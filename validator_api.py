"""
Flask API for the Jinja validator with DOCX support, scan mode, and uploads.

Endpoints:
  GET  /health
  POST /validate   (JSON OR multipart/form-data)

Siili Solutions Oyj
"""

from __future__ import annotations

import csv
import json
import os
import re
from dataclasses import asdict
from io import BytesIO
from pathlib import Path
from typing import Any, Dict, Optional, Tuple, Set, List, Iterable

from flask import Flask, request, jsonify
from flask import Response, send_file, jsonify
from flask_cors import CORS

from pathlib import Path
import yaml

from jinja_validator import (
    JinjaTemplateValidator,
    Policy,
    _extract_docx_with_map,  # accepts bytes/paths/streams
)

# Charset detection for text uploads
try:
    from charset_normalizer import from_bytes as cn_from_bytes  # type: ignore
except Exception:  # pragma: no cover
    cn_from_bytes = None  # type: ignore

# Excel support for variable catalog
try:
    from openpyxl import load_workbook  # type: ignore
    from openpyxl.utils.cell import column_index_from_string  # type: ignore
except Exception:  # pragma: no cover
    load_workbook = None  # type: ignore
    column_index_from_string = None  # type: ignore

# --- Serve UI (HTML/CSS/JS) from ./web-ui ---
ROOT = Path(__file__).resolve().parent
UI_DIR = ROOT / "web-ui"

app = Flask(__name__, static_folder=str(UI_DIR), static_url_path="")

# NOTE: Restrict origins in production (e.g., {"origins": ["https://helsinki.example"]})
allowed = [o.strip() for o in os.getenv("CORS_ALLOW_ORIGINS", "").split(",") if o.strip()]
CORS(app, resources={r"/validate": {"origins": allowed or ["http://localhost:5001", "http://127.0.0.1:5001", "null"]}})
#CORS(app, resources={r"/validate": {"origins": "*"}})

# Cap uploads to ~10 MB by default
app.config["MAX_CONTENT_LENGTH"] = int(os.environ.get("MAX_CONTENT_LENGTH", 10 * 1024 * 1024))

COMMON_ENCODINGS = ("utf-8", "utf-8-sig", "cp1252", "latin-1")

# For docs
OPENAPI_PATH = Path(__file__).with_name("openapi.yaml")


# ------------ helpers ------------

def _decode_bytes(data: bytes, encoding: str = "auto") -> Tuple[str, str]:
    """
    Best-effort decode for uploaded text files (not .docx).
    Returns (text, used_encoding). Raises UnicodeDecodeError on failure.
    """
    if encoding and encoding != "auto":
        return data.decode(encoding), encoding

    last_err: Optional[UnicodeDecodeError] = None
    for enc in COMMON_ENCODINGS:
        try:
            return data.decode(enc), enc
        except UnicodeDecodeError as e:
            last_err = e

    if cn_from_bytes is not None:
        try:
            best = cn_from_bytes(data).best()
            if best is not None:
                return str(best), best.encoding or "detected"
        except Exception:
            pass

    if last_err:
        raise last_err
    raise UnicodeDecodeError("auto", b"", 0, 1, "Unable to decode with common encodings")


def _json_error(message: str, status: int = 400):
    return jsonify({"ok": False, "error": message}), status


# ---- Variable catalog reading (Excel/CSV/JSON) ----

def _normalize_names(items: Iterable[Any]) -> Set[str]:
    """Trim strings, drop empties/None, return a set."""
    out: Set[str] = set()
    for x in items:
        if x is None:
            continue
        s = str(x).strip()
        if s:
            out.add(s)
    return out


def _extract_varlist_from_xlsx(data: bytes, sheet: Optional[str], column: str, has_header: bool = True) -> Set[str]:
    """
    Load variable names from an .xlsx workbook.

    Parameters
    ----------
    sheet : name of worksheet (None => first sheet)
    column: Excel letter (e.g., 'B') OR header text in row 1 (case-sensitive match)
    has_header: when column is a letter, skip row 1 if it contains a header
    """
    if load_wb := load_workbook:
        wb = load_wb(filename=BytesIO(data), read_only=True, data_only=True)
    else:
        raise RuntimeError("Excel support requires 'openpyxl'. Install it on the server.")

    ws = wb[sheet] if sheet else wb.worksheets[0]

    # Header name mode
    if not re.match(r"^[A-Za-z]+$", column):
        header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
        col_idx = None
        for i, h in enumerate(header_row, start=1):
            if h is not None and str(h).strip() == column.strip():
                col_idx = i
                break
        if col_idx is None:
            raise ValueError(f"Column header '{column}' not found in first row of sheet '{ws.title}'.")
        vals: List[Any] = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            vals.append(row[col_idx - 1] if col_idx - 1 < len(row) else None)
        return _normalize_names(vals)

    # Letter mode
    if col_idx_from_str := column_index_from_string:
        col_idx = col_idx_from_str(column)
    else:
        raise RuntimeError("openpyxl missing 'column_index_from_string'")
    start_row = 2 if has_header else 1
    vals: List[Any] = []
    for row in ws.iter_rows(min_row=start_row, values_only=True):
        vals.append(row[col_idx - 1] if col_idx - 1 < len(row) else None)
    return _normalize_names(vals)


def _extract_varlist_from_csv(data: bytes, column: str, encoding: str = "utf-8", has_header: bool = True) -> Set[str]:
    """
    Load variable names from a CSV file.

    Parameters
    ----------
    column : either a 1-based column index given as a string ('1', '2', ...)
             OR a header name (case-sensitive match in the first row).
    """
    text = data.decode(encoding, errors="replace")
    rows = list(csv.reader(text.splitlines()))
    if not rows:
        return set()

    # Header name
    if not column.isdigit():
        header = rows[0]
        if column not in header:
            raise ValueError(f"Column header '{column}' not found in CSV header.")
        idx = header.index(column)
        start = 1
    else:
        idx = max(1, int(column)) - 1
        start = 1 if has_header else 0

    vals = [r[idx] for r in rows[start:] if idx < len(r)]
    return _normalize_names(vals)


# ------------ endpoints ------------

@app.get("/openapi.yaml")
def openapi_yaml():
    # Useful for manual download / editors
    if not OPENAPI_PATH.exists():
        return jsonify({"error": f"openapi.yaml not found at {OPENAPI_PATH}"}), 404
    return send_file(OPENAPI_PATH, mimetype="text/yaml")

@app.get("/openapi.json")
def openapi_json():
    # ReDoc will consume this JSON (avoids YAML parsing surprises)
    try:
        with OPENAPI_PATH.open("r", encoding="utf-8") as f:
            spec = yaml.safe_load(f)
        if not isinstance(spec, dict):
            return jsonify({"error": "OpenAPI root must be a JSON object"}), 500
        return jsonify(spec)
    except FileNotFoundError:
        return jsonify({"error": f"openapi.yaml not found at {OPENAPI_PATH}"}), 404
    except Exception as e:
        return jsonify({"error": f"Failed to load openapi.yaml: {e}"}), 500
    
@app.get("/docs")
def docs():
    html = """
    <!doctype html>
    <html>
      <head>
        <meta charset="utf-8"/>
        <title>Jinja Validator API Docs</title>
        <meta name="viewport" content="width=device-width, initial-scale=1"/>
        <style>html,body{height:100%;margin:0} redoc{height:100%}</style>
      </head>
      <body>
        <redoc spec-url="/openapi.json"></redoc>
        <script src="https://cdn.redoc.ly/redoc/latest/bundles/redoc.standalone.js"></script>
      </body>
    </html>
    """
    return Response(html, mimetype="text/html")


@app.get("/")
def home():
    return app.send_static_file("index.html")

@app.get("/health")
def health():
    return {"status": "ok"}


@app.post("/validate")
def validate():
    """
    Accepts either:
      - JSON: {"template": "...", "data": {...}, "preview": bool,
               "scan": "auto|off|segments|lines",
               "permissive": bool,
               "custom_delims": {"var_start": "...", "var_end": "...", "block_start": "...", "block_end": "..."},
               "varlist": ["name1","name2", ...],
               "varlist_mode": "head"|"fullpath",
               "varlist_exempt": ["is_preview", "..."]  # OPTIONAL — names to NOT check against catalog
               }
      - multipart/form-data:
          file=(.docx|.txt), OR text=(string)
          data_json=(stringified JSON, optional)
          preview, scan, permissive, encoding, and custom_delims fields (optional)
          varlist_file=(.xlsx|.csv)   # upload catalog file
          varlist_sheet=(string)      # for .xlsx only (optional; defaults to first sheet)
          varlist_column=(string)     # Excel letter (A,B,...) OR header name; CSV: header name or 1-based index
          varlist_has_header=(true|false)  # default true
          varlist_encoding=(string)   # CSV only; default utf-8
          varlist_exempt_json=(stringified JSON array)  # OPTIONAL
          varlist_exempt=(string: comma/newline separated)  # OPTIONAL

    Returns ValidationResult as JSON, 200 when ok, 422 when issues were found.
    """
    content_type = (request.content_type or "").lower()

    # Defaults / options
    scan = "auto"
    permissive = False
    data_obj: Optional[Dict[str, Any]] = None
    preview = False
    custom_delims = None
    encoding = "auto"

    # Variable catalog
    allowed_catalog: Optional[Set[str]] = None
    catalog_mode = "head"  # "head" or "fullpath"
    # Exempt names that should never be flagged as unknown when validating var names against a catalogue. 
    # Add more here if needed.
    catalog_exempt: Set[str] = {"is_preview"}

    template_text: Optional[str] = None
    docx_segments = []  # when validating .docx, attach to validator

    # ---------- multipart/form-data ----------
    if "multipart/form-data" in content_type:
        form = request.form

        preview = form.get("preview", "false").lower() == "true"
        permissive = form.get("permissive", "false").lower() == "true"
        scan = form.get("scan", "auto").lower()
        encoding = form.get("encoding", "auto")

        if "custom_delims" in form and form.get("custom_delims"):
            try:
                custom_delims = json.loads(form["custom_delims"])
            except Exception:
                return _json_error("custom_delims must be JSON with {var_start,var_end,block_start,block_end}", 400)

        if form.get("data_json"):
            try:
                data_obj = json.loads(form["data_json"])
                if data_obj is not None and not isinstance(data_obj, dict):
                    return _json_error("data_json must be a JSON object", 400)
            except Exception as e:
                return _json_error(f"Invalid data_json: {e}", 400)

        # Variable catalog via file upload (.xlsx or .csv) OR via text field (JSON array)
        catalog_file = request.files.get("varlist_file") or request.files.get("varlist")
        if catalog_file and catalog_file.filename:
            cname = catalog_file.filename.lower()
            cbytes = catalog_file.read()
            try:
                if cname.endswith(".xlsx"):
                    sheet = form.get("varlist_sheet") or None
                    column = (form.get("varlist_column") or "").strip()
                    if not column:
                        return _json_error("Provide varlist_column for Excel (letter like 'A' or header name).", 400)
                    has_header = form.get("varlist_has_header", "true").lower() == "true"
                    allowed_catalog = _extract_varlist_from_xlsx(cbytes, sheet, column, has_header)
                elif cname.endswith(".csv"):
                    column = (form.get("varlist_column") or "").strip()
                    if not column:
                        return _json_error("Provide varlist_column for CSV (header name or 1-based index).", 400)
                    enc = form.get("varlist_encoding", "utf-8")
                    has_header = form.get("varlist_has_header", "true").lower() == "true"
                    allowed_catalog = _extract_varlist_from_csv(cbytes, column, encoding=enc, has_header=has_header)
                else:
                    return _json_error("Unsupported varlist file type. Use .xlsx or .csv", 400)
            except Exception as e:
                return _json_error(f"Failed to read varlist: {e}", 400)

        if allowed_catalog is None and form.get("varlist_json"):
            try:
                seq = json.loads(form["varlist_json"])
                if not isinstance(seq, list):
                    return _json_error("varlist_json must be a JSON array of strings", 400)
                allowed_catalog = _normalize_names(seq)
            except Exception as e:
                return _json_error(f"Invalid varlist_json: {e}", 400)

        if form.get("varlist_mode"):
            m = form.get("varlist_mode", "head").lower()
            if m in ("head", "fullpath"):
                catalog_mode = m

        # Read exempt list from multipart (JSON array or free-form string)
        if form.get("varlist_exempt_json"):
            try:
                seq = json.loads(form["varlist_exempt_json"])
                if isinstance(seq, list):
                    catalog_exempt |= _normalize_names(seq)
                else:
                    return _json_error("varlist_exempt_json must be a JSON array of strings", 400)
            except Exception as e:
                return _json_error(f"Invalid varlist_exempt_json: {e}", 400)
        elif form.get("varlist_exempt"):
            raw = form.get("varlist_exempt", "")
            # accept comma/newline separated
            parts = [p.strip() for p in re.split(r"[,\\n]+", raw) if p.strip()]
            if parts:
                catalog_exempt |= set(parts)

        # File or text template
        f = request.files.get("file")
        if f and f.filename:
            filename = f.filename.lower()
            data_bytes = f.read()

            if filename.endswith(".docx"):
                # Quick sanity: DOCX should be a ZIP starting with PK\x03\x04
                if not data_bytes.startswith(b"PK\x03\x04"):
                    return _json_error("Uploaded .docx does not look like a valid DOCX (ZIP).", 400)
                try:
                    template_text, docx_segments = _extract_docx_with_map(data_bytes)  # in-memory
                except RuntimeError as e:
                    return _json_error(str(e), 400)
                except Exception as e:
                    return _json_error(f"Failed to read DOCX: {e}", 400)
            else:
                try:
                    template_text, _ = _decode_bytes(data_bytes, encoding=encoding)
                except UnicodeDecodeError as e:
                    tried = ", ".join(COMMON_ENCODINGS) if encoding == "auto" else encoding
                    return _json_error(
                        f"Could not decode file. Try 'encoding=cp1252' or 'latin-1'. (Tried: {tried}). Error: {e}",
                        400
                    )

        elif "text" in form and form.get("text", "").strip():
            template_text = form["text"]
        else:
            return _json_error("Provide either a file (.docx/.txt) or a non-empty 'text' field.", 400)

    # ---------- JSON ----------
    else:
        payload = request.get_json(force=True, silent=True) or {}

        template_text = payload.get("template")
        if template_text is None or not isinstance(template_text, str) or not template_text.strip():
            return _json_error("template (string) is required", 400)

        data_obj = payload.get("data")
        if data_obj is not None and not isinstance(data_obj, dict):
            return _json_error("data must be an object/dict", 400)

        preview = bool(payload.get("preview", False))
        permissive = bool(payload.get("permissive", False))
        scan = str(payload.get("scan", "auto")).lower()
        encoding = str(payload.get("encoding", "auto"))
        custom_delims = payload.get("custom_delims")

        # Variable catalog in JSON
        if isinstance(payload.get("varlist"), list):
            allowed_catalog = _normalize_names(payload["varlist"])
        m = str(payload.get("varlist_mode", "head")).lower()
        if m in ("head", "fullpath"):
            catalog_mode = m

        # Exemptions via JSON (array preferred; also accept comma/newline string)
        if "varlist_exempt" in payload:
            ex = payload["varlist_exempt"]
            if isinstance(ex, list):
                catalog_exempt |= _normalize_names(ex)
            elif isinstance(ex, str):
                parts = [p.strip() for p in re.split(r"[,\\n]+", ex) if p.strip()]
                if parts:
                    catalog_exempt |= set(parts)
            # else: ignore other types silently

    # Build env kwargs (custom delimiters)
    env_kwargs = {}
    if isinstance(custom_delims, dict):
        try:
            env_kwargs = dict(
                variable_start_string=custom_delims["var_start"],
                variable_end_string=custom_delims["var_end"],
                block_start_string=custom_delims["block_start"],
                block_end_string=custom_delims["block_end"],
            )
        except KeyError:
            return _json_error("custom_delims must include var_start, var_end, block_start, block_end", 400)

    # Policy per request
    policy = Policy.permissive() if permissive else Policy.sensible_defaults()

    # Fresh validator per request (avoid cross-request state)
    validator = JinjaTemplateValidator(use_sandbox=True, policy=policy, env_kwargs=env_kwargs)
    validator._scan_mode = scan
    validator._docx_segments = docx_segments

    # Pass catalog whitelist into validator
    validator._allowed_catalog = allowed_catalog
    validator._catalog_mode = catalog_mode
    # Pass exemptions (defaults include "is_preview")
    validator._catalog_exempt = catalog_exempt

    # Validate
    result = validator.validate(template_text or "", sample_data=data_obj, render_preview=preview)

    # HTTP 200 when ok, 422 when issues were found
    return jsonify(asdict(result)), (200 if result.ok else 422)


if __name__ == "__main__":
    # NOTE: production, set debug=False and run behind gunicorn/uwsgi/reverse-proxy.
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5001)), debug=True)
