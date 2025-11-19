"""
A small, validator for Jinja2 templates that favors clear, actionable
feedback for non-technical authors — including *Word-aware* locations
for errors inside .docx files, plus a scan mode to list multiple issues when a
global parse fails.

Highlights
----------
- Parses and reports syntax errors with a small codeframe.
- Collects expected variables + used filters/tests/calls/tags.
- Policy allowlists with sensible defaults (or permissive).
- Optional comparison against sample JSON data (find missing/unused paths).
- Optional preview render using StrictUndefined (early failure on missing keys).
- Resilient text file reading with --encoding and auto-detection.
- DOCX awareness: extract text with a source map so errors/missing vars include
  “Body ▸ Paragraph 12” etc., plus a short anchor string.
- Scan mode: when the global parse fails, optionally scan per-line (TXT)
  or per-paragraph/cell (DOCX) to aggregate multiple issues.
- Variable catalog whitelist — report variables used in templates that are
  NOT present in a catalog list (e.g., from Excel/CSV). See `unknown_variables`.

  Siili Solutions Oyj
"""

from __future__ import annotations

import argparse
import difflib
import io
import json
import re
import sys
from pathlib import Path
from dataclasses import dataclass, asdict, field
from typing import Any, Dict, IO, Iterable, List, Optional, Set, Tuple, Union

from jinja2 import Environment, TemplateSyntaxError, StrictUndefined, UndefinedError
from jinja2.sandbox import SandboxedEnvironment
from jinja2 import nodes
from jinja2.visitor import NodeVisitor

# Encoding detection (graceful fallback if not installed).
try:
    from charset_normalizer import from_path as cn_from_path  # type: ignore
except Exception:  # pragma: no cover
    cn_from_path = None  # type: ignore

# DOCX extraction (graceful fallback if not installed).
try:
    from docx import Document  # type: ignore
except Exception:  # pragma: no cover
    Document = None  # type: ignore


def _flatten_dict(d: Dict[str, Any], prefix: str = "") -> Set[str]:
    """
    Return a set of dotted key paths that exist in a nested dict/list structure.

    Example
    -------
    >>> _flatten_dict({"order": {"items": [{"sku": "A"}]}})
    {'order', 'order.items', 'order.items.0', 'order.items.0.sku'}
    """
    paths: Set[str] = set()

    def walk(value: Any, path_parts: List[str]) -> None:
        dotted = ".".join(path_parts) if path_parts else ""
        if dotted:
            paths.add(dotted)
        if isinstance(value, dict):
            for k, v in value.items():
                walk(v, path_parts + [str(k)])
        elif isinstance(value, list):
            for idx, v in enumerate(value):
                walk(v, path_parts + [str(idx)])

    walk(d, [prefix] if prefix else [])
    return paths


def _resolve_dotted(data: Any, dotted: str) -> Tuple[bool, Optional[Any]]:
    """
    Try to resolve a dotted path (e.g. 'order.customer.name') within `data`.

    Returns
    -------
    (exists, value_if_any)
    """
    cur = data
    if dotted == "":
        return True, cur
    for part in dotted.split("."):
        if isinstance(cur, dict) and part in cur:
            cur = cur[part]
        elif isinstance(cur, list) and part.isdigit():
            i = int(part)
            if 0 <= i < len(cur):
                cur = cur[i]
            else:
                return False, None
        else:
            return False, None
    return True, cur


# -------------------- DOCX source mapping --------------------

@dataclass
class DocxSegment:
    """
    Describe where a chunk of extracted text came from in a .docx file.

    Attributes
    ----------
    scope : str        # 'body', 'header', 'footer'
    section_idx : int  # section index (0 for body)
    kind : str         # 'paragraph' or 'table_cell'
    para_idx : int     # paragraph index if kind == paragraph else -1
    table_idx : int    # table index if kind == table_cell else -1
    row, col : int     # cell coords if kind == table_cell else -1
    text : str         # visible text for this segment
    start_line, end_line : int  # 1-based in concatenated text (one line per segment)
    """
    scope: str
    section_idx: int
    kind: str
    para_idx: int
    table_idx: int
    row: int
    col: int
    text: str
    start_line: int
    end_line: int


def _extract_docx_with_map(source: Union[str, bytes, IO[bytes]]) -> tuple[str, List[DocxSegment]]:
    """
    Read a .docx and return (concatenated_text, segments).

    Accepts:
      - path: str
      - bytes: raw .docx (ZIP) bytes
      - file-like: binary stream positioned at start

    Each paragraph or table cell becomes a single "logical line" in the
    concatenated text, enabling mapping of Jinja error line numbers back
    to a Word location.
    """
    if Document is None:
        raise RuntimeError("python-docx is required for .docx parsing (pip install python-docx)")

    # Accept multiple input forms for flexibility (works better on Windows/API)
    if isinstance(source, (bytes, bytearray)):
        doc = Document(io.BytesIO(source))
    elif hasattr(source, "read"):
        doc = Document(source)  # type: ignore[arg-type]
    else:
        doc = Document(str(source))

    segments: List[DocxSegment] = []
    lines: List[str] = []
    line_no = 1

    def add_para(scope: str, section_i: int, txt: str, para_i: int):
        nonlocal line_no
        start = line_no
        lines.append(txt)
        line_no += 1
        segments.append(DocxSegment(scope, section_i, "paragraph", para_i, -1, -1, -1, txt, start, start))

    def add_cell(scope: str, section_i: int, table_i: int, r: int, c: int, txt: str):
        nonlocal line_no
        start = line_no
        lines.append(txt)
        line_no += 1
        segments.append(DocxSegment(scope, section_i, "table_cell", -1, table_i, r, c, txt, start, start))

    # Body paragraphs
    for pi, p in enumerate(doc.paragraphs):
        add_para("body", 0, "".join(run.text for run in p.runs), pi)

    # Body tables
    for ti, t in enumerate(doc.tables):
        for r, row in enumerate(t.rows):
            for c, cell in enumerate(row.cells):
                cell_txt = "\n".join("".join(run.text for run in p.runs) for p in cell.paragraphs)
                add_cell("body", 0, ti, r, c, cell_txt)

    # Headers/Footers per section
    for si, section in enumerate(doc.sections):
        hdr = section.header
        if hdr:
            for pi, p in enumerate(hdr.paragraphs):
                add_para("header", si, "".join(run.text for run in p.runs), pi)
        ftr = section.footer
        if ftr:
            for pi, p in enumerate(ftr.paragraphs):
                add_para("footer", si, "".join(run.text for run in p.runs), pi)

    full_text = "\n".join(lines)
    return full_text, segments


def _map_lineno_to_docx(lineno: int, segs: List[DocxSegment]) -> dict:
    """Map 1-based concatenated-text line number to a friendly Word location."""
    for s in segs:
        if s.start_line <= lineno <= s.end_line:
            if s.kind == "paragraph":
                label = f"{s.scope.title()} ▸ Paragraph {s.para_idx + 1}"
            else:
                label = f"{s.scope.title()} ▸ Table {s.table_idx + 1} ▸ Row {s.row + 1} ▸ Col {s.col + 1}"
            anchor = (s.text or "")[:120]
            return {"label": label, "anchor": anchor}
    return {"label": "Unknown location", "anchor": ""}


def _anchor_from_line(line: str) -> str:
    """Make a short, stable search needle from a line of text."""
    return (line or "").strip()[:120]


# -------------------- scan mode helpers --------------------

def _has_jinja(text: str) -> bool:
    return ("{{" in text and "}}" in text) or ("{%" in text and "%}" in text) or ("{#" in text and "#}" in text)

def _merge_usage(into: Dict[str, Set[str]], add: "UsageVisitor") -> None:
    into["variables"].update(add.variables)
    into["filters"].update(add.filters)
    into["tests"].update(add.tests)
    into["calls"].update(add.calls)
    into["tags"].update(add.tags)

def _make_excerpt(src: str, err_line: int, context: int = 2) -> str:
    lines = src.splitlines() or [src]
    n = len(lines)
    start = max(1, err_line - context)
    end = min(n, err_line + context)
    out = []
    trans = {ord(c): " " for c in "\x00\x01\x02\x03\x04\x05\x06\x07\x08\x0b\x0c\x0e\x0f"}
    for i in range(start, end + 1):
        mark = ">" if i == err_line else " "
        safe = (lines[i - 1] if 1 <= i <= n else "").translate(trans)
        out.append(f"{mark} {i:4d} | {safe}")
    return "\n".join(out)

def _looks_like_unclosed_block(err: Exception) -> bool:
    """Return True if this TemplateSyntaxError is the typical 'fragment opened a block
    but ended before the closing tag' case. We only use this in scan modes."""
    from jinja2 import TemplateSyntaxError
    if not isinstance(err, TemplateSyntaxError):
        return False
    msg = (err.message or "").lower()
    return msg.startswith("unexpected end of template")

# Detector for obviously missing Jinja delimiters in a fragment
def _detect_missing_delimiters(text: str, env: Any) -> Optional[Dict[str, Any]]:
    """
    Heuristically detect a case where a fragment contains an opening Jinja delimiter
    but not its closing counterpart (e.g., '{{ foo }' or '{% if x %' without '%}').

    Returns a dict like {"open": "{{", "close": "}}", "index": <pos>} or None if not detected.

    This does not fully parse/tolerate nesting; it's intended for short per-line/segment scans.
    """
    opens = [
        (env.variable_start_string, env.variable_end_string),
        (env.block_start_string, env.block_end_string),
        (env.comment_start_string, env.comment_end_string),
    ]
    i = 0
    n = len(text)
    # Scan left-to-right, look for any opening token; ensure its matching close exists after it
    while i < n:
        # Find next opening occurrence among the three kinds
        next_hits = []
        for op, cl in opens:
            pos = text.find(op, i)
            if pos != -1:
                next_hits.append((pos, op, cl))
        if not next_hits:
            return None
        pos, op, cl = min(next_hits, key=lambda t: t[0])  # earliest opening
        # Look for its closer after the opening
        end_pos = text.find(cl, pos + len(op))
        if end_pos == -1:
            return {"open": op, "close": cl, "index": pos}
        # Advance after this complete tag and keep scanning
        i = end_pos + len(cl)
    return None


# -------------------- AST Visitor --------------------

class UsageVisitor(NodeVisitor):
    """
    Walk Jinja2 AST to collect:
      - used variables (as dotted paths),
      - used filters/tests/function calls,
      - used tags.

    Tracks local names introduced by for/set/assign/macro and excludes them
    from expected variables. Also records occurrence line numbers to enable
    Word-aware mapping for missing vars.
    """

    def __init__(self) -> None:
        self.filters: Set[str] = set()
        self.tests: Set[str] = set()
        self.calls: Set[str] = set()
        self.tags: Set[str] = set()
        self.variables: Set[str] = set()
        self._locals: List[Set[str]] = [set()]
        self.var_lines: Dict[str, Set[int]] = {}

    def _chain_from_getattr(self, node: nodes.Getattr) -> Optional[str]:
        parts: List[str] = []
        cur = node
        while isinstance(cur, nodes.Getattr):
            parts.append(cur.attr)
            cur = cur.node  # type: ignore[assignment]
        if isinstance(cur, nodes.Name):
            parts.append(cur.name)
            parts.reverse()
            return ".".join(parts)
        return None

    def _chain_from_getitem(self, node: nodes.Getitem) -> Optional[str]:
        parts: List[str] = []
        cur = node
        while isinstance(cur, nodes.Getitem):
            key = cur.arg
            if isinstance(key, nodes.Const):
                parts.append(str(key.value))
            else:
                parts.append("[]")
            cur = cur.node  # type: ignore[assignment]
        if isinstance(cur, nodes.Name):
            parts.append(cur.name)
            parts.reverse()
            return ".".join(parts)
        elif isinstance(cur, nodes.Getattr):
            prefix = self._chain_from_getattr(cur)
            if prefix is None:
                return None
            parts.reverse()
            return prefix + ("." + ".".join(parts) if parts else "")
        return None

    def _is_local(self, dotted: str) -> bool:
        head = dotted.split(".", 1)[0]
        return any(head in scope for scope in self._locals)

    def _record(self, dotted: Optional[str], lineno: int) -> None:
        if dotted and not self._is_local(dotted) and not dotted.endswith(".[]"):
            self.variables.add(dotted)
            if lineno:
                self.var_lines.setdefault(dotted, set()).add(lineno)

    # Tags / scope
    def visit_For(self, node: nodes.For) -> None:
        self.tags.add("for")
        scope: Set[str] = set()

        def add_target(t: nodes.Node) -> None:
            if isinstance(t, nodes.Name):
                scope.add(t.name)
            elif isinstance(t, (nodes.Tuple, list, tuple)):
                for x in getattr(t, "items", []) or getattr(t, "nodes", []) or []:
                    add_target(x)

        add_target(node.target)
        self._locals.append(scope)
        self.generic_visit(node)
        self._locals.pop()

    def visit_If(self, node: nodes.If) -> None:
        self.tags.add("if")
        self.generic_visit(node)

    def visit_Set(self, node: nodes.Set) -> None:
        self.tags.add("set")
        scope = self._locals[-1]
        if isinstance(node.target, nodes.Name):
            scope.add(node.target.name)
        elif isinstance(node.target, nodes.Tuple):
            for n in node.target.items:
                if isinstance(n, nodes.Name):
                    scope.add(n.name)
        self.generic_visit(node)

    def visit_Assign(self, node: nodes.Assign) -> None:
        self.tags.add("assign")
        if isinstance(node.target, nodes.Name):
            self._locals[-1].add(node.target.name)
        self.generic_visit(node)

    def visit_Macro(self, node: nodes.Macro) -> None:
        self.tags.add("macro")
        scope = set(arg.name for arg in node.args) if node.args else set()
        self._locals.append(scope)
        self.generic_visit(node)
        self._locals.pop()

    def visit_Import(self, node: nodes.Import) -> None:
        self.tags.add("import")
        self.generic_visit(node)

    def visit_FromImport(self, node: nodes.FromImport) -> None:
        self.tags.add("from_import")
        self.generic_visit(node)

    def visit_Block(self, node: nodes.Block) -> None:
        self.tags.add("block")
        self.generic_visit(node)

    def visit_Extends(self, node: nodes.Extends) -> None:
        self.tags.add("extends")
        self.generic_visit(node)

    def visit_Include(self, node: nodes.Include) -> None:
        self.tags.add("include")
        self.generic_visit(node)

    # Expressions / usage
    def visit_Filter(self, node: nodes.Filter) -> None:
        self.filters.add(node.name)
        self.generic_visit(node)

    def visit_Test(self, node: nodes.Test) -> None:
        self.tests.add(node.name)
        self.generic_visit(node)

    def visit_Call(self, node: nodes.Call) -> None:
        n = node.node
        if isinstance(n, nodes.Name):
            self.calls.add(n.name)
        elif isinstance(n, nodes.Getattr):
            chain = self._chain_from_getattr(n)
            if chain:
                self.calls.add(chain)
        self.generic_visit(node)

    def visit_Name(self, node: nodes.Name) -> None:
        if node.ctx == "load":
            self._record(node.name, getattr(node, "lineno", 0))
        self.generic_visit(node)

    def visit_Getattr(self, node: nodes.Getattr) -> None:
        self._record(self._chain_from_getattr(node), getattr(node, "lineno", 0))
        self.generic_visit(node)

    def visit_Getitem(self, node: nodes.Getitem) -> None:
        self._record(self._chain_from_getitem(node), getattr(node, "lineno", 0))
        self.generic_visit(node)


# -------------------- Validation Core --------------------

@dataclass
class Policy:
    """
    Simple allowlists for filters, tests, function calls, and tags.
    Set a field to None to allow everything. Use Policy.permissive() to disable all checks.

    If someone uses a filter/tag/etc not defined here (e.g., map, select, macro), the validator 
    raises a Policy issue, unless Permissive mode is turned on.
    
    Example:
    If "map" isn’t in the allowlist and is used in the template by its author:
        - Permissive OFF → disallowed_filters: ["map"] error.
        - Permissive ON → no policy error (other types of checks still apply).

    """
    allowed_filters: Optional[Set[str]] = None
    allowed_tests: Optional[Set[str]] = None
    allowed_calls: Optional[Set[str]] = None
    allowed_tags: Optional[Set[str]] = None

    @staticmethod
    def permissive() -> "Policy":
        return Policy()

    @staticmethod
    def sensible_defaults() -> "Policy":
        return Policy(
            allowed_filters={
                "upper", "lower", "title", "capitalize", "default", "replace",
                "join", "length", "trim", "striptags", "urlencode", "safe",
                "round", "int", "float", "format", "json", "dictsort", "sort"
            },
            allowed_tests={"defined", "undefined", "equalto", "in", "odd", "even"},
            allowed_calls={"range"},
            allowed_tags={"if", "for", "set", "include"}
        )


@dataclass
class ValidationResult:
    ok: bool
    errors: List[Dict[str, Any]]
    expected_variables: List[str]
    used_filters: List[str]
    used_tests: List[str]
    used_calls: List[str]
    used_tags: List[str]
    missing_from_data: List[Dict[str, Any]]
    unused_data_paths: List[str]
    preview: Optional[str] = None
    # Variables used in template that are not present in the provided catalog
    unknown_variables: List[Dict[str, Any]] = field(default_factory=list)

    def to_json(self) -> str:
        return json.dumps(asdict(self), ensure_ascii=False, indent=2)


class JinjaTemplateValidator:
    """
    Validates a Jinja2 template string.

    DOCX mapping
    ------------
    If `self._docx_segments` (set by caller) is present, syntax errors and
    missing-variable entries include `docx_location` and `anchor`.

    Scan mode
    ---------
    If `_scan_mode` is "auto"/"segments"/"lines" and the global parse fails,
    the validator scans segments/lines independently to aggregate issues.

    Catalog whitelist
    -----------------
    The API may set `_allowed_catalog` (set[str]) and `_catalog_mode` ("head"|"fullpath")
    to report variables not present in an approved list (e.g., Excel/CSV).
    ----
    `_catalog_exempt` (set[str]) lists names that should *never* be reported as unknown
    against the catalog (e.g., special control flags like "is_preview").
    """
    def __init__(
        self,
        use_sandbox: bool = True,
        policy: Optional[Policy] = None,
        env_kwargs: Optional[Dict[str, Any]] = None,
    ) -> None:
        self.policy = policy or Policy.sensible_defaults()
        self.env = (SandboxedEnvironment if use_sandbox else Environment)(
            undefined=StrictUndefined, **(env_kwargs or {})
        )
        self._docx_segments: List[DocxSegment] = []
        self._scan_mode: str = "off"
        # Variable catalog whitelist support
        self._allowed_catalog: Optional[Set[str]] = None
        self._catalog_mode: str = "head"  # "head" (top-level name) or "fullpath"
        # Always-allowed catalog keys (by default, allow 'is_preview')
        self._catalog_exempt: Set[str] = {"is_preview"}

    # --- scan implementations ---

    def _scan_docx_segments(
        self,
        segments: List[DocxSegment],
        sample_data: Optional[Dict[str, Any]],
    ) -> ValidationResult:
        errors: List[Dict[str, Any]] = []
        accum = {"variables": set(), "filters": set(), "tests": set(), "calls": set(), "tags": set()}
        var_lines: Dict[str, Set[int]] = {}

        for seg in segments:
            txt = seg.text or ""

            # catch cases with only an opener (e.g., "{{ foo }") that _has_jinja() would skip
            early_miss = _detect_missing_delimiters(txt, self.env)
            if early_miss and not _has_jinja(txt):
                anchor = (txt[early_miss["index"]:] if isinstance(early_miss.get("index"), int) and early_miss["index"] >= 0 else txt)[:120]
                errors.append({
                    "type": "syntax",
                    "line": seg.start_line,
                    "message": f"Missing closing '{early_miss['close']}' for '{early_miss['open']}' in this paragraph/cell.",
                    "excerpt": _make_excerpt(txt, 1),
                    "docx_location": (
                        f"{seg.scope.title()} ▸ Paragraph {seg.para_idx + 1}"
                        if seg.kind == "paragraph"
                        else f"{seg.scope.title()} ▸ Table {seg.table_idx + 1} ▸ Row {seg.row + 1} ▸ Col {seg.col + 1}"
                    ),
                    "anchor": anchor,
                    "hint": f"Add the closing '{early_miss['close']}'. If a tag intentionally spans multiple paragraphs/cells, run with Scan: off to validate the whole template at once.",
                })
                continue  # handled this segment

            if not _has_jinja(txt):
                continue

            try:
                ast = self.env.parse(txt)
                v = UsageVisitor()
                v.visit(ast)
                _merge_usage(accum, v)
                for var in v.variables:
                    var_lines.setdefault(var, set()).add(seg.start_line)
            except TemplateSyntaxError as e:
                err_line = e.lineno or 1

                # Try to present a clearer "missing closer" message when applicable
                miss = _detect_missing_delimiters(txt, self.env)
                if miss:
                    anchor = (txt[miss["index"]:] if isinstance(miss.get("index"), int) and miss["index"] >= 0 else txt)[:120]
                    errors.append({
                        "type": "syntax",
                        "line": seg.start_line,
                        "message": f"Missing closing '{miss['close']}' for '{miss['open']}' in this paragraph/cell.",
                        "excerpt": _make_excerpt(txt, err_line),
                        "docx_location": (
                            f"{seg.scope.title()} ▸ Paragraph {seg.para_idx + 1}"
                            if seg.kind == "paragraph"
                            else f"{seg.scope.title()} ▸ Table {seg.table_idx + 1} ▸ Row {seg.row + 1} ▸ Col {seg.col + 1}"
                        ),
                        "anchor": anchor,
                    })
                    continue

                # Downgrade fragment "unexpected end of template" to a scan hint
                if _looks_like_unclosed_block(e):
                    errors.append({
                        "type": "scan_unclosed_block",
                        "line": seg.start_line,
                        "message": "Fragment opened a control block but doesn't include its closing tag. Common during per-segment scanning; the full template may be fine.",
                        "excerpt": _make_excerpt(txt, err_line),
                        "docx_location": (
                            f"{seg.scope.title()} ▸ Paragraph {seg.para_idx + 1}"
                            if seg.kind == "paragraph"
                            else f"{seg.scope.title()} ▸ Table {seg.table_idx + 1} ▸ Row {seg.row + 1} ▸ Col {seg.col + 1}"
                        ),
                        "anchor": (txt or "")[:120],
                        "hint": "Run with Scan: off to check the whole template; or keep paired tags in one paragraph if you want segment scanning to pass."
                    })
                else:
                    errors.append({
                        "type": "syntax",
                        "line": seg.start_line,
                        "message": e.message,
                        "excerpt": _make_excerpt(txt, err_line),
                        "docx_location": (
                            f"{seg.scope.title()} ▸ Paragraph {seg.para_idx + 1}"
                            if seg.kind == "paragraph"
                            else f"{seg.scope.title()} ▸ Table {seg.table_idx + 1} ▸ Row {seg.row + 1} ▸ Col {seg.col + 1}"
                        ),
                        "anchor": (txt or "")[:120],
                    })

        if sample_data is None:
            missing, unused = [], set()
        else:
            missing, unused = self._compare_with_data(accum["variables"], sample_data)

        if segments and missing:
            for m in missing:
                occs = []
                for ln in sorted(var_lines.get(m["path"], [])):
                    loc = _map_lineno_to_docx(ln, segments)
                    occs.append({"line": ln, "docx_location": loc["label"], "anchor": loc["anchor"]})
                if occs:
                    m["occurs_at"] = occs

        # Compute unknowns against catalog
        unknowns = self._check_catalog_unknowns(accum["variables"], var_lines)

        # ok ignores scan hints (scan_unclosed_block)
        ok = not [e for e in errors if e.get("type") not in ("scan_unclosed_block",)]
        return ValidationResult(
            ok=ok,
            errors=errors,
            expected_variables=sorted(accum["variables"]),
            used_filters=sorted(accum["filters"]),
            used_tests=sorted(accum["tests"]),
            used_calls=sorted(accum["calls"]),
            used_tags=sorted(accum["tags"]),
            missing_from_data=missing,
            unused_data_paths=sorted(unused)[:2000],
            preview=None,
            unknown_variables=unknowns,
        )

    def _scan_text_lines(
        self,
        template_str: str,
        sample_data: Optional[Dict[str, Any]],
    ) -> ValidationResult:
        errors: List[Dict[str, Any]] = []
        accum = {"variables": set(), "filters": set(), "tests": set(), "calls": set(), "tags": set()}
        var_lines: Dict[str, Set[int]] = {}

        lines = template_str.splitlines()
        for i, line in enumerate(lines, start=1):
            # Catch single-opener cases (e.g., "{{ x }") that _has_jinja() would skip
            early_miss = _detect_missing_delimiters(line, self.env)
            if early_miss and not _has_jinja(line):
                idx = early_miss.get("index")
                anchor_src = line[idx:] if isinstance(idx, int) and idx >= 0 else line
                errors.append({
                    "type": "syntax",
                    "line": i,
                    "message": f"Missing closing '{early_miss['close']}' for '{early_miss['open']}' on this line.",
                    "excerpt": _make_excerpt(line, 1),
                    "anchor": anchor_src[:120],
                    "hint": f"Add the closing '{early_miss['close']}'. If your tag spans multiple lines by design, use Scan: off to validate as one template.",
                })
                continue  # handled this line

            if not _has_jinja(line):
                continue

            try:
                ast = self.env.parse(line)
                v = UsageVisitor()
                v.visit(ast)
                _merge_usage(accum, v)
                for var in v.variables:
                    var_lines.setdefault(var, set()).add(i)
            except TemplateSyntaxError as e:
                err_line = e.lineno or 1

                # Try to present a clearer "missing closer" message when applicable
                miss = _detect_missing_delimiters(line, self.env)
                if miss:
                    idx = miss.get("index")
                    anchor_src = line[idx:] if isinstance(idx, int) and idx >= 0 else line
                    errors.append({
                        "type": "syntax",
                        "line": i,
                        "message": f"Missing closing '{miss['close']}' for '{miss['open']}' on this line.",
                        "excerpt": _make_excerpt(line, err_line),
                        "anchor": anchor_src[:120],
                    })
                    continue

                # Downgrade fragment "unexpected end of template" to a scan hint
                if _looks_like_unclosed_block(e):
                    errors.append({
                        "type": "scan_unclosed_block",
                        "line": i,
                        "message": "Line/fragment opened a control block but doesn't include its closing tag. This often happens with per-line scanning.",
                        "excerpt": _make_excerpt(line, err_line),
                        "hint": "Use Scan: off to validate the whole template; or keep opening/closing tags on the same line if you rely on line scanning."
                    })
                else:
                    errors.append({
                        "type": "syntax",
                        "line": i,
                        "message": e.message,
                        "excerpt": _make_excerpt(line, err_line),
                    })

        if sample_data is None:
            missing, unused = [], set()
        else:
            missing, unused = self._compare_with_data(accum["variables"], sample_data)

        if missing:
            for m in missing:
                occs = [{"line": ln, "anchor": (lines[ln - 1] if 1 <= ln <= len(lines) else "")[:120]}
                        for ln in sorted(var_lines.get(m["path"], []))]
                if occs:
                    m["occurs_at"] = occs

        # Compute unknowns against catalog
        unknowns = self._check_catalog_unknowns(accum["variables"], var_lines)

        # ok ignores scan hints (scan_unclosed_block)
        ok = not [e for e in errors if e.get("type") not in ("scan_unclosed_block",)]
        return ValidationResult(
            ok=ok,
            errors=errors,
            expected_variables=sorted(accum["variables"]),
            used_filters=sorted(accum["filters"]),
            used_tests=sorted(accum["tests"]),
            used_calls=sorted(accum["calls"]),
            used_tags=sorted(accum["tags"]),
            missing_from_data=missing,
            unused_data_paths=sorted(unused)[:2000],
            preview=None,
            unknown_variables=unknowns,
        )

    # --- main validate ---

    def validate(
        self,
        template_str: str,
        sample_data: Optional[Dict[str, Any]] = None,
        render_preview: bool = False,
        preview_max_chars: int = 20000,
    ) -> ValidationResult:
        try:
            ast = self.env.parse(template_str)
        except TemplateSyntaxError as e:
            scan_mode = getattr(self, "_scan_mode", "off")
            if scan_mode == "segments" and self._docx_segments:
                return self._scan_docx_segments(self._docx_segments, sample_data)
            if scan_mode == "lines":
                return self._scan_text_lines(template_str, sample_data)
            if scan_mode == "auto":
                if self._docx_segments:
                    return self._scan_docx_segments(self._docx_segments, sample_data)
                else:
                    return self._scan_text_lines(template_str, sample_data)

            err = self._syntax_error_to_dict(e, template_str)
            if self._docx_segments:
                loc = _map_lineno_to_docx(err["line"], self._docx_segments)
                lines = template_str.splitlines()
                line_txt = lines[err["line"] - 1] if 1 <= err["line"] <= len(lines) else ""
                err["docx_location"] = loc["label"]
                err["anchor"] = loc["anchor"] or _anchor_from_line(line_txt)

            return ValidationResult(
                ok=False,
                errors=[err],
                expected_variables=[],
                used_filters=[],
                used_tests=[],
                used_calls=[],
                used_tags=[],
                missing_from_data=[],
                unused_data_paths=[],
                preview=None,
                unknown_variables=[],
            )

        visitor = UsageVisitor()
        visitor.visit(ast)

        policy_errors = self._check_policy(visitor)
        # Only check against data when it's provided
        if sample_data is None:
            missing, unused = [], set()
        else:
            missing, unused = self._compare_with_data(visitor.variables, sample_data)

        if self._docx_segments and missing:
            lines = template_str.splitlines()
            for m in missing:
                occ_lines = sorted(visitor.var_lines.get(m["path"], []))
                occs: List[Dict[str, Any]] = []
                for ln in occ_lines:
                    loc = _map_lineno_to_docx(ln, self._docx_segments)
                    snippet = lines[ln - 1] if 1 <= ln <= len(lines) else ""
                    occs.append({
                        "line": ln,
                        "docx_location": loc["label"],
                        "anchor": loc["anchor"] or _anchor_from_line(snippet),
                    })
                if occs:
                    m["occurs_at"] = occs

        # Compute unknowns vs catalog (head/fullpath)
        unknowns = self._check_catalog_unknowns(visitor.variables, visitor.var_lines)

        preview: Optional[str] = None
        render_errors: List[Dict[str, Any]] = []
        if render_preview and sample_data is not None:
            try:
                tmpl = self.env.from_string(template_str)
                preview = tmpl.render(**sample_data)
            except UndefinedError as e:
                render_errors.append({"type": "undefined", "message": str(e)})
            except Exception as e:
                render_errors.append({"type": "render", "message": str(e)})

        # NOTE: unknown_variables are considered a blocking issue
        ok = not policy_errors and not render_errors and not unknowns
        return ValidationResult(
            ok=ok,
            errors=policy_errors + render_errors,
            expected_variables=sorted(visitor.variables),
            used_filters=sorted(visitor.filters),
            used_tests=sorted(visitor.tests),
            used_calls=sorted(visitor.calls),
            used_tags=sorted(visitor.tags),
            missing_from_data=missing,
            unused_data_paths=sorted(unused)[:2000],
            preview=(preview[:preview_max_chars] if preview else None),
            unknown_variables=unknowns,
        )

    def _check_policy(self, v: UsageVisitor) -> List[Dict[str, Any]]:
        errors: List[Dict[str, Any]] = []

        def check(kind: str, used: Set[str], allowed: Optional[Set[str]]) -> None:
            if allowed is None:
                return
            disallowed = sorted(x for x in used if x not in allowed)
            if disallowed:
                errors.append(
                    {
                        "type": f"disallowed_{kind}",
                        "message": f"These {kind} are not allowed by policy: {', '.join(disallowed)}",
                        "items": disallowed,
                        "hint": f"Allowed {kind}: {', '.join(sorted(allowed))}",
                    }
                )

        check("filters", v.filters, self.policy.allowed_filters)
        check("tests", v.tests, self.policy.allowed_tests)
        check("calls", v.calls, self.policy.allowed_calls)

        if self.policy.allowed_tags is not None:
            disallowed_tags = sorted(t for t in v.tags if t not in self.policy.allowed_tags)
            if disallowed_tags:
                errors.append(
                    {
                        "type": "disallowed_tags",
                        "message": f"These tags are not allowed by policy: {', '.join(disallowed_tags)}",
                        "items": disallowed_tags,
                        "hint": f"Allowed tags: {', '.join(sorted(self.policy.allowed_tags))}",
                    }
                )
        return errors
    
    def _compare_with_data(
        self, variables: Iterable[str], data: Dict[str, Any]
    ) -> Tuple[List[Dict[str, Any]], Set[str]]:
        """
        Compare expected variables against provided sample `data`.

        Returns
        -------
        (missing, unused_paths)
          - missing: list of dicts like {"path": dotted, "suggestion": closest_or_None}
          - unused_paths: set of dotted paths in `data` that the template never uses

        Notes
        -----
        - Uses `_resolve_dotted` to check existence.
        - Suggests closest match via difflib when available.
        - A data path is considered "used" if it equals an expected var or is nested under it.
        """
        missing: List[Dict[str, Any]] = []
        data_paths = _flatten_dict(data)

        # Check each expected var against data; propose a close match if missing
        for var in variables:
            exists, _ = _resolve_dotted(data, var)
            if not exists:
                suggestion = None
                if data_paths:
                    suggestion = next(iter(difflib.get_close_matches(var, data_paths, n=1, cutoff=0.6)), None)
                missing.append({"path": var, "suggestion": suggestion})

        # Consider a data path "used" if it equals an expected var or is nested under it
        used_prefixes = set(variables)
        unused = set(
            p for p in data_paths
            if not any(p == u or p.startswith(u + ".") for u in used_prefixes)
        )
        return missing, unused

    # Catalog whitelist comparison
    def _check_catalog_unknowns(
        self,
        variables: Iterable[str],
        var_lines: Optional[Dict[str, Set[int]]] = None,
    ) -> List[Dict[str, Any]]:
        """
        Compare variables used in the template against an allowed catalog.

        Returns a list of dicts like:
          {"name": "<full dotted>", "catalog_key":"<head or fullpath>", "suggestion": "<closest or None>",
           "occurs_at":[{line, docx_location?, anchor?}, ...]}

        If `_catalog_mode` == "head", compare only the top-level name (before '.').
        If `_catalog_mode` == "fullpath", compare the entire dotted path.
        ----
        Honors `self._catalog_exempt` — keys listed there are always allowed and
        will not be reported as unknown (e.g., {"is_preview"} by default).
        """
        allowed = self._allowed_catalog
        if not allowed:
            return []

        mode = getattr(self, "_catalog_mode", "head")
        allowed_norm = {s.strip() for s in allowed if isinstance(s, str) and s.strip()}
        exempt = getattr(self, "_catalog_exempt", set())  # names to ignore in catalog check
        unknown: List[Dict[str, Any]] = []

        for var in variables:
            head = var.split(".", 1)[0]
            key = head if mode == "head" else var

            # Skip unknown reporting for exempt keys
            if key in exempt:
                continue

            if key not in allowed_norm:
                suggestion = None
                if allowed_norm:
                    suggestion = next(iter(difflib.get_close_matches(key, list(allowed_norm), n=1, cutoff=0.6)), None)
                entry: Dict[str, Any] = {"name": var, "catalog_key": key, "suggestion": suggestion}

                # Occurrence mapping (line -> Word location if available)
                if var_lines:
                    occs: List[Dict[str, Any]] = []
                    for ln in sorted(var_lines.get(var, set())):
                        loc = _map_lineno_to_docx(ln, self._docx_segments) if self._docx_segments else {"label": "", "anchor": ""}
                        occ = {"line": ln}
                        if loc.get("label"):
                            occ["docx_location"] = loc["label"]
                        if loc.get("anchor"):
                            occ["anchor"] = loc["anchor"]
                        occs.append(occ)
                    if occs:
                        entry["occurs_at"] = occs

                unknown.append(entry)

        return unknown

    @staticmethod
    def _syntax_error_to_dict(err: TemplateSyntaxError, source: str) -> Dict[str, Any]:
        """
        Convert a TemplateSyntaxError to a dict with a codeframe and a heuristic hint
        for the common mistake: two names inside {{ ... }} (e.g., '{{ r projektin_nimi }}').
        """
        line_no = err.lineno or 1
        lines = source.splitlines()
        start = max(1, line_no - 2)
        end = min(len(lines), line_no + 2)
        excerpt = []
        for i in range(start, end + 1):
            prefix = ">" if i == line_no else " "
            safe = lines[i - 1].translate({ord(c): " " for c in "\x00\x01\x02\x03\x04\x05\x06\x07\x08\x0b\x0c\x0e\x0f"})
            excerpt.append(f"{prefix} {i:4d} | {safe}")

        payload = {
            "type": "syntax",
            "line": line_no,
            "message": err.message,
            "excerpt": "\n".join(excerpt),
        }

        offending = lines[line_no - 1] if 1 <= line_no <= len(lines) else ""
        m = re.search(r"\{\{\s*([A-Za-z_]\w*)\s+([A-Za-z_]\w*)\s*\}\}", offending)
        if m:
            a, b = m.group(1), m.group(2)
            payload["hint"] = (
                f"Found two names '{a} {b}' inside {{ }}. "
                f"If '{a}' is a filter, use '{{{{ {b} | {a} }}}}'. "
                f"If it's a function, use '{{{{ {a}({b}) }}}}'. "
                f"Otherwise remove '{a}': '{{{{ {b} }}}}'."
            )
        return payload


# -------------------- Text reading & CLI --------------------

COMMON_ENCODINGS = ("utf-8", "utf-8-sig", "cp1252", "latin-1")


def _read_text_safely(path: str, encoding: str = "auto") -> Tuple[str, str]:
    """
    Read a text file using robust encoding handling.
    Returns (text, used_encoding). Raises UnicodeDecodeError on failure.
    """
    attempts: List[str] = []
    if encoding != "auto":
        attempts = [encoding]
    else:
        attempts = list(COMMON_ENCODINGS)

    last_err: Optional[UnicodeDecodeError] = None
    for enc in attempts:
        try:
            with open(path, "r", encoding=enc) as f:
                return f.read(), enc
        except UnicodeDecodeError as e:
            last_err = e

    if cn_from_path is not None:
        try:
            best = cn_from_path(path).best()
            if best is not None:
                return str(best), best.encoding or "detected"
        except Exception:
            pass

    if last_err:
        raise last_err
    raise UnicodeDecodeError("auto", b"", 0, 1, "Unable to decode file with tried encodings")


def _print_json_error_and_exit(msg: str, kind: str = "encoding"):
    payload = {
        "ok": False,
        "errors": [{"type": kind, "message": msg}],
        "expected_variables": [],
        "used_filters": [],
        "used_tests": [],
        "used_calls": [],
        "used_tags": [],
        "missing_from_data": [],
        "unused_data_paths": [],
        "preview": None,
        "unknown_variables": [],
    }
    sys.stdout.write(json.dumps(payload, ensure_ascii=False, indent=2) + "\n")
    sys.exit(2)


def _load_data(path_or_json: Optional[str]) -> Optional[Dict[str, Any]]:
    if not path_or_json:
        return None
    try:
        p = Path(path_or_json)
        if p.exists():
            with p.open("r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return json.loads(path_or_json)


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="Validate Jinja2 templates (.txt or .docx).")
    parser.add_argument("template", help="Path to a template file (.txt or .docx)")
    parser.add_argument("--data", help="Path to JSON file or a raw JSON string", default=None)
    parser.add_argument("--preview", action="store_true", help="Render preview when data is provided")
    parser.add_argument("--permissive", action="store_true", help="Disable policy checks")
    parser.add_argument("--custom-delims", nargs=4, metavar=("VAR_START", "VAR_END", "BLK_START", "BLK_END"),
                        help="Override Jinja delimiters, e.g., [[ ]] {% %}")
    parser.add_argument("--encoding", default="auto",
                        help="(Plain-text only) File encoding (default: auto). Examples: utf-8, cp1252, latin-1")
    parser.add_argument(
        "--scan",
        choices=["auto", "off", "segments", "lines"],
        default="auto",
        help=("On syntax failure, keep scanning for more issues. "
              "'segments' = DOCX paragraphs/cells, 'lines' = per-line for .txt, "
              "'auto' = segments for .docx, lines for .txt, 'off' = first error only.")
    )
    args = parser.parse_args(argv)

    env_kwargs = {}
    if args.custom_delims:
        var_s, var_e, blk_s, blk_e = args.custom_delims
        env_kwargs = dict(
            variable_start_string=var_s,
            variable_end_string=var_e,
            block_start_string=blk_s,
            block_end_string=blk_e,
        )

    policy = Policy.permissive() if args.permissive else Policy.sensible_defaults()
    validator = JinjaTemplateValidator(use_sandbox=True, policy=policy, env_kwargs=env_kwargs)
    validator._scan_mode = args.scan

    try:
        if args.template.lower().endswith(".docx"):
            if Document is None:
                raise RuntimeError("python-docx is required to read .docx (pip install python-docx)")
            template_str, segs = _extract_docx_with_map(args.template)
            validator._docx_segments = segs
        else:
            template_str, used_enc = _read_text_safely(args.template, args.encoding)
            validator._docx_segments = []
    except UnicodeDecodeError as e:
        tried = ", ".join(COMMON_ENCODINGS) if args.encoding == "auto" else args.encoding
        _print_json_error_and_exit(
            f"Could not decode '{args.template}'. Try --encoding cp1252 or latin-1. (Tried: {tried}). Error: {e}"
        )
    except RuntimeError as e:
        _print_json_error_and_exit(str(e), kind="docx")

    data = _load_data(args.data)
    result = validator.validate(template_str, sample_data=data, render_preview=args.preview)
    print(result.to_json())
    return 0 if result.ok else 2


if __name__ == "__main__":
    sys.exit(main())
