# Jinja Template Validator

A small validator for Jinja2 templates designed for non-technical authors. The validator supports both plain-text templates and **.docx** Word files, and can raise errors with Word-aware locations (e.g., “Body ▸ Paragraph 12”), plus optional scanning to list multiple issues even when the full template doesn’t parse.

The purpose of this validator is to help the user to quickly find possible errors early on in their Jinja2 syntax before uploading the template file into a system (e.g. Kaavapino) to be processed.

The validator also supports a **variable catalog whitelist** (from Excel/CSV or a typed list). This enables flagging template variables that aren’t in the official data dictionary.

The project includes three main components:

- **Backend API (Flask):** `validator_api.py`  
- **Validator library/CLI:** `jinja_validator.py`  
- **Simple Web UI:** `index.html`, `styles.css`, `app.js`

---

## Features

- ✅ **Syntax validation** with small codeframes
- ✅ **StrictUndefined**: catches missing keys early when rendering preview
- ✅ **Policy allowlists** (filters, tests, calls, tags) with sensible defaults
- ✅ **DOCX awareness**: maps Jinja errors and variable usages back to **Word** paragraphs, tables, headers/footers (with short “anchors” to find them)
- ✅ **Scan modes** to aggregate multiple issues:
  - `segments` (per paragraph/cell for .docx)
  - `lines` (per line for text files)
  - `auto` (smart choice)
- ✅ **Scan hints**: per-segment/line “unclosed block” messages (not fatal)
- ✅ **Sample data** (optional): shows **missing variables** only if you provide JSON data
- ✅ **Variable catalog whitelist** (optional): flag variables **not present** in Excel/CSV/typed list
- ✅ **Custom delimiters** (if your system doesn’t use `{{ }}`, `{% %}`)
- ✅ **Encoding auto-detection** (`utf-8`, `utf-8-sig`, `cp1252`, `latin-1`, + optional `charset-normalizer`)
- ✅ **CORS** enabled (lock down in prod)
- ✅ **Whitelisted unknown variable names** Set vars never to be flagged as unknown when validating var names against a catalogue.

---

## Project Structure
```
/
├── jinja_validator.py
├── validator_api.py
├── requirements.txt
├── openapi.yaml
└── web-ui/
    ├── index.html
    ├── styles.css
    └── app.js
```


---

## Quick Start

### 1) Setup Python venv
`python -m venv .venv`
#### Windows:
`.venv\Scripts\activate`
#### macOS/Linux:
`source .venv/bin/activate`

### 2) Install requirements
`pip install -r requirements.txt`

### 3) Run the API
`python validator_api.py`

### 4) Open the Web UI
- Double-click index.html or serve it (e.g., python -m http.server 5500 then open http://localhost:5500).
- If you open via file://, set API base to your API (e.g., http://127.0.0.1:5001) in the field at the top of the page.

---

## Using the Web UI
1. Upload a .docx or text file, or paste template text.
2. (Optional) Paste sample JSON data.
   - If you **don’t** provide sample data, **“Missing variable”** checks are not raised.
3. (Optional) Add a variable catalog:
   - Upload an **Excel** (.xlsx) or **CSV** with your allowed variables (specify sheet and column), **or**
   - Paste a list/JSON array of allowed names.
4. (Optional) Adjust mode:
   - **Scan: auto** (default), segments, lines, off
   - **Preview**: renders the template (fails fast on missing keys)
   - **Permissive**: disables policy checks
   - **Encoding**: for text files if you see decoding errors
   - **Custom delimiters**
5. Click Validate.

You’ll see:
   - **Summary chips** (syntax/render, scan hints, missing vars, unknown vs catalog, policy)
   - **Issues list** with codeframes, Word locations/anchors (for .docx), suggestions
   - **Usage**: expected variables, used tags/filters/tests/calls, unused data paths

**Scan hints** (“unclosed block”) show up when a paragraph/line contains only an opening tag (e.g., {% if %}) but the closing tag is in another paragraph/line. They are not fatal and help you locate control structures

## CLI (optional)
You can run the validator directly on files:
```
python jinja_validator.py path/to/template.docx --scan auto
python jinja_validator.py path/to/template.txt --encoding cp1252 --scan lines
python jinja_validator.py path/to/template.txt --data '{"name":"Kaavahanke"}' --preview
python jinja_validator.py path/to/template.docx --custom-delims "[[ ]]" "{% %}"   # example (see help)
```
`--help` lists all options:
- `-scan` = auto | off | segments | lines
- `-preview` to render (only if you also pass --data …)
- `-permissive` to disable policy checks
- `-encoding` for text files
- `-custom-delims` VAR_START VAR_END BLK_START BLK_END

**Note**: When no sample data is passed, “missing variables” are not reported.

---

## API Reference

### Health
To check if the service is up and running.
```
GET /health
→ 200 {"status":"ok"}
```

### Validate
#### JSON body (no file)
**http**
```
POST /validate
Content-Type: application/json
```
**payload json example**
```
{
  "template": "{% if is_preview %}Hello {{ projektin_nimi }}{% endif %}",
  "scan": "auto",                    // "auto" | "off" | "segments" | "lines"
  "preview": false,                  // renders only if "data" is provided
  "permissive": false,               // disables policy checks if true
  "custom_delims": {                 // optional; only if you use non-default delimiters
    "var_start": "{{",
    "var_end": "}}",
    "block_start": "{%",
    "block_end": "%}"
  },
  "data": {                          // optional sample data; omit to skip "missing variable" checks
    "is_preview": true,
    "projektin_nimi": "Kaavahanke"
  },
  "varlist": [                       // optional variable catalog (whitelist) as an array
    "is_preview",
    "projektin_nimi",
    "kaavanumero"
  ],
  "varlist_mode": "head"             // "head" (top-level name) or "fullpath"
}
```

#### Multipart form (file)
**http**
```
POST /validate
Content-Type: multipart/form-data
```


#### Response
- 200 OK if ok=true
- 422 if issues that fail the run (e.g., syntax errors, policy, unknown catalog vars, preview render errors)

**Example json response**
```
{
  "ok": false,
  "errors": [
    {
      "type": "syntax",
      "message": "...",
      "line": 12,
      "excerpt": "  10 | ...\n> 12 | {% if ...\n",
      "docx_location": "Body ▸ Paragraph 5",
      "anchor": "Short snippet to find it",
      "hint": "…optional guidance…"
    }
  ],
  "expected_variables": ["foo", "bar.baz"],
  "used_filters": ["upper", "default"],
  "used_tests": [],
  "used_calls": [],
  "used_tags": ["if", "for"],
  "missing_from_data": [
    {
      "path": "bar.baz",
      "suggestion": "bar.buzz",
      "occurs_at": [
        { "line": 12, "docx_location": "Body ▸ Paragraph 5", "anchor": "..." }
      ]
    }
  ],
  "unused_data_paths": ["payload.extra"],
  "preview": null,
  "unknown_variables": [
    {
      "name": "foo.bar",
      "catalog_key": "foo",
      "suggestion": "foe",
      "occurs_at": [{ "line": 7, "docx_location": "Body ▸ Paragraph 3" }]
    }
  ]
}
```

**Note**: See **openapi.yaml** for more detailed reference.

---

## Security & Privacy
- The API does not persist uploaded files. Temporary files created during processing are cleaned up.
- CORS: in development, the API may allow all origins. In production, restrict CORS to your UI’s domain:
```
# validator_api.py (example)
CORS(app, resources={r"/validate": {"origins": ["https://your-ui.example"]}})
```
---

## Troubleshooting
- **CORS error in browser console**
  - Restrict or adjust CORS(app, …) to include your UI origin.
- **UnicodeDecodeError (text templates)**
  - Pick a different Encoding (e.g., cp1252 or latin-1) in the UI, or pass --encoding in CLI.
- **“scan hints” everywhere**
  - Means segments/lines open blocks without closing within the same segment/line.
  - Use Scan: off to validate the full template, **or**
  - Keep opening/closing tags in the same paragraph/line if you want per-segment scanning to pass.
- **Missing variable vs. unknown variable**
  - **Missing variable**: referenced in the template but not present in your sample data (only checked when sample data is provided).
  - **Unknown variable**: referenced in the template but not in your official catalog (Excel/CSV/list) when a catalog is provided.



