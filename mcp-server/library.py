"""Read the Open Template Scaffolds library markdown from the repo.

This module is the foundation the MCP tools build on: it locates the library's
content directories and splits a template's YAML front-matter from its body.
The server ships inside the library (`<library-root>/mcp-server/`), so the
library root is simply this file's parent's parent.
"""

import re
from pathlib import Path

import yaml

# This server ships in <library-root>/mcp-server/, so the root is its parent.
LIBRARY_ROOT = Path(__file__).resolve().parents[1]
TEMPLATES_DIR = LIBRARY_ROOT / "templates"
STANDARDS_DIR = LIBRARY_ROOT / "standards"


def split_front_matter(text: str) -> tuple[dict, str]:
    """Split a template's YAML front-matter from its markdown body.

    Returns (front_matter_dict, body). A file without front-matter yields
    ({}, original_text).
    """
    if text.startswith("---"):
        # "---\n<yaml>\n---\n<body>" -> ["", "<yaml>", "<body>"]
        _, front, body = text.split("---", 2)
        return yaml.safe_load(front) or {}, body.lstrip("\n")
    return {}, text


def iter_templates():
    """Yield (path, front_matter, body) for each domain template.

    Skips infrastructure files (prefixed `_`, e.g. `_template-schema.md`,
    `_materialization.md`) and any `README.md`, which are not domain templates.
    """
    for path in sorted(TEMPLATES_DIR.rglob("*.md")):
        if path.name.startswith("_") or path.name.lower() == "readme.md":
            continue
        front, body = split_front_matter(path.read_text(encoding="utf-8"))
        yield path, front, body


def read_standard(name: str) -> str | None:
    """Return the text of a standards-layer file (`standards/<name>.md`).

    Returns None when no such file exists, so callers can report an unresolved
    `standards_layer` entry rather than fail silently.
    """
    path = STANDARDS_DIR / f"{name}.md"
    if path.is_file():
        return path.read_text(encoding="utf-8")
    return None


def iter_standards():
    """Yield (name, content) for every standards-layer file, sorted by name.

    Skips any README.md — that file maps the folder; it isn't a standard —
    matching iter_templates' convention for infrastructure files.
    """
    for path in sorted(STANDARDS_DIR.glob("*.md")):
        if path.name.lower() == "readme.md":
            continue
        yield path.stem, path.read_text(encoding="utf-8")


# --- validate(): format-only rules from templates/_template-schema.md ---
# Each check returns a short "RULE: message" string; an empty list means the
# template is well-formed. No host database is ever opened here.

_TYPE_ENUM = {"table-schema", "vba-scaffold", "form-spec", "spec"}
_STATUS_ENUM = {"draft", "review", "stable"}
_STANDARDS_VALUES = {"audit-columns", "naming-conventions", "error-handling",
                     "query-style", "form-conventions", "design-principles"}
_AUDIT_COLUMNS = {"addedby", "addedon", "modifiedby", "modifiedon"}
_ACCESS_SCALAR_TYPES = {"autonumber", "long", "integer", "byte", "single", "double",
                        "currency", "memo", "date/time", "boolean", "guid"}
_TEXT_TYPE = re.compile(r"^text\(\d+\)$", re.I)
_REQUIRED_FM = ("template", "title", "domain", "type", "version", "status", "standards_layer")
_FIELD_TABLE_HEADER = ["field", "type", "key/req", "purpose&rules"]
_FK_RE = re.compile(r"FK\s*(?:→|->)\s*`?([A-Za-z0-9_]+)`?")


def _h2_sections(body):
    """Split a body into (heading, text) pairs at each level-2 (`## `) heading."""
    sections, head, buf = [], None, []
    for line in body.splitlines():
        s = line.strip()
        if s.startswith("## ") and not s.startswith("### "):
            if head is not None:
                sections.append((head, "\n".join(buf)))
            head, buf = s[3:].strip(), []
        elif head is not None:
            buf.append(line)
    if head is not None:
        sections.append((head, "\n".join(buf)))
    return sections


def _has_section(sections, name):
    return any(h.lower().startswith(name.lower()) for h, _ in sections)


def _section_text(sections, name):
    for h, t in sections:
        if h.lower().startswith(name.lower()):
            return t
    return None


def _h3_headings(text):
    return [s.strip()[4:].strip() for s in (text or "").splitlines()
            if s.strip().startswith("### ") and not s.strip().startswith("#### ")]


def _md_tables(text):
    """Yield (header_cells, [row_cells, ...]) for each GitHub-style table."""
    lines = (text or "").splitlines()
    i = 0
    while i < len(lines):
        row = lines[i].strip()
        sep = lines[i + 1].strip() if i + 1 < len(lines) else ""
        if row.startswith("|") and re.match(r"^\|[\s:|-]+\|$", sep):
            header = [c.strip() for c in row.strip("|").split("|")]
            rows, j = [], i + 2
            while j < len(lines) and lines[j].strip().startswith("|"):
                rows.append([c.strip() for c in lines[j].strip().strip("|").split("|")])
                j += 1
            yield header, rows
            i = j
        else:
            i += 1


def _proc_base(x):
    """Base procedure/name token: strip backticks, drop any signature and trailing words."""
    head = str(x).replace("`", "").split("(")[0].strip()
    return head.split()[0] if head else ""


def _slug_ok(s):
    return bool(re.match(r"^[a-z0-9]+(-[a-z0-9]+)*$", str(s or "")))


def validate_template(front: dict, body: str, stem: str) -> list[str]:
    """Check one template against the canonical format — structure only.

    `stem` is the filename without its extension. Returns human-readable
    "RULE: message" strings; an empty list means well-formed. The caller skips
    `type: spec` infrastructure files. No host database is opened; this speaks
    to well-formedness, never to whether the template fits a given project.
    """
    errors = []
    typ = str(front.get("type", "")).strip()

    # ---- Front-matter (spec section 2) ----
    for key in _REQUIRED_FM:
        val = front.get(key)
        if val is None or (isinstance(val, str) and not val.strip()) or (isinstance(val, list) and not val):
            errors.append(f"FM1: missing/empty required front-matter key '{key}'")
    if typ and typ not in _TYPE_ENUM:
        errors.append(f"FM1: type '{typ}' not one of {sorted(_TYPE_ENUM)}")
    status = str(front.get("status", "")).strip()
    if status and status not in _STATUS_ENUM:
        errors.append(f"FM1: status '{status}' not one of {sorted(_STATUS_ENUM)}")
    slug = str(front.get("template", ""))
    if slug:
        if not _slug_ok(slug):
            errors.append(f"FM2: template slug '{slug}' is not kebab-case")
        elif not (slug == stem or slug.endswith("-" + stem)):
            errors.append(f"FM2: template slug '{slug}' does not end with filename stem '{stem}'")
    if front.get("extends") and not front.get("requires_tables"):
        errors.append("FM3: 'extends' is set but 'requires_tables' is empty")
    for key in ("requires_fields", "seeds"):
        for entry in front.get(key) or []:
            if "." not in str(entry):
                errors.append(f"FM5: {key} entry '{entry}' is not a well-formed Table.Field")
    for sl in front.get("standards_layer") or []:
        if sl not in _STANDARDS_VALUES:
            errors.append(f"FM: unrecognized standards_layer value '{sl}'")
    for ha in front.get("house_assumptions") or []:
        parts = re.split(r"—| - ", str(ha), maxsplit=1)
        if len(parts) < 2 or not parts[1].strip():
            errors.append(f"FM6: house_assumptions entry is not 'Target - rationale': '{str(ha)[:40]}'")
            continue
        token = parts[0].strip().strip("`").split(".")[0].split()[0]
        if token and token.lower() not in body.lower():
            errors.append(f"FM6: house_assumptions Target '{token}' is not named in the template body")

    # ---- Common core (spec section 3) ----
    sections = _h2_sections(body)
    for sec in ("Intent", "Standards Layer", "Extra Options"):
        if not _has_section(sections, sec):
            errors.append(f"CORE: missing required section '## {sec}'")
    h1 = next((l.strip()[2:].strip() for l in body.splitlines()
               if l.strip().startswith("# ") and not l.strip().startswith("## ")), None)
    if front.get("title") and h1 and h1 != str(front["title"]).strip():
        errors.append(f"CORE: H1 '{h1}' does not match title '{front['title']}'")

    # ---- Type-specific ----
    if typ == "table-schema":
        errors += _validate_table_schema(front, sections)
    elif typ == "vba-scaffold":
        errors += _validate_vba_scaffold(front, sections)
    elif typ == "form-spec":
        errors += _validate_form_spec(front, sections)
    return errors


def _validate_table_schema(front, sections):
    errors = []
    if not front.get("new_tables"):
        errors.append("FM: type 'table-schema' requires non-empty 'new_tables'")
    for sec in ("Entities", "Relationships", "Business Rules"):
        if not _has_section(sections, sec):
            errors.append(f"TS: missing required section '## {sec}'")
    if front.get("extends") and not _has_section(sections, "Prerequisites"):
        errors.append("TS: 'extends' is set but there is no '## Prerequisites' section")

    ent = _section_text(sections, "Entities") or ""
    declared = [str(t) for t in front.get("new_tables") or []]
    for t in declared:                                     # every declared table is documented
        if t.lower() not in ent.lower():
            errors.append(f"TS1: new_tables '{t}' is not documented under '## Entities'")
    for h in _h3_headings(ent):                             # every '### entity' is declared
        name = h.strip().strip("`")
        if " " in name:                                    # a descriptive sub-heading, not an entity
            continue
        if name and name not in declared:
            errors.append(f"TS1: '### {name}' under '## Entities' is not in new_tables")

    known = set(declared) | {str(t) for t in front.get("requires_tables") or []}
    for header, rows in _md_tables(ent):
        norm = [c.lower().replace(" ", "") for c in header]
        if norm[:1] != ["field"]:
            continue                                       # not a field-spec table
        if norm != _FIELD_TABLE_HEADER:
            errors.append(f"TS2: field-table header {header} != Field | Type | Key / Req | Purpose & rules")
        for r in rows:
            if len(r) < 3:
                continue
            if r[0].strip("` ").lower() in _AUDIT_COLUMNS:
                errors.append(f"TS5: audit column '{r[0]}' is in a field table (belongs to the standards layer)")
            tval = r[1].strip().split()[0].lower().rstrip(".,") if r[1].strip() else ""
            if tval and tval not in _ACCESS_SCALAR_TYPES and not _TEXT_TYPE.match(tval):
                errors.append(f"TS2: field '{r[0]}' has unknown type '{r[1]}'")
            m = _FK_RE.search(r[2])
            if m and m.group(1) not in known:
                errors.append(f"TS3: FK target '{m.group(1)}' (field {r[0]}) resolves to no known table")
    return errors


def _validate_vba_scaffold(front, sections):
    errors = []
    if not str(front.get("target_module", "")).strip():
        errors.append("VS1: type 'vba-scaffold' requires non-empty 'target_module'")
    if not front.get("new_procedures"):
        errors.append("FM: type 'vba-scaffold' requires non-empty 'new_procedures'")
    if not _has_section(sections, "Procedures"):
        errors.append("VS: missing required section '## Procedures'")
    proc = _section_text(sections, "Procedures") or ""
    declared = [_proc_base(p) for p in front.get("new_procedures") or []]
    documented = [_proc_base(h) for h in _h3_headings(proc)]
    for p in declared:
        if p and p not in documented:
            errors.append(f"VS2: new_procedures '{p}' has no matching '### {p}' under '## Procedures'")
    for d in documented:
        if d and d not in declared:
            errors.append(f"VS2: '### {d}' under '## Procedures' is not in new_procedures")
    for block in re.split(r"^### ", proc, flags=re.M)[1:]:
        name = _proc_base(block.splitlines()[0]) if block.splitlines() else "?"
        if "```vba" not in block.lower():
            errors.append(f"VS3: procedure '### {name}' has no fenced vba block")
    if "error-handling" not in [str(x) for x in front.get("standards_layer") or []]:
        errors.append("VS4: vba-scaffold standards_layer must include 'error-handling'")
    if front.get("implements") and not _slug_ok(front["implements"]):
        errors.append(f"VS5: implements '{front['implements']}' is not a well-formed slug")
    return errors


def _validate_form_spec(front, sections):
    errors = []
    if not str(front.get("record_source", "")).strip():
        errors.append("FS1: type 'form-spec' requires non-empty 'record_source'")
    if not front.get("new_forms"):
        errors.append("FM: type 'form-spec' requires non-empty 'new_forms'")
    for sec in ("Layout", "Features", "Materialization"):
        if not _has_section(sections, sec):
            errors.append(f"FS: missing required section '## {sec}'")
    layout = _section_text(sections, "Layout") or ""
    ctrl_tables = [(h, rows) for h, rows in _md_tables(layout)
                   if h[:1] and h[0].lower() == "control"]
    if not ctrl_tables:
        errors.append("FS3: '## Layout' has no control-inventory table (Control | Type | Bound to | Notes)")
    subform_rows = [r for _, rows in ctrl_tables for r in rows if len(r) > 1 and "subform" in r[1].lower()]
    if len(front.get("new_forms") or []) > 1 and not subform_rows:
        errors.append("FS2: multiple new_forms declared but no 'Subform' control appears in '## Layout'")
    if "form-conventions" not in [str(x) for x in front.get("standards_layer") or []]:
        errors.append("FS4: form-spec standards_layer must include 'form-conventions'")
    if front.get("implements") and not _slug_ok(front["implements"]):
        errors.append(f"FS5: implements '{front['implements']}' is not a well-formed slug")
    return errors


def validate_library(template: str = "") -> dict:
    """Validate one template by id, or the whole library — structure only.

    This is the single shared code path behind both the MCP `validate` tool and
    the CI gate (`run_validate.py`): per-template `validate_template()` checks
    plus the library-wide slug-uniqueness rule. Infrastructure files
    (`type: spec`) are skipped. Returns `{ok, checked, results}`; raises
    ValueError when a given id matches nothing.
    """
    tid = template.strip().lower()
    results = []
    for path, front, body in iter_templates():
        if str(front.get("type", "")).lower() == "spec":
            continue
        if tid and str(front.get("template", "")).lower() != tid:
            continue
        errs = validate_template(front, body, path.stem)
        results.append({"template": front.get("template"), "ok": not errs, "errors": errs})
    if tid and not results:
        raise ValueError(
            f"No template with id '{template}'. "
            "Use list_templates or search_templates to find valid ids."
        )
    counts = {}
    for r in results:
        counts[r["template"]] = counts.get(r["template"], 0) + 1
    for r in results:
        if counts[r["template"]] > 1:
            r["errors"].append(f"FM2: template slug '{r['template']}' is not unique in the library")
            r["ok"] = False
    return {"ok": all(r["ok"] for r in results), "checked": len(results), "results": results}
