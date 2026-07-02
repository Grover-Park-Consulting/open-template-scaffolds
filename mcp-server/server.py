"""Open Template Scaffolds — reference MCP server (thin).

A dependency-light Model Context Protocol server that exposes the library's
markdown templates to MCP-capable AI clients. The tool surface is built across
phase B3; this scaffold provides the foundation plus `list_templates` as a
proof of life.
"""

import re

from mcp.server.fastmcp import FastMCP

from library import iter_templates, read_standard, validate_library

mcp = FastMCP("open-template-scaffolds")

_META_KEYS = ("template", "title", "domain", "type", "status")


def _meta(front: dict) -> dict:
    """Project a template's front-matter to the metadata clients see."""
    return {key: front.get(key) for key in _META_KEYS}


def _intent(body: str) -> str:
    """Return the text of a template's '## Intent' section (empty if absent)."""
    out, capturing = [], False
    for line in body.splitlines():
        if line.lstrip().lower().startswith("## intent"):
            capturing = True
            continue
        if capturing and line.lstrip().startswith("## "):
            break
        if capturing:
            out.append(line)
    return " ".join(out).strip()


@mcp.tool()
def list_templates() -> list[dict]:
    """List the library's templates with their key metadata.

    Returns one entry per domain template (infrastructure files excluded).
    """
    return [_meta(front) for _, front, _ in iter_templates()]


_RANK = {"Likely": 0, "Possible": 1, "Unlikely": 2}


def _relevance(q: str, front: dict, body: str):
    """Rate how query `q` (already lowercased) matches a template.

    Returns (relevance, matched_in), or None when `q` matches nowhere:
      - Likely   -> `q` is in the template's identity (name / title / domain).
      - Possible -> a whole-word hit in the Intent (the template is about it).
      - Unlikely -> only an incidental substring in the Intent (e.g. a
                    cross-reference to another template).
    """
    identity = " ".join(str(front.get(k, "")) for k in ("template", "title", "domain")).lower()
    if q in identity:
        return "Likely", "name/title/domain"
    intent = _intent(body).lower()
    if re.search(r"\b" + re.escape(q) + r"\b", intent):
        return "Possible", "intent"
    if q in intent:
        return "Unlikely", "intent (incidental)"
    return None


@mcp.tool()
def search_templates(query: str = "", domain: str = "", type: str = "") -> list[dict]:
    """Find templates by free-text query, domain, and/or type.

    Filters combine with AND: `domain` and `type` match the front-matter
    case-insensitively (exact); `query` is matched against the template's name,
    title, domain, and Intent.

    Every result carries a `relevance` rating — **Likely**, **Possible**, or
    **Unlikely** — plus a `matched_in` note saying where it hit, so the caller
    can weigh the weaker candidates rather than have them hidden or silently
    acted on. Results come back strongest-first. This tool only surfaces and
    rates candidates; whether any of them actually fit is the developer's call,
    not the tool's.

    With no free-text query, results are exact filter matches (rated Likely).
    When nothing matches at all, returns an empty list.
    """
    q = query.strip().lower()
    dom = domain.strip().lower()
    typ = type.strip().lower()
    results = []
    for _, front, body in iter_templates():
        if dom and dom != str(front.get("domain", "")).lower():
            continue
        if typ and typ != str(front.get("type", "")).lower():
            continue
        if q:
            rated = _relevance(q, front, body)
            if rated is None:
                continue
            relevance, matched_in = rated
        else:
            relevance, matched_in = "Likely", ("filter" if (dom or typ) else "all")
        results.append({**_meta(front), "relevance": relevance, "matched_in": matched_in})
    results.sort(key=lambda r: _RANK[r["relevance"]])
    return results


@mcp.tool()
def get_template(template: str) -> dict:
    """Return a chosen template composed with its active standards layer.

    Looks up `template` by its front-matter id (case-insensitive) and returns
    the full template — front-matter plus body — together with the text of every
    standards file named in the template's `standards_layer`.

    The template and the standards are returned **separately, not merged**: the
    template is the design; the standards are the swappable house layer applied
    on top. Keeping them distinct lets the adopter see and customize each — a
    standard can be swapped without touching the template.

    This is a read/compose tool. It produces the material the AI uses to draft a
    *proposed* schema for the developer to approve; it builds nothing itself.

    Returns a dict with the `_meta` keys, the full `front_matter`, the `body`,
    and a `standards` list of `{name, content}`. If a `standards_layer` entry
    has no matching file, its name is reported under `standards_missing` rather
    than silently dropped. Raises ValueError when no template has the given id
    (use `list_templates` or `search_templates` to find valid ids).
    """
    tid = template.strip().lower()
    for _, front, body in iter_templates():
        if str(front.get("template", "")).lower() == tid:
            standards, missing = [], []
            for name in front.get("standards_layer") or []:
                content = read_standard(name)
                if content is None:
                    missing.append(name)
                else:
                    standards.append({"name": name, "content": content})
            result = {**_meta(front), "front_matter": front, "body": body, "standards": standards}
            if missing:
                result["standards_missing"] = missing
            return result
    raise ValueError(
        f"No template with id '{template}'. "
        "Use list_templates or search_templates to find valid ids."
    )


@mcp.tool()
def validate(template: str = "") -> dict:
    """Check whether a template is *built correctly* — structure only.

    Plain version: this is a "is this template put together right?" checker, like
    a spell-check for a blueprint. It confirms the required sections are present,
    the tables named at the top match the ones described below, a form says which
    data it edits, and nothing is mislabeled or missing.

    Two limits, on purpose: it never opens your real database (that's
    `check_compatibility`), and it cannot tell you whether a template is the
    *right* one for your project — a template can be perfectly well-formed and
    still be the wrong fit. Passing means *well-built*, never *suitable for you*;
    that judgment stays yours.

    With a `template` id it checks just that one; with no argument it checks the
    whole library (the CI-gate mode). Returns `{ok, checked, results}` where each
    result is `{template, ok, errors}` and `errors` names the rule that failed.
    Infrastructure files (`type: spec`) are skipped. Raises ValueError if a given
    id matches nothing.
    """
    return validate_library(template)


@mcp.tool()
def check_compatibility(template: str, db_path: str) -> dict:
    """Check whether a real database already has what a template builds on.

    Plain version: does your actual Access database already contain the tables
    and columns this template expects to find (its `requires_tables` /
    `requires_fields`)? This is the one tool that opens your real database — and
    it only *looks*: it reads the list of table and column names, never your
    data, and never changes anything.

    Two limits, on purpose: "compatible" means the expected pieces are *present*,
    never that this is the *right* template for your project (that judgment stays
    yours). And it is Windows/Access-only — it needs the Microsoft Access ODBC
    driver plus `pyodbc` (the optional `requirements-compat.txt`). When those are
    missing, or the file can't be opened, it returns `available: false` with a
    reason instead of failing — the rest of the server keeps working.

    Looks `template` up by its front-matter id (case-insensitive). A template
    that declares no `requires_*` has no host dependencies, so it is trivially
    compatible. Returns `{available, template, db_path, tables, fields, ok}`
    where `ok` is true only when the check ran and nothing is missing. Raises
    ValueError when no template has the given id.
    """
    from compat import check_db

    tid = template.strip().lower()
    for _, front, _ in iter_templates():
        if str(front.get("template", "")).lower() == tid:
            result = check_db(db_path, front.get("requires_tables"), front.get("requires_fields"))
            result["template"] = front.get("template")
            result["db_path"] = db_path
            if result.get("available"):
                result["ok"] = not (result["tables"]["missing"]
                                    or result["fields"]["missing"]
                                    or result["fields"]["table_absent"])
            else:
                result["ok"] = False
            return result
    raise ValueError(
        f"No template with id '{template}'. "
        "Use list_templates or search_templates to find valid ids."
    )


if __name__ == "__main__":
    mcp.run()  # stdio transport
