"""Open Template Scaffolds — reference MCP server (thin).

A dependency-light Model Context Protocol server that exposes the library's
markdown templates to MCP-capable AI clients. The tool surface is built across
phase B3; this scaffold provides the foundation plus `list_templates` as a
proof of life.
"""

import re

try:
    # FastMCP ships as its own package from version 2 on.
    from fastmcp import FastMCP
except ImportError:
    # mcp 1.x bundled it; mcp 2.0 removed it. Kept so an existing install
    # that predates the split keeps working without a reinstall.
    from mcp.server.fastmcp import FastMCP

from library import iter_standards, iter_templates, read_standard, validate_library

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

# A whole-phrase hit outranks any word-level hit inside the same tier.
_PHRASE = 1000

# Filler words carry no signal about what a template is for, and matching on them
# would rate every template against every sentence.
_STOPWORDS = frozenset(
    """a about all also an and any anything are as at be been being by can do does each every
    everything for from get got has have help how i in into is it its just like made make me
    more most much my need none nothing of on only or over please so some something that the
    their them then there these thing things this to up use using very want was what when
    where which who why will with would you your""".split()
)


def _terms(q: str) -> list[str]:
    """The searchable words of a query: three characters or more, no filler, no repeats.

    A trailing "s" is dropped so a plural finds its singular — "changes" becomes
    "change", which then matches "changed" as well.
    """
    seen, out = set(), []
    for w in re.findall(r"[a-z0-9]+", q):
        if len(w) < 3 or w in _STOPWORDS:
            continue
        if len(w) > 3 and w.endswith("s"):
            w = w[:-1]
        if w not in seen:
            seen.add(w)
            out.append(w)
    return out


def _word_hits(terms: list[str], haystack: str) -> list[str]:
    """Which terms start a word in `haystack`.

    Anchored at the start of a word, open at the end: "log" finds "logging" but
    not "catalog", and "old" does not find "scaffold". Matching an unanchored
    substring rates templates highly on accidental fragments.
    """
    return [t for t in terms if re.search(r"\b" + re.escape(t), haystack)]


def _relevance(q: str, front: dict, body: str):
    """Rate how query `q` (already lowercased) matches a template.

    The whole query is tried first, as a phrase. When that finds nothing, the
    query's individual words are tried, so a caller who describes a need in
    their own words is not told the library has nothing.

    Returns (relevance, matched_in, strength), or None when `q` matches nowhere.
    `strength` orders results inside a tier; it is not reported to the caller.

      - Likely   -> the phrase is in the template's identity (name / title /
                    domain), or every word of the query was found and at least
                    one of them in the identity.
      - Possible -> a whole-word phrase hit in the Intent, or at least half the
                    query's words found (the template is about it).
      - Unlikely -> an incidental substring of the phrase in the Intent (e.g. a
                    cross-reference to another template), or a minority of the
                    query's words.
    """
    identity = " ".join(str(front.get(k, "")) for k in ("template", "title", "domain")).lower()
    if q in identity:
        return "Likely", "name/title/domain", _PHRASE
    intent = _intent(body).lower()
    if re.search(r"\b" + re.escape(q) + r"\b", intent):
        return "Possible", "intent", _PHRASE

    # Word-level pass, so a caller who describes the need in their own words is
    # not told the library has nothing.
    terms = _terms(q)
    if terms:
        in_id = _word_hits(terms, identity)
        in_intent = _word_hits(terms, intent)
        matched = set(in_id) | set(in_intent)
        if matched:
            where = "name/title/domain + intent" if (in_id and in_intent) else (
                "name/title/domain" if in_id else "intent")
            note = f"{where} ({len(matched)} of {len(terms)} words)"
            if len(matched) == len(terms) and in_id:
                return "Likely", note, len(matched)
            if len(matched) * 2 >= len(terms):
                return "Possible", note, len(matched)
            return "Unlikely", note, len(matched)

    if q in intent:
        return "Unlikely", "intent (incidental)", 0
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
            relevance, matched_in, strength = rated
        else:
            relevance, matched_in, strength = "Likely", ("filter" if (dom or typ) else "all"), _PHRASE
        results.append(
            (_RANK[relevance], -strength,
             {**_meta(front), "relevance": relevance, "matched_in": matched_in})
        )
    # Strongest tier first, and inside a tier the entry that matched most of the query.
    results.sort(key=lambda r: (r[0], r[1]))
    return [r[2] for r in results]


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
def get_standards(standard: str = "") -> dict:
    """Return the active standards layer on its own — no template required.

    Plain version: this hands the AI the house rules by themselves. It exists
    for the from-scratch path: when no template fits and the developer approves
    a from-scratch design, that design must still follow the standards layer —
    naming, audit columns, error handling, query style, form conventions. This
    tool loads the layer without going through a template.

    With no argument it returns every standards file; pass a `standard` name
    (e.g. "naming-conventions") for just that one. Any README in `standards/`
    is skipped — it maps the folder; it isn't a standard. Read-only: this
    builds nothing and changes nothing; the never-build-before-approval
    boundary is unaffected.

    Returns `{standards: [{name, content}], count}`. Raises ValueError when a
    given name matches no standards file (the message lists what exists).
    """
    name = standard.strip().lower()
    entries = [{"name": n, "content": c} for n, c in iter_standards()]
    if name:
        matched = [e for e in entries if e["name"].lower() == name]
        if not matched:
            raise ValueError(
                f"No standards file named '{standard}'. "
                f"Available: {', '.join(e['name'] for e in entries)}."
            )
        entries = matched
    return {"standards": entries, "count": len(entries)}


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
