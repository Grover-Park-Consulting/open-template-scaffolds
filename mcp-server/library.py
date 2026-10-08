"""Read the Open Template Scaffolds library markdown from the repo.

This module is the foundation the MCP tools build on: it locates the library's
content directories and splits a template's YAML front-matter from its body.
The server ships inside the library (`<library-root>/mcp-server/`), so the
library root is simply this file's parent's parent.
"""

import hashlib
import json
import re
from pathlib import Path

import yaml

# This server ships in <library-root>/mcp-server/, so the root is its parent.
LIBRARY_ROOT = Path(__file__).resolve().parents[1]
TEMPLATES_DIR = LIBRARY_ROOT / "templates"
STANDARDS_DIR = LIBRARY_ROOT / "standards"
# A template's step files sit in `<template-id>.steps/` beside the master file.
STEPS_SUFFIX = ".steps"


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
    `_materialization.md`), any `README.md`, and step files: those live in a
    `<template-id>.steps/` folder beside their master and belong to it.
    """
    for path in sorted(TEMPLATES_DIR.rglob("*.md")):
        if path.name.startswith("_") or path.name.lower() == "readme.md":
            continue
        if any(p.name.endswith(STEPS_SUFFIX) for p in path.relative_to(TEMPLATES_DIR).parents):
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


MATERIALIZATION = TEMPLATES_DIR / "_materialization.md"
METHOD = TEMPLATES_DIR / "_method.md"


def read_platform_facts() -> tuple[dict[str, dict], list[str]]:
    """Map each fact id in _materialization.md to its section (see _read_marked)."""
    return _read_marked(MATERIALIZATION, "fact")


def read_method() -> tuple[dict[str, dict], list[str]]:
    """Map each method id in _method.md to its section (see _read_marked)."""
    return _read_marked(METHOD, "method")


# Method every run receives, then what the template's features and the route add.
_METHOD_EVERY_RUN = ("run-opening", "design-review", "house-assumptions-and-warnings",
                     "between-questions", "checklist-rule", "build-record")
_METHOD_DESIGN = ("design-only-handover",)
_METHOD_BUILD = ("build-route", "access-gate", "quiet-build", "runbook", "build-records-accumulate")
ROUTES = ("", "design", "build")


def _front_method(front: dict) -> list[str]:
    """Method a template's front matter calls for; delivered with the master."""
    ids = []
    if front.get("related"):
        ids.append("related-after-finish")
    typ = str(front.get("type", ""))
    if typ == "outcome-first":
        ids.append("explore-options")
    if typ == "vba-scaffold":
        ids.append("staged-procedures")
    return ids


def _body_method(body: str) -> list[str]:
    """Method a piece of template text calls for; delivered with the part that holds it."""
    return ["wizard"] if re.search(r"^## Wizard\s*$", body, re.M) else []


def method_for(front: dict | None, body: str, route: str = "") -> list[str]:
    """The method ids a run needs, chosen from what the template contains and the route.

    The server chooses; no template declares method, so none can forget it. `front`
    is None for a run with no template (the from-scratch path). An empty route
    returns both routes' method.
    """
    ids = list(_METHOD_EVERY_RUN)
    if front is not None:
        ids += _front_method(front) + _body_method(body)
    if route in ("", "design"):
        ids += _METHOD_DESIGN
    if route in ("", "build"):
        ids += _METHOD_BUILD
    return ids


def _read_marked(path: Path, kind: str) -> tuple[dict[str, dict], list[str]]:
    """Map each `<!-- kind: id -->` marker in `path` to its section.

    A section is marked by the marker line under its heading and runs from that
    heading to the line before the next heading of the same or higher level.
    Lines inside code fences are never read as headings. A marker may end in
    `route: build` (`<!-- fact: dao-table-build route: build -->`): the section is
    then needed only on the build route, and its `route` is "build" ("" otherwise).
    Returns (sections, duplicate_ids); a duplicated id keeps its first section.
    """
    if not path.is_file():
        return {}, []
    mark = re.compile(r"^<!--\s*" + kind + r":\s*([a-z0-9-]+)(?:\s+route:\s*(build))?\s*-->$")
    lines = path.read_text(encoding="utf-8").splitlines()
    heads, fence = [], False
    for i, ln in enumerate(lines):
        if ln.lstrip().startswith("```"):
            fence = not fence
        elif not fence and ln.startswith("#"):
            heads.append((i, len(ln) - len(ln.lstrip("#"))))
    facts, dupes = {}, []
    for i, ln in enumerate(lines):
        m = mark.match(ln.strip())
        if not m:
            continue
        fid = m.group(1)
        above = [h for h in heads if h[0] < i]
        if not above:
            continue
        h, level = above[-1]
        end = next((x[0] for x in heads if x[0] > i and x[1] <= level), len(lines))
        if fid in facts:
            dupes.append(fid)
            continue
        facts[fid] = {"id": fid, "heading": lines[h].lstrip("#").strip(),
                      "content": "\n".join(lines[h:end]).strip(), "route": m.group(2) or ""}
    return facts, dupes


# --- The run plan: a template delivered as parts, one per call --------------
#
# Claude Code saves any tool answer over 50,000 characters to a file instead of
# showing it (measured 2026-10-08), and an assistant that has to go and read a
# file has been handed a pointer, not the knowledge. So no answer may exceed
# RESPONSE_LIMIT, which leaves headroom for other clients, whose limits are
# unmeasured. A template is a master plus ordered step files; the server turns
# them into parts, and each answer names the next call and when to make it.

RESPONSE_LIMIT = 45000
# Standards read while building; every other standard is read while designing.
_BUILD_STANDARDS = ("error-handling", "query-style", "startup-conventions")
# ST: every template declares steps. Off until the rollout finishes, as FM9 was.
REQUIRE_STEPS = False

# Each `when` is an order on the work, not a label: it says what not to do before the part
# arrives and closes the shortcut of working from what the assistant already knows.
_WHEN = {
    "master": "now",
    "design-standards": ("Fetch this before you ask the standards gate's first question. Ask no "
                         "design question until it has arrived."),
    "design-facts": ("Fetch this before you draft any part of the design. Do not draft from what "
                     "you already know about Access: these facts are what make the design right."),
    "build-method": ("Fetch this only after the developer chooses Build it or Give me the code. "
                     "Open, create or change no database before it has arrived."),
    "build-standards": ("Fetch this straight after the build method. Write no code before it has "
                        "arrived."),
    "build-facts": ("Fetch this straight after the build standards. Write no code before it has "
                    "arrived."),
}
_WHEN_DESIGN_STEP = ("Fetch this before you draft any part of the design. Do not draft from what "
                     "you already know: it carries what the design must satisfy.")
_WHEN_FIRST_BUILD_STEP = (
    "Fetch this as soon as the build parts before it have arrived. Do none of this step's work "
    "before then, even if you already know how from another template or file: this step carries "
    "the facts its work depends on.")
_WHEN_BUILD_STEP = (
    "Fetch this only when every action in step {prev} is done. Do none of this step's work before "
    "it arrives, even if you already know how from another template or file: this step carries the "
    "facts its work depends on.")
_STEP_RULE = (
    "Do this step's work now, and only this step's. Write, import or run nothing a later step "
    "covers, and do not fetch the next step until everything here is done. If you did any of this "
    "step's work before this part arrived, say so in the build record and check that work against "
    "this text now.")
_STANDARDS_NOTE = (
    "These standards govern this part. Their full text arrived in the part named beside each. "
    "If you cannot see that text now (it was summarized away, or you would be quoting it from "
    "memory), fetch that part again before you write anything it governs.")
_DESIGN_END = (
    "The design route ends when the developer approves the design. If they then ask for the "
    "code, call get_part with part 'build-method' and route 'build', follow next from there, "
    "and mark every file you hand over UNVERIFIED.")


def serialized_size(obj) -> int:
    """Characters in an answer as the server sends it (indented JSON, characters unescaped)."""
    return len(json.dumps(obj, indent=2, ensure_ascii=False))


def _sha(*texts: str) -> str:
    return hashlib.sha256("\n".join(texts).encode("utf-8")).hexdigest()[:12]


def steps_folder(path: Path) -> Path:
    return path.with_name(path.stem + STEPS_SUFFIX)


def read_steps(path: Path, front: dict) -> list[dict]:
    """The step files a master lists under `steps`, in its order.

    Each is `{n, step, path, front, body}`; a listed file that does not exist
    comes back with `path` None, for validate to report.
    """
    folder = steps_folder(path)
    out = []
    for n, sid in enumerate(front.get("steps") or [], 1):
        f = folder / f"{sid}.md"
        if f.is_file():
            try:
                sfront, sbody = split_front_matter(f.read_text(encoding="utf-8"))
            except yaml.YAMLError as exc:
                sfront, sbody = {"_error": str(exc).splitlines()[0]}, ""
        else:
            f, sfront, sbody = None, {}, ""
        out.append({"n": n, "step": str(sid), "path": f, "front": sfront, "body": sbody})
    return out


def _pack(items: list[dict], budget: int) -> list[list[dict]]:
    """Group items (each with `content`) in order so no group's text exceeds `budget`."""
    groups, cur, size = [], [], 0
    for it in items:
        n = len(it["content"])
        if cur and size + n > budget:
            groups.append(cur)
            cur, size = [], 0
        cur.append(it)
        size += n
    if cur:
        groups.append(cur)
    return groups


def _split(name: str, items: list[dict]) -> list[tuple[str, list[dict]]]:
    """Name each group: one group keeps the plain name, several are numbered."""
    groups = _pack(items, RESPONSE_LIMIT - 5000)
    if len(groups) == 1:
        return [(name, groups[0])]
    return [(f"{name}-{i}", g) for i, g in enumerate(groups, 1)]


def _facts(ids) -> tuple[list[dict], list[str]]:
    facts, _ = read_platform_facts()
    found = [{"id": str(f), "heading": facts[str(f)]["heading"], "content": facts[str(f)]["content"],
              "route": facts[str(f)]["route"]} for f in ids or [] if str(f) in facts]
    return found, [str(f) for f in ids or [] if str(f) not in facts]


def _standards(names) -> tuple[list[dict], list[str]]:
    found, missing = [], []
    for name in names:
        path = STANDARDS_DIR / f"{name}.md"
        if path.is_file():
            found.append({"name": name, "content": path.read_text(encoding="utf-8")})
        else:
            missing.append(name)
    return found, missing


def _headings(text: str) -> list[str]:
    out, fence = [], False
    for ln in text.splitlines():
        if ln.lstrip().startswith("```"):
            fence = not fence
        elif not fence and ln.startswith("## "):
            out.append(ln[3:].strip())
    return out


def _method(ids) -> tuple[list[dict], list[str]]:
    sections, _ = read_method()
    found = [{k: sections[m][k] for k in ("id", "heading", "content")} for m in ids if m in sections]
    return found, [m for m in ids if m not in sections]


def plan(path: Path, front: dict, body: str) -> list[dict]:
    """Every part a run can fetch for this template, in order, both routes.

    Each part is `{part, title, when, route, ...}` plus what composes it. The
    design route stops before `build-method`.
    """
    layer = [str(s) for s in front.get("standards_layer") or []]
    steps = read_steps(path, front)
    facts, _ = _facts(front.get("platform_facts"))
    design_facts = [f for f in facts if f["route"] != "build"]
    build_facts = [f for f in facts if f["route"] == "build"]

    parts = [{"part": "master", "title": "The template", "route": ""}]
    std, _ = _standards([s for s in layer if s not in _BUILD_STANDARDS])
    for name, group in _split("design-standards", std):
        parts.append({"part": name, "title": "The standards the design follows", "route": "",
                      "standards": group})
    for name, group in _split("design-facts", design_facts):
        parts.append({"part": name, "title": "Platform facts the design depends on", "route": "",
                      "facts": group})
    # A step marked `route: both` is read while designing too (a checklist the design must
    # pass, say), so it arrives with the design parts. validate keeps such steps first.
    for s in steps:
        if str(s["front"].get("route", "")) == "both":
            parts.append({"part": f"step-{s['n']}", "title": str(s["front"].get("title") or s["step"]),
                          "route": "", "step": s})
    parts.append({"part": "build-method", "title": "How the build is conducted", "route": "build"})
    std, _ = _standards([s for s in layer if s in _BUILD_STANDARDS])
    for name, group in _split("build-standards", std):
        parts.append({"part": name, "title": "The standards the code follows", "route": "build",
                      "standards": group})
    if not steps:
        # A template not yet divided into steps: its build facts arrive together.
        for name, group in _split("build-facts", build_facts):
            parts.append({"part": name, "title": "Platform facts the build uses", "route": "build",
                          "facts": group})
    for s in steps:
        if str(s["front"].get("route", "")) != "both":
            parts.append({"part": f"step-{s['n']}", "title": str(s["front"].get("title") or s["step"]),
                          "route": "build", "step": s})
    first_build = next((s["n"] for s in steps if str(s["front"].get("route", "")) != "both"), 0)
    for p in parts:
        p["when"] = _WHEN.get(p["part"].rstrip("0123456789").rstrip("-"), "")
        if p["part"].startswith("step-"):
            n = p["step"]["n"]
            if p["route"] == "":
                p["when"] = _WHEN_DESIGN_STEP
            elif n == first_build:
                p["when"] = _WHEN_FIRST_BUILD_STEP
            else:
                p["when"] = _WHEN_BUILD_STEP.format(prev=n - 1)
        if p["part"].startswith("step-") and p["step"]["front"].get("when"):
            p["when"] = str(p["step"]["front"]["when"])
    return parts


def _in_force(parts: list[dict]) -> list[dict]:
    """Each standard the template follows, the part that carried it, and its sections."""
    out = []
    for p in parts:
        for s in p.get("standards") or []:
            out.append({"name": s["name"], "part": p["part"], "sections": _headings(s["content"])})
    return out


def compose(template: str, part: str = "master", route: str = "",
            have_method: bool = False) -> dict:
    """One part of a template, exactly as the server answers it.

    Raises ValueError for an unknown template or part id.
    """
    tid = template.strip().lower()
    for path, front, body in iter_templates():
        if str(front.get("template", "")).lower() == tid:
            break
    else:
        raise ValueError(f"No template with id '{template}'. "
                         "Use list_templates or search_templates to find valid ids.")
    parts = plan(path, front, body)
    if route == "design":
        visible = [p for p in parts if p["route"] != "build"]
    else:
        visible = parts
    ids = [p["part"] for p in parts]
    pid = part.strip().lower()
    if pid not in ids:
        raise ValueError(f"Template '{template}' has no part '{part}'. Its parts: {', '.join(ids)}.")
    p = parts[ids.index(pid)]
    version = front.get("version")
    result = {"template": front.get("template"), "part": pid, "title": p["title"]}

    if pid == "master":
        own = _front_method(front) + _body_method(body)
        if not have_method:
            own = list(_METHOD_EVERY_RUN) + own + (list(_METHOD_DESIGN) if route != "build" else [])
        method, missing = _method(own)
        result = {**{k: front.get(k) for k in ("template", "title", "domain", "type", "status")},
                  "part": "master", "front_matter": front, "body": body, "method": method,
                  "parts": [{"part": q["part"], "title": q["title"], "when": q["when"]}
                            for q in visible],
                  "served": {"template": front.get("template"), "part": "master",
                             "version": version, "sha": _sha(path.read_text(encoding="utf-8"))}}
        if missing:
            result["method_missing"] = missing
        layer = [str(s) for s in front.get("standards_layer") or []]
        _, std_missing = _standards(layer)
        if std_missing:
            result["standards_missing"] = std_missing
        _, facts_missing = _facts(front.get("platform_facts"))
        if facts_missing:
            result["platform_facts_missing"] = facts_missing
    elif "standards" in p:
        result["standards"] = p["standards"]
        result["served"] = {"template": front.get("template"), "part": pid, "version": version,
                            "sha": _sha(*[s["content"] for s in p["standards"]])}
    elif "facts" in p:
        result["platform_facts"] = [{k: f[k] for k in ("id", "heading", "content")}
                                    for f in p["facts"]]
        result["served"] = {"template": front.get("template"), "part": pid, "version": version,
                            "sha": _sha(*[f["content"] for f in p["facts"]])}
        if pid.startswith("build-facts"):
            result["standards_in_force"] = _in_force(parts)
            result["standards_note"] = _STANDARDS_NOTE
    elif pid == "build-method":
        method, missing = _method(_METHOD_BUILD)
        result["method"] = method
        if missing:
            result["method_missing"] = missing
        result["served"] = {"template": front.get("template"), "part": pid, "version": version,
                            "sha": _sha(*[m["content"] for m in method])}
    else:
        s = p["step"]
        if s["path"] is None:
            raise ValueError(f"Step file for '{s['step']}' is missing from {steps_folder(path).name}.")
        earlier = _body_method(body)
        for q in parts:
            if q.get("step") and q["step"]["n"] < s["n"]:
                earlier += _body_method(q["step"]["body"])
        method, _ = _method([m for m in _body_method(s["body"]) if m not in earlier])
        facts, facts_missing = _facts(s["front"].get("platform_facts"))
        result.update({"step": s["n"], "of": sum(1 for q in parts if q.get("step")),
                       "rule": _STEP_RULE,
                       "body": s["body"],
                       "platform_facts": [{k: f[k] for k in ("id", "heading", "content")}
                                          for f in facts],
                       "method": method,
                       "standards_in_force": _in_force(parts),
                       "standards_note": _STANDARDS_NOTE,
                       "served": {"template": front.get("template"), "part": pid,
                                  "version": version,
                                  "sha": _sha(s["path"].read_text(encoding="utf-8"))}})
        if facts_missing:
            result["platform_facts_missing"] = facts_missing

    # A build part fetched on the design route (the developer asked for the code) chains on
    # through the build parts.
    chain = visible if pid in [q["part"] for q in visible] else parts
    chain_ids = [q["part"] for q in chain]
    if chain_ids.index(pid) + 1 < len(chain_ids):
        nxt = chain[chain_ids.index(pid) + 1]
        result["next"] = {"call": "get_part", "template": front.get("template"),
                          "part": nxt["part"], "route": route, "when": nxt["when"]}
    elif route == "design" and chain is visible:
        result["next"] = None
        result["after_design"] = _DESIGN_END
    else:
        result["next"] = None
    return result


# FM9: every template declares platform_facts (switched on 2026-10-06 when the rollout finished).
REQUIRE_PLATFORM_FACTS = True


# --- validate(): format-only rules from templates/_template-schema.md ---
# Each check returns a short "RULE: message" string; an empty list means the
# template is well-formed. No host database is ever opened here.

_TYPE_ENUM = {"table-schema", "vba-scaffold", "outcome-first", "form-spec", "spec"}
_STATUS_ENUM = {"draft", "review", "stable"}
_STANDARDS_VALUES = {"audit-columns", "naming-conventions", "error-handling",
                     "query-style", "form-conventions", "design-principles",
                     "startup-conventions"}
_AUDIT_COLUMNS = {"addedby", "addedon", "modifiedby", "modifiedon"}
_ACCESS_SCALAR_TYPES = {"autonumber", "long", "integer", "byte", "single", "double",
                        "currency", "memo", "date/time", "boolean", "guid"}
_TEXT_TYPE = re.compile(r"^text\(\d+\)$", re.I)
_REQUIRED_FM = ("template", "title", "domain", "type", "version", "status", "standards_layer")
_FIELD_TABLE_HEADER = ["field", "type", "key/req", "purpose&rules"]
_FK_RE = re.compile(r"FK\s*(?:→|->)\s*`?([A-Za-z0-9_]+)`?")
# A field name is read as words: an all-caps run, or a capital and what follows it.
_SEGMENT = re.compile(r"[A-Z]+(?![a-z])|[A-Z][a-z0-9]*")
_ACRONYM = re.compile(r"^[A-Z]{2,}$")
_TABLE_PREFIX = re.compile(r"^(tlkp|tbl|USys)", re.I)
# `TableA (1) → (∞) TableB` — the one cardinality-arrow shape every '## Relationships'
# bullet in the library uses; tolerant of any cardinality token ("1", "∞", "0..1").
_REL_ARROW_RE = re.compile(r"`([A-Za-z0-9_]+)\s*\([^)]*\)\s*→\s*\([^)]*\)\s*([A-Za-z0-9_]+)`")
# "Business Rule 3" / "Business Rules 2 and 7" — the two citation shapes actually in use.
_BUSINESS_RULE_CITE_RE = re.compile(r"Business\s+Rules?\s+(\d+)(?:\s+and\s+(\d+))?", re.I)


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
    """A section's text; where a master and its steps each carry one, all of them together."""
    found = [t for h, t in sections if h.lower().startswith(name.lower())]
    return "\n".join(found) if found else None


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


def _entity_bases(tables):
    """Entity names behind a set of table names: prefix stripped, plural-tolerant.

    `tblPublication` and `Products` both yield the words a key field may be
    built on, so `PublicationID` and `ProductID` can be recognized as keys of a
    table this template knows about.
    """
    bases = set()
    for t in tables:
        name = _TABLE_PREFIX.sub("", str(t).strip().replace(" ", "")).lower()
        bases.update({name, name + "s"})
        if name.endswith("s"):
            bases.add(name[:-1])
    return bases


def _field_qualified(name, entity_bases):
    """Is this field name qualified, or one bare noun?

    The test is whether the name is qualified, never whether it is reserved, so
    no word list is involved and a reserved word added years from now is caught
    the day it appears: reserved words are single common words by definition.

    A trailing `ID` is not a qualifier. The word in front of it is what gets
    judged, which is why `PublicationID` passes (`Publication` names the table)
    while `StatusID` does not. An acronym standing alone is accepted (`ISBN`),
    with one exception: `ID` by itself, because the standards layer in this
    repository requires `[Entity]ID` for a key. A shop that replaces
    `standards/` with its own may take a different view of that last point.
    """
    if not re.fullmatch(r"[A-Za-z][A-Za-z0-9]*", name):
        return True                                    # other rules speak to these
    if name == "ID":
        return False
    if _ACRONYM.match(name):
        return True
    if name.endswith("ID") and len(name) > 2:
        base = name[:-2]
        return len(_SEGMENT.findall(base)) >= 2 or base.lower() in entity_bases
    return len(_SEGMENT.findall(name)) >= 2


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
    for key in ("requires_fields", "new_fields", "seeds"):
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
    declared = front.get("platform_facts") or []
    if declared:
        facts, dupes = read_platform_facts()
        for fid in declared:
            if str(fid) not in facts:
                errors.append(f"FM7: platform_facts id '{fid}' matches no fact marker in _materialization.md")
            elif str(fid) in dupes:
                errors.append(f"FM8: fact id '{fid}' is defined more than once in _materialization.md")
    elif REQUIRE_PLATFORM_FACTS and typ != "spec":
        errors.append("FM9: missing/empty platform_facts")
    methods, _ = read_method()
    for mid in method_for(front, body):
        if mid not in methods:
            errors.append(f"MT1: method '{mid}' this template receives has no section in _method.md")

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
        errors += _validate_table_schema(front, sections, body)
    elif typ == "vba-scaffold":
        errors += _validate_vba_scaffold(front, sections)
    elif typ == "outcome-first":
        errors += _validate_outcome_first(front, sections)
    elif typ == "form-spec":
        errors += _validate_form_spec(front, sections)
    return errors


def _validate_table_schema(front, sections, body):
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
    bases = _entity_bases(known)
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
            elif not _field_qualified(r[0].strip("` "), bases):
                errors.append(f"TS6: field '{r[0].strip('` ')}' is not qualified (add the entity or "
                              "purpose it belongs to; a trailing 'ID' does not qualify it)")
            tval = r[1].strip().split()[0].lower().rstrip(".,") if r[1].strip() else ""
            if tval and tval not in _ACCESS_SCALAR_TYPES and not _TEXT_TYPE.match(tval):
                errors.append(f"TS2: field '{r[0]}' has unknown type '{r[1]}'")
            m = _FK_RE.search(r[2])
            if m and m.group(1) not in known:
                errors.append(f"TS3: FK target '{m.group(1)}' (field {r[0]}) resolves to no known table")

    # ---- Relationships: every named table resolves, every line states cascade behavior ----
    rel = _section_text(sections, "Relationships") or ""
    for line in rel.splitlines():
        m = _REL_ARROW_RE.search(line)
        if not m:
            continue                                       # not an arrow relationship line
        for name in m.groups():
            if name not in known:
                errors.append(f"TS4: relationship names '{name}', which resolves to no known table")
        if not re.search(r"cascade|restrict", line, re.I):
            errors.append("TS7: relationship line does not state its cascade behavior "
                          f"('cascade' or 'no cascade'/'restrict'): {line.strip()[:80]}")

    # ---- '## Validating the build' is required and must say something ----
    vtb = _section_text(sections, "Validating the build")
    if vtb is None:
        errors.append("TS8: missing required section '## Validating the build'")
    elif not vtb.strip():
        errors.append("TS8: '## Validating the build' is present but empty")

    # ---- Every '(Business Rule N)' citation resolves to a real numbered rule ----
    br_text = _section_text(sections, "Business Rules") or ""
    rule_nums = [int(n) for n in re.findall(r"^\s*(\d+)\.\s", br_text, re.M)]
    max_rule = max(rule_nums) if rule_nums else 0
    if max_rule:
        for m in _BUSINESS_RULE_CITE_RE.finditer(body):
            for g in m.groups():
                if g and not (1 <= int(g) <= max_rule):
                    errors.append(f"TS9: citation 'Business Rule {g}' does not resolve to any of the "
                                  f"{max_rule} numbered items in '## Business Rules'")

    # ---- Every front-matter 'seeds' entry is described somewhere, not just named ----
    for entry in front.get("seeds") or []:
        token = str(entry).split(".")[-1].strip().strip("`")
        if token and token.lower() not in body.lower():
            errors.append(f"TS10: seeds entry '{entry}' is not described anywhere in the template body")

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


def _validate_outcome_first(front, sections):
    """An outcome-first template states the finished condition and no route to it.

    The rules are the mirror image of `_validate_vba_scaffold`: the three sections
    carrying the promise are required, and the three pieces of route specification
    are forbidden. Declaring any of the latter is the one way a template of this
    type stops being one.
    """
    errors = []
    for sec in ("What you end up with",
                "Information and conditions you need to supply",
                "To the AI assistant building this"):
        if not _has_section(sections, sec):
            errors.append(f"OF: missing required section '## {sec}'")
    if "error-handling" not in [str(x) for x in front.get("standards_layer") or []]:
        errors.append("OF1: outcome-first standards_layer must include 'error-handling'")
    if front.get("implements") and not _slug_ok(front["implements"]):
        errors.append(f"OF2: implements '{front['implements']}' is not a well-formed slug")
    for key in ("target_module", "new_procedures"):
        if front.get(key):
            errors.append(f"OF3: type 'outcome-first' must not declare '{key}' - that is route "
                          "specification, which this type exists to leave out")
    if _has_section(sections, "Procedures"):
        errors.append("OF3: type 'outcome-first' must not have a '## Procedures' section - that is "
                      "route specification, which this type exists to leave out")
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


def validate_steps(path: Path, front: dict, body: str, steps: list[dict]) -> list[str]:
    """The master/steps rules (ST) and the answer-size rule (SZ1)."""
    errors = []
    folder = steps_folder(path)
    if not steps:
        if REQUIRE_STEPS and str(front.get("type", "")) != "spec":
            errors.append("ST0: missing/empty steps")
        if folder.is_dir():
            errors.append(f"ST1: {folder.name}/ exists but the master lists no steps")
        return errors
    listed = {s["step"] for s in steps}
    for s in steps:
        if s["path"] is None:
            errors.append(f"ST1: step '{s['step']}' has no file {folder.name}/{s['step']}.md")
            continue
        if "_error" in s["front"]:
            errors.append(f"ST1: {s['step']}.md front matter does not parse: {s['front']['_error']}")
            continue
        if str(s["front"].get("step", "")) != s["step"]:
            errors.append(f"ST1: {s['step']}.md front-matter 'step' must be '{s['step']}'")
        if not str(s["front"].get("title", "")).strip():
            errors.append(f"ST1: {s['step']}.md has no 'title'")
        if "**Who reads this:**" not in s["body"]:
            errors.append(f"ST4: {s['step']}.md does not name its reader ('**Who reads this:**')")
        facts, dupes = read_platform_facts()
        for fid in s["front"].get("platform_facts") or []:
            if str(fid) not in facts:
                errors.append(f"ST2: step '{s['step']}' fact '{fid}' matches no fact marker")
    routes = [str(s["front"].get("route", "")) for s in steps]
    if "both" in routes and any(r != "both" for r in routes[:max(i for i, r in enumerate(routes)
                                                              if r == "both")]):
        errors.append("ST5: a step marked 'route: both' is listed after a build step; list it first")
    for s in steps:
        if str(s["front"].get("route", "")) not in ("", "both"):
            errors.append(f"ST5: step '{s['step']}' route must be 'both' or absent")
    for f in sorted(folder.glob("*.md")) if folder.is_dir() else []:
        if f.stem not in listed:
            errors.append(f"ST1: {folder.name}/{f.name} is not listed under steps")
    facts, _ = read_platform_facts()
    for fid in front.get("platform_facts") or []:
        if facts.get(str(fid), {}).get("route") == "build":
            errors.append(f"ST3: build fact '{fid}' is in the master; it belongs in the step "
                          "that uses it")
    if errors:
        return errors          # sizes are measured only on a template whose steps resolve
    for p in plan(path, front, body):
        for route in ("", "design", "build"):
            if route == "design" and p["route"] == "build":
                continue
            size = serialized_size(compose(str(front.get("template")), p["part"], route, True))
            if size > RESPONSE_LIMIT:
                errors.append(f"SZ1: part '{p['part']}' (route '{route or 'both'}') is {size} "
                              f"characters; the limit is {RESPONSE_LIMIT}")
                break
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
        steps = read_steps(path, front)
        whole = "\n".join([body] + [s["body"] for s in steps])
        errs = validate_template(front, whole, path.stem) + validate_steps(path, front, body, steps)
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
