"""Tests for the MCP tool surface in server.py (except check_compatibility —
see test_compat.py).

FastMCP wraps each @mcp.tool() function; unwrap() reaches the plain function so
tests exercise exactly what a client call reaches, without a running server.
"""

import sys
import unittest
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import library
import server


def unwrap(tool):
    return tool.fn if hasattr(tool, "fn") else tool


list_templates = unwrap(server.list_templates)
search_templates = unwrap(server.search_templates)
get_template = unwrap(server.get_template)
get_part = unwrap(server.get_part)
get_standards = unwrap(server.get_standards)
get_method = unwrap(server.get_method)
validate = unwrap(server.validate)

META_KEYS = ("template", "title", "domain", "type", "status")


class TestListTemplates(unittest.TestCase):
    def test_returns_metadata_for_every_domain_template(self):
        result = list_templates()
        self.assertGreaterEqual(len(result), 5)
        for entry in result:
            for key in META_KEYS:
                self.assertIn(key, entry)
        slugs = [e["template"] for e in result]
        self.assertEqual(len(slugs), len(set(slugs)), "duplicate slugs surfaced")


class TestSearchTemplates(unittest.TestCase):
    def test_identity_hit_is_likely_and_sorted_first(self):
        results = search_templates(query="stocktake")
        self.assertTrue(results)
        self.assertEqual(results[0]["relevance"], "Likely")

    def test_no_match_returns_empty_list(self):
        self.assertEqual(search_templates(query="qqqzzz-nothing"), [])

    def test_domain_and_type_filters(self):
        for entry in search_templates(domain="library"):
            self.assertEqual(entry["domain"], "library")
        for entry in search_templates(type="table-schema"):
            self.assertEqual(entry["type"], "table-schema")

    def test_relevance_tiers(self):
        front = {"template": "a-b", "title": "Title", "domain": "dom"}
        body = "## Intent\n\nAbout warehouse counting.\n\n## Entities\n"
        self.assertEqual(server._relevance("dom", front, body)[0], "Likely")
        self.assertEqual(server._relevance("warehouse", front, body)[0], "Possible")
        self.assertEqual(server._relevance("hous", front, body)[0], "Unlikely")
        self.assertIsNone(server._relevance("zebra", front, body))

    def test_words_match_when_the_whole_phrase_does_not(self):
        """A need described in the caller's own words still finds the template."""
        front = {"template": "a-b", "title": "Title", "domain": "dom"}
        body = "## Intent\n\nAbout warehouse counting.\n\n## Entities\n"
        rated = server._relevance("counting warehouse stock", front, body)
        self.assertIsNotNone(rated, "individual words must match when the phrase cannot")
        self.assertEqual(rated[0], "Possible")
        self.assertIn("2 of 3 words", rated[1])

    def test_word_match_anchors_at_the_start_of_a_word(self):
        """"log" must not match "catalog", nor "old" match "scaffold"."""
        front = {"template": "catalog-schema", "title": "Catalog", "domain": "library"}
        body = "## Intent\n\nA scaffold for catalogs.\n\n## Entities\n"
        self.assertIsNone(server._relevance("log old", front, body))

    def test_plural_query_finds_the_singular(self):
        front = {"template": "a-b", "title": "Title", "domain": "dom"}
        body = "## Intent\n\nRecords every change to a row.\n\n## Entities\n"
        rated = server._relevance("changes", front, body)
        self.assertIsNotNone(rated, "'changes' must find 'change'")


def walk(template, route=""):
    """Every answer a run fetches, master first, by following `next`."""
    got = [get_template(template, route=route, have_method=True)]
    while got[-1].get("next"):
        got.append(get_part(template, got[-1]["next"]["part"], route=route))
    return got


class TestGetTemplate(unittest.TestCase):
    def test_master_is_the_first_part_and_names_the_rest(self):
        result = get_template("stocktake-schema")
        self.assertTrue(result["body"].strip())
        self.assertEqual(result["part"], "master")
        self.assertNotIn("standards", result)
        self.assertNotIn("platform_facts", result)
        self.assertEqual(result["parts"][0]["part"], "master")
        self.assertEqual(result["next"]["part"], result["parts"][1]["part"])
        self.assertEqual(result["served"]["template"], "stocktake-schema")
        self.assertEqual(len(result["served"]["sha"]), 12)

    def test_unknown_id_raises(self):
        with self.assertRaises(ValueError):
            get_template("no-such-template")
        with self.assertRaises(ValueError):
            get_part("stocktake-schema", "no-such-part")

    def test_following_next_delivers_every_standard_and_fact(self):
        for t in list_templates():
            parts = walk(t["template"])
            front = parts[0]["front_matter"]
            stds = [s["name"] for p in parts for s in p.get("standards", [])]
            self.assertEqual(sorted(stds), sorted(front["standards_layer"]), t["template"])
            facts = {f["id"] for p in parts for f in p.get("platform_facts", [])}
            self.assertTrue(set(front["platform_facts"]) <= facts, t["template"])
            self.assertEqual([p["part"] for p in parts],
                             [q["part"] for q in parts[0]["parts"]], t["template"])
            self.assertIsNone(parts[-1]["next"])
            for p in parts:
                self.assertNotIn("platform_facts_missing", p, t["template"])
                self.assertNotIn("standards_missing", p, t["template"])

    def test_design_route_stops_before_the_build(self):
        parts = walk("error-logging-schema", route="design")
        ids = [p["part"] for p in parts]
        self.assertNotIn("build-method", ids)
        self.assertIn("after_design", parts[-1])
        facts = {f["id"] for p in parts for f in p.get("platform_facts", [])}
        self.assertIn("target-file", facts)
        self.assertNotIn("dao-table-build", facts)
        code = get_part("error-logging-schema", "build-method", route="design")
        self.assertIsNotNone(code["next"], "asking for the code chains on through the build")

    def test_method_arrives_where_it_is_used(self):
        ids = lambda r: [m["id"] for m in r["method"]]
        design = get_template("time-off-ledger-outcome-first", route="design")
        self.assertIn("design-only-handover", ids(design))
        self.assertNotIn("access-gate", ids(design))
        self.assertIn("explore-options", ids(design))
        lean = get_template("time-off-ledger-outcome-first", route="build", have_method=True)
        self.assertNotIn("run-opening", ids(lean))
        self.assertIn("access-gate", ids(get_part("time-off-ledger-outcome-first", "build-method")))
        wizard = [p for p in walk("error-logging-scaffold")
                  if "wizard" in [m["id"] for m in p.get("method", [])]]
        self.assertEqual(len(wizard), 1, "the wizard method arrives once")
        with self.assertRaises(ValueError):
            get_template("time-off-ledger-outcome-first", route="sideways")

    def test_every_route_gets_build_record_and_between_questions(self):
        for r in ("design", "build"):
            got = [m["id"] for m in get_method(r)["method"]]
            self.assertIn("build-record", got, r)
            self.assertIn("between-questions", got, r)
        self.assertNotIn("quiet-build", [m["id"] for m in get_method("design")["method"]])

    def test_get_method_build_stage_for_a_run_with_no_template(self):
        opening = [m["id"] for m in get_method("build")["method"]]
        self.assertIn("run-opening", opening)
        self.assertNotIn("access-gate", opening)
        result = get_method("build", stage="build")
        self.assertIn("access-gate", [m["id"] for m in result["method"]])
        self.assertNotIn("method_missing", result)
        with self.assertRaises(ValueError):
            get_method(stage="later")

    def test_build_facts_and_steps_name_the_standards_in_force(self):
        parts = walk("stocktake-schema", route="build")
        last = [p for p in parts if p["part"].startswith(("build-facts", "step-"))]
        self.assertTrue(last)
        for p in last:
            names = {s["name"] for s in p["standards_in_force"]}
            self.assertEqual(names, set(parts[0]["front_matter"]["standards_layer"]))
            for s in p["standards_in_force"]:
                self.assertTrue(s["sections"], s["name"])
            self.assertIn("fetch that part again", p["standards_note"])

    def test_unresolved_standard_reported_not_dropped(self):
        with patch.object(library, "STANDARDS_DIR", library.STANDARDS_DIR / "nowhere"):
            result = get_template("stocktake-schema")
        self.assertEqual(result["standards_missing"],
                         result["front_matter"]["standards_layer"])


class TestSteps(unittest.TestCase):
    PILOTS = ("audit-logging-lite-scaffold", "time-off-ledger-outcome-first",
              "time-off-ledger-scaffold")

    def test_every_part_of_a_divided_template_fits(self):
        for t in self.PILOTS:
            for route in ("design", "build"):
                for p in walk(t, route=route):
                    self.assertLessEqual(library.serialized_size(p), library.RESPONSE_LIMIT,
                                         f"{t} {p['part']} ({route})")

    def test_steps_arrive_in_order_with_their_facts(self):
        parts = walk("time-off-ledger-scaffold", route="build")
        steps = [p for p in parts if p["part"].startswith("step-")]
        self.assertEqual([p["step"] for p in steps], [1, 2, 3, 4])
        self.assertIn("dao-table-build", [f["id"] for f in steps[0]["platform_facts"]])
        self.assertTrue(all(p["served"]["part"] == p["part"] for p in parts))
        self.assertNotIn("dao-table-build", parts[0]["front_matter"]["platform_facts"])

    def test_a_design_step_arrives_before_the_design(self):
        design = [p["part"] for p in walk("time-off-ledger-outcome-first", route="design")]
        self.assertIn("step-1", design)
        self.assertNotIn("step-2", design)
        build = [p["part"] for p in walk("time-off-ledger-outcome-first", route="build")]
        self.assertLess(build.index("step-1"), build.index("build-method"))

    def test_the_wizard_method_comes_with_the_wizard_step(self):
        parts = walk("audit-logging-lite-scaffold")
        holders = [p["part"] for p in parts if "wizard" in [m["id"] for m in p.get("method", [])]]
        self.assertEqual(holders, ["step-1"])

    def test_validate_reports_step_problems(self):
        path = next(p for p, f, _ in library.iter_templates()
                    if f["template"] == "time-off-ledger-scaffold")
        front, body = library.split_front_matter(path.read_text(encoding="utf-8"))
        broken = dict(front, steps=list(front["steps"]) + ["99-missing"],
                      platform_facts=list(front["platform_facts"]) + ["dao-table-build"])
        errs = library.validate_steps(path, broken, body, library.read_steps(path, broken))
        self.assertTrue(any(e.startswith("ST1") and "99-missing" in e for e in errs))
        self.assertTrue(any(e.startswith("ST3") for e in errs))
        unlisted = dict(front, steps=list(front["steps"])[:-1])
        errs = library.validate_steps(path, unlisted, body, library.read_steps(path, unlisted))
        self.assertTrue(any("is not listed" in e for e in errs))

    def test_step_files_are_not_templates(self):
        self.assertFalse(any(".steps" in str(p) for p, _, _ in library.iter_templates()))


class TestGetStandards(unittest.TestCase):
    def test_default_returns_all_standards(self):
        result = get_standards()
        names = [e["name"] for e in result["standards"]]
        self.assertEqual(result["count"], len(names))
        self.assertIn("naming-conventions", names)
        self.assertEqual(names, sorted(names))
        self.assertNotIn("readme", [n.lower() for n in names])
        for entry in result["standards"]:
            self.assertTrue(entry["content"].strip())

    def test_single_fetch_is_case_insensitive(self):
        result = get_standards("Naming-Conventions")
        self.assertEqual(result["count"], 1)
        self.assertEqual(result["standards"][0]["name"], "naming-conventions")

    def test_unknown_name_raises_listing_available(self):
        with self.assertRaises(ValueError) as ctx:
            get_standards("no-such-standard")
        self.assertIn("naming-conventions", str(ctx.exception))


class TestValidateTool(unittest.TestCase):
    def test_whole_library_mode(self):
        result = validate()
        self.assertTrue(result["ok"])
        self.assertGreaterEqual(result["checked"], 5)

    def test_unknown_id_raises(self):
        with self.assertRaises(ValueError):
            validate("no-such-template")


if __name__ == "__main__":
    unittest.main()
