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

import server


def unwrap(tool):
    return tool.fn if hasattr(tool, "fn") else tool


list_templates = unwrap(server.list_templates)
search_templates = unwrap(server.search_templates)
get_template = unwrap(server.get_template)
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


class TestGetTemplate(unittest.TestCase):
    def test_composes_template_with_standards(self):
        result = get_template("stocktake-schema")
        self.assertTrue(result["body"].strip())
        layer = result["front_matter"]["standards_layer"]
        self.assertEqual([s["name"] for s in result["standards"]], layer)
        self.assertNotIn("standards_missing", result)
        for s in result["standards"]:
            self.assertTrue(s["content"].strip())

    def test_unknown_id_raises(self):
        with self.assertRaises(ValueError):
            get_template("no-such-template")

    def test_delivers_declared_platform_facts_and_served(self):
        result = get_template("time-off-ledger-outcome-first")
        declared = result["front_matter"]["platform_facts"]
        self.assertEqual([f["id"] for f in result["platform_facts"]], declared)
        self.assertNotIn("platform_facts_missing", result)
        for f in result["platform_facts"]:
            self.assertTrue(f["content"].startswith("#"))
        self.assertEqual(result["served"]["template"], "time-off-ledger-outcome-first")
        self.assertEqual(len(result["served"]["sha"]), 12)

    def test_method_follows_route_and_features(self):
        ids = lambda r: [m["id"] for m in r["method"]]
        build = get_template("time-off-ledger-outcome-first", route="build")
        design = get_template("time-off-ledger-outcome-first", route="design")
        self.assertIn("access-gate", ids(build))
        self.assertNotIn("design-only-handover", ids(build))
        self.assertIn("design-only-handover", ids(design))
        self.assertNotIn("access-gate", ids(design))
        self.assertIn("explore-options", ids(build))
        self.assertNotIn("wizard", ids(build))
        self.assertIn("wizard", ids(get_template("error-logging-scaffold")))
        self.assertNotIn("method_missing", build)
        with self.assertRaises(ValueError):
            get_template("time-off-ledger-outcome-first", route="sideways")

    def test_design_route_trims_build_facts_and_known_method(self):
        full = get_template("error-logging-schema", route="build")
        lean = get_template("error-logging-schema", route="design", have_method=True)
        self.assertIn("dao-table-build", [f["id"] for f in full["platform_facts"]])
        self.assertNotIn("dao-table-build", [f["id"] for f in lean["platform_facts"]])
        self.assertIn("dao-table-build", lean["platform_facts_omitted"])
        self.assertIn("target-file", [f["id"] for f in lean["platform_facts"]])
        self.assertNotIn("run-opening", [m["id"] for m in lean["method"]])
        self.assertIn("run-opening", lean["method_omitted"])
        wiz = get_template("error-logging-scaffold", route="build", have_method=True)
        self.assertIn("wizard", [m["id"] for m in wiz["method"]])
        self.assertLess(len(str(lean)), len(str(full)))

    def test_every_route_gets_build_record_and_between_questions(self):
        for r in ("design", "build"):
            got = [m["id"] for m in get_method(r)["method"]]
            self.assertIn("build-record", got, r)
            self.assertIn("between-questions", got, r)
        self.assertNotIn("quiet-build", [m["id"] for m in get_method("design")["method"]])

    def test_get_method_for_a_run_with_no_template(self):
        result = get_method("build")
        got = [m["id"] for m in result["method"]]
        self.assertIn("run-opening", got)
        self.assertIn("access-gate", got)
        self.assertNotIn("wizard", got)
        self.assertNotIn("method_missing", result)

    def test_every_template_receives_its_declared_facts(self):
        for t in list_templates():
            result = get_template(t["template"])
            self.assertNotIn("platform_facts_missing", result, t["template"])
            self.assertTrue(result["platform_facts"], t["template"])

    def test_unresolved_standard_reported_not_dropped(self):
        with patch.object(server, "read_standard", return_value=None):
            result = get_template("stocktake-schema")
        self.assertEqual(result["standards"], [])
        self.assertEqual(
            result["standards_missing"],
            result["front_matter"]["standards_layer"],
        )


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
