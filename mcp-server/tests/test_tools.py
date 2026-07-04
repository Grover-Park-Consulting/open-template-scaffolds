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


class TestGetTemplate(unittest.TestCase):
    def test_composes_template_with_standards(self):
        result = get_template("northwind-stocktake-schema")
        self.assertTrue(result["body"].strip())
        layer = result["front_matter"]["standards_layer"]
        self.assertEqual([s["name"] for s in result["standards"]], layer)
        self.assertNotIn("standards_missing", result)
        for s in result["standards"]:
            self.assertTrue(s["content"].strip())

    def test_unknown_id_raises(self):
        with self.assertRaises(ValueError):
            get_template("no-such-template")

    def test_unresolved_standard_reported_not_dropped(self):
        with patch.object(server, "read_standard", return_value=None):
            result = get_template("northwind-stocktake-schema")
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
