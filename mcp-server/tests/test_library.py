"""Tests for library.py — parsing helpers and the validate() rule set.

Run from the repo root:  python -m unittest discover -s mcp-server/tests
Fixture strategy: tests/fixtures/broken holds one deliberately broken template
per validation rule (the B3-13 proof pattern, made permanent); the live corpus
under templates/ must always pass in full.
"""

import sys
import unittest
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import library
from library import (iter_standards, iter_templates, read_standard,
                     split_front_matter, validate_library, validate_template)

FIXTURES = Path(__file__).resolve().parent / "fixtures"
BROKEN = FIXTURES / "broken"
DUPES = FIXTURES / "duplicate_slugs"

# filename stem -> the rule prefix its validation errors must include
EXPECTED_RULE = {
    "fm1_missing_title": "FM1",
    "fm2_bad_slug": "FM2",
    "fm3_extends_no_requires": "FM3",
    "fm5_bad_requires_field": "FM5",
    "fm6_bad_house_assumption": "FM6",
    "fm_unknown_standard": "FM:",
    "core_missing_intent": "CORE",
    "ts1_undeclared_table": "TS1",
    "ts2_unknown_type": "TS2",
    "ts3_unknown_fk": "TS3",
    "ts5_audit_column": "TS5",
    "vs4_missing_error_handling": "VS4",
    "fs1_missing_record_source": "FS1",
}


class TestSplitFrontMatter(unittest.TestCase):
    def test_splits_front_matter_and_body(self):
        front, body = split_front_matter("---\ntemplate: x\n---\n\n# Body\n")
        self.assertEqual(front, {"template": "x"})
        self.assertTrue(body.startswith("# Body"))

    def test_no_front_matter_returns_empty_dict(self):
        front, body = split_front_matter("# Just a body\n")
        self.assertEqual(front, {})
        self.assertEqual(body, "# Just a body\n")


class TestIterTemplates(unittest.TestCase):
    def test_skips_infrastructure_and_readme(self):
        with patch.object(library, "TEMPLATES_DIR", DUPES):
            names = [p.name for p, _, _ in iter_templates()]
        self.assertEqual(sorted(names), ["dup.md", "team-dup.md"])
        self.assertNotIn("_infra-file.md", names)
        self.assertNotIn("README.md", names)


class TestStandardsHelpers(unittest.TestCase):
    def test_read_standard_found_and_missing(self):
        self.assertIsNotNone(read_standard("naming-conventions"))
        self.assertIsNone(read_standard("no-such-standard"))

    def test_iter_standards_sorted_no_readme_nonempty(self):
        entries = list(iter_standards())
        names = [n for n, _ in entries]
        self.assertIn("naming-conventions", names)
        self.assertEqual(names, sorted(names))
        self.assertNotIn("readme", [n.lower() for n in names])
        for name, content in entries:
            self.assertTrue(content.strip(), f"standard '{name}' is empty")


class TestValidateTemplateRules(unittest.TestCase):
    def test_each_broken_fixture_fails_on_its_rule(self):
        for path in sorted(BROKEN.glob("*.md")):
            with self.subTest(fixture=path.stem):
                expected = EXPECTED_RULE[path.stem]
                front, body = split_front_matter(path.read_text(encoding="utf-8"))
                errors = validate_template(front, body, path.stem)
                self.assertTrue(errors, "expected at least one error")
                self.assertTrue(
                    any(e.startswith(expected) for e in errors),
                    f"expected an error starting with '{expected}', got: {errors}",
                )

    def test_every_fixture_has_an_expectation(self):
        stems = {p.stem for p in BROKEN.glob("*.md")}
        self.assertEqual(stems, set(EXPECTED_RULE), "fixture/expectation drift")


class TestValidateLibrary(unittest.TestCase):
    def test_live_corpus_all_pass(self):
        result = validate_library()
        self.assertTrue(result["ok"], f"live corpus failed: {result['results']}")
        self.assertGreaterEqual(result["checked"], 5)
        for r in result["results"]:
            self.assertTrue(r["ok"], f"{r['template']}: {r['errors']}")

    def test_single_template_by_id(self):
        result = validate_library("northwind-stocktake-schema")
        self.assertEqual(result["checked"], 1)
        self.assertTrue(result["ok"])

    def test_unknown_id_raises(self):
        with self.assertRaises(ValueError):
            validate_library("no-such-template")

    def test_duplicate_slugs_flagged_on_both(self):
        with patch.object(library, "TEMPLATES_DIR", DUPES):
            result = validate_library()
        self.assertFalse(result["ok"])
        self.assertEqual(result["checked"], 2)
        for r in result["results"]:
            self.assertTrue(
                any("not unique" in e for e in r["errors"]),
                f"{r['template']}: {r['errors']}",
            )


if __name__ == "__main__":
    unittest.main()
