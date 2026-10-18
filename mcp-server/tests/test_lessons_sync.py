"""Tests for lessons_sync that need no database: parsing, delivery, the text fallback, and the
missing-PowerShell and missing-engine paths."""
import shutil
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import isolate_lessons  # noqa: E402
import lessons_sync as ls  # noqa: E402


class ReadTextLessons(unittest.TestCase):
    def test_seed_set_parses(self):
        mechanisms, lessons, text_hash = ls.read_text_lessons(ls.DEFAULT_LESSONS)
        self.assertEqual(len(lessons), 23)
        self.assertEqual({m["name"] for m in mechanisms if m["always"]}, {"error-handling", "vba-generation", "build-verification"})
        names = {m["name"] for m in mechanisms}
        for lesson in lessons:
            self.assertIn(lesson["mechanism"], names, lesson["slug"])
            self.assertIn(lesson["kind"], ("guardrail", "consequence"), lesson["slug"])
            self.assertTrue(lesson["applies_when"] and lesson["evidence"], lesson["slug"])
            if lesson["kind"] == "consequence":
                self.assertTrue(lesson["triggers_when"], lesson["slug"])
        self.assertEqual(len(text_hash), 64)


class EnsureIndexWithoutDatabase(unittest.TestCase):
    def test_no_powershell_falls_back_to_text(self):
        with tempfile.TemporaryDirectory() as tmp:
            result = ls.ensure_index(ls.DEFAULT_LESSONS, Path(tmp) / "store")  # isolate_lessons: no PowerShell
        self.assertEqual(result["status"], "from-text")
        self.assertIn("PowerShell was not found", result["message"])
        self.assertEqual(len(result["index"]["lessons"]), 23)

    def test_no_engine_falls_back_to_text(self):
        shells = isolate_lessons.REAL_POWERSHELLS()
        if not shells:
            self.skipTest("no PowerShell on this machine")
        with tempfile.TemporaryDirectory() as tmp:
            fake = Path(tmp) / "fake.ps1"
            fake.write_text("exit 20\n", encoding="utf-8")
            original = (ls.SCRIPT, ls._powershells)
            ls.SCRIPT, ls._powershells = fake, lambda: shells
            try:
                result = ls.ensure_index(ls.DEFAULT_LESSONS, Path(tmp) / "store")
            finally:
                ls.SCRIPT, ls._powershells = original
        self.assertEqual(result["status"], "from-text")
        self.assertIn("not available", result["message"])
        self.assertEqual(len(result["index"]["lessons"]), 23)

    def test_missing_lessons_folder_is_reported_not_raised(self):
        with tempfile.TemporaryDirectory() as tmp:
            result = ls.ensure_index(Path(tmp) / "nope", Path(tmp) / "store")
        self.assertEqual(result["status"], "unavailable")
        self.assertIsNone(result["index"])


class Delivery(unittest.TestCase):
    def deliver(self, mechanisms):
        with tempfile.TemporaryDirectory() as tmp:
            return ls.deliver(mechanisms, ls.DEFAULT_LESSONS, Path(tmp) / "store")

    def test_always_mechanisms_arrive_with_no_request(self):
        got = self.deliver([])
        self.assertTrue(got["lessons"])
        self.assertEqual({l["mechanism"] for l in got["lessons"]} - {"error-handling", "vba-generation", "build-verification"}, set())

    def test_requested_mechanism_and_alias_are_added_case_insensitively(self):
        by_name = {l["id"] for l in self.deliver(["Schema-Migration"])["lessons"]}
        by_alias = {l["id"] for l in self.deliver(["table migration"])["lessons"]}
        self.assertIn("widening-does-not-recover-truncated-data", by_name)
        self.assertEqual(by_name, by_alias)

    def test_unknown_name_is_reported_and_does_not_stop_delivery(self):
        got = self.deliver(["no-such-mechanism", "ace-sql"])
        self.assertEqual(got["mechanisms_unknown"], ["no-such-mechanism"])
        self.assertIn("no-count-distinct", {l["id"] for l in got["lessons"]})

    def test_action_follows_kind_and_confidence(self):
        lessons = {l["id"]: l for l in self.deliver(["field-properties", "error-handling"])["lessons"]}
        self.assertEqual(lessons["designer-gives-number-fields-a-default-of-zero"]["action"], "ask")
        self.assertEqual(lessons["failure-value-must-not-be-a-meaningful-value"]["action"], "apply")
        self.assertEqual(lessons["probe-returns-false-without-reporting"]["action"], "state")

    def test_record_line_names_every_delivered_lesson(self):
        got = self.deliver(["ace-sql"])
        self.assertTrue(got["record_line"].startswith(f"Lessons consulted: {len(got['lessons'])} ("))
        for l in got["lessons"]:
            self.assertIn(l["id"], got["record_line"])

    def test_unreadable_store_gives_a_record_line_with_the_reason(self):
        with tempfile.TemporaryDirectory() as tmp:
            got = ls.deliver([], Path(tmp) / "nope", Path(tmp) / "store")
        self.assertEqual(got["lessons"], [])
        self.assertTrue(got["record_line"].startswith("Lessons consulted: none ("))

    def test_potential_lessons_are_held_back(self):
        index = {"mechanisms": {"m": {"always": True, "aliases": []}},
                 "lessons": [{"id": "a", "mechanism": "m", "kind": "guardrail", "confidence": "Potential", "lesson": "x"},
                             {"id": "b", "mechanism": "m", "kind": "guardrail", "confidence": "Observed", "lesson": "y"}]}
        lessons, unknown, held = ls.select(index, [])
        self.assertEqual([l["id"] for l in lessons], ["b"])
        self.assertEqual(held, 1)


if __name__ == "__main__":
    unittest.main()
