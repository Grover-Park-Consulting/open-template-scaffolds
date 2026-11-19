"""Tests for platform facts delivered in pieces: child sections and target platforms.

Run from the repo root:  python -m unittest discover -s mcp-server/tests
"""

import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import library

SAMPLE = """# Facts

### Table build
<!-- fact: build route: build -->

Intro: why the build is done this way.

#### Access sample
<!-- fact: build.sample target: access -->

The DAO sample.

#### Rules
<!-- fact: build.rules -->

The rules that make it run.

### SQL Server DDL
<!-- fact: ddl route: build target: sqlserver -->

CREATE TABLE text.

### Plain
<!-- fact: plain -->

A design fact.
"""


class TestSubFacts(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.TemporaryDirectory()
        path = Path(self.dir.name) / "_materialization.md"
        path.write_text(SAMPLE, encoding="utf-8")
        self.patch = patch.object(library, "MATERIALIZATION", path)
        self.patch.start()

    def tearDown(self):
        self.patch.stop()
        self.dir.cleanup()

    def test_children_inherit_route_and_parent_records_intro(self):
        facts, dupes = library.read_platform_facts()
        self.assertEqual(dupes, [])
        self.assertEqual(facts["build.rules"]["route"], "build")
        self.assertEqual(facts["build.sample"]["target"], "access")
        self.assertIn("Intro", facts["build"]["intro"])
        self.assertNotIn("DAO sample", facts["build"]["intro"])
        self.assertEqual(facts["build"]["children"], ["build.sample", "build.rules"])

    def test_parent_id_sends_the_whole_section(self):
        out, missing = library.resolve_facts(["build"])
        self.assertEqual(missing, [])
        self.assertIn("DAO sample", out[0]["content"])
        self.assertIn("rules that make it run", out[0]["content"])

    def test_child_id_sends_intro_and_that_child_only(self):
        out, _ = library.resolve_facts(["build.rules"])
        self.assertEqual(len(out), 1)
        self.assertIn("Intro", out[0]["content"])
        self.assertIn("rules that make it run", out[0]["content"])
        self.assertNotIn("DAO sample", out[0]["content"])
        self.assertEqual(out[0]["route"], "build")

    def test_two_children_arrive_as_one_entry_in_file_order(self):
        out, _ = library.resolve_facts(["build.rules", "build.sample"])
        self.assertEqual(len(out), 1)
        self.assertEqual(out[0]["id"], "build.sample, build.rules")
        self.assertEqual(out[0]["content"].count("Intro"), 1)

    def test_child_is_skipped_when_its_parent_is_declared(self):
        out, _ = library.resolve_facts(["build", "build.rules"])
        self.assertEqual([f["id"] for f in out], ["build"])

    def test_target_leaves_out_the_other_platform(self):
        out, _ = library.resolve_facts(["build", "ddl"], "sqlserver")
        self.assertEqual([f["id"] for f in out], ["build", "ddl"])
        self.assertNotIn("DAO sample", out[0]["content"])
        out, _ = library.resolve_facts(["build", "ddl"], "access")
        self.assertEqual([f["id"] for f in out], ["build"])
        self.assertIn("DAO sample", out[0]["content"])

    def test_no_target_sends_everything(self):
        out, _ = library.resolve_facts(["build", "ddl", "plain"])
        self.assertEqual(len(out), 3)
        self.assertIn("DAO sample", out[0]["content"])

    def test_unknown_id_is_reported(self):
        _, missing = library.resolve_facts(["build.nothing"])
        self.assertEqual(missing, ["build.nothing"])


class TestTargetOnParts(unittest.TestCase):
    def test_bad_target_is_refused(self):
        with self.assertRaises(ValueError):
            library.compose("stocktake-schema", "master", "", True, "oracle")

    def test_next_carries_the_target(self):
        got = library.compose("stocktake-schema", "master", "build", True, "access")
        self.assertEqual(got["next"].get("target"), "access")


if __name__ == "__main__":
    unittest.main()
