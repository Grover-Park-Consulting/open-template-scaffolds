"""template_status.py: the one command that records a decision about a template's text."""

import io
import sys
import tempfile
import unittest
from contextlib import redirect_stderr, redirect_stdout
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import library  # noqa: E402
import template_status as ts  # noqa: E402

MASTER = ("---\ntemplate: t\nversion: 0.4.6\nstatus: review\nstatus_fingerprint: 000000000000\n"
          "steps: [a]\n---\n\n# T\n\nBody.\n\n**Status last determined:** 2026-09-26.\n")


class TestPieces(unittest.TestCase):
    def test_fix_is_a_patch_and_change_is_a_minor(self):
        self.assertEqual(ts.bump("0.4.6", "fix"), "0.4.7")
        self.assertEqual(ts.bump("0.4.6", "change"), "0.5.0")

    def test_front_line_is_replaced_or_added_and_the_body_is_left_alone(self):
        text = ts.set_front_line(MASTER, "status", "stable")
        self.assertIn("\nstatus: stable\n", text)
        self.assertIn("Body.", text)
        bare = MASTER.replace("status_fingerprint: 000000000000\n", "")
        self.assertIn("status: review\nstatus_fingerprint: abc\n", ts.set_front_line(bare, "status_fingerprint", "abc"))

    def test_the_date_line_is_rewritten_in_both_spellings(self):
        new, found = ts.set_date(MASTER, "2027-01-02")
        self.assertTrue(found and "**Status last determined:** 2027-01-02." in new)
        new, found = ts.set_date("Status last determined: 2026-09-26.\n", "2027-01-02")
        self.assertTrue(found and new.startswith("Status last determined: 2027-01-02"))
        self.assertFalse(ts.set_date("No such line.\n", "2027-01-02")[1])

    def test_the_proposal_follows_the_record(self):
        now = "aaaaaaaaaaaa"
        rec = {"file": "r.md", "passed": 5, "failed": 0, "fingerprint": now}
        self.assertEqual(ts.propose("review", rec, now)[0], "stable")
        self.assertEqual(ts.propose("stable", {**rec, "fingerprint": "bbbbbbbbbbbb"}, now)[0], "review")
        self.assertEqual(ts.propose("stable", {**rec, "failed": 1}, now)[0], "draft")
        self.assertEqual(ts.propose("stable", {**rec, "fingerprint": None}, now)[0], "stable")
        self.assertEqual(ts.propose("draft", None, now)[0], "draft")


class TestEvidence(unittest.TestCase):
    def test_a_tally_line_is_not_counted_as_a_result(self):
        with tempfile.TemporaryDirectory() as d:
            (Path(d) / "t").mkdir()
            (Path(d) / "t" / "r.md").write_text(
                "Result: PASSED: 2\nResult: NOT PASSED: 0\n| 1 | a | Result: PASSED |\n"
                "| 2 | b | Result: PASSED |\n", encoding="utf-8")
            with patch.object(ts, "RECORDS", Path(d)):
                ev = ts.evidence("t")
            self.assertEqual((ev["passed"], ev["failed"]), (2, 0))


class TestCommand(unittest.TestCase):
    def run_cmd(self, d, *argv, record=None):
        d = Path(d)
        (d / "templates").mkdir(exist_ok=True)
        master = d / "templates" / "t.md"
        if not master.exists():
            master.write_bytes(MASTER.encode("utf-8"))
            (d / "templates" / "t.steps").mkdir()
            (d / "templates" / "t.steps" / "a.md").write_text("---\nstep: a\n---\nStep.\n", encoding="utf-8")
        if record is not None:
            (d / "records" / "t").mkdir(parents=True, exist_ok=True)
            (d / "records" / "t" / "2026-10-31-run.md").write_text(record, encoding="utf-8")
        out, err = io.StringIO(), io.StringIO()
        with patch.object(library, "TEMPLATES_DIR", d / "templates"), \
                patch.object(ts, "RECORDS", d / "records"), \
                patch.object(ts, "_git", lambda *a: ""), \
                redirect_stdout(out), redirect_stderr(err):
            code = ts.main(list(argv))
        return code, out.getvalue(), err.getvalue(), master

    def test_typo_rewrites_only_the_fingerprint(self):
        with tempfile.TemporaryDirectory() as d:
            code, out, _, master = self.run_cmd(d, "t", "typo")
            text = master.read_text(encoding="utf-8")
            self.assertEqual(code, 0)
            self.assertNotIn("000000000000", text)
            self.assertIn("version: 0.4.6", text)
            self.assertIn("status: review", text)
            self.assertIn("2026-09-26", text)

    def test_fix_without_status_prints_the_block_and_changes_nothing(self):
        with tempfile.TemporaryDirectory() as d:
            _, _, _, master = self.run_cmd(d, "t", "typo")
            before = master.read_bytes()
            code, out, _, _ = self.run_cmd(d, "t", "fix")
            self.assertEqual(code, 0)
            self.assertIn("v0.4.6 -> v0.4.7", out)
            self.assertIn("Proposed:", out)
            self.assertEqual(master.read_bytes(), before)

    def test_fix_with_status_bumps_dates_and_leaves_a_matching_fingerprint(self):
        with tempfile.TemporaryDirectory() as d:
            code, _, _, master = self.run_cmd(d, "t", "fix", "--status", "draft")
            text = master.read_text(encoding="utf-8")
            self.assertEqual(code, 0)
            self.assertIn("version: 0.4.7", text)
            self.assertIn("status: draft", text)
            self.assertNotIn("2026-09-26", text)
            front, _ = library.split_front_matter(text)
            steps = library.read_steps(master, front)
            self.assertEqual(library.fingerprint_errors("t", master, front, steps), [])

    def test_stable_is_refused_without_a_matching_record_and_allowed_by_hand(self):
        with tempfile.TemporaryDirectory() as d:
            code, _, err, master = self.run_cmd(d, "t", "change", "--status", "stable")
            self.assertEqual(code, 1)
            self.assertIn("stable refused", err)
            self.assertIn("version: 0.4.6", master.read_text(encoding="utf-8"))
            code, _, _, master = self.run_cmd(d, "t", "change", "--status", "stable", "--by-hand")
            self.assertEqual(code, 0)
            self.assertIn("version: 0.5.0", master.read_text(encoding="utf-8"))

    def test_a_record_with_this_texts_fingerprint_allows_stable(self):
        with tempfile.TemporaryDirectory() as d:
            _, out, _, _ = self.run_cmd(d, "t", "fix")
            # No record yet: read the fingerprint the bumped text will have from a dry run.
            ts_now = None
            with patch.object(library, "TEMPLATES_DIR", Path(d) / "templates"):
                for path, front, _b in library.iter_templates():
                    original = path.read_bytes()
                    path.write_bytes(ts.set_front_line(original.decode("utf-8"), "version", "0.4.7").encode("utf-8"))
                    ts_now = library.template_fingerprint(path, library.read_steps(path, front))
                    path.write_bytes(original)
            record = f"served fingerprint {ts_now}\n| 1 | a | Result: PASSED |\n"
            code, _, _, master = self.run_cmd(d, "t", "fix", "--status", "stable", record=record)
            self.assertEqual(code, 0)
            self.assertIn("status: stable", master.read_text(encoding="utf-8"))


class TestServedFingerprint(unittest.TestCase):
    def test_the_master_part_carries_the_templates_fingerprint(self):
        for path, front, _ in library.iter_templates():
            if str(front.get("type", "")).lower() == "spec":
                continue
            got = library.compose(str(front["template"]))
            steps = library.read_steps(path, front)
            self.assertEqual(got["served"]["fingerprint"], library.template_fingerprint(path, steps))
            return


if __name__ == "__main__":
    unittest.main()
