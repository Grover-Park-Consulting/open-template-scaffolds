"""Tests for the compatibility check — the one code path that can touch a real
database.

Split across two layers on purpose:
- compat.check_db  — pure, imports pyodbc lazily, needs neither `mcp` nor
  `pyyaml`. Tested directly, so this module also runs under the 32-bit Python
  used with the x86 ACE driver (which has pyodbc but not mcp).
- server.check_compatibility — the MCP wrapper (template lookup + ok flag).
  Those tests self-skip when `mcp` isn't installed.

CI has no Access ODBC driver, so it exercises graceful degradation only: the
check must return available:false with a reason, never raise. The happy path
needs a real scan-extended Northwind .accdb and runs only when OTS_TEST_ACCDB
points at one (Windows, Python bitness matching the ACE driver — e.g.
`py -3.12-32` on George's box).
"""

import os
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import compat

try:
    import server
    _HAS_MCP = True
except ImportError:
    _HAS_MCP = False


def unwrap(tool):
    return tool.fn if hasattr(tool, "fn") else tool


STOCKTAKE_TABLES = ["Products", "Employees", "SystemSettings"]
STOCKTAKE_FIELDS = [
    "Products.ProductID",
    "Products.SKUBarCode",
    "Products.QuantityInPackage",
]


class TestCheckDbGraceful(unittest.TestCase):
    def test_unusable_database_degrades_gracefully(self):
        result = compat.check_db(
            r"Z:\no\such\place\nothing.accdb", STOCKTAKE_TABLES, STOCKTAKE_FIELDS
        )
        self.assertFalse(result["available"])
        self.assertTrue(result.get("reason"), "graceful failure must say why")


@unittest.skipUnless(_HAS_MCP, "mcp not installed in this Python environment")
class TestServerWrapper(unittest.TestCase):
    def test_unknown_template_raises(self):
        with self.assertRaises(ValueError):
            unwrap(server.check_compatibility)("no-such-template", "x.accdb")

    def test_wrapper_reports_unavailable_as_not_ok(self):
        result = unwrap(server.check_compatibility)(
            "northwind-stocktake-schema", r"Z:\no\such\place\nothing.accdb"
        )
        self.assertFalse(result["available"])
        self.assertFalse(result["ok"])
        self.assertEqual(result["template"], "northwind-stocktake-schema")


@unittest.skipUnless(
    os.environ.get("OTS_TEST_ACCDB"),
    "set OTS_TEST_ACCDB to a scan-extended Northwind .accdb to run",
)
class TestHappyPath(unittest.TestCase):
    def test_real_database_reports_requirements_present(self):
        result = compat.check_db(
            os.environ["OTS_TEST_ACCDB"], STOCKTAKE_TABLES, STOCKTAKE_FIELDS
        )
        self.assertTrue(result["available"], result)
        self.assertEqual(result["tables"]["missing"], [], result)
        self.assertEqual(result["fields"]["missing"], [], result)
        self.assertEqual(result["fields"]["table_absent"], [], result)


if __name__ == "__main__":
    unittest.main()
