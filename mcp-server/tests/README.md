# Tests — reference MCP server

Nothing to install beyond the server's own requirements: the suite uses Python's
built-in `unittest`. Run it from the repository root:

```powershell
python -m unittest discover -s mcp-server/tests
```

## What's covered

| File | Covers |
|---|---|
| `test_library.py` | Parsing helpers (`split_front_matter`, `iter_templates`, `read_standard`, `iter_standards`), every `validate()` rule via broken fixtures, `validate_library` (live corpus + slug uniqueness) |
| `test_tools.py` | The MCP tool surface: `list_templates`, `search_templates` (all relevance tiers), `get_template` (compose / missing-standard / unknown-id), `get_standards`, `validate` |
| `test_compat.py` | `check_db` graceful degradation (no driver → `available:false`, never a crash), the server wrapper, and an opt-in happy path against a real database |

## Fixtures

`fixtures/broken/` holds one deliberately broken template per validation rule —
each file names its rule in its filename and title. If you add a rule to
`validate_template`, add its fixture here and its expectation to
`EXPECTED_RULE` in `test_library.py` (a test fails if the two drift apart).

The live corpus under `templates/` is itself part of the suite: every shipped
template must validate clean.

## The real-database happy path

CI machines have no Access ODBC driver, so the compatibility check's success
path is opt-in:

```powershell
$env:OTS_TEST_ACCDB = 'C:\path\to\a\scan-extended-Northwind.accdb'
py -3.12-32 mcp-server/tests/test_compat.py
```

Use a Python whose bitness matches your installed ACE driver (32-bit Office →
32-bit Python). In that environment `mcp` may not be installed; the
server-wrapper tests skip themselves and the direct `check_db` tests still run.
