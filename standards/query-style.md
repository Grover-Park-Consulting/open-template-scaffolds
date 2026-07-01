# Query Style — GPC Default Standards Layer

> **GPC default; fork-and-replace.** How VBA and saved queries write and run SQL — the house style a
> `vba-scaffold` (or any generated code that touches data) defers to via `standards_layer: [query-style]`.
> A forked practice swaps this file for its own query conventions. Naming of the query *objects* lives
> in `naming-conventions.md`; this file governs how the SQL *inside* them reads. Applies to Microsoft
> Access (DAO/Jet) and, where noted, SQL Server.

---

## 1. Where SQL lives

- **Prefer saved QueryDefs** over SQL strings built in VBA, wherever the query is reusable — they're
  named, optimized, and visible (`qry` / `qspt` / `qapd` / `qupd` / `qdel` per `naming-conventions.md`).
- **Inline SQL** in VBA is for genuinely dynamic statements assembled at runtime; keep it in a single
  `sSql` string, built per this file's rules.
- **Pass-through (`qspt`)** to run **complex SQL directly on the server** (SQL Server), where the server
  engine executes it — faster than pulling rows into Access to process, and the only way to reach
  server-side features. Jet parameters don't apply; the SQL is sent as-is.
- Execute action SQL with `db.Execute strSQL, dbFailOnError [+ dbSeeChanges]`.

---

## 2. Aliasing and qualification

- In any **multi-table** query, give each table a short, meaningful **alias** and **qualify every
  column** with it (`p.ProductName`, `s.ScanQuantity`).
- Do **not rename** a column to disambiguate it — alias the table instead (this is the
  `naming-conventions.md` §3 rule; the aliasing details live here).
- Single-table queries need no alias.
- On **SQL Server**, schema-qualify objects (`dbo.tblProduct`) — see `naming-conventions.md` §1.2.

---

## 3. Formatting

- **SQL keywords UPPERCASE** (`SELECT`, `FROM`, `INNER JOIN`, `WHERE`); object names as cased.
- **Explicit `JOIN ... ON`** — never join in the `WHERE` clause.
- One major clause per line (`SELECT` / `FROM` / `JOIN` / `WHERE` / `ORDER BY`) for readability in
  multi-line strings.

---

## 4. Building criteria

The rule of thumb: **parameterize or safely delimit values; concatenate only structure.**

- **When only the values vary** (the query's shape is fixed): prefer a **saved parameter query**, with
  the parameters driven by **`TempVars`** set at the form level.

  > **House divergence — `TempVars`, not form control references.** GPC drives these parameters from
  > `TempVars`, *not* from `Forms!frm!ctl` references (see the "All or One" combo/list-box selection
  > pattern). Setting the value in a `TempVar` **decouples the query from the form**: the query opens
  > and runs correctly whether or not the form is open. Mainstream Access practice reaches first for the
  > form control reference — a perfectly common approach, and one many prefer; a forked practice that
  > wants it simply swaps this rule. GPC defaults the other way *because* of the decoupling.

- **When the shape of the query varies** — optional criteria, a variable `IN (…)` list, dynamic columns
  or sort — **build the string**. Then:
  - Delimit every concatenated **value** correctly: strings in single quotes (double any embedded
    quote), dates in `#mm/dd/yyyy#` or via a formatting helper, numbers bare and validated.
  - **Whitelist any concatenated identifier** (field, table, or sort name) against known values — never
    pass user free-text as an identifier.

- **Pass-through** can't use Jet parameters — build the server SQL string with the same delimiting
  discipline. A genuinely *parameterized* server call needs **ADO**, which GPC reserves for rare,
  advanced cases.
