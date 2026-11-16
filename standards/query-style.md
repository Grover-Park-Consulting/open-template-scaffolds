# Query Style — OTS Default Standards Layer

**Who reads this:** the AI assistant, applying these rules to what it generates or a shop deciding
what to replace with their own.

If you are building something from a template, you don't need to read this file unless you are curious.
However, we have included clarifying comments to help you interpret what it says, just in case.

**Building from a template?** This file decides how the code asks the database for data:
where those requests are kept, and how they are written. What the requests themselves are called is
decided in [`naming-conventions.md`](naming-conventions.md).

You do not have to do anything with this information to use the templates. The rules apply to the code you receive
whether or not you read these files.

Using the standards as they are is the normal choice. If you came here because a template asked whether
you want to use these rules as they are or make them your own, you don't answer the question here.

*After you've read about the choice here, go back and answer it where it was asked.*

[`README.md`](README.md) lists all seven files if you want to see the others first.

> **This is the OTS default standards layer.** When you fork the library, replace this file with
> your own house query conventions. You can also add to them here; **do so carefully to avoid
> breaking the template.** It covers how VBA and saved queries write and
> run SQL — the house style a `vba-scaffold` (or any generated code that touches data) defers to via
> `standards_layer: [query-style]`. Naming of the query *objects* lives in `naming-conventions.md`;
> this file governs how the SQL *inside* them reads. Applies to Microsoft Access (DAO/Jet) and,
> where noted, SQL Server.

From here on, this file contains instructions for the AI.

**To the AI generating code:**

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
  `naming-conventions.md` §4 rule; the aliasing details live here).
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

  > **House divergence — `TempVars`, not form control references.** This layer drives these parameters from
  > `TempVars`, *not* from `Forms!frm!ctl` references (see the "All or One" combo/list-box selection
  > pattern). Setting the value in a `TempVar` **decouples the query from the form**: the query opens
  > and runs correctly whether or not the form is open. Mainstream Access practice reaches first for the
  > form control reference — a perfectly common approach, and one many prefer; a forked practice that
  > wants it simply swaps this rule. This layer defaults the other way *because* of the decoupling.

- **When the shape of the query varies** — optional criteria, a variable `IN (…)` list, dynamic columns
  or sort — **build the string**. Then:
  - Delimit every concatenated **value** correctly: strings in single quotes (double any embedded
    quote), dates in `#mm/dd/yyyy#` or via a formatting helper, numbers bare and validated.
  - **Whitelist any concatenated identifier** (field, table, or sort name) against known values — never
    pass user free-text as an identifier.

- **Pass-through** can't use Jet parameters — build the server SQL string with the same delimiting
  discipline. A genuinely *parameterized* server call needs **ADO**, which this standards layer
  reserves for rare, advanced cases.

---

## 5. Building the string

- **Optional criteria: start the `WHERE` with `1=1`** and append each condition with `AND`. Whether a
  condition is the first never needs deciding:

  ```vba
  strWhere = "1=1"
  If lngIndividualID > 0 Then strWhere = strWhere & " AND IndividualID = " & lngIndividualID
  strSQL = "SELECT * FROM tblIndividual WHERE " & strWhere
  ```

- **Line continuation limit.** The code editor allows about 25 `_` continuations in one statement. Past
  that, the import or compile fails with "too many line continuations" and names no line. A wide
  `INSERT` or `UPDATE` built one column per line passes it in the low twenties of columns. Plan for it
  before you hit it: collect the statement into several string variables, each well under the limit,
  and join them once at the end.

  ```vba
  strSQL1 = "INSERT INTO tblExample (Field01, Field02, ..." & _
            "Field20) "
  strSQL2 = "VALUES (..."
  strSQL = strSQL1 & strSQL2
  ```

  Splitting one statement into several `Execute` calls changes what runs; do not. A statement this wide
  by accident is better refactored: a loop over a field list, a saved parameter query, or a narrower
  table. Use the variables when the statement is wide by nature.
