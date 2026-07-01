# Naming Conventions — GPC Default Standards Layer

> **This is the GPC default standards layer.** When you fork the library, replace this file
> with your own house naming rules; the templates regenerate the same entities under whatever
> conventions you put here. Nothing in a template body depends on these specific names — that's
> the point of the standards/template split.
>
> Applies to **Microsoft Access** (local tables, queries, VBA objects) and **SQL Server**
> (tables, columns, programmable objects). Audit columns are covered in `audit-columns.md`;
> the VBA error-handling pattern in `error-handling.md`.

---

## 1. Database object naming

### 1.1 Access objects

| Object | Convention | Example |
|---|---|---|
| Local data/entity table | `tbl[Entity]` | `tblIndividual` |
| Lookup table | `tlkp[Entity]` | `tlkpPriority` |
| Junction table | `tbl[EntityA][EntityB]`, parent first | `tblAccessSiteOutlet` |
| ODBC-linked table | source table name, schema stripped (see note) | `tblHousehold` (from `dbo.tblHousehold`) |
| System/config table | `USys[Name]` | `USysAppConfig` |
| Main form | `frm[Entity]_[Purpose]` | `frmIndividual_Edit` |
| Subform | `sfrm[Entity]_[Purpose]` | `sfrmIndividual_Address` |
| Form record source | `qry[Entity]_frm` | `qryIndividual_frm` |
| Listbox source | `qry[Entity]_lst` | `qryIndividual_lst` |
| General query | `qry[Purpose]` | `qryActiveMembers` |
| Pass-through query | `qspt[Purpose]` | `qsptSelectedFoodItem_Lst` |
| Append / Update / Delete query | `qapd[Entity]` / `qupd[Entity]` / `qdel[Entity]` | `qapdInvoice` |
| Standard module | `mod[Purpose]` | `modAppSetup` |
| Callback/utility module | `bas[Purpose]` | `basRibbonCallbacks` |
| Class module | `[ClassName]` (PascalCase) | `AppConstants` |

Form and subform names lead with the **entity**, then an underscore, then the **purpose**
(`frm[Entity]_[Purpose]`) — so a form family groups together and multi-word names stay
unambiguous: `frmIndividual_Edit`, `frmIndividual_List`, `sfrmIndividual_Address`.

> **ODBC link names — strip the schema.** Link a SQL Server table under its bare table name
> (`dbo.tblHousehold` → `tblHousehold`); Access's namespace is flat, so the schema is dropped.
> **Collision tiebreaker** (two schemas exposing the same table name):
> - **`dbo` always keeps the bare name.** When the collision involves `dbo` and another schema,
>   `dbo`'s table links bare and the other appends its schema — `tblGenre` (from `dbo.tblGenre`)
>   and `tblGenre_Movies` (from `Movies.tblGenre`).
> - **Two non-`dbo` schemas:** neither gets the bare name; **both** append their schema —
>   `tblGenre_Movies` and `tblGenre_Catalog`.

### 1.2 SQL Server objects

| Object | Convention | Example |
|---|---|---|
| Data/entity table | `tbl[Entity]` (schema-qualified) | `dbo.tblProject` |
| Lookup table | `tlkp[Entity]` | `tlkpProjectArea` |
| Config table | `USys[Name]` | `USysAppConfig` |
| Junction table | `tbl[EntityA][EntityB]`, parent first | `tblAccessSiteOutlet` |
| View | `vw_[Purpose]` | `vw_RptProjectSummary` |
| Stored procedure | `usp_[Verb][Entity]` (never `sp_`) | `usp_AddUpdateBookcase` |
| Function | `fn_[Purpose]` (schema-qualified) | `dbo.fn_GetProjectSummary` |

Always schema-qualify table, view, procedure, and function references in views and stored procedures.

---

## 2. Lookup tables

A lookup table is defined **functionally, not structurally**: its primary role is to supply the
controlled set of values for a foreign key referenced by other tables, typically surfaced in a
combo/list box for data entry.

1. **Functional test (primary):** is the table referenced by FK from other tables, its main use
   being "populate a combo/list of valid values"? If yes, it's a lookup — regardless of size or
   growth rate.
2. **Shape test (confirmatory):** the classic shape is `[Entity]ID` + a single descriptor (e.g.
   `BrandName`) + optionally `SortOrder` and/or `[Entity]InactiveDate`. Extra columns (audit
   fields, "insert-if-not-exists" support) do **not** disqualify it.
3. **Disqualifier:** a table holding transactional facts (dates, costs, quantities tied to
   *events*) is not a lookup, even if it carries a Name-like field.

Lookup tables take the `tlkp` prefix. User-extensible lookups that grow via "insert if not exists"
are still lookups.

### Junction tables

Junction tables appear in **both Access and SQL Server**.

- Terminology: **junction table** (not bridge/link/associative).
- **Primary key:** a surrogate — **AutoNumber** (Access) / **`INT IDENTITY`** (SQL Server) — named
  `[EntityA][EntityB]ID` (e.g. `AccessSiteOutletID` on `tblAccessSiteOutlet`).
- **Two foreign-key columns**, named by the `[Entity]ID` convention, carrying the related pair.

---

## 3. Field / column naming

- **Primary key:** `[Entity]ID` — `IndividualID`, `ProjectID`. No `pk`/`fk` prefixes on columns.
- **Foreign key:** the `[ReferencedEntity]ID` of the PK it references — a FK to `tblProject.ProjectID`
  is `ProjectID` on the child. When ambiguous in a query, alias it (`p.ProjectID`); do not rename.
- **FK to a lookup table — qualify with the entity:** `ProjectStatusID`, `StepStatusID` — **never
  bare `StatusID`.** Consistency of the qualification pattern outranks the brevity of a bare FK name.
- All field names **PascalCase, no underscores**.

---

## 4. Qualified field names (the core rule)

**Never use an unqualified common English noun as a field name.** Always qualify with entity,
domain, or purpose. Two tiers:

**Tier 1 — Reserved words (hard prohibition).** Reserved in Access (Jet/ACE), T-SQL, or the
ODBC/ISO standard; bare use causes parser errors or ODBC failures on linked tables:
`Name`, `Date`, `Time`, `Timestamp`, `Description`, `Note`, `Text`, `Memo`, `Type`, `Value`,
`Number`, `Level`, `Key`, `Field`, `Group`, `Order`, `Index`, `Table`, `View`, `User`, `Schema`,
`Image`, `Password`, `Count`, `Size`.

**Tier 2 — Ambiguous common nouns (strong avoidance).** Not reserved, but semantically empty
without a qualifier (several are SQL Server future keywords): `Status`, `Category`, `Title`,
`Code`, `Label`, `Flag`, `Comment`, `Address`, `Amount`, `Total`, `Price`, `Active`, `State`,
`Role`, `Sequence`, `Result`, `Class`, `Source`, `Owner`, `Rank`, `Path`, `Reason`, `Body`,
`Subject`.

Qualify with entity or purpose: `ProjectStatus`, `ProductCategory`, `RecordLabel`,
`StreetAddress`, `CancellationReason`, `SortSequence`.

> **Test:** *Would a reader in a query result know what this field refers to without seeing the
> table name?* If no, qualify it.

This applies to Access local tables and SQL Server columns equally — especially for SQL Server
tables linked into Access via ODBC, where both reserved-word lists are active at once.

---

## 5. Data types

### 5.1 Access (local tables)

Use the Access type vocabulary: `AutoNumber`, `Long`, `Integer`, `Byte`, `Single`, `Double`,
`Currency`, `Text(n)`, `Memo`, `Date/Time`, `Boolean`. (This matches the field-spec vocabulary in
`templates/_template-schema.md` §5.)

### 5.2 SQL Server

- **Strings:** `NVARCHAR` universally. Use a defined maximum (`NVARCHAR(100)`, `NVARCHAR(500)`)
  where a practical bound exists; `NVARCHAR(MAX)` for indeterminate free text (notes, descriptions).
- **Dates:** `DATE` (date only); `DATETIME` (the GPC default for date+time — Access-safe);
  `DATETIME2` only when Access will never touch the table.
- **Identity PK:** `INT` (GPC databases don't approach the `BIGINT` threshold).

### 5.3 Index & constraint naming (SQL Server)

| Object | Pattern | Example |
|---|---|---|
| Primary key | `PK_tableName` | `PK_tblProject` |
| Foreign key | `FK_childTable_parentTable_columnName` | `FK_tblStep_tblProject_ProjectID` |
| Unique constraint | `UQ_tableName_columnName` | `UQ_tblProject_ProjectName` |
| Check constraint | `CHK_tableName_columnName` | `CHK_tblStep_StepStatusID` |
| Default constraint | `DF_tableName_columnName` | `DF_tblProject_CreatedDate` |
| Unique index | `UIX_tableName_leadingColumn` | `UIX_tblProject_ProjectName` |
| Non-clustered index | `IX_tableName_leadingColumn` | `IX_tblStep_ProjectID` |

Composite indexes are named by leading column only.

### 5.4 Column ordering in `CREATE TABLE`

1. **Identity PK** (`[Entity]ID`) — always first.
2. **Business/data columns** — identifying before qualifying (`ProjectName` before `ProjectDescription`).
3. **Audit columns** — always last (see `audit-columns.md`).
