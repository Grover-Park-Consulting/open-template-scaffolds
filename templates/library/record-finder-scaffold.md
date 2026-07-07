---
template: library-record-finder-scaffold
title: Record Finder for an Entry Form — VBA Scaffold
domain: library
type: vba-scaffold
version: 0.1.0
status: draft
implements: library-catalog-schema
requires_tables:
  - tblPublication
  - tblPublicationGenre
standards_layer:
  - design-principles
  - error-handling
  - query-style
  - naming-conventions
  - form-conventions
target_module: modRecordFinder
new_procedures:
  - FinderRowSource
  - WithAllRow
  - JumpFormToRecord
  - RewriteWhere
---

# Record Finder for an Entry Form — VBA Scaffold

## Intent

Realize the **record finder** the publications entry form defers to code — the pattern
`library-catalog-publication-form` names as a deferred dependency. On a data-entry form over a large
table, this is how a user gets to the *one* record they want to edit: a filter strip narrows a
pick-list dropdown three ways — by **category**, by **first letter (A–Z)**, or by **keyword** — and
choosing an entry **jumps the form to that record**, with a consistent **"show all"** option in every
mode.

This scaffold supplies the **procedure skeletons** — signatures, the row-source builder, the
server-side jump, and the error-handling frame. It does **not** write the domain logic: each
procedure marks where that goes (which columns are searched, the sentinel text, the primary key).
House style (the central error logger, how SQL is written, control names) is deferred to the
standards layer.

Three layers, kept distinct throughout:

- **`[SCAFFOLD]`** — structure provided here (signatures, the build/jump plumbing, error structure).
- **`[STANDARDS]`** — house style, deferred (`error-handling.md`, `query-style.md`, `naming-conventions.md`, `form-conventions.md`).
- **`[BUSINESS LOGIC]`** — the domain rule you fill in (columns, sentinel text, primary key).

## Prerequisites

| Object | Role |
|---|---|
| `library-catalog-schema` tables (`tblPublication`, `tblPublicationGenre`) | The finder queries these to build the pick list |
| `library-catalog-publication-form` | The paired form-spec; supplies the finder controls this engine drives (`cboCategory`, the A–Z sub-picker, `txtKeyword` + `cmdKeywordGo`, `cboFinder`) |
| The form's base record source | A saved query (or its SQL) whose `WHERE` `JumpFormToRecord` rewrites to land on one record |
| A sort-title column (e.g. `PubSort`) | The finder list and the jumped form order on it |
| A central error logger | `error-handling.md` (GPC default: `codearchive.GlblErrMsg`) |

## Procedures

Each procedure shows its scope, signature, and an annotated skeleton. The **`errHandler` block is
shown in full in `FinderRowSource` and referenced thereafter** — the VBE-reflection form
(`error-handling.md`) is identical in every procedure by design. The two pure string helpers
(`WithAllRow`, `RewriteWhere`) have **no** `errHandler` — nothing in them can raise a trappable data
error — and so, per `error-handling.md`, no line numbers. Line numbers are otherwise deliberately
absent (house-specific; see `error-handling.md`).

### FinderRowSource — `Public Function` → `String`

```vba
Public Function FinderRowSource(ByVal lCategoryID As Long, _
                                ByVal sInitialLetter As String, _
                                ByVal sKeywords As String) As String
    ' [SCAFFOLD] Build the finder dropdown's row source: the entries matching the active
    '            filter, plus an "<All ...>" row that always sorts first. Keyword text,
    '            when present, wins; otherwise filter by category + first letter.
    Dim sWhere As String

    On Error GoTo errHandler

    If Len(Trim$(sKeywords)) > 0 Then
        ' [BUSINESS LOGIC] keyword filter: split sKeywords on comma, trim and escape each
        '            term, OR together LIKE '*term*' over the searchable text columns
        '            (title, notes). An all-blank input yields "" -> caller falls back.
        ' >>> build sWhere per query-style.md <<<
        sWhere = vbNullString
    Else
        ' [BUSINESS LOGIC] category + first-letter filter over the sort-title column.
        '            lCategoryID = 0 means "all categories"; sInitialLetter = "*" means
        '            "all letters"; the four combinations give WHERE 1=1 / letter only /
        '            category only / category AND letter.
        ' >>> build sWhere per query-style.md <<<
        sWhere = vbNullString
    End If

    ' [BUSINESS LOGIC] core SELECT returns (RecordID, DisplayText, 1 AS SortIndex) filtered
    '            by sWhere; WithAllRow adds the sentinel + ordering.
    ' >>> FinderRowSource = WithAllRow( "SELECT DISTINCT ... " & sWhere ) per query-style.md <<<
    FinderRowSource = WithAllRow(vbNullString)

Cleanup:
    Exit Function

errHandler:
    ' [STANDARDS — error-handling.md] house-specific central logger, shown as a demo.
    '            GPC-private — you won't have it and shouldn't look for it; substitute your
    '            own logger or the dependency-free MsgBox block in error-handling.md.
    Call codearchive.GlblErrMsg(iLn:=Erl, _
        sFrm:=Application.VBE.ActiveCodePane.CodeModule, _
        sCtl:=Application.VBE.ActiveCodePane.CodeModule.ProcOfLine(Erl, 0), bLog:=True)
    Resume Cleanup
    Resume
End Function
```

### WithAllRow — `Private Function` → `String`

```vba
Private Function WithAllRow(ByVal sCoreSelect As String) As String
    ' [SCAFFOLD] Append the "<All ...>" sentinel row and the pick-list ordering to a core
    '            SELECT of (RecordID, DisplayText, 1 AS SortIndex) - the one home for the
    '            tail both filter paths share. Pure string builder: no errHandler.
    ' [BUSINESS LOGIC] the sentinel display text (e.g. "<All Publications>"), a 0 RecordID
    '            with SortIndex 0 (so it sorts first), and ORDER BY SortIndex, DisplayText.
    ' >>> WithAllRow = sCoreSelect & " UNION ALL SELECT ... ORDER BY ..." per query-style.md <<<
    WithAllRow = vbNullString
End Function
```

### JumpFormToRecord — `Public Sub`

```vba
Public Sub JumpFormToRecord(ByVal frm As Access.Form, ByVal lRecordID As Long)
    ' [SCAFFOLD] Narrow the form to the picked record - or show all when lRecordID = 0 -
    '            by rewriting the form's record-source WHERE server-side (only the wanted
    '            row is fetched). Uses RewriteWhere so a base query with no WHERE, or a
    '            multi-line one, is handled safely.
    Dim sWhere As String

    On Error GoTo errHandler

    ' [BUSINESS LOGIC] 0 => all rows; otherwise the primary key equals the picked id.
    ' >>> sWhere = IIf(lRecordID = 0, "1=1", "<PK> = " & lRecordID) <<<
    sWhere = IIf(lRecordID = 0, "1=1", vbNullString)

    ' [BUSINESS LOGIC] the base record source (saved-query name's SQL, or the form's own)
    '            and the sort column.
    ' >>> frm.RecordSource = RewriteWhere(baseSQL, sWhere, "<SortColumn>") per query-style.md <<<

Cleanup:
    Exit Sub

errHandler:
    ' [STANDARDS — error-handling.md] standard errHandler block (see FinderRowSource)
    Resume Cleanup
End Sub
```

### RewriteWhere — `Private Function` → `String`

```vba
Private Function RewriteWhere(ByVal sBaseSQL As String, _
                             ByVal sWhere As String, _
                             ByVal sOrderBy As String) As String
    ' [SCAFFOLD] Return the SELECT...FROM... head of sBaseSQL (everything before its first
    '            WHERE / GROUP BY / ORDER BY) with a fresh WHERE and ORDER BY appended.
    '            Pure string builder: no errHandler. This is the robust replacement for
    '            Left$(sql, InStr(sql, "WHERE") - 1), which crashes on a query with no
    '            WHERE and misreads a multi-line one.
    Dim sHead As String

    ' [STANDARDS — query-style.md] how the base SELECT is trimmed:
    ' >>> normalize CR/LF/Tab to spaces (Access stores query SQL multi-line); drop any
    '     trailing ";"; cut at the earliest of " WHERE " / " GROUP BY " / " ORDER BY "
    '     (space-delimited, case-insensitive, so an "...Order..." identifier is safe) <<<
    sHead = sBaseSQL

    RewriteWhere = sHead & _
        " WHERE " & Trim$(sWhere) & _
        IIf(Len(Trim$(sOrderBy)) > 0, " ORDER BY " & Trim$(sOrderBy), vbNullString)
End Function
```

## Form wiring

The engine is only useful wired to the form's events, and that is the point of the pattern: each
handler becomes a **thin delegate** — no domain logic, so (per `error-handling.md`) **no `errHandler`
and no line numbers**. Control names follow `form-conventions.md`; these are *not* part of
`new_procedures` (they live on the form, not in `modRecordFinder`).

```vba
' Category changed -> rebuild the list for the new category + the current letter.
Private Sub cboCategory_AfterUpdate()
    ' [BUSINESS LOGIC] the current A-Z letter is held per form-conventions (a TempVar or a
    '            form property); "*" until the A-Z picker sets one.
    Me.cboFinder.RowSource = FinderRowSource(Nz(Me.cboCategory, 0), "*", vbNullString)
    Me.cboFinder.Requery
End Sub

' A-Z letter chosen on the sub-picker -> rebuild for that letter (record it first).
Private Sub ApplyLetter(ByVal sLetter As String)
    Me.cboFinder.RowSource = FinderRowSource(Nz(Me.cboCategory, 0), sLetter, vbNullString)
    Me.cboFinder.Requery
End Sub

' Keyword "Go" -> keyword search (category/letter ignored while keywords are present).
Private Sub cmdKeywordGo_Click()
    Me.cboFinder.RowSource = FinderRowSource(0, "*", Nz(Me.txtKeyword, vbNullString))
    Me.cboFinder.Requery
End Sub

' An entry (or "<All ...>") picked -> jump the form to it.
Private Sub cboFinder_AfterUpdate()
    JumpFormToRecord Me, Nz(Me.cboFinder, 0)
End Sub
```

## Standards Layer

- **Error handling** — the `errHandler`/`Cleanup` structure, the central logger, and the
  line-number policy come from `error-handling.md`. The `GlblErrMsg` call shown is the GPC default
  (house-specific); a forked practice substitutes its own logger (or the dependency-free MsgBox
  block) and may number lines or not. The two pure helpers carry no handler by the same standard.
- **Query style** — every `>>> ... per query-style.md <<<` marker is SQL written to the house query
  standard: the `LIKE` keyword search and its term escaping, the `SELECT DISTINCT (RecordID,
  DisplayText, SortIndex)` shape, the `<All ...>` union, and the WHERE-trim in `RewriteWhere`.
- **Naming** — procedure, variable, and parameter names follow `naming-conventions.md`.
- **Form conventions** — the finder control names (`cboCategory`, `cboFinder`, `txtKeyword`,
  `cmdKeywordGo`, the A–Z sub-picker) and the thin-delegate handler rule follow `form-conventions.md`.

## Extra Options

*Empty in the base template. Filled per engagement; the filled copy is saved to the developer's own
library, not committed here.*

- **"Recently edited" list** — a most-recently-used shortcut alongside the finder, jumping straight
  to a recent record via `JumpFormToRecord`.
- **Multi-column finder** — show author / year beside the title in `cboFinder` for disambiguation.
- **Debounced keyword search** — search as the user pauses typing instead of on a `Go` click.

## Parked / future considerations (not in this design)

- **Server-side vs. client-side jump** — `JumpFormToRecord` refetches a single row server-side (best
  for a large, linked back end). A small, local record set could instead use the form's `Filter` /
  `FilterOn`, keeping the full set loaded; the choice is the adopter's, by data volume.
- **One owner for the base query** — `JumpFormToRecord` assumes a stable base record source. If the
  same saved query is rewritten from more than one place, a single owner (or a pristine base kept
  untouched) avoids surprises. In the source database this pattern came from, the same rewrite
  primitive was shared by a "recently edited" list and a related form — improved in place there.
