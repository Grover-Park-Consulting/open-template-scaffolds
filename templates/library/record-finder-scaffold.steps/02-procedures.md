---
step: 02-procedures
title: "The procedures"
platform_facts: [mcp-module-import, mcp-line-numbers, vbe-line-continuation, form-live-build, form-hand-validation, mcp-run-procedure]
---

**Who reads this:** the AI assistant, carrying out this step of the build.

## Procedures

Each procedure shows its scope, signature, and an annotated skeleton. The **`errHandler` block is
shown in full in `FinderRowSource` and referenced thereafter** — the errHandler block's form
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
    ' [STANDARDS — error-handling.md] error reporting comes from the standards layer.
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation
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
    ' [SCAFFOLD] This assumes ONE owner for the base query it rewrites. Where the same saved
    '            query is also rewritten from somewhere else - a "recently edited" list, a
    '            related form - this is not a hypothetical to plan for later: it has actually
    '            happened in a real build of this pattern (see "One owner for the base query"
    '            under Parked). Check for other rewriters of the same query before building.
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

**Two limits on the `>>>` block above, neither one this function can detect on its own.**
Normalizing whitespace and searching for the earliest `WHERE`/`GROUP BY`/`ORDER BY` both scan the
whole SQL string as plain text, with no awareness of where a string literal starts or ends:

- **A tab or embedded line break inside a string literal in the base query gets normalized to a
  space along with everything else**, silently changing that literal's actual content. Rare in
  practice — most base queries carry no such literal — but real where one does.
- **A subquery in the base query's `FROM` clause that has its own `WHERE`/`GROUP BY`/`ORDER BY`
  gets cut at that inner keyword**, not at the outer query's own clause boundary, truncating the
  `FROM` clause mid-subquery and producing broken SQL.

Both are a property of scanning as plain text rather than parsing SQL, and this function is
deliberately the plain-text version — a real parser is more than this pattern needs for an ordinary
base query. Where the base query carries either case, check it before building: a literal with
embedded whitespace can usually be rewritten to avoid one, and a subquery-based record source is
outside what `RewriteWhere` safely handles as written.

## Form wiring

The engine is only useful wired to the form's events, and that is the point of the pattern: each
handler becomes a **thin delegate** — no domain logic, so (per `error-handling.md`) **no `errHandler`
and no line numbers**. Control names follow `form-conventions.md`; these are *not* part of
`new_procedures` (they live on the form, not in `modRecordFinder`).

**Six entry points touch the finder's state, not four.** The four below rebuild or act on the
pick-list itself. Two more, already on the paired `form-spec`'s own standard button set, have to
reset or release that state and are just as much a part of wiring the finder in as the other four —
leaving them out is how a form ends up showing a stale selection or a stale `TempVar`:

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

' New record started -> the finder's selection no longer reflects what's on screen; blank it
' rather than leave the last-picked entry showing against a record that isn't it.
Private Sub cmdNew_Click()
    Me.cboFinder = Null
End Sub

' Form closing -> release the TempVar the A-Z picker rode in on, so it doesn't leak a stale
' letter into the next time this form (or another one reading the same TempVar name) opens.
Private Sub Form_Close()
    TempVars.Remove "Alpha1stChar"
End Sub
```
