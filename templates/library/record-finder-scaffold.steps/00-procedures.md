---
step: 00-procedures
title: "The procedures this template writes"
platform_facts: []
route: both
---

**Who reads this:** the AI assistant, drafting the design.

## The procedures

These are this template's procedures as its code steps write them. Draft the design spec from them:
give each procedure the declaration below and say what it does. Draft no signature or behaviour of
your own. A change to either is the developer's decision: ask it before you present the design.

### FinderRowSource

```vba
Public Function FinderRowSource(ByVal lCategoryID As Long, _
                                ByVal sInitialLetter As String, _
                                ByVal sKeywords As String) As String
```

Build the finder dropdown's row source: the entries matching the active filter, plus an "<All ...>" row that always sorts first. Keyword text, when present, wins; otherwise filter by category + first letter.

### WithAllRow

```vba
Private Function WithAllRow(ByVal sCoreSelect As String) As String
```

Append the "<All ...>" sentinel row and the pick-list ordering to a core SELECT of (RecordID, DisplayText, 1 AS SortIndex) - the one home for the tail both filter paths share. Pure string builder: no errHandler.

### JumpFormToRecord

```vba
Public Sub JumpFormToRecord(ByVal frm As Access.Form, ByVal lRecordID As Long)
```

Narrow the form to the picked record - or show all when lRecordID = 0 - by rewriting the form's record-source WHERE server-side (only the wanted row is fetched). Uses RewriteWhere so a base query with no WHERE, or a multi-line one, is handled safely.

### RewriteWhere

```vba
Private Function RewriteWhere(ByVal sBaseSQL As String, _
                             ByVal sWhere As String, _
                             ByVal sOrderBy As String) As String
```

Return the SELECT...FROM... head of sBaseSQL (everything before its first WHERE / GROUP BY / ORDER BY) with a fresh WHERE and ORDER BY appended. Pure string builder: no errHandler. This is the robust replacement for Left$(sql, InStr(sql, "WHERE") - 1), which crashes on a query with no WHERE and misreads a multi-line one.
