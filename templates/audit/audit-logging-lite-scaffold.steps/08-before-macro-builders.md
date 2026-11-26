---
step: 08-before-macro-builders
title: "The Before Change and Before Delete macro builders, and the expression helpers"
platform_facts: [data-macro-rules, vbe-line-continuation]
---

**Who reads this:** the AI assistant, carrying out this step of the build.

## Procedures (continued)

### BuildBeforeChangeMacro — `Private Function` → `String`

**This one macro does two jobs, and it has to.** `LoadFromText` replaces a table's *entire* Data
Macro set rather than adding to it, so the two things that both need to happen on a Before Change —
stamping the house audit columns, and staging Long Text values before they're overwritten — cannot
be loaded as separate macros. The second load would silently delete the first. They are generated
together instead:

1. **Stamp the house audit columns** (`standards/audit-columns.md`), on the tables in
   `AuditStampList` that carry them. On those tables nothing else fills them, so without this a
   `Required` `CreatedBy` **rejects every insert**. A table whose columns a default or the
   database's own code fills is not listed, and its columns are left to that.
2. **Stage Long Text old values** ahead of an update, by calling `BackupLongTextFieldsDM` for each
   Long Text field — the workaround for a Data Macro being unable to read `[Old].[LongTextField]`.

Both key on the same test, `IsNull([Old].[PK])` — true on an insert, false on an update — so they
merge into one conditional block with no conflict.

The function looks at what the table actually has and emits only what applies. A table with
neither stamping to do nor a Long Text field gets an **empty string back**, meaning no Before
Change macro at all — that table's set is just the three After macros.

**Why the branches are assembled before the XML:** a branch with no actions in it would produce an
empty `<Statements></Statements>` block, and we have no evidence Access accepts one. Rather than find
out on someone's live table, each branch's actions are built first; a branch with nothing to do is
left out, and if only the update branch has work the condition is inverted (`Not IsNull(...)`) so an
ordinary If-without-Else is emitted instead.

```vba
Private Function BuildBeforeChangeMacro(strTableName As String, fieldList As Collection, strPrimaryKeyField As String) As String
    ' [SCAFFOLD + STANDARDS — audit-columns.md] BeforeChange carries TWO jobs, because
    '            LoadFromText replaces a table's whole macro set and they cannot be
    '            loaded separately without one destroying the other:
    '              1. Stamp the house audit columns (tables in AuditStampList).
    '              2. Stage Long Text old values ahead of an update (Long Text tables).
    '            Both key on the same discriminator, IsNull([Old].[PK]).
    '            Returns "" when the table needs neither — no BeforeChange macro at all.
    Dim strXml As String
    Dim strInsertActions As String
    Dim strUpdateActions As String
    Dim fieldInfo As Variant
    Dim strFieldName As String
    Dim booHasLongText As Boolean
    Dim booHasCreatedDate As Boolean
    Dim booHasCreatedBy As Boolean
    Dim booHasModifiedDate As Boolean
    Dim booHasModifiedBy As Boolean
    Dim booStamp As Boolean

    ' [STANDARDS — audit-columns.md] Stamp only where AuditStampList says the generator
    '            fills the columns. Matching column names alone would take over columns a
    '            default or the database's own code already fills, and overwrite them.
    '            Within a listed table, presence decides which of the four are stamped;
    '            they arrive with IsAuditable False (always seeded off), so not the flag.
    booStamp = IsNamedInList(AuditStampList, strTableName)
    booHasLongText = False
    For Each fieldInfo In fieldList
        strFieldName = fieldInfo(0)
        If booStamp Then
            If StrComp(strFieldName, AuditCreatedDate, vbTextCompare) = 0 Then booHasCreatedDate = True
            If StrComp(strFieldName, AuditCreatedBy, vbTextCompare) = 0 Then booHasCreatedBy = True
            If StrComp(strFieldName, AuditModifiedDate, vbTextCompare) = 0 Then booHasModifiedDate = True
            If StrComp(strFieldName, AuditModifiedBy, vbTextCompare) = 0 Then booHasModifiedBy = True
        End If
        If fieldInfo(1) = dbMemo And fieldInfo(3) = True Then booHasLongText = True
    Next fieldInfo

    ' ---- INSERT branch: stamp Created*, never touch Modified* ----
    If booHasCreatedDate Then strInsertActions = strInsertActions & AuditSetField(AuditCreatedDate, "Now()")
    If booHasCreatedBy Then strInsertActions = strInsertActions & AuditSetField(AuditCreatedBy, "AuditUser()")

    If booHasLongText Then
        ' [BUSINESS LOGIC - schema Business Rule 4] Nothing to back up on an insert - there
        ' is no prior value. Null is the marker BackupLongTextFieldsDM tests for, and it is
        ' the only value that cannot be a real key: a primary key is Required by definition.
        ' The 0 this replaced was a legal Long Integer, Integer or Byte key value, so a row
        ' keyed 0 had its Long Text backup silently skipped. Measured: a Data Macro can set a
        ' local variable to Null, and VBA receives a genuine Null.
        strInsertActions = strInsertActions & "<Action Name=""SetLocalVar"">"
        strInsertActions = strInsertActions & "<Argument Name=""Name"">lngPKValue</Argument>"
        strInsertActions = strInsertActions & "<Argument Name=""Value"">Null</Argument>"
        strInsertActions = strInsertActions & "</Action>"
    End If

    ' ---- UPDATE branch: stamp Modified*, leave Created* frozen ----
    If booHasModifiedDate Then strUpdateActions = strUpdateActions & AuditSetField(AuditModifiedDate, "Now()")
    If booHasModifiedBy Then strUpdateActions = strUpdateActions & AuditSetField(AuditModifiedBy, "AuditUser()")

    If booHasLongText Then
        strUpdateActions = strUpdateActions & "<Action Name=""SetLocalVar"">"
        strUpdateActions = strUpdateActions & "<Argument Name=""Name"">lngPKValue</Argument>"
        strUpdateActions = strUpdateActions & "<Argument Name=""Value"">=" & AuditKeyExpression("[" & strPrimaryKeyField & "]") & "</Argument>"
        strUpdateActions = strUpdateActions & "</Action>"

        strUpdateActions = strUpdateActions & "<Action Name=""SetLocalVar"">"
        strUpdateActions = strUpdateActions & "<Argument Name=""Name"">strTableName</Argument>"
        strUpdateActions = strUpdateActions & "<Argument Name=""Value"">""" & strTableName & """</Argument>"
        strUpdateActions = strUpdateActions & "</Action>"

        ' One backup call per Long Text field — a data macro CAN call a public VBA
        ' function in the same accdb; that is the hinge of the whole hybrid method.
        For Each fieldInfo In fieldList
            strFieldName = fieldInfo(0)
            If fieldInfo(1) = dbMemo And fieldInfo(3) = True Then
                strUpdateActions = strUpdateActions & "<Action Name=""SetLocalVar"">"
                strUpdateActions = strUpdateActions & "<Argument Name=""Name"">varLongTextBackup</Argument>"
                strUpdateActions = strUpdateActions & "<Argument Name=""Value"">BackupLongTextFieldsDM([strTableName],[lngPKValue],""" & strFieldName & """)</Argument>"
                strUpdateActions = strUpdateActions & "</Action>"
                strUpdateActions = strUpdateActions & AuditRefusalBlock("varLongTextBackup")
            End If
        Next fieldInfo
    End If

    ' Neither job applies to this table
    If Len(strInsertActions) = 0 And Len(strUpdateActions) = 0 Then
        BuildBeforeChangeMacro = ""
        Exit Function
    End If

    ' [SCAFFOLD] Emit only branches that have actions. An empty <Statements></Statements>
    '            is untested against Access, and a live table is the wrong place to find
    '            out — so when only the update branch has work, invert the condition and
    '            emit a plain If with no Else.
    strXml = "<DataMacro Event=""BeforeChange""><Statements>"
    strXml = strXml & "<Comment>" & AuditMacroMarkerFull & " - regenerate rather than edit by hand.</Comment>"
    strXml = strXml & "<ConditionalBlock>"

    If Len(strInsertActions) > 0 Then
        strXml = strXml & "<If><Condition>IsNull([Old].[" & strPrimaryKeyField & "])</Condition>"
        strXml = strXml & "<Statements>" & strInsertActions & "</Statements></If>"
        If Len(strUpdateActions) > 0 Then
            strXml = strXml & "<Else><Statements>" & strUpdateActions & "</Statements></Else>"
        End If
    Else
        strXml = strXml & "<If><Condition>Not IsNull([Old].[" & strPrimaryKeyField & "])</Condition>"
        strXml = strXml & "<Statements>" & strUpdateActions & "</Statements></If>"
    End If

    strXml = strXml & "</ConditionalBlock></Statements></DataMacro>"

    BuildBeforeChangeMacro = strXml
End Function
```

### BuildBeforeDeleteMacro — `Private Function` → `String`

Emits the BeforeDelete fragment for a Long Text table: captures the PK, then stages each Long
Text value through `BackupLongTextFieldsDM` so the AfterDelete macro can log it. Returns an empty
string when the table has no Long Text fields.

```vba
Private Function BuildBeforeDeleteMacro(strTableName As String, fieldList As Collection, strPrimaryKeyField As String) As String
    ' [SCAFFOLD] BeforeDelete: stage Long Text values ahead of a delete
    '            (schema Business Rules 2-3).
    Dim strXml As String
    Dim fieldInfo As Variant
    Dim strFieldName As String
    Dim booHasLongText As Boolean

    booHasLongText = False
    For Each fieldInfo In fieldList
        If fieldInfo(1) = dbMemo And fieldInfo(3) = True Then
            booHasLongText = True
            Exit For
        End If
    Next fieldInfo
    If Not booHasLongText Then
        BuildBeforeDeleteMacro = ""
        Exit Function
    End If

    strXml = "<DataMacro Event=""BeforeDelete""><Statements>"
    strXml = strXml & "<Comment>" & AuditMacroMarkerFull & " - regenerate rather than edit by hand.</Comment>"

    strXml = strXml & "<Action Name=""SetLocalVar"">"
    strXml = strXml & "<Argument Name=""Name"">lngPKValue</Argument>"
    strXml = strXml & "<Argument Name=""Value"">=" & AuditKeyExpression("[" & strPrimaryKeyField & "]") & "</Argument>"
    strXml = strXml & "</Action>"

    strXml = strXml & "<Action Name=""SetLocalVar"">"
    strXml = strXml & "<Argument Name=""Name"">strTableName</Argument>"
    strXml = strXml & "<Argument Name=""Value"">""" & strTableName & """</Argument>"
    strXml = strXml & "</Action>"

    For Each fieldInfo In fieldList
        strFieldName = fieldInfo(0)
        If fieldInfo(1) = dbMemo And fieldInfo(3) = True Then
            strXml = strXml & "<Action Name=""SetLocalVar"">"
            strXml = strXml & "<Argument Name=""Name"">varLongTextBackup</Argument>"
            strXml = strXml & "<Argument Name=""Value"">BackupLongTextFieldsDM([strTableName],[lngPKValue],""" & strFieldName & """)</Argument>"
            strXml = strXml & "</Action>"
            strXml = strXml & AuditRefusalBlock("varLongTextBackup")
        End If
    Next fieldInfo

    strXml = strXml & "</Statements></DataMacro>"

    BuildBeforeDeleteMacro = strXml
End Function
```

### AuditSetField — `Private Function` → `String`

One `SetField` action, in macro XML. `BuildBeforeChangeMacro` calls it four times — twice on the
insert branch and twice on the update branch — to stamp the house audit columns. It exists so the
XML for a stamp appears once rather than four times, which is what makes the stamping easy to read
against `standards/audit-columns.md`.

```vba
Private Function AuditSetField(ByVal strField As String, ByVal strValue As String) As String
    ' [STANDARDS — audit-columns.md] One SetField action for a stamped audit column.
    AuditSetField = "<Action Name=""SetField""><Argument Name=""Field"">" & strField & _
        "</Argument><Argument Name=""Value"">" & strValue & "</Argument></Action>"
End Function
```

### AuditKeyExpression — `Private Function` → `String`

**The audited row's key goes into one local variable, and everything downstream reads that
variable rather than the key field.** Both After macros that retrieve a staged Long Text value
find it with a `LookUpRecord`, and **a Replication ID named inside that `LookUpRecord`'s
`WhereCondition` stops the condition filtering at all** — every clause, not just the one naming
the key — so the lookup returns the first row of the staging table and the log records another
row's data as this row's history. Measured, on both `AfterUpdate` and `AfterDelete`. No string
conversion written into the condition avoids it; the key field simply cannot appear there.

Setting the variable first, with the conversion forced, keeps the key field out of the condition
and was measured to return the right row against three decoys — one differing only in table, one
only in field, one only in key.

**Every build emits the variable, whatever the key type**, so there is one macro shape rather
than two. Verified for whole-number keys as well: a Long key carried in the variable and compared
against a Long column filters exactly as the direct reference it replaces.

**The `&` is built from `Chr(38)`, not typed** — the same rule and the same reason as
`GetComparisonExpression` below. A literal `&amp;` in this source is one un-escaping importer away
from becoming a bare `&`, which breaks the macro XML the moment it loads. Building it from a
character code means no `&`-entity ever exists in the source for a tool to touch. The doubled
quotes come from `Chr(34)` for the same reason, and because counting five consecutive quote
characters in VBA source is a mistake waiting to happen.

**After importing this module, check that `AuditKeyExpression` still reads
`Chr(38) & "amp; " & Chr(34) & Chr(34)`, and repair it if the import changed it.** Do that
silently — it is your job, and the developer has no part in it.

```vba
Private Function AuditKeyExpression(ByVal strKeyRef As String) As String
    ' [BUSINESS LOGIC — schema Business Rule 4] Assigning a Replication ID bare into a text
    '                 column copies its sixteen bytes in as eight characters of unreadable
    '                 text: no error raised, the log row still written, and every check still
    '                 passing. Where the log's key column holds whole numbers the reference is
    '                 used as it stands, so the value keeps its type.
    If AuditGuidKeysAudited Then
        AuditKeyExpression = strKeyRef & " " & Chr(38) & "amp; " & Chr(34) & Chr(34)
    Else
        AuditKeyExpression = strKeyRef
    End If
End Function
```

### AuditValueExpression — `Private Function` → `String`

**The log's before-and-after columns are text too, and they take a Replication ID by the same
route the key column does.** `AuditKeyExpression` above covers the column holding the audited
row's key. It is not the only text column a Replication ID can reach: `OldValue` and `NewValue`
are Long Text, and every audited field's value passes through them.

**This one asks the field's own type, where `AuditKeyExpression` asks the build setting.** The
two questions are different. Whether the key column holds text is a build-wide answer, which is
what `AuditGuidKeysAudited` records. Whether *this* field is a Replication ID is a per-field
answer, and `IsUnauditableFieldType` does not exclude the type — so **a Replication ID that is
not a primary key is audited in every build, including one where `AuditGuidKeysAudited` is
`False`**. Gating this function on the build setting would leave that field wrong.

**The `&` is built from `Chr(38)`, not typed** — same rule and same reason as
`AuditKeyExpression`, and the same check after import.

```vba
Private Function AuditValueExpression(ByVal strFieldRef As String, ByVal lngFldType As Long) As String
    ' [BUSINESS LOGIC — schema Business Rule 4] A field's value as it has to be written into
    '                 OldValue or NewValue. A Replication ID assigned bare into a text column
    '                 copies its sixteen bytes in as eight characters of unreadable text:
    '                 nothing raises an error and the log row is still written, so no check
    '                 catches it. Every other type renders correctly on its own, so only this
    '                 one is converted — asked per field, because a Replication ID that is not
    '                 a primary key is audited whatever AuditGuidKeysAudited says.
    If lngFldType = dbGUID Then
        AuditValueExpression = strFieldRef & " " & Chr(38) & "amp; " & Chr(34) & Chr(34)
    Else
        AuditValueExpression = strFieldRef
    End If
End Function
```

### AuditKeyLocalVar — `Private Function` → `String`

**The one `SetLocalVar` action each After macro emits, and it has to come first.** It wraps
`AuditKeyExpression` in the action that puts the key into `AuditKeyVarName`, so the three
After builders each emit it once — above the per-field loop in the two that have one, because a
`LookUpRecord` inside that loop reads the variable and a variable set inside the loop would be
re-set per field.

`AfterInsert` has no `LookUpRecord` and could read the key field directly, but emits the variable
anyway: one macro shape across all three events is worth more than the action it saves.

```vba
Private Function AuditKeyLocalVar(ByVal strKeyRef As String) As String
    ' [SCAFFOLD] The SetLocalVar action every After macro emits before anything reads the key.
    '            It must come before the LookUpRecord that uses it.
    AuditKeyLocalVar = "<Action Name=""SetLocalVar""><Argument Name=""Name"">" & _
        AuditKeyVarName & "</Argument><Argument Name=""Value"">=" & _
        AuditKeyExpression(strKeyRef) & "</Argument></Action>"
End Function
```

### AuditRefusalBlock — `Private Function` → `String`

**Refuses the save with the sentence `BackupLongTextFieldsDM` returned, when there is one**
(`error-handling.md`, the refusing shape). `BackupLongTextFieldsDM` returns `""` when the old value was
staged, so the condition passes only for a sentence. The `<>` is built from `Chr(38)`, for the same
reason as in `GetComparisonExpression`.

```vba
Private Function AuditRefusalBlock(ByVal strVarName As String) As String
    ' [SCAFFOLD] A ConditionalBlock that raises the local variable's sentence, if it holds one.
    Dim strNotEqual As String
    strNotEqual = Chr(38) & "lt;" & Chr(38) & "gt;"
    AuditRefusalBlock = "<ConditionalBlock><If><Condition>[" & strVarName & "]" & strNotEqual & _
        """""</Condition><Statements><Action Name=""RaiseError"">" & _
        "<Argument Name=""Number"">1</Argument>" & _
        "<Argument Name=""Description"">=[" & strVarName & "]</Argument>" & _
        "</Action></Statements></If></ConditionalBlock>"
End Function
```

### GetComparisonExpression — `Private Function` → `String`

The change test the AfterUpdate macro embeds per field. Ordinary fields compare old and new with
`StrComp` over `Nz`-wrapped values; Long Text returns `True` — always log — because the macro
cannot read the old value to compare (schema Business Rule 6).

**The not-equal operator is built from `Chr(38)`, not typed as a literal `&lt;&gt;`.** A real
build hit error 3870 ("Microsoft Access cannot interpret the text you are pasting as a data
macro") on every table because the VBA source was imported through a tool that **un-escapes HTML
entities on the way in** — a literal `&lt;&gt;` in this function's source became a raw `<>` once
imported, which breaks the macro XML the moment it's loaded (the `<` is read as an opening tag).
Assembling the entity from character codes at runtime means no literal `&`-entity ever exists in
the source for an importer to touch — the string is only ever `&lt;&gt;` in memory, never in text
anyone's tool re-reads. See `templates/_materialization.md` for the general rule this follows.

**After importing this module, check that `strNotEqual` still reads
`Chr(38) & "lt;" & Chr(38) & "gt;"`, and repair it if the import changed it.** Do that silently —
it is your job, and the developer has no part to play in it. Say nothing to them about entity
escaping unless a build actually stops with error 3870; then explain what happened and that you
are fixing it.

```vba
Private Function GetComparisonExpression(strTableName As String, strFieldName As String, lngFldType As Long) As String
    ' [SCAFFOLD] Per-type change test for the AfterUpdate conditional block.
    Dim strNotEqual As String
    ' Built from Chr(38), not typed as a literal &lt;&gt; — see the note above this block.
    ' This is XML for "<>"; the expression lives inside <Condition>...</Condition>.
    strNotEqual = Chr(38) & "lt;" & Chr(38) & "gt;"
    Select Case lngFldType
        Case dbMemo
            ' [BUSINESS LOGIC — schema Business Rule 6] Long Text: the old value cannot be
            '                 read from [Old], so it is compared against the row staged in
            '                 tblLongTextBackup. The caller emits this inside the LookUpRecord
            '                 that reads that row, which is what makes [BackupRec] resolve.
            '                 NZ on both sides, exactly as below: a staged Null against a new
            '                 empty string would otherwise make the whole expression Null, the
            '                 condition would not pass, and a real edit would go unlogged.
            GetComparisonExpression = "StrComp(NZ([" & strTableName & "].[" & strFieldName & "],""""),NZ([BackupRec].[OldValue],""""),0)" & strNotEqual & "0"
        Case Else
            GetComparisonExpression = "StrComp(NZ([" & strTableName & "].[" & strFieldName & "],""""),NZ([Old].[" & strFieldName & "],""""),0)" & strNotEqual & "0"
    End Select
End Function
```
