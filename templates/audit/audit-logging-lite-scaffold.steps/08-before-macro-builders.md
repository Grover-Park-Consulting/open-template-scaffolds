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

1. **Stamp the house audit columns** (`standards/audit-columns.md`), on any table that carries them.
   Without this, a `Required` `CreatedBy` with nothing able to fill it **rejects every insert**.
2. **Stage Long Text old values** ahead of an update, by calling `BackupLongTextFieldsDM` for each
   Long Text field — the workaround for a Data Macro being unable to read `[Old].[LongTextField]`.

Both key on the same test, `IsNull([Old].[PK])` — true on an insert, false on an update — so they
merge into one conditional block with no conflict.

The function looks at what the table actually has and emits only what applies. A table carrying
neither the audit columns nor a Long Text field gets an **empty string back**, meaning no Before
Change macro at all — that table's set is just the three After macros.

**Why the branches are assembled before the XML:** a branch with no actions in it would produce an
empty `<Statements></Statements>` block, and we have no evidence Access accepts one. Rather than find
out on someone's live table, each branch's actions are built first; a branch with nothing to do is
left out, and if only the update branch has work the condition is inverted (`Not IsNull(...)`) so an
ordinary If-without-Else is emitted instead.

```vba
Private Function BuildBeforeChangeMacro(sTableName As String, fieldList As Collection, sPrimaryKeyField As String) As String
    ' [SCAFFOLD + STANDARDS — audit-columns.md] BeforeChange carries TWO jobs, because
    '            LoadFromText replaces a table's whole macro set and they cannot be
    '            loaded separately without one destroying the other:
    '              1. Stamp the house audit columns (any table that carries them).
    '              2. Stage Long Text old values ahead of an update (Long Text tables).
    '            Both key on the same discriminator, IsNull([Old].[PK]).
    '            Returns "" when the table needs neither — no BeforeChange macro at all.
    Dim sXml As String
    Dim sInsertActions As String
    Dim sUpdateActions As String
    Dim fieldInfo As Variant
    Dim sFieldName As String
    Dim bHasLongText As Boolean
    Dim bHasCreatedDate As Boolean
    Dim bHasCreatedBy As Boolean
    Dim bHasModifiedDate As Boolean
    Dim bHasModifiedBy As Boolean

    ' What does this table actually carry? The audit columns arrive with IsAuditable
    ' False (they are always seeded off) — presence is what decides stamping, not the flag.
    bHasLongText = False
    For Each fieldInfo In fieldList
        sFieldName = fieldInfo(0)
        If StrComp(sFieldName, AUDIT_CREATED_DATE, vbTextCompare) = 0 Then bHasCreatedDate = True
        If StrComp(sFieldName, AUDIT_CREATED_BY, vbTextCompare) = 0 Then bHasCreatedBy = True
        If StrComp(sFieldName, AUDIT_MODIFIED_DATE, vbTextCompare) = 0 Then bHasModifiedDate = True
        If StrComp(sFieldName, AUDIT_MODIFIED_BY, vbTextCompare) = 0 Then bHasModifiedBy = True
        If fieldInfo(1) = dbMemo And fieldInfo(3) = True Then bHasLongText = True
    Next fieldInfo

    ' ---- INSERT branch: stamp Created*, never touch Modified* ----
    If bHasCreatedDate Then sInsertActions = sInsertActions & AuditSetField(AUDIT_CREATED_DATE, "Now()")
    If bHasCreatedBy Then sInsertActions = sInsertActions & AuditSetField(AUDIT_CREATED_BY, "AuditUser()")

    If bHasLongText Then
        ' [BUSINESS LOGIC - schema Business Rule 4] Nothing to back up on an insert - there
        ' is no prior value. Null is the marker BackupLongTextFieldsDM tests for, and it is
        ' the only value that cannot be a real key: a primary key is Required by definition.
        ' The 0 this replaced was a legal Long Integer, Integer or Byte key value, so a row
        ' keyed 0 had its Long Text backup silently skipped. Measured: a Data Macro can set a
        ' local variable to Null, and VBA receives a genuine Null.
        sInsertActions = sInsertActions & "<Action Name=""SetLocalVar"">"
        sInsertActions = sInsertActions & "<Argument Name=""Name"">lngPKValue</Argument>"
        sInsertActions = sInsertActions & "<Argument Name=""Value"">Null</Argument>"
        sInsertActions = sInsertActions & "</Action>"
    End If

    ' ---- UPDATE branch: stamp Modified*, leave Created* frozen ----
    If bHasModifiedDate Then sUpdateActions = sUpdateActions & AuditSetField(AUDIT_MODIFIED_DATE, "Now()")
    If bHasModifiedBy Then sUpdateActions = sUpdateActions & AuditSetField(AUDIT_MODIFIED_BY, "AuditUser()")

    If bHasLongText Then
        sUpdateActions = sUpdateActions & "<Action Name=""SetLocalVar"">"
        sUpdateActions = sUpdateActions & "<Argument Name=""Name"">lngPKValue</Argument>"
        sUpdateActions = sUpdateActions & "<Argument Name=""Value"">=" & AuditKeyExpression("[" & sPrimaryKeyField & "]") & "</Argument>"
        sUpdateActions = sUpdateActions & "</Action>"

        sUpdateActions = sUpdateActions & "<Action Name=""SetLocalVar"">"
        sUpdateActions = sUpdateActions & "<Argument Name=""Name"">strTableName</Argument>"
        sUpdateActions = sUpdateActions & "<Argument Name=""Value"">""" & sTableName & """</Argument>"
        sUpdateActions = sUpdateActions & "</Action>"

        ' One backup call per Long Text field — a data macro CAN call a public VBA
        ' function in the same accdb; that is the hinge of the whole hybrid method.
        For Each fieldInfo In fieldList
            sFieldName = fieldInfo(0)
            If fieldInfo(1) = dbMemo And fieldInfo(3) = True Then
                sUpdateActions = sUpdateActions & "<Action Name=""SetLocalVar"">"
                sUpdateActions = sUpdateActions & "<Argument Name=""Name"">varLongTextBackup</Argument>"
                sUpdateActions = sUpdateActions & "<Argument Name=""Value"">BackupLongTextFieldsDM([strTableName],[lngPKValue],""" & sFieldName & """)</Argument>"
                sUpdateActions = sUpdateActions & "</Action>"
            End If
        Next fieldInfo
    End If

    ' Neither job applies to this table
    If Len(sInsertActions) = 0 And Len(sUpdateActions) = 0 Then
        BuildBeforeChangeMacro = ""
        Exit Function
    End If

    ' [SCAFFOLD] Emit only branches that have actions. An empty <Statements></Statements>
    '            is untested against Access, and a live table is the wrong place to find
    '            out — so when only the update branch has work, invert the condition and
    '            emit a plain If with no Else.
    sXml = "<DataMacro Event=""BeforeChange""><Statements>"
    sXml = sXml & "<Comment>" & AUDIT_MACRO_MARKER_FULL & " - regenerate rather than edit by hand.</Comment>"
    sXml = sXml & "<ConditionalBlock>"

    If Len(sInsertActions) > 0 Then
        sXml = sXml & "<If><Condition>IsNull([Old].[" & sPrimaryKeyField & "])</Condition>"
        sXml = sXml & "<Statements>" & sInsertActions & "</Statements></If>"
        If Len(sUpdateActions) > 0 Then
            sXml = sXml & "<Else><Statements>" & sUpdateActions & "</Statements></Else>"
        End If
    Else
        sXml = sXml & "<If><Condition>Not IsNull([Old].[" & sPrimaryKeyField & "])</Condition>"
        sXml = sXml & "<Statements>" & sUpdateActions & "</Statements></If>"
    End If

    sXml = sXml & "</ConditionalBlock></Statements></DataMacro>"

    BuildBeforeChangeMacro = sXml
End Function
```

### BuildBeforeDeleteMacro — `Private Function` → `String`

Emits the BeforeDelete fragment for a Long Text table: captures the PK, then stages each Long
Text value through `BackupLongTextFieldsDM` so the AfterDelete macro can log it. Returns an empty
string when the table has no Long Text fields.

```vba
Private Function BuildBeforeDeleteMacro(sTableName As String, fieldList As Collection, sPrimaryKeyField As String) As String
    ' [SCAFFOLD] BeforeDelete: stage Long Text values ahead of a delete
    '            (schema Business Rules 2-3).
    Dim sXml As String
    Dim fieldInfo As Variant
    Dim sFieldName As String
    Dim bHasLongText As Boolean

    bHasLongText = False
    For Each fieldInfo In fieldList
        If fieldInfo(1) = dbMemo And fieldInfo(3) = True Then
            bHasLongText = True
            Exit For
        End If
    Next fieldInfo
    If Not bHasLongText Then
        BuildBeforeDeleteMacro = ""
        Exit Function
    End If

    sXml = "<DataMacro Event=""BeforeDelete""><Statements>"
    sXml = sXml & "<Comment>" & AUDIT_MACRO_MARKER_FULL & " - regenerate rather than edit by hand.</Comment>"

    sXml = sXml & "<Action Name=""SetLocalVar"">"
    sXml = sXml & "<Argument Name=""Name"">lngPKValue</Argument>"
    sXml = sXml & "<Argument Name=""Value"">=" & AuditKeyExpression("[" & sPrimaryKeyField & "]") & "</Argument>"
    sXml = sXml & "</Action>"

    sXml = sXml & "<Action Name=""SetLocalVar"">"
    sXml = sXml & "<Argument Name=""Name"">strTableName</Argument>"
    sXml = sXml & "<Argument Name=""Value"">""" & sTableName & """</Argument>"
    sXml = sXml & "</Action>"

    For Each fieldInfo In fieldList
        sFieldName = fieldInfo(0)
        If fieldInfo(1) = dbMemo And fieldInfo(3) = True Then
            sXml = sXml & "<Action Name=""SetLocalVar"">"
            sXml = sXml & "<Argument Name=""Name"">varLongTextBackup</Argument>"
            sXml = sXml & "<Argument Name=""Value"">BackupLongTextFieldsDM([strTableName],[lngPKValue],""" & sFieldName & """)</Argument>"
            sXml = sXml & "</Action>"
        End If
    Next fieldInfo

    sXml = sXml & "</Statements></DataMacro>"

    BuildBeforeDeleteMacro = sXml
End Function
```

### AuditSetField — `Private Function` → `String`

One `SetField` action, in macro XML. `BuildBeforeChangeMacro` calls it four times — twice on the
insert branch and twice on the update branch — to stamp the house audit columns. It exists so the
XML for a stamp appears once rather than four times, which is what makes the stamping easy to read
against `standards/audit-columns.md`.

```vba
Private Function AuditSetField(ByVal sField As String, ByVal sValue As String) As String
    ' [STANDARDS — audit-columns.md] One SetField action for a stamped audit column.
    AuditSetField = "<Action Name=""SetField""><Argument Name=""Field"">" & sField & _
        "</Argument><Argument Name=""Value"">" & sValue & "</Argument></Action>"
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
Private Function AuditKeyExpression(ByVal sKeyRef As String) As String
    ' [BUSINESS LOGIC — schema Business Rule 4] Assigning a Replication ID bare into a text
    '                 column copies its sixteen bytes in as eight characters of unreadable
    '                 text: no error raised, the log row still written, and every check still
    '                 passing. Where the log's key column holds whole numbers the reference is
    '                 used as it stands, so the value keeps its type.
    If AUDIT_GUID_KEYS_AUDITED Then
        AuditKeyExpression = sKeyRef & " " & Chr(38) & "amp; " & Chr(34) & Chr(34)
    Else
        AuditKeyExpression = sKeyRef
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
what `AUDIT_GUID_KEYS_AUDITED` records. Whether *this* field is a Replication ID is a per-field
answer, and `IsUnauditableFieldType` does not exclude the type — so **a Replication ID that is
not a primary key is audited in every build, including one where `AUDIT_GUID_KEYS_AUDITED` is
`False`**. Gating this function on the build setting would leave that field wrong.

**The `&` is built from `Chr(38)`, not typed** — same rule and same reason as
`AuditKeyExpression`, and the same check after import.

```vba
Private Function AuditValueExpression(ByVal sFieldRef As String, ByVal lFldType As Long) As String
    ' [BUSINESS LOGIC — schema Business Rule 4] A field's value as it has to be written into
    '                 OldValue or NewValue. A Replication ID assigned bare into a text column
    '                 copies its sixteen bytes in as eight characters of unreadable text:
    '                 nothing raises an error and the log row is still written, so no check
    '                 catches it. Every other type renders correctly on its own, so only this
    '                 one is converted — asked per field, because a Replication ID that is not
    '                 a primary key is audited whatever AUDIT_GUID_KEYS_AUDITED says.
    If lFldType = dbGUID Then
        AuditValueExpression = sFieldRef & " " & Chr(38) & "amp; " & Chr(34) & Chr(34)
    Else
        AuditValueExpression = sFieldRef
    End If
End Function
```

### AuditKeyLocalVar — `Private Function` → `String`

**The one `SetLocalVar` action each After macro emits, and it has to come first.** It wraps
`AuditKeyExpression` in the action that puts the key into `AUDIT_KEY_LOCAL_VAR`, so the three
After builders each emit it once — above the per-field loop in the two that have one, because a
`LookUpRecord` inside that loop reads the variable and a variable set inside the loop would be
re-set per field.

`AfterInsert` has no `LookUpRecord` and could read the key field directly, but emits the variable
anyway: one macro shape across all three events is worth more than the action it saves.

```vba
Private Function AuditKeyLocalVar(ByVal sKeyRef As String) As String
    ' [SCAFFOLD] The SetLocalVar action every After macro emits before anything reads the key.
    '            It must come before the LookUpRecord that uses it.
    AuditKeyLocalVar = "<Action Name=""SetLocalVar""><Argument Name=""Name"">" & _
        AUDIT_KEY_LOCAL_VAR & "</Argument><Argument Name=""Value"">=" & _
        AuditKeyExpression(sKeyRef) & "</Argument></Action>"
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

**After importing this module, check that `sNotEqual` still reads
`Chr(38) & "lt;" & Chr(38) & "gt;"`, and repair it if the import changed it.** Do that silently —
it is your job, and the developer has no part to play in it. Say nothing to them about entity
escaping unless a build actually stops with error 3870; then explain what happened and that you
are fixing it.

```vba
Private Function GetComparisonExpression(sTableName As String, sFieldName As String, lFldType As Long) As String
    ' [SCAFFOLD] Per-type change test for the AfterUpdate conditional block.
    Dim sNotEqual As String
    ' Built from Chr(38), not typed as a literal &lt;&gt; — see the note above this block.
    ' This is XML for "<>"; the expression lives inside <Condition>...</Condition>.
    sNotEqual = Chr(38) & "lt;" & Chr(38) & "gt;"
    Select Case lFldType
        Case dbMemo
            ' [BUSINESS LOGIC — schema Business Rule 6] Long Text: the old value cannot be
            '                 read from [Old], so it is compared against the row staged in
            '                 tblLongTextBackup. The caller emits this inside the LookUpRecord
            '                 that reads that row, which is what makes [BackupRec] resolve.
            '                 NZ on both sides, exactly as below: a staged Null against a new
            '                 empty string would otherwise make the whole expression Null, the
            '                 condition would not pass, and a real edit would go unlogged.
            GetComparisonExpression = "StrComp(NZ([" & sTableName & "].[" & sFieldName & "],""""),NZ([BackupRec].[OldValue],""""),0)" & sNotEqual & "0"
        Case Else
            GetComparisonExpression = "StrComp(NZ([" & sTableName & "].[" & sFieldName & "],""""),NZ([Old].[" & sFieldName & "],""""),0)" & sNotEqual & "0"
    End Select
End Function
```
