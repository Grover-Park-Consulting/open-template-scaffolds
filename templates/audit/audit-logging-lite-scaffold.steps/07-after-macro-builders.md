---
step: 07-after-macro-builders
title: "The After Insert, After Update and After Delete macro builders"
platform_facts: [data-macro-rules]
---

**Who reads this:** the AI assistant, carrying out this step of the build.

## Procedures (continued)

### BuildAfterInsertMacro — `Private Function` → `String`

Emits the AfterInsert `<DataMacro>` fragment: one `tblAuditLog` row per configured field, marking
`OldValue` as `[NEW RECORD]`.

```vba
Private Function BuildAfterInsertMacro(sTableName As String, fieldList As Collection, sPrimaryKeyField As String) As String
    ' [SCAFFOLD] AfterInsert: log every configured field of the new row.
    Dim sXml As String
    Dim fieldInfo As Variant
    Dim sFieldName As String
    Dim lFldType As Long

    sXml = "<DataMacro Event=""AfterInsert""><Statements>"
    sXml = sXml & "<Comment>" & AUDIT_MACRO_MARKER_FULL & " - regenerate rather than edit by hand.</Comment>"

    ' [BUSINESS LOGIC - schema Business Rule 4] The key goes into one local variable and
    '                 everything below reads that, never the key field. See
    '                 AuditKeyLocalVar for why the key field cannot be read directly.
    If sPrimaryKeyField <> "" Then
        sXml = sXml & AuditKeyLocalVar("[" & sTableName & "].[" & sPrimaryKeyField & "]")
    End If

    For Each fieldInfo In fieldList
      If fieldInfo(3) = True Then    ' auditable fields only (schema Business Rule 5)
        sFieldName = fieldInfo(0)
        lFldType = fieldInfo(1)

        sXml = sXml & "<CreateRecord>"
        sXml = sXml & "<Data Alias=""NewAudit""><Reference>tblAuditLog</Reference></Data>"
        sXml = sXml & "<Statements>"

        sXml = sXml & "<Action Name=""SetField"">"
        sXml = sXml & "<Argument Name=""Field"">NewAudit.TableName</Argument>"
        sXml = sXml & "<Argument Name=""Value"">""" & sTableName & """</Argument>"
        sXml = sXml & "</Action>"

        If sPrimaryKeyField <> "" Then
            sXml = sXml & "<Action Name=""SetField"">"
            sXml = sXml & "<Argument Name=""Field"">NewAudit.PrimaryKey</Argument>"
            sXml = sXml & "<Argument Name=""Value"">[" & AUDIT_KEY_LOCAL_VAR & "]</Argument>"
            sXml = sXml & "</Action>"
        End If

        sXml = sXml & "<Action Name=""SetField"">"
        sXml = sXml & "<Argument Name=""Field"">NewAudit.FieldName</Argument>"
        sXml = sXml & "<Argument Name=""Value"">""" & sFieldName & """</Argument>"
        sXml = sXml & "</Action>"

        ' OperationType (schema Business Rule 6); OldValue stays Null on an insert
        sXml = sXml & "<Action Name=""SetField"">"
        sXml = sXml & "<Argument Name=""Field"">NewAudit.OperationType</Argument>"
        sXml = sXml & "<Argument Name=""Value"">""Insert""</Argument>"
        sXml = sXml & "</Action>"

        ' [BUSINESS LOGIC — schema Business Rule 4] Through AuditValueExpression: NewValue is
        '                 a text column, and this loop includes the primary key, so on a
        '                 Replication ID it is the key's own log row that goes in unreadable.
        sXml = sXml & "<Action Name=""SetField"">"
        sXml = sXml & "<Argument Name=""Field"">NewAudit.NewValue</Argument>"
        sXml = sXml & "<Argument Name=""Value"">" & _
            AuditValueExpression("[" & sTableName & "].[" & sFieldName & "]", lFldType) & "</Argument>"
        sXml = sXml & "</Action>"

        sXml = sXml & "<Action Name=""SetField"">"
        sXml = sXml & "<Argument Name=""Field"">NewAudit.DateChanged</Argument>"
        sXml = sXml & "<Argument Name=""Value"">Now()</Argument>"
        sXml = sXml & "</Action>"

        ' [STANDARDS / schema Business Rule 9] Identity — AuditUser() is the PREFERRED choice.
        '            The BeforeChange stamping macro calls it because audit-columns.md does,
        '            so the log names the same person the stamped row does.
        '            CurrentUser() is a named Extra Option: switch this and the three other
        '            ChangedBy sites together (AfterUpdate, AfterDelete, and
        '            BackupLongTextFieldsDM), and switch the stamping in your standards
        '            layer too — half a change puts two names on one edit.
        sXml = sXml & "<Action Name=""SetField"">"
        sXml = sXml & "<Argument Name=""Field"">NewAudit.ChangedBy</Argument>"
        sXml = sXml & "<Argument Name=""Value"">AuditUser()</Argument>"
        sXml = sXml & "</Action>"

        sXml = sXml & "</Statements></CreateRecord>"
      End If
    Next fieldInfo

    sXml = sXml & "</Statements></DataMacro>"
    BuildAfterInsertMacro = sXml
End Function
```

### BuildAfterUpdateMacro — `Private Function` → `String`

Emits the AfterUpdate fragment. Ordinary fields: log only when the value actually changed
(`GetComparisonExpression`), reading the old value from `[Old]`. Long Text fields: the same test,
against the value the BeforeChange macro staged in `tblLongTextBackup` (schema Business Rule 3) —
so the lookup that retrieves it wraps the test rather than sitting inside it.

```vba
Private Function BuildAfterUpdateMacro(sTableName As String, fieldList As Collection, sPrimaryKeyField As String) As String
    ' [SCAFFOLD] AfterUpdate: one conditional block per non-PK field; Long Text goes
    '            through the LookUpRecord retrieval path.
    Dim sXml As String
    Dim fieldInfo As Variant
    Dim sFieldName As String
    Dim lFldType As Long
    Dim bIsLongText As Boolean

    sXml = "<DataMacro Event=""AfterUpdate""><Statements>"
    sXml = sXml & "<Comment>" & AUDIT_MACRO_MARKER_FULL & " - regenerate rather than edit by hand.</Comment>"

    ' [BUSINESS LOGIC - schema Business Rule 4] Emitted once, above the per-field loop, so
    '                 every LookUpRecord below can read it. See AuditKeyLocalVar.
    If sPrimaryKeyField <> "" Then
        sXml = sXml & AuditKeyLocalVar("[" & sTableName & "].[" & sPrimaryKeyField & "]")
    End If

    For Each fieldInfo In fieldList
        sFieldName = fieldInfo(0)
        lFldType = fieldInfo(1)
        bIsLongText = (lFldType = dbMemo)

        ' Auditable fields only (schema Business Rule 5); PK never changes, skip it
        If fieldInfo(3) = True And sFieldName <> sPrimaryKeyField Then
            ' [BUSINESS LOGIC — schema Business Rule 3] Long Text: the staged old value is
            '                 fetched FIRST, because the change test compares against it. The
            '                 lookup wraps the test; the test does not wrap the lookup.
            If bIsLongText Then
                sXml = sXml & "<LookUpRecord>"
                sXml = sXml & "<Data Alias=""BackupRec"">"
                sXml = sXml & "<Reference>tblLongTextBackup</Reference>"
                sXml = sXml & "<WhereCondition>"
                sXml = sXml & "[tblLongTextBackup].[TableName]=""" & sTableName & """ And "
                sXml = sXml & "[tblLongTextBackup].[PrimaryKey]=[" & AUDIT_KEY_LOCAL_VAR & "] And "
                sXml = sXml & "[tblLongTextBackup].[FieldName]=""" & sFieldName & """"
                sXml = sXml & "</WhereCondition>"
                sXml = sXml & "</Data>"
                sXml = sXml & "<Statements>"
            End If

            sXml = sXml & "<ConditionalBlock><If>"
            sXml = sXml & "<Condition>" & GetComparisonExpression(sTableName, sFieldName, lFldType) & "</Condition>"
            sXml = sXml & "<Statements>"

            sXml = sXml & "<CreateRecord>"
            If bIsLongText Then
                sXml = sXml & "<Data><Reference>tblAuditLog</Reference></Data>"
            Else
                sXml = sXml & "<Data Alias=""NewAudit""><Reference>tblAuditLog</Reference></Data>"
            End If
            sXml = sXml & "<Statements>"

            sXml = sXml & "<Action Name=""SetField"">"
            If bIsLongText Then
                sXml = sXml & "<Argument Name=""Field"">tblAuditLog.TableName</Argument>"
            Else
                sXml = sXml & "<Argument Name=""Field"">NewAudit.TableName</Argument>"
            End If
            sXml = sXml & "<Argument Name=""Value"">""" & sTableName & """</Argument>"
            sXml = sXml & "</Action>"

            If sPrimaryKeyField <> "" Then
                sXml = sXml & "<Action Name=""SetField"">"
                If bIsLongText Then
                    sXml = sXml & "<Argument Name=""Field"">tblAuditLog.PrimaryKey</Argument>"
                Else
                    sXml = sXml & "<Argument Name=""Field"">NewAudit.PrimaryKey</Argument>"
                End If
                sXml = sXml & "<Argument Name=""Value"">[" & AUDIT_KEY_LOCAL_VAR & "]</Argument>"
                sXml = sXml & "</Action>"
            End If

            sXml = sXml & "<Action Name=""SetField"">"
            If bIsLongText Then
                sXml = sXml & "<Argument Name=""Field"">tblAuditLog.FieldName</Argument>"
            Else
                sXml = sXml & "<Argument Name=""Field"">NewAudit.FieldName</Argument>"
            End If
            sXml = sXml & "<Argument Name=""Value"">""" & sFieldName & """</Argument>"
            sXml = sXml & "</Action>"

            ' OperationType (schema Business Rule 6)
            sXml = sXml & "<Action Name=""SetField"">"
            If bIsLongText Then
                sXml = sXml & "<Argument Name=""Field"">tblAuditLog.OperationType</Argument>"
            Else
                sXml = sXml & "<Argument Name=""Field"">NewAudit.OperationType</Argument>"
            End If
            sXml = sXml & "<Argument Name=""Value"">""Update""</Argument>"
            sXml = sXml & "</Action>"

            ' OldValue — from the backup for Long Text, from [Old] otherwise. The backup
            ' column is already text; the [Old] reference goes through AuditValueExpression
            ' (schema Business Rule 4) because OldValue is a text column and the field may
            ' be a Replication ID that is not this table's key.
            sXml = sXml & "<Action Name=""SetField"">"
            If bIsLongText Then
                sXml = sXml & "<Argument Name=""Field"">tblAuditLog.OldValue</Argument>"
                sXml = sXml & "<Argument Name=""Value"">[BackupRec].[OldValue]</Argument>"
            Else
                sXml = sXml & "<Argument Name=""Field"">NewAudit.OldValue</Argument>"
                sXml = sXml & "<Argument Name=""Value"">" & _
                    AuditValueExpression("[Old].[" & sFieldName & "]", lFldType) & "</Argument>"
            End If
            sXml = sXml & "</Action>"

            sXml = sXml & "<Action Name=""SetField"">"
            If bIsLongText Then
                sXml = sXml & "<Argument Name=""Field"">tblAuditLog.NewValue</Argument>"
            Else
                sXml = sXml & "<Argument Name=""Field"">NewAudit.NewValue</Argument>"
            End If
            sXml = sXml & "<Argument Name=""Value"">" & _
                AuditValueExpression("[" & sTableName & "].[" & sFieldName & "]", lFldType) & "</Argument>"
            sXml = sXml & "</Action>"

            sXml = sXml & "<Action Name=""SetField"">"
            If bIsLongText Then
                sXml = sXml & "<Argument Name=""Field"">tblAuditLog.DateChanged</Argument>"
            Else
                sXml = sXml & "<Argument Name=""Field"">NewAudit.DateChanged</Argument>"
            End If
            sXml = sXml & "<Argument Name=""Value"">Now()</Argument>"
            sXml = sXml & "</Action>"

            ' [STANDARDS / schema Business Rule 9] identity — see BuildAfterInsertMacro
            ' [STANDARDS / schema Business Rule 9] AuditUser() preferred choice — same person as the
            '            stamp. CurrentUser() is an Extra Option; change all four ChangedBy
            '            sites together, never just one.
            sXml = sXml & "<Action Name=""SetField"">"
            If bIsLongText Then
                sXml = sXml & "<Argument Name=""Field"">tblAuditLog.ChangedBy</Argument>"
            Else
                sXml = sXml & "<Argument Name=""Field"">NewAudit.ChangedBy</Argument>"
            End If
            sXml = sXml & "<Argument Name=""Value"">AuditUser()</Argument>"
            sXml = sXml & "</Action>"

            sXml = sXml & "</Statements></CreateRecord>"
            sXml = sXml & "</Statements></If></ConditionalBlock>"

            If bIsLongText Then
                sXml = sXml & "</Statements></LookUpRecord>"
            End If
        End If
    Next fieldInfo

    sXml = sXml & "</Statements></DataMacro>"
    BuildAfterUpdateMacro = sXml
End Function
```

### BuildAfterDeleteMacro — `Private Function` → `String`

Emits the AfterDelete fragment: one `tblAuditLog` row per field, `NewValue` marked `[DELETED]`,
old values read from `[Old]` — except Long Text, retrieved from the backup the BeforeDelete macro
staged.

```vba
Private Function BuildAfterDeleteMacro(sTableName As String, fieldList As Collection, sPrimaryKeyField As String) As String
    ' [SCAFFOLD] AfterDelete: log every configured field of the deleted row.
    Dim sXml As String
    Dim fieldInfo As Variant
    Dim sFieldName As String
    Dim lFldType As Long
    Dim bIsLongText As Boolean

    sXml = "<DataMacro Event=""AfterDelete""><Statements>"
    sXml = sXml & "<Comment>" & AUDIT_MACRO_MARKER_FULL & " - regenerate rather than edit by hand.</Comment>"

    ' [BUSINESS LOGIC - schema Business Rule 4] The departed row is [Old] here. Measured on
    '                 AfterDelete as well as AfterUpdate. See AuditKeyLocalVar.
    If sPrimaryKeyField <> "" Then
        sXml = sXml & AuditKeyLocalVar("[Old].[" & sPrimaryKeyField & "]")
    End If

    For Each fieldInfo In fieldList
      If fieldInfo(3) = True Then    ' auditable fields only (schema Business Rule 5)
        sFieldName = fieldInfo(0)
        lFldType = fieldInfo(1)
        bIsLongText = (lFldType = dbMemo)

        If bIsLongText Then
            sXml = sXml & "<LookUpRecord>"
            sXml = sXml & "<Data Alias=""BackupRec"">"
            sXml = sXml & "<Reference>tblLongTextBackup</Reference>"
            sXml = sXml & "<WhereCondition>"
            sXml = sXml & "[tblLongTextBackup].[TableName]=""" & sTableName & """ And "
            sXml = sXml & "[tblLongTextBackup].[PrimaryKey]=[" & AUDIT_KEY_LOCAL_VAR & "] And "
            sXml = sXml & "[tblLongTextBackup].[FieldName]=""" & sFieldName & """"
            sXml = sXml & "</WhereCondition>"
            sXml = sXml & "</Data>"
            sXml = sXml & "<Statements>"
        End If

        sXml = sXml & "<CreateRecord>"
        If bIsLongText Then
            sXml = sXml & "<Data><Reference>tblAuditLog</Reference></Data>"
        Else
            sXml = sXml & "<Data Alias=""NewAudit""><Reference>tblAuditLog</Reference></Data>"
        End If
        sXml = sXml & "<Statements>"

        sXml = sXml & "<Action Name=""SetField"">"
        If bIsLongText Then
            sXml = sXml & "<Argument Name=""Field"">tblAuditLog.TableName</Argument>"
        Else
            sXml = sXml & "<Argument Name=""Field"">NewAudit.TableName</Argument>"
        End If
        sXml = sXml & "<Argument Name=""Value"">""" & sTableName & """</Argument>"
        sXml = sXml & "</Action>"

        If sPrimaryKeyField <> "" Then
            sXml = sXml & "<Action Name=""SetField"">"
            If bIsLongText Then
                sXml = sXml & "<Argument Name=""Field"">tblAuditLog.PrimaryKey</Argument>"
            Else
                sXml = sXml & "<Argument Name=""Field"">NewAudit.PrimaryKey</Argument>"
            End If
            sXml = sXml & "<Argument Name=""Value"">[" & AUDIT_KEY_LOCAL_VAR & "]</Argument>"
            sXml = sXml & "</Action>"
        End If

        sXml = sXml & "<Action Name=""SetField"">"
        If bIsLongText Then
            sXml = sXml & "<Argument Name=""Field"">tblAuditLog.FieldName</Argument>"
        Else
            sXml = sXml & "<Argument Name=""Field"">NewAudit.FieldName</Argument>"
        End If
        sXml = sXml & "<Argument Name=""Value"">""" & sFieldName & """</Argument>"
        sXml = sXml & "</Action>"

        ' OperationType (schema Business Rule 6); NewValue stays Null on a delete
        sXml = sXml & "<Action Name=""SetField"">"
        If bIsLongText Then
            sXml = sXml & "<Argument Name=""Field"">tblAuditLog.OperationType</Argument>"
        Else
            sXml = sXml & "<Argument Name=""Field"">NewAudit.OperationType</Argument>"
        End If
        sXml = sXml & "<Argument Name=""Value"">""Delete""</Argument>"
        sXml = sXml & "</Action>"

        ' [BUSINESS LOGIC — schema Business Rule 4] Through AuditValueExpression: OldValue is
        '                 a text column, this loop includes the primary key, and a delete is
        '                 the other place a Replication ID key reaches a value column.
        sXml = sXml & "<Action Name=""SetField"">"
        If bIsLongText Then
            sXml = sXml & "<Argument Name=""Field"">tblAuditLog.OldValue</Argument>"
            sXml = sXml & "<Argument Name=""Value"">[BackupRec].[OldValue]</Argument>"
        Else
            sXml = sXml & "<Argument Name=""Field"">NewAudit.OldValue</Argument>"
            sXml = sXml & "<Argument Name=""Value"">" & _
                AuditValueExpression("[Old].[" & sFieldName & "]", lFldType) & "</Argument>"
        End If
        sXml = sXml & "</Action>"

        sXml = sXml & "<Action Name=""SetField"">"
        If bIsLongText Then
            sXml = sXml & "<Argument Name=""Field"">tblAuditLog.DateChanged</Argument>"
        Else
            sXml = sXml & "<Argument Name=""Field"">NewAudit.DateChanged</Argument>"
        End If
        sXml = sXml & "<Argument Name=""Value"">Now()</Argument>"
        sXml = sXml & "</Action>"

        ' [STANDARDS / schema Business Rule 9] identity — see BuildAfterInsertMacro
        sXml = sXml & "<Action Name=""SetField"">"
        If bIsLongText Then
            sXml = sXml & "<Argument Name=""Field"">tblAuditLog.ChangedBy</Argument>"
        Else
            sXml = sXml & "<Argument Name=""Field"">NewAudit.ChangedBy</Argument>"
        End If
        ' [STANDARDS / schema Business Rule 9] AuditUser() — same person as the stamp.
        sXml = sXml & "<Argument Name=""Value"">AuditUser()</Argument>"
        sXml = sXml & "</Action>"

        sXml = sXml & "</Statements></CreateRecord>"

        If bIsLongText Then
            sXml = sXml & "</Statements></LookUpRecord>"
        End If
      End If
    Next fieldInfo

    sXml = sXml & "</Statements></DataMacro>"
    BuildAfterDeleteMacro = sXml
End Function
```
