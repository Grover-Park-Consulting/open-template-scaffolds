---
step: 07-after-macro-builders
title: "The After Insert, After Update and After Delete macro builders"
platform_facts: [data-macro-rules]
---

**Who reads this:** the AI assistant, carrying out this step of the build.

## Procedures (continued)

### BuildAfterInsertMacro — `Private Function` → `String`

Emits the AfterInsert `<DataMacro>` fragment: one `tblAuditLog` row per configured field, leaving
`OldValue` Null, as schema Business Rule 6 says.

```vba
Private Function BuildAfterInsertMacro(strTableName As String, fieldList As Collection, strPrimaryKeyField As String) As String
    ' [SCAFFOLD] AfterInsert: log every configured field of the new row.
    Dim strXml As String
    Dim fieldInfo As Variant
    Dim strFieldName As String
    Dim lngFldType As Long

    strXml = "<DataMacro Event=""AfterInsert""><Statements>"
    strXml = strXml & "<Comment>" & AUDIT_MACRO_MARKER_FULL & " - regenerate rather than edit by hand.</Comment>"

    ' [BUSINESS LOGIC - schema Business Rule 4] The key goes into one local variable and
    '                 everything below reads that, never the key field. See
    '                 AuditKeyLocalVar for why the key field cannot be read directly.
    If strPrimaryKeyField <> "" Then
        strXml = strXml & AuditKeyLocalVar("[" & strTableName & "].[" & strPrimaryKeyField & "]")
    End If

    For Each fieldInfo In fieldList
      If fieldInfo(3) = True Then    ' auditable fields only (schema Business Rule 5)
        strFieldName = fieldInfo(0)
        lngFldType = fieldInfo(1)

        strXml = strXml & "<CreateRecord>"
        strXml = strXml & "<Data Alias=""NewAudit""><Reference>tblAuditLog</Reference></Data>"
        strXml = strXml & "<Statements>"

        strXml = strXml & "<Action Name=""SetField"">"
        strXml = strXml & "<Argument Name=""Field"">NewAudit.TableName</Argument>"
        strXml = strXml & "<Argument Name=""Value"">""" & strTableName & """</Argument>"
        strXml = strXml & "</Action>"

        If strPrimaryKeyField <> "" Then
            strXml = strXml & "<Action Name=""SetField"">"
            strXml = strXml & "<Argument Name=""Field"">NewAudit.PrimaryKey</Argument>"
            strXml = strXml & "<Argument Name=""Value"">[" & AUDIT_KEY_LOCAL_VAR & "]</Argument>"
            strXml = strXml & "</Action>"
        End If

        strXml = strXml & "<Action Name=""SetField"">"
        strXml = strXml & "<Argument Name=""Field"">NewAudit.FieldName</Argument>"
        strXml = strXml & "<Argument Name=""Value"">""" & strFieldName & """</Argument>"
        strXml = strXml & "</Action>"

        ' OperationType (schema Business Rule 6); OldValue stays Null on an insert
        strXml = strXml & "<Action Name=""SetField"">"
        strXml = strXml & "<Argument Name=""Field"">NewAudit.OperationType</Argument>"
        strXml = strXml & "<Argument Name=""Value"">""Insert""</Argument>"
        strXml = strXml & "</Action>"

        ' [BUSINESS LOGIC — schema Business Rule 4] Through AuditValueExpression: NewValue is
        '                 a text column, and this loop includes the primary key, so on a
        '                 Replication ID it is the key's own log row that goes in unreadable.
        strXml = strXml & "<Action Name=""SetField"">"
        strXml = strXml & "<Argument Name=""Field"">NewAudit.NewValue</Argument>"
        strXml = strXml & "<Argument Name=""Value"">" & _
            AuditValueExpression("[" & strTableName & "].[" & strFieldName & "]", lngFldType) & "</Argument>"
        strXml = strXml & "</Action>"

        strXml = strXml & "<Action Name=""SetField"">"
        strXml = strXml & "<Argument Name=""Field"">NewAudit.DateChanged</Argument>"
        strXml = strXml & "<Argument Name=""Value"">Now()</Argument>"
        strXml = strXml & "</Action>"

        ' [STANDARDS / schema Business Rule 9] Identity — AuditUser() is the PREFERRED choice.
        '            The BeforeChange stamping macro calls it because audit-columns.md does,
        '            so the log names the same person the stamped row does.
        '            CurrentUser() is a named Extra Option: switch this and the three other
        '            ChangedBy sites together (AfterUpdate, AfterDelete, and
        '            BackupLongTextFieldsDM), and switch the stamping in your standards
        '            layer too — half a change puts two names on one edit.
        strXml = strXml & "<Action Name=""SetField"">"
        strXml = strXml & "<Argument Name=""Field"">NewAudit.ChangedBy</Argument>"
        strXml = strXml & "<Argument Name=""Value"">AuditUser()</Argument>"
        strXml = strXml & "</Action>"

        strXml = strXml & "</Statements></CreateRecord>"
      End If
    Next fieldInfo

    strXml = strXml & "</Statements></DataMacro>"
    BuildAfterInsertMacro = strXml
End Function
```

### BuildAfterUpdateMacro — `Private Function` → `String`

Emits the AfterUpdate fragment. Ordinary fields: log only when the value actually changed
(`GetComparisonExpression`), reading the old value from `[Old]`. Long Text fields: the same test,
against the value the BeforeChange macro staged in `tblLongTextBackup` (schema Business Rule 3) —
so the lookup that retrieves it wraps the test rather than sitting inside it.

```vba
Private Function BuildAfterUpdateMacro(strTableName As String, fieldList As Collection, strPrimaryKeyField As String) As String
    ' [SCAFFOLD] AfterUpdate: one conditional block per non-PK field; Long Text goes
    '            through the LookUpRecord retrieval path.
    Dim strXml As String
    Dim fieldInfo As Variant
    Dim strFieldName As String
    Dim lngFldType As Long
    Dim booIsLongText As Boolean

    strXml = "<DataMacro Event=""AfterUpdate""><Statements>"
    strXml = strXml & "<Comment>" & AUDIT_MACRO_MARKER_FULL & " - regenerate rather than edit by hand.</Comment>"

    ' [BUSINESS LOGIC - schema Business Rule 4] Emitted once, above the per-field loop, so
    '                 every LookUpRecord below can read it. See AuditKeyLocalVar.
    If strPrimaryKeyField <> "" Then
        strXml = strXml & AuditKeyLocalVar("[" & strTableName & "].[" & strPrimaryKeyField & "]")
    End If

    For Each fieldInfo In fieldList
        strFieldName = fieldInfo(0)
        lngFldType = fieldInfo(1)
        booIsLongText = (lngFldType = dbMemo)

        ' Auditable fields only (schema Business Rule 5); PK never changes, skip it
        If fieldInfo(3) = True And strFieldName <> strPrimaryKeyField Then
            ' [BUSINESS LOGIC — schema Business Rule 3] Long Text: the staged old value is
            '                 fetched FIRST, because the change test compares against it. The
            '                 lookup wraps the test; the test does not wrap the lookup.
            If booIsLongText Then
                strXml = strXml & "<LookUpRecord>"
                strXml = strXml & "<Data Alias=""BackupRec"">"
                strXml = strXml & "<Reference>tblLongTextBackup</Reference>"
                strXml = strXml & "<WhereCondition>"
                strXml = strXml & "[tblLongTextBackup].[TableName]=""" & strTableName & """ And "
                strXml = strXml & "[tblLongTextBackup].[PrimaryKey]=[" & AUDIT_KEY_LOCAL_VAR & "] And "
                strXml = strXml & "[tblLongTextBackup].[FieldName]=""" & strFieldName & """"
                strXml = strXml & "</WhereCondition>"
                strXml = strXml & "</Data>"
                strXml = strXml & "<Statements>"
            End If

            strXml = strXml & "<ConditionalBlock><If>"
            strXml = strXml & "<Condition>" & GetComparisonExpression(strTableName, strFieldName, lngFldType) & "</Condition>"
            strXml = strXml & "<Statements>"

            strXml = strXml & "<CreateRecord>"
            If booIsLongText Then
                strXml = strXml & "<Data><Reference>tblAuditLog</Reference></Data>"
            Else
                strXml = strXml & "<Data Alias=""NewAudit""><Reference>tblAuditLog</Reference></Data>"
            End If
            strXml = strXml & "<Statements>"

            strXml = strXml & "<Action Name=""SetField"">"
            If booIsLongText Then
                strXml = strXml & "<Argument Name=""Field"">tblAuditLog.TableName</Argument>"
            Else
                strXml = strXml & "<Argument Name=""Field"">NewAudit.TableName</Argument>"
            End If
            strXml = strXml & "<Argument Name=""Value"">""" & strTableName & """</Argument>"
            strXml = strXml & "</Action>"

            If strPrimaryKeyField <> "" Then
                strXml = strXml & "<Action Name=""SetField"">"
                If booIsLongText Then
                    strXml = strXml & "<Argument Name=""Field"">tblAuditLog.PrimaryKey</Argument>"
                Else
                    strXml = strXml & "<Argument Name=""Field"">NewAudit.PrimaryKey</Argument>"
                End If
                strXml = strXml & "<Argument Name=""Value"">[" & AUDIT_KEY_LOCAL_VAR & "]</Argument>"
                strXml = strXml & "</Action>"
            End If

            strXml = strXml & "<Action Name=""SetField"">"
            If booIsLongText Then
                strXml = strXml & "<Argument Name=""Field"">tblAuditLog.FieldName</Argument>"
            Else
                strXml = strXml & "<Argument Name=""Field"">NewAudit.FieldName</Argument>"
            End If
            strXml = strXml & "<Argument Name=""Value"">""" & strFieldName & """</Argument>"
            strXml = strXml & "</Action>"

            ' OperationType (schema Business Rule 6)
            strXml = strXml & "<Action Name=""SetField"">"
            If booIsLongText Then
                strXml = strXml & "<Argument Name=""Field"">tblAuditLog.OperationType</Argument>"
            Else
                strXml = strXml & "<Argument Name=""Field"">NewAudit.OperationType</Argument>"
            End If
            strXml = strXml & "<Argument Name=""Value"">""Update""</Argument>"
            strXml = strXml & "</Action>"

            ' OldValue — from the backup for Long Text, from [Old] otherwise. The backup
            ' column is already text; the [Old] reference goes through AuditValueExpression
            ' (schema Business Rule 4) because OldValue is a text column and the field may
            ' be a Replication ID that is not this table's key.
            strXml = strXml & "<Action Name=""SetField"">"
            If booIsLongText Then
                strXml = strXml & "<Argument Name=""Field"">tblAuditLog.OldValue</Argument>"
                strXml = strXml & "<Argument Name=""Value"">[BackupRec].[OldValue]</Argument>"
            Else
                strXml = strXml & "<Argument Name=""Field"">NewAudit.OldValue</Argument>"
                strXml = strXml & "<Argument Name=""Value"">" & _
                    AuditValueExpression("[Old].[" & strFieldName & "]", lngFldType) & "</Argument>"
            End If
            strXml = strXml & "</Action>"

            strXml = strXml & "<Action Name=""SetField"">"
            If booIsLongText Then
                strXml = strXml & "<Argument Name=""Field"">tblAuditLog.NewValue</Argument>"
            Else
                strXml = strXml & "<Argument Name=""Field"">NewAudit.NewValue</Argument>"
            End If
            strXml = strXml & "<Argument Name=""Value"">" & _
                AuditValueExpression("[" & strTableName & "].[" & strFieldName & "]", lngFldType) & "</Argument>"
            strXml = strXml & "</Action>"

            strXml = strXml & "<Action Name=""SetField"">"
            If booIsLongText Then
                strXml = strXml & "<Argument Name=""Field"">tblAuditLog.DateChanged</Argument>"
            Else
                strXml = strXml & "<Argument Name=""Field"">NewAudit.DateChanged</Argument>"
            End If
            strXml = strXml & "<Argument Name=""Value"">Now()</Argument>"
            strXml = strXml & "</Action>"

            ' [STANDARDS / schema Business Rule 9] identity — see BuildAfterInsertMacro
            ' [STANDARDS / schema Business Rule 9] AuditUser() preferred choice — same person as the
            '            stamp. CurrentUser() is an Extra Option; change all four ChangedBy
            '            sites together, never just one.
            strXml = strXml & "<Action Name=""SetField"">"
            If booIsLongText Then
                strXml = strXml & "<Argument Name=""Field"">tblAuditLog.ChangedBy</Argument>"
            Else
                strXml = strXml & "<Argument Name=""Field"">NewAudit.ChangedBy</Argument>"
            End If
            strXml = strXml & "<Argument Name=""Value"">AuditUser()</Argument>"
            strXml = strXml & "</Action>"

            strXml = strXml & "</Statements></CreateRecord>"
            strXml = strXml & "</Statements></If></ConditionalBlock>"

            If booIsLongText Then
                strXml = strXml & "</Statements></LookUpRecord>"
            End If
        End If
    Next fieldInfo

    strXml = strXml & "</Statements></DataMacro>"
    BuildAfterUpdateMacro = strXml
End Function
```

### BuildAfterDeleteMacro — `Private Function` → `String`

Emits the AfterDelete fragment: one `tblAuditLog` row per field, `NewValue` left Null (Business Rule 6),
old values read from `[Old]` — except Long Text, retrieved from the backup the BeforeDelete macro
staged.

```vba
Private Function BuildAfterDeleteMacro(strTableName As String, fieldList As Collection, strPrimaryKeyField As String) As String
    ' [SCAFFOLD] AfterDelete: log every configured field of the deleted row.
    Dim strXml As String
    Dim fieldInfo As Variant
    Dim strFieldName As String
    Dim lngFldType As Long
    Dim booIsLongText As Boolean

    strXml = "<DataMacro Event=""AfterDelete""><Statements>"
    strXml = strXml & "<Comment>" & AUDIT_MACRO_MARKER_FULL & " - regenerate rather than edit by hand.</Comment>"

    ' [BUSINESS LOGIC - schema Business Rule 4] The departed row is [Old] here. Measured on
    '                 AfterDelete as well as AfterUpdate. See AuditKeyLocalVar.
    If strPrimaryKeyField <> "" Then
        strXml = strXml & AuditKeyLocalVar("[Old].[" & strPrimaryKeyField & "]")
    End If

    For Each fieldInfo In fieldList
      If fieldInfo(3) = True Then    ' auditable fields only (schema Business Rule 5)
        strFieldName = fieldInfo(0)
        lngFldType = fieldInfo(1)
        booIsLongText = (lngFldType = dbMemo)

        If booIsLongText Then
            strXml = strXml & "<LookUpRecord>"
            strXml = strXml & "<Data Alias=""BackupRec"">"
            strXml = strXml & "<Reference>tblLongTextBackup</Reference>"
            strXml = strXml & "<WhereCondition>"
            strXml = strXml & "[tblLongTextBackup].[TableName]=""" & strTableName & """ And "
            strXml = strXml & "[tblLongTextBackup].[PrimaryKey]=[" & AUDIT_KEY_LOCAL_VAR & "] And "
            strXml = strXml & "[tblLongTextBackup].[FieldName]=""" & strFieldName & """"
            strXml = strXml & "</WhereCondition>"
            strXml = strXml & "</Data>"
            strXml = strXml & "<Statements>"
        End If

        strXml = strXml & "<CreateRecord>"
        If booIsLongText Then
            strXml = strXml & "<Data><Reference>tblAuditLog</Reference></Data>"
        Else
            strXml = strXml & "<Data Alias=""NewAudit""><Reference>tblAuditLog</Reference></Data>"
        End If
        strXml = strXml & "<Statements>"

        strXml = strXml & "<Action Name=""SetField"">"
        If booIsLongText Then
            strXml = strXml & "<Argument Name=""Field"">tblAuditLog.TableName</Argument>"
        Else
            strXml = strXml & "<Argument Name=""Field"">NewAudit.TableName</Argument>"
        End If
        strXml = strXml & "<Argument Name=""Value"">""" & strTableName & """</Argument>"
        strXml = strXml & "</Action>"

        If strPrimaryKeyField <> "" Then
            strXml = strXml & "<Action Name=""SetField"">"
            If booIsLongText Then
                strXml = strXml & "<Argument Name=""Field"">tblAuditLog.PrimaryKey</Argument>"
            Else
                strXml = strXml & "<Argument Name=""Field"">NewAudit.PrimaryKey</Argument>"
            End If
            strXml = strXml & "<Argument Name=""Value"">[" & AUDIT_KEY_LOCAL_VAR & "]</Argument>"
            strXml = strXml & "</Action>"
        End If

        strXml = strXml & "<Action Name=""SetField"">"
        If booIsLongText Then
            strXml = strXml & "<Argument Name=""Field"">tblAuditLog.FieldName</Argument>"
        Else
            strXml = strXml & "<Argument Name=""Field"">NewAudit.FieldName</Argument>"
        End If
        strXml = strXml & "<Argument Name=""Value"">""" & strFieldName & """</Argument>"
        strXml = strXml & "</Action>"

        ' OperationType (schema Business Rule 6); NewValue stays Null on a delete
        strXml = strXml & "<Action Name=""SetField"">"
        If booIsLongText Then
            strXml = strXml & "<Argument Name=""Field"">tblAuditLog.OperationType</Argument>"
        Else
            strXml = strXml & "<Argument Name=""Field"">NewAudit.OperationType</Argument>"
        End If
        strXml = strXml & "<Argument Name=""Value"">""Delete""</Argument>"
        strXml = strXml & "</Action>"

        ' [BUSINESS LOGIC — schema Business Rule 4] Through AuditValueExpression: OldValue is
        '                 a text column, this loop includes the primary key, and a delete is
        '                 the other place a Replication ID key reaches a value column.
        strXml = strXml & "<Action Name=""SetField"">"
        If booIsLongText Then
            strXml = strXml & "<Argument Name=""Field"">tblAuditLog.OldValue</Argument>"
            strXml = strXml & "<Argument Name=""Value"">[BackupRec].[OldValue]</Argument>"
        Else
            strXml = strXml & "<Argument Name=""Field"">NewAudit.OldValue</Argument>"
            strXml = strXml & "<Argument Name=""Value"">" & _
                AuditValueExpression("[Old].[" & strFieldName & "]", lngFldType) & "</Argument>"
        End If
        strXml = strXml & "</Action>"

        strXml = strXml & "<Action Name=""SetField"">"
        If booIsLongText Then
            strXml = strXml & "<Argument Name=""Field"">tblAuditLog.DateChanged</Argument>"
        Else
            strXml = strXml & "<Argument Name=""Field"">NewAudit.DateChanged</Argument>"
        End If
        strXml = strXml & "<Argument Name=""Value"">Now()</Argument>"
        strXml = strXml & "</Action>"

        ' [STANDARDS / schema Business Rule 9] identity — see BuildAfterInsertMacro
        strXml = strXml & "<Action Name=""SetField"">"
        If booIsLongText Then
            strXml = strXml & "<Argument Name=""Field"">tblAuditLog.ChangedBy</Argument>"
        Else
            strXml = strXml & "<Argument Name=""Field"">NewAudit.ChangedBy</Argument>"
        End If
        ' [STANDARDS / schema Business Rule 9] AuditUser() — same person as the stamp.
        strXml = strXml & "<Argument Name=""Value"">AuditUser()</Argument>"
        strXml = strXml & "</Action>"

        strXml = strXml & "</Statements></CreateRecord>"

        If booIsLongText Then
            strXml = strXml & "</Statements></LookUpRecord>"
        End If
      End If
    Next fieldInfo

    strXml = strXml & "</Statements></DataMacro>"
    BuildAfterDeleteMacro = strXml
End Function
```
