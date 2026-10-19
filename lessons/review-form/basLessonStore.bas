Option Compare Database
Option Explicit

Private Const MODULE_NAME As String = "basLessonStore"
Private Const STORE_FILE As String = "lessons.accdb"

' Links the nine lesson tables from the lessons database that sits beside this file.
' A link that points anywhere else is replaced. Returns True when all the links are good.
Public Function EnsureStoreLinks() As Boolean
Const PROC_NAME As String = "EnsureStoreLinks"
    On Error GoTo errHandler

Dim db         As DAO.Database
Dim strStore   As String
Dim varTable   As Variant

    strStore = CurrentProject.Path & "\" & STORE_FILE
    If Len(Dir(strStore)) = 0 Then
        Err.Raise vbObjectError + 513, PROC_NAME, "The lessons database was not found beside this file: " & strStore
    End If
    Set db = CurrentDb
    For Each varTable In Array("tlkpMechanism", "tblLesson", "tblLessonEvent", "tlkpLessonKind", "tlkpLessonConfidence", _
                               "tlkpLessonStatus", "tlkpLessonOrigin", "tlkpLessonApproval", "tlkpLessonEventKind")
        If TableExists(db, CStr(varTable)) Then
            If InStr(1, db.TableDefs(CStr(varTable)).Connect, ";DATABASE=" & strStore, vbTextCompare) = 0 Then
                db.TableDefs.Delete CStr(varTable)
                DoCmd.TransferDatabase acLink, "Microsoft Access", strStore, acTable, CStr(varTable), CStr(varTable)
            End If
        Else
            DoCmd.TransferDatabase acLink, "Microsoft Access", strStore, acTable, CStr(varTable), CStr(varTable)
        End If
    Next varTable
    EnsureStoreLinks = True

Cleanup:
    On Error Resume Next
    Set db = Nothing
    Exit Function

errHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation
    Resume Cleanup
    Resume
End Function

Public Function TableExists(ByVal db As DAO.Database, ByVal strName As String) As Boolean
Dim td As DAO.TableDef
    For Each td In db.TableDefs
        If StrComp(td.Name, strName, vbTextCompare) = 0 Then
            TableExists = True
            Exit Function
        End If
    Next td
End Function

' Gives the database its name and makes the review form the one that opens. Safe to run again.
Public Sub ApplyAppSettings()
Const PROC_NAME As String = "ApplyAppSettings"
    On Error GoTo errHandler

    SetStartupProperty "AppTitle", "Lessons Learned from Builds"
    SetStartupProperty "StartUpForm", "frmLesson_Review"
    SetStartupProperty "StartupShowDBWindow", False

Cleanup:
    Exit Sub

errHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation
    Resume Cleanup
    Resume
End Sub

Private Sub SetStartupProperty(ByVal strName As String, ByVal varValue As Variant)
Const PROC_NAME As String = "SetStartupProperty"
    On Error GoTo errHandler

Dim db  As DAO.Database
Dim prp As DAO.Property

    Set db = CurrentDb
    On Error Resume Next
    db.Properties(strName) = varValue
    If Err.Number = 3270 Then
        Err.Clear
        Set prp = db.CreateProperty(strName, IIf(VarType(varValue) = vbBoolean, dbBoolean, dbText), varValue)
        db.Properties.Append prp
    End If
    On Error GoTo errHandler

Cleanup:
    On Error Resume Next
    Set db = Nothing
    Exit Sub

errHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation
    Resume Cleanup
    Resume
End Sub