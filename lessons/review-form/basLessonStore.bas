Option Compare Database
Option Explicit

Private Const MODULE_NAME As String = "basLessonStore"
Private Const STORE_FILE As String = "lessons.accdb"

' Links the three lesson tables from the lessons database that sits beside this file.
' A link that points anywhere else is replaced. Returns True when all three links are good.
Public Function EnsureStoreLinks() As Boolean
      Const PROC_NAME As String = "EnsureStoreLinks"
100       On Error GoTo errHandler

      Dim db         As DAO.Database
      Dim strStore   As String
      Dim varTable   As Variant

110       strStore = CurrentProject.Path & "\" & STORE_FILE
120       If Len(Dir(strStore)) = 0 Then
130           Err.Raise vbObjectError + 513, PROC_NAME, "The lessons database was not found beside this file: " & strStore
          End If
140       Set db = CurrentDb
150       For Each varTable In Array("tlkpMechanism", "tblLesson", "tblLessonEvent")
160           If TableExists(db, CStr(varTable)) Then
170               If InStr(1, db.TableDefs(CStr(varTable)).Connect, ";DATABASE=" & strStore, vbTextCompare) = 0 Then
180                   db.TableDefs.Delete CStr(varTable)
190                   DoCmd.TransferDatabase acLink, "Microsoft Access", strStore, acTable, CStr(varTable), CStr(varTable)
                  End If
              Else
200               DoCmd.TransferDatabase acLink, "Microsoft Access", strStore, acTable, CStr(varTable), CStr(varTable)
              End If
210       Next varTable
220       EnsureStoreLinks = True

Cleanup:
      On Error Resume Next
      Set db = Nothing
      Exit Function

errHandler:
240       MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation
250       Resume Cleanup
260       Resume
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
100       On Error GoTo errHandler

110       SetStartupProperty "AppTitle", "Lessons Learned from Builds"
120       SetStartupProperty "StartUpForm", "frmLesson_Review"

Cleanup:
      Exit Sub

errHandler:
240       MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation
250       Resume Cleanup
260       Resume
End Sub

Private Sub SetStartupProperty(ByVal strName As String, ByVal strValue As String)
      Const PROC_NAME As String = "SetStartupProperty"
100       On Error GoTo errHandler

      Dim db  As DAO.Database
      Dim prp As DAO.Property

110       Set db = CurrentDb
120       On Error Resume Next
130       db.Properties(strName) = strValue
140       If Err.Number = 3270 Then
150           Err.Clear
160           Set prp = db.CreateProperty(strName, dbText, strValue)
170           db.Properties.Append prp
          End If
180       On Error GoTo errHandler

Cleanup:
      On Error Resume Next
      Set db = Nothing
      Exit Sub

errHandler:
240       MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation
250       Resume Cleanup
260       Resume
End Sub