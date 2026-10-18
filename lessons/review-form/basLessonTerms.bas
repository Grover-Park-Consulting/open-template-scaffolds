Option Compare Database
Option Explicit

Private Const MODULE_NAME As String = "basLessonTerms"

' Opens the help popup for one term. Called from the small ? button beside each field.
Public Function ShowTerm(ByVal strTerm As String) As Boolean
      Const PROC_NAME As String = "ShowTerm"
100       On Error GoTo errHandler

110       EnsureTerms
120       DoCmd.OpenForm "frmLessonTerm_Help", acNormal, , "TermName = '" & Replace(strTerm, "'", "''") & "'", acFormReadOnly, acDialog
130       ShowTerm = True

Cleanup:
      Exit Function

errHandler:
240       MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation
250       Resume Cleanup
260       Resume
End Function

' Creates the terms table when it is missing and brings its rows up to date with the text below.
Public Sub EnsureTerms()
      Const PROC_NAME As String = "EnsureTerms"
100       On Error GoTo errHandler

      Dim db As DAO.Database

110       Set db = CurrentDb
120       If Not TableExists(db, "tlkpLessonTerm") Then
130           db.Execute "CREATE TABLE tlkpLessonTerm (LessonTermID COUNTER CONSTRAINT pkLessonTerm PRIMARY KEY, " & _
                  "TermName TEXT(60) NOT NULL, TermOrder LONG NOT NULL, TermDescription LONGTEXT NOT NULL, " & _
                  "CreatedDate DATETIME NOT NULL, CreatedBy TEXT(100) NOT NULL, ModifiedDate DATETIME, ModifiedBy TEXT(100))", dbFailOnError
140           db.Execute "CREATE UNIQUE INDEX uxLessonTermName ON tlkpLessonTerm (TermName)", dbFailOnError
          End If
150       SeedTerms db

Cleanup:
      On Error Resume Next
      Set db = Nothing
      Exit Sub

errHandler:
240       MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation
250       Resume Cleanup
260       Resume
End Sub

Private Sub SeedTerms(ByVal db As DAO.Database)
      Const PROC_NAME As String = "SeedTerms"
100       On Error GoTo errHandler

      Dim s As String

110       AddTerm db, "Lesson", 10, _
              "The short name of one lesson, for example no-count-distinct. Every lesson has a different name." & vbCrLf & vbCrLf & _
              "A lesson that came with the library keeps its name. A lesson you add yourself gets a name you choose: lowercase words joined by hyphens."

120       AddTerm db, "Mechanism", 20, _
              "The part of Access or of the build process that the lesson is about, such as error-handling or schema-migration." & vbCrLf & vbCrLf & _
              "Each template names the mechanisms it uses. Before a design starts, the assistant is given the approved lessons for those mechanisms. " & _
              "Choose from the list. If none fits, ask for a new mechanism to be added to the library's list."

130       AddTerm db, "Kind", 30, _
              "Guardrail: tells the assistant what to do or what to avoid in a build. It applies without interrupting you." & vbCrLf & vbCrLf & _
              "Consequence: warns that a choice you make will conflict with another choice, or will block the build later. " & _
              "The assistant tells you before you commit to it, and you decide."

140       s = "How well the lesson is established." & vbCrLf & vbCrLf
150       s = s & "Observed: seen happen in a build, with evidence. The assistant applies it." & vbCrLf
160       s = s & "Inferred: concluded from what was observed, not seen directly. The assistant tells you about it as something to consider." & vbCrLf
170       s = s & "Assumed: relied on in a build and never tested. The assistant asks you." & vbCrLf
180       s = s & "Potential: noticed but not shown. Kept for future trials and not used in a design." & vbCrLf & vbCrLf
190       s = s & "Two independent builds that see the same thing move an Inferred lesson to Observed."
200       AddTerm db, "Confidence", 40, s

210       AddTerm db, "Origin", 50, _
              "Library: the lesson came with the library. Its text is restored from the library each time the lessons are synchronized, " & _
              "so only its approval can be changed here." & vbCrLf & vbCrLf & _
              "Own: you added it. You can edit every field."

220       s = "Where the lesson stands with you." & vbCrLf & vbCrLf
230       s = s & "Proposed: waiting for your decision. The assistant is not given it." & vbCrLf
240       s = s & "Approved: the assistant is given it before a design starts." & vbCrLf
250       s = s & "Rejected: kept, but never given to the assistant." & vbCrLf & vbCrLf
260       s = s & "Change it with the Approve, Reject and Proposed buttons at the bottom of the form."
270       AddTerm db, "Approval", 60, s

280       AddTerm db, "Lesson text", 70, _
              "What to do or what to expect, in one or two sentences. This is the text the assistant reads." & vbCrLf & vbCrLf & _
              "Leave out names of subjects, templates, machines and people, so the lesson applies to any database."

290       AddTerm db, "Applies when", 80, _
              "The conditions under which the lesson held. Every lesson needs this, because a lesson with no stated conditions cannot be checked."

300       AddTerm db, "Triggers when", 90, _
              "Only for a consequence. The choice that sets the lesson off, and the earlier choice or condition it conflicts with, " & _
              "in the words the template uses for them." & vbCrLf & vbCrLf & _
              "Leave it empty for a guardrail."

310       AddTerm db, "Evidence", 100, _
              "What was seen that supports the lesson: the error number and the object that raised it, the output, and the fix that worked." & vbCrLf & vbCrLf & _
              "A lesson with no evidence for its confidence level should not be entered."

320       AddTerm db, "Environment", 110, _
              "Engine and version facts the lesson depends on, each with the date it was last checked." & vbCrLf & vbCrLf & _
              "Empty means the lesson holds anywhere. The assistant checks these facts before it applies the lesson."

330       AddTerm db, "Sources", 120, _
              "The builds the lesson came from, one per line, as the template name and the date (for example stocktake-schema 09-22). " & _
              "If it did not come from a build, name where it came from."

340       s = "Use this list to record your own experience of this lesson in a build." & vbCrLf & vbCrLf
350       s = s & "Kind: Confirmation means a later build saw the same thing independently. Contradiction means a later build saw something different." & vbCrLf
360       s = s & "Build: the template name and the date of the build, for example stocktake-schema 2026-10-18, so that someone can find it." & vbCrLf
370       s = s & "Origin: filled in for you. Rows you add are Own. Rows that came with the library are Library and cannot be changed." & vbCrLf & vbCrLf
380       s = s & "To add a row: click the empty row at the bottom of the list, choose Confirmation or Contradiction, type the build, and click away. It saves by itself." & vbCrLf & vbCrLf
390       s = s & "Count a build only if it was independent. A build that reused an earlier build's method or code is not a confirmation." & vbCrLf & vbCrLf
400       s = s & "Two independent confirmations move an Inferred lesson to Observed. A contradiction sends the lesson back for review."
410       AddTerm db, "My builds", 130, s

420       s = "Approve, Reject and Proposed set the approval of the lesson on the screen and save it at once." & vbCrLf & vbCrLf
430       s = s & "New starts a lesson of your own. It begins as Proposed." & vbCrLf & vbCrLf
440       s = s & "The box in the top bar shows only the lessons with the approval you pick. All shows every lesson." & vbCrLf & vbCrLf
450       s = s & "The arrows at the bottom move from one lesson to the next."
460       AddTerm db, "Buttons and filter", 140, s

Cleanup:
      Exit Sub

errHandler:
540       MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation
550       Resume Cleanup
560       Resume
End Sub

' Adds the term, or refreshes its text when the stored text differs.
Private Sub AddTerm(ByVal db As DAO.Database, ByVal strName As String, ByVal lngOrder As Long, ByVal strText As String)
      Const PROC_NAME As String = "AddTerm"
100       On Error GoTo errHandler

      Dim rs As DAO.Recordset

110       Set rs = db.OpenRecordset("SELECT * FROM tlkpLessonTerm WHERE TermName = '" & Replace(strName, "'", "''") & "'", dbOpenDynaset)
120       If rs.EOF Then
130           rs.AddNew
140           rs!TermName = strName
150           rs!TermOrder = lngOrder
160           rs!TermDescription = strText
170           rs!CreatedDate = Now()
180           rs!CreatedBy = Environ$("USERNAME")
190           rs.Update
200       ElseIf rs!TermDescription <> strText Or rs!TermOrder <> lngOrder Then
210           rs.Edit
220           rs!TermOrder = lngOrder
230           rs!TermDescription = strText
240           rs!ModifiedDate = Now()
250           rs!ModifiedBy = Environ$("USERNAME")
260           rs.Update
          End If

Cleanup:
      On Error Resume Next
      rs.Close
      Set rs = Nothing
      Exit Sub

errHandler:
340       MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation
350       Resume Cleanup
360       Resume
End Sub