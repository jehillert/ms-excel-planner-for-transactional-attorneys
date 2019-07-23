VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Const TopRow = 1
Private nFree As Integer


'*************************************************************
' WORKSHEET EVENT HANDLERS
'*************************************************************
Private Sub Worksheet_Activate()
   Application.ScreenUpdating = False
   Application.FormulaBarHeight = 1
   Application.Iteration = True
   ZoomToSel "A1,G1"
   FreezeRowPane
   RibbonExpand
   HeadingsYes
   CenterInUpperLeftCorner
   'ActiveSheet.Hyperlinks.Add shp, "", SubAddress:="Sheet2!A5", ScreenTip:=tipstring
   Application.ScreenUpdating = True
   End Sub

'*************************************************************
' COLUMN BUTTONS
'*************************************************************
Private Sub PLUS_BUTTON_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
   On Error Resume Next
   Application.ScreenUpdating = False
   'LBUTTON
   If Button = 1 Then
      AddBlankTimeEntry (1)
   'RBUTTON
   ElseIf Button = 2 Then
      AddBlankTimeEntry (9)
   End If
   Cells(Application.ActiveCell.row, 2).Select
   Application.ScreenUpdating = True
   End Sub
Private Sub MATTER_COLUMN_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
   On Error Resume Next
   Application.ScreenUpdating = False
   'LBUTTON
   If Button = 1 Then
      'DuplicateMatter_ThisTimeEntry
      DuplicateMatter_LastIfAboveThisIfBelow
   'RBUTTON
   ElseIf Button = 2 Then
      MsgBox "You right clicked."
   End If
   Application.ScreenUpdating = True
   End Sub
Private Sub DATE_COLUMN_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
   On Error Resume Next
   Application.ScreenUpdating = False
   'LBUTTON
   If Button = 1 Then
      InsertTimeEntryDate
   'RBUTTON
   ElseIf Button = 2 Then
      MsgBox "You right clicked."
   End If
   Application.ScreenUpdating = True
   End Sub
Private Sub START_COLUMN_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
   On Error Resume Next
   Application.ScreenUpdating = False
   'LBUTTON - RecordStartTime
   If Button = 1 Then
      Cells(RowToChange(3), 3).Value = Now()
   'RBUTTON
   ElseIf Button = 2 Then
      MsgBox "You right clicked."
   End If
   Application.ScreenUpdating = True
   End Sub
Private Sub END_COLUMN_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
   On Error Resume Next
   Application.ScreenUpdating = False
   'LBUTTON
   If Button = 1 Then
      RecordTime ("D")
   'RBUTTON
   ElseIf Button = 2 Then
      MsgBox "You right clicked."
   End If
   Application.ScreenUpdating = True
   End Sub
Private Sub TIME_COLUMN_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
   On Error Resume Next
   Application.ScreenUpdating = False
   'LBUTTON
   If Button = 1 Then
      HideStartStopTimes
   'RBUTTON
   ElseIf Button = 2 Then
      MsgBox "BLANK FOR NOW"
   End If
   Application.ScreenUpdating = True
   End Sub
'*************************************************************
' SHAPE BUTTONS
'*************************************************************
Private Sub ARROW_BUTTON_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
   On Error Resume Next
   Application.ScreenUpdating = False
   'LBUTTON
   If Button = 1 Then
      ArchiveEntries
   'RBUTTON -
   ElseIf Button = 2 Then
      MsgBox "You right clicked."
   End If
   Application.ScreenUpdating = True
   End Sub
Private Sub MINUS_BUTTON_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
   On Error Resume Next
   Application.ScreenUpdating = False
   'LBUTTON
   If Button = 1 Then
      DeleteThisRow
   'RBUTTON
   ElseIf Button = 2 Then
      MsgBox "You right clicked."
   End If
   Application.ScreenUpdating = True
   End Sub
Private Sub FULL_SCREEN_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
   On Error Resume Next
   Application.ScreenUpdating = False
   'LBUTTON
   If Button = 1 Then
      HideStartStopTimes
   'RBUTTON -
   ElseIf Button = 2 Then
      MsgBox "You right clicked."
   End If
   Application.ScreenUpdating = True
   End Sub
Private Sub WIDE_BAR_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
   On Error Resume Next
   Application.ScreenUpdating = False
   'LBUTTON
   If Button = 1 Then
      ShowHideComplete
   'RBUTTON -
   ElseIf Button = 2 Then
      ShowHideNarratives
   End If
   Application.ScreenUpdating = True
   End Sub
'*************************************************************
' SUBROUTINES/FUNCTIONS
'*************************************************************
Private Sub AddBlankTimeEntry(Optional NumberOfNewEntries As Integer = 1)
   On Error Resume Next
   Application.ScreenUpdating = False
   Dim NewRow, RowAbove, RowBelow As Integer
   RowAbove = Selection.row
   RowBelow = Selection.row + 1
      NewRow = RowBelow
   For i = 1 To NumberOfNewEntries
      Cells(NewRow, 1).EntireRow.Insert
      Cells(NewRow, 1).Formula = "=IF(B" & NewRow & "<>"""",IF(A" & NewRow & "="""",TODAY(),A" & NewRow & "),"""")"
      'Cells(NewRow, 3).Formula = "=IF(B" & NewRow & "<>"""",IF(C" & NewRow & "="""",NOW(),C" & NewRow & "),"""")"
      Cells(NewRow, 5).Formula = "=(((D" & NewRow & "-C" & NewRow & ")-INT(D" & NewRow & "-C" & NewRow & "))*24)"
   Next i
   Application.ScreenUpdating = True
   End Sub
Private Sub ArchiveEntries()
   Application.ScreenUpdating = False
   On Error Resume Next
   Dim rng As Range
   Dim rw, rwCount As Integer
   Set rng = Selection
   rw = rng.Cells(1, 1).row 'row number of first row in selection
   'DO NOT DELETE UI CONTROLS
   If rw > 2 Then
      rwCount = rng.Rows.Count 'rows to delete
      ' copy, insert elsewhere, delete and slide up
      ' next row until all rows have been transferred
      For i = 1 To rwCount
         Sheets("Time").Rows(rw).EntireRow.Copy
         If Sheets("Time").Cells(rw, 5).Value = 0 Then
            'MsgBox "send to ARCHIVE"
            Sheets("Archive").Rows("3:3").Insert Shift:=xlDown
         Else
            'MsgBox "send to COMPLETE"
            Sheets("Complete").Rows("3:3").Insert Shift:=xlDown
         End If
         Sheets("Time").Rows(rw).EntireRow.Delete
      Next i
      Sheets("Time").Cells(rw, 1).Select
   End If
   Application.ScreenUpdating = False
   End Sub
Private Sub DuplicateMatter_LastIfAboveThisIfBelow()
   Application.ScreenUpdating = False
   Dim SourceRow, TargetRow As Integer
   TargetRow = NextFreeRow("B")
   If TargetRow <= ActiveCell.row Then
      SourceRow = TargetRow - 1
   Else
      SourceRow = ActiveCell.row
   End If
   Range("B" & TargetRow).Value = Range("B" & SourceRow).Value
   Range("C" & TargetRow).Value = Now()
   Application.ScreenUpdating = True
   End Sub
Private Sub DuplicateMatter_LastTimeEntry()
   Application.ScreenUpdating = False
   Dim SourceRow, TargetRow As Integer
   TargetRow = NextFreeRow("B")
   SourceRow = TargetRow - 1
   Range("B" & TargetRow).Value = Range("B" & SourceRow).Value
   Range("C" & TargetRow).Value = Now()
   Application.ScreenUpdating = True
   End Sub
Private Sub DuplicateMatter_ThisTimeEntry()
   Application.ScreenUpdating = False
   Dim SourceRow, TargetRow As Integer
   TargetRow = NextFreeRow("B")
   SourceRow = ActiveCell.row
   Range("B" & TargetRow).Value = Range("B" & SourceRow).Value
   Range("C" & TargetRow).Value = Now()
   Application.ScreenUpdating = True
   End Sub
Private Sub HideStartStopTimes()
   On Error Resume Next
   Application.ScreenUpdating = False
   ActiveSheet.Columns("C:D").Hidden = Not ActiveSheet.Columns("C:D").Hidden
   SetResetZoom
   Application.ScreenUpdating = True
   End Sub
Private Sub InsertTimeEntryDate()
   lastRowIndex = ActiveSheet.Cells(ActiveSheet.Rows.Count, "B").End(xlUp).row
   Cells(lastRowIndex + 1, 1).Value = Date
   End Sub
Private Function RowToChange(Optional c As Long = 0) As Long
   Dim r1, r2, c1, c2 As Long
   Dim c1_ABC As String
   If c = 0 Then
      c1 = ActiveCell.Column
   Else
      c1 = c
   End If
   c2 = c1 + 1
   c1_ABC = ColumnAsLetter(c)
   r2 = NextFreeRow(c1_ABC)
   r1 = r2 - 1
'   MsgBox c1
'   MsgBox Cells(r1, c2).Range.Value
'    MsgBox Cells(r1, c2).Value
   If Cells(r1, c2).Value = "" Then
      RowToChange = r1
   Else
      RowToChange = r2
   End If
   'MsgBox RowToChange
   End Function
Private Function RecordTime(myCol As String)
   On Error Resume Next
   nFree = NextFreeRow(myCol)
   Range(myCol & nFree).Value = Now()
   End Function
Private Sub SendTimeEntry_Click()
   MailSelectedOpen
   End Sub
Private Sub ShowHideArchive()
   Application.ScreenUpdating = False
   Dim NewVisibilityState As Boolean
   NewVisibilityState = Not Sheets("Archive").Visible
   If NewVisibilityState = True Then
      HideNormallyHiddenSheets
   Else
      Sheets("TF").Visible = True
   End If
   Sheets("Archive").Visible = NewVisibilityState
   Sheets("Complete").Visible = NewVisibilityState
   Application.ScreenUpdating = True
   End Sub
Private Sub ShowHideComplete()
   Application.ScreenUpdating = False
   Dim NewVisibilityState As Boolean
   NewVisibilityState = Not Sheets("Complete").Visible
   If NewVisibilityState = True Then
      HideNormallyHiddenSheets
   Else
      Sheets("TF").Visible = True
   End If
   Sheets("Archive").Visible = NewVisibilityState
   Sheets("Complete").Visible = NewVisibilityState
   Sheets("Complete").Activate
   Application.ScreenUpdating = True
   End Sub
Private Sub ShowHideNarratives()
   Application.ScreenUpdating = False
   Dim NewVisibilityState As Boolean
   NewVisibilityState = Not Sheets("Narratives").Visible
   If NewVisibilityState = True Then
      HideNormallyHiddenSheets
   Else
      Sheets("TF").Visible = True
   End If
   Sheets("Narratives").Visible = NewVisibilityState
   Application.ScreenUpdating = True
   End Sub
Private Sub TimeStamp_Click()
   On Error Resume Next
   ActiveCell.Value = Now()
   ActiveCell.NumberFormat = "hh:mm"
   End Sub
