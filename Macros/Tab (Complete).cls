VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Const TopRow = 1
Private nFree As Integer
Private Sub Worksheet_Activate()
   Application.ScreenUpdating = False
   Application.FormulaBarHeight = 1
   ZoomToSel "A1,G1", , 0
   FreezeRowPane
   RibbonExpand
   HeadingsYes
   CenterInUpperLeftCorner
   Application.ScreenUpdating = True
   End Sub
Private Sub AddBlankTimeEntries_Click()
   On Error Resume Next
   AddBlankTimeEntry (1): Cells(Application.ActiveCell.row, 2).Select
   End Sub
Private Sub AddBlankTimeEntries_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
   On Error Resume Next
   AddBlankTimeEntry (9): Cells(Application.ActiveCell.row, 2).Select
   End Sub
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
      Cells(NewRow, 5).Formula = "=(((D" & NewRow & "-C" & NewRow & ")-INT(D" & NewRow & "-C" & NewRow & "))*24)"
   Next i
   Application.ScreenUpdating = True
   End Sub
Private Sub DeleteRow_Click()
   DeleteThisRow
   End Sub
Private Sub HideStartStopTimes_Click()
   On Error Resume Next
   Application.ScreenUpdating = False
   ActiveSheet.Columns("C:D").Hidden = Not ActiveSheet.Columns("C:D").Hidden
   ZoomToSel "A1,G1", , 0
   Application.ScreenUpdating = True
   End Sub
Private Sub HideThisSheet_Click()
   MoveAndHide "Time"
   End Sub
