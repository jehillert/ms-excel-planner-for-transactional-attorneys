VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Activate()
   Application.ScreenUpdating = False
   Application.FormulaBarHeight = 1
   ZoomToSel "A1,G1"
   FreezeRowPane 2
   RibbonExpand
   HeadingsYes
   CenterInUpperLeftCorner
   Application.ScreenUpdating = True
   End Sub
Private Sub Rowhide_Click()
   On Error Resume Next
   Application.ScreenUpdating = False
   Selection.EntireRow.Hidden = True
   Application.ScreenUpdating = True
   End Sub
Private Sub RowUnhide_Click()
   On Error Resume Next
   Application.ScreenUpdating = False
   Dim RowAbove, RowBelow As Integer
   RowAbove = Selection.Cells(1, 1).row - 1
   RowBelow = Selection.Cells(1, 1).row + 1
   Rows(RowAbove).EntireRow.Hidden = False
   Rows(RowBelow).EntireRow.Hidden = False
   'Selection.EntireRow.Hidden = False
   'MsgBox (Selection.Cells(1, 1).row)
   Application.ScreenUpdating = True
   End Sub
Private Sub HideThisSheet_Click()
   MoveAndHide "Time"
   End Sub
Private Sub ShowHideSkinnyButton_Click()
   Sheets("Time").Activate
   Sheets("Complete").Visible = False
   Sheets("Archive").Visible = False
   End Sub
Private Sub SearchDownButton_Click()
   Dim strFindWhat As String
   strFindWhat = UserSearch.Text
   On Error GoTo ErrorMessage
   Cells.Find(what:=strFindWhat, After:=ActiveCell, LookIn:=xlFormulas, LookAt _
   :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
   False).Select
   Exit Sub
ErrorMessage:
   MsgBox ("The data you are searching for does not exist")
   End Sub
Private Sub SearchUpButton_Click()
   Dim strFindWhat As String
   strFindWhat = UserSearch.Text
   On Error GoTo ErrorMessage
   Cells.Find(what:=strFindWhat, After:=ActiveCell, LookIn:=xlFormulas, LookAt _
   :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:= _
   False).Select
   Exit Sub
ErrorMessage:
   MsgBox ("The data you are searching for does not exist")
   End Sub
