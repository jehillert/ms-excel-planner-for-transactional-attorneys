VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Activate()
   Application.ScreenUpdating = False
   Application.FormulaBarHeight = 1
   ZoomToSel "A1,F1", , 1
   RibbonCollapse
   FreezeRowPane
   HeadingsYes
   CenterInUpperLeftCorner
   Application.ScreenUpdating = True
   End Sub
Private Sub HideThisSheet_Click()
   MoveAndHide "Time"
   End Sub
Private Sub ToggleFormulaBarHeight_Click()
   If Application.FormulaBarHeight = 1 Then
      Application.FormulaBarHeight = 3
   ElseIf Application.FormulaBarHeight = 3 Then
      Application.FormulaBarHeight = 5
   Else
      Application.FormulaBarHeight = 1
   End If
   End Sub

