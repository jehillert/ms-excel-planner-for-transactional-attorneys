VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Activate()
   Application.ScreenUpdating = False
   Application.FormulaBarHeight = 1
   CenterInUpperLeftCorner
   ZoomToSel "A1,G1"
   RibbonExpand
   HeadingsNo
   Application.ScreenUpdating = True
   End Sub
Private Sub HideVariablesSheet()
   MoveAndHide "VARS"
   If Sheets("Narratives").Visible = True Then
      Sheets("Time").Activate
      Sheets("Narratives").Visible = False
   Else
      Sheets("Payroll").Activate
   End If
   Sheets("VARS").Visible = False
   End Sub
Private Sub CATEGORY_BUTTON_Click()
   MoveAndHide "Payroll"
   End Sub
Private Sub TYPE_BUTTON_Click()
   MoveAndHide "Payroll"
   End Sub

