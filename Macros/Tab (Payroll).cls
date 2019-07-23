VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'*************************************************************
' WORKSHEET EVENT HANDLERS
'*************************************************************
Private Sub Worksheet_Activate()
   Application.ScreenUpdating = False
   ActiveSheet.Rows("2:4").Hidden = True
   Application.FormulaBarHeight = 1
   CenterInUpperLeftCorner 5, 1
   ZoomToSel "A1,H1"
   FreezeRowPane 4
   RibbonCollapse
   HeadingsYes
   Application.ScreenUpdating = True
   End Sub
'Private Sub Worksheet_SelectionChange(ByVal Target As Range)
'   ActiveWindow.ScrollRow = Target.Row
'   ActiveWindow.ScrollColumn = Target.Column
'   End Sub
'*************************************************************
' CLICK EVENT HANDLERS
'*************************************************************
Private Sub ExpandContract_Click()
   On Error Resume Next
   Application.ScreenUpdating = False
   If ActiveWindow.Zoom = 110 Then
      HeadingsYes
      Application.DisplayFormulaBar = True
      ActiveWindow.Zoom = 120
   ElseIf ActiveWindow.Zoom = 120 Then
      HeadingsYes
      Application.DisplayFormulaBar = True
      ActiveWindow.Zoom = 110
   End If
   Application.ScreenUpdating = True
   End Sub
Private Sub ShowHide_Click()
   On Error Resume Next
   Application.ScreenUpdating = False
   FindShowHide_Below ("Archive")
   Range("A20").Select
   Application.ScreenUpdating = True
   End Sub
Private Sub RowDelete_Click()
   On Error Resume Next
   Application.ScreenUpdating = False
   Selection.EntireRow.Delete
   Cells(Application.ActiveCell.row, 6).Select
   Application.ScreenUpdating = True
   End Sub
Private Sub RowInsert_Click()
   On Error Resume Next
   Application.ScreenUpdating = False
   ActiveCell.EntireRow.Select
   Rows("4:4").Copy
   Selection.EntireRow.Insert
   Selection.EntireRow.Hidden = False
   Cells(Application.ActiveCell.row, 6).Select
   Application.ScreenUpdating = True
   End Sub
Private Sub WIDE_BAR_Click()
   Application.ScreenUpdating = False
   Dim NewVisibilityState As Boolean
   NewVisibilityState = Not Sheets("VARS").Visible
   If NewVisibilityState = True Then
      HideNormallyHiddenSheets
   Else
      Sheets("TF").Visible = True
   End If
   Sheets("VARS").Visible = NewVisibilityState
   Application.ScreenUpdating = True
   End Sub
'*************************************************************
' SUBROUTINES/FUNCTIONS
'*************************************************************
Private Sub RDPViewer1_OnAttendeeConnected(ByVal pAttendee As Object)
   End Sub
