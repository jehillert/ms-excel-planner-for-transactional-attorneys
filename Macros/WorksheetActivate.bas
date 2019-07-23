Attribute VB_Name = "WorksheetActivate"
' WORKSHEET ACTIVATION SUBS
Sub GoToLastActiveSheet()
   LastActiveSheet.Activate
   End Sub
Sub HideNormallyHiddenSheets(Optional VisibilityStateTF As Boolean = False)
   Sheets("Narratives").Visible = False
   Sheets("Complete").Visible = False
   Sheets("Archive").Visible = False
   Sheets("VARS").Visible = False
   Sheets("TF").Visible = VisibilityStateTF
   End Sub
Sub MoveAndHide(Optional DestinationSheet As String = "Unspecified")
   Application.ScreenUpdating = False
   If DestinationSheet = "Unspecified" Then
      GoToLastActiveSheet
   Else
      Sheets(DestinationSheet).Activate
   End If
   HideNormallyHiddenSheets (True)
   Application.ScreenUpdating = True
   End Sub
'FREEZE PANE FUNCTIONS
Sub ToggleFreezePanes()
   On Error Resume Next
   Dim r As Range
   Set r = ActiveCell
   If ActiveWindow.FreezePanes = False Then
      Range("A2").Select
      ActiveWindow.SplitColumn = 0
      ActiveWindow.SplitRow = 1
      ActiveWindow.FreezePanes = True
   Else
      ActiveWindow.SplitColumn = 0
      ActiveWindow.SplitRow = 0
      ActiveWindow.FreezePanes = False
   End If
      r.Select
   End Sub
Function UnfreezePanes()
   On Error Resume Next
   If ActiveWindow.FreezePanes = True Then
      ActiveWindow.SplitColumn = 0
      ActiveWindow.SplitRow = 0
      ActiveWindow.FreezePanes = False
   End If
   End Function
Function FreezeRowPane(Optional myRow As Integer = 1, Optional myCol As Integer = 0)
   If ActiveWindow.FreezePanes = False Then
      ActiveWindow.SplitColumn = myCol
      ActiveWindow.SplitRow = myRow
      ActiveWindow.FreezePanes = True
   End If
   End Function



