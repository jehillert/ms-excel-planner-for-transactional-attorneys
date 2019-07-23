Attribute VB_Name = "WorksheetZoom"
Function ZoomToSel(ZoomToRange As String, Optional SelectedAtEnd As String = "A1", Optional ZoomOffset As Integer = 2)
   On Error Resume Next
   Range(ZoomToRange).Select
   ActiveWindow.Zoom = True
   Range(SelectedAtEnd).Select
   ActiveWindow.Zoom = ActiveWindow.Zoom + ZoomOffset
   End Function
Sub SetResetZoom()
   If ActiveWindow.ActiveSheet.Name = "Archive" Or _
      ActiveWindow.ActiveSheet.Name = "Complete" Or _
      ActiveWindow.ActiveSheet.Name = "Time" Then
      ZoomToSel "A1:G1"
   ElseIf ActiveWindow.ActiveSheet.Name = "Calendar" Then
      ZoomToSel "A1:M1"
   ElseIf ActiveWindow.ActiveSheet.Name = "Payroll" Then
      ZoomToSel "A1:H1"
   ElseIf ActiveWindow.ActiveSheet.Name = "VARS" Then
      ZoomToSel "A1:H1", 0
   ElseIf ActiveWindow.ActiveSheet.Name = "Narratives" Then
      ZoomToSel "A1:B1", 1
   End If
   End Sub
