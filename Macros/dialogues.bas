Attribute VB_Name = "dialogues"
Sub displayVisibleRange()
   Dim xRg As Range
   Dim xTxt As String
   Dim xCell As Range
   Dim xStr As String
   Dim xRow As Long
   Dim xCol As Long
   On Error Resume Next
   xTxt = Application.ActiveWindow.VisibleRange.AddressLocal
   Set xRg = Application.InputBox("Please select range:", "Current Visible Range", xTxt, , , , , 8)
   If xRg Is Nothing Then Exit Sub
   On Error Resume Next
   For xRow = 1 To xRg.Rows.Count
    For xCol = 1 To xRg.Columns.Count
       xStr = xStr & xRg.Cells(xRow, xCol).Value & vbTab
    Next
    xStr = xStr & vbCrLf
   Next
   MsgBox xStr, vbInformation, "Kutools for Excel"
End Sub
