Attribute VB_Name = "a_TEMPLATES"
'DO NOT DELETE.  MAKE ALL SUBS PRIVATE.  THAT SHOULD PREVENT INTERACTIONS WITH REGULAR USE
Private Sub CommandButton1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
   On Error Resume Next
   Application.ScreenUpdating = False
   'LBUTTON
   If Button = 1 Then
      MsgBox "You left clicked."
   'RBUTTON
   ElseIf Button = 2 Then
      MsgBox "You right clicked."
   End If
   Application.ScreenUpdating = True
   End Sub

Sub PerformOperationsOnEachCellInSelection()
   Dim cell As Range
   Dim i As Integer
   For Each cell In Selection
      'cell.Font.Name = MY_FONT
   Next
   End Sub
