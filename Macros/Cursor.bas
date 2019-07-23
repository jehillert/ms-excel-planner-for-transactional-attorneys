Attribute VB_Name = "Cursor"
Public Declare PtrSafe Function GetCursorPos Lib "user32" _
(lpPoint As POINTAPI) As Long
Public Type POINTAPI
x As Long
y As Long
End Type

'Sub PositionXY()
'Dim lngCurPos As POINTAPI
'Do
'GetCursorPos lngCurPos
'
''Temporary code line to visually identify the X and Y coordinates.
''It will be replaced by the real line of code to run a macro
''when the coordinates are met.
'Range("G1").Value = "X: " & lngCurPos.x & " Y: " & lngCurPos.y
'
'DoEvents
'Loop
'End Sub

Sub PositionXY()
Dim lngCurPos As POINTAPI
Do
GetCursorPos lngCurPos
If (lngCurPos.x >= 121 And lngCurPos.x <= 411) And _
(lngCurPos.y >= 159 And lngCurPos.y <= 223) Then
UserForm1.Show
Exit Sub
End If
DoEvents
Loop
End Sub

`
