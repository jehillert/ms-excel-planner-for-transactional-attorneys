Attribute VB_Name = "ShapeTools"
Sub CenterIt()
   ActiveSheet.Shapes("SwapUpTriangle").Select 'get the object
   With Selection
   .Left = Range("K1").Left + (Range("K1:L1").Width - Selection.Width) / 4
   .Top = Range("K1").Top + (Range("K1").Height - Selection.Height) / 2
   End With
   End Sub
