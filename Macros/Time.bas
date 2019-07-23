Attribute VB_Name = "Time"
Sub StampTime()
   On Error Resume Next
   ActiveCell.Value = Now()
   ActiveCell.NumberFormat = "hh:mm"
   End Sub
'https://www.myonlinetraininghub.com/pausing-or-delaying-vba-using-wait-sleep-or-a-loop
Sub WasteTime(Finish As Long)
   Dim NowTick As Long
   Dim EndTick As Long
   EndTick = GetTickCount + (Finish * 1000)
   Do
    NowTick = GetTickCount
    DoEvents
   Loop Until NowTick >= EndTick
   End Sub

