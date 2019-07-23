VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
   DisplayFilenameOnlyInWinCaption
   Application.OnKey "^+c", "CopyCellContents"
   Application.OnKey "^+x", "CopyCellContents"
   Application.OnKey "^+v", "MoveCellContents"
   Application.OnKey "^+z", "ClearActiveCellContentAndComments"
   Application.OnKey "%+{Up}", "SwapCalendarEntryUp"
   Application.OnKey "%+{Down}", "SwapCalendarEntryDown"
   Application.ScreenUpdating = False
   Application.WindowState = xlMaximized
   ActiveWindow.WindowState = xlMaximized
   'Sheets("Time").Activate
   Application.Iteration = True
   Application.ScreenUpdating = True
End Sub
Private Sub Workbook_BeforeClose(Cancel As Boolean)
   If Me.Saved = False Then Me.Save
   End Sub
Private Sub Workbook_SheetDeactivate(ByVal Sh As Object)
   Set LastActiveSheet = Sh
   End Sub
'THIS GOES IN A MODULE
'Private Sub AutoOpen()
'   Application.ScreenUpdating = False
'   Application.WindowState = xlMaximized
'   ActiveWindow.WindowState = xlMaximized
'   Application.ScreenUpdating = True
'   End Sub
Private Sub DisplayFilenameOnlyInWinCaption()
   Dim Filename As String
   Filename = Left(ActiveWorkbook.Name, InStr(ActiveWorkbook.Name, ".") - 1)
   ActiveWindow.Caption = Filename
   ActiveWindow.Application.Caption = " "
   End Sub

