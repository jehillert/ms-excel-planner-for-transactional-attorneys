Attribute VB_Name = "WorksheetTools"
Function CenterInUpperLeftCorner(Optional row As Integer = 1, Optional col As Integer = 1)
   ActiveWindow.ScrollRow = row
   ActiveWindow.ScrollColumn = col
   End Function
Function ColumnAsLetter(ColumnNumber As Long) As String
   ColumnAsLetter = Split(Cells(1, ColumnNumber).Address, "$")(1)
   End Function
Function ColumnAsNumber(ColumnLetter As String) As Integer
   ColumnAsNumber = Range(ColumnLetter & 1).Column
   End Function
Function FindShowHide_Below(target As String)
   On Error Resume Next
   Dim ws As Excel.Worksheet
   Dim FoundCell As Excel.Range
   Dim HideStart, HideStop As Integer
   Set ws = ActiveSheet
   Set FoundCell = ws.Range("A1:K5000").Find(what:=target)
   ActiveSheet.Rows(FoundCell.row).Hidden = True
   HideStart = FoundCell.row + 1
   HideStop = LastRow + 1
   If ActiveSheet.Rows(HideStart & ":" & HideStop).Hidden = True Then
      ActiveSheet.Rows(HideStart & ":" & HideStop).Hidden = False
      ActiveSheet.Outline.ShowLevels RowLevels:=2
   Else
      ActiveSheet.Rows(HideStart & ":" & HideStop).Hidden = True
   End If
   Set FoundCell = ws.Range("A1:K5000").Find(what:="")
   End Function
Function LastRow() As Long
   LastRow = Cells.Find("*", [A1], , , xlByRows, xlPrevious).row
   End Function
Sub MarkAsComplete()
   ActiveCell.Replace what:="> ", Replacement:="", LookAt:=xlPart, _
      SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
      ReplaceFormat:=False
   ActiveCell.Replace what:="ü ", Replacement:="", LookAt:=xlPart, _
      SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
      ReplaceFormat:=False
   ActiveCell.Font.Name = MY_FONT
   ActiveCell.FormulaR1C1 = "ü " & ActiveCell.Value
   With ActiveCell.Characters(Start:=1, Length:=1).Font
      .Name = "Wingdings"
   End With
   Selection.Font.Bold = False
   ActiveCell.Offset(1, 0).Range("A1:B1").Select
   End Sub
Function NextFreeRow(Optional myCol As String = "ActiveColumn", Optional tcRow As String = 1) As Integer
   If myCol = "ActiveColumn" Then myCol = ColumnAsLetter(ActiveCell.Column)
   NextFreeRow = Range(myCol & tcRow & ":" & myCol & Rows.Count).Cells.SpecialCells(xlCellTypeBlanks).row
   End Function
Sub NumlockPlannerFunctions()
   Dim NewCol As Integer
   If ActiveWindow.ActiveSheet.Name = "Time" Then
      If ActiveCell.Column = 6 Then
         ActiveCell.Offset(1, -5).Range("A1").Select
      Else
         ActiveCell.Offset(0, 1).Range("A1").Select
      End If
   Else
      ActiveCell.Offset(0, 1).Range("A1").Select
   End If
End Sub
Sub SelectOutOfView()
   Dim r As Range
   Application.ScreenUpdating = False
   Set r = Application.ActiveWindow.VisibleRange
   r(r.Rows.Count + 1, r.Columns.Count + 1).Select
   Application.ScreenUpdating = True
   End Sub
Sub SendToArchive()
   On Error Resume Next
   Application.ScreenUpdating = False
   Rows(ActiveCell.row).Cut
   Sheets("Archive").Rows("3:3").Insert Shift:=xlDown
   Rows(ActiveCell.row).Delete
   Cells(ActiveCell.row, 1).Select
   Application.ScreenUpdating = True
   End Sub
Function ToggleWorksheet(HomeSheet As String, HideSheet As String)
   On Error Resume Next
   If Sheets(HideSheet).Visible = xlSheetVisible Then
      Sheets(HideSheet).Visible = xlHidden
      Sheets(HomeSheet).Activate
   Else
      Sheets(HideSheet).Visible = True
      Sheets(HideSheet).Activate
   End If
   End Function
Sub ToggleNSheet()
   Call ToggleWorksheet("Time", "Narratives")
   End Sub
Sub ToggleArchiveSheet()
   On Error Resume Next
   Application.ScreenUpdating = False
   Call ToggleWorksheet("Time", "Archive")
   Application.ScreenUpdating = True
   End Sub
Function ToggleTimeSheets()
   On Error Resume Next
   Application.ScreenUpdating = False
   If Sheets("Complete").Visible = True Or Sheets("Archive").Visible = True Then
      If ActiveWindow.ActiveSheet.Name = "Archive" Or ActiveWindow.ActiveSheet.Name = "Complete" _
         Then Sheets("Time").Activate
      If Sheets("Complete").Visible = True _
         Then Sheets("Complete").Visible = False
      If Sheets("Archive").Visible = True _
         Then Sheets("Archive").Visible = False
   Else
      If Sheets("Complete").Visible = False _
         Then Sheets("Complete").Visible = True
      If Sheets("Archive").Visible = False _
         Then Sheets("Archive").Visible = True
   End If
   Application.ScreenUpdating = True
   End Function
