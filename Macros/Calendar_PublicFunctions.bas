Attribute VB_Name = "Calendar_PublicFunctions"
Public CellText As String
Public CellComment As String
Public OriginCell As Range

Sub ClearActiveCellContentAndComments()
   On Error Resume Next
   Application.ScreenUpdating = False
   Dim cell
   Dim targetRange As Range
   Set targetRange = Selection
   targetRange.ClearComments
   targetRange.ClearContents
   Application.ScreenUpdating = True
   End Sub
Sub ClearOriginCell()
   If (OriginCell Is Nothing) Or (ActiveWindow.ActiveSheet.Name <> "Calendar") _
      Then Exit Sub
   OriginCell.ClearComments
   OriginCell.ClearContents
   Set OriginCell = Nothing
   End Sub
Sub CopyCellContents()
   If (ActiveWindow.ActiveSheet.Name <> "Calendar") Or (Selection.Count > 2) Then Exit Sub
   On Error Resume Next
   Set OriginCell = Selection
   End Sub
Sub MoveCellContents()
   If (OriginCell Is Nothing) Or (ActiveWindow.ActiveSheet.Name <> "Calendar") _
      Then Exit Sub
   On Error Resume Next
   Application.ScreenUpdating = False
   ActiveCell.ClearComments
   ActiveCell.Value = OriginCell.Cells(1, 1).Value
   ActiveCell.AddComment Text:=OriginCell.Cells(1, 1).Comment.Text
   'ActiveCell.Comment.Shape.TextFrame.AutoSize = True
   FormatCellComent ActiveCell.Comment
   FixSymbolFont
   ClearOriginCell
   Application.ScreenUpdating = True
   End Sub
Function RemoveFormatsDarkGray()
   On Error Resume Next
   Application.ScreenUpdating = False
   With Selection.Interior
      .Pattern = xlSolid
      .PatternColorIndex = xlAutomatic
      .ThemeColor = xlThemeColorDark1
      .TintAndShade = -0.149998474074526
      .PatternTintAndShade = 0
   End With
   Selection.FormatConditions.Delete
   Application.ScreenUpdating = True
   End Function
Sub SwapCalendarEntryUp()
   If ActiveWindow.ActiveSheet.Name <> "Calendar" Then Exit Sub
   On Error Resume Next
   Application.ScreenUpdating = False
   Dim cell_content
   Dim cell_comment As String
   cell_comment = ActiveCell.Comment.Text
   ActiveCell.ClearComments
   ActiveCell.AddComment Text:=ActiveCell.Offset(-1, 0).Comment.Text
   ActiveCell.Offset(-1, 0).ClearComments
   ActiveCell.Offset(-1, 0).AddComment Text:=cell_comment
   'ActiveCell.Comment.Shape.TextFrame.AutoSize = True
   FormatCellComent ActiveCell.Comment
   ActiveCell.Offset(0, -1).Comment.Shape.TextFrame.AutoSize = True
   cell_content = ActiveCell.Value
   ActiveCell.Value = ActiveCell.Offset(-1, 0).Value
   ActiveCell.Offset(-1, 0).Value = cell_content
   FixSymbolFont
   ActiveCell.Offset(-1, 0).Range("A1:B1").Select
   FixSymbolFont
   'ActiveCell.Comment.Shape.TextFrame.AutoSize = True
   FormatCellComent ActiveCell.Comment
   Application.ScreenUpdating = True
   End Sub
Sub SwapCalendarEntryDown()
   If ActiveWindow.ActiveSheet.Name <> "Calendar" Then Exit Sub
   On Error Resume Next
   Application.ScreenUpdating = False
   Dim cell_content
   Dim cell_comment As String
   cell_comment = ActiveCell.Comment.Text
   ActiveCell.ClearComments
   ActiveCell.AddComment Text:=ActiveCell.Offset(1, 0).Comment.Text
   ActiveCell.Offset(1, 0).ClearComments
   ActiveCell.Offset(1, 0).AddComment Text:=cell_comment
   'ActiveCell.Comment.Shape.TextFrame.AutoSize = True
   FormatCellComent ActiveCell.Comment
   ActiveCell.Offset(0, -1).Comment.Shape.TextFrame.AutoSize = True
   cell_content = ActiveCell.Value
   ActiveCell.Value = ActiveCell.Offset(1, 0).Value
   FixSymbolFont
   ActiveCell.Offset(1, 0).Value = cell_content
   ActiveCell.Offset(1, 0).Range("A1:B1").Select
   FixSymbolFont
   'ActiveCell.Comment.Shape.TextFrame.AutoSize = True
   FormatCellComent ActiveCell.Comment
   Application.ScreenUpdating = True
   End Sub
