VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private x As Excel.Range
Private LastTarget As Range
Private Const RowsPerDay = 7

'README - THIS IS HOW YOU SET BUTTON POSITIONS.  ADD IT TO THISWORKBOOK AS ERROR CONTROL

Sub Tester()
   Dim rng As Range
   Set rng = ActiveSheet.Range("C2:D2")
   With ActiveSheet.OLEObjects("MONDAY_BUTTON")
      .Top = rng.Top
      .Left = rng.Left
      .Width = rng.Width
      .Height = rng.RowHeight
   End With
End Sub

'*************************************************************
' WORKSHEET EVENT HANDLERS
'*************************************************************
Private Sub Worksheet_Activate()
   Application.ScreenUpdating = False
'   ActiveSheet.Columns(1).Hidden = True
'   ActiveSheet.Columns(13).Hidden = True
   Application.FormulaBarHeight = 1
   ZoomToSel "A1,N1"
   RibbonCollapse
   HeadingsNo
   CenterInUpperLeftCorner
   FreezeRowPane 2
   Application.ScreenUpdating = True
   End Sub
Private Sub Worksheet_SelectionChange(ByVal target As Range)
   If (LastTarget Is Nothing) = False Then
      If Not LastTarget.Comment Is Nothing Then
         LastTarget.Comment.Visible = False
      End If
    End If
   Set target = ActiveCell
   If Not target.Comment Is Nothing Then
      'ActiveCell.Comment.Shape.TextFrame.AutoSize = True
      FormatCellComent ActiveCell.Comment
      target.Comment.Visible = True
      PositionCalendarCellComment
   End If
   Set LastTarget = ActiveCell
   End Sub
Private Sub Worksheet_BeforeDoubleClick(ByVal target As Range, Cancel As Boolean)
   AddOrEditComment
   End Sub
'*************************************************************
' CLICK EVENT HANDLERS
'*************************************************************
Private Sub COLUMN_B_BUTTON_Click()
   AdvancePlaceholderOneWeek
   End Sub

Private Sub COLUMN_C_BUTTON_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
   On Error Resume Next
   Application.ScreenUpdating = False
   'LBUTTON
   If Button = 1 Then
      ClearActiveCellContentAndComments
   'RBUTTON
   ElseIf Button = 2 Then
      RemoveFormatsDarkGray
   End If
   Application.ScreenUpdating = True
   End Sub
Private Sub COLUMN_D_1_BUTTON_Click()
   FixSymbolFontForEachCellInSelectedRange
   End Sub
Private Sub COLUMN_D_2_BUTTON_Click()
   MarkAsComplete
   End Sub
Private Sub COLUMN_L_1_BUTTON_Click()
   SwapCalendarEntryDown
   End Sub
Private Sub COLUMN_L_2_BUTTON_Click()
   SwapCalendarEntryUp
   End Sub
Private Sub COLUMN_L_3_BUTTON_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
   On Error Resume Next
   Application.ScreenUpdating = False
   'LBUTTON
   If Button = 1 Then
      ToggleHiddenRows
   'RBUTTON
   ElseIf Button = 2 Then
      ToggleHiddenRows
   End If
   Application.ScreenUpdating = True
   End Sub
Private Sub MONDAY_Click()
   ShadeDayIfPast (3)
   End Sub
Private Sub TUESDAY_Click()
   ShadeDayIfPast (5)
   End Sub
Private Sub WEDNESDAY_Click()
   ShadeDayIfPast (7)
   End Sub
Private Sub THURSDAY_Click()
   ShadeDayIfPast (9)
   End Sub
Private Sub FRIDAY_Click()
   Application.ScreenUpdating = True
   ShadeDayIfPast (11)
   AdvancePlaceholderOneWeek
   Application.ScreenUpdating = False
   End Sub
Private Sub ROW2_COLUMN_B_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
   On Error Resume Next
   Application.ScreenUpdating = False
   'LBUTTON
   If Button = 1 Then
      ActiveWindow.DisplayHeadings = Not ActiveWindow.DisplayHeadings = True
   'RBUTTON
   ElseIf Button = 2 Then
      MsgBox "Toggle row/column headings."
   End If
   Application.ScreenUpdating = True
   End Sub
Private Sub Hide_Click()
   On Error Resume Next
   Application.ScreenUpdating = False
   Const WHAT_TO_FIND As String = ". . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . ."
   Dim ws As Excel.Worksheet
   Dim FoundCell As Excel.Range
   Set ws = ActiveSheet
   Set FoundCell = ws.Range("A1:K5000").Find(what:=WHAT_TO_FIND)
   ActiveSheet.Rows(3 & ":" & FoundCell.row).Hidden = True
   Application.ScreenUpdating = True
   End Sub
Private Sub resetConditionalFormatting_Click()
   On Error Resume Next
   Application.ScreenUpdating = False
   getInteriorPropertiesFromSelection
   Selection.FormatConditions.Delete
   Range("A1:B1").Copy
   Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.ScreenUpdating = True
   End Sub
Private Function RowThisWeek() As Long
   On Error Resume Next
   Application.ScreenUpdating = False
   Const WHAT_TO_FIND As String = ". . . . ."
   Dim ws As Excel.Worksheet
   Dim FoundCell As Excel.Range
   Set ws = ActiveSheet
   Set FoundCell = ws.Range("A1:K5000").Find(what:=WHAT_TO_FIND)
   RowThisWeek = FoundCell.row + 1
   End Function
Private Sub WIDE_BAR_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
   On Error Resume Next
   Application.ScreenUpdating = False
   'LBUTTON
   If Button = 1 Then
      ToggleHiddenRows
   'RBUTTON
   ElseIf Button = 2 Then
      ToggleHiddenRows
   End If
   Application.ScreenUpdating = True
   End Sub
Private Sub ToggleHiddenRows(Optional ToggleOption As String = "toggle")
   On Error Resume Next
   Application.ScreenUpdating = False
   Const WHAT_TO_FIND As String = ". . . . ."
   Dim ws As Excel.Worksheet
   Dim FoundCell As Excel.Range
   Set ws = ActiveSheet
   Set FoundCell = ws.Range("A1:K5000").Find(what:=WHAT_TO_FIND)
   If ((ToggleOption = "toggle") And (ActiveSheet.Rows(3 & ":" & FoundCell.row).Hidden = True)) Or _
      (ToggleOption = "hide") Then
      Columns("M:M").EntireColumn.Hidden = False
      ZoomToSel "A1,N1"
      ActiveSheet.Rows(3 & ":" & FoundCell.row).Hidden = False
      'Call FreezeRowPane(2, 13)
      FreezeRowPane 2
      ActiveWindow.ScrollRow = FoundCell.row - 12
      Cells(Application.ActiveCell.row, 1).Select
      Set FoundCell = ws.Range("A1:K5000").Find(what:="")
   ElseIf ((ToggleOption = "toggle") And ActiveSheet.Rows(3 & ":" & FoundCell.row).Hidden = False) Or _
         (ToggleOption = "show") Then
      Columns("M:M").EntireColumn.Hidden = True
      ZoomToSel "A1,N1"
      Set ws = ActiveSheet
      Set FoundCell = ws.Range("A1:K5000").Find(what:=WHAT_TO_FIND)
      ActiveSheet.Rows(3 & ":" & FoundCell.row).Hidden = True
      ActiveSheet.Cells(FoundCell.row + 1, 2).Select
      Cells(Application.ActiveCell.row, 1).Select
      Set FoundCell = ws.Range("A1:K5000").Find(what:="")
   End If
   Application.ScreenUpdating = True
   End Sub
Private Sub FreezeCalendarPanes()
   If ActiveWindow.FreezePanes = False Then
      ActiveWindow.SplitColumn = 13
      ActiveWindow.SplitRow = 2
      ActiveWindow.FreezePanes = True
   End If
   End Sub
'*************************************************************
' SUBROUTINES/FUNCTIONS
'*************************************************************
Private Sub AdvancePlaceholderOneWeek()
   On Error Resume Next
   Cells.Find(what:=". . . . .", After:=ActiveCell, LookIn:=xlFormulas, _
   LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
   MatchCase:=False, SearchFormat:=False).Activate

   Selection.Copy
   ActiveCell.Offset(7, 0).Range("A1:B1").Select
   ActiveSheet.Paste
   ActiveCell.Offset(-7, 0).Range("A1:B1").Select
   Application.CutCopyMode = False
   Selection.ClearContents

   Const WHAT_TO_FIND As String = ". . . . ."
   Dim ws As Excel.Worksheet
   Dim FoundCell As Excel.Range

   Set ws = ActiveSheet
   Set FoundCell = ws.Range("A1:K5000").Find(what:=WHAT_TO_FIND)
   ActiveSheet.Rows(3 & ":" & FoundCell.row).Hidden = True

   ActiveSheet.Cells(FoundCell.row + 1, 2).Select
   Cells(Application.ActiveCell.row, 1).Select
   Set FoundCell = ws.Range("A1:K5000").Find(what:="")
   End Sub
Private Sub getInteriorPropertiesFromSelection()
   With Range("A1").Interior
      .Pattern = ActiveCell.Interior.Pattern
      .PatternColorIndex = ActiveCell.Interior.PatternColorIndex
      .Color = ActiveCell.Interior.Color
      .TintAndShade = ActiveCell.Interior.TintAndShade
      .PatternTintAndShade = ActiveCell.Interior.PatternTintAndShade
   End With
   End Sub
Private Sub MarkAsComplete()
   'For original code, see MarkAsComplete_OLD_CODE_Click() in ARCHIVE module
   Dim stringBefore, stringAfter As String
   stringBefore = ActiveCell.Text
   If Left(stringBefore, 1) = "!" Or Left(stringBefore, 1) = ">" Or Left(stringBefore, 1) = "|" Then
      stringAfter = Right(stringBefore, Len(stringBefore) - 1)
      stringAfter = Trim(stringAfter)
      ActiveCell.Value = stringAfter
   End If
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
Private Function ShadeDayIfPast(DayColumn As Long)
   On Error Resume Next
   Dim cell, CalDay As Range
   Dim r1, r2 As Long
   Dim c1, c2 As String
   Dim i As Integer
   Dim PriorScreenUpdating As Boolean
   PriorScreenUpdating = Application.ScreenUpdating
   If Application.ScreenUpdating = True Then Application.ScreenUpdating = False
   r1 = RowThisWeek
   r2 = r1 + RowsPerDay - 1
   c1 = ColumnAsLetter(DayColumn)
   c2 = ColumnAsLetter(DayColumn + 1)
   Set CalDay = Range(c1 & r1 & ":" & c2 & r2)
   'Msgbox Range(c1 & r1 & ":" & c2 & r2).Cells(1, 1).Value
   For Each cell In CalDay
      With cell.Interior
      .Pattern = xlSolid
      .PatternColorIndex = xlAutomatic
      .ThemeColor = xlThemeColorDark1
      .TintAndShade = -0.149998474074526
      .PatternTintAndShade = 0
      End With
      cell.FormatConditions.Delete
   Next
   If PriorScreenUpdating Then Application.ScreenUpdating = True
   End Function
Private Sub Strikethrough()
    On Error Resume Next
   Application.ScreenUpdating = False
   Selection.Font.Strikethrough = Not Selection.Font.Strikethrough
   Selection.FormatConditions.Delete
   Selection.Font.Bold = False
   Application.ScreenUpdating = True
   End Sub




