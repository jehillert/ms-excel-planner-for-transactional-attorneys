Attribute VB_Name = "Miscellaneous"
'settings
Option Explicit
' DIAGNOSTICS
'Public Declare Function GetSystemMetrics Lib "user32.dll" (ByVal index As Long) As Long
Public Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal index As Long) As Long

Public Const SM_CXSCREEN = 0
Public Const SM_CYSCREEN = 1
Sub RibbonHeight()
   MsgBox Application.CommandBars("Ribbon").Height
'     On Dell 19" monitors, max res:
'        collapsed height   108
'        expanded height   194
End Sub

Sub ScreenDimensions()
   MsgBox GetSystemMetrics(SM_CXSCREEN) & "x" & GetSystemMetrics(SM_CYSCREEN)
   End Sub
Sub ShowCommandBars()
   For Each cbar In CommandBars
      Debug.Print cbar.Name, cbar.NameLocal, cbar.Visible
   Next
   End Sub
Sub UnhideAllSheets()
   Dim ws As Worksheet
   For Each ws In ActiveWorkbook.Worksheets
      ws.Visible = xlSheetVisible
   Next ws
   End Sub
' MISCELLANEOUS
Sub AnchorImage()
    Selection.Placement = xlMoveAndSize
   End Sub
Sub Center()
   With Selection
      .HorizontalAlignment = xlCenter
      .VerticalAlignment = xlDistributed
      .WrapText = False
      .Orientation = 0
      .AddIndent = False
      .IndentLevel = 0
      .ShrinkToFit = False
      .ReadingOrder = xlContext
      .MergeCells = True
   End With
   End Sub
Sub ClearFormats()
   Selection.FormatConditions.Delete
   End Sub
Sub CopyFormat()
   Selection.Copy
   End Sub
Sub FormatPainter()
   If TypeName(Selection) = "Range" Then
      Selection.Copy
      ActiveWorkbook.Names.Add "FormatPainter.Flag", RefersTo:=True, Visible:=False
   End If
   End Sub
Public Sub Clear_FormatPainter_Flag()
   On Error Resume Next
   ActiveWorkbook.Names("FormatPainter.Flag").Delete
   End Sub
Public Function NameExists(ByVal sName As String) As Boolean
   On Error Resume Next
   NameExists = ActiveWorkbook.Names(sName).Name = sName
   End Function
Sub PasteFormat()
   Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
   SkipBlanks:=False, Transpose:=False
   Application.CutCopyMode = False
   End Sub
Sub SideBySide()
   Windows.Arrange ArrangeStyle:=xlVertical
   End Sub
Sub ToggleHeadingsFormulaBar()
   If ActiveWindow.DisplayHeadings = "True" Then
      HeadingsNo
      Application.DisplayFormulaBar = False
   ElseIf ActiveWindow.DisplayHeadings = "False" Then
      HeadingsYes
      Application.DisplayFormulaBar = True
   End If
   End Sub
' OUTLOOK MACROS
Sub MailSelectedOpen()
   Dim OutApp As Outlook.Application
   Dim objOutlookMsg As Outlook.MailItem
   Dim objOutlookRecip As Recipient
   Dim Recipients As Recipients
   Dim RngCopied As Range
'   Set RngCopied = Selection
   Selection.Copy
   Set OutApp = CreateObject("Outlook.Application")
   Set objOutlookMsg = OutApp.CreateItem(olMailItem)
   'add recipients
   Set Recipients = objOutlookMsg.Recipients
   Set objOutlookRecip = Recipients.Add("kschoen@scheinbergip.com")
   objOutlookRecip.Type = 1
   Set objOutlookRecip = Recipients.Add("nelson@scheinbergip.com")
   objOutlookRecip.Type = olCC
   Set objOutlookRecip = Recipients.Add("tray@scheinbergip.com")
   objOutlookRecip.Type = olCC
   'add subject & body
   objOutlookMsg.SentOnBehalfOfName = "John Hillert"
   objOutlookMsg.Subject = "Time Entry - "
   objOutlookMsg.HTMLBody = ""
   'resolve each recipient's name
   For Each objOutlookRecip In objOutlookMsg.Recipients
      objOutlookRecip.Resolve
   Next
   'send or display message
   'objOutlookMsg.Send
   objOutlookMsg.Display
   Set OutApp = Nothing
   SendKeys "^v", True
   SendKeys "^{HOME}{TAB}^c+{TAB}+{TAB}{END} ^v", True
   
   End Sub
Sub MailSelectedSend()
   Dim OutApp As Object
   Dim OutMail As Object
   Dim RngCopied As Range
   Set OutApp = CreateObject("Outlook.Application")
   Set OutMail = OutApp.CreateItem(0)
   Set RngCopied = Selection
   'On Error Resume Next
   With OutMail
      .To = "jhillert@scheinbergip.com"
      .CC = ""
      .BCC = ""
      .Subject = "Time Entry - "
      .HTMLBody = RangetoHTML(RngCopied)
      '.Attachments.Add ActiveWorkbook.FullName
      .Send
   End With
   On Error GoTo 0
   Set OutMail = Nothing
   Set OutApp = Nothing
End Sub
'RANGE TO HTML Function - Changed by Ron de Bruin 28-Oct-2006. Working in Office 2000-2010
Function RangetoHTML(rng As Range)
   Dim fso As Object
   Dim ts As Object
   Dim TempFile As String
   Dim TempWB As Workbook
   TempFile = Environ$("temp") & "/" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"
   'Copy the range and create a new workbook to past the data in
   rng.Copy
   Set TempWB = Workbooks.Add(1)
   With TempWB.Sheets(1)
      .Cells(1).PasteSpecial Paste:=8
      .Cells(1).PasteSpecial xlPasteValues, , False, False
      .Cells(1).PasteSpecial xlPasteFormats, , False, False
      .Cells(1).Select
      Application.CutCopyMode = False
      On Error Resume Next
      .DrawingObjects.Visible = True
      .DrawingObjects.Delete
      On Error GoTo 0
   End With
   'Publish the sheet to a htm file
   With TempWB.PublishObjects.Add( _
      SourceType:=xlSourceRange, _
      Filename:=TempFile, _
      Sheet:=TempWB.Sheets(1).Name, _
      Source:=TempWB.Sheets(1).UsedRange.Address, _
      HtmlType:=xlHtmlStatic)
      .Publish (True)
   End With
   'Read all data from the htm file into RangetoHTML
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
   RangetoHTML = ts.ReadAll
   ts.Close
   RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
             "align=left x:publishsource=")
   'Close TempWB
   TempWB.Close SaveChanges:=False
   'Delete the htm file we used in this function
   Kill TempFile
   Set ts = Nothing
   Set fso = Nothing
   Set TempWB = Nothing
   End Function
' ROWS, COLUMNS, GROUPS
Sub collapse_all_groups()
   ActiveSheet.Outline.ShowLevels RowLevels:=1, ColumnLevels:=1
   End Sub
Sub DeleteThisRow()
   Application.ScreenUpdating = False
   Selection.EntireRow.Delete
   Cells(Application.ActiveCell.row, 1).Select
   Application.ScreenUpdating = True
   End Sub
Sub expand_all_groups()
   ActiveSheet.Outline.ShowLevels RowLevels:=8, ColumnLevels:=8
End Sub
Sub InsertColumnLeft()
   On Error Resume Next
   ActiveCell.EntireColumn.Insert: Selection.Offset(0, 1).Select
   End Sub
Sub InsertColumnRight()
   ActiveCell.EntireColumn.Offset(0, 1).Insert
   End Sub
Sub insert_row()
   Selection.EntireRow.Insert
   Cells(Application.ActiveCell.row, 1).Select
   End Sub
Sub InsertRowAbove()
   On Error Resume Next
   ActiveCell.EntireRow.Insert: Selection.Offset(-1, 0).Select
   End Sub
Sub InsertRowBelow()
   ActiveCell.Offset(1).EntireRow.Insert
   End Sub
Sub MoveCellUp()
   On Error Resume Next
   Selection.Cut: Selection.Offset(-1, 0).Select: Selection.Insert Shift:=xlDown
   End Sub
Sub MoveCellDown()
   On Error Resume Next
   Selection.Cut: Selection.Offset(2, 0).Select: Selection.Insert Shift:=xlDown
   End Sub
Sub MoveRowUp()
   On Error Resume Next
   Selection.EntireRow.Select: Selection.Cut: Selection.Offset(-1, 0).Select: Selection.Insert Shift:=xlDown
   End Sub
Sub MoveRowDown()
   On Error Resume Next
   Selection.EntireRow.Select: Selection.Cut: Selection.Offset(2, 0).Select: Selection.Insert Shift:=xlUp
   End Sub
Sub RowAdd()
   On Error Resume Next
   Selection.EntireRow.Insert: Cells(Application.ActiveCell.row, 1).Select
   End Sub
Sub RowDelete()
   On Error Resume Next
   Selection.EntireRow.Delete: Cells(Application.ActiveCell.row, 1).Select
   End Sub
Sub RowHide()
   Selection.EntireRow.Hidden = True
   End Sub
Sub RowUnhide()
   Selection.EntireRow.Hidden = False
   End Sub
' TEXT
Sub RedText()
   Selection.Font.Color = -16776961
   End Sub
Sub BlackText()
   Selection.Font.ThemeColor = xlThemeColorLight1
   End Sub
Function ScrollRow(myCell As Range)
   On Error Resume Next
   Application.ScreenUpdating = False
   Dim r As Range
   Set r = myCell
   numColum = r.Columns.Count
   numRow = r.Rows.Count
   Range("A2").Select
   With ActiveWindow
      .FreezePanes = False
      .ScrollRow = 1
      .ScrollColumn = 1
      .FreezePanes = True
      .ScrollRow = r.row
   End With
   r.Select
   Application.ScreenUpdating = True
   End Function
Sub tryscroll()
   Call ScrollRow(ActiveSheet.Range("A36"))
   End Sub

