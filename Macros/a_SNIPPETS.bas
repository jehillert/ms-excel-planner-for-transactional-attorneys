Attribute VB_Name = "a_SNIPPETS"
Sub TestSnippet()
   ActiveWindow.SmallScroll Down:=-3
   End Sub
Sub MYSNIPPETS()
   Selection.FormatConditions.Delete
   Selection.ClearContents
   ActiveWindow.SmallScroll Down:=-3
   Application.ScreenUpdating = False
   Application.ScreenUpdating = True
End Sub

'If ActiveWindow.ActiveSheet.Name = "_____" Then

'EXCELLENT GUIDE FOR COMMENTS:
'     http://www.contextures.com/xlcomments03.html
'     www.contextures.com/xlcomments03.html

'   SNAPPING CONTROLS
'     You can "snap" any shape such as a Command Button in a cell by dragging one of its
'     corner handles while in Design mode and (here's the trick) pressing the Alt key. Then
'     do the same thing with the opposite corner handle (example bottom right), again, while
'     holding down the Alt key.
'     Release the Alt key, set the property for Move and Size with cells, and exit Design mode.

'CONTEXT MENU
'     great reference:
'     https://msdn.microsoft.com/en-us/library/office/gg469862(v=office.14).aspx
'     Copy the following two event procedures into the ThisWorkbook module of your workbook.
'     These events automatically add the controls to the Cell context menu when you open or
'     activate the workbook and delete the controls when you close or deactivate the workbook.

'Dim x As Range
'Private Sub Worksheet_SelectionChange(ByVal Target As Range)
'   On Error Resume Next
'   If Not x Is Nothing Then x.Comment.Visible = False
'   Target.Comment.Visible = True
'   Set x = Target
'   End Sub

'Private Sub Workbook_AddinUninstall()
'Private Sub Workbook_BeforeClose(Cancel As Boolean)
'Private Sub Workbook_BeforePrint(Cancel As Boolean)
'Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
'Private Sub Workbook_Deactivate()
'Private Sub Workbook_NewSheet(ByVal Sh As Object)
'Private Sub Workbook_Open()
'Private Sub Workbook_SheetActivate(ByVal Sh As Object)
'Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
'Private Sub Workbook_SheetBeforeRightClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
'Private Sub Workbook_SheetCalculate(ByVal Sh As Object)
'Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
'Private Sub Workbook_SheetDeactivate(ByVal Sh As Object)
'Private Sub Workbook_SheetFollowHyperlink(ByVal Sh As Object, ByVal Target As Hyperlink)
'Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
'Private Sub Workbook_WindowActivate(ByVal Wn As Window)
'Private Sub Workbook_WindowDeactivate(ByVal Wn As Window)
'Private Sub Workbook_WindowResize(ByVal Wn As Window)
Private Sub Workbook_Deactivate()

End Sub
