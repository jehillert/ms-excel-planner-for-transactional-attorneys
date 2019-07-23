Attribute VB_Name = "Calendar_CheckmarkFunctions"
Sub ChangeCharIfMatchesCondition(ByVal cell As Range)
   On Error Resume Next
   Application.ScreenUpdating = False
   Dim i As Integer
   Dim FindChar As String
   Dim SearchString As String
   SearchString = cell.Value
   FindChar = Chr(252)
   For i = 1 To Len(SearchString)
      If Mid(SearchString, i, 1) = FindChar Then cell.Characters(i, 1).Font.Name = "Wingdings"
   Next i
   Application.ScreenUpdating = True
   End Sub
Sub FixSymbolFont()
   On Error Resume Next
   Dim i As Integer
   Dim FindChar As String
   Dim SearchString As String
   ActiveCell.Font.Name = MY_FONT
   SearchString = ActiveCell.Value
   FindChar = Chr(252)
   For i = 1 To Len(SearchString)
   If Mid(SearchString, i, 1) = FindChar Then
      ActiveCell.Characters(i, 1).Font.Name = "Wingdings"
   End If
   Next i
   End Sub
Sub FixSymbolFontForEachCellInSelectedRange()
   Dim cell As Range
   Dim i As Integer
   Dim FindChar As String
   Dim SearchString As String
   For Each cell In Selection
      cell.Font.Name = MY_FONT
      SearchString = cell.Value
      FindChar = Chr(252)
      For i = 1 To Len(SearchString)
      If Mid(SearchString, i, 1) = FindChar Then
         cell.Characters(i, 1).Font.Name = "Wingdings"
      End If
      Next i
   Next
   End Sub
