Attribute VB_Name = "Comments"
Sub AutoFitComments()
   Dim xComment As Comment
   For Each xComment In Application.ActiveSheet.Comments
       xComment.Shape.TextFrame.AutoSize = True
   Next
   End Sub
Sub ScaleAllCommentsToDefault()
   Dim xComment As Comment
   For Each xComment In Application.ActiveSheet.Comments
      xComment.Shape.TextFrame.AutoSize = False
'      xComment.Shape.ScaleHeight DefaultScaleHeight, 0, 0
'      xComment.Shape.ScaleWidth DefaultScaleWidth, 0, 0
'xComment.Shape.TextEffect.FontName = 'Arial Nova'
      xComment.Shape.Width = DefaultCommentWidth
      xComment.Shape.Height = DefaultCommentHeight
   Next
   End Sub
Sub CopyFromAbove()
   On Error Resume Next
   Selection.FillDown
   ActiveCell.Offset(0, 1).Select
   End Sub
Sub AddOrEditComment()
   On Error GoTo CommentErrHandler
   Dim cmt As Comment
   Dim CommentText As String
   Set cmt = ActiveCell.Comment
   Cancel = True
   If Not cmt Is Nothing Then
      CommentText = ActiveCell.Comment.Text
   Else
      CommentText = ""
   End If
   ActiveCell.ClearComments
   ActiveCell.AddComment Text:=CommentText
   PositionCalendarCellComment
   ActiveCell.Comment.Visible = True
   FormatCellComent ActiveCell.Comment
   'PositionCalendarCellComment
   ActiveCell.Comment.Shape.Select True
Done:
   Exit Sub
CommentErrHandler:
   End Sub
Sub PasteCBTxtToCell()
   On Error Resume Next
   Dim DataObj As MSForms.DataObject
   Set DataObj = New MSForms.DataObject
   DataObj.GetFromClipboard
   strPaste = DataObj.GetText(1)
   ActiveCell.FormulaR1C1 = strPaste
   Cells(Application.ActiveCell.row, 1).Select
   End Sub

'RESIZING, SCALING & POSITIONING
Sub ScaleToDefault()
   ActiveCell.Comment.Shape.TextFrame.AutoSize = True
   ActiveCell.Comment.Shape.ScaleHeight DefaultScaleHeight, 0, 0
   ActiveCell.Comment.Shape.ScaleWidth DefaultScaleWidth, 0, 0
   End Sub
Sub lkjlkj()
   FormatCellComent ActiveCell.Comment
End Sub
'FORMATTING CELL COMMENTS
Sub FormatCellComent(myComment As Comment)
   Application.ScreenUpdating = False
   With myComment
      .Shape.TextFrame.AutoSize = True
       If .Shape.Width > DefaultCommentWidth Then
         .Shape.TextFrame.AutoSize = True
         .Shape.Width = DefaultCommentWidth
         .Shape.Height = DefaultCommentHeight
       End If
      .Shape.AutoShapeType = msoShapeRoundedRectangle
      .Shape.TextFrame.Characters.Font.Name = "Tahoma"
      .Shape.TextFrame.Characters.Font.Size = 8
      .Shape.TextFrame.Characters.Font.ColorIndex = 2
      .Shape.Line.ForeColor.RGB = RGB(0, 0, 0)
      .Shape.Line.BackColor.RGB = RGB(255, 255, 255)
      .Shape.Fill.Visible = msoTrue
      .Shape.Fill.ForeColor.RGB = RGB(58, 82, 184)
      .Shape.Fill.OneColorGradient msoGradientDiagonalUp, 1, 0.23
   End With
   Application.ScreenUpdating = True
   End Sub
Sub PositionCalendarCellComment()
   If ActiveWindow.ActiveSheet.Name <> "Calendar" Then Exit Sub
   With ActiveCell
   If .Left + .Comment.Shape.Width >= .Offset(0, 1).Left And _
          (.Column = 11 Or .Column = 12) Then
      .Comment.Shape.Left = .Offset(0, 1).Left - .Comment.Shape.Width * 1.01
      .Comment.Shape.Top = .Top + .Height * 1.01
   Else
      .Comment.Shape.Left = .Offset(0, 1).Left * 1.01
      .Comment.Shape.Top = .Top
   End If
   End With
End Sub
