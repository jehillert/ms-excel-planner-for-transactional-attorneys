Attribute VB_Name = "UI"
Sub HeadingsYes()
   ActiveWindow.DisplayHeadings = True
   End Sub
Sub HeadingsNo()
   ActiveWindow.DisplayHeadings = False
   End Sub
Sub EnterInDesignMode()
   Application.CommandBars.FindControl(ID:=1605).Execute
   End Sub
Sub ExitInDesignMode()
   Dim sTemp As String
   With Application.CommandBars("Exit Design Mode")
      sTemp = .Controls(1).Caption
   End With
   End Sub
Sub RemoveToolbars()
   On Error Resume Next
   With Application
      .DisplayFullScreen = True
      .CommandBars("Full Screen").Visible = False
      .CommandBars("MyToolbar").Enabled = True
      .CommandBars("MyToolbar").Visible = True
      .CommandBars("Worksheet Menu Bar").Enabled = False
   End With
   On Error GoTo 0
   End Sub
Sub RestoreToolbars()
   On Error Resume Next
   With Application
      .DisplayFullScreen = False
      .CommandBars("MyToolbar").Enabled = False
      .CommandBars("Worksheet Menu Bar").Enabled = True
   End With
   On Error GoTo 0
   End Sub
Sub RibbonCollapse()
   If Application.CommandBars("Ribbon").Height > 150 _
      Then CommandBars.ExecuteMso "MinimizeRibbon"
   End Sub
Sub RibbonExpand()
   If Application.CommandBars("Ribbon").Height < 150 _
      Then CommandBars.ExecuteMso "MinimizeRibbon"
   End Sub
Sub RibbonHide()
   Application.ExecuteExcel4Macro "show.toolbar(""Ribbon"",False)"
   End Sub
Sub RibbonShow()
   Application.ExecuteExcel4Macro "show.toolbar(""Ribbon"",True)"
   End Sub

