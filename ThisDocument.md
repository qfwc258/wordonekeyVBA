Private Sub Document_Open()
Dim NewMenu As CommandBarPopup
    Dim MenuItem As CommandBarControl
    On Error Resume Next
    '如果菜单已存在,则删除该菜单
    Application.CommandBars(1).Controls("自动排版").Delete
    Set NewMenu = Application.CommandBars(1).Controls.Add(Type:=msoControlPopup, Before:=11)
     
    '添加菜单标题并指定热键
    NewMenu.Caption = "自动排版"
   
    '添加第一个菜单项
    Set MenuItem = NewMenu.Controls.Add _
      (Type:=msoControlButton)
    With MenuItem
        .Caption = "自动排版"
        .OnAction = "自动排版"
    End With
   
    Set MenuItem = NewMenu.Controls.Add _
      (Type:=msoControlButton)
    With MenuItem
        .Caption = "红头文件"
        .OnAction = "红头文件"
    End With
   
    Set MenuItem = NewMenu.Controls.Add _
      (Type:=msoControlButton)
    With MenuItem
        .Caption = "章-省政府"
        .OnAction = "gz"
    End With
   
    Set MenuItem = NewMenu.Controls.Add _
      (Type:=msoControlButton)
    With MenuItem
        .Caption = "章-中国报道"
        .OnAction = "zgbd"
    End With
   
    End Sub
