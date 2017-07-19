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
    
    ******************************************
Sub gz()
On Error Resume Next
    Selection.InlineShapes.AddPicture FileName:=ThisDocument.Path & "\湖南省人民政府章.png", _
        LinkToFile:=False, SaveWithDocument:=True
End Sub
Sub zgbd()
On Error Resume Next
    Selection.InlineShapes.AddPicture FileName:=ThisDocument.Path & "\中国报道.png", _
        LinkToFile:=False, SaveWithDocument:=True
End Sub
Sub 初始化()
Dim file As Field
For Each file In ActiveDocument.Fields
file.Unlink
Next
ActiveDocument.Content.ListFormat.ConvertNumbersToText
Selection.WholeStory
End Sub
Sub 页面设置()
With ActiveDocument.PageSetup
.TopMargin = CentimetersToPoints(3.7)
.BottomMargin = CentimetersToPoints(3.5)
.LeftMargin = CentimetersToPoints(2.8)
.RightMargin = CentimetersToPoints(2.6)
.PaperSize = wdPaperA4
.HeaderDistance = CentimetersToPoints(1.5)
.FooterDistance = CentimetersToPoints(2.8)
.CharsLine = 42
.OddAndEvenPagesHeaderFooter = False
End With
End Sub
Sub 段落设置()
With ActiveDocument.Paragraphs
.Alignment = wdAlignParagraphJustify
.LeftIndent = CentimetersToPoints(0)
.CharacterUnitFirstLineIndent = 2
.LineSpacingRule = wdLineSpaceExactly
.LineSpacing = 28
End With
End Sub
Sub 字体()
Selection.WholeStory
Selection.Font.Name = "仿宋"
Selection.Font.Size = 16
ActiveDocument.Paragraphs(1).Range.Select
With Selection
.ClearFormatting
With .Font
.Size = 22
.Name = "方正小标宋简体"
End With
With .ParagraphFormat
.CharacterUnitFirstLineIndent = 0
.Alignment = wdAlignParagraphCenter
.LineSpacingRule = wdLineSpaceExactly
.LineSpacing = 35
.SpaceAfter = 20
End With
.HomeKey Unit:=wdStory
End With
End Sub
Sub 页脚格式()
WordBasic.ViewFooterOnly
'WordBasic.GoToFooter
With Selection
.WholeStory
.Delete
.TypeText "－－"
.MoveLeft wdCharacter, 1
.Fields.Add Selection.Range, -1, "Page \ * mergeformat", False
.MoveRight wdCharacter, 1
.WholeStory
With .Font
.Name = "宋体"
.Size = 16
End With
.ParagraphFormat.Alignment = wdAlignParagraphRight
End With
WordBasic.RemoveHeader
Selection.EscapeKey
End Sub
Sub 奇偶页相同()
ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
End Sub
Sub 奇偶页不同()
ActiveDocument.PageSetup.OddAndEvenPagesHeaderFooter = True
ActiveWindow.ActivePane.View.SeekView = wdSeekEvenPagesFooter
Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
End Sub
Sub 页脚格式判断()
If PRD = 2 And (Selection.Information(wdNumberOfPagesInDocument) > 1) Then
ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
End If
End Sub
Sub 字符格式()
With ActiveDocument.Content.Font
.Name = "仿宋"
.Size = 16
.Color = wdColorBlack
End With
End Sub
Sub 自动排版()
PRT = 2
初始化
页面设置
段落设置
字体
页脚格式
页脚格式判断
'红头文件
End Sub

Sub 红头文件()
For i = 1 To 5
    Selection.TypeParagraph
    Next
Dim myShape As Shape
For Each myShape In ActiveDocument.Shapes
myShape.Delete
Next
ActiveDocument.Shapes.AddTextbox(msoTextOrientationHorizontal, 80.8, _
        185, 439.35, 100.3).Select
    Selection.ShapeRange.Fill.Visible = msoFalse
    Selection.ShapeRange.Line.Visible = msoFalse
    Selection.ShapeRange.TextFrame.TextRange.Select
    Selection.TypeText Text:="湖南省人民政府文件"
    Selection.WholeStory
    With Selection.Font
        .NameFarEast = "方正小标宋简体"
        .NameAscii = "方正小标宋简体"
        .NameOther = "+西文正文"
        .Name = "方正小标宋简体"
        .Size = 60
        .Bold = False
        .Italic = False
        .Underline = wdUnderlineNone
        .UnderlineColor = wdColorAutomatic
        .StrikeThrough = False
        .DoubleStrikeThrough = False
        .Outline = False
        .Emboss = False
        .Shadow = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = False
        .Color = wdColorRed
        .Engrave = False
        .Superscript = False
        .Subscript = False
        .Spacing = 4
        .Scaling = 60
        .Position = 0
        .Kerning = 1
        .Animation = wdAnimationNone
        .DisableCharacterSpaceGrid = False
        .EmphasisMark = wdEmphasisMarkNone
    End With
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
  ActiveDocument.Shapes.AddConnector(msoConnectorStraight, 70#, 352.55, _
        346.7, 0#).Select
With Selection.ShapeRange
    .Width = 439.35
    .Line.Weight = 2.25
    .Left = wdShapeCenter
    .Line.ForeColor.RGB = RGB(255, 0, 0)
End With
 ActiveDocument.Shapes.AddTextbox(msoTextOrientationHorizontal, 195.4, _
        310.4, 203.25, 31.65).Select
With Selection
   .ShapeRange.TextFrame.TextRange.Select
       .TypeText Text:="湘政发"
    .InsertSymbol Font:="+中文正文", CharacterNumber:=12308, Unicode:= _
        True
      .TypeText Text:="2017"
    .InsertSymbol Font:="+中文正文", CharacterNumber:=12309, Unicode:= _
        True
    .TypeText Text:="00000号"
   .WholeStory
    .ParagraphFormat.Alignment = wdAlignParagraphCenter
    .Font.Name = "仿宋"
    .Font.Size = 16
    .ShapeRange.Fill.Visible = msoFalse
    .ShapeRange.Line.Visible = msoFalse
End With
End Sub
