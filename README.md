# wordonekeyVBA
VBA机关排版自用源码,基于office2007
#统一清除word文档格式、空格等
****************************************************************
Sub 统一清除格式()
'清除格式
Selection.WholeStory
Selection.ClearFormatting
'清除空行
    ActiveDocument.Content.Find.Execute _
    findtext:="[^11^13]{1,}", MatchWildcards:=True, _
    replacewith:="^p", Replace:=wdReplaceAll
'清除空格
空格
全角
End Sub
Sub 空格()
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        '假空格
        .Text = "^u160"
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll

        '换行符前空格
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "^32{1,}[^11^13]"
        .Replacement.Text = "^13"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll

        '换行符后空格
        .Text = "^13^32{1,}"
        .Replacement.Text = "^13"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll

        '行首后空格
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "<^32{1,}"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    
    '中文空格
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = " "
        .Replacement.Text = ""
        .Forward = True
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = True
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        Selection.Find.Execute Replace:=wdReplaceAll
    End With
  
End Sub
Sub 全角()
Dim temp As String
temp = ActiveDocument.Content.Text
temp = StrConv(temp, vbWide)
ActiveDocument.Content.Text = temp
End Sub
