VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FirstLevelTitle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'一级中文标题
'段落：
'wdLineSpaceSingle单倍行距
'LineUnitBefore段前0.5行距
'LineUnitAfter段后0行距

'字体：
'中文字体为黑体
'英文字体为Times New Roman
'小2(18)

Public Sub changeFormat()

    setFormat
    
    If checkFormat() = False Then
        '如果格式不对
        setFormat
    End If

End Sub
Public Function checkFormat() As Boolean
    
    '检查格式是否正确
    checkFormat = CheckFont() And CheckParagraphFormat()

End Function


Private Function CheckParagraphFormat() As Boolean

    If Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceSingle And _
        Selection.ParagraphFormat.LineUnitBefore = 0.5 And _
        Selection.ParagraphFormat.LineUnitAfter = 0 Then
        '格式正确，返回true
        CheckParagraphFormat = True
    Else
        '格式不对，打印提示
        popMsg ("段落错误")
        CheckParagraphFormat = False
    End If
    
    Exit Function
    
End Function

Private Function CheckFont() As Boolean
    '检查字体是否正确
    
    If Selection.Font.NameFarEast = "黑体" And _
        Selection.Font.NameAscii = "Times New Roman" And _
        Selection.Font.Size = 18 And _
        Selection.Font.Bold = True Then
        '格式正确，返回true
        CheckFont = True
    Else
        '格式不对，打印提示
        popMsg ("字体错误")
        CheckFont = False
    End If

    Exit Function
    
End Function


Private Sub setFormat()

    '这是用录制宏命令的代码
    With Selection.ParagraphFormat
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 2.5
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceSingle
        .Alignment = wdAlignParagraphJustify
        .WidowControl = False
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0.5
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
        .AutoAdjustRightIndent = True
        .DisableLineHeightGrid = False
        .FarEastLineBreakControl = True
        .WordWrap = True
        .HangingPunctuation = True
        .HalfWidthPunctuationOnTopOfLine = False
        .AddSpaceBetweenFarEastAndAlpha = True
        .AddSpaceBetweenFarEastAndDigit = True
        .BaseLineAlignment = wdBaselineAlignAuto
    End With


    With Selection.Font
        .NameFarEast = "黑体" '设置中文字体为黑体
        .NameAscii = "Times New Roman" '设置英文字体为Times New Roman
        .Size = 18 '小2
        .Bold = True '加粗
    End With

End Sub

Private Sub popMsg(Msg As String)

    Dim Style, title, Help, Ctxt, Response, MyString
    'Msg = "一级标题格式错误 ?"    ' 定义信息。
    Style = vbYesNo + vbCritical + vbDefaultButton2    ' 定义按钮。
    title = "检查一级标题格式"    ' 定义标题。
    Help = "DEMO.HLP"    ' 定义帮助文件。
    Ctxt = 1000    ' 定义标题
            ' 上下文。
            ' 显示信息。
    Response = MsgBox(Msg, Style, title, Help, Ctxt)
End Sub



