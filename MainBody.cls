VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MainBody"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'正文文本
'段落：
'正文内容：宋体小4号，行距固定值20磅；英文用Times New Roman，小4号
'注意，字体并没有设置，因为涉及到上下标的问题


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


    If Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly And _
        Selection.ParagraphFormat.LineSpacing = 20 And _
        Selection.ParagraphFormat.LineUnitBefore = 0 And _
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
    
    If Selection.Font.NameFarEast = "宋体" And _
        Selection.Font.NameAscii = "Times New Roman" And _
        Selection.Font.NameOther = "Times New Roman" And _
        Selection.Font.Name = "" And _
        Selection.Font.Bold = False Then
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

    'Selection.ClearFormatting


    '这是用录制宏命令的代码
    With Selection.Font
        .NameFarEast = "宋体"
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
       ' .Size = 12  注意这里没有设字体为小4，因为涉及到上下标的问题
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
        .Color = wdColorAutomatic
        .Engrave = False
        .Spacing = 0
        .Scaling = 100
        .Position = 0
        .Kerning = 1
        .Animation = wdAnimationNone
        .DisableCharacterSpaceGrid = False
        .EmphasisMark = wdEmphasisMarkNone
        .Ligatures = wdLigaturesNone
        .NumberSpacing = wdNumberSpacingDefault
        .NumberForm = wdNumberFormDefault
        .StylisticSet = wdStylisticSetDefault
        .ContextualAlternates = 0
    End With
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = 20
        .Alignment = wdAlignParagraphJustify
        .WidowControl = False
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = CentimetersToPoints(0.35)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 2
        .LineUnitBefore = 0
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
