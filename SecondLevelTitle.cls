VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SecondLevelTitle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'��������
'���䣺
'wdLineSpaceExactly�о�̶�ֵ��20��
'LineUnitBefore��ǰ0.5�о�
'LineUnitAfter�κ�0.5�о�

'���壺
'��������Ϊ����
'Ӣ������ΪTimes New Roman
'С3(15)

Public Sub changeFormat()

    setFormat
    
    If checkFormat() = False Then
        '�����ʽ����
        setFormat
    End If

End Sub
Public Function checkFormat() As Boolean
    
    '����ʽ�Ƿ���ȷ
    checkFormat = CheckFont() And CheckParagraphFormat()

End Function


Private Function CheckParagraphFormat() As Boolean


    If Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly And _
        Selection.ParagraphFormat.LineSpacing = 20 And _
        Selection.ParagraphFormat.LineUnitBefore = 0.5 And _
        Selection.ParagraphFormat.LineUnitAfter = 0.5 Then
        '��ʽ��ȷ������true
        CheckParagraphFormat = True
    Else
        '��ʽ���ԣ���ӡ��ʾ
        popMsg ("�������")
        CheckParagraphFormat = False
    End If
    
    Exit Function
    
End Function

Private Function CheckFont() As Boolean
    '��������Ƿ���ȷ
    
    If Selection.Font.NameFarEast = "����" And _
        Selection.Font.NameAscii = "Times New Roman" And _
        Selection.Font.NameOther = "Times New Roman" And _
        Selection.Font.Name = "" And _
        Selection.Font.Size = 15 And _
        Selection.Font.Bold = False Then
        '��ʽ��ȷ������true
        CheckFont = True
    Else
        '��ʽ���ԣ���ӡ��ʾ
        popMsg ("�������")
        CheckFont = False
    End If

    Exit Function
    
End Function


Private Sub setFormat()

    Selection.ClearFormatting


    '������¼�ƺ�����Ĵ���
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 2.5
        .SpaceBeforeAuto = False
        .SpaceAfter = 2.5
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
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevel2
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0.5
        .LineUnitAfter = 0.5
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
        .NameFarEast = "����"
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .Name = ""
        .Size = 15
        .Bold = False
'        .BoldBi = False
'        .Italic = False
'        .Underline = wdUnderlineNone
'        .UnderlineColor = wdColorAutomatic
'        .StrikeThrough = False
'        .DoubleStrikeThrough = False
'        .Outline = False
'        .Emboss = False
'        .Shadow = False
'        .Hidden = False
'        .SmallCaps = False
'        .AllCaps = False
'        .Color = wdColorAutomatic
'        .Engrave = False
'        .Superscript = False
'        .Subscript = False
'        .Spacing = 0
'        .Scaling = 100
'        .Position = 0
'        .Kerning = 1
'        .Animation = wdAnimationNone
'        .DisableCharacterSpaceGrid = False
'        .EmphasisMark = wdEmphasisMarkNone
'        .Ligatures = wdLigaturesNone
'        .NumberSpacing = wdNumberSpacingDefault
'        .NumberForm = wdNumberFormDefault
'        .StylisticSet = wdStylisticSetDefault
'        .ContextualAlternates = 0
    End With

End Sub

Private Sub popMsg(Msg As String)

    Dim Style, title, Help, Ctxt, Response, MyString
    'Msg = "һ�������ʽ���� ?"    ' ������Ϣ��
    Style = vbYesNo + vbCritical + vbDefaultButton2    ' ���尴ť��
    title = "���һ�������ʽ"    ' ������⡣
    Help = "DEMO.HLP"    ' ��������ļ���
    Ctxt = 1000    ' �������
            ' �����ġ�
            ' ��ʾ��Ϣ��
    Response = MsgBox(Msg, Style, title, Help, Ctxt)
End Sub



