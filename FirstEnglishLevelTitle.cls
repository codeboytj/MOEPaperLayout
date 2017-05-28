VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FirstEnglishLevelTitle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'һ��Ӣ�ı���
'���䣺
'wdLineSpaceSingle�����о�
'LineUnitBefore��ǰ0�о�
'LineUnitAfter�κ�0.5�о�

'���壺
'Ӣ������ΪTimes New Roman
'С2(18)

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

    If Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceSingle And _
        Selection.ParagraphFormat.LineUnitBefore = 0 And _
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
        Selection.Font.Size = 18 And _
        Selection.Font.Bold = True Then
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

    '������¼�ƺ�����Ĵ���
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 2.5
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
        .LineUnitBefore = 0
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
        .NameFarEast = "����" '������������Ϊ����
        .NameAscii = "Times New Roman" '����Ӣ������ΪTimes New Roman
        .Size = 18 'С2
        .Bold = True '�Ӵ�
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

