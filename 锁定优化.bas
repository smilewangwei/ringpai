Sub ����()
    Dim x
    Dim y
    Dim password As Integer
    y = VBA.Timer
    password = Application.InputBox("��������������", "��������", 1101, , , , , 1)
    For x = Sheets.Count To 1 Step -1
        Application.ScreenUpdating = True '���������Ļ

        If Sheets(x).Name = "�ɱ���-08" Then
            Debug.Print Sheets(x).Name
            Sheets(Sheets(x).Name).Select '�������
            'ActionGroup1 (password)
            'Call CostSchedule
            ActionGroup2 Sheets(x).Name, password
        ElseIf Sheets(x).Name = "ҽ��-�Ĳ�-07" Or Sheets(x).Name = "��Ʒ-06" Or Sheets(x).Name = "����-05" Or Sheets(x).Name = "����-04" Then
            Debug.Print Sheets(x).Name
            'On Error Resume Next
            Sheets(Sheets(x).Name).Select '�������
            'ActionGroup1 (password)
            'Call money_delete
            Call Optimization()
            ActionGroup2 Sheets(x).Name, password
        ElseIf Sheets(x).Name = "����������-03" Then
            Debug.Print Sheets(x).Name
            '�ɱ�sheet��Ҫִ�еĶ���
            Sheets(Sheets(x).Name).Select '�������
            'ActionGroup1 (password)
            'Call ordermanagement
            ActionGroup2 Sheets(x).Name, password
        Else
            Debug.Print "������ִ�в���"
        End If
    Next
    Sheets("�ɱ���-08").Select
    VBA.Interaction.MsgBox("Time" & VBA.Timer - y & "s")
    Debug.Print("Time" & VBA.Timer - y & "s")
End Sub

'����������
Sub unlock_workbook()
    ActiveWorkbook.Unprotect password:=9987
End Sub

'����������
Sub locking_workbook(password As Integer)
    ActiveWorkbook.Protect password:=1102, Structure:=True, Windows:=False
End Sub

'�ɱ����������,08
Sub CostSchedule()
    '���ֲ���
    With Range("A40")
        .Value = "��������"
        .Font.Size = 14
    End With

    '��ʷ
    Range("A1:H8").Copy
    With Range("A41")
        .PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        .PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    End With

    '����
    Range("G2:G7").Copy
    Range("H2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

End Sub

'��ĩ������,04,05,06,07
Sub money_delete()
    Dim Title_MinRow
    Dim Title_MaxRow As Integer
    Dim Col1
    Dim Col2
    Dim Col3
    Dim Col4
    Dim Col5 As String
    Title_MinRow = ChearCol(2, "��Ʒ����") + 1
    Title_MaxRow = ActiveSheet.Range(ChearCol(1, "��Ʒ����") & "65536").End(xlUp).Row + 1
    Col1 = ChearCol(1, "��ĩ����")
    Col2 = ChearCol(1, "��ĩ����")

    '������ĩ����,��ĩ����
    Range(Col1 & Title_MinRow & ":" & Col2 & Title_MaxRow).Copy
    Col3 = ChearCol(1, "�ڳ�����")
    'ճ�������ڵ���
    Range(Col3 & Title_MinRow).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    '�̵�ʵ�����

    Col4 = ChearCol(1, "�̵�ʵ��")
    Range(Col4 & Title_MinRow & ":" & Col4 & Title_MaxRow).ClearContents
End Sub

'����������-03
Sub ordermanagement()
    Dim Title_MinRow
    Dim Title_MaxRow As Integer
    Dim Col1
    Dim Col2
    Dim Col3
    Dim Col4
    Dim Col5 As String

    Title_MinRow = ChearCol(2, "��Ʒ����") + 1
    Title_MaxRow = ActiveSheet.Range(ChearCol(1, "��Ʒ����") & "65536").End(xlUp).Row + 1
    Col1 = ChearCol(1, "�������")
    Col2 = ChearCol(1, "��Ʊ����")

    '�������
    Range(Col1 & Title_MinRow & ":" & Col2 & Title_MaxRow).ClearContents

    Col3 = ChearCol(1, "��������")
    '��������
    Range(Col3 & Title_MinRow & ":" & Col3 & Title_MaxRow).ClearContents

    Col4 = ChearCol(1, "�����")
    Range(Col4 & Title_MinRow).FormulaR1C1 = "=RC[-2]*RC[-1]"
    Range(Col4 & Title_MinRow).AutoFill Destination:=Range(Col4 & Title_MinRow & ":" & Col4 & Title_MaxRow)

End Sub

Sub ActionGroup1(password As Integer)
    'Sheets(Sheets(x).Name).Select '�������

    With ActiveSheet
        .Unprotect password:=password   '����sheet
        .Cells.EntireColumn.Hidden = False '��ʾ������
        .AutoFilterMode = False 'ȡ��ɸѡ
    End With
End Sub

Sub ActionGroup2(Name As String, password As Integer)
    Dim PositioningValue As String
    PositioningValue = "��Ʒ����"
    Range("A1").Select

    '�ɱ༭��Ԫ�����������ͳһ��ʽ
    ActiveSheet.AutoFilterMode = False 'ȡ��ɸѡ
    With ActiveWindow
        .FreezePanes = False 'ȡ������
        .Zoom = 100
    End With

    With Cells
        .Font.Name = "΢���ź�"
        .Font.Size = 10
        .RowHeight = 17
        .Locked = False
        .FormulaHidden = False
        .SpecialCells(xlCellTypeFormulas, 23).Locked = True

        If Name = "����������-03" Then
            .SpecialCells(xlCellTypeConstants, 7).Locked = False
            .Find(PositioningValue).EntireRow.AutoFilter
            Range(ChearCol(1, PositioningValue) & ChearCol(2, PositioningValue) + 1).Select
            ActiveWindow.FreezePanes = True
        ElseIf Name <> "�ɱ���-08" Then
            .SpecialCells(xlCellTypeConstants, 7).Locked = True
            .Find(PositioningValue).EntireRow.AutoFilter
            Range(ChearCol(1, PositioningValue) & ChearCol(2, PositioningValue) + 1).Select
            ActiveWindow.FreezePanes = True
        End If

        ActiveSheet.Protect password:=password, DrawingObjects:=False, Contents:=True, Scenarios:=False, AllowFormattingColumns:=True, AllowFormattingRows:=True, AllowFiltering:=True
    End With


End Sub

'��ȡ���Գ���
Function ChearCol(SplitNumber As Integer, Value As String)
    ChearCol = Split(Cells.Find(Value, LookIn:=xlValues, lookat:=xlPart).Address, "$")(SplitNumber)
End Function



Sub Optimization()
    Dim arr, arr2
    Dim SheetName As String
    Dim dic, dic2
    Set dic = CreateObject("Scripting.Dictionary")
    Set dic2 = CreateObject("Scripting.Dictionary")
    SheetName = "����������-03"

    arr2 = Array("��Ʒ����", "�������", "��ⵥ��", "�����", "��Ʊ���", "��Ʊ����")
    For i = LBound(arr2) To UBound(arr2)
        dic2.Add Split(Sheets(SheetName).Cells.Find(CStr(arr2(i)), LookIn:=xlValues, lookat:=xlPart).Address, "$")(1), arr2(i)
    Next
    DicKyes2 = dic2.Keys() 'ͨ������Set�󶨲���ֱ��ʹ��dic.keys(i)

    '    DicItem2 = dic2.Items()
    '    For i = LBound(arr2) To UBound(arr2)
    '        Debug.Print i & "---" & DicKyes2(i) & "---" & DicItem2(i)
    '    Next
    '0---F---��Ʒ����
    '1---O---�������
    '2---P---��ⵥ��
    '3---Q---�����
    '4---R---��Ʊ���
    '5---S---��Ʊ����

    arr = Array("���", "��Ʒ����", "��Ʒ����", "��Ʒ���", "�ڳ�����", "�ڳ�����", "�ڳ����", "��ⵥ��", "�������", "�ɹ����", "���ⵥ��", "��������", "��������", "���۳ɱ�", "�̵�����", "������", "��ĩ����", "��ĩ����", "��ĩ���", "�̵�ʵ��", "У��")
    For i = LBound(arr) To UBound(arr)
        dic.Add ChearCol(1, CStr(arr(i))), arr(i)
    Next
    DicKyes = dic.Keys() 'ͨ������Set�󶨲���ֱ��ʹ��dic.keys(i)
    DicItems = dic.Items() 'ͨ������Set�󶨲���ֱ��ʹ��dic.keys(i)

    '    'For i = LBound(arr) To UBound(arr)
    '    '    Debug.Print i & "---" & k(i) & "---" & it(i)
    '    'Next
    '    '0---A---���
    '    '1---B---��Ʒ����
    '    '2---C---��Ʒ����
    '    '3---D---��Ʒ���
    '    '4---E---�ڳ�����
    '    '5---F---�ڳ�����
    '    '6---G---�ڳ����
    '    '7---H---��ⵥ��
    '    '8---I---�������
    '    '9---J---�ɹ����
    '    '10---K---���ⵥ��
    '    '11---L---��������
    '    '12---M---��������
    '    '13---N---���۳ɱ�
    '    '14---O---�̵��������
    '    '15---P---������
    '    '16---Q---��ĩ����
    '    '17---R---��ĩ����
    '    '18---S---��ĩ���
    '    '19---T---�̵�ʵ��
    '    '20---U---У��
    '    '"&DicKyes(0)&"

    '��ʽУ��
    Dim Title_MinRow, Title_MaxRow As Integer
    Title_MinRow = ChearCol(2, CStr(DicItems(2))) + 1
    Title_MaxRow = ActiveSheet.Range(ChearCol(1, CStr(DicItems(2))) & "65536").End(xlUp).Row + 100
    Range(DicKyes(0) & Title_MinRow).Formula = "=row()"
    Range(DicKyes(6) & Title_MinRow).Formula = "=" & DicKyes(4) & Title_MinRow & "*" & DicKyes(5) & Title_MinRow
    Range(DicKyes(7) & Title_MinRow).Formula = "=SUMIF('" & SheetName & "'!" & DicKyes2(0) & ":" & DicKyes2(0) & "," & DicKyes(2) & Title_MinRow & ",'" & SheetName & "'!" & DicKyes2(1) & ":" & DicKyes2(1) & ")"
    Range(DicKyes(8) & Title_MinRow).Formula = "=SUMIF('" & SheetName & "'!" & DicKyes2(0) & ":" & DicKyes2(0) & "," & DicKyes(2) & Title_MinRow & ",'" & SheetName & "'!" & DicKyes2(2) & ":" & DicKyes2(2) & ")"
    Range(DicKyes(9) & Title_MinRow).Formula = "=SUMIF('" & SheetName & "'!" & DicKyes2(0) & ":" & DicKyes2(0) & "," & DicKyes(2) & Title_MinRow & ",'" & SheetName & "'!" & DicKyes2(3) & ":" & DicKyes2(3) & ")"
    Range(DicKyes(10) & Title_MinRow).Formula = "=IF((" & DicKyes(5) & Title_MinRow & "+" & DicKyes(8) & Title_MinRow & ")=0,0,(" & DicKyes(6) & Title_MinRow & "+" & DicKyes(9) & Title_MinRow & ")/(" & DicKyes(5) & Title_MinRow & "+" & DicKyes(8) & Title_MinRow & "))"
    Range(DicKyes(11) & Title_MinRow).Formula = ""
    Range(DicKyes(12) & Title_MinRow).Formula = "=" & DicKyes(11) & Title_MinRow
    Range(DicKyes(13) & Title_MinRow).Formula = "=" & DicKyes(10) & Title_MinRow & "*" & DicKyes(12) & Title_MinRow
    Range(DicKyes(14) & Title_MinRow).Formula = ""
    Range(DicKyes(15) & Title_MinRow).Formula = "=" & DicKyes(10) & Title_MinRow & "*" & DicKyes(14) & Title_MinRow
    Range(DicKyes(16) & Title_MinRow).Formula = "=IFERROR(" & DicKyes(18) & Title_MinRow & "/" & DicKyes(17) & Title_MinRow & ",0)"
    Range(DicKyes(17) & Title_MinRow).Formula = "=" & DicKyes(5) & Title_MinRow & "+" & DicKyes(8) & Title_MinRow & "-" & DicKyes(11) & Title_MinRow & "+" & DicKyes(14) & Title_MinRow
    Range(DicKyes(18) & Title_MinRow).Formula = "=" & DicKyes(10) & Title_MinRow & "*" & DicKyes(17) & Title_MinRow
    Range(DicKyes(19) & Title_MinRow).Formula = ""
    Range(DicKyes(20) & Title_MinRow).Formula = "=" & DicKyes(19) & Title_MinRow & "-" & DicKyes(17) & Title_MinRow


    Range(DicKyes(0) & Title_MinRow & ":" & DicKyes(0) & Title_MaxRow).FillDown
    Range(DicKyes(6) & Title_MinRow & ":" & DicKyes(10) & Title_MaxRow).FillDown
    '��������,�������,�̵�ʱ��,У��
    If False Then
        Range(DicKyes(11) & Title_MinRow & ":" & DicKyes(11) & Title_MaxRow).FillDown
        Range(DicKyes(14) & Title_MinRow & ":" & DicKyes(14) & Title_MaxRow).FillDown
        Range(DicKyes(19) & Title_MinRow & ":" & DicKyes(19) & Title_MaxRow).FillDown
    ElseIf True Then
        Range("" & DicKyes(11) & Title_MinRow & ":" & DicKyes(11) & Title_MaxRow & "," & DicKyes(14) & Title_MinRow & ":" & DicKyes(14) & Title_MaxRow & "," & DicKyes(19) & Title_MinRow & ":" & DicKyes(19) & Title_MaxRow & "").Clear
    End If
    Range(DicKyes(12) & Title_MinRow & ":" & DicKyes(13) & Title_MaxRow).FillDown
    Range(DicKyes(15) & Title_MinRow & ":" & DicKyes(18) & Title_MaxRow).FillDown
    Range(DicKyes(20) & Title_MinRow & ":" & DicKyes(20) & Title_MaxRow).FillDown

    '��ʽͳһ
    '    Range(DicKyes(0) & Title_MinRow & ":" & DicKyes(20) & Title_MaxRow).Select
    '    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    '    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    '
    '    With Selection.Borders(xlEdgeLeft)
    '        .LineStyle = xlContinuous
    '        .ColorIndex = xlAutomatic
    '        .TintAndShade = 0
    '        .Weight = xlHairline
    '    End With
    '
    '    With Selection.Borders(xlEdgeTop)
    '        .LineStyle = xlContinuous
    '        .ColorIndex = xlAutomatic
    '        .TintAndShade = 0
    '        .Weight = xlHairline
    '    End With
    '
    '    With Selection.Borders(xlEdgeBottom)
    '        .LineStyle = xlContinuous
    '        .ColorIndex = xlAutomatic
    '        .TintAndShade = 0
    '        .Weight = xlHairline
    '    End With
    '
    '    With Selection.Borders(xlEdgeRight)
    '        .LineStyle = xlContinuous
    '        .ColorIndex = xlAutomatic
    '        .TintAndShade = 0
    '        .Weight = xlHairline
    '    End With
    '
    '    With Selection.Borders(xlInsideVertical)
    '        .LineStyle = xlContinuous
    '        .ColorIndex = xlAutomatic
    '        .TintAndShade = 0
    '        .Weight = xlHairline
    '    End With
    '
    '    With Selection.Borders(xlInsideHorizontal)
    '        .LineStyle = xlContinuous
    '        .ColorIndex = xlAutomatic
    '        .TintAndShade = 0
    '        .Weight = xlHairline
    '    End With
    '
    '    With Selection.Font
    '        .Name = "΢���ź�"
    '        .FontStyle = "����"
    '        .Size = 10
    '        .Strikethrough = False
    '        .Superscript = False
    '        .Subscript = False
    '        .OutlineFont = False
    '        .Shadow = False
    '        .Underline = xlUnderlineStyleNone
    '        .ColorIndex = xlAutomatic
    '        .TintAndShade = 0
    '        .ThemeFont = xlThemeFontNone
    '    End With
    '
    '    With Selection.Interior
    '        .Pattern = xlSolid
    '        .PatternColorIndex = xlAutomatic
    '        .ThemeColor = xlThemeColorAccent4
    '        .TintAndShade = 0.799981688894314
    '        .PatternTintAndShade = 0
    '    End With

    With Range("" & DicKyes(11) & Title_MinRow & ":" & DicKyes(11) & Title_MaxRow & "," & DicKyes(14) & Title_MinRow & ":" & DicKyes(14) & Title_MaxRow & "," & DicKyes(19) & Title_MinRow & ":" & DicKyes(19) & Title_MaxRow & "").Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

End Sub