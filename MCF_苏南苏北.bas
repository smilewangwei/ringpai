'���=
'˵��=��������
Sub �����ձ�()
    Dim x, y
    Dim password As Integer
    y = VBA.Timer
    password = 1102
'    Application.Calculation = xlCalculationManual
'    Application.ScreenUpdating = False
'    Application.DisplayStatusBar = False
'    Application.EnableEvents = False
    For x = Sheets.Count To 1 Step -1
        Application.ScreenUpdating = True '���������Ļ
        If Sheets(x).Name = "�ɱ���-08" Then
            Debug.Print Sheets(x).Name
            Sheets(Sheets(x).Name).Select '�������
            ActionGroup1 (password)
            Call CostSchedule
            ActionGroup2 Sheets(x).Name, password
        ElseIf Sheets(x).Name = "ҽ��-�Ĳ�-07" Or Sheets(x).Name = "��Ʒ-06" Or Sheets(x).Name = "����-05" Or Sheets(x).Name = "����-04" Then
            Debug.Print Sheets(x).Name
            'On Error Resume Next
            Sheets(Sheets(x).Name).Select '�������
            ActionGroup1 (password)
            Call money_delete
            ActionGroup2 Sheets(x).Name, password

        ElseIf Sheets(x).Name = "����������-03" Then
            Debug.Print Sheets(x).Name
            '�ɱ�sheet��Ҫִ�еĶ���
            Sheets(Sheets(x).Name).Select '�������
            ActionGroup1 (password)
            Call ordermanagement
            ActionGroup2 Sheets(x).Name, password
        Else
            Debug.Print "������ִ�в���"
        End If

    Next
'    Application.Calculation = xlCalculationManual
'    Application.ScreenUpdating = True
'    Application.DisplayStatusBar = True
'    Application.EnableEvents = True

    Sheets("�ɱ���-08").Select
    VBA.Interaction.MsgBox ("Time" & VBA.Timer - y & "s")
    Debug.Print ("Time" & VBA.Timer - y & "s")
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
    Dim Title_MinRow, Title_MaxRow As Integer
    Dim Col1, Col2, Col3, Col4, Col5, Col6 As String
    Title_MinRow = ChearCol(2, "��Ʒ����") + 1
    Title_MaxRow = ActiveSheet.Range(ChearCol(1, "��Ʒ����") & "65536").End(xlUp).Row + 1
    Col1 = ChearCol(1, "��ĩ����")
    Col2 = ChearCol(1, "��ĩ���")
    '������ĩ����,��ĩ����
    Range(Col1 & Title_MinRow & ":" & Col2 & Title_MaxRow).Copy
    Col3 = ChearCol(1, "�ڳ�����")
    'ճ�������ڵ���
    Range(Col3 & Title_MinRow).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    '�̵�ʵ�����

    Col4 = ChearCol(1, "�̵�ʵ��")
    Range(Col4 & Title_MinRow & ":" & Col4 & Title_MaxRow).ClearContents
    Col5 = ChearCol(1, "��������")
    Range(Col5 & Title_MinRow & ":" & Col5 & Title_MaxRow).ClearContents
    Col6 = ChearCol(1, "�̵�����")
    Range(Col6 & Title_MinRow & ":" & Col6 & Title_MaxRow).ClearContents
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
    Range(Col4 & Title_MinRow & ":" & Col4 & Title_MaxRow).FormulaR1C1 = "=RC[-2]*RC[-1]"
End Sub

Sub ActionGroup1(password As Integer)
    'Sheets(Sheets(x).Name).Select '�������
    With ActiveSheet
        .Unprotect password:=password   '����sheet
        .Cells.EntireColumn.Hidden = False '��ʾ������
        .AutoFilterMode = False 'ȡ��ɸѡ
    End With
    With ActiveWindow
    .FreezePanes = False 'ȡ������
    .Zoom = 100
    End With
End Sub






Sub ActionGroup2(Name As String, password As Integer)
    Dim PositioningValue As String
    PositioningValue = "��Ʒ����"
    Range("A1").Select

    '�ɱ༭��Ԫ�����������ͳһ��ʽ

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

        ActiveSheet.Protect password:=1102, DrawingObjects:=False, Contents:=True, Scenarios:=False, AllowFormattingColumns:=True, AllowFormattingRows:=True, AllowFiltering:=True

        '���ֲ��ֶ�������
    End With

End Sub



'��ȡ���Գ���
Function ChearCol(SplitNumber As Integer, Value As String)
ChearCol = Split(Cells.Find(Value).Address, "$")(SplitNumber)
End Function
