'���=
'˵��=��˵��
Sub ���˽ű����ɻ���()
 With Cells
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("B4:C5").FillDown
    Dim n As Integer
    Dim TableName, StartCol As String

    n = [C65536].End(xlUp).Row
    TableName = "����͸�ӱ�1"
    StartCol = "B"
    
    For i = 5 To n
        If Range("B" & i).Value = "" Then
            Range("B" & i).Value = Range("B" & i - 1).Value
        End If
        Range("D" & i).Value = Range("q" & i).Value + Range("aa" & i).Value
    Next
    
    
    Range(StartCol & "5:D" & n).Select
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        StartCol & "5:D" & n, Version:=6).CreatePivotTable TableDestination:=Range(StartCol & n + 5), TableName:=TableName, DefaultVersion:=6
    Range(StartCol & n + 10).Select
    With ActiveSheet.PivotTables(TableName).PivotFields("��������")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables(TableName).PivotFields("����")
        .Orientation = xlColumnField
        .Position = 1
    End With

    ActiveSheet.PivotTables(TableName).AddDataField ActiveSheet.PivotTables(TableName).PivotFields("ȫ������������"), "�����:ȫ������������", xlSum
    
    With ActiveSheet.PivotTables(TableName).PivotFields("�����:ȫ������������")
        .NumberFormat = "0_ "
    End With
    Range(StartCol & n + 10).Select
End Sub
