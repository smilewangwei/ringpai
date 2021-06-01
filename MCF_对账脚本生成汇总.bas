'类别=
'说明=无说明
Sub 对账脚本生成汇总()
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
    TableName = "数据透视表1"
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
    With ActiveSheet.PivotTables(TableName).PivotFields("店面名称")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables(TableName).PivotFields("日期")
        .Orientation = xlColumnField
        .Position = 1
    End With

    ActiveSheet.PivotTables(TableName).AddDataField ActiveSheet.PivotTables(TableName).PivotFields("全部馈赠额消费"), "求和项:全部馈赠额消费", xlSum
    
    With ActiveSheet.PivotTables(TableName).PivotFields("求和项:全部馈赠额消费")
        .NumberFormat = "0_ "
    End With
    Range(StartCol & n + 10).Select
End Sub
