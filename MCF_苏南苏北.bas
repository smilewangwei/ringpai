'类别=
'说明=其他地区
Sub 苏南苏北()
    Dim x, y
    Dim password As Integer
    y = VBA.Timer
    password = 1102
'    Application.Calculation = xlCalculationManual
'    Application.ScreenUpdating = False
'    Application.DisplayStatusBar = False
'    Application.EnableEvents = False
    For x = Sheets.Count To 1 Step -1
        Application.ScreenUpdating = True '允许更新屏幕
        If Sheets(x).Name = "成本表-08" Then
            Debug.Print Sheets(x).Name
            Sheets(Sheets(x).Name).Select '活动工作表
            ActionGroup1 (password)
            Call CostSchedule
            ActionGroup2 Sheets(x).Name, password
        ElseIf Sheets(x).Name = "医疗-耗材-07" Or Sheets(x).Name = "用品-06" Or Sheets(x).Name = "美容-05" Or Sheets(x).Name = "诊疗-04" Then
            Debug.Print Sheets(x).Name
            'On Error Resume Next
            Sheets(Sheets(x).Name).Select '活动工作表
            ActionGroup1 (password)
            Call money_delete
            ActionGroup2 Sheets(x).Name, password

        ElseIf Sheets(x).Name = "订单入库管理-03" Then
            Debug.Print Sheets(x).Name
            '成本sheet需要执行的动作
            Sheets(Sheets(x).Name).Select '活动工作表
            ActionGroup1 (password)
            Call ordermanagement
            ActionGroup2 Sheets(x).Name, password
        Else
            Debug.Print "其他不执行操作"
        End If

    Next
'    Application.Calculation = xlCalculationManual
'    Application.ScreenUpdating = True
'    Application.DisplayStatusBar = True
'    Application.EnableEvents = True

    Sheets("成本表-08").Select
    VBA.Interaction.MsgBox ("Time" & VBA.Timer - y & "s")
    Debug.Print ("Time" & VBA.Timer - y & "s")
End Sub

'解锁工作簿
Sub unlock_workbook()
    ActiveWorkbook.Unprotect password:=9987
End Sub

'锁定工作簿
Sub locking_workbook(password As Integer)
    ActiveWorkbook.Protect password:=1102, Structure:=True, Windows:=False
End Sub

'成本表需操作项,08
Sub CostSchedule()
    '文字部分
    With Range("A40")
        .Value = "上月数据"
        .Font.Size = 14
    End With
    
    '历史
    Range("A1:H8").Copy
    
    With Range("A41")
        .PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        .PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    End With
    
    
    '返回
    Range("G2:G7").Copy
    Range("H2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    

End Sub

'期末金额调整,04,05,06,07
Sub money_delete()
    Dim Title_MinRow, Title_MaxRow As Integer
    Dim Col1, Col2, Col3, Col4, Col5, Col6 As String
    Title_MinRow = ChearCol(2, "产品名称") + 1
    Title_MaxRow = ActiveSheet.Range(ChearCol(1, "产品名称") & "65536").End(xlUp).Row + 1
    Col1 = ChearCol(1, "期末单价")
    Col2 = ChearCol(1, "期末金额")
    '复制期末单价,期末数量
    Range(Col1 & Title_MinRow & ":" & Col2 & Title_MaxRow).Copy
    Col3 = ChearCol(1, "期初单价")
    '粘贴到初期单价
    Range(Col3 & Title_MinRow).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    '盘点实存清空

    Col4 = ChearCol(1, "盘点实存")
    Range(Col4 & Title_MinRow & ":" & Col4 & Title_MaxRow).ClearContents
    Col5 = ChearCol(1, "出库数量")
    Range(Col5 & Title_MinRow & ":" & Col5 & Title_MaxRow).ClearContents
    Col6 = ChearCol(1, "盘点损益")
    Range(Col6 & Title_MinRow & ":" & Col6 & Title_MaxRow).ClearContents
End Sub


'订单入库管理-03
Sub ordermanagement()
    Dim Title_MinRow
    Dim Title_MaxRow As Integer
    Dim Col1
    Dim Col2
    Dim Col3
    Dim Col4
    Dim Col5 As String
    
    Title_MinRow = ChearCol(2, "产品名称") + 1
    Title_MaxRow = ActiveSheet.Range(ChearCol(1, "产品名称") & "65536").End(xlUp).Row + 1
    Col1 = ChearCol(1, "入库数量")
    Col2 = ChearCol(1, "开票日期")

    '入库数量
    Range(Col1 & Title_MinRow & ":" & Col2 & Title_MaxRow).ClearContents

    Col3 = ChearCol(1, "订单日期")
    '订单日期
    Range(Col3 & Title_MinRow & ":" & Col3 & Title_MaxRow).ClearContents

    Col4 = ChearCol(1, "入库金额")
    Range(Col4 & Title_MinRow & ":" & Col4 & Title_MaxRow).FormulaR1C1 = "=RC[-2]*RC[-1]"
End Sub

Sub ActionGroup1(password As Integer)
    'Sheets(Sheets(x).Name).Select '活动工作表
    With ActiveSheet
        .Unprotect password:=password   '解锁sheet
        .Cells.EntireColumn.Hidden = False '显示隐藏列
        .AutoFilterMode = False '取消筛选
    End With
    With ActiveWindow
    .FreezePanes = False '取消冻结
    .Zoom = 100
    End With
End Sub






Sub ActionGroup2(Name As String, password As Integer)
    Dim PositioningValue As String
    PositioningValue = "产品名称"
    Range("A1").Select

    '可编辑单元格对象锁定并统一格式

    With Cells
        .Font.Name = "微软雅黑"
        .Font.Size = 10
        .RowHeight = 17
        .Locked = False
        .FormulaHidden = False
        .SpecialCells(xlCellTypeFormulas, 23).Locked = True

        If Name = "订单入库管理-03" Then
            .SpecialCells(xlCellTypeConstants, 7).Locked = False
            .Find(PositioningValue).EntireRow.AutoFilter
            Range(ChearCol(1, PositioningValue) & ChearCol(2, PositioningValue) + 1).Select
            ActiveWindow.FreezePanes = True
        ElseIf Name <> "成本表-08" Then
            .SpecialCells(xlCellTypeConstants, 7).Locked = True
            .Find(PositioningValue).EntireRow.AutoFilter
            Range(ChearCol(1, PositioningValue) & ChearCol(2, PositioningValue) + 1).Select
            ActiveWindow.FreezePanes = True
        End If

        ActiveSheet.Protect password:=1102, DrawingObjects:=False, Contents:=True, Scenarios:=False, AllowFormattingColumns:=True, AllowFormattingRows:=True, AllowFiltering:=True

        '文字部分对象锁定
    End With

End Sub



'获取不对称列
Function ChearCol(SplitNumber As Integer, Value As String)
ChearCol = Split(Cells.Find(Value).Address, "$")(SplitNumber)
End Function
