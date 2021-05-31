Sub 锁定()
    Dim x
    Dim y
    Dim password As Integer
    y = VBA.Timer
    password = Application.InputBox("请输入锁定密码", "输入密码", 1101, , , , , 1)
    For x = Sheets.Count To 1 Step -1
        Application.ScreenUpdating = True '允许更新屏幕

        If Sheets(x).Name = "成本表-08" Then
            Debug.Print Sheets(x).Name
            Sheets(Sheets(x).Name).Select '活动工作表
            'ActionGroup1 (password)
            'Call CostSchedule
            ActionGroup2 Sheets(x).Name, password
        ElseIf Sheets(x).Name = "医疗-耗材-07" Or Sheets(x).Name = "用品-06" Or Sheets(x).Name = "美容-05" Or Sheets(x).Name = "诊疗-04" Then
            Debug.Print Sheets(x).Name
            'On Error Resume Next
            Sheets(Sheets(x).Name).Select '活动工作表
            'ActionGroup1 (password)
            'Call money_delete
            Call Optimization()
            ActionGroup2 Sheets(x).Name, password
        ElseIf Sheets(x).Name = "订单入库管理-03" Then
            Debug.Print Sheets(x).Name
            '成本sheet需要执行的动作
            Sheets(Sheets(x).Name).Select '活动工作表
            'ActionGroup1 (password)
            'Call ordermanagement
            ActionGroup2 Sheets(x).Name, password
        Else
            Debug.Print "其他不执行操作"
        End If
    Next
    Sheets("成本表-08").Select
    VBA.Interaction.MsgBox("Time" & VBA.Timer - y & "s")
    Debug.Print("Time" & VBA.Timer - y & "s")
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
    Dim Title_MinRow
    Dim Title_MaxRow As Integer
    Dim Col1
    Dim Col2
    Dim Col3
    Dim Col4
    Dim Col5 As String
    Title_MinRow = ChearCol(2, "产品名称") + 1
    Title_MaxRow = ActiveSheet.Range(ChearCol(1, "产品名称") & "65536").End(xlUp).Row + 1
    Col1 = ChearCol(1, "期末单价")
    Col2 = ChearCol(1, "期末数量")

    '复制期末单价,期末数量
    Range(Col1 & Title_MinRow & ":" & Col2 & Title_MaxRow).Copy
    Col3 = ChearCol(1, "期初单价")
    '粘贴到初期单价
    Range(Col3 & Title_MinRow).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    '盘点实存清空

    Col4 = ChearCol(1, "盘点实存")
    Range(Col4 & Title_MinRow & ":" & Col4 & Title_MaxRow).ClearContents
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
    Range(Col4 & Title_MinRow).FormulaR1C1 = "=RC[-2]*RC[-1]"
    Range(Col4 & Title_MinRow).AutoFill Destination:=Range(Col4 & Title_MinRow & ":" & Col4 & Title_MaxRow)

End Sub

Sub ActionGroup1(password As Integer)
    'Sheets(Sheets(x).Name).Select '活动工作表

    With ActiveSheet
        .Unprotect password:=password   '解锁sheet
        .Cells.EntireColumn.Hidden = False '显示隐藏列
        .AutoFilterMode = False '取消筛选
    End With
End Sub

Sub ActionGroup2(Name As String, password As Integer)
    Dim PositioningValue As String
    PositioningValue = "产品名称"
    Range("A1").Select

    '可编辑单元格对象锁定并统一格式
    ActiveSheet.AutoFilterMode = False '取消筛选
    With ActiveWindow
        .FreezePanes = False '取消冻结
        .Zoom = 100
    End With

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

        ActiveSheet.Protect password:=password, DrawingObjects:=False, Contents:=True, Scenarios:=False, AllowFormattingColumns:=True, AllowFormattingRows:=True, AllowFiltering:=True
    End With


End Sub

'获取不对称列
Function ChearCol(SplitNumber As Integer, Value As String)
    ChearCol = Split(Cells.Find(Value, LookIn:=xlValues, lookat:=xlPart).Address, "$")(SplitNumber)
End Function



Sub Optimization()
    Dim arr, arr2
    Dim SheetName As String
    Dim dic, dic2
    Set dic = CreateObject("Scripting.Dictionary")
    Set dic2 = CreateObject("Scripting.Dictionary")
    SheetName = "订单入库管理-03"

    arr2 = Array("产品名称", "入库数量", "入库单价", "入库金额", "开票金额", "开票日期")
    For i = LBound(arr2) To UBound(arr2)
        dic2.Add Split(Sheets(SheetName).Cells.Find(CStr(arr2(i)), LookIn:=xlValues, lookat:=xlPart).Address, "$")(1), arr2(i)
    Next
    DicKyes2 = dic2.Keys() '通过后期Set绑定不可直接使用dic.keys(i)

    '    DicItem2 = dic2.Items()
    '    For i = LBound(arr2) To UBound(arr2)
    '        Debug.Print i & "---" & DicKyes2(i) & "---" & DicItem2(i)
    '    Next
    '0---F---产品名称
    '1---O---入库数量
    '2---P---入库单价
    '3---Q---入库金额
    '4---R---开票金额
    '5---S---开票日期

    arr = Array("序号", "产品分类", "产品名称", "产品规格", "期初单价", "期初数量", "期初金额", "入库单价", "入库数量", "采购金额", "出库单价", "出库数量", "销售数量", "销售成本", "盘点损益", "损益金额", "期末单价", "期末数量", "期末金额", "盘点实存", "校验")
    For i = LBound(arr) To UBound(arr)
        dic.Add ChearCol(1, CStr(arr(i))), arr(i)
    Next
    DicKyes = dic.Keys() '通过后期Set绑定不可直接使用dic.keys(i)
    DicItems = dic.Items() '通过后期Set绑定不可直接使用dic.keys(i)

    '    'For i = LBound(arr) To UBound(arr)
    '    '    Debug.Print i & "---" & k(i) & "---" & it(i)
    '    'Next
    '    '0---A---序号
    '    '1---B---产品分类
    '    '2---C---产品名称
    '    '3---D---产品规格
    '    '4---E---期初单价
    '    '5---F---期初数量
    '    '6---G---期初金额
    '    '7---H---入库单价
    '    '8---I---入库数量
    '    '9---J---采购金额
    '    '10---K---出库单价
    '    '11---L---出库数量
    '    '12---M---销售数量
    '    '13---N---销售成本
    '    '14---O---盘点损益调整
    '    '15---P---损益金额
    '    '16---Q---期末单价
    '    '17---R---期末数量
    '    '18---S---期末金额
    '    '19---T---盘点实存
    '    '20---U---校验
    '    '"&DicKyes(0)&"

    '公式校对
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
    '出库数量,损益调整,盘点时存,校验
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

    '格式统一
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
    '        .Name = "微软雅黑"
    '        .FontStyle = "常规"
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