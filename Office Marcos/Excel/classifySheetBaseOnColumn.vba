Sub 筛选拆分表格()
    Dim d As Object, sht As Worksheet, arr, brr, r, kr, i&, j&, k&, x&
    Dim Rng As Range, Rg As Range, tRow&, tCol&, aCol&, pd&
    Application.ScreenUpdating = False '关闭屏幕更新
    Application.DisplayAlerts = False '关闭警告信息提示
    Set d = CreateObject("scripting.dictionary") 'set字典
    Set Rg = Application.InputBox("请框选拆分依据列！只能选择单列单元格区域！", Title:="提示", Type:=8)
    '用户选择的拆分依据列
    tCol = Rg.Column '取拆分依据列列标
    tRow = Val(Application.InputBox("请输入总表标题行的行数？"))
    '用户设置总表的标题行数
    If tRow = 0 Then MsgBox "你未输入标题行行数，程序退出。": Exit Sub
    Set Rng = ActiveSheet.UsedRange '总表的数据区域
    arr = Rng '数据范围装入数组arr
    tCol = tCol - Rng.Column + 1 '计算依据列在数组中的位置
    aCol = UBound(arr, 2) '数据源的列数
    For i = tRow + 1 To UBound(arr) '遍历数组arr
        If Not d.exists(arr(i, tCol)) Then
            d(arr(i, tCol)) = i '字典中不存在关键词则将行号装入字典
        Else
            d(arr(i, tCol)) = d(arr(i, tCol)) & "," & i '如果存在则合并行号，以逗号间隔
        End If
    Next
    For Each sht In Worksheets '遍历一遍工作表，如果字典中存在则删除
        If d.exists(sht.Name) Then sht.Delete
    Next
    kr = d.keys '字典的key集
    For i = 0 To UBound(kr) '遍历字典key值
        If kr(i) <> "" Then '如果key不为空
            r = Split(d(kr(i)), ",") '取出item里储存的行号
            ReDim brr(1 To UBound(r) + 1, 1 To aCol) '声明放置结果的数组brr
            k = 0
            For x = 0 To UBound(r)
                k = k + 1 '累加记录行数
                For j = 1 To aCol '循环读取列
                    brr(k, j) = arr(r(x), j)
                Next
            Next
            With Worksheets.Add(, Sheets(Sheets.Count))
            '新建一个工作表，位置在所有已存在sheet的后面
                .Name = kr(i) '表格命名
                .[a1].Resize(tRow, aCol) = arr '放标题行
                .[a1].Offset(tRow, 0).Resize(k, aCol) = brr '放置数据区域
                Rng.Copy '复制粘贴总表的格式
                .[a1].PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                .[a1].Select
            End With
            Worksheets(kr(i)).Cells.EntireColumn.AutoFit '自动调节表的各列宽度
        End If
    Next
    Sheets(1).Activate '激活第一个表格
    Set d = Nothing '释放字典
    Erase arr: Erase brr '释放数组
    MsgBox "数据拆分完成！"
    Application.ScreenUpdating = True '恢复屏幕更新
    Application.DisplayAlerts = True '恢复警示
End Sub
