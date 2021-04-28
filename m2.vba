Option Explicit

Dim txtlpdm As String ' 楼P代码
Dim txtgcmc As String ' 工程名称
Dim txtqyjx As String ' 区域简写
Dim txtjhdh As String ' 计划单号
Dim txtbmcl As String ' 表面处理
Dim txtxdsj As String ' 下单时间
Dim txtgyxm As String ' 工艺姓名
Dim txtgydh As String ' 工艺电话
Dim txtshxm As String ' 审核姓名
Dim txtscch As String ' 生产厂号

' TODO 输入检查
Sub init()
    txtlpdm = "楼P代码"
    txtgcmc = "工程名称"
    txtqyjx = "区域简写"
    txtjhdh = "计划单号"
    txtbmcl = "表面处理"
    txtxdsj = "下单时间"
    txtgyxm = "工艺姓名"
    txtgydh = "工艺电话"
    txtshxm = "审核姓名"
    txtscch = "生产厂号"
End Sub

' TODO，上一步做完，将fpqdFilename罗列在工作表里，等待处理
Sub testFB1()
    Dim fpqdFilename As String
    Dim scqdFilename As String
    fpqdFilename = "C:\Users\u03013112\Documents\J\new-412-1\分配清单\C-494.xlsx"
    scqdFilename = "C:\Users\u03013112\Documents\J\new-412-1\分配清单\" & txtlpdm & txtgcmc & txtqyjx & "-生产单.xlsx"
    Call FB1(fpqdFilename, scqdFilename)
End Sub
Sub testFB2()
    Dim scqdFilename As String
    scqdFilename = "C:\Users\u03013112\Documents\J\new-412-1\分配清单\" & txtlpdm & txtgcmc & txtqyjx & "-生产单.xlsx"
    Call FB2(scqdFilename)
End Sub
Sub testFB3()
    Dim scqdFilename As String
    scqdFilename = "C:\Users\u03013112\Documents\J\new-412-1\分配清单\" & txtlpdm & txtgcmc & txtqyjx & "-生产单.xlsx"
    Call FB3(scqdFilename)
End Sub

' 实际不再需要手选目标了，但是简单处理还是先分开
' fpqdFilename 配清单中的零件图，每一个都需要处理
Sub FB1(fpqdFilename As String, scqdFilename As String)
    If fileIsExist(scqdFilename) Then
        Kill scqdFilename
    End If
    Call createExcel(scqdFilename)

    Dim wb As Workbook
    Set wb = Workbooks.Open(scqdFilename)
    wb.Windows(1).Visible = False
    ThisWorkbook.Activate

    ' 新建一个sheet以供操作
    If isSheetExist(wb, "临时") Then
        wb.Sheets("临时").Delete
    End If
    wb.Sheets.Add().Name = "临时"
    ' Call copySheet(ThisWorkbook.Sheets("模板生产计划单"), wb.Sheets("临时"))
    ThisWorkbook.Sheets("模板生产计划单").Cells.Copy wb.Sheets("临时").[A1]

    ' ZXD拷贝过去
    ' TODO: 这些数据都应该是追加的，现在全都是替换，之后看看是生成多个文件，还是汇总起来
    If isSheetExist(wb, "ZXD") Then
        wb.Sheets("ZXD").Delete
    End If
    wb.Sheets.Add().Name = "ZXD"
    ' Call copySheet(ThisWorkbook.Sheets("模板生产计划单"), wb.Sheets("ZXD"))
    ThisWorkbook.Sheets("ZXD").Cells.Copy wb.Sheets("ZXD").[A1]

    ' 重新起名字，为了后面代码和他保持一致
    Dim qysr, txtsdsj, gydh, gcmc, gyxm, shxm, jhdh, ptfs, scch
    qysr = txtqyjx
    txtsdsj = Date
    gydh = txtgydh
    gcmc = txtgcmc
    gyxm = txtgyxm
    shxm = txtshxm
    jhdh = txtjhdh
    ptfs = txtbmcl
    scch = txtscch

    With wb.Sheets("临时")
        .Range("B3") = txtlpdm
        .Range("B2") = txtgcmc
        .Range("G2") = qysr
        .Range("I2") = txtsdsj
        .Range("G3") = txtjhdh
        .Range("K3") = txtbmcl
        If Len(.Range("B5")) = 0 Then
                .Range("B5") = txtgyxm
        End If
        If Len(.Range("F5")) = 0 Then
                .Range("F5") = txtshxm
        End If
    End With

    wb.Sheets("ZXD").Range("B2") = gcmc & qysr
    wb.Sheets("ZXD").Range("A1") = "模板转序记录表 (" & ptfs & ")"
    Dim wbTmp As Workbook
    Set wbTmp = Workbooks.Open(fpqdFilename)
    If isSheetExist(wb, "erp") Then
        wb.Sheets("erp").Delete
    End If
    wb.Sheets.Add().Name = "erp"
    wbTmp.Sheets(1).Cells.Copy wb.Sheets("erp").[A1]
    wbTmp.Close False

    Dim endd, Slhj
    If txtqyjx <> "BZJ" And txtqyjx <> "bzj" Then
        If wb.Sheets("erp").Range("A1") = "序号" Then wb.Sheets("erp").Rows("1:1").Delete
        endd = wb.Sheets("erp").[d65536].End(xlUp).Row
        wb.Sheets("erp").Range("j1:j" & endd) = qysr
        wb.Sheets("erp").Range("A1:J" & endd).Interior.Pattern = xlNone
        wb.Sheets("erp").Range("A1:J" & endd).Borders.Weight = 2

        If Left(qysr, 2) = "TP" Or Left(qysr, 3) = " TP" Then
            wb.Sheets("erp").Range("k1:k" & endd) = "带配件"
            wb.Sheets("erp").Range("k1:k" & endd).Interior.Pattern = xlNone
            wb.Sheets("erp").Range("k1:k" & endd).Borders.Weight = 2
        End If
        
        wb.Sheets("erp").Columns("C:C").EntireColumn.AutoFit '调整模板编号列的列宽
        Slhj = Application.WorksheetFunction.Sum(wb.Sheets("erp").Range("D1:D" & endd)) '数量合计
        wb.Sheets("erp").Columns("A:k").FormatConditions.Delete '清空条件格式
        wb.Sheets("erp").Columns("A:J").HorizontalAlignment = xlCenter '水平方向居中
        Application.ScreenUpdating = True
        wb.Windows(1).Visible = True
        wb.Close (True)
        MsgBox "总数量： " & Slhj & " 件"
    End If
End Sub
' 合并同类模板并分类
Sub FB2(scqdFilename As String)
    Dim wb As Workbook
    Set wb = Workbooks.Open(scqdFilename)
    'wb.Windows(1).Visible = False
    ThisWorkbook.Activate

    'Application.ScreenUpdating = False
    On Error Resume Next
    wb.Sheets("erp").Columns("C:C").Replace "（", "("
    wb.Sheets("erp").Columns("C:C").Replace "）", ")"
    wb.Sheets("erp").Columns("C:F").Replace " ", ""
    '单元格匹配替换W2的0值
    wb.Sheets("erp").Columns("F:F").Replace What:="0", Replacement:="", LookAt:=xlWhole, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    wb.Sheets("erp").Range("A1") = "1"
    Dim endb
    endb = wb.Sheets("erp").[b60000].End(xlUp).Row
    wb.Sheets("erp").Range("A1").AutoFill Destination:=wb.Sheets("erp").Range("A1:A" & endb), Type:=xlFillSeries
   
    wb.Sheets("erp").Columns("H:H").NumberFormatLocal = "G/通用格式"
    Dim i
    For i = 1 To endb
        If InStr(wb.Sheets("erp").Range("H" & i), "-N-") > 0 Then wb.Sheets("erp").Range("B" & i) = "转角"
        If InStr(wb.Sheets("erp").Range("H" & i), "ZW-Q-") > 0 Then wb.Sheets("erp").Range("B" & i) = "墙柱C槽"
    Next
    '对现有内容进行排序
    With wb.Sheets("erp").Sort.SortFields
            .Clear
        .Add Key:=.Range("B2"), Order:=1   '模板名称
        .Add Key:=.Range("E2"), Order:=1   'W1
        .Add Key:=.Range("F2"), Order:=1   'W2
        .Add Key:=.Range("H2"), Order:=1   '图纸编号
        .Add Key:=.Range("G2"), Order:=1   'L
    End With
    With wb.Sheets("erp").Sheets("erp").Sort
        .SetRange Range("b2:M" & endb)
        .Header = 2 '没有标题
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    For i = 1 To endb
        wb.Sheets("erp").Range("L" & i) = wb.Sheets("erp").Range("G" & i) & ""  '做宽度连接
        wb.Sheets("erp").Range("K" & i) = wb.Sheets("erp").Range("E" & i) & wb.Sheets("erp").Range("F" & i)  '做宽度连接
        wb.Sheets("erp").Range("M" & i) = wb.Sheets("erp").Range("B" & i) & "&" & wb.Sheets("erp").Range("E" & i) & wb.Sheets("erp").Range("F" & i) & wb.Sheets("erp").Range("J" & i)
    Next
    '宽度连接为数值
    wb.Sheets("erp").Range("E:G").Delete Shift:=xlToLeft '删除以前的两列宽度及单件面积和总面积列
    wb.Sheets("erp").Columns("H:I").Cut
    wb.Sheets("erp").Columns("E:E").Insert Shift:=xlToRight
    ' wb.Sheets("erp").Columns("A:J").Select

    wb.Sheets("erp").[k1] = 1
    wb.Sheets("erp").Range("K2").FormulaR1C1 = "=IF(RC[-1]<>R[-1]C[-1],R[-1]C+1,R[-1]C)"
    wb.Sheets("erp").Range("K2").AutoFill Destination:=wb.Sheets("erp").Range("K2:K" & endb)
    Dim zhz
    zhz = wb.Sheets("erp").Range("K" & endb).Value
    
    wb.Sheets("erp").Range("K1:K" & endb) = wb.Sheets("erp").Range("K1:K" & endb).Value
    Dim jia
    For jia = 1 To zhz
        wb.Sheets("erp").Range("K" & endb + jia) = jia
    Next
    wb.Sheets("erp").Sort.SortFields.Clear
    wb.Sheets("erp").Sort.SortFields.Add Key:=wb.Sheets("erp").Range("K1:K" & endb + zhz), _
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
    xlSortTextAsNumbers
    With wb.Sheets("erp").Sort
        .SetRange Range("A1:k" & endb + zhz)
        .Header = 2
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    wb.Sheets("erp").Rows("1:1").Insert Shift:=xlDown
    wb.Sheets("erp").Columns("J:K").Delete Shift:=xlToLeft
    wb.Sheets("erp").Columns("A:A").Delete Shift:=xlToLeft
    wb.Sheets("erp").Range("A1:H" & endb + zhz).Borders.LineStyle = xlContinuous
    wb.Sheets("erp").Columns("A:A").Cut
    wb.Sheets("erp").Columns("H:H").Insert Shift:=xlToRight
    '把型材截面号放到i列,定尺放到H列
    Dim zjh
    zjh = wb.Sheets("erp").Range("A6000").End(xlUp).Row
    'Columns("C:C").Replace " ", ""
    Dim xcki
    Dim xch
    Dim xck
    Dim dingchi
    Dim hangshu
    Dim Slhj
    Dim dys
    Dim k, kk
    For xcki = 2 To zjh '表示型材宽度的计数
        xch = ""
        dingchi = ""
        If Len(wb.Sheets("erp").Range("C" & xcki)) <> 0 Then
            xck = wb.Sheets("erp").Cells(xcki, 3) '型材宽度
            If ThisWorkbook.Sheets("库(待补充)").Columns(1).Find(xck, LookAt:=xlWhole, SearchDirection:=xlPrevious) Is Nothing Then
                xch = "请输入型材"
                dingchi = "输入定尺"
            Else
                hangshu = ThisWorkbook.Sheets("库(待补充)").Columns(1).Find(xck, LookAt:=xlWhole, SearchDirection:=xlPrevious).Row
                xch = ThisWorkbook.Sheets("库(待补充)").Range("B" & hangshu) '型材截面号
                dingchi = ThisWorkbook.Sheets("库(待补充)").Range("C" & hangshu)
            End If
                wb.Sheets("erp").Range("I" & xcki) = xch
                wb.Sheets("erp").Range("J" & xcki) = dingchi
                If xch = "请输入型材" Then
                    wb.Sheets("erp").Range("I" & xcki).Interior.Color = RGB(255, 20, 20)
                    wb.Sheets("erp").Range("J" & xcki).Interior.Color = RGB(255, 20, 20)
                End If
            End If
        If wb.Sheets("erp").Range("G" & xcki) = "K板" Then
            If Left(wb.Sheets("erp").Range("I" & xcki), 4) <> "YK-P" Or wb.Sheets("erp").Range("I" & xcki) <> "ZWGYC-2370" Then
                    wb.Sheets("erp").Range("I" & xcki).Interior.Color = RGB(255, 20, 20)
            End If
        End If
        If InStr(wb.Sheets("erp").Range("G" & xcki), "堵板") > 0 And wb.Sheets("erp").Range("I" & xcki) <> "ZWGYC-3273" Then wb.Sheets("erp").Range("I" & xcki).Interior.Color = RGB(255, 20, 20)
        '2018-11-02  新增,出现例如300D600 这样的情况,是需要用一代板的,这里增加一个判断提示
        If (wb.Sheets("erp").Range("G" & xcki) = "平板" Or wb.Sheets("erp").Range("G" & xcki) = "平面板" Or wb.Sheets("erp").Range("G" & xcki) = "PK板") Then
            If InStr(Split(wb.Sheets("erp").Range("A" & xcki), "-")(1), "D") > 1 Then
                wb.Sheets("erp").Range("I" & xcki).Interior.Color = RGB(255, 20, 20)
            End If
        End If
        '1107 增加了5XC 5SC颜色
        If InStr(wb.Sheets("erp").Range("A" & xcki), "XC") + InStr(wb.Sheets("erp").Range("A" & xcki), "SC") > 1 Then
            wb.Sheets("erp").Range("I" & xcki).Interior.Color = RGB(255, 20, 20)
        End If
        If InStr(wb.Sheets("erp").Range("G" & xcki), "龙骨") + InStr(wb.Sheets("erp").Range("G" & xcki), "铝梁") > 0 And wb.Sheets("erp").Range("I" & xcki) <> "HLD-31" Then wb.Sheets("erp").Range("I" & xcki).Interior.Color = RGB(255, 20, 20)
    Next
    endb = wb.Sheets("erp").[B65536].End(xlUp).Row
    Slhj = Application.WorksheetFunction.Sum(wb.Sheets("erp").Range("b1:b" & endb)) '数量合计
    wb.Sheets("erp").Range("A1:J" & endb).Borders.LineStyle = xlContinuous
    dys = 1
    wb.Sheets("erp").Range("T1:T1").Value = "=RIGHT((MID(A1, FIND(""K"",A1)-4,4)),LEN((MID(A1, FIND(""K"",A1)-4,4)))-FIND(""-"",(MID(A1, FIND(""K"",A1)-4,4))))"
    ' wb.Sheets("erp").Range("T1:T6000").Select
    ' Selection.FillDown
    wb.Sheets("erp").Range("T1").AutoFill Destination:=wb.Sheets("erp").Range("T1:T" & endb), Type:=xlFillSeries
    Do While dys <= endb
        If wb.Sheets("erp").Range("B" & dys) > 0 Then
            k = k + wb.Sheets("erp").Range("B" & dys)
            kk = kk + 1
        Else
            k = 0
            kk = 0
        End If
        If k >= 50 And kk < 25 Then
            If wb.Sheets("erp").Range("B" & dys - 1) > 0 Then
                wb.Sheets("erp").Rows(dys).Insert
                k = 0
                kk = 0
                endb = endb + 1
            Else
                k = 0
                kk = 0
            End If
        ElseIf k < 50 And kk >= 25 Then
            If wb.Sheets("erp").Range("B" & dys - 1) > 0 Then
                wb.Sheets("erp").Rows(dys).Insert
                k = 0
                kk = 0
                endb = endb + 1
            Else
                k = 0
                kk = 0
            End If
        ElseIf k >= 50 And kk >= 25 Then
            If wb.Sheets("erp").Range("B" & dys - 1) > 0 Then
                wb.Sheets("erp").Rows(dys).Insert
                k = 0
                kk = 0
                endb = endb + 1
            Else
                k = 0
                kk = 0
            End If
        End If
        dys = dys + 1
    Loop
    wb.Sheets("erp").Columns("H:H").Cut
    wb.Sheets("erp").Columns("K:K").Insert Shift:=xlToRight
    wb.Sheets("erp").Columns("G:G").Insert Shift:=xlToRight
    wb.Sheets("erp").Columns("A:J").HorizontalAlignment = xlCenter
    Dim enda
    enda = wb.Sheets("erp").Range("A6000").End(xlUp).Row
    '先用设计提供的模板名称进行判断
    Dim XCHI, TZBH, MM
    For XCHI = 2 To enda '型材号计数i
        On Error Resume Next
        TZBH = wb.Sheets("erp").Range("H" & XCHI) '图纸编号
        '背孔打筋判断-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        MM = wb.Sheets("erp").Range("U" & XCHI).Value
        ' MsgBox MM
        If (20 < MM) Then
            If InStr(TZBH, "平面板") Or InStr(TZBH, "平板") Or InStr(TZBH, "普板") Or InStr(TZBH, "墙板") Or InStr(TZBH, "PK板") > 0 Or InStr(TZBH, "平面板切斜") > 0 Then '判断打筋
                If wb.Sheets("erp").Range("I" & XCHI) = "ZWGYC-3265" Then '100P
                    If 45 <= MM Or MM <= 55 Then
                        wb.Sheets("erp").Range("I" & XCHI) = "ZWGYC-3265"
                    Else
                        ' Range("I" & XCHI) = "YK-P002"
                        wb.Sheets("erp").Range("I" & XCHI).Interior.Color = RGB(127, 255, 212) '青色
                    End If
                End If
                If wb.Sheets("erp").Range("I" & XCHI) = "ZWGYC-3266" Then '150P
                    If 45 <= MM Or MM <= 105 Then
                        wb.Sheets("erp").Range("I" & XCHI) = "ZWGYC-3266"
                    Else
                        '  Range("I" & XCHI) = "YK-P003"
                        wb.Sheets("erp").Range("I" & XCHI).Interior.Color = RGB(127, 255, 212) '青色
                    End If
                End If
                If wb.Sheets("erp").Range("I" & XCHI) = "ZWGYC-3267" Then '200P
                    If 45 <= MM Or MM <= 155 Then
                        wb.Sheets("erp").Range("I" & XCHI) = "ZWGYC-3267"
                    Else
                        '  Range("I" & XCHI) = "YK-P004"
                        wb.Sheets("erp").Range("I" & XCHI).Interior.Color = RGB(127, 255, 212) '青色
                    End If
                End If
                '-----------------------------------------------------------------------------------------------------
                If wb.Sheets("erp").Range("I" & XCHI) = "ZWGYC-3268" Then '250P，W
                    If InStr(wb.Sheets("erp").Range("A" & XCHI), "W") > 0 Then   '铣槽
                        If (45 <= MM And MM <= 100) Or (150 <= MM And MM <= 205) Then
                            wb.Sheets("erp").Range("I" & XCHI) = "ZWGYC-3888"
                        ElseIf (50 <= MM And MM <= 85) Or (110 <= MM And MM <= 140) Or (165 <= MM And MM <= 200) Then '铣槽二代板
                            wb.Sheets("erp").Range("I" & XCHI) = "ZWGYC-3268"
                            wb.Sheets("erp").Range("I" & XCHI).Interior.Color = RGB(255, 193, 37) '橙色
                        Else
                            wb.Sheets("erp").Range("I" & XCHI) = "YK-P005" '铣槽一代板
                            wb.Sheets("erp").Range("I" & XCHI).Interior.Color = RGB(127, 255, 212) '青色
                        End If
                    Else
                        If (50 <= MM And MM <= 85) Or (110 <= MM And MM <= 140) Or (165 <= MM And MM <= 200) Then '二代板
                            wb.Sheets("erp").Range("I" & XCHI) = "ZWGYC-3268"
                        Else
                            '  wb.Sheets("erp").Range("I" & XCHI) = "YK-P005" '一代板
                            wb.Sheets("erp").Range("I" & XCHI).Interior.Color = RGB(127, 255, 212) '青色
                        End If
                    End If
                End If
                '-----------------------------------------------------------------------------------------------------
                If wb.Sheets("erp").Range("I" & XCHI) = "ZWGYC-3269" Then '300P，W
                    If InStr(wb.Sheets("erp").Range("A" & XCHI), "W") > 0 Then   '铣槽
                        If (45 <= MM And MM <= 125) Or (175 <= MM And MM <= 255) Then
                            wb.Sheets("erp").Range("I" & XCHI) = "ZWGYC-3887"
                        ElseIf (50 <= MM And MM <= 110) Or (135 <= MM And MM <= 165) Or (190 <= MM And MM <= 250) Then '铣槽二代板
                            wb.Sheets("erp").Range("I" & XCHI) = "ZWGYC-3269"
                            wb.Sheets("erp").Range("I" & XCHI).Interior.Color = RGB(255, 193, 37) '橙色
                        Else
                            wb.Sheets("erp").Range("I" & XCHI) = "YK-P006" '铣槽一代板
                            wb.Sheets("erp").Range("I" & XCHI).Interior.Color = RGB(127, 255, 212) '青色
                        End If
                    Else
                        If (50 <= MM And MM <= 110) Or (135 <= MM And MM <= 165) Or (190 <= MM And MM <= 250) Then '二代板
                            wb.Sheets("erp").Range("I" & XCHI) = "ZWGYC-3269"
                        Else
                            ' wb.Sheets("erp").Range("I" & XCHI) = "YK-P006" '一代板
                            wb.Sheets("erp").Range("I" & XCHI).Interior.Color = RGB(127, 255, 212) '青色
                        End If
                    End If
                End If
                '-----------------------------------------------------------------------------------------------------
                If wb.Sheets("erp").Range("I" & XCHI) = "ZWGYC-3270" Then '350P，W
                    If InStr(wb.Sheets("erp").Range("A" & XCHI), "W") > 0 Then   '铣槽
                        If (45 <= MM And MM <= 150) Or (200 <= MM And MM <= 305) Then
                            wb.Sheets("erp").Range("I" & XCHI) = "ZWGYC-3886"
                        ElseIf (50 <= MM And MM <= 115) Or (140 <= MM And MM <= 210) Or (235 <= MM And MM <= 300) Then '铣槽二代板
                            wb.Sheets("erp").Range("I" & XCHI) = "ZWGYC-3270"
                            wb.Sheets("erp").Range("I" & XCHI).Interior.Color = RGB(255, 193, 37) '橙色
                        Else
                            wb.Sheets("erp").Range("I" & XCHI) = "YK-P007" '铣槽一代板
                            wb.Sheets("erp").Range("I" & XCHI).Interior.Color = RGB(127, 255, 212) '青色
                        End If
                    Else
                        If (50 <= MM And MM <= 115) Or (140 <= MM And MM <= 210) Or (235 <= MM And MM <= 300) Then '二代板
                            wb.Sheets("erp").Range("I" & XCHI) = "ZWGYC-3270"
                        Else
                            '  wb.Sheets("erp").Range("I" & XCHI) = "YK-P007" '一代板
                            wb.Sheets("erp").Range("I" & XCHI).Interior.Color = RGB(127, 255, 212) '青色
                        End If
                    End If
                End If
                '-----------------------------------------------------------------------------------------------------
                If wb.Sheets("erp").Range("I" & XCHI) = "ZWGYC-3271" Then '400P，W
                    If InStr(wb.Sheets("erp").Range("A" & XCHI), "W") > 0 Then   '铣槽
                        If (45 <= MM And MM <= 175) Or (225 <= MM And MM <= 355) Then
                            wb.Sheets("erp").Range("I" & XCHI) = "ZWGYC-3885"
                        ElseIf (50 <= MM And MM <= 160) Or (185 <= MM And MM <= 215) Or (240 <= MM And MM <= 350) Then '铣槽二代板
                            wb.Sheets("erp").Range("I" & XCHI) = "ZWGYC-3271"
                            wb.Sheets("erp").Range("I" & XCHI).Interior.Color = RGB(255, 193, 37) '橙色
                        Else
                            wb.Sheets("erp").Range("I" & XCHI) = "ZWGYC-2370" '铣槽一代板
                            wb.Sheets("erp").Range("I" & XCHI).Interior.Color = RGB(127, 255, 212) '青色
                        End If
                    Else
                        If (50 <= MM And MM <= 160) Or (185 <= MM And MM <= 215) Or (240 <= MM And MM <= 350) Then '二代板
                            wb.Sheets("erp").Range("I" & XCHI) = "ZWGYC-3271"
                    Else
                        '  wb.Sheets("erp").Range("I" & XCHI) = "ZWGYC-2370" '一代板
                        wb.Sheets("erp").Range("I" & XCHI).Interior.Color = RGB(127, 255, 212) '青色
                    End If
                End If
            End If
            '-----------------------------------------------------------------------------------------------------
            If wb.Sheets("erp").Range("I" & XCHI) = "ZWGYC-3272" Then '500P，W
                If InStr(wb.Sheets("erp").Range("A" & XCHI), "W") > 0 Then   '铣槽
                    If (45 <= MM And MM <= 162.5) Or (212.5 <= MM And MM <= 287.5) Or (337.5 <= MM And MM <= 455) Then
                        wb.Sheets("erp").Range("I" & XCHI) = "ZWGYC-3884"
                        ElseIf (50 <= MM And MM <= 152.5) Or (197.5 <= MM And MM <= 302.5) Or (347.5 <= MM And MM <= 450) Then '铣槽二代板
                            wb.Sheets("erp").Range("I" & XCHI) = "ZWGYC-3272"
                            wb.Sheets("erp").Range("I" & XCHI).Interior.Color = RGB(255, 193, 37) '橙色
                        Else
                            '无一代板
                            wb.Sheets("erp").Range("I" & XCHI).Interior.Color = RGB(255, 0, 0) '红色
                        End If
                    Else
                        If (50 <= MM And MM <= 152.5) Or (197.5 <= MM And MM <= 302.5) Or (347.5 <= MM And MM <= 450) Then '二代板
                            wb.Sheets("erp").Range("I" & XCHI) = "ZWGYC-3272"
                        Else
                            wb.Sheets("erp").Range("I" & XCHI).Interior.Color = RGB(127, 255, 212) '青色
                        End If
                    End If
                End If
                '-----------------------------------------------------------------------------------------------------
                If InStr(wb.Sheets("erp").Range("H" & XCHI), "XP") > 0 Then
                    wb.Sheets("erp").Range("I" & XCHI) = "ZWMB-07"
                    wb.Sheets("erp").Range("I" & XCHI).Interior.Color = xlNone '无色
                End If
            End If
        End If
        'C区---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------++
        If InStr(TZBH, "墙柱") > 0 Then  '如果H列写的是墙柱C槽，判断D列长度，小于等于1200，则显示为C1X，否则为C2X
            If wb.Sheets("erp").Range("I" & XCHI) = "HLD-03" Or wb.Sheets("erp").Range("I" & XCHI) = "HLD-15" Then
                If wb.Sheets("erp").Range("D" & XCHI) <= 1200 Then
                    wb.Sheets("erp").Range("G" & XCHI) = "L1"
                Else
                    wb.Sheets("erp").Range("G" & XCHI) = "L2"
                End If
            Else
                If wb.Sheets("erp").Range("D" & XCHI) <= 1200 Then
                    wb.Sheets("erp").Range("G" & XCHI) = "C1X"
                Else
                    wb.Sheets("erp").Range("G" & XCHI) = "C2X"
                End If
            End If
        ElseIf wb.Sheets("erp").Range("H" & XCHI) = "C槽XC" Or wb.Sheets("erp").Range("H" & XCHI) = "C槽SC" Then
            If InStr(wb.Sheets("erp").Range("A" & XCHI), "5SC") > 0 Or InStr(wb.Sheets("erp").Range("A" & XCHI), "5XC") > 0 Then
                If InStr(wb.Sheets("erp").Range("H" & XCHI), "C槽SC") > 0 Then
                    wb.Sheets("erp").Range("I" & XCHI) = "ZWMB-11"
                End If
                If InStr(wb.Sheets("erp").Range("H" & XCHI), "C槽XC") > 0 Then
                    wb.Sheets("erp").Range("I" & XCHI) = "ZWMB-12"
                End If
                If wb.Sheets("erp").Range("D" & XCHI) <= 1200 Then
                    wb.Sheets("erp").Range("G" & XCHI) = "C1"
                Else
                    wb.Sheets("erp").Range("G" & XCHI) = "C2"
                End If
            Else
                If InStr(wb.Sheets("erp").Range("H" & XCHI), "C槽SC") > 0 Then
                    wb.Sheets("erp").Range("I" & XCHI) = "ZWMB-56"
                    wb.Sheets("erp").Range("J" & XCHI) = "6000"
                End If
                If InStr(wb.Sheets("erp").Range("H" & XCHI), "C槽XC") > 0 Then
                    wb.Sheets("erp").Range("I" & XCHI) = "ZWMB-56+HLD-15"
                    wb.Sheets("erp").Range("J" & XCHI) = "6000"
                End If
                If wb.Sheets("erp").Range("D" & XCHI) <= 1200 Then
                    wb.Sheets("erp").Range("G" & XCHI) = "L1"
                Else
                    wb.Sheets("erp").Range("G" & XCHI) = "L2"
                End If
            End If
        ElseIf InStr(TZBH, "C槽") > 0 And InStr(TZBH, "阴角") + InStr(TZBH, "转角") = 0 And InStr(TZBH, "墙柱") = 0 Then '如果H列写的是C槽，判断D列长度，小于等于1200，则显示为C1，否则为C2
            If wb.Sheets("erp").Range("I" & XCHI) = "HLD-03" Or wb.Sheets("erp").Range("I" & XCHI) = "HLD-15" Or wb.Sheets("erp").Range("I" & XCHI) = "请输入型材" Then
                If InStr(wb.Sheets("erp").Range("A" & XCHI), "AS") > 0 Or InStr(wb.Sheets("erp").Range("A" & XCHI), "AL") > 0 Or InStr(wb.Sheets("erp").Range("A" & XCHI), "AR") > 0 Then
                    wb.Sheets("erp").Range("G" & XCHI) = "L1X"
                Else
                    If wb.Sheets("erp").Range("D" & XCHI) <= 1200 Then
                        wb.Sheets("erp").Range("G" & XCHI) = "L1"
                    Else
                        wb.Sheets("erp").Range("G" & XCHI) = "L2"
                    End If
                End If
            ElseIf InStr(wb.Sheets("erp").Range("A" & XCHI), "AS") > 0 Or InStr(wb.Sheets("erp").Range("A" & XCHI), "AL") > 0 Or InStr(wb.Sheets("erp").Range("A" & XCHI), "AR") > 0 Then
                If wb.Sheets("erp").Range("D" & XCHI) <= 1200 Then
                    wb.Sheets("erp").Range("G" & XCHI) = "C1X"
                Else
                    wb.Sheets("erp").Range("G" & XCHI) = "C2X"
                End If
            Else
                If wb.Sheets("erp").Range("D" & XCHI) <= 1200 Then
                    wb.Sheets("erp").Range("G" & XCHI) = "C1"
                Else
                    wb.Sheets("erp").Range("G" & XCHI) = "C2"
                End If
            End If
        'N区---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------++
        ElseIf InStr(TZBH, "阴角C槽") + InStr(TZBH, "转角") > 0 Then   '如果是阴角C槽或者是写的是转角，则先看型材是不是L 板，
            If wb.Sheets("erp").Range("I" & XCHI) = "HLD-03" Or wb.Sheets("erp").Range("I" & XCHI) = "HLD-15" Or wb.Sheets("erp").Range("I" & XCHI) = "请输入型材" Then
                wb.Sheets("erp").Range("G" & XCHI) = "N2"
            ElseIf InStr(wb.Sheets("erp").Range("A" & XCHI), "A") > 0 Or InStr(wb.Sheets("erp").Range("A" & XCHI), "V") > 0 Then
                wb.Sheets("erp").Range("G" & XCHI) = "N1X"
            ElseIf InStr(wb.Sheets("erp").Range("A" & XCHI), "CN") > 0 Then
                If wb.Sheets("erp").Range("D" & XCHI) <= 1200 Then
                    wb.Sheets("erp").Range("G" & XCHI) = "C1"
                Else
                    wb.Sheets("erp").Range("G" & XCHI) = "C2"
                End If
            ElseIf InStr(wb.Sheets("erp").Range("A" & XCHI), "QN") > 0 Then
                If wb.Sheets("erp").Range("D" & XCHI) <= 1200 Then
                    wb.Sheets("erp").Range("G" & XCHI) = "C1X"
                Else
                    wb.Sheets("erp").Range("G" & XCHI) = "C2X"
                End If
            Else
                wb.Sheets("erp").Range("G" & XCHI) = "N1"
            End If
        'QT区---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------++
        ElseIf InStr(TZBH, "支撑") + InStr(TZBH, "固顶") > 0 Then '支撑
            If InStr(wb.Sheets("erp").Range("A" & XCHI), "LTZ") > 0 Then
                If wb.Sheets("erp").Range("D" & XCHI) <= 1200 Then
                    wb.Sheets("erp").Range("G" & XCHI) = "L1"
                Else
                    wb.Sheets("erp").Range("G" & XCHI) = "L2"
                End If
            Else
                wb.Sheets("erp").Range("G" & XCHI) = "ZC"
            End If
        ElseIf InStr(TZBH, "龙骨") + InStr(TZBH, "铝梁") > 0 Then '龙骨或是铝梁
            wb.Sheets("erp").Range("G" & XCHI) = "LG"
        ElseIf InStr(TZBH, "堵") + InStr(TZBH, "梁底") > 0 Then '堵板或梁底板
            wb.Sheets("erp").Range("G" & XCHI) = "D"
        ElseIf wb.Sheets("erp").Range("H" & XCHI) = "截面异形板" Then '截面异形板
            wb.Sheets("erp").Range("G" & XCHI) = "P1X"
        ElseIf InStr(TZBH, "铝盒子") Or InStr(TZBH, "传料箱") Or InStr(TZBH, "泵送盒") Or InStr(TZBH, "放线口") Or InStr(wb.Sheets("erp").Range("A" & XCHI), "XH") > 0 Or InStr(wb.Sheets("erp").Range("E" & XCHI), "XH") > 0 Then '铝盒子
            wb.Sheets("erp").Range("G" & XCHI) = "CLX"
        ElseIf wb.Sheets("erp").Range("H" & XCHI) = "L型倒角板" Then 'L型倒角板
            wb.Sheets("erp").Range("G" & XCHI) = "N1"
        'P区---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ElseIf InStr(TZBH, "平面板") Or InStr(TZBH, "平板") Or InStr(TZBH, "普板") Or InStr(TZBH, "墙板") Or InStr(TZBH, "PK板") > 0 Or InStr(TZBH, "平面板切斜") > 0 Then '如果是平面板或平板或普板或楼梯墙板，则先需要判断型号是不是HLD-03或15，如果是则为L板，再判断是L1还是L2，如果不是l板则根据宽度判断，然后判断长度
            Dim W1W2, l
            W1W2 = wb.Sheets("erp").Range("C" & XCHI)
            l = wb.Sheets("erp").Range("C" & XCHI)
            If wb.Sheets("erp").Range("I" & XCHI) = "HLD-03" Or wb.Sheets("erp").Range("I" & XCHI) = "HLD-15" Or InStr(wb.Sheets("erp").Range("I" & XCHI), "+") > 0 Or wb.Sheets("erp").Range("I" & XCHI) = "请输入型材" Then
                If wb.Sheets("erp").Range("D" & XCHI) <= 1200 Then '小于1200
                    If InStr(wb.Sheets("erp").Range("A" & XCHI), "TQ") > 0 Or InStr(TZBH, "平面板切斜") Then 'TQ切斜
                        wb.Sheets("erp").Range("G" & XCHI) = "L1X"
                    ElseIf InStr(wb.Sheets("erp").Range("A" & XCHI), "W") > 0 Then '铣槽
                        wb.Sheets("erp").Range("G" & XCHI) = "L1W"
                    Else
                        wb.Sheets("erp").Range("G" & XCHI) = "L1" '其余小于1200的
                    End If
                Else '大于1200的，同上
                    If InStr(wb.Sheets("erp").Range("A" & XCHI), "TQ") > 0 Or InStr(TZBH, "平面板切斜") Then 'TQ切斜
                        wb.Sheets("erp").Range("G" & XCHI) = "L1X"
                    ElseIf InStr(wb.Sheets("erp").Range("A" & XCHI), "W") > 0 Then
                        wb.Sheets("erp").Range("G" & XCHI) = "L2W"
                    Else
                        wb.Sheets("erp").Range("G" & XCHI) = "L2"
                    End If
                End If
            ElseIf wb.Sheets("erp").Range("I" & XCHI) = "HLD-37" Or wb.Sheets("erp").Range("I" & XCHI) = "HLD-40" Or wb.Sheets("erp").Range("I" & XCHI) = "ZWGYC-3139" Or wb.Sheets("erp").Range("I" & XCHI) = "ZWGYC-3138" Or wb.Sheets("erp").Range("I" & XCHI) = "HLD-68" Or wb.Sheets("erp").Range("I" & XCHI) = "HLD-42" Or wb.Sheets("erp").Range("I" & XCHI) = "ZWGYC-2370" Or Left(wb.Sheets("erp").Range("I" & XCHI), 4) = "YK-P" Then
                If InStr(wb.Sheets("erp").Range("A" & XCHI), "TQ") > 0 Or InStr(TZBH, "平面板切斜") Then 'TQ切斜
                    If wb.Sheets("erp").Range("D" & XCHI) <= 1500 Then
                        wb.Sheets("erp").Range("G" & XCHI) = "K1X"
                    Else
                        wb.Sheets("erp").Range("G" & XCHI) = "K2X"
                    End If
                ElseIf InStr(wb.Sheets("erp").Range("A" & XCHI), "W") > 0 Then '铣槽
                    If wb.Sheets("erp").Range("D" & XCHI) <= 1200 Then
                        wb.Sheets("erp").Range("G" & XCHI) = "P1W"
                    Else
                        wb.Sheets("erp").Range("G" & XCHI) = "P2W"
                    End If
                Else
                    If wb.Sheets("erp").Range("D" & XCHI) <= 1500 Then '普通K板
                        wb.Sheets("erp").Range("G" & XCHI) = "K1"
                    Else
                        wb.Sheets("erp").Range("G" & XCHI) = "K2"
                    End If
                End If
            ElseIf wb.Sheets("erp").Range("I" & XCHI) = "ZWGYC-3265" Or wb.Sheets("erp").Range("I" & XCHI) = "ZWGYC-3266" Or wb.Sheets("erp").Range("I" & XCHI) = "ZWGYC-3267" Or wb.Sheets("erp").Range("I" & XCHI) = "ZWMB-07" Then      '判断P1,P2
                If wb.Sheets("erp").Range("D" & XCHI) <= 1200 Then '小于1200
                    If InStr(wb.Sheets("erp").Range("A" & XCHI), "TQ") > 0 Or InStr(TZBH, "平面板切斜") Then 'TQ切斜
                        wb.Sheets("erp").Range("G" & XCHI) = "P1X"
                    ElseIf InStr(wb.Sheets("erp").Range("A" & XCHI), "W") > 0 Then '铣槽
                        wb.Sheets("erp").Range("G" & XCHI) = "P1W"
                    Else
                        wb.Sheets("erp").Range("G" & XCHI) = "P1" '其余大于1200的
                    End If
                Else '大于1200的，同上
                    If InStr(wb.Sheets("erp").Range("A" & XCHI), "TQ") > 0 Or InStr(TZBH, "平面板切斜") Then
                        wb.Sheets("erp").Range("G" & XCHI) = "P2X"
                    ElseIf InStr(wb.Sheets("erp").Range("A" & XCHI), "W") > 0 Then
                        wb.Sheets("erp").Range("G" & XCHI) = "P2W"
                    Else
                        wb.Sheets("erp").Range("G" & XCHI) = "P2"
                    End If
                End If
            ElseIf wb.Sheets("erp").Range("I" & XCHI) = "ZWGYC-3268" Or wb.Sheets("erp").Range("I" & XCHI) = "ZWGYC-3269" Or wb.Sheets("erp").Range("I" & XCHI) = "ZWGYC-3270" Or wb.Sheets("erp").Range("I" & XCHI) = "ZWGYC-3271" Or wb.Sheets("erp").Range("I" & XCHI) = "ZWGYC-3888" Or wb.Sheets("erp").Range("I" & XCHI) = "ZWGYC-3887" Or wb.Sheets("erp").Range("I" & XCHI) = "ZWGYC-3886" Or wb.Sheets("erp").Range("I" & XCHI) = "ZWGYC-3885" Or wb.Sheets("erp").Range("I" & XCHI) = "ZWGYC-3884" Then '判断P3,P4
                If wb.Sheets("erp").Range("D" & XCHI) <= 1200 Then '小于1200
                    If InStr(wb.Sheets("erp").Range("A" & XCHI), "TQ") > 0 Or InStr(TZBH, "平面板切斜") Then 'TQ切斜
                        wb.Sheets("erp").Range("G" & XCHI) = "P3X"
                    ElseIf InStr(wb.Sheets("erp").Range("A" & XCHI), "W") > 0 Then '铣槽
                        wb.Sheets("erp").Range("G" & XCHI) = "P3W"
                    Else
                        wb.Sheets("erp").Range("G" & XCHI) = "P3" '其余大于1200的
                    End If
                Else '大于1200的，同上
                    If InStr(wb.Sheets("erp").Range("A" & XCHI), "TQ") > 0 Or InStr(TZBH, "平面板切斜") Then
                        wb.Sheets("erp").Range("G" & XCHI) = "P4X"
                    ElseIf InStr(wb.Sheets("erp").Range("A" & XCHI), "W") > 0 Then
                        wb.Sheets("erp").Range("G" & XCHI) = "P4W"
                    Else
                        wb.Sheets("erp").Range("G" & XCHI) = "P4"
                    End If
                End If
            ElseIf wb.Sheets("erp").Range("I" & XCHI) = "ZWGYC-3272" Or wb.Sheets("erp").Range("I" & XCHI) = "ZWGYC-3884" Then '判断P5,P6
                If wb.Sheets("erp").Range("D" & XCHI) <= 1200 Then '小于1200
                    If InStr(wb.Sheets("erp").Range("A" & XCHI), "TQ") > 0 Or InStr(TZBH, "平面板切斜") Then 'TQ切斜
                        wb.Sheets("erp").Range("G" & XCHI) = "P5X"
                    ElseIf InStr(wb.Sheets("erp").Range("A" & XCHI), "W") > 0 Then '铣槽
                        wb.Sheets("erp").Range("G" & XCHI) = "P5W"
                    Else
                        wb.Sheets("erp").Range("G" & XCHI) = "P5" '其余小于1200的
                    End If
                Else '大于1200的，同上
                    If InStr(wb.Sheets("erp").Range("A" & XCHI), "TQ") > 0 Or InStr(TZBH, "平面板切斜") Then
                        wb.Sheets("erp").Range("G" & XCHI) = "P6X"
                    ElseIf InStr(wb.Sheets("erp").Range("A" & XCHI), "W") > 0 Then
                        wb.Sheets("erp").Range("G" & XCHI) = "P6W"
                    Else
                        wb.Sheets("erp").Range("G" & XCHI) = "P6"
                    End If
                End If
            ElseIf Left(wb.Sheets("erp").Range("I" & XCHI), 4) = "YK-P" Then
                If wb.Sheets("erp").Range("D" & XCHI) <= 1500 Then '这个是判断长度的
                    wb.Sheets("erp").Range("G" & XCHI) = "K1"
                Else
                    wb.Sheets("erp").Range("G" & XCHI) = "K2"
                End If
            End If
        'K区---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------++
        ElseIf TZBH = "K板" Then '如果H列写的是K板，判断D列长度，小于等于1500，则显示为K1，否则为K2,另外如果型号不是开头是"YK-P",则对应的I列的型材单元格底色为红色
            If wb.Sheets("erp").Range("D" & XCHI) <= 1500 Then '这个是判断长度的
                wb.Sheets("erp").Range("G" & XCHI) = "K1"
            Else
                wb.Sheets("erp").Range("G" & XCHI) = "K2"
            End If
        'J区---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------++
        ElseIf wb.Sheets("erp").Range("H" & XCHI) = "角铝" Or wb.Sheets("erp").Range("H" & XCHI) = "底角铝" Then
            If InStr(wb.Sheets("erp").Range("E" & XCHI), "ZW-J-08") > 0 Or InStr(wb.Sheets("erp").Range("E" & XCHI), "ZW-J-09") > 0 Or InStr(wb.Sheets("erp").Range("E" & XCHI), "ZW-J-10") > 0 Then
                wb.Sheets("erp").Range("G" & XCHI) = "JX2"
            ElseIf wb.Sheets("erp").Range("H" & XCHI) = "角铝" And InStr(wb.Sheets("erp").Range("A" & XCHI), "D-JL") > 0 Then 'Z形角铝
                wb.Sheets("erp").Range("G" & XCHI) = "L1"
            ElseIf wb.Sheets("erp").Range("H" & XCHI) = "底角铝" And InStr(wb.Sheets("erp").Range("A" & XCHI), "NJ") > 0 Or InStr(wb.Sheets("erp").Range("A" & XCHI), "ZW-J-08") > 0 Or InStr(wb.Sheets("erp").Range("A" & XCHI), "ZW-J-09") > 0 Or InStr(wb.Sheets("erp").Range("A" & XCHI), "ZW-J-10") > 0 Then
                wb.Sheets("erp").Range("G" & XCHI) = "JX2"
            Else
                wb.Sheets("erp").Range("G" & XCHI) = "J"
            End If
        ElseIf wb.Sheets("erp").Range("H" & XCHI) = "7字角铝" Or wb.Sheets("erp").Range("H" & XCHI) = "角铝封板" Then '如果是7字角铝或者是角铝封板，类型为JX1
            wb.Sheets("erp").Range("G" & XCHI) = "JX1"
        ElseIf InStr(wb.Sheets("erp").Range("E" & XCHI), "LDJ") > 0 Then  '如果是LDJ大样图
            If wb.Sheets("erp").Range("E" & XCHI) = "ZW-LDJ-01" Or wb.Sheets("erp").Range("E" & XCHI) = "ZW-LDJ-02" Then 'LDJ-01或02大样图为J
                wb.Sheets("erp").Range("G" & XCHI) = "J"
            Else
                wb.Sheets("erp").Range("G" & XCHI) = "L1"
            End If
        ElseIf wb.Sheets("erp").Range("H" & XCHI) = "封板" Or wb.Sheets("erp").Range("H" & XCHI) = "铝板" Then  '吊模中的板子
            If InStr(wb.Sheets("erp").Range("A" & XCHI), "J") > 0 And InStr(wb.Sheets("erp").Range("A" & XCHI), "F") + InStr(wb.Sheets("erp").Range("A" & XCHI), "L") > 0 Then
                wb.Sheets("erp").Range("G" & XCHI) = "JX1"
            ElseIf InStr(wb.Sheets("erp").Range("A" & XCHI), "FB") > 0 Then
                wb.Sheets("erp").Range("G" & XCHI) = "FB"
            End If
        'LT区---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------++
        ElseIf InStr(TZBH, "抬头板") > 0 Then '如果H列写的是折板
            wb.Sheets("erp").Range("G" & XCHI) = "L1X"
        ElseIf InStr(TZBH, "挡板") > 0 Then '楼梯挡板
            wb.Sheets("erp").Range("G" & XCHI) = "DB"
        ElseIf InStr(TZBH, "盖板") > 0 Then '楼梯盖板
            wb.Sheets("erp").Range("G" & XCHI) = "GB"
        ElseIf InStr(TZBH, "侧板") > 0 Then '楼梯侧板
            wb.Sheets("erp").Range("G" & XCHI) = "CB"
        ElseIf InStr(TZBH, "狗牙") > 0 Then  '狗牙或是楼梯狗牙
            If wb.Sheets("erp").Range("D" & XCHI) <= 1500 Then '这个是判断长度的
                wb.Sheets("erp").Range("G" & XCHI) = "T1"
            Else
                wb.Sheets("erp").Range("G" & XCHI) = "T2"
            End If
            If InStr(wb.Sheets("erp").Range("I" & XCHI), "+") > 0 And InStr(wb.Sheets("erp").Range("H" & XCHI), "狗牙") + InStr(wb.Sheets("erp").Range("H" & XCHI), "封板") + InStr(wb.Sheets("erp").Range("H" & XCHI), "角铝") = 0 Then
                If wb.Sheets("erp").Range("D" & XCHI) <= 1200 Then
                    wb.Sheets("erp").Range("G" & XCHI) = "L1"
                Else
                    wb.Sheets("erp").Range("G" & XCHI) = "L2"
                End If
            End If
        End If
        If Len(wb.Sheets("erp").Range("H" & XCHI)) > 0 And Len(wb.Sheets("erp").Range("G" & XCHI)) = 0 Then
            wb.Sheets("erp").Range("G" & XCHI).Interior.Color = RGB(255, 0, 0)
        End If
        If Mid(wb.Sheets("erp").Range("F" & XCHI), 2, 1) = "-" Or Mid(wb.Sheets("erp").Range("F" & XCHI), 3, 1) = "-" And Len("A" & XCHI) > 0 Then
            wb.Sheets("erp").Range("M" & XCHI) = Split(wb.Sheets("erp").Range("F" & XCHI), "-")(0)
            If Left(wb.Sheets("erp").Range("E" & XCHI), 2) <> "ZW" Then
                wb.Sheets("erp").Range("N" & XCHI) = wb.Sheets("erp").Range("E" & XCHI)
            End If
        ElseIf Len("A" & XCHI) = 0 Then
            wb.Sheets("erp").Range("M" & XCHI) = ""
        End If
    Next

    ' wb.Sheets("erp").Columns("U:U").Select
    ' Selection.Delete
    wb.Sheets("erp").Columns("U:U").Delete

    '根据零件的分区编号得出来是哪个区域, '这一步将每个图纸用到的非标件图纸取出来
    ' wb.Sheets("erp").Columns("M:N").Select
    wb.Sheets("erp").Range("M1") = "分区"
    wb.Sheets("erp").Range("N1") = "图纸编号"
    wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "erp!R1C13:R1048576C14", Version:=xlPivotTableVersion14).CreatePivotTable TableDestination:= _
        "erp!R1C18", TableName:="数据透视表2", DefaultVersion:=xlPivotTableVersion14
    wb.ShowPivotTableFieldList = True
    With wb.Sheets("erp").PivotTables("数据透视表2").PivotFields("分区")
            .Orientation = xlRowField
        .Position = 1
    End With
    With wb.Sheets("erp").PivotTables("数据透视表2").PivotFields("图纸编号")
        .Orientation = xlRowField
        .Position = 2
    End With
    wb.ShowPivotTableFieldList = False
    ' Columns("R:R").Select
    ' Selection.Copy
    ' Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    '     :=False, Transpose:=False
    ' Application.CutCopyMode = False
    Dim endr
    endr = wb.Sheets("erp").[r60000].End(xlUp).Row
    wb.Sheets("erp").Range("P1:P" & endr) = wb.Sheets("erp").Range("R1:R" & endr).Value
    
    wb.Sheets("erp").Range("P" & endr) = ""
    wb.Sheets("erp").Range("P" & endr - 1) = ""
    wb.Sheets("erp").Range("P1") = "分区非大样图"
    wb.Sheets("erp").Columns("P:P").Replace "*白*", ""
    
    wb.Sheets("erp").Columns("R:R").Delete
    
    wb.Sheets("erp").Columns("M:O").Delete Shift:=xlToLeft
    wb.Sheets("erp").Columns("I:J").Columns.AutoFit
    'MsgBox ("总数量：" & Slhj & " 件" & Chr(10) & "蓝色格为生产单的第25个零件号" & Chr(10) & "注意特殊型材" & Chr(10) & "如平面板编号包含D,判断是否用一代板")
   ' MsgBox ("注意特殊型材" & Chr(10) & "如平面板编号包含D,判断是否用一代板")
    wb.Windows(1).Visible = True
    wb.Close (True)
End Sub

' 填充单号
Sub FB3(scqdFilename As String)
    Dim wb As Workbook
    Set wb = Workbooks.Open(scqdFilename)
    wb.Windows(1).Visible = False
    ThisWorkbook.Activate

    wb.Sheets("erp").Columns("A:A").Interior.Pattern = xlNone
    enda = Range("A60000").End(xlUp).Row
    For dys = 1 To enda '单页数，一个单子要放的数量
        If Len(wb.Sheets("erp").Range("B" & dys)) = 0 And Len(wb.Sheets("erp").Range("B" & dys + 1)) > 0 Then
            k = 0
        Else
            k = k + 1
        End If
        Remainder = k Mod 26  '余数
        If k > 0 And Remainder = 0 Then
            wb.Sheets("erp").Range("A" & dys - 1).Interior.Color = RGB(232, 159, 187)
        lj = lj + 1 '累计次数
        End If
    Next
    If lj > 0 Then
        MsgBox "还有超过单页25的生产单,调整后重新点击填充单号"
        Exit Sub
    End If
    mbmc = ""
    wb.Sheets("erp").Range("A2:G" & enda).Borders.LineStyle = xlContinuous
    If Len(wb.Sheets("erp").Range("A1")) = 0 Then a = wb.Sheets("临时").Range("B3")
    wb.Sheets("erp").Range("A1") = a & "-" & wb.Sheets("erp").Range("K2") & "-1"
    k = 1
    For ih = 1 To enda
        If Len(wb.Sheets("erp").Range("A" & ih)) = 0 Then
            fenqulast = wb.Sheets("erp").Range("K" & ih - 1)
            fenqu = wb.Sheets("erp").Range("K" & ih + 1)
            If fenqu = fenqulast Then
                k = k + 1
                wb.Sheets("erp").Range("A" & ih) = a & "-" & fenqu & "-" & k
            Else
                k = 1
                wb.Sheets("erp").Range("A" & ih) = a & "-" & fenqu & "-" & k
            End If
        End If
    Next
    wb.Sheets("erp").Columns("H:H").FormatConditions.Delete
  '  MsgBox ("如果图号有5XC,5SC,请注意修改型材")
  '  MsgBox ("如需修改型材，请直接在erp表修改,注意切斜模板")
    wb.Windows(1).Visible = True
    wb.Close (True)
End Sub


' copySheet效果很差，基本作废了，建议用excel自带的copy替代
Private Sub copySheet(src As Worksheet, dst As Worksheet)
    Dim ur As Range
    Dim rowCount As Long
    Dim ColumnCount As Long
    Set ur = src.UsedRange
    ColumnCount = ur.Columns.count
    rowCount = ur.Rows.count

    dst.[A1].Resize(rowCount, ColumnCount) = src.UsedRange.Value
End Sub

Private Function isSheetExist(wb As Workbook, shtName As String) As Boolean
    Dim sht As Worksheet
    For Each sht In wb.Sheets
        If sht.Name = shtName Then
            isSheetExist = True
            Exit Function
        End If
    Next
    isSheetExist = False
End Function

Function fileIsExist(fileFullPath As String) As Boolean
 Dim fso As Object
 Dim ret As Boolean
 Set fso = CreateObject("Scripting.FileSystemObject")
 ret = False
 If fso.FileExists(fileFullPath) = True Then
     ret = True
 End If
  Set fso = Nothing
  fileIsExist = ret
End Function

Sub createExcel(fileFullPath As String)
    Dim excelApp, excelWB As Object
    Dim savePath, saveName As String

    Set excelApp = CreateObject("Excel.Application")
    Set excelWB = excelApp.Workbooks.Add

    excelWB.SaveAs fileFullPath
    excelApp.Quit
End Sub

