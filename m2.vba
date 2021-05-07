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
    ' fpqdFilename = "C:\Users\u03013112\Documents\002\C-494.xlsx"
    
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
Sub testFB4()
    Dim scqdFilename As String
    Dim sjkFilename As String
    scqdFilename = "C:\Users\u03013112\Documents\J\new-412-1\分配清单\" & txtlpdm & txtgcmc & txtqyjx & "-生产单.xlsx"
    sjkFilename = "C:\Users\u03013112\Documents\J\new-412-1\分配清单\" & txtlpdm & txtgcmc & txtqyjx & "-数据库.xlsx"
    Call FB4(scqdFilename, sjkFilename)
End Sub
Sub testFB5()
    Dim scqdFilename As String
    scqdFilename = "C:\Users\u03013112\Documents\J\new-412-1\分配清单\" & txtlpdm & txtgcmc & txtqyjx & "-生产单.xlsx"
    Call FB5(scqdFilename)
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
        .Add Key:=wb.Sheets("erp").Range("B2"), Order:=1   '模板名称
        .Add Key:=wb.Sheets("erp").Range("E2"), Order:=1   'W1
        .Add Key:=wb.Sheets("erp").Range("F2"), Order:=1   'W2
        .Add Key:=wb.Sheets("erp").Range("H2"), Order:=1   '图纸编号
        .Add Key:=wb.Sheets("erp").Range("G2"), Order:=1   'L
    End With
    With wb.Sheets("erp").Sort
        .SetRange wb.Sheets("erp").Range("b2:M" & endb)
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
        .SetRange wb.Sheets("erp").Range("A1:k" & endb + zhz)
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
    Dim enda, dys
    enda = wb.Sheets("erp").Range("A60000").End(xlUp).Row
    Dim k, kk
    Dim Remainder, lj
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
    Dim mbmc
    Dim a
    Dim ih
    mbmc = ""
    wb.Sheets("erp").Range("A2:G" & enda).Borders.LineStyle = xlContinuous
    If Len(wb.Sheets("erp").Range("A1")) = 0 Then a = wb.Sheets("临时").Range("B3")
    wb.Sheets("erp").Range("A1") = a & "-" & wb.Sheets("erp").Range("K2") & "-1"
    k = 1
    Dim fenqulast, fenqu
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

' 转序单生产单
Sub FB4(scqdFilename As String, sjkFilename As String)
    Dim wb As Workbook
    Set wb = Workbooks.Open(scqdFilename)
    'wb.Windows(1).Visible = False
    ' ThisWorkbook.Activate

    Dim p&, zhs&, i1&, jch&, PE&, sn&

    ' Application.ScreenUpdating = False
    
    zhs = wb.Sheets("erp").[b6000].End(xlUp).Row
    
    wb.Sheets("erp").Columns("M:O").Delete Shift:=xlToLeft
    wb.Sheets("erp").Columns("A:J").HorizontalAlignment = xlCenter
    wb.Sheets("erp").Range("A1:j6000").Interior.Pattern = xlNone
    wb.Sheets("erp").Columns("B:C").Insert
    wb.Sheets("erp").Columns("H:H").Insert
    
    Dim wz
    Dim jsp
    wz = Mid(wb.Sheets("erp").Cells(1, 1), 1, 4)
    
    For p = 1 To zhs '合并图纸编号单元格，以方便放入转序单
        If Mid(wb.Sheets("erp").Cells(p, 1), 1, 4) = wz Then
            jsp = jsp + 1 '计数p
        Else
            wb.Sheets("erp").Cells(p, 1).Resize(, 3).Merge
            wb.Sheets("erp").Cells(p, 7).Resize(, 2).Merge
        End If
    Next
    
    ' Sheet5.Activate
    'Sheets("ZXD").Range("A2") = "项目名称：" & Sheet5.Range("B2") & Sheet5.Range("G2")
    '=======================================================================================
    '191211按模板厂需求调整生产计划单格式
    
    Dim aaa As String '用于存储生产单号
    Dim shp
    
    wb.Sheets("临时").Rows(3).RowHeight = 70
    wb.Sheets("临时").Rows("7:31").RowHeight = 19
    wb.Sheets("临时").Range("B3").Font.Size = 36
    
    '=======================================================================================
    ' 原本在下面循环里的代码，复制之前做，提高效率
    Dim pic
    For Each pic In wb.Sheets("临时").Pictures
        pic.Left = (pic.TopLeftCell.Width - pic.Width) / 2 + 2.3 * pic.TopLeftCell.Left
        pic.Top = (pic.TopLeftCell.Height - pic.Height) / 2 + 1.05 * pic.TopLeftCell.Top
    Next

    Dim sht
    For Each sht In wb.Worksheets
        If InStr(sht.Name, "临时表") > 0 Then
            sht.Delete
        End If
    Next

    Dim zjb
    For zjb = 1 To jsp '增加表，原有逻辑少加了一张表，为的是 临时 也可以使用
        ' Sheet5.Select
        ' Sheet5.Copy Before:=Sheets("erp")
        wb.Sheets.Add(before:=wb.Sheets("erp")).Name = "临时表" & zjb
        wb.Sheets("临时").Cells.Copy wb.Sheets("临时表" & zjb).[A1]
    Next
       
    ' Sheets("erp").Activate
    
    wb.Sheets("erp").Range("A1:H" & zhs).Borders.LineStyle = xlContinuous  '合并居中加边框
    
    For i1 = 1 To zhs + jsp * 2
        If Mid(wb.Sheets("erp").Cells(i1, 1), 1, 4) = wz Then
            wb.Sheets("erp").Rows(i1).Insert
            wb.Sheets("erp").Rows(i1 + 2).Insert
            i1 = i1 + 1
        End If
    Next
    
    Dim bs As Long
    Dim r, xck, mbmc, Quyu, xch, dingchi
    bs = 1
    For jch = 1 To zhs + jsp * 2 '加插入行的总行数
        On Error Resume Next
        If Mid(wb.Sheets("erp").Cells(jch, 1), 1, 4) = wz Then
            Set sht = wb.Sheets("临时表" & bs)
            bs = bs + 1
            r = wb.Sheets("erp").Cells(jch + 2, 1).CurrentRegion.Rows.count '连续区域的行数
            'c = wb.Sheets("erp").Cells(jch + 2, 1).CurrentRegion.Columns.Count '连续区域的列数，然并卵
            xck = wb.Sheets("erp").Cells(jch + 2, 5) '型材宽度
            mbmc = wb.Sheets("erp").Cells(jch + 2, 1).Text '模板名称
            wb.Sheets("erp").Cells(jch, 1).Copy sht.[B3:E3] '转序单号复制
            
            '=======================================================================================
            '191211按模板厂需求调整生产计划单格式
            
            ' sht.Activate
            
            With sht
                ' 这段实在看不懂，什么SB逻辑
                aaa = .Range("B3")
                .Range("B3") = "=code128(""" & aaa & """,B3,,230,)"
                .Range("B3") = aaa
                .Range("B3").Font.ColorIndex = 2
            End With
            
            ' Sheets("erp").Activate
            '=======================================================================================

            
            wb.Sheets("erp").Cells(jch + 2, 1).Resize(r, 4).Copy sht.[B7] '图号及数量复制
            wb.Sheets("erp").Cells(jch + 2, 7).Resize(r, 4).Copy sht.[F7] '图纸编号及分区
            
            wb.Sheets("erp").Cells(jch + 2, 15).Resize(r, 2).Copy
            sht.[J7].PasteSpecial Paste:=xlPasteValues  '备注的复制
            
            Quyu = wb.Sheets("erp").Range("N" & jch + 2)
            xch = wb.Sheets("erp").Range("L" & jch + 2)
            dingchi = wb.Sheets("erp").Range("M" & jch + 2)
            
            If Left(Quyu, 2) = "TP" Then sht.Range("J7:J" & (6 + r)) = "带配件"
            
            ' sht.Activate
            
            sht.[G2] = Quyu
            sht.[B4:E4] = xch  '型材截面号的输入
            sht.[G4] = dingchi '定尺输入

            wb.Sheets("计算用表").Range("B2:C31").ClearContents
            
            ' Sheets("erp").Activate
                
            wb.Sheets("erp").Cells(jch + 2, 4).Resize(r, 1).Copy wb.Sheets("计算用表").[c2]
            wb.Sheets("erp").Cells(jch + 2, 6).Resize(r, 1).Copy wb.Sheets("计算用表").[B2]
            
            ' Sheets("计算用表").Activate
                
            wb.Sheets("计算用表").[f1] = dingchi
                
            Call ZYouHua(wb)
            
            If wb.Sheets("计算用表").[f21] = 0 Then
                sht.[I4:K4] = 1
            Else
                wb.Sheets("计算用表").[f21].Copy sht.[I4：K4]
            End If
                
            ' Sheets("erp").Activate
        End If
    Next

    ' Sheet5.Activate

    Dim xh As Long
    xh = 0
    For PE = 1 To wb.Sheets.count
        ' If InStr(Sheets(PE).[A1], "模板") > 0 Then
        If InStr(wb.Sheets(PE).Name, "临时表") > 0 Then
            xh = xh + 1
            wb.Sheets("ZXD").Cells(xh + 3, 1) = xh
            wb.Sheets("ZXD").Cells(xh + 3, 2) = wb.Sheets("临时表" & xh).Cells(3, 2)
            wb.Sheets("ZXD").Cells(xh + 3, 3) = wb.Sheets("临时表" & xh).Cells(3, 9)
            'wb.Sheets("ZXD").Cells(xh + 3, 10) = wb.Sheets("临时表" & xh).Cells(4, 2)
            'wb.Sheets("ZXD").Cells(xh + 3, 13) = wb.Sheets("临时表" & xh).Cells(4, 6)
            'wb.Sheets("ZXD").Cells(xh + 3, 11) = wb.Sheets("临时表" & xh).Cells(4, 7)
            'wb.Sheets("ZXD").Cells(xh + 3, 12) = wb.Sheets("临时表" & xh).Cells(4, 9)
            
            wb.Sheets("临时表" & xh).Name = wb.Sheets("临时表" & xh).Cells(3, 2)
        
            If wb.Sheets("ZXD").Cells(xh + 3, 2) = "" Then
                wb.Sheets("ZXD").Cells(xh + 3, 3) = ""
                wb.Sheets("ZXD").Cells(xh + 3, 10) = ""
                wb.Sheets("ZXD").Cells(xh + 3, 11) = ""
                wb.Sheets("ZXD").Cells(xh + 3, 12) = ""
            End If
        End If
    Next
   
    ' wb.Sheets("erp").Activate
    
    wb.Sheets("erp").Columns("B:C").Delete
    wb.Sheets("erp").Columns("F:F").Delete
    
    wb.Sheets("erp").Columns("A:A").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    
    '=======================================================================================
    '191203按模板厂需求增加erp工作表整理导出环节
    Dim brr
    
    Dim t As Integer
    Dim enderp As Integer
    Dim scdh As String
    Dim erpfzl As String
    Dim erpjs As Integer
    
    Dim qcbm As String
    
    ' Sheets("erp").Copy after:=Sheets("erp")
    ' ActiveSheet.Name = ("erp库")
    wb.Sheets.Add(after:=wb.Sheets("erp")).Name = "erp库"
    wb.Sheets("erp").Cells.Copy wb.Sheets("erp库").[A1]
    
    ' Sheets("erp库").Activate
    
    enderp = wb.Sheets("erp库").Range("B65536").End(xlUp).Row
    For t = 1 To enderp
        If Len(wb.Sheets("erp库").Range("B" & t)) = 0 Then
            scdh = wb.Sheets("erp库").Range("A" & t)
        Else
            wb.Sheets("erp库").Range("L" & t) = scdh
            wb.Sheets("erp库").Range("M" & t) = scdh & wb.Sheets("erp库").Range("K" & t) & wb.Sheets("erp库").Range("G" & t)
        End If
    Next t
    
    wb.Sheets("erp库").Columns("B:B").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    wb.Sheets("erp库").Columns("L:L").Cut
    wb.Sheets("erp库").Columns("A:A").Insert Shift:=xlToRight
    wb.Sheets("erp库").Columns("B:B").Delete
    wb.Sheets("erp库").Columns("C:F").Delete
    wb.Sheets("erp库").Columns("D:F").Delete
    wb.Sheets("erp库").Columns("B:B").Cut
    wb.Sheets("erp库").Columns("E:E").Insert Shift:=xlToRight
    wb.Sheets("erp库").Columns("C:C").Cut
    wb.Sheets("erp库").Columns("B:B").Insert Shift:=xlToRight
    
    wb.Sheets("erp库").Rows(1).Insert
    brr = Array("生产单号", "区域简写", "生产单类型", "支数", "参考列", "支数合计")
    wb.Sheets("erp库").[A1].Resize(1, UBound(brr) + 1) = brr
    
    enderp = wb.Sheets("erp库").Range("B65536").End(xlUp).Row
    
    wb.Sheets("erp库").Columns("A:F").EntireColumn.AutoFit
    wb.Sheets("erp库").Range("A1:F" & enderp).HorizontalAlignment = xlCenter
    wb.Sheets("erp库").Range("A1:F" & enderp).Borders.LineStyle = xlContinuous
    
    With wb.Sheets("erp库").Sort.SortFields
        .Clear
        .Add Key:=wb.Sheets("erp库").Range("A2"), Order:=1
        .Add Key:=wb.Sheets("erp库").Range("B2"), Order:=1
        .Add Key:=wb.Sheets("erp库").Range("C2"), Order:=1
    End With
    
    With wb.Sheets("erp库").Sort
        .SetRange wb.Sheets("erp库").Range("A2:E" & enderp)
        .Header = 2
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    erpfzl = wb.Sheets("erp库").Range("E2")
    erpjs = 0
    
    For t = 2 To enderp + 1
        If wb.Sheets("erp库").Range("E" & t) <> erpfzl Then
            wb.Sheets("erp库").Range("F" & t - 1) = erpjs
            erpfzl = wb.Sheets("erp库").Range("E" & t)
            erpjs = wb.Sheets("erp库").Range("D" & t)
        Else
            erpjs = erpjs + wb.Sheets("erp库").Range("D" & t)
        End If
    Next t
    
    wb.Sheets("erp库").Columns("F:F").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    wb.Sheets("erp库").Columns("D:E").Delete
    ' ActiveWindow.ScrollRow = 1
    wb.Sheets("erp库").ScrollRow = 1
    ' TODO: 上面还有一个函数没有搞，下面这个另存为暂时还没看懂，分不开
    ' qcbm = Replace(wb.Name, "生产单", "库数据")
    
    ' Worksheets(Array("erp库")).Copy
    ' ActiveWorkbook.SaveAs FileName:=ThisWorkbook.path & "\" & qcbm, FileFormat:=51
    ' ActiveWorkbook.Close SaveChanges:=True
    
    If fileIsExist(sjkFilename) Then
        Kill sjkFilename
    End If
    Call createExcel(sjkFilename)
    Dim wbSjk As Workbook
    Set wbSjk = Workbooks.Open(sjkFilename)
    wbSjk.Windows(1).Visible = False
    ThisWorkbook.Activate
    If isSheetExist(wbSjk, "erp库") Then
        wbSjk.Sheets("erp库").Delete
    End If
    wbSjk.Sheets.Add().Name = "erp库"
    wb.Sheets("erp库").Cells.Copy wbSjk.Sheets("erp库").[A1]
    wbSjk.Windows(1).Visible = True
    wbSjk.Close (True)

    Application.DisplayAlerts = False
    
    wb.Sheets("erp库").Delete
    
    Application.DisplayAlerts = True
    
    '=======================================================================================

    
    ' wb.Sheets("ZXD").Activate
    Dim zxdzh
    
    zxdzh = wb.Sheets("ZXD").Range("B65536").End(xlUp).Row
    
    wb.Sheets("ZXD").Columns("A:H").HorizontalAlignment = xlCenter
    wb.Sheets("ZXD").Range("B2").HorizontalAlignment = xlLeft
    Dim v, XHZ, JISHU1, crhj, HS
    For xh = 4 To zxdzh
        v = Split(wb.Sheets("ZXD").Range("B" & xh), "-")
        XHZ = Right(wb.Sheets("ZXD").Range("B" & xh), Len(v(UBound(v))))
        wb.Sheets("ZXD").Range("A" & xh) = XHZ
    Next
    For xh = 4 To zxdzh
        If wb.Sheets("ZXD").Range("A" & xh) = 1 Then
            JISHU1 = JISHU1 + 1
        End If
    Next
    
    If JISHU1 > 1 Then
        For crhj = 5 To zxdzh + (JISHU1 - 1) * 3 '插入合计
            If wb.Sheets("ZXD").Range("A" & crhj) = "1" Then
                wb.Sheets("ZXD").Rows(crhj & ":" & (crhj + 2)).Insert
                crhj = crhj + 3
            End If
        Next
        crhj = ""
        zxdzh = wb.Sheets("ZXD").Range("B65536").End(xlUp).Row
        For crhj = 5 To zxdzh
            If wb.Sheets("ZXD").Range("A" & crhj) = "1" Then
                wb.Sheets("ZXD").Range("B" & (crhj - 2)) = "合计"
                wb.Sheets("ZXD").Range("C" & (crhj - 2)) = Application.WorksheetFunction.Sum(wb.Sheets("ZXD").Range("C" & (crhj - 3) & ": C" & ((crhj - 3) - HS + 2)))
                HS = 1
            Else
                HS = HS + 1
            End If
        Next
        
        wb.Sheets("ZXD").Range("B" & zxdzh + 2) = "合计"
        wb.Sheets("ZXD").Range("C" & zxdzh + 2) = Application.WorksheetFunction.Sum(wb.Sheets("ZXD").Range("C" & zxdzh & ": C" & (zxdzh - HS + 1)))
        Dim i
        i = "'"
        For i = 4 To zxdzh + 2
            If wb.Sheets("ZXD").Range("B" & i) = "合计" Then
                wb.Sheets("ZXD").HPageBreaks.Add wb.Sheets("ZXD").Range("b" & i + 2)
            End If
        Next
    Else
        wb.Sheets("ZXD").Range("B" & zxdzh + 2) = "合计"
        wb.Sheets("ZXD").Range("C" & zxdzh + 2) = WorksheetFunction.Sum(wb.Sheets("ZXD").Range("C" & zxdzh & ": C4"))
    End If
    
    wb.Sheets("ZXD").Range("A" & zxdzh + 1 & ":A" & (zxdzh + 100)).ClearContents
    wb.Sheets("ZXD").PageSetup.PrintArea = "$A$1:$H$" & zxdzh + 2
    wb.Sheets("ZXD").PageSetup.PrintTitleRows = "$1:$3"
    wb.Sheets("ZXD").Range("A4:H" & zxdzh + 2).Borders.Weight = 2
    'wb.Sheets("ZXD").Range("A4:H" & ZXDZH + 2).BorderAround , 3
    wb.Sheets("ZXD").Rows("4:" & zxdzh + 2).RowHeight = 20
    
    wb.Sheets("ZXD").Columns("J:J").EntireColumn.AutoFit
    
    Application.ScreenUpdating = True
    
    MsgBox ("备料单中型材须拆分为一个单元格只有一种型材型号")
    wb.Windows(1).Visible = True
    wb.Close (True)
End Sub

Sub ZYouHua(wb As Workbook) '优化，感谢大神
    '准备工作与转存数组
    wb.Sheets("计算用表").Range("f12:f16,f21:f24").ClearContents
    wb.Sheets("计算用表").Range("i2:i" & Rows.count).ClearContents
    wb.Sheets("计算用表").Range("k2:CH" & Rows.count).ClearContents
    Dim m%, n&, arr, i&, j&, k&, ylcd#, xlzd#, js&, gd#
    ylcd = wb.Sheets("计算用表").Range("f1").Value '整料
    xlzd = wb.Sheets("计算用表").Range("f3").Value '短料

    m = wb.Sheets("计算用表").Range("f4").Value '种数
    If m = 0 Then MsgBox "请输入下料长度与数量！", , "友情提示"
    n = wb.Sheets("计算用表").Range("f5").Value '根数: Exit Sub
    ReDim yssj#(1 To n), crr#(1 To m, 1 To 2) '原始数据
    arr = wb.Sheets("计算用表").Range("b2:c100").Value
    k = 0 '计数初始化
    For i = 1 To 99 '排除arr中空值，形成crr数组
        If arr(i, 1) <> "" And arr(i, 2) <> "" Then
            k = k + 1
            crr(k, 1) = arr(i, 1): crr(k, 2) = arr(i, 2)
        End If
    Next i
    k = 0
    For i = 1 To m '转存一维数组
        For j = 1 To crr(i, 2)
            k = k + 1
            yssj(k) = crr(i, 1)
        Next j
    Next i

    '排除整料与准整料（即余料小于短料）
    ReDim jg11(1 To 1, 1 To m), jg12(1 To 1, 1 To m) '结果1
    js = 0 '输出第一部分结果时的行数
    For i = 1 To m
        If crr(i, 1) > ylcd - xlzd Then
            js = js + 1
            jg11(1, js) = crr(i, 1)
            jg12(1, js) = crr(i, 2)
        End If
    Next i
    If js > 0 Then
        ReDim Preserve jg11(1 To 1, 1 To js), jg12(1 To 1, 1 To js)
        wb.Sheets("计算用表").Range("i2").Resize(js) = WorksheetFunction.Transpose(jg12) '输出根数
        wb.Sheets("计算用表").Range("k2").Resize(js) = WorksheetFunction.Transpose(jg11) '输出规格
    End If

    '整理原始数据
    Dim n1& '需搭配的下料根数
    n1 = n - WorksheetFunction.Sum(jg12)
    If n1 = 0 Then MsgBox "下料长度不能搭配，不需要优化！", , "友情提示": Exit Sub
    ReDim SJ#(1 To n1), sj1#(1 To n1), sj2$(1 To n1), jg$(1 To n1)
    k = 0
    For i = 1 To n
        If yssj(i) <= ylcd - xlzd Then
            k = k + 1
            SJ(k) = yssj(i) '排除整料和准整料后的原始数据
            sj1(k) = yssj(i) '用于搭配求和
        End If
    Next i
    wb.Sheets("计算用表").Range("f12") = WorksheetFunction.Max(SJ) '搭配料最长
    wb.Sheets("计算用表").Range("f13") = WorksheetFunction.Min(SJ) '搭配料最短
    wb.Sheets("计算用表").Range("f14") = m - js '搭配料种数
    wb.Sheets("计算用表").Range("f15") = n1 '搭配料根数
    wb.Sheets("计算用表").Range("f16") = WorksheetFunction.Sum(SJ) '搭配料总长
    'If n1 = n Then
    '    MsgBox "没有整料和准整料！" & Chr(10) & Chr(10) & "下面即将开始可以搭配下料的组合优化，请耐心等待！", , "友情提示"
    'Else
    '    MsgBox "不能搭配下料的整料和准整料已排除完毕！" & Chr(10) & Chr(10) & "下面即将开始可以搭配下料的组合优化，请耐心等待！", , "友情提示"
    'End If

    '需搭配的下料组合优化（★★★核心代码★★★）
    Dim t#, pd As Boolean, yl#, minyl#, r&, h%, l%, brr, sxgs&
    Dim ylgs&, minylgs&, ylfc#, ylfc1#, maxylfc#, sjcs&, ii&, i3&, sjxb&, dapaizhi# '随机次数、随机下标、搭配值
    t = Timer
    minyl = wb.Sheets("计算用表").Range("f16").Value '最小余料
    sxgs = 0 '实现根数
    minylgs = n1 '最小余料根数
    maxylfc = 0 '最大余料方差
    sjcs = wb.Sheets("计算用表").Range("f27").Value * 1000
    Randomize
    For ii = 1 To sjcs
        For k = 1 To n1 - 1
            If sj1(k) > 0 Then '跳过已搭配的料
                dapaizhi = sj1(k) '当前搭配值
                sj2(k) = sj1(k) '搭配组合连接初值
100:
                ReDim sj3#(1 To n1)
                i3 = 0
                For i = k + 1 To n1
                    If sj1(i) > 0 And sj1(i) <= ylcd - dapaizhi Then '筛选出当前dapaizhi的可搭配料存入sj3，sj3的元素可变
                        i3 = i3 + 1
                        sj3(i3) = sj1(i)
                    End If
                Next i
                If i3 > 0 Then '存在可搭配料时
                    ReDim Preserve sj3#(1 To i3)
                    sjxb = Int(Rnd() * i3) + 1 '★随机搭配可搭配的下料规格★
                    dapaizhi = dapaizhi + sj3(sjxb)
                    sj1(k) = dapaizhi
                    sj2(k) = sj2(k) & "+" & sj3(sjxb)
                    For i = k + 1 To n1
                        If sj1(i) = sj3(sjxb) Then sj1(i) = 0: Exit For '已搭配料归0
                    Next i
                    GoTo 100
                End If
            End If
        Next k
        If sj1(n1) > 0 Then sj2(n1) = sj1(n1) '防止最后一个值未被搭配时遗漏
        yl = 0: ylgs = 0: ylfc = 0
        For i = 1 To n1
            If sj1(i) <> 0 Then
                yl = yl + (ylcd - sj1(i)) '计算余料
                If sj1(i) < ylcd Then ylgs = ylgs + 1 '计算余料根数（搭配后小于整料长度）
                ylfc = ylfc + (ylcd - sj1(i)) ^ 2 '计算余料方差
            End If
        Next i
        If yl <= minyl And ylgs <= minylgs Then '记录相对较优方案：余料少、余料根数少
            If ylfc > maxylfc Then '余料方差大
                For i = 1 To n1 '结果出口1
                    jg(i) = sj2(i)
                Next i
                ylfc1 = ylfc '用于正确输出结果
                pd = True
            End If
            If ylgs < minylgs Then
                For i = 1 To n1 '结果出口2
                    jg(i) = sj2(i)
                Next i
                maxylfc = ylfc '★余料根数减少时重新记录最大余料方差★
                pd = False
            End If
            minyl = yl: minylgs = ylgs
        End If
        Randomize
        For i = 1 To n1 '利用【香川经典数组洗牌法】乱序恢复sj1数组数据，准备下一次随机搭配
            r = Int(Rnd() * (n1 - i + 1)) + i
            gd = SJ(r): SJ(r) = SJ(i): SJ(i) = gd
            sj1(i) = SJ(i): sj2(i) = ""
        Next i
    Next ii

    '
    ''输出结果至工作表
    wb.Sheets("计算用表").Range("f22").Value = minyl
    wb.Sheets("计算用表").Range("f23").Value = minylgs: If minylgs = 0 Then minylgs = 1
    wb.Sheets("计算用表").Range("f24").Value = Format(Sqr(IIf(pd, ylfc1, maxylfc)) / minylgs, "0.00")
    ReDim jgjs(1 To n1) '结果计数
    For i = 1 To n1 '对结果字串中连接的下料规格做从大到小排序处理
        If jg(i) <> "" Then
            brr = Split(jg(i), "+")
            For j = 0 To UBound(brr) - 1
                For k = j + 1 To UBound(brr)
                    If Val(brr(j)) < Val(brr(k)) Then gd = brr(k): brr(k) = brr(j): brr(j) = gd
                Next k
            Next j
            jg(i) = ""
            For j = 0 To UBound(brr)
                jg(i) = jg(i) & "+" & brr(j)
            Next j
        End If
    Next i
    For i = 1 To n1 - 1 '去重并计数
        If jg(i) <> "" Then
            For j = i + 1 To n1
                If jg(j) = jg(i) Then jgjs(i) = jgjs(i) + 1: jg(j) = ""
            Next j
        End If
    Next i
    h = js + 3 '行
    For i = 1 To n1
        If jg(i) <> "" Then
            Cells(h, "i").Value = jgjs(i) + 1
            sxgs = sxgs + jgjs(i) + 1
            brr = Split(jg(i), "+")
            l = 11 'K列
            For j = 1 To UBound(brr)
                Cells(h, l).Value = brr(j)
                l = l + 1
            Next j
            h = h + 1
        End If
    Next i
    wb.Sheets("计算用表").Range("f21").Value = sxgs
    ''MsgBox "用时：" & Format(Timer - t, "0.0000") & "秒。", , "友情提示"
End Sub

' 生成ERP
Sub FB5(scqdFilename As String)
    Dim wb As Workbook
    Set wb = Workbooks.Open(scqdFilename)
    'wb.Windows(1).Visible = False
    ' ThisWorkbook.Activate

    Application.ScreenUpdating = False
    ' wb.Sheets("erp").Activate
    Dim endc
    endc = wb.Sheets("erp").[C65536].End(xlUp).Row '对C列的最后一行进行定位，在拆分明细的库里找一下C列的宽度，如果没有就显示宽度为红色
    Dim i, xck, k
    For i = 2 To endc
        If Len(wb.Sheets("erp").Range("C" & i)) > 0 Then
            xck = wb.Sheets("erp").Cells(i, 3) '型材宽度
            If wb.Sheets("erp").Sheets("库(待补充)").Columns(6).Find(xck, LookAt:=xlWhole, SearchDirection:=xlPrevious) Is Nothing Then
                wb.Sheets("erp").Range("C" & i).Interior.Color = RGB(230, 100, 100) '
                k = k + 1 '如果K大于零，则需要在库里加新的拆分明细表
            End If
        End If
    Next
    
    If k > 0 Then
        MsgBox ("需在库中添加明细后，重新生成erp")
        GoTo 100
    End If
    
    If isSheetExist(wb, "拆分明细") Then
        wb.Sheets("拆分明细").Delete
    End If
    wb.Sheets.Add().Name = "拆分明细"
    wb.Sheets("erp").Cells.Copy wb.Sheets("拆分明细").[A1]
    
    wb.Sheets("拆分明细").Columns("E:M").Delete
    
    wb.Sheets("拆分明细").Columns("b:b").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    wb.Sheets("拆分明细").Range("A1:D" & wb.Sheets("拆分明细").UsedRange.Rows.count).Borders.LineStyle = xlContinuous
    Dim BKTOP
    For BKTOP = 1 To wb.Sheets("拆分明细").UsedRange.Rows.count '设置行的上边框
        wb.Sheets("拆分明细").Range("A" & BKTOP & ":D" & BKTOP).Borders(xlEdgeTop).Weight = xlMedium
    Next BKTOP

'    Columns("A:A").Replace What:=" ", Replacement:=""
    wb.Sheets("拆分明细").Columns("D:E").Insert
    Dim kuandu, hangshu, r, m
    For i = 1 To wb.Sheets("拆分明细").UsedRange.Rows.count * 10
        On Error Resume Next
        kuandu = wb.Sheets("拆分明细").Cells(i, 3)
        If kuandu = "" Then
            Exit For
        End If
        hangshu = ThisWorkbook.Sheets("库(待补充)").Columns(6).Find(kuandu, LookAt:=xlWhole, SearchDirection:=xlPrevious).Row
        ' ThisWorkbook.Sheets("库(待补充)").Activate
        ' ThisWorkbook.Sheets("库(待补充)").Range("F" & hangshu).Resize(, 12).Select
        ' r = Selection.Rows.Count
        r = ThisWorkbook.Sheets("库(待补充)").Range("F" & hangshu).Resize(, 12).Rows.count
        ' Sheets("拆分明细").Activate
        wb.Sheets("拆分明细").Rows(i + 1 & ":" & i + r).Insert
        m = m + 1
        ThisWorkbook.Sheets("库(待补充)").Range("F" & hangshu).Resize(r, 12).Copy wb.Sheets("拆分明细").Cells(i + 1, 1)
        wb.Sheets("拆分明细").Cells(i, 3).ClearContents
        ' Range("A" & i & ":F" & i).Copy
        ' Range("B" & i + 1).PasteSpecial SkipBlanks:=True
        wb.Sheets("拆分明细").Range("B" & i + 1) = wb.Sheets("拆分明细").Range("A" & i & ":F" & i).Value
        wb.Sheets("拆分明细").Rows(i).ClearContents
        wb.Sheets("拆分明细").Cells(i + 1, 1) = m
        i = i + r
    Next i
    
    ' Columns("F:F").Select
    ' Selection.Copy
    ' Selection.PasteSpecial Paste:=xlPasteValues
    
    wb.Sheets("拆分明细").Columns("F:F") = wb.Sheets("拆分明细").Columns("F:F").Value

    wb.Sheets("拆分明细").Columns("F:F").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    Dim j
    For j = wb.Sheets("拆分明细").UsedRange.Rows.count To 1 Step -1
        If wb.Sheets("拆分明细").Cells(j, 6).Text = "" Then
            wb.Sheets("拆分明细").Rows(j).EntireRow.Delete
        End If
    Next j
    Dim arr
    wb.Sheets("拆分明细").Rows("1:1").Insert
    arr = Array("序号", "模板编号", "数量", "图纸名称", "型材截面", "材质", "长度", "数量", "总数量", "理论重量", "总重kg", "型材类型")
    wb.Sheets("拆分明细").[A1].Resize(1, UBound(arr) + 1) = arr
    '设置表头的格式
    With wb.Sheets("拆分明细").Range("A1:L1")
        .HorizontalAlignment = xlCenter
        .Borders.Weight = 2
        .BorderAround , 3
    End With
    '设置宽度自适应，然后调整比较窄的列，列宽为8
    wb.Sheets("拆分明细").Columns("A:L").EntireColumn.AutoFit
    wb.Sheets("拆分明细").Range("A:A,C:C,H:H,I:I").ColumnWidth = 8
    
    wb.Sheets("拆分明细").Rows("1:1").RowHeight = 25
    Call FB6(wb)
100:
    Application.ScreenUpdating = True
    wb.Windows(1).Visible = True
    wb.Close (True)
End Sub

' 如果修改明细重新算配件
Sub FB6(wb As Workbook)
    ' wb.Sheets("拆分明细").Activate
    Application.ScreenUpdating = False
    
    ' wb.Sheets("拆分明细").Columns("E:L").Select
    wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        wb.Sheets("拆分明细").Range("E1:L65535"), Version:=xlPivotTableVersion14).CreatePivotTable _
        TableDestination:=wb.Sheets("拆分明细").Range("N1"), TableName:="数据透视表1", DefaultVersion:= _
        xlPivotTableVersion14
    With wb.Sheets("拆分明细").PivotTables("数据透视表1").PivotFields("型材类型")
        .Orientation = xlPageField
        .Position = 1
    End With
    With wb.Sheets("拆分明细").PivotTables("数据透视表1").PivotFields("型材截面")
        .Orientation = xlRowField
        .Position = 1
    End With
    With wb.Sheets("拆分明细").PivotTables("数据透视表1").PivotFields("长度")
        .Orientation = xlRowField
        .Position = 2
    End With
    wb.Sheets("拆分明细").PivotTables("数据透视表1").AddDataField wb.Sheets("拆分明细").PivotTables("数据透视表1" _
        ).PivotFields("总数量"), "总计数量", xlSum
        
    wb.Sheets("拆分明细").PivotTables("数据透视表1").RowAxisLayout xlTabularRow
'    wb.Sheets("拆分明细").PivotTables("数据透视表1").RepeatAllLabels xlRepeatLabels
    Dim p As PivotField
    For Each p In wb.Sheets("拆分明细").PivotTables("数据透视表1").PivotFields
        p.Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
    Next
        
    With wb.Sheets("拆分明细").PivotTables("数据透视表1").PivotFields("型材类型")
        .PivotItems("").Visible = False
'        .PivotItems("主板").Visible = False
        .PivotItems("(blank)").Visible = False
    End With
    '对行和列禁用汇总
    With wb.Sheets("拆分明细").PivotTables("数据透视表1")
        .ColumnGrand = False
        .RowGrand = False
    End With
    '对型材截面进行粘贴为值，对应库找到型材对应的定尺，然后用计算用表获取支数
    ' Columns("N:N").Copy
    ' Columns("R:R").PasteSpecial Paste:=xlPasteValues
    Dim endn
    endn = wb.Sheets("拆分明细").Range("N65535").End(xlUp).Row
    wb.Sheets("拆分明细").Range("R1:R" & endn) = wb.Sheets("拆分明细").Range("N1:N" & endn).Value

    wb.Sheets("拆分明细").Range("r1") = ""
    wb.Sheets("拆分明细").Range("r3") = "配件型材截面"
    wb.Sheets("拆分明细").Columns("R:R").SpecialCells(xlCellTypeBlanks).Delete Shift:=xlUp
    wb.Sheets("拆分明细").PivotTables("数据透视表1").RepeatAllLabels xlRepeatLabels
    
    endr = wb.Sheets("拆分明细").Range("R5000").End(xlUp).Row
    For i = 2 To endr
        xcjm = wb.Sheets("拆分明细").Range("r" & i) '型材截面
        If ThisWorkbook.Sheets("库(待补充)").Columns(2).Find(xcjm, LookAt:=xlWhole, SearchDirection:=xlprerious) Is Nothing Then
            dingchi = "6000"
        Else
            hangshu = ThisWorkbook.Sheets("库(待补充)").Columns(2).Find(xcjm, LookAt:=xlWhole, SearchDirection:=xlprerious).Row
            dingchi = ThisWorkbook.Sheets("库(待补充)").Range("C" & hangshu)
        End If
        wb.Sheets("拆分明细").Range("s" & i) = dingchi
        
        hangshumin = Columns("N:N").Find(xcjm, LookAt:=xlWhole, SearchDirection:=xlprerious).Row
        wb.Sheets("拆分明细").Range("T" & i) = hangshumin
    Next
    wb.Sheets("拆分明细").Range("s1") = "定尺"
   
    wb.Sheets("计算用表").Range("B2:C100").ClearContents
    wb.Sheets("拆分明细").Activate
    endn = wb.Sheets("拆分明细").Range("N5000").End(xlUp).Row
    endr = wb.Sheets("拆分明细").Range("R5000").End(xlUp).Row
    For i = 2 To endr
        ' wb.Sheets("拆分明细").Activate
        py = wb.Sheets("拆分明细").Range("T" & i) '偏移起始单元格
        If i < endr Then
            pyfw = wb.Sheets("拆分明细").Range("T" & (i + 1)) - wb.Sheets("拆分明细").Range("T" & i)
        Else
            pyfw = endn + 1 - wb.Sheets("拆分明细").Range("T" & i)
        End If
        
        wb.Sheets("拆分明细").Range("O" & py).Resize(pyfw, 2).Copy wb.Sheets("计算用表").[B2]
        wb.Sheets("计算用表").[f1] = wb.Sheets("拆分明细").Range("S" & i)
        wb.Sheets("计算用表").Activate
        Call ZYouHua(wb)
        If wb.Sheets("计算用表").[f21] = 0 Then
            wb.Sheets("拆分明细").Range("U" & i) = 1
        Else
            wb.Sheets("计算用表").[f21].Copy wb.Sheets("拆分明细").Range("U" & i)
        End If
        wb.Sheets("计算用表").Range("B2:C100").ClearContents
    Next
    ' wb.Sheets("拆分明细").Activate
    wb.Sheets("拆分明细").Range("U1") = "支数"
    wb.Sheets("拆分明细").Columns("T:T").Delete
    wb.Sheets("拆分明细").Columns("N:P").Delete
    
    With wb.Sheets("拆分明细").Range("O1:Q" & endr)
        .HorizontalAlignment = xlCenter
        .Borders.Weight = 2
    End With
    wb.Sheets("拆分明细").Columns("O:O").EntireColumn.AutoFit
    For i = 2 To endr
        If InStr(wb.Sheets("拆分明细").Range("O" & i).Text, "板材") > 0 Then
            wb.Sheets("拆分明细").Range("O" & i).Interior.Color = RGB(230, 100, 100)
            bcsl = bcsl + 1
        End If
    Next
    If bcsl > 0 Then MsgBox ("辅料用到板材，将数量按照面积换算成小数")
    Application.ScreenUpdating = True
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