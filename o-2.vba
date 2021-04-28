
'Callback for an1 onAction
Sub Marco(control As IRibbonControl)
    Select Case control.ID
        Case "anfb1"
            Call FBA显示信息栏
        Case "anfb2"
            Call FBB合并同类模板并分类
        Case "anfb3"
            Call FBC填充单号
        Case "anfb4"
            Call FBD转序单生产单
        Case "anfb5"
            Call FBE上传
        Case "anfb6"
            Call FBF生成ERP
        Case "anfb7"
            Call FBG如果修改明细重新算配件
        Case "anbz1"
            Call BZA显示信息栏
        Case "anbz2"
        
            Call BZB修改后缀后运行
        Case "anbz3"
        
            Call BZC插入空白行
         Case "anbz4"
            Call BZD生成转序单
        Case "anhz1"
            Call 非标35种及标准件汇总
        Case "anhz2"
            Call 备料型材汇总
            
        Case "anhz3"
            Call A1合并设计清单
        Case "anhz4"
            Call A2分出标准件非标件
            
            
        Case "anhz5"
            Call A3拆分到工作簿
            
        Case "antj1"
            Call Z1合并设计清单
                
        Case "antj2"
            Call Z2标准件数据统计
        
        Case "antj3"
            Call Z2修改后缀后运行
        Case "antj4"
            Call Z3_BK吊架非标数据统计
            
            
            
    End Select


End Sub


Sub FBA显示信息栏()
    
    FBJ.Show
'    txtgyxm.Text = lastgyxm: txtgydh.Text = lastgydh: txtshxm.Text = lastshxm
    
End Sub
Sub FBB合并同类模板并分类()

    Application.ScreenUpdating = False
    
    On Error Resume Next
    
    Columns("C:C").Replace "（", "("
    Columns("C:C").Replace "）", ")"
    Columns("C:F").Replace " ", ""
    
    '单元格匹配替换W2的0值
    Columns("F:F").Replace What:="0", Replacement:="", LookAt:=xlWhole, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    
    Range("A1") = "1"
    
    endb = [b60000].End(xlUp).Row
    Range("A1").AutoFill Destination:=Range("A1:A" & endb), Type:=xlFillSeries
   
    Columns("H:H").NumberFormatLocal = "G/通用格式"
    
    For i = 1 To endb
        
        If InStr(Range("H" & i), "-N-") > 0 Then Range("B" & i) = "转角"
        If InStr(Range("H" & i), "ZW-Q-") > 0 Then Range("B" & i) = "墙柱C槽"
        
    Next
    
    '对现有内容进行排序
    With Sheets("erp").Sort.SortFields
        
        .Clear
        .Add Key:=Range("B2"), Order:=1 '模板名称
        .Add Key:=Range("E2"), Order:=1 'W1
        .Add Key:=Range("F2"), Order:=1 'W2
        .Add Key:=Range("H2"), Order:=1 '图纸编号
        .Add Key:=Range("G2"), Order:=1 'L
        
    End With
    
    With Sheets("erp").Sort
        
        .SetRange Range("b2:M" & endb)
        .Header = 2 '没有标题
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
        
    End With
    
    For i = 1 To endb
        
        Range("L" & i) = Range("G" & i) & ""  '做宽度连接
        Range("K" & i) = Range("E" & i) & Range("F" & i)  '做宽度连接
        Range("M" & i) = Range("B" & i) & "&" & Range("E" & i) & Range("F" & i) & Range("J" & i)
        
    Next
    
    
    '宽度连接为数值
    Range("E:G").Delete Shift:=xlToLeft '删除以前的两列宽度及单件面积和总面积列
    Columns("H:I").Cut
    Columns("E:E").Insert Shift:=xlToRight
    Columns("A:J").Select

    [k1] = 1
    Range("K2").FormulaR1C1 = "=IF(RC[-1]<>R[-1]C[-1],R[-1]C+1,R[-1]C)"
    Range("K2").AutoFill Destination:=Range("K2:K" & endb)
    
    zhz = Sheets("erp").Range("K" & endb).Value
    
    Columns("K:K").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    
    For jia = 1 To zhz
            
        Range("K" & endb + jia) = jia
        
    Next
    
    ActiveWorkbook.Worksheets("erp").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("erp").Sort.SortFields.Add Key:=Range("K1:K" & endb + zhz), _
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
    xlSortTextAsNumbers
    
    With ActiveWorkbook.Worksheets("erp").Sort
    
        .SetRange Range("A1:k" & endb + zhz)
        .Header = 2
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    
    End With
    
    Rows("1:1").Insert Shift:=xlDown
    Columns("J:K").Delete Shift:=xlToLeft
    Columns("A:A").Delete Shift:=xlToLeft
    
    Range("A1:H" & endb + zhz).Borders.LineStyle = xlContinuous
    
    Columns("A:A").Cut
    Columns("H:H").Insert Shift:=xlToRight
    
    '把型材截面号放到i列,定尺放到H列
    zjh = Range("A6000").End(xlUp).Row
    'Columns("C:C").Replace " ", ""
    
    For xcki = 2 To zjh '表示型材宽度的计数
        
        xch = ""
        dingchi = ""
        
        If Len(Range("C" & xcki)) <> 0 Then
            
            xck = Cells(xcki, 3) '型材宽度
            
            If Sheets("库(待补充)").Columns(1).Find(xck, LookAt:=xlWhole, SearchDirection:=xlprerious) Is Nothing Then
                
                xch = "请输入型材"
                dingchi = "输入定尺"
            
            Else
                
                hangshu = Sheets("库(待补充)").Columns(1).Find(xck, LookAt:=xlWhole, SearchDirection:=xlprerious).Row
                
                xch = Sheets("库(待补充)").Range("B" & hangshu) '型材截面号
                dingchi = Sheets("库(待补充)").Range("C" & hangshu)
                
            End If
            
            Range("I" & xcki) = xch
            Range("J" & xcki) = dingchi
            
            If xch = "请输入型材" Then
            
                Range("I" & xcki).Interior.Color = RGB(255, 20, 20)
                Range("J" & xcki).Interior.Color = RGB(255, 20, 20)
            
            End If
        
        End If
        
        If Range("G" & xcki) = "K板" Then
            
            If Left(Range("I" & xcki), 4) <> "YK-P" Or Range("I" & xcki) <> "ZWGYC-2370" Then
                
                Range("I" & xcki).Interior.Color = RGB(255, 20, 20)
            
            End If
        
        End If
        
        If InStr(Range("G" & xcki), "堵板") > 0 And Range("I" & xcki) <> "ZWGYC-3273" Then Range("I" & xcki).Interior.Color = RGB(255, 20, 20)
        
        '2018-11-02  新增,出现例如300D600 这样的情况,是需要用一代板的,这里增加一个判断提示
        If (Range("G" & xcki) = "平板" Or Range("G" & xcki) = "平面板" Or Range("G" & xcki) = "PK板") Then
            
            If InStr(Split(Range("A" & xcki), "-")(1), "D") > 1 Then
                
                Range("I" & xcki).Interior.Color = RGB(255, 20, 20)
            
            End If
        
        End If
        
        '1107 增加了5XC 5SC颜色
        If InStr(Range("A" & xcki), "XC") + InStr(Range("A" & xcki), "SC") > 1 Then
            
            Range("I" & xcki).Interior.Color = RGB(255, 20, 20)
            
        End If
        
        
        If InStr(Range("G" & xcki), "龙骨") + InStr(Range("G" & xcki), "铝梁") > 0 And Range("I" & xcki) <> "HLD-31" Then Range("I" & xcki).Interior.Color = RGB(255, 20, 20)
    
    Next
    
    endb = [B65536].End(xlUp).Row
    Slhj = Application.WorksheetFunction.Sum(Range("b1:b" & endb)) '数量合计
    
    Range("A1:J" & endb).Borders.LineStyle = xlContinuous
    
    dys = 1
    
    
      Range("T1:T1").Select
       
     ActiveCell.Value = "=RIGHT((MID(A1, FIND(""K"",A1)-4,4)),LEN((MID(A1, FIND(""K"",A1)-4,4)))-FIND(""-"",(MID(A1, FIND(""K"",A1)-4,4))))"
      
      Range("T1:T6000").Select
      Selection.FillDown
    
    
    
    
    Do While dys <= endb
    
        If Range("B" & dys) > 0 Then
            
            k = k + Range("B" & dys)
            kk = kk + 1
        
        Else
            
            k = 0
            kk = 0
        
        End If
        
        If k >= 50 And kk < 25 Then
            
            If Range("B" & dys - 1) > 0 Then
            
                Rows(dys).Insert
                k = 0
                kk = 0
                
                endb = endb + 1
                
            Else
            
                k = 0
                kk = 0
        
            End If
            
        ElseIf k < 50 And kk >= 25 Then
        
            If Range("B" & dys - 1) > 0 Then
            
                Rows(dys).Insert
                k = 0
                kk = 0
                
                endb = endb + 1
                
            Else
            
                k = 0
                kk = 0
        
            End If
            
        ElseIf k >= 50 And kk >= 25 Then
        
            If Range("B" & dys - 1) > 0 Then
            
                Rows(dys).Insert
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
    
    Columns("H:H").Cut
    Columns("K:K").Insert Shift:=xlToRight
    Columns("G:G").Insert Shift:=xlToRight
    
    Columns("A:J").HorizontalAlignment = xlCenter
    
    enda = Range("A6000").End(xlUp).Row
    
    '先用设计提供的模板名称进行判断
    For XCHI = 2 To enda '型材号计数i
    
        On Error Resume Next
    
        TZBH = Range("H" & XCHI) '图纸编号
        
        

 
 '背孔打筋判断-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
 
      MM = Range("U" & XCHI).Value
     ' MsgBox MM
      
  If (20 < MM) Then
         
    If InStr(TZBH, "平面板") Or InStr(TZBH, "平板") Or InStr(TZBH, "普板") Or InStr(TZBH, "墙板") Or InStr(TZBH, "PK板") > 0 Or InStr(TZBH, "平面板切斜") > 0 Then '判断打筋
        
        If Range("I" & XCHI) = "ZWGYC-3265" Then '100P
        
             If 45 <= MM Or MM <= 55 Then
             
                 Range("I" & XCHI) = "ZWGYC-3265"
                
            Else
            
                ' Range("I" & XCHI) = "YK-P002"
        
                 Range("I" & XCHI).Interior.Color = RGB(127, 255, 212) '青色
                
            
            End If
            
        End If
        
   
     
     
        If Range("I" & XCHI) = "ZWGYC-3266" Then '150P
        
             If 45 <= MM Or MM <= 105 Then
            
                Range("I" & XCHI) = "ZWGYC-3266"
               
            Else
            
              '  Range("I" & XCHI) = "YK-P003"
                
                Range("I" & XCHI).Interior.Color = RGB(127, 255, 212) '青色
            
                            
            End If
            
        End If
        
   
     
     
      If Range("I" & XCHI) = "ZWGYC-3267" Then '200P
        
             If 45 <= MM Or MM <= 155 Then
            
                Range("I" & XCHI) = "ZWGYC-3267"
            
            Else
            
              '  Range("I" & XCHI) = "YK-P004"
                 
                Range("I" & XCHI).Interior.Color = RGB(127, 255, 212) '青色
            
            End If
            
      End If
                          '-----------------------------------------------------------------------------------------------------
      
         If Range("I" & XCHI) = "ZWGYC-3268" Then '250P，W
         
             If InStr(Range("A" & XCHI), "W") > 0 Then   '铣槽
                
                    If (45 <= MM And MM <= 100) Or (150 <= MM And MM <= 205) Then
                
                         Range("I" & XCHI) = "ZWGYC-3888"
                         
                    ElseIf (50 <= MM And MM <= 85) Or (110 <= MM And MM <= 140) Or (165 <= MM And MM <= 200) Then '铣槽二代板
         
                         Range("I" & XCHI) = "ZWGYC-3268"
                    
                         Range("I" & XCHI).Interior.Color = RGB(255, 193, 37) '橙色
                    
                    Else
                
                         Range("I" & XCHI) = "YK-P005" '铣槽一代板
                    
                         Range("I" & XCHI).Interior.Color = RGB(127, 255, 212) '青色
                         
                         
                    End If
                    
             Else
             
                    If (50 <= MM And MM <= 85) Or (110 <= MM And MM <= 140) Or (165 <= MM And MM <= 200) Then '二代板
         
                         Range("I" & XCHI) = "ZWGYC-3268"
                    
                                         
                    Else
                
                       '  Range("I" & XCHI) = "YK-P005" '一代板
                    
                         Range("I" & XCHI).Interior.Color = RGB(127, 255, 212) '青色
                         
                    End If
                          
             End If
             
          End If
          
          
                          '-----------------------------------------------------------------------------------------------------
      
         If Range("I" & XCHI) = "ZWGYC-3269" Then '300P，W
         
             If InStr(Range("A" & XCHI), "W") > 0 Then   '铣槽
                
                    If (45 <= MM And MM <= 125) Or (175 <= MM And MM <= 255) Then
                
                         Range("I" & XCHI) = "ZWGYC-3887"
                         
                    ElseIf (50 <= MM And MM <= 110) Or (135 <= MM And MM <= 165) Or (190 <= MM And MM <= 250) Then '铣槽二代板
         
                         Range("I" & XCHI) = "ZWGYC-3269"
                    
                         Range("I" & XCHI).Interior.Color = RGB(255, 193, 37) '橙色
                    
                    Else
                
                         Range("I" & XCHI) = "YK-P006" '铣槽一代板
                    
                         Range("I" & XCHI).Interior.Color = RGB(127, 255, 212) '青色
                         
                         
                    End If
                    
             Else
             
                    If (50 <= MM And MM <= 110) Or (135 <= MM And MM <= 165) Or (190 <= MM And MM <= 250) Then '二代板
         
                         Range("I" & XCHI) = "ZWGYC-3269"
                    
                                         
                    Else
                
                        ' Range("I" & XCHI) = "YK-P006" '一代板
                    
                         Range("I" & XCHI).Interior.Color = RGB(127, 255, 212) '青色
                         
                    End If
                          
             End If
             
          End If
          
          
          
                          '-----------------------------------------------------------------------------------------------------
      
         If Range("I" & XCHI) = "ZWGYC-3270" Then '350P，W
         
             If InStr(Range("A" & XCHI), "W") > 0 Then   '铣槽
                
                    If (45 <= MM And MM <= 150) Or (200 <= MM And MM <= 305) Then
                
                         Range("I" & XCHI) = "ZWGYC-3886"
                         
                    ElseIf (50 <= MM And MM <= 115) Or (140 <= MM And MM <= 210) Or (235 <= MM And MM <= 300) Then '铣槽二代板
         
                         Range("I" & XCHI) = "ZWGYC-3270"
                    
                         Range("I" & XCHI).Interior.Color = RGB(255, 193, 37) '橙色
                    
                    Else
                
                         Range("I" & XCHI) = "YK-P007" '铣槽一代板
                    
                         Range("I" & XCHI).Interior.Color = RGB(127, 255, 212) '青色
                         
                    End If
                    
             Else
             
                    If (50 <= MM And MM <= 115) Or (140 <= MM And MM <= 210) Or (235 <= MM And MM <= 300) Then '二代板
         
                         Range("I" & XCHI) = "ZWGYC-3270"
                                        
                    Else
                
                       '  Range("I" & XCHI) = "YK-P007" '一代板
                    
                         Range("I" & XCHI).Interior.Color = RGB(127, 255, 212) '青色
                         
                    End If
                          
             End If
             
          End If
          
          
          
                          '-----------------------------------------------------------------------------------------------------
      
         If Range("I" & XCHI) = "ZWGYC-3271" Then '400P，W
         
             If InStr(Range("A" & XCHI), "W") > 0 Then   '铣槽
                
                    If (45 <= MM And MM <= 175) Or (225 <= MM And MM <= 355) Then
                
                         Range("I" & XCHI) = "ZWGYC-3885"
                         
                    ElseIf (50 <= MM And MM <= 160) Or (185 <= MM And MM <= 215) Or (240 <= MM And MM <= 350) Then '铣槽二代板
         
                         Range("I" & XCHI) = "ZWGYC-3271"
                    
                         Range("I" & XCHI).Interior.Color = RGB(255, 193, 37) '橙色
                    
                    Else
                
                         Range("I" & XCHI) = "ZWGYC-2370" '铣槽一代板
                    
                         Range("I" & XCHI).Interior.Color = RGB(127, 255, 212) '青色
                         
                         
                    End If
                    
             Else
             
                    If (50 <= MM And MM <= 160) Or (185 <= MM And MM <= 215) Or (240 <= MM And MM <= 350) Then '二代板
         
                         Range("I" & XCHI) = "ZWGYC-3271"
                                                             
                    Else
                
                       '  Range("I" & XCHI) = "ZWGYC-2370" '一代板
                    
                         Range("I" & XCHI).Interior.Color = RGB(127, 255, 212) '青色
                         
                    End If
                          
             End If
             
          End If
          
          
                          '-----------------------------------------------------------------------------------------------------
      
         If Range("I" & XCHI) = "ZWGYC-3272" Then '500P，W
         
             If InStr(Range("A" & XCHI), "W") > 0 Then   '铣槽
                
                    If (45 <= MM And MM <= 162.5) Or (212.5 <= MM And MM <= 287.5) Or (337.5 <= MM And MM <= 455) Then
                
                         Range("I" & XCHI) = "ZWGYC-3884"
                         
                    ElseIf (50 <= MM And MM <= 152.5) Or (197.5 <= MM And MM <= 302.5) Or (347.5 <= MM And MM <= 450) Then '铣槽二代板
         
                         Range("I" & XCHI) = "ZWGYC-3272"
                    
                         Range("I" & XCHI).Interior.Color = RGB(255, 193, 37) '橙色
                    
                    Else
                
                         '无一代板
                    
                         Range("I" & XCHI).Interior.Color = RGB(255, 0, 0) '红色
                         
                         
                    End If
                    
             Else
             
                    If (50 <= MM And MM <= 152.5) Or (197.5 <= MM And MM <= 302.5) Or (347.5 <= MM And MM <= 450) Then '二代板
         
                         Range("I" & XCHI) = "ZWGYC-3272"
                                         
                    Else
                                                             
                          Range("I" & XCHI).Interior.Color = RGB(127, 255, 212) '青色
                         
                    End If
                          
             End If
             
          End If
          
          
                          '-----------------------------------------------------------------------------------------------------
        If InStr(Range("H" & XCHI), "XP") > 0 Then
            
            Range("I" & XCHI) = "ZWMB-07"
            
             Range("I" & XCHI).Interior.Color = xlNone '无色
            
       End If
            
    End If
    
   End If
             
       
        
 'C区---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------++
      
   If InStr(TZBH, "墙柱") > 0 Then  '如果H列写的是墙柱C槽，判断D列长度，小于等于1200，则显示为C1X，否则为C2X
        
            If Range("I" & XCHI) = "HLD-03" Or Range("I" & XCHI) = "HLD-15" Then
                
                If Range("D" & XCHI) <= 1200 Then
                    
                    Range("G" & XCHI) = "L1"
                
                Else
                    
                    Range("G" & XCHI) = "L2"
                
                End If
                
            
            Else
                
                If Range("D" & XCHI) <= 1200 Then
                    
                    Range("G" & XCHI) = "C1X"
                
                Else
                    
                    Range("G" & XCHI) = "C2X"
                
                End If
            
            End If
            
            
       ElseIf Range("H" & XCHI) = "C槽XC" Or Range("H" & XCHI) = "C槽SC" Then
        
            If InStr(Range("A" & XCHI), "5SC") > 0 Or InStr(Range("A" & XCHI), "5XC") > 0 Then
            
                    If InStr(Range("H" & XCHI), "C槽SC") > 0 Then
                         
                         Range("I" & XCHI) = "ZWMB-11"
                         
                    End If
             
                    If InStr(Range("H" & XCHI), "C槽XC") > 0 Then
                         
                         Range("I" & XCHI) = "ZWMB-12"
                         
                    End If
                
                    If Range("D" & XCHI) <= 1200 Then
                    
                        Range("G" & XCHI) = "C1"
                        
                    Else
                    
                        Range("G" & XCHI) = "C2"
                        
                  End If
                  
           Else
    
                   If InStr(Range("H" & XCHI), "C槽SC") > 0 Then
                         
                         Range("I" & XCHI) = "ZWMB-56"
                         
                         Range("J" & XCHI) = "6000"
                         
                    End If
             
                    If InStr(Range("H" & XCHI), "C槽XC") > 0 Then
                         
                         Range("I" & XCHI) = "ZWMB-56+HLD-15"
                                                  
                         Range("J" & XCHI) = "6000"
                         
                    End If
                       
                    If Range("D" & XCHI) <= 1200 Then
                    
                        Range("G" & XCHI) = "L1"
                        
                    Else
                    
                        Range("G" & XCHI) = "L2"
                        
                  End If
                  
            End If
    
                  
       ElseIf InStr(TZBH, "C槽") > 0 And InStr(TZBH, "阴角") + InStr(TZBH, "转角") = 0 And InStr(TZBH, "墙柱") = 0 Then '如果H列写的是C槽，判断D列长度，小于等于1200，则显示为C1，否则为C2
        
            If Range("I" & XCHI) = "HLD-03" Or Range("I" & XCHI) = "HLD-15" Or Range("I" & XCHI) = "请输入型材" Then
                
               If InStr(Range("A" & XCHI), "AS") > 0 Or InStr(Range("A" & XCHI), "AL") > 0 Or InStr(Range("A" & XCHI), "AR") > 0 Then
                
                    Range("G" & XCHI) = "L1X"
                 
               Else
                               
                    If Range("D" & XCHI) <= 1200 Then
                    
                         Range("G" & XCHI) = "L1"
                
                    Else
                    
                          Range("G" & XCHI) = "L2"
                
                   End If
                                   
                End If
                
                                                       
         ElseIf InStr(Range("A" & XCHI), "AS") > 0 Or InStr(Range("A" & XCHI), "AL") > 0 Or InStr(Range("A" & XCHI), "AR") > 0 Then
                   
            If Range("D" & XCHI) <= 1200 Then
                    
                      Range("G" & XCHI) = "C1X"
                     
                Else
                    
                      Range("G" & XCHI) = "C2X"
                
            End If
            
                                            
            Else
               
                If Range("D" & XCHI) <= 1200 Then
                    
                    Range("G" & XCHI) = "C1"
                
                Else
                    
                    Range("G" & XCHI) = "C2"
                
                End If
                
                
        End If
 'N区---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------++
            
         ElseIf InStr(TZBH, "阴角C槽") + InStr(TZBH, "转角") > 0 Then   '如果是阴角C槽或者是写的是转角，则先看型材是不是L 板，
            
            If Range("I" & XCHI) = "HLD-03" Or Range("I" & XCHI) = "HLD-15" Or Range("I" & XCHI) = "请输入型材" Then
                
                Range("G" & XCHI) = "N2"
                
                
            ElseIf InStr(Range("A" & XCHI), "A") > 0 Or InStr(Range("A" & XCHI), "V") > 0 Then
                
                Range("G" & XCHI) = "N1X"
                
                
            ElseIf InStr(Range("A" & XCHI), "CN") > 0 Then
                
                 If Range("D" & XCHI) <= 1200 Then
                    
                    Range("G" & XCHI) = "C1"
                
                Else
                    
                    Range("G" & XCHI) = "C2"
                    
                End If
                
                
            ElseIf InStr(Range("A" & XCHI), "QN") > 0 Then
             
                 If Range("D" & XCHI) <= 1200 Then
                    
                    Range("G" & XCHI) = "C1X"
                
                Else
                    
                    Range("G" & XCHI) = "C2X"
                
                End If
                            
            Else
                
                Range("G" & XCHI) = "N1"
            
            End If
            
  'QT区---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------++
        ElseIf InStr(TZBH, "支撑") + InStr(TZBH, "固顶") > 0 Then '支撑
            
            If InStr(Range("A" & XCHI), "LTZ") > 0 Then
                
                If Range("D" & XCHI) <= 1200 Then
                    
                    Range("G" & XCHI) = "L1"
                
                Else
                    
                    Range("G" & XCHI) = "L2"
                
                End If
            
            Else
            
                Range("G" & XCHI) = "ZC"
            
            End If
            
        
        ElseIf InStr(TZBH, "龙骨") + InStr(TZBH, "铝梁") > 0 Then '龙骨或是铝梁
            
            Range("G" & XCHI) = "LG"
            
        
        ElseIf InStr(TZBH, "堵") + InStr(TZBH, "梁底") > 0 Then '堵板或梁底板
            
            Range("G" & XCHI) = "D"
            
            
        ElseIf Range("H" & XCHI) = "截面异形板" Then '截面异形板
            
            Range("G" & XCHI) = "P1X"
            
            
        ElseIf InStr(TZBH, "铝盒子") Or InStr(TZBH, "传料箱") Or InStr(TZBH, "泵送盒") Or InStr(TZBH, "放线口") Or InStr(Range("A" & XCHI), "XH") > 0 Or InStr(Range("E" & XCHI), "XH") > 0 Then '铝盒子
            
            Range("G" & XCHI) = "CLX"
            
            
        ElseIf Range("H" & XCHI) = "L型倒角板" Then 'L型倒角板
            
            Range("G" & XCHI) = "N1"
 'P区---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        ElseIf InStr(TZBH, "平面板") Or InStr(TZBH, "平板") Or InStr(TZBH, "普板") Or InStr(TZBH, "墙板") Or InStr(TZBH, "PK板") > 0 Or InStr(TZBH, "平面板切斜") > 0 Then '如果是平面板或平板或普板或楼梯墙板，则先需要判断型号是不是HLD-03或15，如果是则为L板，再判断是L1还是L2，如果不是l板则根据宽度判断，然后判断长度
            
            W1W2 = Range("C" & XCHI)
            l = Range("C" & XCHI)
            
                            
            If Range("I" & XCHI) = "HLD-03" Or Range("I" & XCHI) = "HLD-15" Or InStr(Range("I" & XCHI), "+") > 0 Or Range("I" & XCHI) = "请输入型材" Then
                
                    
                    If Range("D" & XCHI) <= 1200 Then '小于1200
                    
                      If InStr(Range("A" & XCHI), "TQ") > 0 Or InStr(TZBH, "平面板切斜") Then 'TQ切斜
                     
                              Range("G" & XCHI) = "L1X"
                                                      
                      ElseIf InStr(Range("A" & XCHI), "W") > 0 Then '铣槽
                              
                              Range("G" & XCHI) = "L1W"
                      
                      Else
                              Range("G" & XCHI) = "L1" '其余小于1200的
                        
                      End If
                      
                    Else '大于1200的，同上
                        
                      If InStr(Range("A" & XCHI), "TQ") > 0 Or InStr(TZBH, "平面板切斜") Then 'TQ切斜
                     
                              Range("G" & XCHI) = "L1X"
                        
                      ElseIf InStr(Range("A" & XCHI), "W") > 0 Then
                              
                              Range("G" & XCHI) = "L2W"
                      
                      Else
                              Range("G" & XCHI) = "L2"
                              
                      End If
                      
                    End If
                
            ElseIf Range("I" & XCHI) = "HLD-37" Or Range("I" & XCHI) = "HLD-40" Or Range("I" & XCHI) = "ZWGYC-3139" Or Range("I" & XCHI) = "ZWGYC-3138" Or Range("I" & XCHI) = "HLD-68" Or Range("I" & XCHI) = "HLD-42" Or Range("I" & XCHI) = "ZWGYC-2370" Or Left(Range("I" & XCHI), 4) = "YK-P" Then
                      
                     If InStr(Range("A" & XCHI), "TQ") > 0 Or InStr(TZBH, "平面板切斜") Then 'TQ切斜
                                                                           
                         If Range("D" & XCHI) <= 1500 Then
                        
                             Range("G" & XCHI) = "K1X"
                    
                         Else
                        
                             Range("G" & XCHI) = "K2X"
                    
                         End If
                         
                      ElseIf InStr(Range("A" & XCHI), "W") > 0 Then '铣槽
                       
                         If Range("D" & XCHI) <= 1200 Then
                        
                             Range("G" & XCHI) = "P1W"
                    
                         Else
                        
                             Range("G" & XCHI) = "P2W"
                    
                         End If
                         
                      Else
                      
                         If Range("D" & XCHI) <= 1500 Then '普通K板
                        
                             Range("G" & XCHI) = "K1"
                    
                         Else
                        
                             Range("G" & XCHI) = "K2"
                    
                         End If
                      
                      End If
                    
                
            ElseIf Range("I" & XCHI) = "ZWGYC-3265" Or Range("I" & XCHI) = "ZWGYC-3266" Or Range("I" & XCHI) = "ZWGYC-3267" Or Range("I" & XCHI) = "ZWMB-07" Then      '判断P1,P2
                    
                    If Range("D" & XCHI) <= 1200 Then '小于1200
                    
                      If InStr(Range("A" & XCHI), "TQ") > 0 Or InStr(TZBH, "平面板切斜") Then 'TQ切斜
                     
                              Range("G" & XCHI) = "P1X"
                                                      
                      ElseIf InStr(Range("A" & XCHI), "W") > 0 Then '铣槽
                              
                              Range("G" & XCHI) = "P1W"
                      
                      Else
                              Range("G" & XCHI) = "P1" '其余大于1200的
                        
                      End If
                      
                    Else '大于1200的，同上
                        
                      If InStr(Range("A" & XCHI), "TQ") > 0 Or InStr(TZBH, "平面板切斜") Then
                     
                              Range("G" & XCHI) = "P2X"
                        
                      ElseIf InStr(Range("A" & XCHI), "W") > 0 Then
                              
                              Range("G" & XCHI) = "P2W"
                      
                      Else
                              Range("G" & XCHI) = "P2"
                              
                      End If
                      
                    End If
                
            
                    
            ElseIf Range("I" & XCHI) = "ZWGYC-3268" Or Range("I" & XCHI) = "ZWGYC-3269" Or Range("I" & XCHI) = "ZWGYC-3270" Or Range("I" & XCHI) = "ZWGYC-3271" Or Range("I" & XCHI) = "ZWGYC-3888" Or Range("I" & XCHI) = "ZWGYC-3887" Or Range("I" & XCHI) = "ZWGYC-3886" Or Range("I" & XCHI) = "ZWGYC-3885" Or Range("I" & XCHI) = "ZWGYC-3884" Then '判断P3,P4
                    
                    If Range("D" & XCHI) <= 1200 Then '小于1200
                    
                      If InStr(Range("A" & XCHI), "TQ") > 0 Or InStr(TZBH, "平面板切斜") Then 'TQ切斜
                     
                              Range("G" & XCHI) = "P3X"
                                                      
                      ElseIf InStr(Range("A" & XCHI), "W") > 0 Then '铣槽
                              
                              Range("G" & XCHI) = "P3W"
                      
                      Else
                              Range("G" & XCHI) = "P3" '其余大于1200的
                        
                      End If
                      
                    Else '大于1200的，同上
                        
                      If InStr(Range("A" & XCHI), "TQ") > 0 Or InStr(TZBH, "平面板切斜") Then
                     
                              Range("G" & XCHI) = "P4X"
                        
                      ElseIf InStr(Range("A" & XCHI), "W") > 0 Then
                              
                              Range("G" & XCHI) = "P4W"
                      
                      Else
                              Range("G" & XCHI) = "P4"
                              
                      End If
                      
                    End If
                
             ElseIf Range("I" & XCHI) = "ZWGYC-3272" Or Range("I" & XCHI) = "ZWGYC-3884" Then '判断P5,P6
                    
                    If Range("D" & XCHI) <= 1200 Then '小于1200
                    
                      If InStr(Range("A" & XCHI), "TQ") > 0 Or InStr(TZBH, "平面板切斜") Then 'TQ切斜
                     
                              Range("G" & XCHI) = "P5X"
                                                      
                      ElseIf InStr(Range("A" & XCHI), "W") > 0 Then '铣槽
                              
                              Range("G" & XCHI) = "P5W"
                      
                      Else
                              Range("G" & XCHI) = "P5" '其余小于1200的
                        
                      End If
                      
                    Else '大于1200的，同上
                        
                      If InStr(Range("A" & XCHI), "TQ") > 0 Or InStr(TZBH, "平面板切斜") Then
                     
                              Range("G" & XCHI) = "P6X"
                        
                      ElseIf InStr(Range("A" & XCHI), "W") > 0 Then
                              
                              Range("G" & XCHI) = "P6W"
                      
                      Else
                              Range("G" & XCHI) = "P6"
                              
                      End If
                                                        
                End If
                
             ElseIf Left(Range("I" & XCHI), 4) = "YK-P" Then
                
                If Range("D" & XCHI) <= 1500 Then '这个是判断长度的
                    
                    Range("G" & XCHI) = "K1"
                
                Else
                    
                    Range("G" & XCHI) = "K2"
                
                End If
                
                            
            End If
            
 'K区---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------++
        ElseIf TZBH = "K板" Then '如果H列写的是K板，判断D列长度，小于等于1500，则显示为K1，否则为K2,另外如果型号不是开头是"YK-P",则对应的I列的型材单元格底色为红色
            
            If Range("D" & XCHI) <= 1500 Then '这个是判断长度的
                
                Range("G" & XCHI) = "K1"
            
            Else
                
                Range("G" & XCHI) = "K2"
            
            End If
                                      
'J区---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------++
        ElseIf Range("H" & XCHI) = "角铝" Or Range("H" & XCHI) = "底角铝" Then
        
            If InStr(Range("E" & XCHI), "ZW-J-08") > 0 Or InStr(Range("E" & XCHI), "ZW-J-09") > 0 Or InStr(Range("E" & XCHI), "ZW-J-10") > 0 Then
                
                    Range("G" & XCHI) = "JX2"
                
            ElseIf Range("H" & XCHI) = "角铝" And InStr(Range("A" & XCHI), "D-JL") > 0 Then 'Z形角铝
                        
                    Range("G" & XCHI) = "L1"
                
            ElseIf Range("H" & XCHI) = "底角铝" And InStr(Range("A" & XCHI), "NJ") > 0 Or InStr(Range("A" & XCHI), "ZW-J-08") > 0 Or InStr(Range("A" & XCHI), "ZW-J-09") > 0 Or InStr(Range("A" & XCHI), "ZW-J-10") > 0 Then
                
                    Range("G" & XCHI) = "JX2"
             
            Else
                
                     Range("G" & XCHI) = "J"
            
            End If
                 
            
        ElseIf Range("H" & XCHI) = "7字角铝" Or Range("H" & XCHI) = "角铝封板" Then '如果是7字角铝或者是角铝封板，类型为JX1

            Range("G" & XCHI) = "JX1"
            
            
        ElseIf InStr(Range("E" & XCHI), "LDJ") > 0 Then  '如果是LDJ大样图
        
             If Range("E" & XCHI) = "ZW-LDJ-01" Or Range("E" & XCHI) = "ZW-LDJ-02" Then 'LDJ-01或02大样图为J
            
                  Range("G" & XCHI) = "J"
            
             Else

                  Range("G" & XCHI) = "L1"
                  
             End If
            
                       
        ElseIf Range("H" & XCHI) = "封板" Or Range("H" & XCHI) = "铝板" Then  '吊模中的板子
            
            If InStr(Range("A" & XCHI), "J") > 0 And InStr(Range("A" & XCHI), "F") + InStr(Range("A" & XCHI), "L") > 0 Then
                
                Range("G" & XCHI) = "JX1"
                
            ElseIf InStr(Range("A" & XCHI), "FB") > 0 Then
                
                Range("G" & XCHI) = "FB"
            
            End If
 'LT区---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------++
        
                              
        ElseIf InStr(TZBH, "抬头板") > 0 Then '如果H列写的是折板
            
                          
                Range("G" & XCHI) = "L1X"
                
                
        ElseIf InStr(TZBH, "挡板") > 0 Then '楼梯挡板
            
            Range("G" & XCHI) = "DB"
            
        
        ElseIf InStr(TZBH, "盖板") > 0 Then '楼梯盖板
            
            Range("G" & XCHI) = "GB"
            
        
        ElseIf InStr(TZBH, "侧板") > 0 Then '楼梯侧板
            
            Range("G" & XCHI) = "CB"
            
                
        ElseIf InStr(TZBH, "狗牙") > 0 Then  '狗牙或是楼梯狗牙
         
            If Range("D" & XCHI) <= 1500 Then '这个是判断长度的
                
                Range("G" & XCHI) = "T1"
            
            Else
                
                Range("G" & XCHI) = "T2"
            
            End If
        
        If InStr(Range("I" & XCHI), "+") > 0 And InStr(Range("H" & XCHI), "狗牙") + InStr(Range("H" & XCHI), "封板") + InStr(Range("H" & XCHI), "角铝") = 0 Then
            
            If Range("D" & XCHI) <= 1200 Then
                
                Range("G" & XCHI) = "L1"
            
            Else
                
                Range("G" & XCHI) = "L2"
            
            End If
        
        End If
        
  End If
       
        If Len(Range("H" & XCHI)) > 0 And Len(Range("G" & XCHI)) = 0 Then Range("G" & XCHI).Interior.Color = RGB(255, 0, 0)
        
        If Mid(Range("F" & XCHI), 2, 1) = "-" Or Mid(Range("F" & XCHI), 3, 1) = "-" And Len("A" & XCHI) > 0 Then
        
            Range("M" & XCHI) = Split(Range("F" & XCHI), "-")(0)
            
            If Left(Range("E" & XCHI), 2) <> "ZW" Then Range("N" & XCHI) = Range("E" & XCHI)
        
        ElseIf Len("A" & XCHI) = 0 Then
        
            Range("M" & XCHI) = ""
        
        End If
   
    
    Next
 
            
            Columns("U:U").Select
             Selection.Delete
    
    '根据零件的分区编号得出来是哪个区域, '这一步将每个图纸用到的非标件图纸取出来
    Columns("M:N").Select
    Range("M1") = "分区"
    Range("N1") = "图纸编号"
    
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "erp!R1C13:R1048576C14", Version:=xlPivotTableVersion14).CreatePivotTable TableDestination:= _
        "erp!R1C18", TableName:="数据透视表2", DefaultVersion:=xlPivotTableVersion14

    ActiveWorkbook.ShowPivotTableFieldList = True
    
    With ActiveSheet.PivotTables("数据透视表2").PivotFields("分区")
        
        .Orientation = xlRowField
        .Position = 1
    
    End With
    
    With ActiveSheet.PivotTables("数据透视表2").PivotFields("图纸编号")
        
        .Orientation = xlRowField
        .Position = 2
    
    End With
    
    ActiveWorkbook.ShowPivotTableFieldList = False
    
    Columns("R:R").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    endR = Sheets("erp").[R6000].End(xlUp).Row
    Range("R" & endR) = ""
    Range("R" & endR - 1) = ""
    
    Range("R1") = "分区非大样图"
    Columns("Q:R").Replace "*白*", ""
    
    Columns("R:R").Cut
    Columns("P:P").Select
    ActiveSheet.Paste
    
    Columns("M:O").Delete Shift:=xlToLeft
    
    Columns("I:J").Columns.AutoFit
    
    'MsgBox ("总数量：" & Slhj & " 件" & Chr(10) & "蓝色格为生产单的第25个零件号" & Chr(10) & "注意特殊型材" & Chr(10) & "如平面板编号包含D,判断是否用一代板")
   ' MsgBox ("注意特殊型材" & Chr(10) & "如平面板编号包含D,判断是否用一代板")

End Sub

Sub FBC填充单号()

    Columns("A:A").Interior.Pattern = xlNone
    
    enda = Range("A60000").End(xlUp).Row
    
    For dys = 1 To enda '单页数，一个单子要放的数量
        
        If Len(Range("B" & dys)) = 0 And Len(Range("B" & dys + 1)) > 0 Then
            
            k = 0
        
        Else
            
            k = k + 1
        
        End If
        
        Remainder = k Mod 26  '余数
        
        If k > 0 And Remainder = 0 Then
            
            Range("A" & dys - 1).Interior.Color = RGB(232, 159, 187)
            lj = lj + 1 '累计次数
        
        End If
    
    Next
    
    If lj > 0 Then
        
        MsgBox "还有超过单页25的生产单,调整后重新点击填充单号"
        
        Exit Sub
    
    End If
    
    mbmc = ""
    
    Range("A2:G" & enda).Borders.LineStyle = xlContinuous
    
    If Len(Range("A1")) = 0 Then a = Sheets("Sheet1").Range("B3")
    
    Range("A1") = a & "-" & Range("K2") & "-1"
    
    k = 1
    
    For ih = 1 To enda
        
        If Len(Range("A" & ih)) = 0 Then
            
            fenqulast = Range("K" & ih - 1)
            fenqu = Range("K" & ih + 1)
            
            If fenqu = fenqulast Then
                
                k = k + 1
                Range("A" & ih) = a & "-" & fenqu & "-" & k
            
            Else
                
                k = 1
                Range("A" & ih) = a & "-" & fenqu & "-" & k
            
            End If
        
        End If
    
    Next
    
    Columns("H:H").FormatConditions.Delete
    
  '  MsgBox ("如果图号有5XC,5SC,请注意修改型材")
  '  MsgBox ("如需修改型材，请直接在erp表修改,注意切斜模板")
       
End Sub

Sub FBD转序单生产单()

    MsgBox ("注意切斜模板的类型要加'X'")
    
    Dim p&, zhs&, i1&, jch&, PE&, sn&

    Application.ScreenUpdating = False
    
    Sheets("erp").Activate
    
    zhs = Sheets("erp").[b6000].End(xlUp).Row
    
    Columns("M:O").Delete Shift:=xlToLeft
    Columns("A:J").HorizontalAlignment = xlCenter
    Range("A1:j6000").Interior.Pattern = xlNone
    
    Columns("B:C").Insert
    Columns("H:H").Insert
    
    wz = Mid(Cells(1, 1), 1, 4)
    
    For p = 1 To zhs '合并图纸编号单元格，以方便放入转序单
    
        If Mid(Cells(p, 1), 1, 4) = wz Then
            
            jsp = jsp + 1 '计数p
        
        Else
            
            Cells(p, 1).Resize(, 3).Merge
            Cells(p, 7).Resize(, 2).Merge
        
        End If
    
    Next
    
    Sheet5.Activate
    'Sheets("ZXD").Range("A2") = "项目名称：" & Sheet5.Range("B2") & Sheet5.Range("G2")
    
    '=======================================================================================
    '191211按模板厂需求调整生产计划单格式
    
    Dim aaa As String '用于存储生产单号
    Dim shp
    
    Rows(3).RowHeight = 70
    Rows("7:31").RowHeight = 19
    Range("B3").Font.Size = 36
    
    '=======================================================================================

    
    For zjb = 1 To jsp - 1 '增加表
        
        Sheet5.Select
        Sheet5.Copy Before:=Sheets("erp")
    
    Next
       
    Sheets("erp").Activate
    
    Range("A1:H" & zhs).Borders.LineStyle = xlContinuous  '合并居中加边框
    
    For i1 = 1 To zhs + jsp * 2
        
        If Mid(Cells(i1, 1), 1, 4) = wz Then
            
            Rows(i1).Insert
            Rows(i1 + 2).Insert
            
            i1 = i1 + 1
        
        End If
    
    Next
    
    For jch = 1 To zhs + jsp * 2 '加插入行的总行数
    
        On Error Resume Next
        
        If Mid(Cells(jch, 1), 1, 4) = wz Then
            
            bs = bs + 1
            r = Cells(jch + 2, 1).CurrentRegion.Rows.Count '连续区域的行数
            'c = Cells(jch + 2, 1).CurrentRegion.Columns.Count '连续区域的列数，然并卵
            xck = Cells(jch + 2, 5) '型材宽度
            mbmc = Cells(jch + 2, 1).Text '模板名称
            
            Cells(jch, 1).Copy Sheets(bs + 1).[B3:E3] '转序单号复制
            
            '=======================================================================================
            '191211按模板厂需求调整生产计划单格式
            
            Sheets(bs + 1).Activate
            
            With Sheets(bs + 1)
            
                aaa = .Range("B3")
                .Range("B3") = "=code128(""" & aaa & """,B3,,230,)"
                .Range("B3") = aaa
                .Range("B3").Font.ColorIndex = 2
                
            End With
            
            For Each shp In ActiveSheet.Pictures
        
                shp.Left = (shp.TopLeftCell.Width - shp.Width) / 2 + 2.3 * shp.TopLeftCell.Left
                shp.Top = (shp.TopLeftCell.Height - shp.Height) / 2 + 1.05 * shp.TopLeftCell.Top
        
            Next
            
            Sheets("erp").Activate
            
            '=======================================================================================

            
            Cells(jch + 2, 1).Resize(r, 4).Copy Sheets(bs + 1).[B7] '图号及数量复制
            Cells(jch + 2, 7).Resize(r, 4).Copy Sheets(bs + 1).[F7] '图纸编号及分区
            
            Cells(jch + 2, 15).Resize(r, 2).Copy
            Sheets(bs + 1).[J7].PasteSpecial Paste:=xlPasteValues  '备注的复制
            
            Quyu = Range("N" & jch + 2)
            xch = Range("L" & jch + 2)
            dingchi = Range("M" & jch + 2)
            
            If Left(Quyu, 2) = "TP" Then Sheets(bs + 1).Range("J7:J" & (6 + r)) = "带配件"
            
            Sheets(bs + 1).Activate
            
            Sheets(bs + 1).[G2] = Quyu
            Sheets(bs + 1).[B4:E4] = xch  '型材截面号的输入
            Sheets(bs + 1).[G4] = dingchi '定尺输入

            Sheets("计算用表").Range("B2:C31").ClearContents
            
            Sheets("erp").Activate
                
            Cells(jch + 2, 4).Resize(r, 1).Copy Sheets("计算用表").[C2]
            Cells(jch + 2, 6).Resize(r, 1).Copy Sheets("计算用表").[B2]
            
            Sheets("计算用表").Activate
                
            Sheets("计算用表").[f1] = dingchi
                
            Call Z优化
            
            If Sheets("计算用表").[f21] = 0 Then
                    
                Sheets(bs + 1).[I4:K4] = 1
                    
            Else
                
                Sheets("计算用表").[f21].Copy Sheets(bs + 1).[I4：K4]
                
            End If
                
            Sheets("erp").Activate
        
        End If
    
    Next

    Sheet5.Activate

    For PE = 2 To Sheets.Count
    
        If InStr(Sheets(PE).[A1], "模板") > 0 Then
            
            xh = xh + 1
        
            Sheet1.Cells(PE + 2, 1) = xh
            Sheet1.Cells(PE + 2, 2) = Sheets(PE).Cells(3, 2)
            Sheet1.Cells(PE + 2, 3) = Sheets(PE).Cells(3, 9)
            'Sheet1.Cells(PE + 2, 10) = Sheets(PE).Cells(4, 2)
            'Sheet1.Cells(PE + 2, 13) = Sheets(PE).Cells(4, 6)
            'Sheet1.Cells(PE + 2, 11) = Sheets(PE).Cells(4, 7)
            'Sheet1.Cells(PE + 2, 12) = Sheets(PE).Cells(4, 9)
            
            Sheets(PE).Name = Sheets(PE).Cells(3, 2)
        
            If Sheet1.Cells(PE + 2, 2) = "" Then
            
                Sheet1.Cells(PE + 2, 3) = ""
                Sheet1.Cells(PE + 2, 10) = ""
                Sheet1.Cells(PE + 2, 11) = ""
                Sheet1.Cells(PE + 2, 12) = ""
        
            End If
    
        End If
    
    Next
   
    Sheets("erp").Activate
    
    Columns("B:C").Delete
    Columns("F:F").Delete
    
    Columns("A:A").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    
    '=======================================================================================
    '191203按模板厂需求增加erp工作表整理导出环节
    Dim brr
    
    Dim t As Integer
    Dim enderp As Integer
    Dim scdh As String
    Dim erpfzl As String
    Dim erpjs As Integer
    
    Dim qcbm As String
    
    Sheets("erp").Copy after:=Sheets("erp")
    ActiveSheet.Name = ("erp库")
    
    Sheets("erp库").Activate
    
    enderp = Sheets("erp库").Range("B65536").End(xlUp).Row
    
    For t = 1 To enderp
    
        If Len(Range("B" & t)) = 0 Then
        
            scdh = Range("A" & t)
            
        Else
            
            Range("L" & t) = scdh
            Range("M" & t) = scdh & Range("K" & t) & Range("G" & t)
            
        End If
        
    Next t
    
    Columns("B:B").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    Columns("L:L").Cut
    Columns("A:A").Insert Shift:=xlToRight
    Columns("B:B").Delete
    Columns("C:F").Delete
    Columns("D:F").Delete
    Columns("B:B").Cut
    Columns("E:E").Insert Shift:=xlToRight
    Columns("C:C").Cut
    Columns("B:B").Insert Shift:=xlToRight
    
    Rows(1).Insert
    brr = Array("生产单号", "区域简写", "生产单类型", "支数", "参考列", "支数合计")
    [A1].Resize(1, UBound(brr) + 1) = brr
    
    enderp = Sheets("erp库").Range("B65536").End(xlUp).Row
    
    Columns("A:F").EntireColumn.AutoFit
    Range("A1:F" & enderp).HorizontalAlignment = xlCenter
    Range("A1:F" & enderp).Borders.LineStyle = xlContinuous
    
    With Sheets("erp库").Sort.SortFields
    
        .Clear
        .Add Key:=Range("A2"), Order:=1
        .Add Key:=Range("B2"), Order:=1
        .Add Key:=Range("C2"), Order:=1
        
    End With
    
    With Sheets("erp库").Sort
    
        .SetRange Range("A2:E" & enderp)
        .Header = 2
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
        
    End With
    
    erpfzl = Range("E2")
    erpjs = 0
    
    For t = 2 To enderp + 1
    
        If Range("E" & t) <> erpfzl Then
            
            Range("F" & t - 1) = erpjs
            
            erpfzl = Range("E" & t)
            erpjs = Range("D" & t)
            
        Else
        
            erpjs = erpjs + Range("D" & t)
            
        End If
        
    Next t
    
    Columns("F:F").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    Columns("D:E").Delete
    ActiveWindow.ScrollRow = 1
    
    qcbm = Replace(ThisWorkbook.Name, "生产单.xlsm", "库数据")
    
    Worksheets(Array("erp库")).Copy
    ActiveWorkbook.SaveAs FileName:=ThisWorkbook.path & "\" & qcbm, FileFormat:=51
    ActiveWorkbook.Close SaveChanges:=True
    
    Application.DisplayAlerts = False
    
    Sheets("erp库").Delete
    
    Application.DisplayAlerts = True
    
    '=======================================================================================

    
    Sheets("ZXD").Activate
    
    zxdzh = Sheets("ZXD").Range("B65536").End(xlUp).Row
    
    Columns("A:H").HorizontalAlignment = xlCenter
    Sheets("ZXD").Range("B2").HorizontalAlignment = xlLeft
    
    For xh = 4 To zxdzh
        
        v = Split(Range("B" & xh), "-")
        XHZ = Right(Range("B" & xh), Len(v(UBound(v))))
        
        Range("A" & xh) = XHZ
    
    Next
    
    For xh = 4 To zxdzh
        
        If Range("A" & xh) = 1 Then
             
             JISHU1 = JISHU1 + 1
        
        End If
    
    Next
    
    If JISHU1 > 1 Then
        
        For crhj = 5 To zxdzh + (JISHU1 - 1) * 3 '插入合计
            
            If Range("A" & crhj) = "1" Then
            
                Rows(crhj & ":" & (crhj + 2)).Insert
                crhj = crhj + 3
            
            End If
        
        Next
        
        crhj = ""
        
        zxdzh = Sheets("ZXD").Range("B65536").End(xlUp).Row
        
        For crhj = 5 To zxdzh
            
            If Range("A" & crhj) = "1" Then
            
                Range("B" & (crhj - 2)) = "合计"
                Range("C" & (crhj - 2)) = Application.WorksheetFunction.Sum(Range("C" & (crhj - 3) & ": C" & ((crhj - 3) - HS + 2)))
                
                HS = 1
            
            Else
            
                HS = HS + 1
            
            End If
        
        Next
        
        Range("B" & zxdzh + 2) = "合计"
        Range("C" & zxdzh + 2) = Application.WorksheetFunction.Sum(Range("C" & zxdzh & ": C" & (zxdzh - HS + 1)))
        
        i = "'"
        
        For i = 4 To zxdzh + 2
            
            If Range("B" & i) = "合计" Then
                
                Sheets("ZXD").HPageBreaks.Add Range("b" & i + 2)
            
            End If
        
        Next
    
    Else
        
        Range("B" & zxdzh + 2) = "合计"
        Range("C" & zxdzh + 2) = WorksheetFunction.Sum(Range("C" & zxdzh & ": C4"))
    
    End If
    
    Range("A" & zxdzh + 1 & ":A" & (zxdzh + 100)).ClearContents
    
    Sheets("ZXD").PageSetup.PrintArea = "$A$1:$H$" & zxdzh + 2
    Sheets("ZXD").PageSetup.PrintTitleRows = "$1:$3"
    Range("A4:H" & zxdzh + 2).Borders.Weight = 2
    'Range("A4:H" & ZXDZH + 2).BorderAround , 3
    Rows("4:" & zxdzh + 2).RowHeight = 20
    
    Columns("J:J").EntireColumn.AutoFit
    
    Application.ScreenUpdating = True
    
    MsgBox ("备料单中型材须拆分为一个单元格只有一种型材型号")

End Sub
Sub FBF生成ERP()
Application.ScreenUpdating = False
    Sheets("erp").Activate
    ENDC = [C65536].End(xlUp).Row '对C列的最后一行进行定位，在拆分明细的库里找一下C列的宽度，如果没有就显示宽度为红色
    For i = 2 To ENDC
        If Len(Range("C" & i)) > 0 Then
            xck = Cells(i, 3) '型材宽度
            If Sheets("库(待补充)").Columns(6).Find(xck, LookAt:=xlWhole, SearchDirection:=xlprerious) Is Nothing Then
                Range("C" & i).Interior.Color = RGB(230, 100, 100) '
                k = k + 1 '如果K大于零，则需要在库里加新的拆分明细表
            End If
        End If
    Next
    
    If k > 0 Then
        MsgBox ("需在库中添加明细后，重新生成erp")
        GoTo 100
    End If
    
    Sheets("erp").Copy Before:=Sheets("库(待补充)")
    ActiveSheet.Name = ("拆分明细")
    Columns("E:M").Delete
    
    Columns("b:b").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    Range("A1:D" & ActiveSheet.UsedRange.Rows.Count).Borders.LineStyle = xlContinuous
    For BKTOP = 1 To ActiveSheet.UsedRange.Rows.Count '设置行的上边框
        Range("A" & BKTOP & ":D" & BKTOP).Borders(xlEdgeTop).Weight = xlMedium
    Next BKTOP

'    Columns("A:A").Replace What:=" ", Replacement:=""
    Columns("D:E").Insert
    For i = 1 To Sheets("拆分明细").UsedRange.Rows.Count * 10
     On Error Resume Next
        kuandu = Sheets("拆分明细").Cells(i, 3)
        If kuandu = "" Then Exit For
        hangshu = Sheets("库(待补充)").Columns(6).Find(kuandu, LookAt:=xlWhole, SearchDirection:=xlprerious).Row
        Sheets("库(待补充)").Activate
        Range("F" & hangshu).Resize(, 12).Select
        r = Selection.Rows.Count
        Sheets("拆分明细").Activate
        Rows(i + 1 & ":" & i + r).Insert
        m = m + 1
        Sheets("库(待补充)").Range("F" & hangshu).Resize(r, 12).Copy Sheets("拆分明细").Cells(i + 1, 1)
        Sheets("拆分明细").Cells(i, 3).ClearContents
        Range("A" & i & ":F" & i).Copy
        Range("B" & i + 1).PasteSpecial SkipBlanks:=True
        Sheets("拆分明细").Rows(i).ClearContents
        Cells(i + 1, 1) = m
        i = i + r
    Next i
    
    Columns("F:F").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues
    
    Columns("F:F").SpecialCells(xlCellTypeBlanks).EntireRow.Delete

    For j = ActiveSheet.UsedRange.Rows.Count To 1 Step -1
        If Cells(j, 6).Text = "" Then
            Rows(j).EntireRow.Delete
        End If
    Next j

    Rows("1:1").Insert
    arr = Array("序号", "模板编号", "数量", "图纸名称", "型材截面", "材质", "长度", "数量", "总数量", "理论重量", "总重kg", "型材类型")
    [A1].Resize(1, UBound(arr) + 1) = arr
    '设置表头的格式
     With Range("A1:L1")
        .HorizontalAlignment = xlCenter
        .Borders.Weight = 2
        .BorderAround , 3
    End With
    '设置宽度自适应，然后调整比较窄的列，列宽为8
    Columns("A:L").EntireColumn.AutoFit
    Range("A:A,C:C,H:H,I:I").ColumnWidth = 8
    
    Rows("1:1").RowHeight = 25
    Call FBG如果修改明细重新算配件


100:
Application.ScreenUpdating = True
End Sub


Sub FBG如果修改明细重新算配件()
    Sheets("拆分明细").Activate
    Application.ScreenUpdating = False
    
    Columns("E:L").Select
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "拆分明细!R1C5:R1048576C12", Version:=xlPivotTableVersion14).CreatePivotTable _
        TableDestination:="拆分明细!R1C14", TableName:="数据透视表1", DefaultVersion:= _
        xlPivotTableVersion14
    With ActiveSheet.PivotTables("数据透视表1").PivotFields("型材类型")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("数据透视表1").PivotFields("型材截面")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("数据透视表1").PivotFields("长度")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables("数据透视表1").AddDataField ActiveSheet.PivotTables("数据透视表1" _
        ).PivotFields("总数量"), "总计数量", xlSum
        
    ActiveSheet.PivotTables("数据透视表1").RowAxisLayout xlTabularRow
'    ActiveSheet.PivotTables("数据透视表1").RepeatAllLabels xlRepeatLabels
    Dim p As PivotField
    For Each p In ActiveSheet.PivotTables("数据透视表1").PivotFields
        p.Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
    Next
        
    With ActiveSheet.PivotTables("数据透视表1").PivotFields("型材类型")
        .PivotItems("").Visible = False
'        .PivotItems("主板").Visible = False
        .PivotItems("(blank)").Visible = False
    End With
    '对行和列禁用汇总
    With ActiveSheet.PivotTables("数据透视表1")
        .ColumnGrand = False
        .RowGrand = False
    End With
    '对型材截面进行粘贴为值，对应库找到型材对应的定尺，然后用计算用表获取支数
    Columns("N:N").Copy
    Columns("R:R").PasteSpecial Paste:=xlPasteValues
    Range("r1") = ""
    Range("r3") = "配件型材截面"
    Columns("R:R").SpecialCells(xlCellTypeBlanks).Delete Shift:=xlUp
    ActiveSheet.PivotTables("数据透视表1").RepeatAllLabels xlRepeatLabels
    endR = Range("R5000").End(xlUp).Row
    For i = 2 To endR
        xcjm = Range("r" & i) '型材截面
        If Sheets("库(待补充)").Columns(2).Find(xcjm, LookAt:=xlWhole, SearchDirection:=xlprerious) Is Nothing Then
            dingchi = "6000"
        Else
            hangshu = Sheets("库(待补充)").Columns(2).Find(xcjm, LookAt:=xlWhole, SearchDirection:=xlprerious).Row
            dingchi = Sheets("库(待补充)").Range("C" & hangshu)
        End If
        Range("s" & i) = dingchi
        
        hangshumin = Columns("N:N").Find(xcjm, LookAt:=xlWhole, SearchDirection:=xlprerious).Row
        Range("T" & i) = hangshumin
    Next
    Range("s1") = "定尺"
   
    Sheets("计算用表").Range("B2:C100").ClearContents
    Sheets("拆分明细").Activate
    endN = Range("N5000").End(xlUp).Row
    endR = Range("R5000").End(xlUp).Row
    For i = 2 To endR
        Sheets("拆分明细").Activate
        py = Range("T" & i) '偏移起始单元格
        If i < endR Then
            pyfw = Range("T" & (i + 1)) - Range("T" & i)
        Else
            pyfw = endN + 1 - Range("T" & i)
        End If
        
        Range("O" & py).Resize(pyfw, 2).Copy Sheets("计算用表").[B2]
        Sheets("计算用表").[f1] = Range("S" & i)
        Sheets("计算用表").Activate
        Call Z优化
        If Sheets("计算用表").[f21] = 0 Then
            Sheets("拆分明细").Range("U" & i) = 1
        Else
            Sheets("计算用表").[f21].Copy Sheets("拆分明细").Range("U" & i)
        End If
        Sheets("计算用表").Range("B2:C100").ClearContents
    Next
    Sheets("拆分明细").Activate
    Range("U1") = "支数"
    Columns("T:T").Delete
    Columns("N:P").Delete
    
    With Range("O1:Q" & endR)
        .HorizontalAlignment = xlCenter
        .Borders.Weight = 2
    End With
    Columns("O:O").EntireColumn.AutoFit
    For i = 2 To endR
        If InStr(Range("O" & i).Text, "板材") > 0 Then
            Range("O" & i).Interior.Color = RGB(230, 100, 100)
            bcsl = bcsl + 1
        End If
    Next
    If bcsl > 0 Then MsgBox ("辅料用到板材，将数量按照面积换算成小数")
    Application.ScreenUpdating = True
    

End Sub
Sub FBE上传()
    Dim sh As Worksheet, m&
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    For Each sh In Worksheets
        gzbm = sh.Name
        aaa = InStr(gzbm, "-")
        If sh.Visible = xlSheetVisible And InStr(gzbm, "M") > 0 And InStr(gzbm, "-") = 3 Then
            m = m + 1
            If m = 1 Then sh.Select Else sh.Select False
        End If
    Next
    ActiveWindow.SelectedSheets.Copy
    
    'p代码后面的数字超过1000,不能再用前6位表示
    pdm = Left(ThisWorkbook.Sheets("erp").Range("A1").Text, 3) & Split(ThisWorkbook.Sheets("erp").Range("A1"), "-")(1)
    XMQU = ThisWorkbook.Sheets("ZXD").Range("B2").Value '项目区域
    ActiveWorkbook.SaveAs FileName:=ThisWorkbook.path & "\" & pdm & XMQU & "-上传.xlsm", FileFormat:=xlOpenXMLWorkbookMacroEnabled
    ActiveWorkbook.Close
    MsgBox ("上传文件生成完成")
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub


