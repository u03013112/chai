Option Explicit

' 基于原有宏修改，主要修改方式为把原有修改本文件另存为的方式改为在本地处理之后新建文件，写入后保存，然后恢复本地文件
' 生成目录也改为在本文件夹下生成，不去改原有目录，这步还需要考虑
' 另外的需求就是再做一个界面用于展示，这样样貌会比较漂亮
' 目前的文件大部分保持不动

Dim fso As Object, arr(1 To 100, 1 To 1), i
Dim dg As FileDialog

Sub A合并拆分非标件清单()
    
    '选择标准件及标准件带配件所在文件夹
    Dim strfile As String
    Dim brr
    
    Set dg = Application.FileDialog(msoFileDialogFolderPicker)
    
    If dg.Show = -1 Then
    
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        If isSheetExist(ThisWorkbook, "设计非标件清单") Then
            ThisWorkbook.Sheets("设计非标件清单").Delete
        End If
        If isSheetExist(ThisWorkbook, "设计标准件清单") Then
            ThisWorkbook.Sheets("设计标准件清单").Delete
        End If
        If isSheetExist(ThisWorkbook, "设计打包清单") Then
            ThisWorkbook.Sheets("设计打包清单").Delete
        End If
        Sheets.Add().Name = "设计非标件清单"
        Sheets.Add().Name = "设计标准件清单"
        Sheets.Add().Name = "设计打包清单"
        ' Sheets("ZXD").Delete
        ' Sheets("Sheet1").Delete
        ' Sheets("erp").Delete
        ' Sheets("计算用表").Delete
        Application.DisplayAlerts = True
        
        ThisWorkbook.Sheets("设计标准件清单").Activate
        
        brr = Array("序号", "模板名称", "模板编号", "W1", "W2", "L", "单件面积", "数量", "总件面积", "图纸编号", "工作表名", "是否带配件")
        
        [A1].Resize(1, UBound(brr) + 1) = brr
        Sheets("设计非标件清单").[A1].Resize(1, UBound(brr) + 1) = brr
        
        brr = Array("序号", "模板名称", "数量", "打包表名")
        
        Sheets("设计打包清单").[A1].Resize(1, UBound(brr) + 1) = brr
        
        strfile = dg.InitialFileName
        Set fso = CreateObject("scripting.filesystemobject")
        
        Erase arr()
        i = 0
        合并设计传递 dg.SelectedItems(1)
        
    Else
    
        Exit Sub
    
    End If
    
    Dim wjj_name As String
    Dim i_mbmc  As Integer '遍历模板名称的遍历字符
    Dim endb As Integer
    
    wjj_name = Split(dg.SelectedItems(1), "\")(UBound(Split(dg.SelectedItems(1), "\")))
    
    With Sheets("设计非标件清单")
        
        endb = .Cells(Rows.Count, 2).End(xlUp).Row
        
        .Columns("B:B").Replace "平板", "平面板"
        .Columns("B:B").Replace "转角C槽", "转角"
        
        '模板名称是C槽,模板编号带N 则将模板名称改为转角
        '模板名称是平面板,模板编号带小数点,则将模板名称改为平面板切斜
        
        For i_mbmc = 2 To endb
            
            If (.Range("B" & i_mbmc) = "C槽" Or .Range("B" & i_mbmc) = "阴角") And InStr(.Range("C" & i_mbmc), "N") > 0 Then
                
                .Range("B" & i_mbmc) = "转角"
            
            End If
            
            If .Range("B" & i_mbmc) = "C槽" And InStr(.Range("C" & i_mbmc), "XC") > 0 Then
                
                .Range("B" & i_mbmc) = "C槽XC"
            
            End If
            
            If .Range("B" & i_mbmc) = "C槽" And InStr(.Range("C" & i_mbmc), "SC") > 0 Then
                
                .Range("B" & i_mbmc) = "C槽SC"
            
            End If
            
            If .Range("B" & i_mbmc) = "平面板" And InStr(.Range("C" & i_mbmc), "XP") > 0 Then
                
                .Range("B" & i_mbmc) = "平面板XP"
            
            End If
            
            If .Range("B" & i_mbmc) = "平面板" And InStr(.Range("C" & i_mbmc), ".") > 0 Then
                
                .Range("B" & i_mbmc) = "平面板切斜"
            
            End If
            
        Next i_mbmc
        
    End With
    
    Set dg = Nothing
    Set fso = Nothing
    
    Sheets("设计非标件清单").Tab.ColorIndex = 3
    Sheets("设计非标件清单").Activate
    
    ' Dim hbqdFilename As String
    ' hbqdFilename = strfile & wjj_name & "\" & wjj_name & "-合并清单.xlsm"
    ' Call saveHbqd(hbqdFilename)
    
    ' ThisWorkbook.SaveAs FileName:=strfile & wjj_name & "\" & wjj_name & "-合并清单.xlsm"
    
    '---写出现有模板名称对应的生产单名称----------------------------------------------------------
    
    Dim end_O As Integer
    Dim mbmc As String 'o列的模板名称
    Dim scdmc As String '生产单名称
    Dim hangshu As Integer
    
    Columns("B:B").Copy
    Columns("O:O").Select
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Application.CutCopyMode = False
    
    ActiveSheet.Range("$O$1:$O$" & endb).RemoveDuplicates Columns:=1, Header:=xlNo
    
    ActiveSheet.Columns("O:P").EntireColumn.AutoFit
    
    end_O = Range("O6000").End(xlUp).Row
    
    For i = 1 To end_O
    
        mbmc = Range("O" & i) '型材宽度
        
        If Sheets("库(待补充)").Columns(4).Find(mbmc, LookAt:=xlWhole, SearchDirection:=xlPrevious) Is Nothing Then
            
            scdmc = "QT"
        
        Else
            
            hangshu = Sheets("库(待补充)").Columns(4).Find(mbmc, LookAt:=xlWhole, SearchDirection:=xlPrevious).Row
            scdmc = Sheets("库(待补充)").Range("E" & hangshu) '生产单命名
        
        End If
        
        Range("P" & i) = scdmc
    
    Next
    
    Application.DisplayAlerts = False
    
    ' Sheets("库(待补充)").Delete
    
    Application.DisplayAlerts = True
    
    Call 分出标准件非标件
    
    Call 清单差异比对
    
    Application.DisplayAlerts = False
    
    If Sheets("清单差异比对").Cells(Rows.Count, 1).End(xlUp).Row > 1 Then
        
        Sheets("清单汇总处理").Delete
        
        ThisWorkbook.Worksheets("清单差异比对").Columns("A:E").EntireColumn.AutoFit
        
        Worksheets(Array("设计打包清单", "设计标准件清单", "设计非标件清单", "清单差异比对")).Copy
        ActiveWorkbook.SaveAs FileName:=strfile & wjj_name & "\" & wjj_name & "-清单差异", FileFormat:=51
        ActiveWorkbook.Close SaveChanges:=True
        
        ThisWorkbook.Sheets("清单差异比对").Activate
        
        MsgBox "与设计核对打包数量与设计清单差异"
        
        Exit Sub
        
    Else
    
        Sheets("清单汇总处理").Delete
        Sheets("清单差异比对").Delete
        
    End If
    
    Sheets.Add(after:=Sheets("设计打包清单")).Name = "非标带配件"
    Sheets.Add(after:=Sheets("设计打包清单")).Name = "非标不带配件"
    Sheets.Add(after:=Sheets("设计打包清单")).Name = "打包分区编号汇总"
    
    Sheets("设计标准件清单").Delete
    Sheets("设计非标件清单").Delete
    
    Application.DisplayAlerts = True
    
    Call 打包清单分类
    Call 拆分到工作簿
    
    Dim hbqdFilename As String
    hbqdFilename = strfile & wjj_name & "\" & wjj_name & "-合并清单.xlsm"
    Call saveHbqd(hbqdFilename)

    Application.ScreenUpdating = True
    
    MsgBox "拆分完毕"
    
End Sub

Private Sub 合并设计传递(MyPath As String)

    Dim Folder As Object, SubFolder As Object
    Dim FileCollection As Object, FileName As Object
    
    '拿一个文件夹所有的文件
    Set Folder = fso.GetFolder(MyPath) 'getfolder返回文件夹对象交给变量 folder
    Set FileCollection = Folder.Files
    
    For Each FileName In FileCollection
    
        If InStr(Split(FileName.Name, ".")(UBound(Split(FileName.Name, "."))), "xl") > 0 Then
    
            i = i + 1: arr(i, 1) = FileName
            设计清单复制
        
        End If
        
    Next
    
    '取文件夹进行递归操作
    For Each SubFolder In Folder.SubFolders
    
        'SubFolders 返回由指定文件夹中所有子文件夹(包括隐藏文件夹和系统文件夹)
        合并设计传递 SubFolder.path '递归
        
    Next
    
End Sub

Private Sub 设计清单复制()

    Dim wb As Workbook
    Dim irow As Integer
    Dim k As Integer
    Dim endb As Integer '打开目录的B列检测最后一行
    Dim enda As Integer
    Dim start_row As Integer
    Dim arra
    Dim endthisa As Integer
    Dim wb_name '打开的工作簿名称
    Dim Target_Sheet As String
    Dim qufj As String '区域附加,对汇总后的
    Dim gzbm As String '工作表名,用来提取工作表中的数字加字母
    Dim bc_qufj As String '变层区域附加,用来标记是否是变层
    Dim p_qufj As String '配件区域附加,用来标记是否有配件
    
    Dim range_target '目标区域,查找"数量"的区域
    Dim r_target '查找"数量"的行数
    Dim c_target '查找"数量"的列数
    Dim czgzbm '记录工作表名
    
    
    If InStr(arr(i, 1), ThisWorkbook.Name) = 0 And InStr(arr(i, 1), "~") = 0 Then  '如果不是缓冲文件就打开
    
        Set wb = Workbooks.Open(arr(i, 1)) '打开表格
        wb_name = Split(wb.FullName, "\")(UBound(Split(wb.FullName, "\")))
        
        If InStr(wb_name, "孔") = 0 And InStr(wb_name, "标准板") + InStr(wb_name, "标准件") > 0 Then
            
            qufj = "BZJ":
            
        ElseIf InStr(wb_name, "孔") > 0 And InStr(wb_name, "标准板") + InStr(wb_name, "标准件") > 0 Then
            
            qufj = "BK"
        
        ElseIf InStr(wb_name, "墙") > 0 And InStr(wb_name, "标准板") + InStr(wb_name, "标准件") = 0 Then
            
            qufj = "Q"
        
        ElseIf InStr(wb_name, "梁") > 0 Then
            
            qufj = "L"
        
        ElseIf InStr(wb_name, "顶板") > 0 Or InStr(wb_name, "楼面") > 0 Then
            
            qufj = "LM"
        
        ElseIf InStr(wb_name, "吊模") > 0 Then
            
            qufj = "DM"
            
        ElseIf InStr(wb_name, "吊架") > 0 Then
            
            qufj = "DJ"
        
        ElseIf InStr(wb_name, "节点") > 0 Then
            
            qufj = "JD"
        
        ElseIf InStr(wb_name, "楼梯") > 0 Then
            
            qufj = "LT"
        
        End If
        
        If InStr(wb.FullName, "带配件") > 0 And InStr(wb.FullName, "不带配件") = 0 Then
            
            p_qufj = "带配件"
        
        End If
        
        If InStr(wb.FullName, "变层") > 0 And InStr(wb.FullName, "基本层") = 0 Then
            
            bc_qufj = "BC"
        
        Else
            
            bc_qufj = ""
        
        End If
        
        '通过工作簿名称知道是那个位置区域的,以方便和工作表名称一起做编号依据
        If InStr(wb.FullName, "打包") > 0 Then
        
            Target_Sheet = "设计打包清单"
            start_row = 3
        
        Else
        
            If InStr(wb_name, "孔") = 0 And InStr(wb_name, "标准板") + InStr(wb_name, "标准件") > 0 Then
                
                Target_Sheet = "设计标准件清单"
            
            Else
                
                Target_Sheet = "设计非标件清单"
            
            End If
            
            start_row = 9
        
        End If
        
        For k = 1 To wb.Worksheets.Count
            
            If InStr(wb.FullName, "打包") = 0 Then
                
                '根据"数量"所在的位置调整行或者列
                czgzbm = Worksheets(k).Name
                
                'Set range_target = wb.Sheets(czgzbm).Range("A1:K9")
                r_target = Sheets(czgzbm).Range("A1:K9").Find(What:="数量").Row
                c_target = Sheets(czgzbm).Range("A1:K9").Find(What:="数量").Column
                
                If r_target = 5 Then
                    
                    wb.Sheets(k).Rows("5:6").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                
                ElseIf r_target = 6 Then
                    
                    wb.Sheets(k).Rows("5:5").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                
                End If
    
                If c_target = 7 Then
                   
                   wb.Sheets(k).Columns("G:G").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
                
                End If
            
            End If
            
            gzbm = wb.Sheets(k).Name '工作表名称
            
            With CreateObject("VBSCRIPT.REGEXP")
                
                .Global = True
                .Pattern = "[^!-~]"
                gzbm = .Replace(gzbm, "") '从工作表名称获取是A区还是B区
            
            End With
            
            irow = ThisWorkbook.Sheets(Target_Sheet).UsedRange.Rows.Count + 1 '获取已使用区域非空的下一行
            endb = wb.Sheets(k).Cells(65535, 2).End(xlUp).Row '
            enda = wb.Sheets(k).Cells(65535, 1).End(xlUp).Row '两侧检测以免数量列的最后一行不是非空单元格
            
            If endb - enda > 2 Then
                
                endb = enda - 1
            
            End If
            
            arra = wb.Sheets(k).Range("A" & start_row & ":J" & endb)  '设计清单标题是8行,合并从第9行开始
            endthisa = ThisWorkbook.Worksheets(Target_Sheet).Cells(Rows.Count, 1).End(xlUp).Row
            ThisWorkbook.Worksheets(Target_Sheet).Range("a" & endthisa + 1).Resize(UBound(arra), 10) = arra
            
            If Len(gzbm) > 0 Then
                
                If gzbm = "()" Or gzbm = "（）" Then
                    
                    gzbm = ""
                
                Else
                    
                    gzbm = "-" & gzbm
                
                End If
            
            Else
            
                gzbm = ""
            
            End If
            

            If Target_Sheet = "设计打包清单" Then
                
                If InStr(wb.FullName, "备用") > 0 Then
                    
                    ThisWorkbook.Worksheets(Target_Sheet).Range("D" & endthisa + 1).Resize(UBound(arra), 1) = qufj & bc_qufj & gzbm & "-(BYJ)"
                
                Else
                    
                    ThisWorkbook.Worksheets(Target_Sheet).Range("D" & endthisa + 1).Resize(UBound(arra), 1) = qufj & bc_qufj & gzbm
                
                End If
            
            Else
                ThisWorkbook.Worksheets(Target_Sheet).Range("k" & endthisa + 1).Resize(UBound(arra), 1) = qufj & bc_qufj & gzbm
                If Len(p_qufj) > 0 Then ThisWorkbook.Worksheets(Target_Sheet).Range("L" & endthisa + 1).Resize(UBound(arra), 1) = p_qufj
            End If
            
        Next k
        
        wb.Close 0
    
    End If
    
End Sub

Private Sub 分出标准件非标件()

    Dim i As Integer '用于遍历第一个设计打包清单中的各个编号
    Dim mbmc As String '模板名称
    Dim enda As Integer
    Dim hangshu As Integer
    Dim k As Integer
    Dim Quyu As String
    Dim arr
    Dim brr
    Dim endb As Integer
    
    '提取图纸编号的辅助列,即去掉前缀以后的部分
    With Sheets("设计非标件清单")
     
        endb = .Cells(Rows.Count, 2).End(xlUp).Row
        
        For i = 2 To endb
            
            If Mid(.Range("C" & i), 2, 1) = "-" Then
                
                .Range("m" & i) = Mid(.Range("C" & i), 3, Len(.Range("C" & i)))
            
            Else
            
                .Range("m" & i) = .Range("C" & i)
            
            End If
        
        Next i
    
    End With

    Range("N1") = "类型"
    Range("N2").FormulaR1C1 = "=VLOOKUP(RC[-12],C[1]:C[2],2,0)"
    Range("N2").AutoFill Destination:=Range("N2:N" & endb)

    Columns("N:N").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Application.CutCopyMode = False

    Sheets("设计打包清单").Activate
    
    brr = Array("序号", "模板名称", "数量", "打包表名", "分区编号", "W1", "W2", "L", "非标图纸编号", "图纸类别", "是否带配件", "辅助列", "生产单类型")
    Sheets("设计打包清单").[A1].Resize(1, UBound(brr) + 1) = brr
    
    enda = Sheets("设计打包清单").Cells(Rows.Count, 1).End(xlUp).Row
    Quyu = ""
    
    For i = 2 To enda
        
        mbmc = Range("B" & i)
        
        '在标准件清单中找设计打包清单中的模板名称,如果找到就标注是标准件,没找到看打包名称和上面的是否一样,一样的话就是编号+1,不一样的话就自己开头
        If Sheets("设计非标件清单").Columns(3).Find(mbmc, LookAt:=xlWhole, SearchDirection:=xlPrevious) Is Nothing Then
            
            If Sheets("设计标准件清单").Columns(3).Find(mbmc, LookAt:=xlWhole, SearchDirection:=xlPrevious) Is Nothing Then
                
                Range("E" & i) = "生产清单中没有"
            
            Else
                
                Range("E" & i) = "标准件"
                
            End If
        
        Else
            
            hangshu = Sheets("设计非标件清单").Columns(3).Find(mbmc, LookAt:=xlWhole, SearchDirection:=xlPrevious).Row
            
            arr = Sheets("设计非标件清单").Range("D" & hangshu & ":F" & hangshu)
            
            Range("F" & i).Resize(1, 3) = arr
            Range("I" & i) = Sheets("设计非标件清单").Range("J" & hangshu)
            Range("J" & i) = Sheets("设计非标件清单").Range("B" & hangshu)
            Range("K" & i) = Sheets("设计非标件清单").Range("L" & hangshu)
            Range("L" & i) = Sheets("设计非标件清单").Range("M" & hangshu)
            Range("M" & i) = Sheets("设计非标件清单").Range("N" & hangshu)
            
        End If
        
        If Len(Range("E" & i)) = 0 Then
        
            If Range("D" & i) = Quyu Then
            
                k = k + 1
                
            Else
            
                k = 1
            
            End If
            
            Range("E" & i) = Range("D" & i) & "-" & k
            Quyu = Range("D" & i).Text
            
        End If
        
    Next
    
End Sub

Private Sub 清单差异比对()
    
    Sheets.Add(after:=Sheets("设计非标件清单")).Name = "清单差异比对"
    Sheets.Add(after:=Sheets("设计非标件清单")).Name = "清单汇总处理"
    
    ThisWorkbook.Sheets("清单差异比对").Activate
    
    Dim krd As Integer
    Dim krh As Integer
    Dim krf As Integer
    Dim krj As Integer
    Dim krk As Integer
    Dim krl As Integer
    Dim cyhangshu As Integer
    Dim dbhzhangshu As Integer
    Dim schzhangshu As Integer
    Dim hdyhangshu As Integer
    Dim mbbh As String
    
    krf = 2
    krj = 2
    krl = 2
    
    Sheets("清单差异比对").Columns("A:A").HorizontalAlignment = Excel.xlCenter
    Sheets("清单差异比对").Columns("B:B").HorizontalAlignment = Excel.xlLeft
    Sheets("清单差异比对").Columns("C:F").HorizontalAlignment = Excel.xlCenter
    Sheets("清单差异比对").Columns("A:F").Font.Name = "宋体"
    Sheets("清单差异比对").Rows("1:65535").RowHeight = 18

    Dim srr
    srr = Array("序号", "模板编号", "打包清单支数", "生产清单支数", "备注")
    Sheets("清单差异比对").[A1].Resize(1, UBound(srr) + 1) = srr
    
    srr = Array("模板编号", "打包清单支数", "", "模板编号", "生产清单支数")
    Sheets("清单汇总处理").[A1].Resize(1, UBound(srr) + 1) = srr
    
    Sheets("清单差异比对").Cells(2, 1).Select
    ActiveWindow.FreezePanes = True
    
    For krd = 2 To Sheets("设计打包清单").Cells(Rows.Count, 1).End(xlUp).Row
    
        If Sheets("设计打包清单").Range("E" & krd).Value = "生产清单中未找到" Then
    
            cyhangshu = krd
            
            Sheets("清单差异比对").Range("A" & krf) = krf - 1
            Sheets("清单差异比对").Range("B" & krf) = Sheets("设计打包清单").Range("B" & cyhangshu)
            Sheets("清单差异比对").Range("C" & krf) = Sheets("设计打包清单").Range("C" & cyhangshu)
            Sheets("清单差异比对").Range("D" & krf) = 0
            Sheets("清单差异比对").Range("E" & krf) = "打包清单中有 生产清单中没有的模板编号"
        
            Sheets("清单差异比对").Range("A" & krf & ":" & "E" & krf).Interior.ColorIndex = 38
            
            krf = krf + 1
        
        Else
        
            dbhzhangshu = krd
        
            Sheets("清单汇总处理").Range("A" & krj) = Sheets("设计打包清单").Range("B" & dbhzhangshu)
            Sheets("清单汇总处理").Range("B" & krj) = Sheets("设计打包清单").Range("C" & dbhzhangshu)
        
            krj = krj + 1
        
        End If
    
    Next krd
    
    For krd = 2 To Sheets("设计标准件清单").Cells(Rows.Count, 1).End(xlUp).Row
    
        schzhangshu = krd
        
        Sheets("清单汇总处理").Range("D" & krl) = Sheets("设计标准件清单").Range("C" & schzhangshu)
        Sheets("清单汇总处理").Range("E" & krl) = Sheets("设计标准件清单").Range("H" & schzhangshu)
        
        krl = krl + 1
    
    Next krd
    
    For krd = 2 To Sheets("设计非标件清单").Cells(Rows.Count, 1).End(xlUp).Row
        
        schzhangshu = krd
        
        Sheets("清单汇总处理").Range("D" & krl) = Sheets("设计非标件清单").Range("C" & schzhangshu)
        Sheets("清单汇总处理").Range("E" & krl) = Sheets("设计非标件清单").Range("H" & schzhangshu)
        
        krl = krl + 1
        
    Next krd

    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "清单汇总处理!R1C1:R" & (krj - 1) & "C2", Version:=xlPivotTableVersion10).CreatePivotTable _
        TableDestination:="清单汇总处理!R1C7", TableName:="打包清单汇总透视表", DefaultVersion:= _
        xlPivotTableVersion10
        
    Sheets("清单汇总处理").PivotTables("打包清单汇总透视表").AddFields RowFields:=Array("模板编号")
    
    With Sheets("清单汇总处理").PivotTables("打包清单汇总透视表")
        
        .AddDataField .PivotFields("打包清单支数"), " 数量", xlSum
    
    End With
    
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "清单汇总处理!R1C4:R" & (krl - 1) & "C5", Version:=xlPivotTableVersion10).CreatePivotTable _
        TableDestination:="清单汇总处理!R1C10", TableName:="生产清单汇总透视表", DefaultVersion:= _
        xlPivotTableVersion10
        
    Sheets("清单汇总处理").PivotTables("生产清单汇总透视表").AddFields RowFields:=Array("模板编号")
    
    With Sheets("清单汇总处理").PivotTables("生产清单汇总透视表")
        
        .AddDataField .PivotFields("生产清单支数"), " 数量", xlSum
    
    End With
    
    For krd = 3 To Sheets("清单汇总处理").Cells(Rows.Count, 10).End(xlUp).Row - 1
    
    mbbh = Sheets("清单汇总处理").Range("J" & krd)
    
    If Sheets("清单汇总处理").Columns(7).Find(mbbh, LookAt:=xlWhole, SearchDirection:=xlPrevious) Is Nothing Then
        
        Sheets("清单差异比对").Range("A" & krf) = krf - 1
        Sheets("清单差异比对").Range("B" & krf) = Sheets("清单汇总处理").Range("J" & krd)
        Sheets("清单差异比对").Range("C" & krf) = 0
        Sheets("清单差异比对").Range("D" & krf) = Sheets("清单汇总处理").Range("K" & krd)
        Sheets("清单差异比对").Range("E" & krf) = "生产清单中有 打包清单中没有的模板编号"
        
        Sheets("清单差异比对").Range("A" & krf & ":" & "E" & krf).Interior.ColorIndex = 36
        krf = krf + 1
    
    Else
        
        hdyhangshu = Sheets("清单汇总处理").Columns(7).Find(mbbh, LookAt:=xlWhole, SearchDirection:=xlPrevious).Row
        
        If Sheets("清单汇总处理").Range("H" & hdyhangshu) <> Sheets("清单汇总处理").Range("K" & krd) Then
        
            Sheets("清单差异比对").Range("A" & krf) = krf - 1
            Sheets("清单差异比对").Range("B" & krf) = Sheets("清单汇总处理").Range("J" & krd)
            Sheets("清单差异比对").Range("C" & krf) = Sheets("清单汇总处理").Range("H" & hdyhangshu)
            Sheets("清单差异比对").Range("D" & krf) = Sheets("清单汇总处理").Range("K" & krd)
            Sheets("清单差异比对").Range("E" & krf) = "打包清单与生产清单支数不符"
        
            Sheets("清单差异比对").Range("A" & krf & ":" & "E" & krf).Interior.ColorIndex = 37
            krf = krf + 1
        
        End If
        
    End If
    
    Next krd
    
End Sub

Private Sub 打包清单分类()

    Dim cnn As Object, rs As Object
    
    Set cnn = CreateObject("adodb.connection")
    Set rs = CreateObject("adodb.recordset")
    
    Dim SQL As String
    Dim a As Long
    Dim title_arr
    Dim i As Integer
    Dim endb As Integer
    Dim k As Integer
    
    cnn.Open "provider=Microsoft.ACE.OLEDB.12.0;extended properties='excel 12.0 Macro;hdr=yes';data source=" & ThisWorkbook.FullName

    SQL = "select 模板名称,数量,打包表名,分区编号,是否带配件 from [设计打包清单$] where 分区编号<>'标准件'and 分区编号<>'生产清单中没有' "
    
    Set rs = cnn.Execute(SQL)
    Sheets("打包分区编号汇总").Range("B2").CopyFromRecordset rs
    
    title_arr = Array("序号", "模板编号", "数量", "打包表名", "分区编号", "是否带配件", "备注")
    Sheets("打包分区编号汇总").[A1].Resize(1, UBound(title_arr) + 1) = title_arr
    
    rs.Close: Set rs = Nothing
    
    With Sheets("打包分区编号汇总")
        
        endb = .Cells(Rows.Count, 2).End(xlUp).Row
        
        For i = 2 To endb
            
            .Range("A" & i) = i - 1
        
        Next i
        
        .Range("A1:G" & endb).Interior.Pattern = xlNone
        .Range("A1:G" & endb).Borders.Weight = 2
        .Columns("A:G").HorizontalAlignment = xlCenter
        .Columns("B:B").EntireColumn.AutoFit
    
    End With
    
    Sheets("打包分区编号汇总").Move
    
    Dim dbnum  As String '打包num,即打包分区编号汇总移动出来后新的表格名字
    
    ' dbnum = Replace(ThisWorkbook.Name, "合并清单.xlsm", "打包分区编号汇总.xlsx")
    
    ' ActiveWorkbook.SaveAs FileName:=ThisWorkbook.path & "\" & dbnum
    ActiveWorkbook.SaveAs FileName:=ThisWorkbook.path & "\打包分区编号汇总.xlsx"
    ActiveWorkbook.Close
    
    '先对W1,W2做一下调整
    Dim W1_num As Integer
    Dim W2_num As Integer
    
    Sheets("非标不带配件").Activate
    
    SQL = "select *  from [设计打包清单$] where 分区编号<>'标准件'and 分区编号<>'生产清单中没有'and 是否带配件 is null order by 生产单类型,模板名称,W1,W2,辅助列,非标图纸编号"
    
    Set rs = cnn.Execute(SQL)
    Sheets("非标不带配件").Range("A2").CopyFromRecordset rs
    
    rs.Close: Set rs = Nothing
    
    title_arr = Array("序号", "模板名称", "模板编号", "数量", "W1", "W2", "L", "图纸编号", "分区编号", "辅助列", "生产单类型")
    
    With Sheets("非标不带配件")
        
        .Columns("J:J").Cut
        .Columns("B:B").Insert Shift:=xlToRight
        .Columns("A:A").ClearContents
        .Columns("F:F").Cut
        .Columns("K:K").Insert Shift:=xlToRight
        .Columns("E:E").Delete Shift:=xlToLeft
        .Columns("j:j").Delete Shift:=xlToLeft
        .[A1].Resize(1, UBound(title_arr) + 1) = title_arr
        
        endb = .Cells(Rows.Count, 2).End(xlUp).Row
        
        For i = 2 To endb
             
            If InStr(Range("B" & i), "C槽") + InStr(Range("B" & i), "转角") > 0 Then
                
                If .Range("F" & i) = 100 Then
                    
                    W2_num = .Range("E" & i)
                    .Range("E" & i) = 100
                    .Range("f" & i) = W2_num
                
                ElseIf .Range("F" & i) = 150 And .Range("E" & i) <> 100 Then
                    
                    W2_num = .Range("E" & i)
                    .Range("E" & i) = 150
                    .Range("F" & i) = W2_num
                
                Else
                
                    W1_num = .Range("E" & i)
                    W2_num = .Range("F" & i)
                    .Range("E" & i) = W1_num
                    .Range("F" & i) = W2_num
                    
                End If
              
            ElseIf InStr(Range("B" & i), "角铝") > 0 And Len(Range("F" & i)) > 0 Then

                If .Range("F" & i) = 65 And .Range("E" & i) <> 65 Then
                    
                    W2_num = .Range("E" & i)
                    .Range("E" & i) = 65
                    .Range("F" & i) = W2_num
                
                Else
                    
                    W1_num = .Range("E" & i)
                    W2_num = .Range("F" & i)
                    .Range("E" & i) = W1_num
                    .Range("F" & i) = W2_num
                    
                End If
                
            End If
            
            .Range("A" & i) = i - 1
           
        Next i
        
    End With
    
    '对现有内容进行排序
    With Sheets("非标不带配件").Sort.SortFields
        
        .Clear
        .Add Key:=Range("K2"), Order:=1 '生产单类型
        .Add Key:=Range("B2"), Order:=1 '模板名称
        .Add Key:=Range("E2"), Order:=1 'W1
        .Add Key:=Range("F2"), Order:=1 'W2
        .Add Key:=Range("H2"), Order:=1 '图纸编号
        .Add Key:=Range("J2"), Order:=1 '辅助列
    
    End With

    With Sheets("非标不带配件").Sort
        
        .SetRange Range("b2:L" & endb)
        .Header = 2 '没有标题
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
         
    End With
    
    k = 0
    
    With Sheets("非标不带配件")
    
        '添加颜色填一个彩云色,希望给我们工作带来好心情,夏天时候可以将余数由0,5,4,3,2,1排列有一种清爽的感觉
        For i = 2 To endb
            
            If .Range("B" & i) <> .Range("B" & i - 1) Then k = k + 1
            
            If k Mod 6 = 1 Then
                
                .Range("A" & i & ":K" & i).Interior.Color = RGB(248, 230, 158)
            
            ElseIf k Mod 6 = 2 Then
                
                .Range("A" & i & ":K" & i).Interior.Color = RGB(232, 159, 187)
            
            ElseIf k Mod 6 = 3 Then
                
                .Range("A" & i & ":K" & i).Interior.Color = RGB(6, 103, 163)
            
            ElseIf k Mod 6 = 4 Then
                
                .Range("A" & i & ":K" & i).Interior.Color = RGB(28, 140, 185)
            
            ElseIf k Mod 6 = 5 Then
                
                .Range("A" & i & ":K" & i).Interior.Color = RGB(126, 202, 221)
            
            ElseIf k Mod 6 = 0 Then
                
                .Range("A" & i & ":K" & i).Interior.Color = RGB(221, 236, 242)
            
            End If
        
        Next
    
    End With
    
    Sheets("非标不带配件").Columns("J:J").Delete Shift:=xlToLeft
    
    Sheets("非标带配件").Activate

    SQL = "select *  from [设计打包清单$] where 分区编号<>'标准件'and 分区编号<>'生产清单中没有'and 是否带配件='带配件' order by 生产单类型,模板名称,W1,W2,辅助列,非标图纸编号"
    
    Set rs = cnn.Execute(SQL)
    Sheets("非标带配件").Range("A2").CopyFromRecordset rs
    
    rs.Close: Set rs = Nothing
    
    With Sheets("非标带配件")
        
        .Columns("K:K").ClearContents
        .Columns("J:J").Cut
        .Columns("B:B").Insert Shift:=xlToRight
        .Columns("A:A").ClearContents
        .Columns("F:F").Cut
        .Columns("K:K").Insert Shift:=xlToRight
        .Columns("E:E").Delete Shift:=xlToLeft
        .[A1].Resize(1, UBound(title_arr) + 1) = title_arr
        
        endb = .Cells(Rows.Count, 2).End(xlUp).Row
        
        For i = 2 To endb
             
            If InStr(Range("B" & i), "C槽") + InStr(Range("B" & i), "转角") > 0 Then
                
                If .Range("F" & i) = 100 Then
                    
                    W2_num = .Range("E" & i)
                    .Range("E" & i) = 100
                    .Range("f" & i) = W2_num
                
                ElseIf .Range("F" & i) = 150 And .Range("E" & i) <> 100 Then
                    
                    W2_num = .Range("E" & i)
                    .Range("E" & i) = 150
                    .Range("F" & i) = W2_num
                
                Else
                
                    W1_num = .Range("E" & i)
                    W2_num = .Range("F" & i)
                    .Range("E" & i) = W1_num
                    .Range("F" & i) = W2_num
                    
                End If
             
            ElseIf InStr(Range("B" & i), "角铝") > 0 And Len(Range("F" & i)) > 0 Then

                If .Range("F" & i) = 65 And .Range("E" & i) <> 65 Then
                    
                    W2_num = .Range("E" & i)
                    .Range("E" & i) = 65
                    .Range("F" & i) = W2_num
                
                Else
                
                    W1_num = .Range("E" & i)
                    W2_num = .Range("F" & i)
                    .Range("E" & i) = W1_num
                    .Range("F" & i) = W2_num
                    
                End If
            
            End If
            
            .Range("A" & i) = i - 1
            
        Next i
        
    End With
    
     '对现有内容进行排序
    With Sheets("非标带配件").Sort.SortFields
        
        .Clear
        .Add Key:=Range("B2"), Order:=1 '模板名称
        .Add Key:=Range("E2"), Order:=1 'W1
        .Add Key:=Range("F2"), Order:=1 'W2
        .Add Key:=Range("H2"), Order:=1 '辅助列
        .Add Key:=Range("K2"), Order:=1 '辅助列
    
    End With

    With Sheets("非标带配件").Sort
        
        .SetRange Range("b2:K" & endb)
        .Header = 2 '没有标题
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
         
    End With
    
    k = 0
    
    With Sheets("非标带配件")
    
        .Range("J1") = "生产单类型"
    
        For i = 2 To endb
        
            If .Range("B" & i) <> .Range("B" & i - 1) Then k = k + 1
            
            If k Mod 6 = 1 Then
                
                .Range("A" & i & ":K" & i).Interior.Color = RGB(248, 230, 158)
            
            ElseIf k Mod 6 = 2 Then
                
                .Range("A" & i & ":K" & i).Interior.Color = RGB(232, 159, 187)
            
            ElseIf k Mod 6 = 3 Then
                
                .Range("A" & i & ":K" & i).Interior.Color = RGB(6, 103, 163)
            
            ElseIf k Mod 6 = 4 Then
                
                .Range("A" & i & ":K" & i).Interior.Color = RGB(28, 140, 185)
            
            ElseIf k Mod 6 = 5 Then
                
                .Range("A" & i & ":K" & i).Interior.Color = RGB(126, 202, 221)
            
            ElseIf k Mod 6 = 0 Then
                
                .Range("A" & i & ":K" & i).Interior.Color = RGB(221, 236, 242)
            
            End If
            
            .Range("J" & i) = "TP"
        
        Next
    
    End With
    
    Sheets("非标带配件").Columns("K:L").Delete Shift:=xlToLeft
    cnn.Close: Set cnn = Nothing
    
    '将各种类型的生产单类型及模板名称列出表格
    Sheets("非标不带配件").Activate
    
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "非标不带配件!R1C2:R1048576C10", Version:=xlPivotTableVersion14).CreatePivotTable TableDestination:= _
        "非标不带配件!R1C15", TableName:="数据透视表1", DefaultVersion:=xlPivotTableVersion14
    
    Sheets("非标不带配件").Cells(1, 15).Select
    
    ActiveWorkbook.ShowPivotTableFieldList = True
    
    With ActiveSheet.PivotTables("数据透视表1").PivotFields("生产单类型")
        
        .Orientation = xlRowField
        .Position = 1
    
    End With
    
    With ActiveSheet.PivotTables("数据透视表1").PivotFields("模板名称")
        
        .Orientation = xlRowField
        .Position = 2
    
    End With
    
    ActiveSheet.PivotTables("数据透视表1").AddDataField ActiveSheet.PivotTables("数据透视表1" _
        ).PivotFields("数量"), "总计数量", xlSum
    
    ActiveWorkbook.ShowPivotTableFieldList = False
    
    Range("O4").Select
    
    ActiveSheet.PivotTables("数据透视表1").RowAxisLayout xlTabularRow
    ActiveSheet.PivotTables("数据透视表1").PivotFields("模板名称").Subtotals = Array(False, _
        False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("数据透视表1").PivotFields("生产单类型").Subtotals = Array(False, _
        False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("数据透视表1").RepeatAllLabels xlRepeatLabels
    
    Columns("O:Q").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    Dim end_O As Integer
    end_O = Cells(Rows.Count, 15).End(xlUp).Row - 1
    Range("O" & end_O & ": Q" & end_O).ClearContents
    
    Sheets("非标带配件").Activate
    
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "非标带配件!R1C2:R1048576C10", Version:=xlPivotTableVersion14).CreatePivotTable TableDestination:= _
        "非标带配件!R1C15", TableName:="数据透视表2", DefaultVersion:=xlPivotTableVersion14
    
    Sheets("非标带配件").Select
    Cells(1, 15).Select
    
    ActiveWorkbook.ShowPivotTableFieldList = True
     
    With ActiveSheet.PivotTables("数据透视表2").PivotFields("生产单类型")
        
        .Orientation = xlRowField
        .Position = 1
    
    End With
    
    With ActiveSheet.PivotTables("数据透视表2").PivotFields("模板名称")
        
        .Orientation = xlRowField
        .Position = 2
    
    End With
    
    ActiveSheet.PivotTables("数据透视表2").AddDataField ActiveSheet.PivotTables("数据透视表2" _
        ).PivotFields("数量"), "总计数量", xlSum
    
    ActiveWorkbook.ShowPivotTableFieldList = False
    
    Range("O4").Select
    
    ActiveSheet.PivotTables("数据透视表2").RowAxisLayout xlTabularRow
    ActiveSheet.PivotTables("数据透视表2").PivotFields("模板名称").Subtotals = Array(False, _
        False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("数据透视表2").PivotFields("生产单类型").Subtotals = Array(False, _
        False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("数据透视表2").RepeatAllLabels xlRepeatLabels
    
    Columns("O:Q").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    end_O = Cells(Rows.Count, 15).End(xlUp).Row - 1
    Range("O" & end_O & ": Q" & end_O).ClearContents
    
End Sub

Private Sub 拆分到工作簿()

    Dim ary(), arr, brr, sh As Worksheet, d As Object, k, t, a, i&, j&, m&, l&
    Dim arr1, k1, T1
    Dim heji As Integer
    Dim path As String
    Dim ws As Worksheet

    path = ThisWorkbook.path & "\分配清单"

    If Dir(path, vbDirectory) = "" Then
        
        MkDir path
        
    End If
    
    Application.DisplayAlerts = False
    
    For Each ws In Worksheets
        
        If Left(ws.Name, 2) <> "设计" And (ws.Name <> "Sheet1") And (ws.Name <> "库(待补充)") Then
            
            Set d = CreateObject("scripting.dictionary")
            
            arr = ws.[A1].CurrentRegion
            
            ReDim ary(1 To 200000, 1 To UBound(arr, 2))
            
            For i = 2 To UBound(arr)
                
                m = m + 1
                d(arr(i, 10)) = d(arr(i, 10)) & "," & m
                
                For j = 1 To UBound(arr, 2)
                    
                    ary(m, j) = arr(i, j)
                
                Next
            
            Next
            
            k = d.Keys
            t = d.Items
            
            brr = [A1].Resize(65536, UBound(arr, 2))
            
            For i = 0 To d.Count - 1
                
                m = 1
                a = Split(t(i), ",")
                
                For j = 1 To UBound(a)
                    
                    m = m + 1
                    
                    For l = 1 To UBound(arr, 2)
                        
                        brr(m, l) = ary(a(j), l)
                        
                    Next
                    
                    heji = brr(m, 4) + heji
                
                Next
                
                With Workbooks.Add(xlWBATWorksheet)
                    
                    With .Sheets(1).[A1].Resize(m, UBound(brr, 2))
                        
                        .Value = brr
                        .Borders.LineStyle = xlContinuous
                        .EntireColumn.AutoFit
                    
                    End With
                    
                    .SaveAs FileName:=ThisWorkbook.path & "\分配清单\" & Replace(k(i), Chr(9), "") & "-" & heji & ".xlsx"
                    .Close
                    
                    heji = 0
                
                End With
            
            Next
            
            Set d = Nothing
            
            m = 0
            
            Erase arr
            Erase brr
            
            Set k = Nothing
            Set t = Nothing
        End If
    
    Next
   
    Application.DisplayAlerts = True

End Sub

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

Private Sub copySheet(src As Worksheet, dst As Worksheet)
    Dim ur As Range
    Dim rowCount As Long
    Dim ColumnCount As Long
    Set ur = src.UsedRange
    ColumnCount = ur.Columns.Count
    rowCount = ur.Rows.Count

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

' Private Sub saveHbqd(filename As String)
Sub saveHbqd(FileName As String)
    'Dim filename As String
    'filename = "C:\Users\u03013112\Documents\new-412-1\a.xlsx"
    If fileIsExist(FileName) Then
    ' TODO 询问是否要清除重来
    Else
        Call createExcel(FileName)
    End If

    Dim thisWb As Workbook
    Dim wb As Workbook
    Set thisWb = ActiveWorkbook
    Set wb = Workbooks.Open(FileName)
    
   ' 暂时只复制3个表，其他的貌似不需要
   Application.DisplayAlerts = False
    If isSheetExist(wb, "设计打包清单") Then
        wb.Sheets("设计打包清单").Delete
    End If
    wb.Sheets.Add().Name = "设计打包清单"
    Call copySheet(thisWb.Sheets("设计打包清单"), wb.Sheets("设计打包清单"))

    If isSheetExist(wb, "非标不带配件") Then
        wb.Sheets("非标不带配件").Delete
    End If
    wb.Sheets.Add().Name = "非标不带配件"
    Call copySheet(thisWb.Sheets("非标不带配件"), wb.Sheets("非标不带配件"))

    If isSheetExist(wb, "非标带配件") Then
        wb.Sheets("非标带配件").Delete
    End If
    wb.Sheets.Add().Name = "非标带配件"
    Call copySheet(thisWb.Sheets("非标带配件"), wb.Sheets("非标带配件"))
    Application.DisplayAlerts = True
End Sub




