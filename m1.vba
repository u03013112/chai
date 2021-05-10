Option Explicit

' 整体思路再次更改，仍旧需要在本文件中进行主要数据操作，保证各种透视表能够正常工作，然后将需要的表格保存到指定文件
' 中间拆分的过程仍然用新的这版，保障中间上下文明了

' TODO:search的写法有提高空间，目前每次查找都多查了至少一遍
' TODO:Dbqdfl中间替代了数据库算法，改为数组遍历，排序与原版不同
' TODO:可以把过程再拆一拆，让他自己GC，以免内存不足报错

Function dirIsExist(dirFullPath As String) As Boolean
 Dim fso As Object
 Dim ret As Boolean
 Set fso = CreateObject("Scripting.FileSystemObject")
 ret = False
 If fso.FolderExists(dirFullPath) = True Then
     ret = True
 End If
  Set fso = Nothing
  dirIsExist = ret
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

Sub Log(shtName As String, cellName As String, str As String,color As Long)
    Sheets(shtName).Range(cellName) = str
    Sheets(shtName).Range(cellName).Interior.Color = color
End Sub

' TODO:在打开的时候把临时状态都清了
Sub statusInit()
    ThisWorkbook.Sheets("main").Range("D2") = "等待开始"
    ThisWorkbook.Sheets("main").Range("D3:D13").ClearContents
    ThisWorkbook.Sheets("main").Range("D3:D13").Interior.Color = RGB(240, 240, 0)
End Sub

Sub chose1()
    Dim fso As Object, arr(1 To 10 ^ 2, 1 To 1), i
    Dim dg As FileDialog

    Dim strfile As String
    Dim brr
    Dim fileFolderName

    Dim outputDir As String
    Dim hbqdFilename As String  ' 合并清单文件名
    Dim dbfqhzFilename As String    ' 打包分区编号汇总文件名
    Dim qdcyFilename As String ' 清单差异文件名
    Dim fpqdDirname As String  ' 分配清单目录名

    Call Log("main", "D2", "请等待..",RGB(240,240,0))
    Set dg = Application.FileDialog(msoFileDialogFolderPicker)
    If dg.Show = -1 Then
        '递归所选目录，找到所有excel文件
        Dim excelFilenames As Variant
        excelFilenames = getAllExcelFile(dg.SelectedItems(1))
        ' [f5] = getArrLen(excelFilenames)
        'TODO 检测找到的文件是否合格

        '在本地进行临时处理
        Call Log("main", "D3", "已选择目录:" & dg.SelectedItems(1),RGB(0,240,0))
        fileFolderName = Split(dg.SelectedItems(1), "\")(UBound(Split(dg.SelectedItems(1), "\")))
        outputDir = ThisWorkbook.path & "\" & fileFolderName
        hbqdFilename = outputDir & "\" & fileFolderName & "-合并清单.xlsx"
        dbfqhzFilename = outputDir & "\" & fileFolderName & "-打包分区编号汇总.xlsx"
        qdcyFilename = outputDir & "\" & fileFolderName & "-清单差异.xlsx"
        fpqdDirname = outputDir & "\分配清单\"
        If dirIsExist(outputDir) = True Then
            Dim result
            Call Log("main", "D4", "存有旧状态，或已完成文件，需要清理才能继续",RGB(240,240,0))
            result = MsgBox("检测到有同名项目已存在，是否删除重做？", 4, "选择否将中断拆图")
            If result = vbNo Then Exit Sub
            
            CreateObject("scripting.filesystemobject").GetFolder(outputDir).Delete True
            ' delDIr (outputDir)
            Call Log("main", "D4", "清理完成",RGB(0,240,0))
        End If
        
        VBA.MkDir (outputDir)
        '提前建立合并清单文件
        createExcel (hbqdFilename)
        
        ThisWorkbook.Activate
        ' 拆分步骤，每一步都相对独立
        Call HbqdStep1(hbqdFilename, excelFilenames)
        Call HbqdStep2(hbqdFilename)
        Call HbqdStep3(hbqdFilename, qdcyFilename)

        Call Dbqdfl(hbqdFilename, dbfqhzFilename)
        Call Cfdgzb(hbqdFilename, fpqdDirname)
        
        Call Log("main", "D2", "已完成",RGB(0,240,0))
        Exit Sub
    Else
        Exit Sub
    End If
    ' MsgBox "拆分完毕"
End Sub

Sub HbqdStep1(hbqdFilename As String, excelFilenames As Variant)
    Dim wb As Workbook
    Set wb = Workbooks.Open(hbqdFilename)
    wb.Windows(1).Visible = True
    ThisWorkbook.Activate

    wb.Sheets.Add().Name = "设计非标件清单"
    wb.Sheets.Add().Name = "设计标准件清单"
    wb.Sheets.Add().Name = "设计打包清单"
    
    Application.DisplayAlerts = True
         
    Dim brr
    brr = Array("序号", "模板名称", "模板编号", "W1", "W2", "L", "单件面积", "数量", "总件面积", "图纸编号", "工作表名", "是否带配件")
    wb.Sheets("设计标准件清单").[A1].Resize(1, UBound(brr) + 1) = brr
    wb.Sheets("设计非标件清单").[A1].Resize(1, UBound(brr) + 1) = brr
        
    brr = Array("序号", "模板名称", "数量", "打包表名")
    wb.Sheets("设计打包清单").[A1].Resize(1, UBound(brr) + 1) = brr

    Call Log("main", "D5", "共检测到" & getArrLen(excelFilenames) & "个excel文件",RGB(0,240,0))
    Dim excelFilename As Variant
    Dim count As Long
    count = 1
    For Each excelFilename In excelFilenames
        If excelFilename = Empty Then
            Exit For
        End If
        Call Log("main", "D6", "正在处理第" & count & "个文件：" & excelFilename,RGB(240,240,0))
        Application.ScreenUpdating = False
        Call SjqdCopy(CStr(excelFilename), wb)
        Application.ScreenUpdating = True
        count = count + 1
    Next
    Call Log("main", "D6", "已完成",RGB(0,240,0))
    wb.Windows(1).Visible = True
    wb.Close (True)
End Sub

Sub HbqdStep2(hbqdFilename As String)
    Dim wb As Workbook
    Set wb = Workbooks.Open(hbqdFilename)
    wb.Windows(1).Visible = True
    ThisWorkbook.Activate

    Dim endb As Integer
    Dim i_mbmc  As Integer '遍历模板名称的遍历字符
    With wb.Sheets("设计非标件清单")
        endb = .Cells(65535, 2).End(xlUp).Row
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
        .Tab.ColorIndex = 3
        .Range("O:O") = .Range("B:B").Value
        .Range("$O$1:$O$" & endb).RemoveDuplicates Columns:=1, Header:=xlNo
        .Columns("O:P").EntireColumn.AutoFit
        Dim end_O As Integer
        end_O = .Range("O6000").End(xlUp).Row
        Dim i
        Dim mbmc As String 'o列的模板名称
        Dim scdmc As String '生产单名称
        Dim hangshu As Integer
        For i = 1 To end_O
            mbmc = .Range("O" & i) '型材宽度
            If ThisWorkbook.Sheets("库(待补充)").Columns(4).Find(mbmc, LookAt:=xlWhole, SearchDirection:=xlPrevious) Is Nothing Then
                scdmc = "QT"
            Else
                hangshu = ThisWorkbook.Sheets("库(待补充)").Columns(4).Find(mbmc, LookAt:=xlWhole, SearchDirection:=xlPrevious).Row
                scdmc = ThisWorkbook.Sheets("库(待补充)").Range("E" & hangshu) '生产单命名
            End If
            .Range("P" & i) = scdmc
        Next
    End With
    wb.Windows(1).Visible = True
    wb.Close (True)
End Sub

Sub HbqdStep3Test()
    Dim hbqdFilename As String
    Dim qdcyFilename As String
    hbqdFilename = "C:\Users\u03013112\Documents\J\new-412-1\new-412-1-合并清单.xlsx"
    qdcyFilename = "C:\Users\u03013112\Documents\J\new-412-1\new-412-1-清单差异.xlsx"

    Call HbqdStep3(hbqdFilename, qdcyFilename)
End Sub

Sub HbqdStep4Test()
    Dim hbqdFilename As String
    Dim dbfqhzFilename As String
    hbqdFilename = "C:\Users\u03013112\Documents\J\new-412-1\new-412-1-合并清单.xlsx"
    dbfqhzFilename = "C:\Users\u03013112\Documents\J\new-412-1\new-412-1-打包分区编号汇总.xlsx"

    Call Dbqdfl(hbqdFilename, dbfqhzFilename)
End Sub

Sub HbqdStep5Test()
    Dim hbqdFilename As String
    Dim fpqdDirname As String
    hbqdFilename = "C:\Users\u03013112\Documents\J\new-412-1\new-412-1-合并清单.xlsx"
    fpqdDirname = "C:\Users\u03013112\Documents\J\new-412-1\分配清单\"

    Call Cfdgzb(hbqdFilename, fpqdDirname)
End Sub

' TODO: 完全不再使用透视表，使用Dict替代透视表
Sub HbqdStep3(hbqdFilename As String, qdcyFilename As String)
    Call Log("main", "D7", "开始检查数据，核对打包清单",RGB(240,240,0))
    Application.DisplayAlerts = False
    

    Dim wb As Workbook
    Set wb = Workbooks.Open(hbqdFilename)
    wb.Windows(1).Visible = True
    ThisWorkbook.Activate

    Call StdOrNoStd(wb)
    Application.ScreenUpdating = False
    Call QdDiff(wb)
    Application.DisplayAlerts = False
    If wb.Sheets("清单差异比对").Cells(65535, 1).End(xlUp).Row > 1 Then
        wb.Sheets("清单汇总处理").Delete
        wb.Worksheets("清单差异比对").Columns("A:E").EntireColumn.AutoFit
        Worksheets(Array("设计打包清单", "设计标准件清单", "设计非标件清单", "清单差异比对")).Copy
        ActiveWorkbook.SaveAs filename:=qdcyFilename, FileFormat:=51
        ActiveWorkbook.Close SaveChanges:=True
        wb.Sheets("清单差异比对").Activate
        MsgBox "与设计核对打包数量与设计清单差异"
        Exit Sub
    Else
        'wb.Sheets("清单汇总处理").Delete
        'wb.Sheets("清单差异比对").Delete
    End If
    wb.Sheets.Add().Name = "非标带配件"
    wb.Sheets.Add().Name = "非标不带配件"
    wb.Sheets.Add().Name = "打包分区编号汇总"
    'wb.Sheets("设计标准件清单").Delete
    'wb.Sheets("设计非标件清单").Delete
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    wb.Close (True)
End Sub

Sub createExcel(fileFullPath As String)
    Dim excelApp, excelWB As Object
    Dim savePath, saveName As String

    Set excelApp = CreateObject("Excel.Application")
    Set excelWB = excelApp.Workbooks.Add

    excelWB.SaveAs fileFullPath
    excelApp.Quit
End Sub

Private Function getAllExcelFile(MyPath As String) As Variant
    Dim arr(1 To 300)
    Dim arrTmp As Variant
    Dim i As Long
    Dim Folder As Object, SubFolder As Object
    Dim FileCollection As Object
    Dim filename As Variant
    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set Folder = fso.GetFolder(MyPath)
    Set FileCollection = Folder.Files
    
    i = 1
    For Each filename In FileCollection
        If InStr(Split(filename.Name, ".")(UBound(Split(filename.Name, "."))), "xl") > 0 And InStr(filename.Name, "~") = 0 Then
            arr(i) = filename
            i = i + 1
        End If
    Next
    For Each SubFolder In Folder.SubFolders
        arrTmp = getAllExcelFile(SubFolder.path) '递归
        For Each filename In arrTmp
            If filename = Empty Then
                Exit For
            End If
            arr(i) = filename
            i = i + 1
        Next
    Next
    getAllExcelFile = arr
End Function

Private Function getArrLen(arr As Variant) As Long
    Dim v As Variant
    Dim i As Long
    i = 1
    For Each v In arr
        If arr(i) = Empty Then
            i = i - 1
            Exit For
        End If
        i = i + 1
    Next
    getArrLen = i
End Function

Private Function delDIr(MyPath As String)
    Dim Folder As Object, SubFolder As Object
    Dim FileCollection As Object
    Dim filename As Variant
    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set Folder = fso.GetFolder(MyPath)
    Set FileCollection = Folder.Files
    
    For Each filename In FileCollection
        Kill filename
    Next
    For Each SubFolder In Folder.SubFolders
        delDIr (SubFolder.path) '递归
        RmDir SubFolder.path
    Next
End Function

'设计清单复制 ：沿用了旧名字，不明白意义，不改名
Private Sub SjqdCopy(filename As String, wbTarget As Workbook)
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
    
    Set wb = Workbooks.Open(filename) '打开表格
    wb.Windows(1).Visible = True
    wb_name = Split(wb.FullName, "\")(UBound(Split(wb.FullName, "\")))
    '区域附加
    If InStr(wb_name, "孔") = 0 And InStr(wb_name, "标准板") + InStr(wb_name, "标准件") > 0 Then
        qufj = "BZJ"
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
     '配件区域附加
    If InStr(wb.FullName, "带配件") > 0 And InStr(wb.FullName, "不带配件") = 0 Then
        p_qufj = "带配件"
    End If
    '变层区域附加
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
    
    For k = 1 To wb.Worksheets.count
        If InStr(wb.FullName, "打包") = 0 Then
            '根据"数量"所在的位置调整行或者列
            czgzbm = wb.Worksheets(k).Name
            'Set range_target = wb.Sheets(czgzbm).Range("A1:K9")
            r_target = wb.Sheets(czgzbm).Range("A1:K9").Find(What:="数量").Row
            c_target = wb.Sheets(czgzbm).Range("A1:K9").Find(What:="数量").Column
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
        
        irow = wbTarget.Sheets(Target_Sheet).UsedRange.Rows.count + 1 '获取已使用区域非空的下一行
        endb = wb.Sheets(k).Cells(wb.Sheets(k).Rows.count, 2).End(xlUp).Row '
        enda = wb.Sheets(k).Cells(wb.Sheets(k).Rows.count, 1).End(xlUp).Row '两侧检测以免数量列的最后一行不是非空单元格
        
        If endb - enda > 2 Then
            endb = enda - 1
        End If
        
        arra = wb.Sheets(k).Range("A" & start_row & ":J" & endb)  '设计清单标题是8行,合并从第9行开始
        endthisa = wbTarget.Worksheets(Target_Sheet).Cells(65535, 1).End(xlUp).Row
        wbTarget.Worksheets(Target_Sheet).Range("a" & endthisa + 1).Resize(UBound(arra), 10) = arra
        
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
                wbTarget.Worksheets(Target_Sheet).Range("D" & endthisa + 1).Resize(UBound(arra), 1) = qufj & bc_qufj & gzbm & "-(BYJ)"
            Else
                wbTarget.Worksheets(Target_Sheet).Range("D" & endthisa + 1).Resize(UBound(arra), 1) = qufj & bc_qufj & gzbm
            End If
        Else
            wbTarget.Worksheets(Target_Sheet).Range("k" & endthisa + 1).Resize(UBound(arra), 1) = qufj & bc_qufj & gzbm
            If Len(p_qufj) > 0 Then wbTarget.Worksheets(Target_Sheet).Range("L" & endthisa + 1).Resize(UBound(arra), 1) = p_qufj
        End If
    Next
    wb.Close 0
End Sub

' 分出标准件非标件 ：沿用了旧名字，不明白意义，不改名
Private Sub StdOrNoStd(wb As Workbook)
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
    With wb.Sheets("设计非标件清单")
        endb = .Cells(65535, 2).End(xlUp).Row
        For i = 2 To endb
            If Mid(.Range("C" & i), 2, 1) = "-" Then
                .Range("m" & i) = Mid(.Range("C" & i), 3, Len(.Range("C" & i)))
            Else
                .Range("m" & i) = .Range("C" & i)
            End If
        Next i
        .Range("N1") = "类型"
        .Range("N2").FormulaR1C1 = "=VLOOKUP(RC[-12],C[1]:C[2],2,0)"
        .Range("N2").AutoFill Destination:=.Range("N2:N" & endb)
    End With
    
    With wb.Sheets("设计打包清单")
        brr = Array("序号", "模板名称", "数量", "打包表名", "分区编号", "W1", "W2", "L", "非标图纸编号", "图纸类别", "是否带配件", "辅助列", "生产单类型")
        .[A1].Resize(1, UBound(brr) + 1) = brr
        enda = .Cells(65535, 1).End(xlUp).Row
        Quyu = ""
        Call Log("main", "D8", "共发现零件:" & enda - 1 & "种",RGB(0,240,0))
        For i = 2 To enda
            Call Log("main", "D9", "已完成:" & i - 1,RGB(240,240,0))
            
            mbmc = .Range("B" & i)
            '在标准件清单中找设计打包清单中的模板名称,如果找到就标注是标准件,没找到看打包名称和上面的是否一样,一样的话就是编号+1,不一样的话就自己开头
            If wb.Sheets("设计非标件清单").Columns(3).Find(mbmc, LookAt:=xlWhole, SearchDirection:=xlPrevious) Is Nothing Then
                If wb.Sheets("设计标准件清单").Columns(3).Find(mbmc, LookAt:=xlWhole, SearchDirection:=xlPrevious) Is Nothing Then
                    .Range("E" & i) = "生产清单中没有"
                Else
                    .Range("E" & i) = "标准件"
                End If
            Else
                hangshu = wb.Sheets("设计非标件清单").Columns(3).Find(mbmc, LookAt:=xlWhole, SearchDirection:=xlPrevious).Row
                arr = wb.Sheets("设计非标件清单").Range("D" & hangshu & ":F" & hangshu)
                .Range("F" & i).Resize(1, 3) = arr
                .Range("I" & i) = wb.Sheets("设计非标件清单").Range("J" & hangshu)
                .Range("J" & i) = wb.Sheets("设计非标件清单").Range("B" & hangshu)
                .Range("K" & i) = wb.Sheets("设计非标件清单").Range("L" & hangshu)
                .Range("L" & i) = wb.Sheets("设计非标件清单").Range("M" & hangshu)
                .Range("M" & i) = wb.Sheets("设计非标件清单").Range("N" & hangshu)
            End If
            If Len(.Range("E" & i)) = 0 Then
                If .Range("D" & i) = Quyu Then
                    k = k + 1
                Else
                    k = 1
                End If
                .Range("E" & i) = .Range("D" & i) & "-" & k
                Quyu = .Range("D" & i).Text
            End If
        Next i
        Call Log("main", "D7", "已完成",RGB(0,240,0))
        Call Log("main", "D9", "已完成",RGB(0,240,0))
    End With
End Sub

' 清单差异比对：沿用了旧名字，不明白意义，不改名
Private Sub QdDiff(wb As Workbook)
    If isSheetExist(wb, "清单差异比对") Then
        wb.Sheets("清单差异比对").Delete
    End If
    If isSheetExist(wb, "清单汇总处理") Then
        wb.Sheets("清单汇总处理").Delete
    End If
    wb.Sheets.Add().Name = "清单差异比对"
    wb.Sheets.Add().Name = "清单汇总处理"
    wb.Sheets("清单差异比对").Activate
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
    wb.Sheets("清单差异比对").Columns("A:A").HorizontalAlignment = Excel.xlCenter
    wb.Sheets("清单差异比对").Columns("B:B").HorizontalAlignment = Excel.xlLeft
    wb.Sheets("清单差异比对").Columns("C:F").HorizontalAlignment = Excel.xlCenter
    wb.Sheets("清单差异比对").Columns("A:F").Font.Name = "宋体"
    wb.Sheets("清单差异比对").Rows("1:65535").RowHeight = 18

    Dim srr
    srr = Array("序号", "模板编号", "打包清单支数", "生产清单支数", "备注")
    wb.Sheets("清单差异比对").[A1].Resize(1, UBound(srr) + 1) = srr
    srr = Array("模板编号", "打包清单支数", "", "模板编号", "生产清单支数")
    wb.Sheets("清单汇总处理").[A1].Resize(1, UBound(srr) + 1) = srr
    wb.Sheets("清单差异比对").Cells(2, 1).Select
    ActiveWindow.FreezePanes = True
    For krd = 2 To wb.Sheets("设计打包清单").Cells(65535, 1).End(xlUp).Row
        If wb.Sheets("设计打包清单").Range("E" & krd).Value = "生产清单中未找到" Then
            cyhangshu = krd
            wb.Sheets("清单差异比对").Range("A" & krf) = krf - 1
            wb.Sheets("清单差异比对").Range("B" & krf) = wb.Sheets("设计打包清单").Range("B" & cyhangshu)
            wb.Sheets("清单差异比对").Range("C" & krf) = wb.Sheets("设计打包清单").Range("C" & cyhangshu)
            wb.Sheets("清单差异比对").Range("D" & krf) = 0
            wb.Sheets("清单差异比对").Range("E" & krf) = "打包清单中有 生产清单中没有的模板编号"
            wb.Sheets("清单差异比对").Range("A" & krf & ":" & "E" & krf).Interior.ColorIndex = 38
            krf = krf + 1
        Else
            dbhzhangshu = krd
            wb.Sheets("清单汇总处理").Range("A" & krj) = wb.Sheets("设计打包清单").Range("B" & dbhzhangshu)
            wb.Sheets("清单汇总处理").Range("B" & krj) = wb.Sheets("设计打包清单").Range("C" & dbhzhangshu)
            krj = krj + 1
        End If
    Next krd
    For krd = 2 To wb.Sheets("设计标准件清单").Cells(65535, 1).End(xlUp).Row
        schzhangshu = krd
        wb.Sheets("清单汇总处理").Range("D" & krl) = wb.Sheets("设计标准件清单").Range("C" & schzhangshu)
        wb.Sheets("清单汇总处理").Range("E" & krl) = wb.Sheets("设计标准件清单").Range("H" & schzhangshu)
        krl = krl + 1
    Next krd
    For krd = 2 To wb.Sheets("设计非标件清单").Cells(65535, 1).End(xlUp).Row
        schzhangshu = krd
        wb.Sheets("清单汇总处理").Range("D" & krl) = wb.Sheets("设计非标件清单").Range("C" & schzhangshu)
        wb.Sheets("清单汇总处理").Range("E" & krl) = wb.Sheets("设计非标件清单").Range("H" & schzhangshu)
        krl = krl + 1
    Next krd

    Dim cache
    Set cache = wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        wb.Sheets("清单汇总处理").Range("A1:B" & (krj - 1)), Version:=xlPivotTableVersion10)
    cache.CreatePivotTable _
        TableDestination:=wb.Sheets("清单汇总处理").Range("G1"), TableName:="打包清单汇总透视表", DefaultVersion:= _
        xlPivotTableVersion10
        wb.Sheets("清单汇总处理").PivotTables("打包清单汇总透视表").AddFields RowFields:=Array("模板编号")
    With wb.Sheets("清单汇总处理").PivotTables("打包清单汇总透视表")
        .AddDataField .PivotFields("打包清单支数"), " 数量", xlSum
    End With
    wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "清单汇总处理!R1C4:R" & (krl - 1) & "C5", Version:=xlPivotTableVersion10).CreatePivotTable _
        TableDestination:="清单汇总处理!R1C10", TableName:="生产清单汇总透视表", DefaultVersion:= _
        xlPivotTableVersion10
    wb.Sheets("清单汇总处理").PivotTables("生产清单汇总透视表").AddFields RowFields:=Array("模板编号")
    With wb.Sheets("清单汇总处理").PivotTables("生产清单汇总透视表")
        .AddDataField .PivotFields("生产清单支数"), " 数量", xlSum
    End With
    For krd = 3 To wb.Sheets("清单汇总处理").Cells(65535, 10).End(xlUp).Row - 1
        mbbh = wb.Sheets("清单汇总处理").Range("J" & krd)
        If wb.Sheets("清单汇总处理").Columns(7).Find(mbbh, LookAt:=xlWhole, SearchDirection:=xlPrevious) Is Nothing Then
            wb.Sheets("清单差异比对").Range("A" & krf) = krf - 1
            wb.Sheets("清单差异比对").Range("B" & krf) = wb.Sheets("清单汇总处理").Range("J" & krd)
            wb.Sheets("清单差异比对").Range("C" & krf) = 0
            wb.Sheets("清单差异比对").Range("D" & krf) = wb.Sheets("清单汇总处理").Range("K" & krd)
            wb.Sheets("清单差异比对").Range("E" & krf) = "生产清单中有 打包清单中没有的模板编号"
            wb.Sheets("清单差异比对").Range("A" & krf & ":" & "E" & krf).Interior.ColorIndex = 36
            krf = krf + 1
        Else
            hdyhangshu = wb.Sheets("清单汇总处理").Columns(7).Find(mbbh, LookAt:=xlWhole, SearchDirection:=xlPrevious).Row
            If wb.Sheets("清单汇总处理").Range("H" & hdyhangshu) <> wb.Sheets("清单汇总处理").Range("K" & krd) Then
                wb.Sheets("清单差异比对").Range("A" & krf) = krf - 1
                wb.Sheets("清单差异比对").Range("B" & krf) = wb.Sheets("清单汇总处理").Range("J" & krd)
                wb.Sheets("清单差异比对").Range("C" & krf) = wb.Sheets("清单汇总处理").Range("H" & hdyhangshu)
                wb.Sheets("清单差异比对").Range("D" & krf) = wb.Sheets("清单汇总处理").Range("K" & krd)
                wb.Sheets("清单差异比对").Range("E" & krf) = "打包清单与生产清单支数不符"
                wb.Sheets("清单差异比对").Range("A" & krf & ":" & "E" & krf).Interior.ColorIndex = 37
                krf = krf + 1
            End If
        End If
    Next krd
End Sub

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

' 打包清单分类 ：沿用了旧名字，不明白意义，不改名
Private Sub Dbqdfl(hbqdFilename As String, dbfqhzFilename As String)
    Dim wb As Workbook
    Set wb = Workbooks.Open(hbqdFilename)
    wb.Windows(1).Visible = True
    ThisWorkbook.Activate
    Dim a As Long
    Dim title_arr
    Dim i As Integer
    Dim endb As Integer
    Dim k As Integer

    Dim c1 As Long
    Dim c2 As Long
    Call Log("main", "D10", "正在编制 《打包分区编号汇总》",RGB(240,240,0))
    endb = wb.Sheets("设计打包清单").Cells(65535, 1).End(xlUp).Row
    c2 = 2
    For c1 = 2 To endb
        If wb.Sheets("设计打包清单").Cells(c1, 5) <> "标准件" And wb.Sheets("设计打包清单").Cells(c1, 5) <> "生产清单中没有" Then
            wb.Sheets("打包分区编号汇总").Range("B" & c2) = wb.Sheets("设计打包清单").Range("B" & c1).Value
            wb.Sheets("打包分区编号汇总").Range("C" & c2) = wb.Sheets("设计打包清单").Range("C" & c1).Value
            wb.Sheets("打包分区编号汇总").Range("D" & c2) = wb.Sheets("设计打包清单").Range("D" & c1).Value
            wb.Sheets("打包分区编号汇总").Range("E" & c2) = wb.Sheets("设计打包清单").Range("E" & c1).Value
            wb.Sheets("打包分区编号汇总").Range("F" & c2) = wb.Sheets("设计打包清单").Range("K" & c1).Value
            c2 = c2 + 1
        End If
    Next c1

    title_arr = Array("序号", "模板编号", "数量", "打包表名", "分区编号", "是否带配件", "备注")
    wb.Sheets("打包分区编号汇总").[A1].Resize(1, UBound(title_arr) + 1) = title_arr

    With wb.Sheets("打包分区编号汇总")
            endb = .Cells(65535, 2).End(xlUp).Row
            For i = 2 To endb
                .Range("A" & i) = i - 1
            Next i
            .Range("A1:G" & endb).Interior.Pattern = xlNone
        .Range("A1:G" & endb).Borders.Weight = 2
        .Columns("A:G").HorizontalAlignment = xlCenter
        .Columns("B:B").EntireColumn.AutoFit
    End With

    If fileIsExist(dbfqhzFilename) Then
        Kill dbfqhzFilename
    End If
    Call createExcel(dbfqhzFilename)
    
    Dim dbfqhzWb As Workbook
    Set dbfqhzWb = Workbooks.Open(dbfqhzFilename)
    dbfqhzWb.Windows(1).Visible = True
    ThisWorkbook.Activate
    dbfqhzWb.Sheets.Add().Name = "打包分区编号汇总"
    Call copySheet(wb.Sheets("打包分区编号汇总"), dbfqhzWb.Sheets("打包分区编号汇总"))
    ' TODO :可能需要把自带的sheet1删了
    ' 打包分区编号汇总在wb里已经可以删了
    dbfqhzWb.Windows(1).Visible = True
    dbfqhzWb.Close (True)
    Set dbfqhzWb = Nothing
    Call Log("main", "D10", "《打包分区编号汇总》编制完成",RGB(0,240,0))
    '先对W1,W2做一下调整
    Dim W1_num As Integer
    Dim W2_num As Integer
    Call Log("main", "D11", "正在 处理《非标不带配件》",RGB(240,240,0))
    endb = wb.Sheets("设计打包清单").Cells(65535, 1).End(xlUp).Row
    c2 = 2
    For c1 = 2 To endb
        If wb.Sheets("设计打包清单").Cells(c1, 5) <> "标准件" And wb.Sheets("设计打包清单").Cells(c1, 5) <> "生产清单中没有" And wb.Sheets("设计打包清单").Cells(c1, 11) = "" Then
            wb.Sheets("非标不带配件").Range("A" & c2 & ":M" & c2) = wb.Sheets("设计打包清单").Range("A" & c1 & ":M" & c1).Value
            c2 = c2 + 1
        End If
    Next c1

    title_arr = Array("序号", "模板名称", "模板编号", "数量", "W1", "W2", "L", "图纸编号", "分区编号", "辅助列", "生产单类型")
    With wb.Sheets("非标不带配件")
        .Columns("J:J").Cut
        .Columns("B:B").Insert Shift:=xlToRight
        .Columns("A:A").ClearContents
        .Columns("F:F").Cut
        .Columns("K:K").Insert Shift:=xlToRight
        .Columns("E:E").Delete Shift:=xlToLeft
        .Columns("j:j").Delete Shift:=xlToLeft
        .[A1].Resize(1, UBound(title_arr) + 1) = title_arr
        endb = .Cells(65535, 2).End(xlUp).Row
        For i = 2 To endb
            If InStr(.Range("B" & i), "C槽") + InStr(.Range("B" & i), "转角") > 0 Then
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
            ElseIf InStr(.Range("B" & i), "角铝") > 0 And Len(.Range("F" & i)) > 0 Then
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
    With wb.Sheets("非标不带配件").Sort.SortFields
        .Clear
        .Add Key:=wb.Sheets("非标不带配件").Range("K2"), Order:=1  '生产单类型
        .Add Key:=wb.Sheets("非标不带配件").Range("B2"), Order:=1  '模板名称
        .Add Key:=wb.Sheets("非标不带配件").Range("E2"), Order:=1  'W1
        .Add Key:=wb.Sheets("非标不带配件").Range("F2"), Order:=1  'W2
        .Add Key:=wb.Sheets("非标不带配件").Range("H2"), Order:=1  '图纸编号
        .Add Key:=wb.Sheets("非标不带配件").Range("J2"), Order:=1  '辅助列
    End With

    With wb.Sheets("非标不带配件").Sort
        .SetRange wb.Sheets("非标不带配件").Range("b2:L" & endb)
        .Header = 2 '没有标题
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
         End With
    k = 0
    With wb.Sheets("非标不带配件")
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
    wb.Sheets("非标不带配件").Columns("J:J").Delete Shift:=xlToLeft
    Call Log("main", "D11", "《非标带配件》处理完成",RGB(0,240,0))

    Call Log("main", "D12", "正在 处理《非标带配件》",RGB(240,240,0))
    endb = wb.Sheets("设计打包清单").Cells(65535, 1).End(xlUp).Row
    c2 = 2
    For c1 = 2 To endb
        If wb.Sheets("设计打包清单").Cells(c1, 5) <> "标准件" And wb.Sheets("设计打包清单").Cells(c1, 5) <> "生产清单中没有" And wb.Sheets("设计打包清单").Cells(c1, 11) = "带配件" Then
            wb.Sheets("非标带配件").Range("A" & c2 & ":M" & c2) = wb.Sheets("设计打包清单").Range("A" & c1 & ":M" & c1).Value
            c2 = c2 + 1
        End If
    Next c1

    With wb.Sheets("非标带配件")
        .Columns("K:K").ClearContents
        .Columns("J:J").Cut
        .Columns("B:B").Insert Shift:=xlToRight
        .Columns("A:A").ClearContents
        .Columns("F:F").Cut
        .Columns("K:K").Insert Shift:=xlToRight
        .Columns("E:E").Delete Shift:=xlToLeft
        .[A1].Resize(1, UBound(title_arr) + 1) = title_arr
        endb = .Cells(65535, 2).End(xlUp).Row
        For i = 2 To endb
            If InStr(.Range("B" & i), "C槽") + InStr(.Range("B" & i), "转角") > 0 Then
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
            ElseIf InStr(.Range("B" & i), "角铝") > 0 And Len(.Range("F" & i)) > 0 Then
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
    With wb.Sheets("非标带配件").Sort.SortFields
        .Clear
        .Add Key:=wb.Sheets("非标带配件").Range("B2"), Order:=1  '模板名称
        .Add Key:=wb.Sheets("非标带配件").Range("E2"), Order:=1  'W1
        .Add Key:=wb.Sheets("非标带配件").Range("F2"), Order:=1  'W2
        .Add Key:=wb.Sheets("非标带配件").Range("H2"), Order:=1  '辅助列
        .Add Key:=wb.Sheets("非标带配件").Range("K2"), Order:=1  '辅助列
    End With

    With wb.Sheets("非标带配件").Sort
        .SetRange wb.Sheets("非标带配件").Range("b2:K" & endb)
        .Header = 2 '没有标题
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    k = 0
    With wb.Sheets("非标带配件")
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
    wb.Sheets("非标带配件").Columns("K:L").Delete Shift:=xlToLeft
    ' cnn.Close: Set cnn = Nothing
    '将各种类型的生产单类型及模板名称列出表格
    ' wb.Sheets("非标不带配件").Activate
    wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "非标不带配件!R1C2:R1048576C10", Version:=xlPivotTableVersion14).CreatePivotTable TableDestination:= _
        "非标不带配件!R1C15", TableName:="数据透视表1", DefaultVersion:=xlPivotTableVersion14
    'wb.Sheets("非标不带配件").Cells(1, 15).Select
    wb.ShowPivotTableFieldList = True
    With wb.Sheets("非标不带配件").PivotTables("数据透视表1").PivotFields("生产单类型")
        .Orientation = xlRowField
        .Position = 1
    End With
    With wb.Sheets("非标不带配件").PivotTables("数据透视表1").PivotFields("模板名称")
        .Orientation = xlRowField
        .Position = 2
    End With
    wb.Sheets("非标不带配件").PivotTables("数据透视表1").AddDataField wb.Sheets("非标不带配件").PivotTables("数据透视表1" _
        ).PivotFields("数量"), "总计数量", xlSum
    wb.ShowPivotTableFieldList = False
    ' Range("O4").Select
    wb.Sheets("非标不带配件").PivotTables("数据透视表1").RowAxisLayout xlTabularRow
    wb.Sheets("非标不带配件").PivotTables("数据透视表1").PivotFields("模板名称").Subtotals = Array(False, _
        False, False, False, False, False, False, False, False, False, False, False)
    wb.Sheets("非标不带配件").PivotTables("数据透视表1").PivotFields("生产单类型").Subtotals = Array(False, _
        False, False, False, False, False, False, False, False, False, False, False)
    wb.Sheets("非标不带配件").PivotTables("数据透视表1").RepeatAllLabels xlRepeatLabels
    Dim end_O As Integer
    end_O = wb.Sheets("非标不带配件").Cells(65535, 15).End(xlUp).Row

    wb.Sheets("非标不带配件").Range("R" & end_O & ":T" & end_O) = wb.Sheets("非标不带配件").Range("O" & end_O & ":Q" & end_O).Value
    ' 这里由于不能覆盖写入，原有的O列改为R列
    
    wb.Sheets("非标不带配件").Range("R" & end_O - 1 & ": T" & end_O - 1).ClearContents
    ' wb.Sheets("非标不带配件").Range("O"&end_O&":Q"&end_O).Delete Shift:=xlLeft
    wb.Sheets("非标不带配件").Range("O:Q").Delete Shift:=xlLeft

    ' wb.Sheets("非标带配件").Activate
    wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "非标带配件!R1C2:R1048576C10", Version:=xlPivotTableVersion14).CreatePivotTable TableDestination:= _
        "非标带配件!R1C15", TableName:="数据透视表2", DefaultVersion:=xlPivotTableVersion14
    ' wb.Sheets("非标带配件").Select
    ' wb.Sheets("非标带配件").Cells(1, 15).Select
    wb.ShowPivotTableFieldList = True
    With wb.Sheets("非标带配件").PivotTables("数据透视表2").PivotFields("生产单类型")
        .Orientation = xlRowField
        .Position = 1
    End With
    With wb.Sheets("非标带配件").PivotTables("数据透视表2").PivotFields("模板名称")
        .Orientation = xlRowField
        .Position = 2
    End With
    wb.Sheets("非标带配件").PivotTables("数据透视表2").AddDataField wb.Sheets("非标带配件").PivotTables("数据透视表2" _
        ).PivotFields("数量"), "总计数量", xlSum
    wb.ShowPivotTableFieldList = False
    ' Range("O4").Select
    wb.Sheets("非标带配件").PivotTables("数据透视表2").RowAxisLayout xlTabularRow
    wb.Sheets("非标带配件").PivotTables("数据透视表2").PivotFields("模板名称").Subtotals = Array(False, _
        False, False, False, False, False, False, False, False, False, False, False)
    wb.Sheets("非标带配件").PivotTables("数据透视表2").PivotFields("生产单类型").Subtotals = Array(False, _
        False, False, False, False, False, False, False, False, False, False, False)
    wb.Sheets("非标带配件").PivotTables("数据透视表2").RepeatAllLabels xlRepeatLabels
    
    end_O = wb.Sheets("非标带配件").Cells(65535, 15).End(xlUp).Row
    wb.Sheets("非标带配件").Range("R" & end_O & ":T" & end_O) = wb.Sheets("非标带配件").Range("O" & end_O & ":Q" & end_O).Value
    wb.Sheets("非标带配件").Range("R" & end_O - 1 & ": T" & end_O - 1).ClearContents
    ' wb.Sheets("非标带配件").Range("O" & end_O & ":Q" & end_O).Delete Shift:=xlLeft
    wb.Sheets("非标带配件").Range("O:Q").Delete Shift:=xlLeft
    Call Log("main", "D12", "《非标带配件》处理完成",RGB(0,240,0))
    Call Log("main", "D10", "已完成",RGB(0,240,0))
    wb.Windows(1).Visible = True
    wb.Close (True)
End Sub



' 拆分到工作簿 :
Private Sub Cfdgzb(hbqdFilename As String, fpqdDirname As String)
    Dim wb As Workbook
    Set wb = Workbooks.Open(hbqdFilename)
    wb.Windows(1).Visible = True
    ThisWorkbook.Activate

    Dim ary(), arr, brr, sh As Worksheet, d As Object, k, t, a, i&, j&, m&, l&
    Dim arr1, k1, T1
    Dim heji As Integer
    Dim path As String
    Dim ws As Worksheet
    path = fpqdDirname
    If Dir(path, vbDirectory) = "" Then
        MkDir path
    End If
    Call Log("main", "D13", "正在将配件 拆分到《分配清单》目录中",RGB(240,240,0))
    Application.DisplayAlerts = False
    For Each ws In wb.Worksheets
        If ws.Name = "非标带配件" Or ws.Name = "非标不带配件" Then
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
            brr = ws.[A1].Resize(65536, UBound(arr, 2))
            For i = 0 To d.count - 1
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
                    .SaveAs filename:=path & Replace(k(i), Chr(9), "") & "-" & heji & ".xlsx"
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
    Call Log("main", "D13", "已完成",RGB(0,240,0))
    Application.DisplayAlerts = True
    wb.Windows(1).Visible = True
    wb.Close (True)
End Sub



