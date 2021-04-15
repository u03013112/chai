Option Explicit

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

Sub chose1()
    Dim fso As Object, arr(1 To 10 ^ 2, 1 To 1), i
    Dim dg As FileDialog

    Dim strfile As String
    Dim brr
    Dim fileFolderName, hbqdFilename As String
    
    Set dg = Application.FileDialog(msoFileDialogFolderPicker)
    
    If dg.Show = -1 Then
        '递归所选目录，找到所有excel文件
        Dim excelFilenames As Variant
        excelFilenames = getAllFile(dg.SelectedItems(1))
        [f5] = getArrLen(excelFilenames)
        'TODO 检测找到的文件是否合格
        
        '在本地进行临时处理
        'TODO 检测是否有临时文件，如果有是否要按照进度仅需处理，或者清理之后重新做
        
        'TODO 在本地新建一个目录，用于存储对应的临时文件和结果
        fileFolderName = Split(dg.SelectedItems(1), "\")(UBound(Split(dg.SelectedItems(1), "\")))
        Dim outputDir As String
        outputDir = ThisWorkbook.Path & "\" & fileFolderName
        If dirIsExist(outputDir) = True Then
            MsgBox (outputDir & " already exist")
        Else
            VBA.MkDir (outputDir)
        End If
        '提前建立合并清单文件
        hbqdFilename = outputDir & "\" & fileFolderName & "-合并清单.xlsx"
        If fileIsExist(hbqdFilename) Then
            MsgBox (hbqdFilename & " already exist")
        Else
            createExcel (hbqdFilename)
            initHBQD (hbqdFilename)
        End If
        
        '[F2] = dg.SelectedItems(1)
        '[F3] = hbqdFileName
        
        Dim c As Variant
        
        For Each excelFilename In excelFilenames
            If excelFilename = enmpty Then
                Exit For
            End If
            MsgBocx (excelFilename)
            '设计清单复制 (excelFilename.Name)
        Next
        Exit Sub
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
    
    ThisWorkbook.SaveAs filename:=strfile & wjj_name & "\" & wjj_name & "-合并清单.xlsx"
    
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
    
    Sheets("库(待补充)").Delete
    
    Application.DisplayAlerts = True
    
    'Call 分出标准件非标件
    
    'Call 清单差异比对
    
    Application.DisplayAlerts = False
    
    If Sheets("清单差异比对").Cells(Rows.Count, 1).End(xlUp).Row > 1 Then
        
        Sheets("清单汇总处理").Delete
        
        ThisWorkbook.Worksheets("清单差异比对").Columns("A:E").EntireColumn.AutoFit
        
        Worksheets(Array("设计打包清单", "设计标准件清单", "设计非标件清单", "清单差异比对")).Copy
        ActiveWorkbook.SaveAs filename:=strfile & wjj_name & "\" & wjj_name & "-清单差异", FileFormat:=51
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
    
    'Call 打包清单分类
    'Call 拆分到工作簿
    
    Application.ScreenUpdating = True
    
    MsgBox "拆分完毕"
End Sub

' 初始化合并清单，目前多次打开文件，有优化空间
' 只是建立了Sheet，添加了第一行
Sub initHBQD(hbqdFilename As String)
    Dim wb As Workbook
    Dim arr(100, 1)
    Set wb = Workbooks.Open(hbqdFilename)
    
    Application.ScreenUpdating = False
        
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

Private Function getAllFile(MyPath As String) As Variant
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
        arrTmp = getAllFile(SubFolder.Path) '递归
        For Each filename In arrTmp
            If arrTmp(i) = Empty Then
                Exit For
            End If
            arr(i) = filename
            i = i + 1
        Next
    Next
    getAllFile = arr
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

'沿用了旧名字，不明白意义，不改名
Private Sub 设计清单复制(filename As String)
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
        endb = wb.Sheets(k).Cells(Rows.Count, 2).End(xlUp).Row '
        enda = wb.Sheets(k).Cells(Rows.Count, 1).End(xlUp).Row '两侧检测以免数量列的最后一行不是非空单元格
        
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
        wb.Close 0
End Sub
