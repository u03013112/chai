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
    scqdFilename = "C:\Users\u03013112\Documents\J\new-412-1\分配清单\" & txtlpdm & txtgcmc & txtqyjx & "生产单.xlsx"
    Call FB1(fpqdFilename, scqdFilename)
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
    Call copySheet(ThisWorkbook.Sheets("模板生产计划单"), wb.Sheets("临时"))

    ' ZXD拷贝过去
    ' TODO: 这些数据都应该是追加的，现在全都是替换，之后看看是生成多个文件，还是汇总起来
    If isSheetExist(wb, "ZXD") Then
        wb.Sheets("ZXD").Delete
    End If
    wb.Sheets.Add().Name = "ZXD"
    Call copySheet(ThisWorkbook.Sheets("模板生产计划单"), wb.Sheets("ZXD"))

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

    wb.Sheets("ZXD").Range("B2") = gcmc & qysr
    wb.Sheets("ZXD").Range("A1") = "模板转序记录表 (" & ptfs & ")"
    
    
    Dim wbTmp As Workbook
    Set wbTmp = Workbooks.Open(fpqdFilename)
    wbTmp.Sheets(1).Cells.Copy wb.Worksheets("erp").[A1]
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
        MsgBox "总数量： " & Slhj & " 件"
        wb.Close (True)
    End If
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