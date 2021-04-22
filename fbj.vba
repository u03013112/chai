



Private Sub txtbmcl_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then txtxdsj.SetFocus
End Sub

Private Sub txtgyxm_Change()

End Sub

Private Sub txtqyjx_Change()

End Sub

Private Sub txtshxm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then Call CommandButton1_Click
End Sub

Private Sub UserForm_Initialize()
 
lastgyxm = GetSetting("MyApp", "Startup", "gyxm", "")
lastgydh = GetSetting("MyApp", "Startup", "gydh", "")
lastshxm = GetSetting("MyApp", "Startup", "shxm", "")

txtgyxm.Text = lastgyxm: txtgydh.Text = lastgydh: txtshxm.Text = lastshxm
txtxdsj.Text = Date
End Sub


Private Sub CommandButton1_Click()
   FBJ.Hide

    Dim wjm$, wb As Workbook, ws As Worksheet, rngs As Range, arr
    Dim cell As Range
    Dim jctem As String
    Dim gcmc As String '工程名称
    Dim Quyu As String '区域
    Dim Xdrq '下单日期
    Dim jhdh '计划单号
    
    Range("A1:K1").Borders.LineStyle = xlContinuous
    ThisWorkbook.Sheets("Sheet1").Activate
    
    With Sheets("Sheet1")
    '------------------------------------------------------------------------
    
    '用来保存 GetSetting 函数所返回之二维数组数据的变量。
    Dim MySettings As Variant
    ' 在注册区中添加项目。

    .Range("B3") = txtlpdm
    .Range("B2") = txtgcmc
    
     qysr = txtqyjx.Text '区域输入
     txtsdsj = Date
    .Range("G2") = qysr
    .Range("I2") = txtsdsj
    .Range("G3") = txtjhdh
    
    .Range("K3") = txtbmcl
    
        If Len(.Range("B5")) = 0 Then
             .Range("B5") = txtgyxm
             SaveSetting "MyApp", "Startup", "gyxm", txtgyxm
        End If
        gydh = txtgydh
        SaveSetting "MyApp", "Startup", "gydh", txtgydh
        If Len(.Range("F5")) = 0 Then
             .Range("F5") = txtshxm
             SaveSetting "MyApp", "Startup", "shxm", txtshxm
        End If
    '------------------------------------------------------------------------
        gcmc = .Range("B2")
        gyxm = .Range("B5")
        shxm = .Range("F5")
        jhdh = .Range("G3")
        ptfs = .Range("K3") '喷涂方式
    End With
    Sheets("ZXD").Range("B2") = gcmc & qysr
    Sheets("ZXD").Range("A1") = "模板转序记录表 (" & ptfs & ")"
    scch = txtscch.Text
    Dim strfile As String
    Dim nm
    nm = Application.GetOpenFilename("Excel 文件 ,*.xls*;*.xlsx")
    If nm = False Then
        MsgBox "请选择文件"
    Else
        strfile = Left(nm, Len(nm) - Len(Split(nm, "\")(UBound(Split(nm, "\")))))
    End If
    
    Set wb = Workbooks.Open(nm)
    wb.Sheets(1).Cells.Copy ThisWorkbook.Worksheets("erp").[A1]
    wb.Close False


    If txtqyjx.Text <> "BZJ" And txtqyjx.Text <> "bzj" Then
    
    
        ThisWorkbook.Sheets("erp").Activate
        
        If Range("A1") = "序号" Then Rows("1:1").Delete
        endd = [d65536].End(xlUp).Row
        Range("j1:j" & endd) = qysr
        Range("A1:J" & endd).Interior.Pattern = xlNone
        Range("A1:J" & endd).Borders.Weight = 2

        If Left(qysr, 2) = "TP" Or Left(qysr, 3) = " TP" Then
            Range("k1:k" & endd) = "带配件"
            Range("k1:k" & endd).Interior.Pattern = xlNone
            Range("k1:k" & endd).Borders.Weight = 2
        End If

        
        Columns("C:C").EntireColumn.AutoFit '调整模板编号列的列宽
        Slhj = Application.WorksheetFunction.Sum(Range("D1:D" & endd)) '数量合计
        Columns("A:k").FormatConditions.Delete '清空条件格式
        Columns("A:J").HorizontalAlignment = xlCenter '水平方向居中
        Application.ScreenUpdating = True
        MsgBox "总数量： " & Slhj & " 件"
        ThisWorkbook.SaveAs FileName:=strfile & txtlpdm & gcmc & qysr & "生产单.xlsm"

        
    
    End If
    
    Unload Me
End Sub

Private Sub CommandButton2_Click()
    Unload Me
End Sub
