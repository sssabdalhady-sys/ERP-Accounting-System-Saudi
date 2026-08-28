Option Explicit

' =====================================================
' وحدة القوائم المالية (Financial Statements Module)
' Income Statement & Balance Sheet Generation
' =====================================================

' =========================================
' 1. قائمة الدخل (Income Statement)
' =========================================

Public Sub GenerateIncomeStatement()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim sheetExists As Boolean
    Dim row As Integer
    Dim totalRevenue As Double
    Dim totalExpenses As Double
    Dim netIncome As Double
    
    sheetExists = False
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Sheets
        If sh.Name = "Income_Statement" Then
            sheetExists = True
            Set ws = sh
            Exit For
        End If
    Next sh
    
    If Not sheetExists Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "Income_Statement"
    Else
        ws.Cells.Clear
    End If
    
    ' العنوان
    With ws.Range("A1:C1")
        .Merge
        .Value = "قائمة الدخل (Income Statement)"
        .Font.Size = 14
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(0, 102, 204)
        .Font.Color = RGB(255, 255, 255)
    End With
    
    ws.Range("A2").Value = "للفترة من: " & Format(Date - 30, "dd/mm/yyyy") & " إلى: " & Format(Date, "dd/mm/yyyy")
    ws.Range("A2").Font.Italic = True
    
    row = 4
    
    ' الإيرادات
    With ws
        .Cells(row, 1).Value = "الإيرادات (REVENUES)"
        .Cells(row, 1).Font.Bold = True
        .Cells(row, 1).Interior.Color = RGB(200, 220, 255)
    End With
    row = row + 1
    
    totalRevenue = GetAccountGroupTotal("4001", "4031")
    
    ws.Cells(row, 1).Value = "  مبيعات التكييفات"
    ws.Cells(row, 2).Value = GetAccountBalance("4001")
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    row = row + 1
    
    ws.Cells(row, 1).Value = "  مبيعات قطع الغيار"
    ws.Cells(row, 2).Value = GetAccountBalance("4002")
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    row = row + 1
    
    ws.Cells(row, 1).Value = "  مبيعات التركيب والصيانة"
    ws.Cells(row, 2).Value = GetAccountBalance("4003")
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    row = row + 1
    
    ws.Cells(row, 1).Value = "  إيرادات المشاريع"
    ws.Cells(row, 2).Value = GetAccountBalance("4010")
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    row = row + 1
    
    ws.Cells(row, 1).Value = "  إيرادات أخرى"
    ws.Cells(row, 2).Value = GetAccountBalance("4020")
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    row = row + 1
    
    ws.Cells(row, 1).Value = "  (خصم وعودات)"
    ws.Cells(row, 2).Value = -GetAccountBalance("4030")
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    row = row + 1
    
    With ws
        .Cells(row, 1).Value = "إجمالي الإيرادات"
        .Cells(row, 1).Font.Bold = True
        .Cells(row, 2).Value = totalRevenue
        .Cells(row, 2).NumberFormat = "#,##0.00"
        .Cells(row, 2).Font.Bold = True
        .Cells(row, 2).Interior.Color = RGB(220, 240, 255)
    End With
    row = row + 2
    
    ' تكل��ة المبيعات
    With ws
        .Cells(row, 1).Value = "تكلفة المبيعات (COST OF GOODS SOLD)"
        .Cells(row, 1).Font.Bold = True
        .Cells(row, 1).Interior.Color = RGB(200, 220, 255)
    End With
    row = row + 1
    
    ws.Cells(row, 1).Value = "  تكلفة التكييفات"
    ws.Cells(row, 2).Value = GetAccountBalance("5001")
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    row = row + 1
    
    ws.Cells(row, 1).Value = "  تكلفة قطع الغيار"
    ws.Cells(row, 2).Value = GetAccountBalance("5002")
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    row = row + 1
    
    ws.Cells(row, 1).Value = "  تكلفة الخدمات"
    ws.Cells(row, 2).Value = GetAccountBalance("5003")
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    row = row + 1
    
    Dim totalCOGS As Double
    totalCOGS = GetAccountBalance("5001") + GetAccountBalance("5002") + GetAccountBalance("5003")
    
    With ws
        .Cells(row, 1).Value = "إجمالي تكلفة المبيعات"
        .Cells(row, 1).Font.Bold = True
        .Cells(row, 2).Value = totalCOGS
        .Cells(row, 2).NumberFormat = "#,##0.00"
        .Cells(row, 2).Font.Bold = True
        .Cells(row, 2).Interior.Color = RGB(220, 240, 255)
    End With
    row = row + 1
    
    Dim grossProfit As Double
    grossProfit = totalRevenue - totalCOGS
    
    With ws
        .Cells(row, 1).Value = "الربح الإجمالي"
        .Cells(row, 1).Font.Bold = True
        .Cells(row, 2).Value = grossProfit
        .Cells(row, 2).NumberFormat = "#,##0.00"
        .Cells(row, 2).Font.Bold = True
        .Cells(row, 2).Interior.Color = RGB(240, 255, 220)
    End With
    row = row + 2
    
    ' المصروفات الإدارية
    With ws
        .Cells(row, 1).Value = "المصروفات الإدارية والعامة (OPERATING EXPENSES)"
        .Cells(row, 1).Font.Bold = True
        .Cells(row, 1).Interior.Color = RGB(200, 220, 255)
    End With
    row = row + 1
    
    ws.Cells(row, 1).Value = "  رواتب الموظفين"
    ws.Cells(row, 2).Value = GetAccountBalance("5010")
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    row = row + 1
    
    ws.Cells(row, 1).Value = "  بدلات الموظفين"
    ws.Cells(row, 2).Value = GetAccountBalance("5011")
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    row = row + 1
    
    ws.Cells(row, 1).Value = "  إيجار المكتب"
    ws.Cells(row, 2).Value = GetAccountBalance("5020")
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    row = row + 1
    
    ws.Cells(row, 1).Value = "  الكهرباء والمياه"
    ws.Cells(row, 2).Value = GetAccountBalance("5021")
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    row = row + 1
    
    ws.Cells(row, 1).Value = "  الهاتف والانترنت"
    ws.Cells(row, 2).Value = GetAccountBalance("5022")
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    row = row + 1
    
    ws.Cells(row, 1).Value = "  صيانة وتصليحات"
    ws.Cells(row, 2).Value = GetAccountBalance("5030")
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    row = row + 1
    
    ws.Cells(row, 1).Value = "  مصروفات النقل"
    ws.Cells(row, 2).Value = GetAccountBalance("5033")
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    row = row + 1
    
    ws.Cells(row, 1).Value = "  إعلان وتسويق"
    ws.Cells(row, 2).Value = GetAccountBalance("5040")
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    row = row + 1
    
    ws.Cells(row, 1).Value = "  استهلاك الأصول"
    ws.Cells(row, 2).Value = GetAccountBalance("5053")
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    row = row + 1
    
    ws.Cells(row, 1).Value = "  مصروفات أخرى"
    ws.Cells(row, 2).Value = GetAccountBalance("5054")
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    row = row + 1
    
    totalExpenses = GetAccountGroupTotal("5010", "5054")
    
    With ws
        .Cells(row, 1).Value = "إجمالي المصروفات"
        .Cells(row, 1).Font.Bold = True
        .Cells(row, 2).Value = totalExpenses
        .Cells(row, 2).NumberFormat = "#,##0.00"
        .Cells(row, 2).Font.Bold = True
        .Cells(row, 2).Interior.Color = RGB(220, 240, 255)
    End With
    row = row + 2
    
    ' صافي الدخل
    netIncome = grossProfit - totalExpenses
    
    With ws
        .Cells(row, 1).Value = "صافي الدخل (NET INCOME)"
        .Cells(row, 1).Font.Bold = True
        .Cells(row, 1).Font.Size = 12
        .Cells(row, 2).Value = netIncome
        .Cells(row, 2).NumberFormat = "#,##0.00"
        .Cells(row, 2).Font.Bold = True
        .Cells(row, 2).Font.Size = 12
        If netIncome > 0 Then
            .Cells(row, 2).Interior.Color = RGB(144, 238, 144)
        Else
            .Cells(row, 2).Interior.Color = RGB(255, 99, 71)
        End If
    End With
    
    ws.Columns("A:C").AutoFit
    MsgBox "✅ تم إنشاء قائمة الدخل بنجاح!", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "❌ خطأ في إنشاء قائمة الدخل: " & Err.Description, vbCritical
End Sub

' =========================================
' 2. قائمة المركز المالي (Balance Sheet)
' =========================================

Public Sub GenerateBalanceSheet()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim sheetExists As Boolean
    Dim row As Integer
    Dim totalAssets As Double
    Dim totalLiabilities As Double
    Dim totalEquity As Double
    
    sheetExists = False
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Sheets
        If sh.Name = "Balance_Sheet" Then
            sheetExists = True
            Set ws = sh
            Exit For
        End If
    Next sh
    
    If Not sheetExists Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "Balance_Sheet"
    Else
        ws.Cells.Clear
    End If
    
    ' العنوان
    With ws.Range("A1:C1")
        .Merge
        .Value = "قائمة المركز المالي - الميزانية العمومية (Balance Sheet)"
        .Font.Size = 14
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(0, 102, 204)
        .Font.Color = RGB(255, 255, 255)
    End With
    
    ws.Range("A2").Value = "كما في تاريخ: " & Format(Date, "dd/mm/yyyy")
    ws.Range("A2").Font.Italic = True
    
    row = 4
    
    ' الأصول
    With ws
        .Cells(row, 1).Value = "الأصول (ASSETS)"
        .Cells(row, 1).Font.Bold = True
        .Cells(row, 1).Font.Size = 12
        .Cells(row, 1).Interior.Color = RGB(200, 220, 255)
    End With
    row = row + 1
    
    ' الأصول المتداولة
    With ws
        .Cells(row, 1).Value = "الأصول المتداولة (Current Assets)"
        .Cells(row, 1).Font.Bold = True
        .Cells(row, 1).Interior.Color = RGB(220, 240, 255)
    End With
    row = row + 1
    
    Dim currentAssets As Double
    
    ws.Cells(row, 1).Value = "  النقد في الصندوق"
    ws.Cells(row, 2).Value = GetAccountBalance("1001")
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    currentAssets = currentAssets + ws.Cells(row, 2).Value
    row = row + 1
    
    ws.Cells(row, 1).Value = "  النقد في البنك"
    ws.Cells(row, 2).Value = GetAccountBalance("1002") + GetAccountBalance("1003")
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    currentAssets = currentAssets + ws.Cells(row, 2).Value
    row = row + 1
    
    ws.Cells(row, 1).Value = "  الذمم المدينة"
    ws.Cells(row, 2).Value = GetAccountBalance("1010")
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    currentAssets = currentAssets + ws.Cells(row, 2).Value
    row = row + 1
    
    ws.Cells(row, 1).Value = "  المخزون"
    ws.Cells(row, 2).Value = GetAccountBalance("1020") + GetAccountBalance("1021") + GetAccountBalance("1022")
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    currentAssets = currentAssets + ws.Cells(row, 2).Value
    row = row + 1
    
    With ws
        .Cells(row, 1).Value = "إجمالي الأصول المتداولة"
        .Cells(row, 1).Font.Bold = True
        .Cells(row, 2).Value = currentAssets
        .Cells(row, 2).NumberFormat = "#,##0.00"
        .Cells(row, 2).Interior.Color = RGB(230, 245, 255)
    End With
    row = row + 2
    
    ' الأصول الثابتة
    With ws
        .Cells(row, 1).Value = "الأصول الثابتة (Fixed Assets)"
        .Cells(row, 1).Font.Bold = True
        .Cells(row, 1).Interior.Color = RGB(220, 240, 255)
    End With
    row = row + 1
    
    Dim fixedAssets As Double
    
    ws.Cells(row, 1).Value = "  السيارات"
    ws.Cells(row, 2).Value = GetAccountBalance("1501")
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    fixedAssets = fixedAssets + ws.Cells(row, 2).Value
    row = row + 1
    
    ws.Cells(row, 1).Value = "  الأثاث والتجهيزات"
    ws.Cells(row, 2).Value = GetAccountBalance("1502")
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    fixedAssets = fixedAssets + ws.Cells(row, 2).Value
    row = row + 1
    
    ws.Cells(row, 1).Value = "  أجهزة الكمبيوتر"
    ws.Cells(row, 2).Value = GetAccountBalance("1503")
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    fixedAssets = fixedAssets + ws.Cells(row, 2).Value
    row = row + 1
    
    ws.Cells(row, 1).Value = "  (مجمع الاستهلاك)"
    ws.Cells(row, 2).Value = GetAccountBalance("1599")
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    fixedAssets = fixedAssets + ws.Cells(row, 2).Value
    row = row + 1
    
    With ws
        .Cells(row, 1).Value = "إجمالي الأصول الثابتة"
        .Cells(row, 1).Font.Bold = True
        .Cells(row, 2).Value = fixedAssets
        .Cells(row, 2).NumberFormat = "#,##0.00"
        .Cells(row, 2).Interior.Color = RGB(230, 245, 255)
    End With
    row = row + 1
    
    totalAssets = currentAssets + fixedAssets
    
    With ws
        .Cells(row, 1).Value = "إجمالي الأصول"
        .Cells(row, 1).Font.Bold = True
        .Cells(row, 1).Font.Size = 12
        .Cells(row, 2).Value = totalAssets
        .Cells(row, 2).NumberFormat = "#,##0.00"
        .Cells(row, 2).Font.Bold = True
        .Cells(row, 2).Font.Size = 12
        .Cells(row, 2).Interior.Color = RGB(200, 220, 255)
    End With
    row = row + 2
    
    ' الالتزامات
    With ws
        .Cells(row, 1).Value = "الالتزامات (LIABILITIES)"
        .Cells(row, 1).Font.Bold = True
        .Cells(row, 1).Font.Size = 12
        .Cells(row, 1).Interior.Color = RGB(255, 200, 200)
    End With
    row = row + 1
    
    ' الالتزامات المتداولة
    With ws
        .Cells(row, 1).Value = "الالتزامات المتداولة (Current Liabilities)"
        .Cells(row, 1).Font.Bold = True
        .Cells(row, 1).Interior.Color = RGB(255, 220, 220)
    End With
    row = row + 1
    
    Dim currentLiabilities As Double
    
    ws.Cells(row, 1).Value = "  الذمم الدائنة"
    ws.Cells(row, 2).Value = GetAccountBalance("2001")
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    currentLiabilities = currentLiabilities + ws.Cells(row, 2).Value
    row = row + 1
    
    ws.Cells(row, 1).Value = "  الرواتب المستحقة"
    ws.Cells(row, 2).Value = GetAccountBalance("2010")
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    currentLiabilities = currentLiabilities + ws.Cells(row, 2).Value
    row = row + 1
    
    ws.Cells(row, 1).Value = "  ضريبة القيمة المضافة المستحقة"
    ws.Cells(row, 2).Value = GetAccountBalance("2030")
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    currentLiabilities = currentLiabilities + ws.Cells(row, 2).Value
    row = row + 1
    
    With ws
        .Cells(row, 1).Value = "إجمالي الالتزامات المتداولة"
        .Cells(row, 1).Font.Bold = True
        .Cells(row, 2).Value = currentLiabilities
        .Cells(row, 2).NumberFormat = "#,##0.00"
        .Cells(row, 2).Interior.Color = RGB(255, 230, 230)
    End With
    row = row + 1
    
    totalLiabilities = currentLiabilities
    
    With ws
        .Cells(row, 1).Value = "إجمالي الالتزامات"
        .Cells(row, 1).Font.Bold = True
        .Cells(row, 1).Font.Size = 12
        .Cells(row, 2).Value = totalLiabilities
        .Cells(row, 2).NumberFormat = "#,##0.00"
        .Cells(row, 2).Font.Bold = True
        .Cells(row, 2).Font.Size = 12
        .Cells(row, 2).Interior.Color = RGB(255, 200, 200)
    End With
    row = row + 2
    
    ' حقوق الملكية
    With ws
        .Cells(row, 1).Value = "حقوق الملكية (EQUITY)"
        .Cells(row, 1).Font.Bold = True
        .Cells(row, 1).Font.Size = 12
        .Cells(row, 1).Interior.Color = RGB(200, 255, 200)
    End With
    row = row + 1
    
    ws.Cells(row, 1).Value = "  رأس المال"
    ws.Cells(row, 2).Value = GetAccountBalance("3001")
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    row = row + 1
    
    ws.Cells(row, 1).Value = "  الأرباح المحتفظ بها"
    ws.Cells(row, 2).Value = GetAccountBalance("3010")
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    row = row + 1
    
    ws.Cells(row, 1).Value = "  الأرباح/(الخسائر) للسنة"
    ws.Cells(row, 2).Value = GetAccountBalance("3020")
    ws.Cells(row, 2).NumberFormat = "#,##0.00"
    row = row + 1
    
    totalEquity = GetAccountBalance("3001") + GetAccountBalance("3010") + GetAccountBalance("3020")
    
    With ws
        .Cells(row, 1).Value = "إجمالي حقوق الملكية"
        .Cells(row, 1).Font.Bold = True
        .Cells(row, 1).Font.Size = 12
        .Cells(row, 2).Value = totalEquity
        .Cells(row, 2).NumberFormat = "#,##0.00"
        .Cells(row, 2).Font.Bold = True
        .Cells(row, 2).Font.Size = 12
        .Cells(row, 2).Interior.Color = RGB(200, 255, 200)
    End With
    row = row + 1
    
    Dim totalLiabilitiesAndEquity As Double
    totalLiabilitiesAndEquity = totalLiabilities + totalEquity
    
    With ws
        .Cells(row, 1).Value = "إجمالي الالتزامات وحقوق الملكية"
        .Cells(row, 1).Font.Bold = True
        .Cells(row, 1).Font.Size = 12
        .Cells(row, 2).Value = totalLiabilitiesAndEquity
        .Cells(row, 2).NumberFormat = "#,##0.00"
        .Cells(row, 2).Font.Bold = True
        .Cells(row, 2).Font.Size = 12
        .Cells(row, 2).Interior.Color = RGB(200, 255, 200)
    End With
    
    ws.Columns("A:C").AutoFit
    
    ' التحقق من المعادلة المحاسبية
    If Abs(totalAssets - totalLiabilitiesAndEquity) < 0.01 Then
        MsgBox "✅ تم إنشاء قائمة المركز المالي بنجاح!" & vbCrLf & "المعادلة محققة: الأصول = الالتزامات + حقوق الملكية", vbInformation
    Else
        MsgBox "⚠️ تحذير: عدم توازن المعادلة المحاسبية!" & vbCrLf & "الفرق: " & Format(totalAssets - totalLiabilitiesAndEquity, "#,##0.00"), vbExclamation
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "❌ خطأ في إنشاء قائمة المركز المالي: " & Err.Description, vbCritical
End Sub

' =========================================
' 3. ميزان المراجعة (Trial Balance)
' =========================================

Public Sub GenerateTrialBalance()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim chartWs As Worksheet
    Dim sheetExists As Boolean
    Dim row As Integer
    Dim i As Integer
    Dim lastRow As Long
    Dim totalDebit As Double
    Dim totalCredit As Double
    
    sheetExists = False
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Sheets
        If sh.Name = "Trial_Balance" Then
            sheetExists = True
            Set ws = sh
            Exit For
        End If
    Next sh
    
    If Not sheetExists Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "Trial_Balance"
    Else
        ws.Cells.Clear
    End If
    
    Set chartWs = ThisWorkbook.Sheets("Chart_of_Accounts")
    
    ' العنوان
    With ws.Range("A1:D1")
        .Merge
        .Value = "ميزان المراجعة (Trial Balance)"
        .Font.Size = 14
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(0, 102, 204)
        .Font.Color = RGB(255, 255, 255)
    End With
    
    ws.Range("A2").Value = "كما في تاريخ: " & Format(Date, "dd/mm/yyyy")
    ws.Range("A2").Font.Italic = True
    
    ' رؤوس الأعمدة
    With ws.Range("A4:D4")
        .Value = Array("Account Code", "Account Name", "Debit (مدين)", "Credit (دائن)")
        .Font.Bold = True
        .Interior.Color = RGB(200, 220, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    row = 5
    lastRow = chartWs.Cells(chartWs.Rows.Count, 1).End(xlUp).Row
    
    ' جلب جميع الحسابات من شجرة الحسابات
    For i = 2 To lastRow
        Dim accountCode As String
        Dim accountName As String
        Dim accountType As String
        Dim balance As Double
        
        accountCode = chartWs.Cells(i, 1).Value
        accountName = chartWs.Cells(i, 2).Value
        accountType = chartWs.Cells(i, 3).Value
        balance = chartWs.Cells(i, 5).Value
        
        If balance <> 0 Then
            ws.Cells(row, 1).Value = accountCode
            ws.Cells(row, 2).Value = accountName
            
            ' توزيع على المدين والدائن حسب نوع الحساب
            If accountType = "Asset" Or accountType = "Expense" Then
                ws.Cells(row, 3).Value = balance
                totalDebit = totalDebit + balance
            Else
                ws.Cells(row, 4).Value = balance
                totalCredit = totalCredit + balance
            End If
            
            ws.Cells(row, 3).NumberFormat = "#,##0.00"
            ws.Cells(row, 4).NumberFormat = "#,##0.00"
            row = row + 1
        End If
    Next i
    
    ' صف الإجمالي
    row = row + 1
    With ws
        .Cells(row, 1).Value = "الإجمالي"
        .Cells(row, 1).Font.Bold = True
        .Cells(row, 2).Font.Bold = True
        .Cells(row, 3).Value = totalDebit
        .Cells(row, 3).Font.Bold = True
        .Cells(row, 3).NumberFormat = "#,##0.00"
        .Cells(row, 3).Interior.Color = RGB(200, 220, 255)
        .Cells(row, 4).Value = totalCredit
        .Cells(row, 4).Font.Bold = True
        .Cells(row, 4).NumberFormat = "#,##0.00"
        .Cells(row, 4).Interior.Color = RGB(200, 220, 255)
    End With
    
    ws.Columns("A:D").AutoFit
    
    ' التحقق من التوازن
    If Abs(totalDebit - totalCredit) < 0.01 Then
        MsgBox "✅ ميزان المراجعة متوازن!" & vbCrLf & "المجموع المدين = " & Format(totalDebit, "#,##0.00"), vbInformation
    Else
        MsgBox "⚠️ تحذير: عدم توازن ميزان المراجعة!" & vbCrLf & "الفرق: " & Format(Abs(totalDebit - totalCredit), "#,##0.00"), vbExclamation
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "❌ خطأ في إنشاء ميزان المراجعة: " & Err.Description, vbCritical
End Sub

' =========================================
' دوال مساعدة
' =========================================

Private Function GetAccountBalance(accountCode As String) As Double
    On Error GoTo ErrorHandler
    
    Dim chartWs As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim balance As Double
    
    Set chartWs = ThisWorkbook.Sheets("Chart_of_Accounts")
    lastRow = chartWs.Cells(chartWs.Rows.Count, 1).End(xlUp).Row
    
    balance = 0
    
    For i = 2 To lastRow
        If chartWs.Cells(i, 1).Value = accountCode Then
            balance = chartWs.Cells(i, 5).Value
            Exit For
        End If
    Next i
    
    GetAccountBalance = balance
    Exit Function
    
ErrorHandler:
    GetAccountBalance = 0
End Function

Private Function GetAccountGroupTotal(startCode As String, endCode As String) As Double
    On Error GoTo ErrorHandler
    
    Dim chartWs As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim total As Double
    Dim currentCode As Integer
    Dim startCodeInt As Integer
    Dim endCodeInt As Integer
    
    Set chartWs = ThisWorkbook.Sheets("Chart_of_Accounts")
    lastRow = chartWs.Cells(chartWs.Rows.Count, 1).End(xlUp).Row
    
    startCodeInt = CLng(startCode)
    endCodeInt = CLng(endCode)
    total = 0
    
    For i = 2 To lastRow
        currentCode = CLng(chartWs.Cells(i, 1).Value)
        If currentCode >= startCodeInt And currentCode <= endCodeInt Then
            total = total + chartWs.Cells(i, 5).Value
        End If
    Next i
    
    GetAccountGroupTotal = total
    Exit Function
    
ErrorHandler:
    GetAccountGroupTotal = 0
End Function
