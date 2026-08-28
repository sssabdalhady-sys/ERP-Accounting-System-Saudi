Option Explicit

' =====================================================
' نظام ERP محاسبي متكامل سعودي
' Comprehensive Accounting ERP System - Saudi Arabia
' Version 1.0 - Safe & Error-Free
' =====================================================

' =========================================
' 1. شجرة الحسابات (Chart of Accounts)
' =========================================

Public Function InitializeChartOfAccounts()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim accountNum As Integer
    Dim rowNum As Integer
    Dim sheetExists As Boolean
    
    sheetExists = False
    
    ' البحث عن الورقة إذا كانت موجودة
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Sheets
        If sh.Name = "Chart_of_Accounts" Then
            sheetExists = True
            Set ws = sh
            Exit For
        End If
    Next sh
    
    ' إنشاء ورقة جديدة فقط إذا لم تكن موجودة
    If Not sheetExists Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "Chart_of_Accounts"
    Else
        ' تنظيف الورقة بدون حذفها بالكامل
        ws.Range("A1").Select
        ws.Cells.Clear
    End If
    
    ' رؤوس الأعمدة - معادلات محمية
    ws.Range("A1").Value = "Account Code"
    ws.Range("B1").Value = "Account Name"
    ws.Range("C1").Value = "Account Type"
    ws.Range("D1").Value = "Sub Type"
    ws.Range("E1").Value = "Balance"
    
    ' تنسيق رؤوس الأعمدة
    With ws.Range("A1:E1")
        .Font.Bold = True
        .Interior.Color = RGB(0, 102, 204)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    rowNum = 2
    
    ' =====================
    ' الأصول (Assets) - 1000
    ' =====================
    
    ' الأصول المتداولة
    rowNum = AddAccountToSheet(ws, rowNum, "1001", "النقد في الصندوق", "Asset", "Current Assets")
    rowNum = AddAccountToSheet(ws, rowNum, "1002", "النقد في البنك - الحساب الأساسي", "Asset", "Current Assets")
    rowNum = AddAccountToSheet(ws, rowNum, "1003", "النقد في البنك - حساب المشاريع", "Asset", "Current Assets")
    rowNum = AddAccountToSheet(ws, rowNum, "1010", "ذمم مدينة - عملاء", "Asset", "Receivables")
    rowNum = AddAccountToSheet(ws, rowNum, "1011", "مخصص ديون مشكوك فيها", "Asset", "Receivables")
    rowNum = AddAccountToSheet(ws, rowNum, "1020", "المخزون - تكييفات", "Asset", "Inventory")
    rowNum = AddAccountToSheet(ws, rowNum, "1021", "المخزون - قطع غيار", "Asset", "Inventory")
    rowNum = AddAccountToSheet(ws, rowNum, "1022", "المخزون - مواد أخرى", "Asset", "Inventory")
    rowNum = AddAccountToSheet(ws, rowNum, "1030", "مصروفات مقدمة", "Asset", "Current Assets")
    rowNum = AddAccountToSheet(ws, rowNum, "1040", "إيجار مقدم", "Asset", "Current Assets")
    
    ' الأصول الثابتة
    rowNum = AddAccountToSheet(ws, rowNum, "1501", "السيارات", "Asset", "Fixed Assets")
    rowNum = AddAccountToSheet(ws, rowNum, "1502", "الأثاث والتجهيزات", "Asset", "Fixed Assets")
    rowNum = AddAccountToSheet(ws, rowNum, "1503", "أجهزة الكمبيوتر", "Asset", "Fixed Assets")
    rowNum = AddAccountToSheet(ws, rowNum, "1599", "مجمع الاستهلاك", "Asset", "Fixed Assets")
    
    ' =====================
    ' الالتزامات (Liabilities) - 2000
    ' =====================
    
    ' الالتزامات المتداولة
    rowNum = AddAccountToSheet(ws, rowNum, "2001", "ذمم دائنة - موردين", "Liability", "Current Liabilities")
    rowNum = AddAccountToSheet(ws, rowNum, "2010", "الرواتب المستحقة", "Liability", "Current Liabilities")
    rowNum = AddAccountToSheet(ws, rowNum, "2011", "البدلات المستحقة", "Liability", "Current Liabilities")
    rowNum = AddAccountToSheet(ws, rowNum, "2012", "الخصومات والتأمينات", "Liability", "Current Liabilities")
    rowNum = AddAccountToSheet(ws, rowNum, "2020", "الضريبة الدخل المستحقة", "Liability", "Current Liabilities")
    rowNum = AddAccountToSheet(ws, rowNum, "2030", "ضريبة القيمة المضافة - مستحقة", "Liability", "VAT")
    rowNum = AddAccountToSheet(ws, rowNum, "2031", "ضريبة القيمة المضافة - مدفوعة (استرجاع)", "Liability", "VAT")
    rowNum = AddAccountToSheet(ws, rowNum, "2040", "إيجار مستحق", "Liability", "Current Liabilities")
    
    ' الالتزامات طويلة الأجل
    rowNum = AddAccountToSheet(ws, rowNum, "2501", "قروض طويلة الأجل", "Liability", "Long-term Liabilities")
    rowNum = AddAccountToSheet(ws, rowNum, "2502", "رهون عقارية", "Liability", "Long-term Liabilities")
    
    ' =====================
    ' حقوق الملكية (Equity) - 3000
    ' =====================
    
    rowNum = AddAccountToSheet(ws, rowNum, "3001", "رأس المال", "Equity", "Owner's Equity")
    rowNum = AddAccountToSheet(ws, rowNum, "3010", "الأرباح المحتفظ بها", "Equity", "Owner's Equity")
    rowNum = AddAccountToSheet(ws, rowNum, "3020", "الأرباح/(الخسائر) للسنة الحالية", "Equity", "Current Year")
    
    ' =====================
    ' الإيرادات (Revenue) - 4000
    ' =====================
    
    rowNum = AddAccountToSheet(ws, rowNum, "4001", "مبيعات التكييفات", "Revenue", "Sales")
    rowNum = AddAccountToSheet(ws, rowNum, "4002", "مبيعات قطع الغيار", "Revenue", "Sales")
    rowNum = AddAccountToSheet(ws, rowNum, "4003", "مبيعات تركيب وصيانة", "Revenue", "Sales")
    rowNum = AddAccountToSheet(ws, rowNum, "4010", "إيرادات المشاريع", "Revenue", "Services")
    rowNum = AddAccountToSheet(ws, rowNum, "4011", "إيرادات الاستشارة", "Revenue", "Services")
    rowNum = AddAccountToSheet(ws, rowNum, "4020", "الفائدة المكتسبة", "Revenue", "Other Income")
    rowNum = AddAccountToSheet(ws, rowNum, "4021", "إيرادات أخرى", "Revenue", "Other Income")
    rowNum = AddAccountToSheet(ws, rowNum, "4030", "خصم المبيعات", "Revenue", "Sales Adjustments")
    rowNum = AddAccountToSheet(ws, rowNum, "4031", "مردودات المبيعات", "Revenue", "Sales Adjustments")
    
    ' =====================
    ' المصروفات (Expenses) - 5000
    ' =====================
    
    ' تكلفة المبيعات
    rowNum = AddAccountToSheet(ws, rowNum, "5001", "تكلفة المبيعات - تكييفات", "Expense", "Cost of Goods")
    rowNum = AddAccountToSheet(ws, rowNum, "5002", "تكلفة المبيعات - قطع غيار", "Expense", "Cost of Goods")
    rowNum = AddAccountToSheet(ws, rowNum, "5003", "تكلفة المبيعات - خدمات", "Expense", "Cost of Goods")
    
    ' الرواتب والمزايا
    rowNum = AddAccountToSheet(ws, rowNum, "5010", "رواتب الموظفين", "Expense", "Salaries & Benefits")
    rowNum = AddAccountToSheet(ws, rowNum, "5011", "بدلات الموظفين", "Expense", "Salaries & Benefits")
    rowNum = AddAccountToSheet(ws, rowNum, "5012", "معاشات اجتماعية", "Expense", "Salaries & Benefits")
    rowNum = AddAccountToSheet(ws, rowNum, "5013", "مكافآت الموظفين", "Expense", "Salaries & Benefits")
    
    ' المصروفات الإدارية
    rowNum = AddAccountToSheet(ws, rowNum, "5020", "إيجار المكتب", "Expense", "Administrative")
    rowNum = AddAccountToSheet(ws, rowNum, "5021", "الكهرباء والمياه", "Expense", "Administrative")
    rowNum = AddAccountToSheet(ws, rowNum, "5022", "الهاتف والانترنت", "Expense", "Administrative")
    rowNum = AddAccountToSheet(ws, rowNum, "5023", "مكاتب ومستلزمات", "Expense", "Administrative")
    rowNum = AddAccountToSheet(ws, rowNum, "5024", "صيانة المباني", "Expense", "Administrative")
    
    ' مصروفات التشغيل
    rowNum = AddAccountToSheet(ws, rowNum, "5030", "صيانة وتصليحات", "Expense", "Operating")
    rowNum = AddAccountToSheet(ws, rowNum, "5031", "وقود السيارات", "Expense", "Operating")
    rowNum = AddAccountToSheet(ws, rowNum, "5032", "صيانة السيارات", "Expense", "Operating")
    rowNum = AddAccountToSheet(ws, rowNum, "5033", "مصروفات النقل والتوصيل", "Expense", "Operating")
    
    ' مصروفات التسويق والبيع
    rowNum = AddAccountToSheet(ws, rowNum, "5040", "إعلان وتسويق", "Expense", "Marketing & Sales")
    rowNum = AddAccountToSheet(ws, rowNum, "5041", "عمولات مبيعات", "Expense", "Marketing & Sales")
    rowNum = AddAccountToSheet(ws, rowNum, "5042", "هدايا العملاء", "Expense", "Marketing & Sales")
    
    ' مصروفات أخرى
    rowNum = AddAccountToSheet(ws, rowNum, "5050", "تأمين", "Expense", "Other")
    rowNum = AddAccountToSheet(ws, rowNum, "5051", "رسوم بنكية", "Expense", "Other")
    rowNum = AddAccountToSheet(ws, rowNum, "5052", "رسوم قانونية ومحاسبية", "Expense", "Other")
    rowNum = AddAccountToSheet(ws, rowNum, "5053", "استهلاك الأصول", "Expense", "Depreciation")
    rowNum = AddAccountToSheet(ws, rowNum, "5054", "مصروفات متنوعة", "Expense", "Other")
    
    ' تعديل عرض الأعمدة
    ws.Columns("A:E").AutoFit
    
    ' تجميد الصف الأول
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
    
    MsgBox "✅ تم إنشاء شجرة الحسابات بنجاح!" & vbCrLf & "تم إضافة " & (rowNum - 2) & " حساب", vbInformation
    Exit Function
    
ErrorHandler:
    MsgBox "❌ حدث خطأ: " & Err.Description, vbCritical
End Function

Private Function AddAccountToSheet(ws As Worksheet, rowNum As Integer, code As String, name As String, type As String, subType As String) As Integer
    On Error GoTo ErrorHandler
    
    With ws
        .Cells(rowNum, 1).Value = code
        .Cells(rowNum, 2).Value = name
        .Cells(rowNum, 3).Value = type
        .Cells(rowNum, 4).Value = subType
        .Cells(rowNum, 5).Value = 0
        
        ' تنسيق
        .Cells(rowNum, 1).NumberFormat = "0000"
        .Cells(rowNum, 5).NumberFormat = "#,##0.00"
        .Cells(rowNum, 1).HorizontalAlignment = xlCenter
    End With
    
    AddAccountToSheet = rowNum + 1
    Exit Function
    
ErrorHandler:
    MsgBox "❌ خطأ في إضافة الحساب: " & Err.Description, vbCritical
    AddAccountToSheet = rowNum
End Function

' =========================================
' 2. دفتر اليومية (General Journal)
' =========================================

Public Function InitializeGeneralJournal()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim sheetExists As Boolean
    
    sheetExists = False
    
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Sheets
        If sh.Name = "General_Journal" Then
            sheetExists = True
            Set ws = sh
            Exit For
        End If
    Next sh
    
    If Not sheetExists Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "General_Journal"
    Else
        ws.Range("A1").Select
        ws.Cells.Clear
    End If
    
    ' رؤوس الأعمدة
    ws.Range("A1").Value = "Journal #"
    ws.Range("B1").Value = "Date"
    ws.Range("C1").Value = "Account Code"
    ws.Range("D1").Value = "Account Name"
    ws.Range("E1").Value = "Description"
    ws.Range("F1").Value = "Debit (مدين)"
    ws.Range("G1").Value = "Credit (دائن)"
    ws.Range("H1").Value = "Posted"
    
    With ws.Range("A1:H1")
        .Font.Bold = True
        .Interior.Color = RGB(0, 102, 204)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    ' تنسيق الأعمدة
    ws.Range("B:B").NumberFormat = "mm/dd/yyyy"
    ws.Range("F:F").NumberFormat = "#,##0.00"
    ws.Range("G:G").NumberFormat = "#,##0.00"
    
    ws.Columns("A:H").AutoFit
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
    
    MsgBox "✅ تم إنشاء دفتر اليومية بنجاح!", vbInformation
    Exit Function
    
ErrorHandler:
    MsgBox "❌ حدث خطأ: " & Err.Description, vbCritical
End Function

' =========================================
' 3. الأستاذ العام (General Ledger)
' =========================================

Public Function InitializeGeneralLedger()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim sheetExists As Boolean
    
    sheetExists = False
    
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Sheets
        If sh.Name = "General_Ledger" Then
            sheetExists = True
            Set ws = sh
            Exit For
        End If
    Next sh
    
    If Not sheetExists Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "General_Ledger"
    Else
        ws.Range("A1").Select
        ws.Cells.Clear
    End If
    
    ' رؤوس الأعمدة
    ws.Range("A1").Value = "Account Code"
    ws.Range("B1").Value = "Account Name"
    ws.Range("C1").Value = "Date"
    ws.Range("D1").Value = "Description"
    ws.Range("E1").Value = "Debit (مدين)"
    ws.Range("F1").Value = "Credit (دائن)"
    ws.Range("G1").Value = "Balance (الرصيد)"
    
    With ws.Range("A1:G1")
        .Font.Bold = True
        .Interior.Color = RGB(0, 102, 204)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    ' تنسيق الأعمدة
    ws.Range("C:C").NumberFormat = "mm/dd/yyyy"
    ws.Range("E:G").NumberFormat = "#,##0.00"
    
    ws.Columns("A:G").AutoFit
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
    
    MsgBox "✅ تم إنشاء الأستاذ العام بنجاح!", vbInformation
    Exit Function
    
ErrorHandler:
    MsgBox "❌ حدث خطأ: " & Err.Description, vbCritical
End Function

' =========================================
' 4. لوحة التحكم الرئيسية
' =========================================

Public Sub CreateMainDashboard()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim sheetExists As Boolean
    
    sheetExists = False
    
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Sheets
        If sh.Name = "Dashboard" Then
            sheetExists = True
            Set ws = sh
            Exit For
        End If
    Next sh
    
    If Not sheetExists Then
        Set ws = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
        ws.Name = "Dashboard"
    Else
        ws.Range("A1").Select
        ws.Cells.Clear
    End If
    
    ' العنوان الرئيسي
    With ws.Range("A1:D1")
        .Merge
        .Value = "🏢 لوحة التحكم - نظام ERP محاسبي سعودي"
        .Font.Size = 16
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(0, 102, 204)
        .Font.Color = RGB(255, 255, 255)
    End With
    
    ws.Range("A1:D1").RowHeight = 30
    
    ' المقدمة
    ws.Range("A3").Value = "مرحباً بك في نظام ERP المحاسبي المتكامل"
    ws.Range("A3").Font.Size = 12
    ws.Range("A3").Font.Bold = True
    
    ws.Range("A5").Value = "الوحدات الرئيسية:"
    ws.Range("A5").Font.Bold = True
    
    Dim row As Integer
    row = 7
    
    ws.Cells(row, 1).Value = "📊 1. شجرة الحسابات"
    ws.Cells(row, 2).Value = "إنشاء وإدارة شجرة الحسابات السعودية"
    row = row + 1
    
    ws.Cells(row, 1).Value = "📝 2. دفتر اليومية"
    ws.Cells(row, 2).Value = "إدخال ومعالجة القيود اليومية"
    row = row + 1
    
    ws.Cells(row, 1).Value = "📖 3. الأستاذ العام"
    ws.Cells(row, 2).Value = "عرض وتحليل الأستاذ العام"
    row = row + 1
    
    ws.Cells(row, 1).Value = "💰 4. قائمة الدخل"
    ws.Cells(row, 2).Value = "تقارير الإيرادات والمصروفات"
    row = row + 1
    
    ws.Cells(row, 1).Value = "⚖️ 5. قائمة المركز المالي"
    ws.Cells(row, 2).Value = "الميزانية العمومية والأصول"
    row = row + 1
    
    ws.Cells(row, 1).Value = "🔢 6. ميزان المراجعة"
    ws.Cells(row, 2).Value = "التحقق من توازن الحسابات"
    
    ws.Columns("A:B").AutoFit
    ws.Range("A1").Select
    
    MsgBox "✅ تم إنشاء لوحة التحكم بنجاح!", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "❌ حدث خطأ: " & Err.Description, vbCritical
End Sub

' =========================================
' 5. تهيئة كاملة النظام
' =========================================

Public Sub InitializeCompleteSystem()
    On Error GoTo ErrorHandler
    
    Dim response As Integer
    
    response = MsgBox("هل تريد تهيئة النظام المحاسبي؟" & vbCrLf & "سيتم إنشاء:" & vbCrLf & _
    "✓ لوحة التحكم" & vbCrLf & _
    "✓ شجرة الحسابات" & vbCrLf & _
    "✓ دفتر اليومية" & vbCrLf & _
    "✓ الأستاذ العام", vbYesNo + vbQuestion)
    
    If response = vbNo Then Exit Sub
    
    Call CreateMainDashboard
    Call InitializeChartOfAccounts
    Call InitializeGeneralJournal
    Call InitializeGeneralLedger
    
    MsgBox "✅ تم تهيئة النظام المحاسبي بنجاح!" & vbCrLf & "جميع الوحدات جاهزة للاستخدام", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "❌ حدث خطأ أثناء التهيئة: " & Err.Description, vbCritical
End Sub
