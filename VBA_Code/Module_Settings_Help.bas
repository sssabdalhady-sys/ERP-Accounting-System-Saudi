Option Explicit

' =====================================================
' وحدة الإعدادات والمساعدة (Settings & Help Module)
' System Settings & User Guide
' =====================================================

Public Sub OpenSettings()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim sheetExists As Boolean
    
    sheetExists = False
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Sheets
        If sh.Name = "Settings" Then
            sheetExists = True
            Set ws = sh
            Exit For
        End If
    Next sh
    
    If Not sheetExists Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "Settings"
    Else
        ws.Cells.Clear
    End If
    
    Dim row As Integer
    row = 1
    
    ' العنوان
    With ws.Range("A1:C1")
        .Merge
        .Value = "⚙️ إعدادات النظام"
        .Font.Size = 14
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(0, 102, 204)
        .Font.Color = RGB(255, 255, 255)
    End With
    
    ws.Range("A1:C1").RowHeight = 25
    row = 3
    
    ' بيانات الشركة
    With ws
        .Cells(row, 1).Value = "معلومات الشركة"
        .Cells(row, 1).Font.Bold = True
        .Cells(row, 1).Font.Size = 12
        .Cells(row, 1).Interior.Color = RGB(200, 220, 255)
    End With
    row = row + 1
    
    With ws
        .Cells(row, 1).Value = "اسم الشركة:"
        .Cells(row, 2).Value = "[أدخل اسم الشركة]"
        row = row + 1
        
        .Cells(row, 1).Value = "السجل التجاري:"
        .Cells(row, 2).Value = "[أدخل رقم السجل التجاري]"
        row = row + 1
        
        .Cells(row, 1).Value = "رقم الضريبة:"
        .Cells(row, 2).Value = "[أدخل رقم الضريبة]"
        row = row + 1
        
        .Cells(row, 1).Value = "العنوان:"
        .Cells(row, 2).Value = "[أدخل عنوان الشركة]"
        row = row + 1
        
        .Cells(row, 1).Value = "الهاتف:"
        .Cells(row, 2).Value = "[أد��ل رقم الهاتف]"
        row = row + 1
        
        .Cells(row, 1).Value = "البريد الإلكتروني:"
        .Cells(row, 2).Value = "[أدخل البريد الإلكتروني]"
    End With
    
    row = row + 2
    
    ' إعدادات الضريبة
    With ws
        .Cells(row, 1).Value = "إعدادات ضريبة القيمة المضافة"
        .Cells(row, 1).Font.Bold = True
        .Cells(row, 1).Font.Size = 12
        .Cells(row, 1).Interior.Color = RGB(200, 220, 255)
    End With
    row = row + 1
    
    With ws
        .Cells(row, 1).Value = "نسبة الضريبة (%):"
        .Cells(row, 2).Value = 15
        .Cells(row, 2).NumberFormat = "0.00"
        row = row + 1
        
        .Cells(row, 1).Value = "تاريخ بدء التطبيق:"
        .Cells(row, 2).Value = Date
        .Cells(row, 2).NumberFormat = "mm/dd/yyyy"
    End With
    
    row = row + 2
    
    ' إعدادات العملة
    With ws
        .Cells(row, 1).Value = "إعدادات العملة"
        .Cells(row, 1).Font.Bold = True
        .Cells(row, 1).Font.Size = 12
        .Cells(row, 1).Interior.Color = RGB(200, 220, 255)
    End With
    row = row + 1
    
    With ws
        .Cells(row, 1).Value = "العملة الأساسية:"
        .Cells(row, 2).Value = "ريال سعودي (SAR)"
        row = row + 1
        
        .Cells(row, 1).Value = "رمز العملة:"
        .Cells(row, 2).Value = "ر.س"
    End With
    
    ws.Columns("A:C").AutoFit
    MsgBox "✅ تم فتح صفحة الإعدادات بنجاح!", vbInformation
    
    Exit Sub
    
ErrorHandler:
    MsgBox "❌ خطأ: " & Err.Description, vbCritical
End Sub

Public Sub OpenUserGuide()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim sheetExists As Boolean
    
    sheetExists = False
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Sheets
        If sh.Name = "User_Guide" Then
            sheetExists = True
            Set ws = sh
            Exit For
        End If
    Next sh
    
    If Not sheetExists Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "User_Guide"
    Else
        ws.Cells.Clear
    End If
    
    Dim row As Integer
    row = 1
    
    ' العنوان
    With ws.Range("A1:C1")
        .Merge
        .Value = "📖 دليل الاستخدام"
        .Font.Size = 14
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(0, 102, 204)
        .Font.Color = RGB(255, 255, 255)
    End With
    
    ws.Range("A1:C1").RowHeight = 25
    row = 3
    
    With ws
        .Cells(row, 1).Value = "1. البدء السريع"
        .Cells(row, 1).Font.Bold = True
        .Cells(row, 1).Interior.Color = RGB(200, 220, 255)
        row = row + 1
        
        .Cells(row, 1).Value = "- افتح الملف وفعّل Macros"
        row = row + 1
        
        .Cells(row, 1).Value = "- اذهب إلى لوحة التحكم (Dashboard)"
        row = row + 1
        
        .Cells(row, 1).Value = "- اختر الوحدة المطلوبة (مبيعات، مشتريات، إلخ)"
        row = row + 1
        
        .Cells(row, 1).Value = ""
        row = row + 1
        
        .Cells(row, 1).Value = "2. إدخال فاتورة مبيعات"
        .Cells(row, 1).Font.Bold = True
        .Cells(row, 1).Interior.Color = RGB(200, 220, 255)
        row = row + 1
        
        .Cells(row, 1).Value = "- اذهب إلى وحدة 'Sales_Invoices'"
        row = row + 1
        
        .Cells(row, 1).Value = "- أدخل رقم الفاتورة والتاريخ"
        row = row + 1
        
        .Cells(row, 1).Value = "- أضف بيانات العميل والسعر"
        row = row + 1
        
        .Cells(row, 1).Value = "- سيتم الترحيل تلقائياً إلى اليومية والأستاذ"
        row = row + 1
        
        .Cells(row, 1).Value = ""
        row = row + 1
        
        .Cells(row, 1).Value = "3. إدخال طلبية شراء"
        .Cells(row, 1).Font.Bold = True
        .Cells(row, 1).Interior.Color = RGB(200, 220, 255)
        row = row + 1
        
        .Cells(row, 1).Value = "- اذهب إلى وحدة 'Purchase_Orders'"
        row = row + 1
        
        .Cells(row, 1).Value = "- أدخل رقم الطلبية والتاريخ"
        row = row + 1
        
        .Cells(row, 1).Value = "- أضف بيانات الموردين والسعر"
        row = row + 1
        
        .Cells(row, 1).Value = "- سيتم الترحيل تلقائياً"
        row = row + 1
        
        .Cells(row, 1).Value = ""
        row = row + 1
        
        .Cells(row, 1).Value = "4. معالجة الرواتب"
        .Cells(row, 1).Font.Bold = True
        .Cells(row, 1).Interior.Color = RGB(200, 220, 255)
        row = row + 1
        
        .Cells(row, 1).Value = "- اذهب إلى وحدة 'Payroll'"
        row = row + 1
        
        .Cells(row, 1).Value = "- أضف كل موظف مع الراتب والبدلات والخصومات"
        row = row + 1
        
        .Cells(row, 1).Value = "- سيتم حساب الراتب الصافي تلقائياً"
        row = row + 1
        
        .Cells(row, 1).Value = ""
        row = row + 1
        
        .Cells(row, 1).Value = "5. عرض القوائم المالية"
        .Cells(row, 1).Font.Bold = True
        .Cells(row, 1).Interior.Color = RGB(200, 220, 255)
        row = row + 1
        
        .Cells(row, 1).Value = "- قائمة الدخل: استخدم وحدة 'Income_Statement'"
        row = row + 1
        
        .Cells(row, 1).Value = "- قائمة المركز المالي: استخدم 'Balance_Sheet'"
        row = row + 1
        
        .Cells(row, 1).Value = "- ميزان المراجعة: استخدم 'Trial_Balance'"
        row = row + 1
        
        .Cells(row, 1).Value = ""
        row = row + 1
        
        .Cells(row, 1).Value = "6. ملاحظات مهمة"
        .Cells(row, 1).Font.Bold = True
        .Cells(row, 1).Interior.Color = RGB(200, 220, 255)
        row = row + 1
        
        .Cells(row, 1).Value = "✓ جميع المعاملات ترحل تلقائياً بدون أخطاء"
        row = row + 1
        
        .Cells(row, 1).Value = "✓ ضريبة القيمة المضافة تحسب تلقائياً (15%)"
        row = row + 1
        
        .Cells(row, 1).Value = "✓ لا تحذف أي أوراق - كل شيء مترابط"
        row = row + 1
        
        .Cells(row, 1).Value = "✓ احفظ الملف بانتظام لتجنب فقدان البيانات"
        row = row + 1
        
        .Cells(row, 1).Value = "✓ يمكنك عمل نسخة احتياطية من الملف"
    End With
    
    ws.Columns("A:C").AutoFit
    ws.Range("A:A").ColumnWidth = 70
    MsgBox "✅ تم فتح دليل الاستخدام بنجاح!", vbInformation
    
    Exit Sub
    
ErrorHandler:
    MsgBox "❌ خطأ: " & Err.Description, vbCritical
End Sub

Public Sub ShowAboutSystem()
    MsgBox "نظام ERP محاسبي متكامل سعودي" & vbCrLf & vbCrLf & _
    "الإصدار: 1.0" & vbCrLf & _
    "تاريخ الإنشاء: " & Format(Date, "dd/mm/yyyy") & vbCrLf & vbCrLf & _
    "المميزات:" & vbCrLf & _
    "✓ شجرة حسابات سعودية قياسية" & vbCrLf & _
    "✓ ترحيل آلي بدون أخطاء" & vbCrLf & _
    "✓ ضريبة القيمة المضافة 15%" & vbCrLf & _
    "✓ قوائم مالية دقيقة" & vbCrLf & _
    "✓ إدارة متكاملة للمبيعات والمشتريات" & vbCrLf & _
    "✓ نظام رواتب آلي" & vbCrLf & _
    "✓ سهل الاستخدام وآمن" & vbCrLf & vbCrLf & _
    "© جميع الحقوق محفوظة", vbInformation, "عن النظام"
End Sub

Public Sub CheckSystemHealth()
    On Error GoTo ErrorHandler
    
    Dim requiredSheets As Variant
    Dim i As Integer
    Dim sheetFound As Boolean
    Dim missingSheets As String
    
    requiredSheets = Array("Dashboard", "Chart_of_Accounts", "General_Journal", "General_Ledger", _
                           "Income_Statement", "Balance_Sheet", "Trial_Balance")
    
    missingSheets = ""
    
    For i = LBound(requiredSheets) To UBound(requiredSheets)
        sheetFound = False
        Dim sh As Worksheet
        For Each sh In ThisWorkbook.Sheets
            If sh.Name = requiredSheets(i) Then
                sheetFound = True
                Exit For
            End If
        Next sh
        
        If Not sheetFound Then
            missingSheets = missingSheets & vbCrLf & "- " & requiredSheets(i)
        End If
    Next i
    
    If missingSheets = "" Then
        MsgBox "✅ صحة النظام تمام التمام!" & vbCrLf & vbCrLf & _
        "جميع الوحدات موجودة وجاهزة للعمل", vbInformation, "فحص صحة النظام"
    Else
        MsgBox "⚠️ تحذير: بعض الوحدات مفقودة:" & missingSheets & vbCrLf & vbCrLf & _
        "يرجى إعادة تهيئة النظام", vbExclamation, "فحص صحة النظام"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "❌ خطأ في الفحص: " & Err.Description, vbCritical
End Sub

' =====================================================
' قائمة تقارير (Reports Menu)
' =====================================================

Public Sub GenerateAllReports()
    On Error GoTo ErrorHandler
    
    Dim response As Integer
    
    response = MsgBox("اختر التقرير المطلوب:" & vbCrLf & vbCrLf & _
    "1 - قائمة الدخل" & vbCrLf & _
    "2 - قائمة المركز المالي" & vbCrLf & _
    "3 - ميزان المراجعة" & vbCrLf & _
    "4 - جميع التقارير", _
    vbYesNoCancel + vbQuestion)
    
    Select Case response
        Case vbYes
            Call GenerateIncomeStatement
        Case vbNo
            Call GenerateBalanceSheet
        Case vbCancel
            Call GenerateTrialBalance
    End Select
    
    Exit Sub
    
ErrorHandler:
    MsgBox "❌ خطأ: " & Err.Description, vbCritical
End Sub

' =====================================================
' وظائف الدعم الفني
' =====================================================

Public Sub ExportDataToBackup()
    On Error GoTo ErrorHandler
    
    Dim response As Integer
    
    response = MsgBox("هل تريد إنشاء نسخة احتياطية من البيانات؟", vbYesNo + vbQuestion)
    
    If response = vbYes Then
        Dim backupPath As String
        backupPath = ThisWorkbook.Path & "\Backup_" & Format(Date, "ddMMyyyy_hhmmss") & ".xlsx"
        
        ThisWorkbook.SaveCopyAs backupPath
        MsgBox "✅ تم إنشاء نسخة احتياطية بنجاح!" & vbCrLf & "المسار: " & backupPath, vbInformation
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "❌ خطأ في النسخ الاحتياطي: " & Err.Description, vbCritical
End Sub

Public Sub ClearAllData()
    On Error GoTo ErrorHandler
    
    Dim response As Integer
    
    response = MsgBox("⚠️ تحذير: هذه العملية ستحذف جميع البيانات!" & vbCrLf & vbCrLf & _
    "هل أنت متأكد؟ (هذا لا يمكن التراجع عنه)", vbYesNo + vbCritical)
    
    If response = vbYes Then
        Dim sh As Worksheet
        For Each sh In ThisWorkbook.Sheets
            If sh.Name <> "Dashboard" Then
                On Error Resume Next
                sh.Cells.Clear
                On Error GoTo ErrorHandler
            End If
        Next sh
        
        MsgBox "✅ تم حذف جميع البيانات بنجاح!", vbInformation
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "❌ خطأ: " & Err.Description, vbCritical
End Sub
