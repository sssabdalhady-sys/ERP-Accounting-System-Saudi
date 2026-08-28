Option Explicit

' =====================================================
' وحدة الترحيل الآلي (Auto Posting Module)
' Automatic Journal Posting - SAE & VAT System
' =====================================================

' =========================================
' 1. ترحيل المبيعات تلقائياً
' =========================================

Public Function AutoPostSales(invoiceNum As String, invDate As Date, customerAccount As String, totalAmount As Double, vatAmount As Double, netAmount As Double, description As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim journalWs As Worksheet
    Dim ledgerWs As Worksheet
    Dim journalRow As Long
    Dim debitAccount As String
    Dim creditAccount As String
    Dim creditVATAccount As String
    
    Set journalWs = ThisWorkbook.Sheets("General_Journal")
    Set ledgerWs = ThisWorkbook.Sheets("General_Ledger")
    
    ' البحث عن آخر صف في دفتر اليومية
    journalRow = journalWs.Cells(journalWs.Rows.Count, 1).End(xlUp).Row + 1
    
    debitAccount = "1010" ' ذمم مدينة - عملاء
    creditAccount = "4001" ' مبيعات التكييفات
    creditVATAccount = "2030" ' ضريبة القيمة المضافة المستحقة
    
    ' إدخال القيد الأول - المدين (الذمم المدينة)
    With journalWs
        .Cells(journalRow, 1).Value = invoiceNum
        .Cells(journalRow, 2).Value = invDate
        .Cells(journalRow, 3).Value = debitAccount
        .Cells(journalRow, 4).Value = "ذمم مدينة - عملاء"
        .Cells(journalRow, 5).Value = description & " - فاتورة مبيعات"
        .Cells(journalRow, 6).Value = totalAmount ' المبلغ الإجمالي (متضمن الضريبة)
        .Cells(journalRow, 7).Value = 0
        .Cells(journalRow, 8).Value = "Pending"
        
        journalRow = journalRow + 1
        
        ' إدخال القيد الثاني - الدائن (الإيرادات)
        .Cells(journalRow, 1).Value = invoiceNum
        .Cells(journalRow, 2).Value = invDate
        .Cells(journalRow, 3).Value = creditAccount
        .Cells(journalRow, 4).Value = "مبيعات التكييفات"
        .Cells(journalRow, 5).Value = description & " - إيرادات"
        .Cells(journalRow, 6).Value = 0
        .Cells(journalRow, 7).Value = netAmount ' صافي المبيعات (بدون ضريبة)
        .Cells(journalRow, 8).Value = "Pending"
        
        journalRow = journalRow + 1
        
        ' إدخال القيد الثالث - ضريبة القيمة المضافة
        .Cells(journalRow, 1).Value = invoiceNum
        .Cells(journalRow, 2).Value = invDate
        .Cells(journalRow, 3).Value = creditVATAccount
        .Cells(journalRow, 4).Value = "ضريبة القيمة المضافة - مستحقة"
        .Cells(journalRow, 5).Value = description & " - ضريبة"
        .Cells(journalRow, 6).Value = 0
        .Cells(journalRow, 7).Value = vatAmount
        .Cells(journalRow, 8).Value = "Pending"
    End With
    
    ' ترحيل تلقائي إلى الأستاذ العام
    Call PostToGeneralLedger(debitAccount, "ذمم مدينة - عملاء", invDate, description & " - فاتورة", totalAmount, 0)
    Call PostToGeneralLedger(creditAccount, "مبيعات التكييفات", invDate, description, 0, netAmount)
    Call PostToGeneralLedger(creditVATAccount, "ضريبة القيمة المضافة - مستحقة", invDate, description, 0, vatAmount)
    
    ' تحديث شجرة الحسابات بالرصيد الجديد
    Call UpdateAccountBalance(debitAccount, totalAmount)
    Call UpdateAccountBalance(creditAccount, -netAmount)
    Call UpdateAccountBalance(creditVATAccount, -vatAmount)
    
    AutoPostSales = True
    Exit Function
    
ErrorHandler:
    MsgBox "❌ خطأ في ترحيل المبيعات: " & Err.Description, vbCritical
    AutoPostSales = False
End Function

' =========================================
' 2. ترحيل المشتريات تلقائياً
' =========================================

Public Function AutoPostPurchase(poNum As String, purchDate As Date, supplierAccount As String, totalAmount As Double, vatAmount As Double, netAmount As Double, description As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim journalWs As Worksheet
    Dim ledgerWs As Worksheet
    Dim journalRow As Long
    Dim debitAccount As String
    Dim creditAccount As String
    Dim debitVATAccount As String
    
    Set journalWs = ThisWorkbook.Sheets("General_Journal")
    Set ledgerWs = ThisWorkbook.Sheets("General_Ledger")
    
    journalRow = journalWs.Cells(journalWs.Rows.Count, 1).End(xlUp).Row + 1
    
    debitAccount = "5001" ' تكلفة المبيعات
    creditAccount = "2001" ' ذمم دائنة - موردين
    debitVATAccount = "2031" ' ضريبة القيمة المضافة - مدفوعة (استرجاع)
    
    ' إدخال القيود
    With journalWs
        ' القيد الأول - تكلفة المشتريات
        .Cells(journalRow, 1).Value = poNum
        .Cells(journalRow, 2).Value = purchDate
        .Cells(journalRow, 3).Value = debitAccount
        .Cells(journalRow, 4).Value = "تكلفة المبيعات"
        .Cells(journalRow, 5).Value = description & " - فاتورة شراء"
        .Cells(journalRow, 6).Value = netAmount
        .Cells(journalRow, 7).Value = 0
        .Cells(journalRow, 8).Value = "Pending"
        
        journalRow = journalRow + 1
        
        ' القيد الثاني - الذمم الدائنة
        .Cells(journalRow, 1).Value = poNum
        .Cells(journalRow, 2).Value = purchDate
        .Cells(journalRow, 3).Value = creditAccount
        .Cells(journalRow, 4).Value = "ذمم دائنة - موردين"
        .Cells(journalRow, 5).Value = description & " - دين"
        .Cells(journalRow, 6).Value = 0
        .Cells(journalRow, 7).Value = totalAmount
        .Cells(journalRow, 8).Value = "Pending"
        
        journalRow = journalRow + 1
        
        ' القيد الثالث - ضريبة مدفوعة
        .Cells(journalRow, 1).Value = poNum
        .Cells(journalRow, 2).Value = purchDate
        .Cells(journalRow, 3).Value = debitVATAccount
        .Cells(journalRow, 4).Value = "ضريبة القيمة المضافة - مدفوعة"
        .Cells(journalRow, 5).Value = description & " - ضريبة"
        .Cells(journalRow, 6).Value = vatAmount
        .Cells(journalRow, 7).Value = 0
        .Cells(journalRow, 8).Value = "Pending"
    End With
    
    ' ترحيل تلقائي
    Call PostToGeneralLedger(debitAccount, "تكلفة المبيعات", purchDate, description, netAmount, 0)
    Call PostToGeneralLedger(creditAccount, "ذمم دائنة - موردين", purchDate, description, 0, totalAmount)
    Call PostToGeneralLedger(debitVATAccount, "ضريبة القيمة المضافة - مدفوعة", purchDate, description, vatAmount, 0)
    
    ' تحديث الأرصدة
    Call UpdateAccountBalance(debitAccount, netAmount)
    Call UpdateAccountBalance(creditAccount, totalAmount)
    Call UpdateAccountBalance(debitVATAccount, vatAmount)
    
    AutoPostPurchase = True
    Exit Function
    
ErrorHandler:
    MsgBox "❌ خطأ في ترحيل المشتريات: " & Err.Description, vbCritical
    AutoPostPurchase = False
End Function

' =========================================
' 3. ترحيل الرواتب تلقائياً
' =========================================

Public Function AutoPostPayroll(payrollDate As Date, employeeName As String, salary As Double, allowances As Double, deductions As Double, netPay As Double) As Boolean
    On Error GoTo ErrorHandler
    
    Dim journalWs As Worksheet
    Dim journalRow As Long
    Dim salaryAccount As String
    Dim allowanceAccount As String
    Dim deductionAccount As String
    Dim payableAccount As String
    Dim totalExpense As Double
    
    Set journalWs = ThisWorkbook.Sheets("General_Journal")
    
    journalRow = journalWs.Cells(journalWs.Rows.Count, 1).End(xlUp).Row + 1
    
    salaryAccount = "5010" ' رواتب الموظفين
    allowanceAccount = "5011" ' بدلات الموظفين
    deductionAccount = "5012" ' خصومات
    payableAccount = "2010" ' الرواتب المستحقة
    
    totalExpense = salary + allowances
    
    With journalWs
        ' قيد الراتب الأساسي
        .Cells(journalRow, 1).Value = "PR-" & Format(payrollDate, "ddMMyyyy")
        .Cells(journalRow, 2).Value = payrollDate
        .Cells(journalRow, 3).Value = salaryAccount
        .Cells(journalRow, 4).Value = "رواتب الموظفين"
        .Cells(journalRow, 5).Value = "راتب " & employeeName
        .Cells(journalRow, 6).Value = salary
        .Cells(journalRow, 7).Value = 0
        .Cells(journalRow, 8).Value = "Pending"
        
        journalRow = journalRow + 1
        
        ' قيد البدلات
        If allowances > 0 Then
            .Cells(journalRow, 1).Value = "PR-" & Format(payrollDate, "ddMMyyyy")
            .Cells(journalRow, 2).Value = payrollDate
            .Cells(journalRow, 3).Value = allowanceAccount
            .Cells(journalRow, 4).Value = "بدلات الموظفين"
            .Cells(journalRow, 5).Value = "بدلات " & employeeName
            .Cells(journalRow, 6).Value = allowances
            .Cells(journalRow, 7).Value = 0
            .Cells(journalRow, 8).Value = "Pending"
            
            journalRow = journalRow + 1
        End If
        
        ' قيد الخصومات
        If deductions > 0 Then
            .Cells(journalRow, 1).Value = "PR-" & Format(payrollDate, "ddMMyyyy")
            .Cells(journalRow, 2).Value = payrollDate
            .Cells(journalRow, 3).Value = deductionAccount
            .Cells(journalRow, 4).Value = "خصومات الموظفين"
            .Cells(journalRow, 5).Value = "خصومات " & employeeName
            .Cells(journalRow, 6).Value = 0
            .Cells(journalRow, 7).Value = deductions
            .Cells(journalRow, 8).Value = "Pending"
            
            journalRow = journalRow + 1
        End If
        
        ' قيد الرواتب المستحقة
        .Cells(journalRow, 1).Value = "PR-" & Format(payrollDate, "ddMMyyyy")
        .Cells(journalRow, 2).Value = payrollDate
        .Cells(journalRow, 3).Value = payableAccount
        .Cells(journalRow, 4).Value = "الرواتب المستحقة"
        .Cells(journalRow, 5).Value = "دفع " & employeeName
        .Cells(journalRow, 6).Value = 0
        .Cells(journalRow, 7).Value = netPay
        .Cells(journalRow, 8).Value = "Pending"
    End With
    
    ' الترحيل الآلي
    Call PostToGeneralLedger(salaryAccount, "رواتب الموظفين", payrollDate, "راتب " & employeeName, salary, 0)
    If allowances > 0 Then
        Call PostToGeneralLedger(allowanceAccount, "بدلات الموظفين", payrollDate, "بدلات " & employeeName, allowances, 0)
    End If
    If deductions > 0 Then
        Call PostToGeneralLedger(deductionAccount, "خصومات الموظفين", payrollDate, "خصومات " & employeeName, 0, deductions)
    End If
    Call PostToGeneralLedger(payableAccount, "الرواتب المستحقة", payrollDate, "دفع " & employeeName, 0, netPay)
    
    ' تحديث الأرصدة
    Call UpdateAccountBalance(salaryAccount, salary)
    If allowances > 0 Then Call UpdateAccountBalance(allowanceAccount, allowances)
    If deductions > 0 Then Call UpdateAccountBalance(deductionAccount, -deductions)
    Call UpdateAccountBalance(payableAccount, -netPay)
    
    AutoPostPayroll = True
    Exit Function
    
ErrorHandler:
    MsgBox "❌ خطأ في ترحيل الرواتب: " & Err.Description, vbCritical
    AutoPostPayroll = False
End Function

' =========================================
' 4. ترحيل المصروفات تلقائياً
' =========================================

Public Function AutoPostExpense(expenseDate As Date, expenseType As String, accountCode As String, accountName As String, amount As Double, description As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim journalWs As Worksheet
    Dim journalRow As Long
    Dim bankAccount As String
    
    Set journalWs = ThisWorkbook.Sheets("General_Journal")
    
    journalRow = journalWs.Cells(journalWs.Rows.Count, 1).End(xlUp).Row + 1
    bankAccount = "1002" ' النقد في البنك
    
    With journalWs
        ' قيد المصروف
        .Cells(journalRow, 1).Value = "EXP-" & Format(expenseDate, "ddMMyyyy")
        .Cells(journalRow, 2).Value = expenseDate
        .Cells(journalRow, 3).Value = accountCode
        .Cells(journalRow, 4).Value = accountName
        .Cells(journalRow, 5).Value = description
        .Cells(journalRow, 6).Value = amount
        .Cells(journalRow, 7).Value = 0
        .Cells(journalRow, 8).Value = "Pending"
        
        journalRow = journalRow + 1
        
        ' قيد البنك (الدائن)
        .Cells(journalRow, 1).Value = "EXP-" & Format(expenseDate, "ddMMyyyy")
        .Cells(journalRow, 2).Value = expenseDate
        .Cells(journalRow, 3).Value = bankAccount
        .Cells(journalRow, 4).Value = "النقد في البنك"
        .Cells(journalRow, 5).Value = "دفع " & description
        .Cells(journalRow, 6).Value = 0
        .Cells(journalRow, 7).Value = amount
        .Cells(journalRow, 8).Value = "Pending"
    End With
    
    ' الترحيل الآلي
    Call PostToGeneralLedger(accountCode, accountName, expenseDate, description, amount, 0)
    Call PostToGeneralLedger(bankAccount, "النقد في البنك", expenseDate, "دفع " & description, 0, amount)
    
    ' تحديث الأرصدة
    Call UpdateAccountBalance(accountCode, amount)
    Call UpdateAccountBalance(bankAccount, -amount)
    
    AutoPostExpense = True
    Exit Function
    
ErrorHandler:
    MsgBox "❌ خطأ في ترحيل المصروف: " & Err.Description, vbCritical
    AutoPostExpense = False
End Function

' =========================================
' 5. ترحيل إلى الأستاذ العام
' =========================================

Private Function PostToGeneralLedger(accountCode As String, accountName As String, postDate As Date, description As String, debit As Double, credit As Double) As Boolean
    On Error GoTo ErrorHandler
    
    Dim ledgerWs As Worksheet
    Dim ledgerRow As Long
    Dim previousBalance As Double
    Dim newBalance As Double
    
    Set ledgerWs = ThisWorkbook.Sheets("General_Ledger")
    
    ' البحث عن آخر رصيد للحساب
    previousBalance = GetAccountBalance(accountCode)
    newBalance = previousBalance + debit - credit
    
    ' إضافة القيد
    ledgerRow = ledgerWs.Cells(ledgerWs.Rows.Count, 1).End(xlUp).Row + 1
    
    With ledgerWs
        .Cells(ledgerRow, 1).Value = accountCode
        .Cells(ledgerRow, 2).Value = accountName
        .Cells(ledgerRow, 3).Value = postDate
        .Cells(ledgerRow, 4).Value = description
        .Cells(ledgerRow, 5).Value = debit
        .Cells(ledgerRow, 6).Value = credit
        .Cells(ledgerRow, 7).Value = newBalance
        
        ' تنسيق
        .Cells(ledgerRow, 3).NumberFormat = "mm/dd/yyyy"
        .Cells(ledgerRow, 5).NumberFormat = "#,##0.00"
        .Cells(ledgerRow, 6).NumberFormat = "#,##0.00"
        .Cells(ledgerRow, 7).NumberFormat = "#,##0.00"
    End With
    
    PostToGeneralLedger = True
    Exit Function
    
ErrorHandler:
    MsgBox "❌ خطأ في الترحيل: " & Err.Description, vbCritical
    PostToGeneralLedger = False
End Function

' =========================================
' 6. الحصول على رصيد الحساب الحالي
' =========================================

Private Function GetAccountBalance(accountCode As String) As Double
    On Error GoTo ErrorHandler
    
    Dim ledgerWs As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim balance As Double
    
    Set ledgerWs = ThisWorkbook.Sheets("General_Ledger")
    lastRow = ledgerWs.Cells(ledgerWs.Rows.Count, 1).End(xlUp).Row
    
    balance = 0
    
    ' البحث عن آخر رصيد للحساب
    For i = lastRow To 2 Step -1
        If ledgerWs.Cells(i, 1).Value = accountCode Then
            balance = ledgerWs.Cells(i, 7).Value
            Exit For
        End If
    Next i
    
    GetAccountBalance = balance
    Exit Function
    
ErrorHandler:
    GetAccountBalance = 0
End Function

' =========================================
' 7. تحديث رصيد الحساب في شجرة الحسابات
' =========================================

Private Function UpdateAccountBalance(accountCode As String, changeAmount As Double) As Boolean
    On Error GoTo ErrorHandler
    
    Dim chartWs As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Set chartWs = ThisWorkbook.Sheets("Chart_of_Accounts")
    lastRow = chartWs.Cells(chartWs.Rows.Count, 1).End(xlUp).Row
    
    ' البحث عن الحساب وتحديث رصيده
    For i = 2 To lastRow
        If chartWs.Cells(i, 1).Value = accountCode Then
            chartWs.Cells(i, 5).Value = chartWs.Cells(i, 5).Value + changeAmount
            chartWs.Cells(i, 5).NumberFormat = "#,##0.00"
            UpdateAccountBalance = True
            Exit Function
        End If
    Next i
    
    UpdateAccountBalance = False
    Exit Function
    
ErrorHandler:
    MsgBox "❌ خطأ في تحديث الرصيد: " & Err.Description, vbCritical
    UpdateAccountBalance = False
End Function

' =========================================
' 8. ترحيل جماعي لجميع المعاملات المعلقة
' =========================================

Public Sub AutoPostAllPendingTransactions()
    On Error GoTo ErrorHandler
    
    Dim journalWs As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim postedCount As Integer
    
    Set journalWs = ThisWorkbook.Sheets("General_Journal")
    lastRow = journalWs.Cells(journalWs.Rows.Count, 1).End(xlUp).Row
    
    postedCount = 0
    
    ' ترحيل جميع المعاملات غير المرحلة
    For i = 2 To lastRow
        If journalWs.Cells(i, 8).Value = "Pending" Then
            journalWs.Cells(i, 8).Value = "Posted"
            postedCount = postedCount + 1
        End If
    Next i
    
    MsgBox "✅ تم ترحيل " & postedCount & " معاملة بنجاح!", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "❌ خطأ في الترحيل الجماعي: " & Err.Description, vbCritical
End Sub
