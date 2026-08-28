Option Explicit

' =====================================================
' وحدة المبيعات (Sales Module)
' Sales Invoices & Customer Management
' =====================================================

Public Function CreateSalesInvoice(invoiceNum As String, invDate As Date, customerName As String, customerPhone As String, items As Collection, totalBeforeVAT As Double) As Boolean
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim sheetExists As Boolean
    Dim row As Long
    Dim i As Integer
    Dim vatAmount As Double
    Dim totalAmount As Double
    
    sheetExists = False
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Sheets
        If sh.Name = "Sales_Invoices" Then
            sheetExists = True
            Set ws = sh
            Exit For
        End If
    Next sh
    
    If Not sheetExists Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "Sales_Invoices"
        
        ' رؤوس الأعمدة
        With ws.Range("A1:H1")
            .Value = Array("Invoice #", "Date", "Customer", "Phone", "Item Description", "Quantity", "Unit Price", "Line Total")
            .Font.Bold = True
            .Interior.Color = RGB(0, 102, 204)
            .Font.Color = RGB(255, 255, 255)
        End With
        
        ' صف ملخص في الأعلى
        ws.Range("A2").Value = "Total Before VAT"
        ws.Range("A3").Value = "VAT (15%)"
        ws.Range("A4").Value = "Total Amount"
    End If
    
    vatAmount = totalBeforeVAT * 0.15
    totalAmount = totalBeforeVAT + vatAmount
    
    row = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    ' إدخال بيانات الفاتورة
    For i = 0 To items.Count - 1
        With ws
            .Cells(row, 1).Value = invoiceNum
            .Cells(row, 2).Value = invDate
            .Cells(row, 2).NumberFormat = "mm/dd/yyyy"
            .Cells(row, 3).Value = customerName
            .Cells(row, 4).Value = customerPhone
            
            ' تفاصيل الصنف (المتوقع أن يكون كل صنف array من 3 عناصر: وصف، كمية، سعر)
            ' يمكن تعديل حسب احتياجك
            .Cells(row, 5).Value = "Item " & (i + 1)
            .Cells(row, 6).Value = 1
            .Cells(row, 7).Value = totalBeforeVAT / items.Count
            .Cells(row, 8).Value = .Cells(row, 6).Value * .Cells(row, 7).Value
            .Cells(row, 8).NumberFormat = "#,##0.00"
        End With
        row = row + 1
    Next i
    
    ' ترحيل آلي
    Call AutoPostSales(invoiceNum, invDate, customerName, totalAmount, vatAmount, totalBeforeVAT, customerName & " - فاتورة رقم " & invoiceNum)
    
    ' إضافة العميل إلى قائمة العملاء
    Call AddOrUpdateCustomer(customerName, customerPhone, totalAmount)
    
    CreateSalesInvoice = True
    MsgBox "✅ تم إنشاء فاتورة المبيعات رقم " & invoiceNum & " بنجاح!" & vbCrLf & "الإجمالي: " & Format(totalAmount, "#,##0.00") & " ريال", vbInformation
    Exit Function
    
ErrorHandler:
    MsgBox "❌ خطأ في إنشاء فاتورة المبيعات: " & Err.Description, vbCritical
    CreateSalesInvoice = False
End Function

Public Sub ManageSalesInvoices()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim sheetExists As Boolean
    
    sheetExists = False
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Sheets
        If sh.Name = "Sales_Invoices" Then
            sheetExists = True
            Set ws = sh
            Exit For
        End If
    Next sh
    
    If Not sheetExists Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "Sales_Invoices"
    End If
    
    With ws.Range("A1:H1")
        .Value = Array("Invoice #", "Date", "Customer", "Phone", "Item Description", "Quantity", "Unit Price", "Line Total")
        .Font.Bold = True
        .Interior.Color = RGB(0, 102, 204)
        .Font.Color = RGB(255, 255, 255)
    End With
    
    ws.Columns("A:H").AutoFit
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
    
    Exit Sub
    
ErrorHandler:
    MsgBox "❌ خطأ: " & Err.Description, vbCritical
End Sub

' =====================================================
' وحدة المشتريات (Purchases Module)
' Purchase Orders & Supplier Management
' =====================================================

Public Function CreatePurchaseOrder(poNum As String, purchDate As Date, supplierName As String, supplierPhone As String, items As Collection, totalBeforeVAT As Double) As Boolean
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim sheetExists As Boolean
    Dim row As Long
    Dim i As Integer
    Dim vatAmount As Double
    Dim totalAmount As Double
    
    sheetExists = False
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Sheets
        If sh.Name = "Purchase_Orders" Then
            sheetExists = True
            Set ws = sh
            Exit For
        End If
    Next sh
    
    If Not sheetExists Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "Purchase_Orders"
        
        With ws.Range("A1:H1")
            .Value = Array("PO #", "Date", "Supplier", "Phone", "Item Description", "Quantity", "Unit Price", "Line Total")
            .Font.Bold = True
            .Interior.Color = RGB(0, 102, 204)
            .Font.Color = RGB(255, 255, 255)
        End With
    End If
    
    vatAmount = totalBeforeVAT * 0.15
    totalAmount = totalBeforeVAT + vatAmount
    
    row = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    For i = 0 To items.Count - 1
        With ws
            .Cells(row, 1).Value = poNum
            .Cells(row, 2).Value = purchDate
            .Cells(row, 2).NumberFormat = "mm/dd/yyyy"
            .Cells(row, 3).Value = supplierName
            .Cells(row, 4).Value = supplierPhone
            .Cells(row, 5).Value = "Item " & (i + 1)
            .Cells(row, 6).Value = 1
            .Cells(row, 7).Value = totalBeforeVAT / items.Count
            .Cells(row, 8).Value = .Cells(row, 6).Value * .Cells(row, 7).Value
            .Cells(row, 8).NumberFormat = "#,##0.00"
        End With
        row = row + 1
    Next i
    
    ' ترحيل آلي
    Call AutoPostPurchase(poNum, purchDate, supplierName, totalAmount, vatAmount, totalBeforeVAT, supplierName & " - طلبية شراء رقم " & poNum)
    
    ' إضافة الموردين إلى قائمة الموردين
    Call AddOrUpdateSupplier(supplierName, supplierPhone, totalAmount)
    
    CreatePurchaseOrder = True
    MsgBox "✅ تم إنشاء طلبية الشراء رقم " & poNum & " بنجاح!" & vbCrLf & "الإجمالي: " & Format(totalAmount, "#,##0.00") & " ريال", vbInformation
    Exit Function
    
ErrorHandler:
    MsgBox "❌ خطأ في إنشاء طلبية الشراء: " & Err.Description, vbCritical
    CreatePurchaseOrder = False
End Function

Public Sub ManagePurchaseOrders()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim sheetExists As Boolean
    
    sheetExists = False
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Sheets
        If sh.Name = "Purchase_Orders" Then
            sheetExists = True
            Set ws = sh
            Exit For
        End If
    Next sh
    
    If Not sheetExists Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "Purchase_Orders"
    End If
    
    With ws.Range("A1:H1")
        .Value = Array("PO #", "Date", "Supplier", "Phone", "Item Description", "Quantity", "Unit Price", "Line Total")
        .Font.Bold = True
        .Interior.Color = RGB(0, 102, 204)
        .Font.Color = RGB(255, 255, 255)
    End With
    
    ws.Columns("A:H").AutoFit
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
    
    Exit Sub
    
ErrorHandler:
    MsgBox "❌ خطأ: " & Err.Description, vbCritical
End Sub

' =====================================================
' وحدة الرواتب (Payroll Module)
' Employee Management & Salary Calculation
' =====================================================

Public Function ProcessPayroll(payrollMonth As Date) As Boolean
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim sheetExists As Boolean
    Dim row As Long
    
    sheetExists = False
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Sheets
        If sh.Name = "Payroll" Then
            sheetExists = True
            Set ws = sh
            Exit For
        End If
    Next sh
    
    If Not sheetExists Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "Payroll"
        
        With ws.Range("A1:H1")
            .Value = Array("Employee", "Salary", "Allowances", "Deductions", "Net Pay", "Date", "Status", "Posted")
            .Font.Bold = True
            .Interior.Color = RGB(0, 102, 204)
            .Font.Color = RGB(255, 255, 255)
        End With
    End If
    
    MsgBox "✅ تم فتح وحدة الرواتب لإدخال بيانات موظفي شهر " & Format(payrollMonth, "mmmm/yyyy"), vbInformation
    ProcessPayroll = True
    
    Exit Function
    
ErrorHandler:
    MsgBox "❌ خطأ في معالجة الرواتب: " & Err.Description, vbCritical
    ProcessPayroll = False
End Function

Public Function AddEmployeeToPayroll(employeeName As String, salary As Double, allowances As Double, deductions As Double) As Boolean
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim row As Long
    Dim netPay As Double
    
    Set ws = ThisWorkbook.Sheets("Payroll")
    row = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    netPay = salary + allowances - deductions
    
    With ws
        .Cells(row, 1).Value = employeeName
        .Cells(row, 2).Value = salary
        .Cells(row, 2).NumberFormat = "#,##0.00"
        .Cells(row, 3).Value = allowances
        .Cells(row, 3).NumberFormat = "#,##0.00"
        .Cells(row, 4).Value = deductions
        .Cells(row, 4).NumberFormat = "#,##0.00"
        .Cells(row, 5).Value = netPay
        .Cells(row, 5).NumberFormat = "#,##0.00"
        .Cells(row, 6).Value = Date
        .Cells(row, 6).NumberFormat = "mm/dd/yyyy"
        .Cells(row, 7).Value = "Pending"
        .Cells(row, 8).Value = "No"
    End With
    
    ' الترحيل الآلي
    Call AutoPostPayroll(Date, employeeName, salary, allowances, deductions, netPay)
    
    AddEmployeeToPayroll = True
    MsgBox "✅ تم إضافة " & employeeName & " إلى كشف الرواتب!" & vbCrLf & "الراتب الصافي: " & Format(netPay, "#,##0.00") & " ريال", vbInformation
    
    Exit Function
    
ErrorHandler:
    MsgBox "❌ خطأ في إضافة الموظف: " & Err.Description, vbCritical
    AddEmployeeToPayroll = False
End Function

Public Sub ManageEmployees()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim sheetExists As Boolean
    
    sheetExists = False
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Sheets
        If sh.Name = "Employees" Then
            sheetExists = True
            Set ws = sh
            Exit For
        End If
    Next sh
    
    If Not sheetExists Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "Employees"
    End If
    
    With ws.Range("A1:G1")
        .Value = Array("ID", "Name", "Position", "Salary", "Phone", "Email", "Hire Date")
        .Font.Bold = True
        .Interior.Color = RGB(0, 102, 204)
        .Font.Color = RGB(255, 255, 255)
    End With
    
    ws.Columns("A:G").AutoFit
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
    
    Exit Sub
    
ErrorHandler:
    MsgBox "❌ خطأ: " & Err.Description, vbCritical
End Sub

' =====================================================
' وحدة المخزون (Inventory Module)
' =====================================================

Public Sub ManageInventory()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim sheetExists As Boolean
    
    sheetExists = False
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Sheets
        If sh.Name = "Inventory" Then
            sheetExists = True
            Set ws = sh
            Exit For
        End If
    Next sh
    
    If Not sheetExists Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "Inventory"
    End If
    
    With ws.Range("A1:H1")
        .Value = Array("Product ID", "Product Name", "Category", "Beginning Stock", "Quantity In", "Quantity Out", "Ending Stock", "Unit Cost")
        .Font.Bold = True
        .Interior.Color = RGB(0, 102, 204)
        .Font.Color = RGB(255, 255, 255)
    End With
    
    ' إضافة بعض الأصناف الافتراضية
    Dim row As Integer
    row = 2
    
    With ws
        .Cells(row, 1).Value = "AC-001"
        .Cells(row, 2).Value = "تكييف 1.5 حصان"
        .Cells(row, 3).Value = "Air Conditioners"
        .Cells(row, 4).Value = 10
        .Cells(row, 5).Value = 0
        .Cells(row, 6).Value = 0
        .Cells(row, 7).Formula = "=D" & row & "+E" & row & "-F" & row
        .Cells(row, 8).Value = 2000
        .Cells(row, 8).NumberFormat = "#,##0.00"
        
        row = row + 1
        
        .Cells(row, 1).Value = "AC-002"
        .Cells(row, 2).Value = "تكييف 2 حصان"
        .Cells(row, 3).Value = "Air Conditioners"
        .Cells(row, 4).Value = 8
        .Cells(row, 5).Value = 0
        .Cells(row, 6).Value = 0
        .Cells(row, 7).Formula = "=D" & row & "+E" & row & "-F" & row
        .Cells(row, 8).Value = 2500
        .Cells(row, 8).NumberFormat = "#,##0.00"
        
        row = row + 1
        
        .Cells(row, 1).Value = "SP-001"
        .Cells(row, 2).Value = "قطع غيار مختلفة"
        .Cells(row, 3).Value = "Spare Parts"
        .Cells(row, 4).Value = 50
        .Cells(row, 5).Value = 0
        .Cells(row, 6).Value = 0
        .Cells(row, 7).Formula = "=D" & row & "+E" & row & "-F" & row
        .Cells(row, 8).Value = 500
        .Cells(row, 8).NumberFormat = "#,##0.00"
    End With
    
    ws.Columns("A:H").AutoFit
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
    
    Exit Sub
    
ErrorHandler:
    MsgBox "❌ خطأ: " & Err.Description, vbCritical
End Sub

' =====================================================
' وحدة إدارة العملاء (Customer Management)
' =====================================================

Private Function AddOrUpdateCustomer(customerName As String, customerPhone As String, balance As Double) As Boolean
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim sheetExists As Boolean
    Dim row As Long
    Dim lastRow As Long
    Dim found As Boolean
    
    sheetExists = False
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Sheets
        If sh.Name = "Customers" Then
            sheetExists = True
            Set ws = sh
            Exit For
        End If
    Next sh
    
    If Not sheetExists Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "Customers"
        
        With ws.Range("A1:E1")
            .Value = Array("Customer ID", "Name", "Phone", "Balance Due", "Registration Date")
            .Font.Bold = True
            .Interior.Color = RGB(0, 102, 204)
            .Font.Color = RGB(255, 255, 255)
        End With
    End If
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    found = False
    
    ' البحث عن العميل وتحديث الرصيد
    For row = 2 To lastRow
        If ws.Cells(row, 2).Value = customerName Then
            ws.Cells(row, 4).Value = ws.Cells(row, 4).Value + balance
            found = True
            Exit For
        End If
    Next row
    
    ' إضافة عميل جديد
    If Not found Then
        row = lastRow + 1
        With ws
            .Cells(row, 1).Value = "CUST-" & Format(row, "00000")
            .Cells(row, 2).Value = customerName
            .Cells(row, 3).Value = customerPhone
            .Cells(row, 4).Value = balance
            .Cells(row, 4).NumberFormat = "#,##0.00"
            .Cells(row, 5).Value = Date
            .Cells(row, 5).NumberFormat = "mm/dd/yyyy"
        End With
    End If
    
    ws.Columns("A:E").AutoFit
    AddOrUpdateCustomer = True
    
    Exit Function
    
ErrorHandler:
    AddOrUpdateCustomer = False
End Function

' =====================================================
' وحدة إدارة الموردين (Supplier Management)
' =====================================================

Private Function AddOrUpdateSupplier(supplierName As String, supplierPhone As String, balance As Double) As Boolean
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim sheetExists As Boolean
    Dim row As Long
    Dim lastRow As Long
    Dim found As Boolean
    
    sheetExists = False
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Sheets
        If sh.Name = "Suppliers" Then
            sheetExists = True
            Set ws = sh
            Exit For
        End If
    Next sh
    
    If Not sheetExists Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "Suppliers"
        
        With ws.Range("A1:E1")
            .Value = Array("Supplier ID", "Name", "Phone", "Amount Due", "Registration Date")
            .Font.Bold = True
            .Interior.Color = RGB(0, 102, 204)
            .Font.Color = RGB(255, 255, 255)
        End With
    End If
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    found = False
    
    ' البحث عن الموردين وتحديث الديون
    For row = 2 To lastRow
        If ws.Cells(row, 2).Value = supplierName Then
            ws.Cells(row, 4).Value = ws.Cells(row, 4).Value + balance
            found = True
            Exit For
        End If
    Next row
    
    ' إضافة موردين جديد
    If Not found Then
        row = lastRow + 1
        With ws
            .Cells(row, 1).Value = "SUPP-" & Format(row, "00000")
            .Cells(row, 2).Value = supplierName
            .Cells(row, 3).Value = supplierPhone
            .Cells(row, 4).Value = balance
            .Cells(row, 4).NumberFormat = "#,##0.00"
            .Cells(row, 5).Value = Date
            .Cells(row, 5).NumberFormat = "mm/dd/yyyy"
        End With
    End If
    
    ws.Columns("A:E").AutoFit
    AddOrUpdateSupplier = True
    
    Exit Function
    
ErrorHandler:
    AddOrUpdateSupplier = False
End Function
