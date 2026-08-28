Option Explicit

' =====================================================
' وحدة الرواتب والأجور المتقدمة (Advanced Payroll Module)
' Employee Salaries, Allowances, Deductions & Benefits
' =====================================================

' =========================================
' 1. إضافة موظف جديد
' =========================================

Public Function AddNewEmployee(empName As String, empID As String, empPosition As String, basicSalary As Double, _
                               empPhone As String, empEmail As String, hireDate As Date, department As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim sheetExists As Boolean
    Dim row As Long
    
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
        
        ' رؤوس الأعمدة
        With ws.Range("A1:H1")
            .Value = Array("Employee ID", "Name", "Position", "Department", "Basic Salary", "Phone", "Email", "Hire Date")
            .Font.Bold = True
            .Interior.Color = RGB(0, 102, 204)
            .Font.Color = RGB(255, 255, 255)
        End With
    End If
    
    row = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    With ws
        .Cells(row, 1).Value = empID
        .Cells(row, 2).Value = empName
        .Cells(row, 3).Value = empPosition
        .Cells(row, 4).Value = department
        .Cells(row, 5).Value = basicSalary
        .Cells(row, 5).NumberFormat = "#,##0.00"
        .Cells(row, 6).Value = empPhone
        .Cells(row, 7).Value = empEmail
        .Cells(row, 8).Value = hireDate
        .Cells(row, 8).NumberFormat = "mm/dd/yyyy"
    End With
    
    ws.Columns("A:H").AutoFit
    
    AddNewEmployee = True
    MsgBox "✅ تم إضافة الموظف " & empName & " بنجاح!", vbInformation
    Exit Function
    
ErrorHandler:
    MsgBox "❌ خطأ في إضافة الموظف: " & Err.Description, vbCritical
    AddNewEmployee = False
End Function

' =========================================
' 2. وحدة البدلات (Allowances)
' =========================================

Public Function AddAllowances(empID As String, empName As String, allowanceType As String, allowanceAmount As Double, payrollMonth As Date) As Boolean
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim sheetExists As Boolean
    Dim row As Long
    
    sheetExists = False
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Sheets
        If sh.Name = "Allowances" Then
            sheetExists = True
            Set ws = sh
            Exit For
        End If
    Next sh
    
    If Not sheetExists Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "Allowances"
        
        ' رؤوس الأعمدة
        With ws.Range("A1:E1")
            .Value = Array("Employee ID", "Employee Name", "Allowance Type", "Amount", "Payroll Month")
            .Font.Bold = True
            .Interior.Color = RGB(100, 200, 100)
            .Font.Color = RGB(255, 255, 255)
        End With
    End If
    
    row = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    With ws
        .Cells(row, 1).Value = empID
        .Cells(row, 2).Value = empName
        .Cells(row, 3).Value = allowanceType
        .Cells(row, 4).Value = allowanceAmount
        .Cells(row, 4).NumberFormat = "#,##0.00"
        .Cells(row, 5).Value = payrollMonth
        .Cells(row, 5).NumberFormat = "mm/yyyy"
    End With
    
    ws.Columns("A:E").AutoFit
    
    AddAllowances = True
    MsgBox "✅ تم إضافة البدلة: " & allowanceType & " بمبلغ " & Format(allowanceAmount, "#,##0.00"), vbInformation
    Exit Function
    
ErrorHandler:
    MsgBox "❌ خطأ في إضافة البدلة: " & Err.Description, vbCritical
    AddAllowances = False
End Function

' =========================================
' 3. وحدة الخصومات (Deductions)
' =========================================

Public Function AddDeductions(empID As String, empName As String, deductionType As String, deductionAmount As Double, payrollMonth As Date) As Boolean
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim sheetExists As Boolean
    Dim row As Long
    
    sheetExists = False
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Sheets
        If sh.Name = "Deductions" Then
            sheetExists = True
            Set ws = sh
            Exit For
        End If
    Next sh
    
    If Not sheetExists Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "Deductions"
        
        ' رؤوس الأعمدة
        With ws.Range("A1:E1")
            .Value = Array("Employee ID", "Employee Name", "Deduction Type", "Amount", "Payroll Month")
            .Font.Bold = True
            .Interior.Color = RGB(255, 100, 100)
            .Font.Color = RGB(255, 255, 255)
        End With
    End If
    
    row = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    With ws
        .Cells(row, 1).Value = empID
        .Cells(row, 2).Value = empName
        .Cells(row, 3).Value = deductionType
        .Cells(row, 4).Value = deductionAmount
        .Cells(row, 4).NumberFormat = "#,##0.00"
        .Cells(row, 5).Value = payrollMonth
        .Cells(row, 5).NumberFormat = "mm/yyyy"
    End With
    
    ws.Columns("A:E").AutoFit
    
    AddDeductions = True
    MsgBox "✅ تم إضافة الخصم: " & deductionType & " بمبلغ " & Format(deductionAmount, "#,##0.00"), vbInformation
    Exit Function
    
ErrorHandler:
    MsgBox "❌ خطأ في إضافة الخصم: " & Err.Description, vbCritical
    AddDeductions = False
End Function

' =========================================
' 4. معالجة كشف الرواتب الشهري
' =========================================

Public Function ProcessMonthlyPayroll(payrollMonth As Date) As Boolean
    On Error GoTo ErrorHandler
    
    Dim empWs As Worksheet
    Dim allowWs As Worksheet
    Dim dedWs As Worksheet
    Dim payrollWs As Worksheet
    Dim sheetExists As Boolean
    Dim empRow As Long
    Dim payRow As Long
    Dim lastEmpRow As Long
    Dim empID As String
    Dim empName As String
    Dim basicSalary As Double
    Dim totalAllowances As Double
    Dim totalDeductions As Double
    Dim netSalary As Double
    Dim i As Long
    
    ' التحقق من وجود الأوراق المطلوبة
    On Error Resume Next
    Set empWs = ThisWorkbook.Sheets("Employees")
    Set allowWs = ThisWorkbook.Sheets("Allowances")
    Set dedWs = ThisWorkbook.Sheets("Deductions")
    On Error GoTo ErrorHandler
    
    If empWs Is Nothing Then
        MsgBox "❌ لم يتم العثور على قائمة الموظفين!", vbCritical
        ProcessMonthlyPayroll = False
        Exit Function
    End If
    
    ' إنشاء ورقة كشف الرواتب
    sheetExists = False
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Sheets
        If sh.Name = "Monthly_Payroll_" & Format(payrollMonth, "mmyyyy") Then
            sheetExists = True
            Set payrollWs = sh
            Exit For
        End If
    Next sh
    
    If Not sheetExists Then
        Set payrollWs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        payrollWs.Name = "Monthly_Payroll_" & Format(payrollMonth, "mmyyyy")
    Else
        payrollWs.Cells.Clear
    End If
    
    ' العنوان
    With payrollWs.Range("A1:H1")
        .Merge
        .Value = "كشف الرواتب الشهري - " & Format(payrollMonth, "mmmm yyyy")
        .Font.Size = 14
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(0, 102, 204)
        .Font.Color = RGB(255, 255, 255)
    End With
    
    ' رؤوس الأعمدة
    With payrollWs.Range("A3:H3")
        .Value = Array("Employee ID", "Name", "Basic Salary", "Allowances", "Deductions", "Gross Salary", "Net Salary", "Status")
        .Font.Bold = True
        .Interior.Color = RGB(200, 220, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    payRow = 4
    lastEmpRow = empWs.Cells(empWs.Rows.Count, 1).End(xlUp).Row
    
    ' معالجة كل موظف
    For i = 2 To lastEmpRow
        empID = empWs.Cells(i, 1).Value
        empName = empWs.Cells(i, 2).Value
        basicSalary = empWs.Cells(i, 5).Value
        
        ' حساب البدلات الإجمالية
        totalAllowances = 0
        If Not allowWs Is Nothing Then
            Dim allowRow As Long
            Dim allowLastRow As Long
            allowLastRow = allowWs.Cells(allowWs.Rows.Count, 1).End(xlUp).Row
            
            For allowRow = 2 To allowLastRow
                If allowWs.Cells(allowRow, 1).Value = empID And _
                   Month(allowWs.Cells(allowRow, 5).Value) = Month(payrollMonth) And _
                   Year(allowWs.Cells(allowRow, 5).Value) = Year(payrollMonth) Then
                    totalAllowances = totalAllowances + allowWs.Cells(allowRow, 4).Value
                End If
            Next allowRow
        End If
        
        ' حساب الخصومات الإجمالية
        totalDeductions = 0
        If Not dedWs Is Nothing Then
            Dim dedRow As Long
            Dim dedLastRow As Long
            dedLastRow = dedWs.Cells(dedWs.Rows.Count, 1).End(xlUp).Row
            
            For dedRow = 2 To dedLastRow
                If dedWs.Cells(dedRow, 1).Value = empID And _
                   Month(dedWs.Cells(dedRow, 5).Value) = Month(payrollMonth) And _
                   Year(dedWs.Cells(dedRow, 5).Value) = Year(payrollMonth) Then
                    totalDeductions = totalDeductions + dedWs.Cells(dedRow, 4).Value
                End If
            Next dedRow
        End If
        
        ' حساب الراتب الصافي
        Dim grossSalary As Double
        grossSalary = basicSalary + totalAllowances
        netSalary = grossSalary - totalDeductions
        
        ' إدخال البيانات في كشف الرواتب
        With payrollWs
            .Cells(payRow, 1).Value = empID
            .Cells(payRow, 2).Value = empName
            .Cells(payRow, 3).Value = basicSalary
            .Cells(payRow, 3).NumberFormat = "#,##0.00"
            .Cells(payRow, 4).Value = totalAllowances
            .Cells(payRow, 4).NumberFormat = "#,##0.00"
            .Cells(payRow, 5).Value = totalDeductions
            .Cells(payRow, 5).NumberFormat = "#,##0.00"
            .Cells(payRow, 6).Value = grossSalary
            .Cells(payRow, 6).NumberFormat = "#,##0.00"
            .Cells(payRow, 7).Value = netSalary
            .Cells(payRow, 7).NumberFormat = "#,##0.00"
            .Cells(payRow, 8).Value = "Processed"
            
            ' تلوين صف الموظف
            If payRow Mod 2 = 0 Then
                .Range("A" & payRow & ":H" & payRow).Interior.Color = RGB(245, 245, 245)
            End If
        End With
        
        ' الترحيل الآلي للرواتب
        Call AutoPostPayroll(payrollMonth, empName, basicSalary, totalAllowances, totalDeductions, netSalary)
        
        payRow = payRow + 1
    Next i
    
    ' صف الإجمالي
    payRow = payRow + 1
    With payrollWs
        .Cells(payRow, 1).Value = "الإجمالي"
        .Cells(payRow, 1).Font.Bold = True
        .Cells(payRow, 3).Formula = "=SUM(C4:C" & (payRow - 2) & ")"
        .Cells(payRow, 3).NumberFormat = "#,##0.00"
        .Cells(payRow, 3).Font.Bold = True
        .Cells(payRow, 3).Interior.Color = RGB(200, 220, 255)
        .Cells(payRow, 4).Formula = "=SUM(D4:D" & (payRow - 2) & ")"
        .Cells(payRow, 4).NumberFormat = "#,##0.00"
        .Cells(payRow, 4).Font.Bold = True
        .Cells(payRow, 4).Interior.Color = RGB(200, 220, 255)
        .Cells(payRow, 5).Formula = "=SUM(E4:E" & (payRow - 2) & ")"
        .Cells(payRow, 5).NumberFormat = "#,##0.00"
        .Cells(payRow, 5).Font.Bold = True
        .Cells(payRow, 5).Interior.Color = RGB(200, 220, 255)
        .Cells(payRow, 6).Formula = "=SUM(F4:F" & (payRow - 2) & ")"
        .Cells(payRow, 6).NumberFormat = "#,##0.00"
        .Cells(payRow, 6).Font.Bold = True
        .Cells(payRow, 6).Interior.Color = RGB(200, 220, 255)
        .Cells(payRow, 7).Formula = "=SUM(G4:G" & (payRow - 2) & ")"
        .Cells(payRow, 7).NumberFormat = "#,##0.00"
        .Cells(payRow, 7).Font.Bold = True
        .Cells(payRow, 7).Interior.Color = RGB(200, 220, 255)
    End With
    
    payrollWs.Columns("A:H").AutoFit
    payrollWs.Range("A1").Select
    ActiveWindow.FreezePanes = True
    
    ProcessMonthlyPayroll = True
    MsgBox "✅ تم معالجة كشف الرواتب لشهر " & Format(payrollMonth, "mmmm yyyy") & " بنجاح!" & vbCrLf & _
    "عدد الموظفين: " & (payRow - 4), vbInformation
    
    Exit Function
    
ErrorHandler:
    MsgBox "❌ خطأ في معالجة الرواتب: " & Err.Description, vbCritical
    ProcessMonthlyPayroll = False
End Function

' =========================================
' 5. تقرير الرواتب (Payroll Report)
' =========================================

Public Sub GeneratePayrollReport(payrollMonth As Date)
    On Error GoTo ErrorHandler
    
    Dim reportWs As Worksheet
    Dim payrollWs As Worksheet
    Dim sheetExists As Boolean
    Dim row As Integer
    
    ' البحث عن ورقة الرواتب
    On Error Resume Next
    Set payrollWs = ThisWorkbook.Sheets("Monthly_Payroll_" & Format(payrollMonth, "mmyyyy"))
    On Error GoTo ErrorHandler
    
    If payrollWs Is Nothing Then
        MsgBox "❌ لم يتم العثور على كشف الرواتب لهذا الشهر!", vbCritical
        Exit Sub
    End If
    
    ' إنشاء ورقة التقرير
    sheetExists = False
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Sheets
        If sh.Name = "Payroll_Report_" & Format(payrollMonth, "mmyyyy") Then
            sheetExists = True
            Set reportWs = sh
            Exit For
        End If
    Next sh
    
    If Not sheetExists Then
        Set reportWs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        reportWs.Name = "Payroll_Report_" & Format(payrollMonth, "mmyyyy")
    Else
        reportWs.Cells.Clear
    End If
    
    ' العنوان
    With reportWs.Range("A1:D1")
        .Merge
        .Value = "تقرير الرواتب الشهري - " & Format(payrollMonth, "mmmm yyyy")
        .Font.Size = 14
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(0, 102, 204)
        .Font.Color = RGB(255, 255, 255)
    End With
    
    row = 3
    
    With reportWs
        .Cells(row, 1).Value = "البيان"
        .Cells(row, 2).Value = "المبلغ"
        .Cells(row, 1).Font.Bold = True
        .Cells(row, 2).Font.Bold = True
        .Cells(row, 1).Interior.Color = RGB(200, 220, 255)
        .Cells(row, 2).Interior.Color = RGB(200, 220, 255)
    End With
    
    row = row + 1
    
    ' استخراج البيانات من كشف الرواتب
    Dim lastRow As Long
    lastRow = payrollWs.Cells(payrollWs.Rows.Count, 1).End(xlUp).Row
    
    ' إجمالي الرواتب الأساسية
    With reportWs
        .Cells(row, 1).Value = "إجمالي الرواتب الأساسية"
        .Cells(row, 2).Formula = "=SUMIF('Monthly_Payroll_" & Format(payrollMonth, "mmyyyy") & "'!A:A,""<>Employee ID"",'" & _
                                 "Monthly_Payroll_" & Format(payrollMonth, "mmyyyy") & "'!C:C)"
        .Cells(row, 2).NumberFormat = "#,##0.00"
    End With
    row = row + 1
    
    With reportWs
        .Cells(row, 1).Value = "إجمالي البدلات"
        .Cells(row, 2).Formula = "=SUMIF('Monthly_Payroll_" & Format(payrollMonth, "mmyyyy") & "'!A:A,""<>Employee ID"",'" & _
                                 "Monthly_Payroll_" & Format(payrollMonth, "mmyyyy") & "'!D:D)"
        .Cells(row, 2).NumberFormat = "#,##0.00"
    End With
    row = row + 1
    
    With reportWs
        .Cells(row, 1).Value = "إجمالي الخصومات"
        .Cells(row, 2).Formula = "=SUMIF('Monthly_Payroll_" & Format(payrollMonth, "mmyyyy") & "'!A:A,""<>Employee ID"",'" & _
                                 "Monthly_Payroll_" & Format(payrollMonth, "mmyyyy") & "'!E:E)"
        .Cells(row, 2).NumberFormat = "#,##0.00"
    End With
    row = row + 1
    
    With reportWs
        .Cells(row, 1).Value = "إجمالي الأجور الإجمالية"
        .Cells(row, 1).Font.Bold = True
        .Cells(row, 2).Formula = "=SUMIF('Monthly_Payroll_" & Format(payrollMonth, "mmyyyy") & "'!A:A,""<>Employee ID"",'" & _
                                 "Monthly_Payroll_" & Format(payrollMonth, "mmyyyy") & "'!F:F)"
        .Cells(row, 2).NumberFormat = "#,##0.00"
        .Cells(row, 2).Font.Bold = True
        .Cells(row, 2).Interior.Color = RGB(200, 220, 255)
    End With
    row = row + 1
    
    With reportWs
        .Cells(row, 1).Value = "إجمالي الأجور الصافية"
        .Cells(row, 1).Font.Bold = True
        .Cells(row, 2).Formula = "=SUMIF('Monthly_Payroll_" & Format(payrollMonth, "mmyyyy") & "'!A:A,""<>Employee ID"",'" & _
                                 "Monthly_Payroll_" & Format(payrollMonth, "mmyyyy") & "'!G:G)"
        .Cells(row, 2).NumberFormat = "#,##0.00"
        .Cells(row, 2).Font.Bold = True
        .Cells(row, 2).Interior.Color = RGB(144, 238, 144)
    End With
    
    reportWs.Columns("A:D").AutoFit
    
    MsgBox "✅ تم إنشاء تقرير الرواتب بنجاح!", vbInformation
    
    Exit Sub
    
ErrorHandler:
    MsgBox "❌ خطأ في إنشاء التقرير: " & Err.Description, vbCritical
End Sub

' =========================================
' 6. واجهة معالجة الرواتب
' =========================================

Public Sub PayrollInterface()
    On Error GoTo ErrorHandler
    
    Dim response As Integer
    
    response = MsgBox("اختر الخيار المطلوب:" & vbCrLf & vbCrLf & _
    "1 - إضافة موظف جديد" & vbCrLf & _
    "2 - إضافة بدلات" & vbCrLf & _
    "3 - إضافة خصومات" & vbCrLf & _
    "4 - معالجة كشف الرواتب الشهري" & vbCrLf & _
    "5 - عرض تقرير الرواتب", _
    vbYesNoCancel + vbQuestion)
    
    Select Case response
        Case vbYes
            Call ManageEmployees
        Case vbNo
            Call ManagePayroll
        Case vbCancel
            MsgBox "تم الإلغاء", vbInformation
    End Select
    
    Exit Sub
    
ErrorHandler:
    MsgBox "❌ خطأ: " & Err.Description, vbCritical
End Sub

Public Sub ManagePayroll()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim sheetExists As Boolean
    
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
    End If
    
    With ws.Range("A1:H1")
        .Value = Array("Employee ID", "Employee Name", "Basic Salary", "Allowances", "Deductions", "Gross Salary", "Net Salary", "Date")
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
