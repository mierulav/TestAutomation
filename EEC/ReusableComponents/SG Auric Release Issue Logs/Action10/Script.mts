Option Explicit

Dim Home, ExpenseReport, bResult, arrTestData, strAvail, strSearch
Set Home = Home_Page
Set ExpenseReport = ExpenseReport_Page
arrTestData = Split(Parameter("strTestData"), ",")
strAvail = arrTestData(1)
strSearch = arrTestData(0)

'Step 1: Navigate to My Expense Report page
Home.NavigateToMyExpenseReport

'Step 2: Search for "Ready to process" status
Select Case Lcase(strSearch)
	
	Case "submitted"
		If ExpenseReport.ValidateExpenseReportSubmitted(strSearch) Then
			ExpenseReport.EditExpenseReport
		Else
			Parameter("bResult") = False
			ExitAction
		End If
	
	Case "pending"
		If ExpenseReport.ValidateExpenseReportPending(strSearch) Then
			ExpenseReport.EditExpenseReport
		Else
			Parameter("bResult") = False
			ExitAction
		End If
		
	Case "revise"
		If ExpenseReport.ValidateExpenseReportRevise(strSearch) Then
			ExpenseReport.EditExpenseReport
		Else
			Parameter("bResult") = False
			ExitAction
		End If
	
	Case "ready to process"
		If ExpenseReport.ValidateExpenseReportReadyToProcess(strSearch) Then
			ExpenseReport.EditExpenseReport
		Else
			Parameter("bResult") = False
			ExitAction
		End If
	
	Case Else
		Parameter("bResult") = False
		ExitAction
		
End Select

'Step 3: Validate Remarks field
If ExpenseReport.ValidateRemarksField = strAvail Then
	Parameter("bResult") = True
Else
	Parameter("bResult") = False	
End If


