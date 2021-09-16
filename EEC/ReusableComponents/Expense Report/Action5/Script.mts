Option Explicit

Dim Homepage : Set Homepage = Home_Page()
Dim ExpenseReport : Set ExpenseReport = ExpenseReport_Page()
Dim strStatus : strStatus =  Parameter("strStatus")
Dim strRefNo

'Navigate to My expense report
Homepage.NavigateToMyExpenseReport

'Search for expense report
if ExpenseReport.SearchExpenseReport(strStatus) = False Then 
	Parameter("bResult") = False
	ExitAction
End If 

'Click Edit of the expense report
ExpenseReport.EditExpenseReport

'Store searched expense report's RefNo
strRefNo = ExpenseReport.GetReferenceNumber

'Withdraw expense report
ExpenseReport.WithdrawExpenseReport

'Search for withdrawn expense report
If ExpenseReport.ValidateExpenseReportWithdrawn(strRefNo) Then
	Reporter.ReportEvent micPass, "Validate Expense Report is in Withdrawn status from Expense Report list table", "Pass"
Else
	Reporter.ReportEvent micPass, "Validate Expense Report is in Withdrawn status from Expense Report list table", "Pass"
	Parameter("bResult") =  False
	ExitAction
End If

'Click Edit of the expense report
'ExpenseReport.EditExpenseReport
ExpenseReport.EditSpecificExpense(strStatus)

Wait(3)

'Validate Audit Log
If ExpenseReport.ValidateAuditLogWithdrawnStatus Then
	Reporter.ReportEvent micPass, "Withdraw is populated under Action column of Audit Log", "Pass"
	Parameter("bResult") =  True
Else
	Reporter.ReportEvent micFail, "Withdraw is populated under Action column of Audit Log", "Fail"
	Parameter("bResult") =  False
End If
