Option Explicit

Dim Homepage : Set Homepage = Home_Page()
Dim ExpenseReport : Set ExpenseReport = ExpenseReport_Page()
Dim strSearch, strRefNo, strStatus
strSearch = Parameter("strSearch")
strStatus = Parameter("strSearch")

'Navigate to My expense report
Homepage.NavigateToMyExpenseReport

'Search for expense report
if ExpenseReport.SearchExpenseReport(strSearch) = False Then 
	Reporter.ReportEvent micDone, "No status " & strStatus & " found", "Done" 
	Parameter("bResult") = False
	ExitAction
End If

'Edit the first item from the list
'ExpenseReport.EditExpenseReport
ExpenseReport.EditSpecificExpense(strSearch)

'Get EEC RefNO
strRefNo = ExpenseReport.GetReferenceNumber

'Click on delete button
ExpenseReport.DeleteExpenseReport

'Validate the expense report is cancelled
If ExpenseReport.ValidateExpenseReportCancelled(strRefNo) Then
	Reporter.ReportEvent micPass, "Cancel Expense Report", "Pass"
Else
	Reporter.ReportEvent micFail, "Cancel Expense Report", "Fail"
	Parameter("bResult") = False
	ExitAction
End If 

'To validate Audit Log
'ExpenseReport.EditExpenseReport
ExpenseReport.EditSpecificExpense(strSearch)

Wait(3)

If ExpenseReport.ValidateAuditLogCancelledStatus Then
	Reporter.ReportEvent micPass, "Cancel is populated under Action column of Audit Log", "Pass"
	Parameter("bResult") = True
Else
	Reporter.ReportEvent micFail, "Cancel is populated under Action column of Audit Log", "Fail"
	Parameter("bResult") = False
End If

