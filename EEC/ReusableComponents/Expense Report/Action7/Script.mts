Option Explicit

Dim Home : Set Home = Home_Page()
Dim Expense : Set Expense = ExpenseReport_Page()
Dim strTitle : strTitle = Parameter("strRefNo")
Dim strStatus : strStatus = Parameter("strStatus")
Dim bRes
Parameter("bResult") = True

'Step 1: Navigate to My Expense Report
Home.NavigateToMyExpenseReport

'Step 2: Validate Expense Report status in Expense Report List page
Expense.ValidateExpenseReportRevise(strTitle)

'Click Edit of the expense report
Expense.EditSpecificExpense(strTitle)

Wait(3)

'Validate Audit Log
Select Case LCase(strStatus)
	
	Case "revise"
		bRes = Expense.ValidateAuditLogRevisedStatus
	
	Case "reject"
		bRes = Expense.ValidateAuditLogRejectedStatus
		
	Case "submit"
		bRes = Expense.ValidateAuditLogSubmittedStatus
End Select


If bRes Then
	Reporter.ReportEvent micPass, "Status is populated under Action column of Audit Log", "Pass"
	Parameter("bResult") =  True
Else
	Reporter.ReportEvent micFail, "Status is populated under Action column of Audit Log", "Fail"
	Parameter("bResult") =  False
End If
