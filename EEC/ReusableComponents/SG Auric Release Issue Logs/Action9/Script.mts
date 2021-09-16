Option Explicit

Dim bResult, strExpenseTitle(2), Home, ExpenseReport, i, x
Set Home = Home_Page
Set ExpenseReport = ExpenseReport_Page

For i = 0 To Ubound(strExpenseTitle)
	'Step 1: Create claim with amount more than whats in policy
	RunAction "Create Complete Claim (No Tax) [Quick Add]", oneIteration, Parameter("strExpenseData"), bResult

	'Step 2: Submit expenses
	RunAction "Submit Expense Report (Returns ExpenseTitle) [Expense Report]", oneIteration, Parameter("strEmails"), bResult, strExpenseTitle(i)
Next

'Step 7: Wait for system validation period for 3 minutes
Wait(180)

'Step 8: Search for expense report
For x = 0 To Ubound(strExpenseTitle)
	Parameter("bResult") = ExpenseReport.ValidateExpenseReportPending(strExpenseTitle(x))
	If Parameter("bResult") = False Then
		ExitAction
	End If
Next

Set Home = nothing
Set ExpenseReport = nothing
