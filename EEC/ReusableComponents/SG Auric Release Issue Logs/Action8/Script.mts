Option Explicit

Dim bResult, strExpenseTitle, Home, ExpenseReport
Set Home = Home_Page
Set ExpenseReport = ExpenseReport_Page

'Step 1: Create claim with amount more than whats in policy
RunAction "Create Complete Claim (No Tax) [Quick Add]", oneIteration, Parameter("strExpenseData"), bResult

'Step 2: Submit expenses
RunAction "Submit Expense Report (Returns ExpenseTitle) [Expense Report]", oneIteration, Parameter("strEmails"), bResult, strExpenseTitle

'Step 3: Wait for system validation period for 3 minutes
Wait(180)

'Step 4: Search for expense report
Parameter("bResult") = ExpenseReport.ValidateExpenseReportRevise(strExpenseTitle)


Set Home = nothing
Set ExpenseReport = nothing
