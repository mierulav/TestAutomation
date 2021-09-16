Option Explicit

Dim Home : Set Home = Home_Page()
Dim Expense : Set Expense = ExpenseReport_Page()
Dim Timestamp : TimeStamp = GetStringDate
Dim strTitle : strTitle = "AT" & TimeStamp
Dim strEmails : strEmails = Parameter("strEmails")
Parameter("bResult") = True

'Step 1: Navigate to My Expense Report
Home.NavigateToMyExpenseReport

'Step 2: Click on Create Expense Report
Expense.CreateExpenseReport

'Step 3: Modify Expense Report title
Expense.SetExpenseReportTitle strTitle

'Validate Step 3 - Validate that user is able to change the title of the expense report.
If Expense.GetExpenseTitle = strTitle Then
	Reporter.ReportEvent micPass, "Expense Report title changed successfully", "Pass"
Else
	Reporter.ReportEvent micFail, "Expense Report title changed successfully", "Fail"
End If

'Step 4: Fill emails to cc separated by semi colon 
Expense.SetExpenseReportEmailCC strEmails

'Step 5: Fill personal emails separated by semi colon
Expense.SetExpenseReportPersonalEmail strEmails

'Step 6: Fill in remarks
Expense.SetExpenseReportRemarks strTitle

'Step 7: Tick on the self-certification checkbox.
Expense.SetExpenseReportReceiptCertified

'Step 8: Submit Expense Report
If  Expense.SubmitExpenseReport and Expense.ValidateExpenseReportSubmitted(strTitle) Then
	Reporter.ReportEvent micPass, "Submit Expense Report", "Pass"
Else
	Reporter.ReportEvent micFail, "Submit Expense Report", "Fail"
	Parameter("bResult") = False
	ExitAction
End If

'Click Edit of the expense report
'Expense.EditExpenseReport
Expense.EditSpecificExpense(strTitle)

Wait(3)

'Validate Audit Log
If Expense.ValidateAuditLogSubmittedStatus Then
	Reporter.ReportEvent micPass, "Submit is populated under Action column of Audit Log", "Pass"
	Parameter("bResult") =  True
Else
	Reporter.ReportEvent micFail, "Submit is populated under Action column of Audit Log", "Fail"
	Parameter("bResult") =  False
End If
