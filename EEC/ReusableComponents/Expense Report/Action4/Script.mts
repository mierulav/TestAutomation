Option Explicit

Dim Home : Set Home = Home_Page()
Dim Expense : Set Expense = ExpenseReport_Page()
Dim Timestamp : TimeStamp = GetStringDate
Dim strTitle : strTitle = "AT" & TimeStamp
Dim strEmails : strEmails = Parameter("strEmails")
Parameter("bResult") = True

'Mavigate to My Expense Report page
Home.NavigateToMyExpenseReport

'The My expense report items listing page is displayed.(Step 16).
If Expense.ValidateMyExpenseReportPage Then
	Reporter.ReportEvent micPass, "My Expense Report page is displayed", "Pass"
Else
	Reporter.ReportEvent micFail, "My Expense Report page is displayed", "Fail"
End If

'Create expense report
Expense.CreateExpenseReport

'Change expense report title
Expense.SetExpenseReportTitle strTitle

'Validate that user is able to change the title of the expense report.(Step 18)
If Expense.GetExpenseTitle = strTitle Then
	Reporter.ReportEvent micPass, "Expense Report title changed successfully", "Pass"
Else
	Reporter.ReportEvent micFail, "Expense Report title changed successfully", "Fail"
End If

'Fill emails to cc separated by semi colon 
Expense.SetExpenseReportEmailCC strEmails

'Fill personal emails separated by semi colon
Expense.SetExpenseReportPersonalEmail strEmails

'Fill in remarks
Expense.SetExpenseReportRemarks strTitle

'Tick on the self-certification checkbox.
Expense.SetExpenseReportReceiptCertified

'Validate that the status of the expense report is Draft
Expense.SaveDraftExpenseReport
If Expense.ValidateExpenseReportDraft(strTitle) Then
	Reporter.ReportEvent micPass, "Save draft Expense Report", "Pass"
Else
	Reporter.ReportEvent micFail, "Save draft Expense Report", "Fail"
	Parameter("bResult") = False
	ExitAction
End If

'Search expense report
Expense.EditSpecificExpense(strTitle)

Wait(3)

'Validate Audit Log
If Expense.ValidateAuditLogDraftStatus Then
	Reporter.ReportEvent micPass, "Save as draft is populated under Action column of Audit Log", "Pass"
	Parameter("bResult") = True
Else
	Reporter.ReportEvent micPass, "Save as draft is populated under Action column of Audit Log", "Fail"
	Parameter("bResult") = False
End If
	


