Option Explicit

Dim Home : Set Home = Home_Page
Dim ExpenseReport : Set ExpenseReport = ExpenseReport_Page
Dim arrTestData : arrTestData = Split(Parameter("strTestData"), ",")
Dim strSearchCriteria : strSearchCriteria =  arrTestData(0)
Dim strVal	: strVal = arrTestData(1) 

'Step 1: Navigate to My Expense Report
Home.NavigateToMyExpenseReport

'Step 2: Search expense report
Select Case Lcase(strSearchCriteria)
	
	Case "submission date"
	Parameter("bResult") = ExpenseReport.ValidateSearchBySubmissionDate(strVal)
	
	Case "reference no"
	Parameter("bResult") = ExpenseReportExpenseReport.ValidateSearchByReferenceNo(strVal)

	Case "title"
	Parameter("bResult") = ExpenseReport.ValidateSearchByTitle(strVal)
	
	Case Else
	Parameter("bResult") = False
	ExitAction
	
End Select

