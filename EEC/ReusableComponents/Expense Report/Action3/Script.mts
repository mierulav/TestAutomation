Option Explicit

Dim Home : Set Home = Home_Page()
Dim ExpenseReport : Set ExpenseReport = ExpenseReport_Page()
Parameter("bResult") = True

'Navigate to Expense Report 
Home.NavigateToMyExpenseReport

'Search for expense report by title and return reference no
 if ExpenseReport.SearchExpenseReport(Parameter("strTitle")) Then
 	Parameter("strRefNo") = ExpenseReport.GetCellDataReferenceNo
 Else
 	Parameter("bResult") = False
 End If



