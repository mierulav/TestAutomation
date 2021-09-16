Option Explicit

'Declarations
Dim Homepage : Set Homepage = Home_Page()
Dim Functions : Set Functions = Function_Page()
Dim strEmployeeName, strEmployeeID, strEmployeeEmail, strView
strEmployeeName = Parameter("strEmployeeName")
strEmployeeID = Parameter("strEmployeeID")
strEmployeeEmail = Parameter("strEmployeeEmail")
strView = Parameter("RoleView")

'Navigate to Function
Homepage.NavigateToFunction

'Validate Function layout - 'HR' or 'finance' view
Functions.ValidateFunctionView(strView)

'Validate search by employee name within country or business unit
If Functions.ValidateSearchByEmployeeName(strEmployeeName) Then
	Reporter.ReportEvent micPass, "Validate Search function using Employee Name: " & strEmployeeName, "Pass"
Else
	Reporter.ReportEvent micFail, "Validate Search function using Employee Name: " & strEmployeeName, "Fail"
End if

'Validate search by employee id within country or business unit
If Functions.ValidateSearchByEmployeeID(strEmployeeID, strEmployeeName) Then
	Reporter.ReportEvent micPass, "Validate Search function using Employee ID: " & strEmployeeID, "Pass"
Else
	Reporter.ReportEvent micFail, "Validate Search function using Employee ID: " & strEmployeeID, "Fail"
End If

'Validate search by employee email within country or business unit
If Functions.ValidateSearchByEmailAddress(strEmployeeEmail, strEmployeeName) Then
	Reporter.ReportEvent micPass, "Validate Search function using Employee Email: " & strEmployeeEmail, "Pass"
Else
	Reporter.ReportEvent micFail, "Validate Search function using Employee Email: " & strEmployeeEmail, "Fail"
End If

'Search for Processed status expense and Download expense report
Functions.SearchExpenseReport("Processed")
Functions.DownloadExpenseReport

'Validate download page is opened and allow user to download
If Functions.ValidateExpenseReportDownloaded Then
	Reporter.ReportEvent micPass, "Verify user is allow to download expense report", "Pass"
Else
	Reporter.ReportEvent micPass, "Verify user is allow to download expense report", "Pass"
End If 
	
'View expense report detail
Functions.ViewExpenseReportDetails

'Validate all fields disabled
If Functions.ValidateAllExpenseReportFieldsDisabled Then
	Reporter.ReportEvent micPass, "Verify User is not able to make any changes to the details", "Pass"
Else
	Reporter.ReportEvent micFail, "Verify User is not able to make any changes to the details", "Fail"
End If





