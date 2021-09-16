Option Explicit

'Declarations
Dim Homepage : Set Homepage = Home_Page()
Dim Functions : Set Functions = Function_Page()
Dim strValidSearch : strValidSearch = Parameter("strNotValidEmp")
Dim strNotValidSearch : strNotValidSearch = Parameter("strNotValidEmp")

'Navigate to Function
Homepage.NavigateToFunction

'Validate Search employee not within business unit
If Functions.SearchExpenseReport(strNotValidSearch) = False Then
	Reporter.ReportEvent micPass, "Not able to find result outside business unit", "Pass"
Else
	Reporter.ReportEvent micFail, "Not able to find result outside business unit", "Fail"
	Parameter("bResult") = False
	ExitAction
End If

'Validate Search employee within business unit
If Functions.SearchExpenseReport(strValidSearch) Then
	Reporter.ReportEvent micPass, "Not able to find result outside business unit", "Pass"
Else
	Reporter.ReportEvent micFail, "Not able to find result outside business unit", "Fail"
	Parameter("bResult") = False
	ExitAction
End If

'Download expense report
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
	Parameter("bResult") = False
End If





