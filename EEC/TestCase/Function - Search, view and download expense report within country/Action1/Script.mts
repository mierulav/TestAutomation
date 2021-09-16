Option Explicit

'Declarations
Dim i, j, arrVal, bResult, Initiate, EECLogin, strHREmail, strHRPassword

'Login
Set Initiate = Init()
Initiate.OpenURL Initiate.GetURL

'Get Test data
With Datatable
	.ImportSheet Initiate.GetTestDataCountryFile, "Main", "Global"
	strHREmail = .Value("HREmail", "Global")
	strHRPassword = .Value("HRPassword", "Global")

	Set EECLogin = Login_Page() : EECLogin.Login strHREmail, strHRPassword
	
	.ImportSheet Initiate.GetTestDataCountryFile, Initiate.GetCompanyCode, "Local"
	
	For i = 1 To .GetSheet("Local").GetRowCount
	.GetSheet("Local").SetCurrentRow(i)
		If .Value("ToTest", "Local") = "Y" Then
			RunAction "Search, View and Download Employee Expense Report [Function]", oneIteration, .Value("EmployeeID", "Local"), _
			.Value("EmployeeName", "Local"), .Value("EmployeeEmail", "Local"), "hr"
		End If
	Next
	
End With
