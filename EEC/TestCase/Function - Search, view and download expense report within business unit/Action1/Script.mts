Option Explicit

'Declarations
Dim i, j, arrVal, bResult, Initiate, EECLogin, strFinanceEmail, strFinancePassword

'Login
Set Initiate = Init()
Initiate.OpenURL Initiate.GetURL

'Get Test data
With Datatable
	.ImportSheet Initiate.GetTestDataCountryFile, "Main", "Global"
	strFinanceEmail = .Value("FinanceEmail", "Global")
	strFinancePassword = .Value("FinancePassword", "Global")

	Set EECLogin = Login_Page() : EECLogin.Login strFinanceEmail, strFinancePassword
	
	.ImportSheet Initiate.GetTestDataCountryFile, Initiate.GetCompanyCode, "Local"
	
	For i = 1 To .GetSheet("Local").GetRowCount
	.GetSheet("Local").SetCurrentRow(i)
		If .Value("ToTest", "Local") = "Y" Then
			RunAction "Search, View and Download Employee Expense Report [Function]", oneIteration, .Value("EmployeeID", "Local"), _
			.Value("EmployeeName", "Local"), .Value("EmployeeEmail", "Local"), "finance"
		End If
	Next
	
End With
