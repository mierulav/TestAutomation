Option Explicit

Dim Initiate : Set Initiate = Init()
Dim Home : Set Home = Home_Page()
Dim EECLogin : Set EECLogin = Login_Page()

'Declaration
Dim bResult, bResult2, strTitle, strReason, i, x

'Import test data
Datatable.ImportSheet Initiate.GetTestDataFile, "Main", "Global"

For i = 1 To Datatable.GetSheet("Global").GetRowCount
	Datatable.GetSheet("Global").SetCurrentRow(i)
	
	'Login
	Initiate.OpenURL Initiate.GetURL
	EECLogin.Login Initiate.GetUsername, Initiate.GetPassword
	
	If UCase(Datatable.Value("ToTest", "Global")) = "Y" and UCase(Datatable.Value("TestName", "Global")) = "REJECT" Then
		Datatable.ImportSheet Initiate.GetTestDataFile, Datatable.Value("Tab", "Global"), "Local"
		Datatable.GetSheet("Local").SetCurrentRow(1)
		ReDim arrVal(Datatable.GetSheet("Local").GetParameterCount-1)
		For x = 0 To Datatable.GetSheet("Local").GetParameterCount-1
			arrVal(x) = Datatable.GetSheet("Local").GetParameter(x+1)
		Next
		
		'Create and submit expense	
		RunAction "Create Claim Item [Claim Item]", oneIteration, Join(arrVal, ","), bResult
		RunAction "Submit Expense Report (Returns ExpenseTitle) [Expense Report]", oneIteration, Datatable.Value("SupervisorEmail", "Global"), bResult, strTitle
		
		'Logout and login using supervisor's
		Home.Logoff
		Wait(2)
		EECLogin.Login Datatable.Value("SupervisorEmail", "Global"), Datatable.Value("SupervisorPassword", "Global")
		
		'Reject
		RunAction "Reject Expense Report [My Team]", oneIteration, strTitle, "Reject", bResult
		
		'Logout
		Home.Logoff
		Wait(2)
		'Validate uses status
		EECLogin.Login Initiate.GetUsername, Initiate.GetPassword
		RunAction "Validate Expense Report Status [Expense Report]", oneIteration, strTitle, "reject", bResult2
		
		'Logout
		Home.Logoff
		
	End If
	
	If UCase(Datatable.Value("ToTest", "Global") = "Y") and UCase(Datatable.Value("TestName", "Global")) = "REVISE" Then
		Datatable.ImportSheet Initiate.GetTestDataFile, Datatable.Value("Tab", "Global"), "Local"
		Datatable.GetSheet("Local").SetCurrentRow(1)
		ReDim arrVal(Datatable.GetSheet("Local").GetParameterCount-1)
		For x = 0 To Datatable.GetSheet("Local").GetParameterCount-1
			arrVal(x) = Datatable.GetSheet("Local").GetParameter(x+1)
		Next
		
		'Create and submit expense	
		RunAction "Create Claim Item [Claim Item]", oneIteration, Join(arrVal, ","), bResult
		RunAction "Submit Expense Report (Returns ExpenseTitle) [Expense Report]", oneIteration, Datatable.Value("SupervisorEmail", "Global"), bResult, strTitle
		
		'Logout and login using supervisor's
		Home.Logoff
		Wait(2)
		EECLogin.Login Datatable.Value("SupervisorEmail", "Global"), Datatable.Value("SupervisorPassword", "Global")
		
		'Revise
		RunAction "Revise Expense Report [My Team]", oneIteration, strTitle, "Revise", bResult
		
		'Logout
		Home.Logoff
		Wait(2)
		'Validate uses status
		EECLogin.Login Initiate.GetUsername, Initiate.GetPassword
		RunAction "Validate Expense Report Status [Expense Report]", oneIteration, strTitle, "revise", bResult2
		
		'Logout
		Home.Logoff
		
	End If
	
	If bResult and bResult2 Then
		Reporter.ReportEvent micPass, "Rejects/Revise expense report", "Pass"
	Else
		Reporter.ReportEvent micFail, "Rejects/Revise expense report", "Fail"
	End If
	
Next





