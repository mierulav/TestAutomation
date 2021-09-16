Option Explicit

'Declarations
Dim Initiate : Set Initiate = Init()
Dim EECLogin : Set EECLogin = Login_Page()
Dim i, j, k
Dim bResult, strVal

'Login
Initiate.OpenURL Initiate.GetURL
EECLogin.Login Initiate.GetUsername, Initiate.GetPassword

'Get test data
Datatable.ImportSheet Initiate.GetTestDataFile, "Main", "Global"

For i = 1 To Datatable.GetSheet("Global").GetRowCount
	Datatable.GetSheet("Global").SetCurrentRow(i)
	If Datatable.Value("ToTest", "Global") = "Y" Then
		Datatable.ImportSheet Initiate.GetTestDataFile, Datatable.Value("Tab", "Global"), "Local"
		Select Case Datatable.Value("Tab", "Global")
			Case "1"
				Datatable.GetSheet("Local").SetCurrentRow(1)
				ReDim arrVal(Datatable.GetSheet("Local").GetParameterCount-1)
				For j = 0 To Datatable.GetSheet("Local").GetParameterCount-1
					arrVal(j) = Datatable.GetSheet("Local").GetParameter(j+1)
				Next
				
				'Run create action
				RunAction "Create Claim Item [Claim Item]", oneIteration, Join(arrVal, ","), bResult
				
			Case "2"
				Datatable.GetSheet("Local").SetCurrentRow(1)
				ReDim arrVal(Datatable.GetSheet("Local").GetParameterCount-1)
				For j = 0 To Datatable.GetSheet("Local").GetParameterCount-1
					arrVal(j) = Datatable.GetSheet("Local").GetParameter(j+1)
				Next
				
				'Run Edit action
				RunAction "Edit Claim Item [Claim Item]", oneIteration, Join(arrVal, ","), bResult
				
			Case "3"
				Datatable.GetSheet("Local").SetCurrentRow(1)
				ReDim arrVal(Datatable.GetSheet("Local").GetParameterCount-1)
				For j = 0 To Datatable.GetSheet("Local").GetParameterCount-1
					arrVal(j) = Datatable.GetSheet("Local").GetParameter(j+1)
				Next
				
				'Run Cancel action
				RunAction "Cancel Claim Item [Claim Item]", oneIteration, Join(arrVal, ","), bResult
				
			Case "4"
				'Run Delete action
				RunAction "Remove Claim Item [Claim Item]", oneIteration, "Accommodation", bResult
				
		End Select
		
		If bResult Then
			Reporter.ReportEvent micPass, "Claim Functionality of " & Datatable.Value("TestName", "Global") & " is working as expected", "Pass"
		Else
			Reporter.ReportEvent micFail, "Claim Functionality of " & Datatable.Value("TestName", "Global") & " is working as expected", "Fail"	
		End If
	End If
Next
