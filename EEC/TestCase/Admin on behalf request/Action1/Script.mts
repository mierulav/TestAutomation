Option Explicit

'Declarations
Dim strAdminUsername, strAdminPassword, strEmployeeName, strAdminName
Dim EECLogin, Initiate, Homepage
Set Initiate = Init()

With Datatable
	
'Import Test Data
.ImportSheet Initiate.GetTestDataFile, "Main", "Global"
Dim i, j, x, bResult, bResult2
For i = 1 To .GetRowCount
	.GetSheet("Global").SetCurrentRow(i)
	If .Value("ToTest", "Global") = "Y" Then
		.SetCurrentRow(i)
		strAdminUsername = .Value("AdminEmail", "Global")
		strAdminPassword = .Value("AdminPassword", "Global")
		strEmployeeName = .Value("OnBehalfOf", "Global")
		strAdminName = .Value("AdminName", "Global")
		
		'Login as Personal Admin
		If i = 1 Then
			Initiate.OpenURL Initiate.GetURL
			Set EECLogin = Login_Page() : EECLogin.Login strAdminUsername, strAdminPassword
			Set Homepage = Home_Page() 
			
			'Validate impersonation
			If Homepage.ImpersonateEmployee(strEmployeeName) And Homepage.ValidateImpersonatingViewableModules Then
				Reporter.ReportEvent micPass, "Employee is found and successfully impersonated", "Pass"
			Else
				Reporter.ReportEvent micFail, "Employee is found and successfully impersonated", "Fail"
			End If
		End If
		
		.ImportSheet Initiate.GetTestDataFile, .Value("Tab", "Global"), "Local"
		
		Select Case .Value("Tab", "Global")
		
			Case "1" 'Create calim item test case
				.GetSheet("Local").SetCurrentRow(1)
				ReDim arrVal(.GetSheet("Local").GetParameterCount-1)
				For x = 0 To .GetSheet("Local").GetParameterCount-1
					arrVal(x) = .GetSheet("Local").GetParameter(x+1)
				Next
				'Run action 	
				RunAction "Create Claim Item [Claim Item]", oneIteration, Join(arrVal, ","), bResult
				
			Case "2" 'Edit claim item
				.GetSheet("Local").SetCurrentRow(1)
				ReDim arrVal(.GetSheet("Local").GetParameterCount-1)
				For x = 0 To .GetSheet("Local").GetParameterCount-1
					arrVal(x) = .GetSheet("Local").GetParameter(x+1)
				Next
				'Run action 
				RunAction "Edit Claim Item [Claim Item]", oneIteration, Join(arrVal, ","), bResult

				
			Case "3" 'Delete claim item
				.GetSheet("Local").SetCurrentRow(1)
				'Run action
				RunAction "Remove Claim Item [Claim Item]", oneIteration, .Value("ClaimCategory", "Local"), bResult
				
			Case "4" 'Cancel expense report
				.GetSheet("Local").SetCurrentRow(1)
				Dim strTitle, bResult3, bResult4, bResult5
				ReDim arrVal(.GetSheet("Local").GetParameterCount-1)
				For x = 0 To .GetSheet("Local").GetParameterCount-1
					arrVal(x) = .GetSheet("Local").GetParameter(x+1)
				Next
				
				'Create and submit expense	
				RunAction "Create Claim Item [Claim Item]", oneIteration, Join(arrVal, ","), bResult
				RunAction "Submit Expense Report (Returns ExpenseTitle) [Expense Report]", oneIteration, Datatable.Value("AdminEmail", "Global"), bResult2, strTitle

				'Cancel submitted
				If bResult and bResult2 Then
					RunAction "Cancel Expense Report [Expense Report]", oneIteration, strTitle, bResult4
				End If
				
				'Create, submit and withdraw expense
				RunAction "Create Claim Item [Claim Item]", oneIteration, Join(arrVal, ","), bResult
				RunAction "Submit Expense Report (Returns ExpenseTitle) [Expense Report]", oneIteration, Datatable.Value("AdminEmail", "Global"), bResult2, strTitle
				RunAction "Withdraw Expense Report [Expense Report]", oneIteration, strTitle, bResult3
				
				'Cancel withdrawn
				If bResult and bResult2 and bResult3 Then
					RunAction "Cancel Expense Report [Expense Report]", oneIteration, strTitle, bResult5
				End If
				
				'Final validation
				If bResult4 and bResult5 Then
					bResult = True
				Else
					bResult = False
				End If
					
			Case "5" 'Withdraw expense report
				.GetSheet("Local").SetCurrentRow(1)
				'Run action
			
			Case "6" 'Create and submit expense report
				.GetSheet("Local").SetCurrentRow(1)
				ReDim arrVal(.GetSheet("Local").GetParameterCount-1)
				For x = 0 To .GetSheet("Local").GetParameterCount-1
					arrVal(x) = .GetSheet("Local").GetParameter(x+1)
				Next
				'Run action
				RunAction "Create Claim Item [Claim Item]", oneIteration, Join(arrVal, ","), bResult
				RunAction "Submit Expense Report (Returns ExpenseTitle) [Expense Report]", oneIteration, Datatable.Value("AdminEmail", "Global"), bResult2, strTitle

				If bResult and bResult2 Then
					bResult = True
				Else
					bResult = False
				End If
				
		End Select
		
		If i = 1 Then
			If Homepage.UnmaskEmployee(strAdminUsername) Then
				Reporter.ReportEvent micPass, "Unmasked employee successfully", "Pass"
			Else
				Reporter.ReportEvent micFail, "Fail to unmask employee", "Fail"
			End If
			
			'impersonate again
			 Homepage.ImpersonateEmployee(strEmployeeName)
		End If
		
		If bResult Then
			Reporter.ReportEvent micPass, .Value("TestName", "Global"), "Pass"
		Else
			Reporter.ReportEvent micFail, .Value("TestName", "Global"), "Fail"
		End If
		
'		Homepage.Logoff
'		Initiate.CloseAllBrowser
		
	End If
		
Next

End  With
