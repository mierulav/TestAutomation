Option Explicit

'Declarations
Dim EECLogin, Initiate, Homepage
Set Initiate = Init()

With Datatable
	
'Import Test Data
.ImportSheet Initiate.GetTestDataFile, "Main", "Global"

'Login
Initiate.OpenURL Initiate.GetURL
Set EECLogin = Login_Page() : EECLogin.Login Init.GetUsername, Init.GetPassword
Set Homepage = Home_Page() 

Dim i, j, x, bResult, bResult2
For i = 1 To .GetRowCount
	.GetSheet("Global").SetCurrentRow(i)
	If .Value("ToTest", "Global") = "Y" Then
		.SetCurrentRow(i)		
		.ImportSheet Initiate.GetTestDataFile, .Value("Tab", "Global"), "Local"
		
		Select Case .Value("Tab", "Global")
		
			Case "1" 'Save draft expense report
				.GetSheet("Local").SetCurrentRow(1)
				ReDim arrVal(.GetSheet("Local").GetParameterCount-1)
				For x = 0 To .GetSheet("Local").GetParameterCount-1
					arrVal(x) = .GetSheet("Local").GetParameter(x+1)
				Next
				'Run action 	
				RunAction "Create Claim Item [Claim Item]", oneIteration, Join(arrVal, ","), bResult
				RunAction "Save Draft Expense Report [Expense Report]", oneIteration, Init.GetUsername, bResult
				
			Case "2" 'Withdraw expense report
				.GetSheet("Local").SetCurrentRow(1)
				ReDim arrVal(.GetSheet("Local").GetParameterCount-1)
				For x = 0 To .GetSheet("Local").GetParameterCount-1
					arrVal(x) = .GetSheet("Local").GetParameter(x+1)
				Next
				'Run action 
				RunAction "Create Claim Item [Claim Item]", oneIteration, Join(arrVal, ","), bResult
				RunAction "Submit Expense Report (Returns ExpenseTitle) [Expense Report]", oneIteration, Init.GetUsername, bResult2, strTitle
				RunAction "Withdraw Expense Report [Expense Report]", oneIteration, strTitle, bResult3
				
				
			Case "3" 'Cancel expense report
				.GetSheet("Local").SetCurrentRow(1)
				Dim strTitle, bResult3, bResult4, bResult5
				ReDim arrVal(.GetSheet("Local").GetParameterCount-1)
				For x = 0 To .GetSheet("Local").GetParameterCount-1
					arrVal(x) = .GetSheet("Local").GetParameter(x+1)
				Next
				
				'Create and submit expense	
				RunAction "Create Claim Item [Claim Item]", oneIteration, Join(arrVal, ","), bResult
				RunAction "Submit Expense Report (Returns ExpenseTitle) [Expense Report]", oneIteration, Init.GetUsername, bResult2, strTitle

				'Cancel submitted
				If bResult and bResult2 Then
					RunAction "Cancel Expense Report [Expense Report]", oneIteration, strTitle, bResult4
				End If
				
				'Create, submit and withdraw expense
				RunAction "Create Claim Item [Claim Item]", oneIteration, Join(arrVal, ","), bResult
				RunAction "Submit Expense Report (Returns ExpenseTitle) [Expense Report]", oneIteration, Init.GetUsername, bResult2, strTitle
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
					
		End Select
		
		
	End If
		
Next

End  With
