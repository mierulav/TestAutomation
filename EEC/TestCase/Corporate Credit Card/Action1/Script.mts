Option Explicit

'Declarations
Dim Initiate : Set Initiate = Init()
Dim EECLogin : Set EECLogin = Login_Page()
Dim bResult, bResult2, i, j, x

'Login
Initiate.OpenURL Initiate.GetURL
EECLogin.Login Initiate.GetUsername, Initiate.GetPassword

'Run action
Datatable.ImportSheet Initiate.GetTestDataFile, "Sheet1", "Global"

For i = 1 To Datatable.GetSheet("Global").GetRowCount
	Datatable.GetSheet("Global").SetCurrentRow(i)
		
		'CC Save Draft
		If Datatable.Value("Tab", "Global") = "1" and Datatable.Value("ToTest", "Global") = "Y" Then
			For x = 1 To CInt(Datatable.Value("NumberOfClaim", "Global"))
				RunAction "Select and Save Credit Card Expense [Corporate Credit Card]", oneIteration, Datatable.Value("EECRefNo", "Global"), Datatable.Value("ExpenseCategory", "Global"), x
			Next
			RunAction "Save Draft Expense Report [Corporate Credit Card]", oneIteration, Datatable.Value("EECRefNo", "Global"), bResult
			If bResult Then
				bResult2 = bResult
			End If
		End If
		
		'CC Valid
		If Datatable.Value("Tab", "Global") = "2" and Datatable.Value("ToTest", "Global") = "Y" Then
			For x = 1 To CInt(Datatable.Value("NumberOfClaim", "Global"))
				RunAction "Select and Save Credit Card Expense [Corporate Credit Card]", oneIteration, Datatable.Value("EECRefNo", "Global"), Datatable.Value("ExpenseCategory", "Global"), x
			Next	
			RunAction "Submit Expense Report [Corporate Credit Card]", oneIteration, Datatable.Value("EECRefNo", "Global"), bResult, bResult2
		End If
		
		'CC Reversal
		If Datatable.Value("Tab", "Global") = "3" and Datatable.Value("ToTest", "Global") = "Y" Then
			For x = 1 To CInt(Datatable.Value("NumberOfClaim", "Global"))
				RunAction "Select and Save Credit Card Expense [Corporate Credit Card]", oneIteration, Datatable.Value("EECRefNo", "Global"), Datatable.Value("ExpenseCategory", "Global"), x
			Next	
			RunAction "Submit Expense Report [Corporate Credit Card]", oneIteration, Datatable.Value("EECRefNo", "Global"), bResult, bResult2
		End If
		
		''CC Dispute
		If Datatable.Value("Tab", "Global") = "4" and Datatable.Value("ToTest", "Global") = "Y" Then
			For x = 1 To CInt(Datatable.Value("NumberOfClaim", "Global"))
				RunAction "Select and Save Credit Card Expense [Corporate Credit Card]", oneIteration, Datatable.Value("EECRefNo", "Global"), Datatable.Value("ExpenseCategory", "Global"), x
			Next	
			RunAction "Submit Expense Report [Corporate Credit Card]", oneIteration, Datatable.Value("EECRefNo", "Global"), bResult, bResult2
		End If
		
		'CC Personal Expense
		If Datatable.Value("Tab", "Global") = "5" and Datatable.Value("ToTest", "Global") = "Y" Then
			For x = 1 To CInt(Datatable.Value("NumberOfClaim", "Global"))
				RunAction "Select and Save Credit Card Expense [Corporate Credit Card]", oneIteration, Datatable.Value("EECRefNo", "Global"), Datatable.Value("ExpenseCategory", "Global"), x
			Next	
			RunAction "Submit Expense Report [Corporate Credit Card]", oneIteration, Datatable.Value("EECRefNo", "Global"), bResult, bResult2
		End If

		
		If bResult and bResult2 Then
			Reporter.ReportEvent micPass, Datatable.Value("TestName", "Global"), "Pass"
		Else
			Reporter.ReportEvent micFail, Datatable.Value("TestName", "Global"), "Fail"
		End If
Next

Set Initiate = nothing
