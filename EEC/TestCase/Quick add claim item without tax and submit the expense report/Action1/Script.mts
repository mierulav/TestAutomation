Option Explicit

Dim Initiate : Set Initiate = Init()
Dim Login : Set Login = Login_Page()

'Login
Initiate.OpenURL Initiate.GetURL
Login.Login Initiate.GetUsername, Initiate.GetPassword

Dim TestDataFile : TestDataFile = Initiate.GetTestDataEntitlement & "\" & Initiate.GetCountry & ".xls"
	
With DataTable

.AddSheet "Entitlement"
.AddSheet "ExpenseData"
.ImportSheet TestDataFile, "Main", "Entitlement"

Dim x, i, j, ColCount, bResult, strTitle, strVal, strEmails, strRefNo, strAmount, arrValueToExcel

For i = 1 To .GetSheet("Entitlement").GetRowCount
	.GetSheet("Entitlement").SetCurrentRow(i)
	If .Value(Initiate.GetCompanyCode, "Entitlement") = "XT" or .Value(Initiate.GetCompanyCode, "Entitlement") = "V" Then
		.ImportSheet TestDataFile, .Value("Tab", "Entitlement"), "ExpenseData"
		For j = 1 To .GetSheet("ExpenseData").GetRowCount
			.GetSheet("ExpenseData").SetCurrentRow(j)
			If .Value("ToTest", "ExpenseData") = "Y" or .Value("ToTest", "ExpenseData") = "XT" Then
				ColCount = .GetSheet("ExpenseData").GetParameterCount
				ReDim arrVal(ColCount-3)
				For x = 0 To Ubound(arrVal)
					arrVal(x) = .GetSheet("ExpenseData").GetParameter(3+x)
				Next
				
				strVal = Join(arrVal, ",")
				strEmails = Initiate.GetUsername & ";" & "amirul.saddam@dksh.com"
					
				RunAction "Create Complete Claim (No Tax) [Quick Add]", oneIteration, strVal, bResult, strAmount
				
				If bResult Then
					RunAction "Submit Expense Report (Returns ExpenseTitle) [Expense Report]", oneIteration, strEmails, bResult, strTitle
	 				RunAction "Get Expense Report Reference Number [Expense Report]", oneIteration, strTitle, strRefNo, bResult
	
					arrValueToExcel = Array(Initiate.GetCompanyCode, "", "", Initiate.GetUsername, strRefNo, Datatable.Value("ExpenseCategory", "Entitlement"), strAmount, "N", "")
					Datatable.GetSheet("Entitlement").AddParameter "Result" & j, Join(arrValueToExcel, ",")
				End If
				
				
			End If
		Next
	End If
Next

Datatable.ExportSheet Initiate.GetTestDataGlobal & "\ExportedTestData\" & Initiate.GetTestName & ".xls", "Entitlement", Initiate.GetCompanyCode

End With

Set Initiate = nothing

