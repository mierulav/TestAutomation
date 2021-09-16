Option Explicit

Dim Initiate : Set Initiate = Init()
Dim Login : Set Login = Login_Page()

'Login
Initiate.OpenURL Initiate.GetURL
Login.Login Initiate.GetUsername, Initiate.GetPassword

'Start Test
Dim TestDataFile : TestDataFile = Initiate.GetTestDataEntitlement & "\" & Initiate.GetCountry & ".xls"
	
With DataTable

.ImportSheet TestDataFile, "Main", "Global"

Dim x, i, j, ColCount, bResult, strTitle, strVal, strEmails, strRefNo, strAmount, strTaxAmount, arrValueToExcel

For i = 1 To .GetSheet("Global").GetRowCount
	.GetSheet("Global").SetCurrentRow(i)
	If .Value(Initiate.GetCompanyCode, "Global") = "V" or .Value(Initiate.GetCompanyCode, "Global") = "MT" Then
		.ImportSheet TestDataFile, .Value("Tab"), "Local"
		For j = 1 To .GetSheet("Local").GetRowCount
			.GetSheet("Local").SetCurrentRow(j)
			If .Value("ToTest", "Local") = "MT" Then
				ColCount = .GetSheet("Local").GetParameterCount
				ReDim arrVal(ColCount-3)
				For x = 0 To Ubound(arrVal)
					arrVal(x) = .GetSheet("Local").GetParameter(3+x)
				Next
				
				strVal = Join(arrVal, ",")
				strEmails = Initiate.GetUsername & ";" & "amirul.saddam@dksh.com"
				
				'Run action
				RunAction "Create Complete Claim (Single Tax for MultiTax Selection) [Quick Add]", oneIteration, strVal, .Value("Remarks", "Local"), strAmount, strTaxAmount
				RunAction "Submit Expense Report (Returns ExpenseTitle) [Expense Report]", oneIteration, strEmails, bResult, strTitle
				RunAction "Get Expense Report Reference Number [Expense Report]", oneIteration, strTitle, strRefNo, bResult

				arrValueToExcel = Array(Initiate.GetCompanyCode, "", "", Initiate.GetUsername, strRefNo, Datatable.Value("ExpenseCategory", "Global"), strAmount, "Y", strTaxAmount)
				Datatable.GetSheet("Global").AddParameter "Result" & j, Join(arrValueToExcel, ",")
	
			End If
		Next
	End If
Next

Datatable.ExportSheet Initiate.GetTestDataGlobal & "\ExportedTestData\" & Initiate.GetTestName & ".xls", "Global", Initiate.GetCompanyCode

End With

Set Initiate = nothing

