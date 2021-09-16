Option Explicit

Dim Initiate : Set Initiate = Init
Dim Home : Set Home = Home_Page
Dim ClaimItem : Set ClaimItem = ClaimItem_Page
Dim i, x

DataTable.ImportSheet Initiate.GetTestCaseData & "\MasterData-Country.xls", "Sheet1", "Local"
ReDim arrVal(DataTable.GetSheet("Local").GetRowCount-1)
For x = 1 To DataTable.GetSheet("Local").GetRowCount
	Datatable.GetSheet("Local").SetCurrentRow(x)
	arrVal(x-1) = DataTable.Value("CountryName", "Local")
Next

'Step 1: Navigate to claim item page
Home.NavigateToClaimItem

'Step 2: Click on create claim item link
ClaimItem.CreateClaimItem

Wait(3)

'Step 3: Select expense category
ClaimItem.SelectExpenseCategory Parameter("strExpenseCategory")

'Step 4: Get country field list
Dim arrCountryListed : arrCountryListed = ClaimItem.GetCountryFieldListed
Dim strCountries : strCountries = Join(arrVal, ";")

'Step 5: Validate country master data with country field listed
Dim falseCount : falseCount = 0
For i = 0 To Ubound(arrCountryListed)
	If Instr(strCountries, arrCountryListed(i)) = 0 and arrCountryListed(i) <> "Other"  Then
		falseCount = falseCount + 1
		Reporter.AddRunInformation "Issue", "This country(s) not available in the master data - " & arrCountryListed(i)
	End If	
Next

If falseCount = 0 Then
	Parameter("bResult") = True
Else
	Parameter("bResult") = False
End If





