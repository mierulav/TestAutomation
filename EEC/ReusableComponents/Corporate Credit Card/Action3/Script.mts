Option Explicit

'Declarations
Dim Initiate : Set Initiate = Init()
Dim Homepage : Set Homepage = Home_Page()
Dim ClaimItem : Set ClaimItem = ClaimItem_Page()
Dim ExpenseReport : Set ExpenseReport = ExpenseReport_Page()
Dim strSearch, strRefNo, strType
strSearch = Parameter("strRefNo")
strType = Parameter("strType")
strIndex = Parameter("strIndex")

'Navigate to My expense report
Homepage.NavigateToMyExpenseReport

'Search for expense title starts wih System Generate *
ExpenseReport.SearchExpenseReport strSearch

'Get first from the list with draft status
ExpenseReport.EditSpecificExpense(strSearch)

'Click on edit claim item
ExpenseReport.EditSpecificClaim(intIndex)

Wait(5)

'Set Category to Personal expense, disputed, or any valid category
If strType <> "" Then
	ClaimItem.SelectExpenseCategory strType
End If

ClaimItem.UploadFile Initiate.GetTestDataGlobal & "\ReceiptToUpload\TestPurpose.xlsx"

'Save and exit claim
ClaimItem.SaveClaimItem
