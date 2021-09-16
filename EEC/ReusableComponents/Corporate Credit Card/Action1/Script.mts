Option Explicit

'Declarations
Dim Initiate : Set Initiate = Init()
Dim Homepage : Set Homepage = Home_Page()
Dim ClaimItem : Set ClaimItem = ClaimItem_Page()
Dim ExpenseReport : Set ExpenseReport = ExpenseReport_Page()
Dim strSearch, strRefNo, strType
strRefNo = Parameter("strRefNo")
'strType = Parameter("strType")
'
''Navigate to My expense report
'Homepage.NavigateToMyExpenseReport
'
''Search for expense title starts wih System Generate *
'ExpenseReport.SearchExpenseReport strSearch
'
''Get first from the list with draft status
'ExpenseReport.EditExpenseReport
'
''Get Ref No of the expense report
'strRefNo = ExpenseReport.GetReferenceNumber
'
''Click on edit claim item
'ExpenseReport.EditClaimItem
'
'Wait(5)
'
''Set Category to Personal expense, disputed, or any valid category
'If strType <> "" Then
'	ClaimItem.SelectExpenseCategory strType
'End If
'
'ClaimItem.UploadFile Initiate.GetTestDataGlobal & "\ReceiptToUpload\TestPurpose.xlsx"
'
''Save and exit claim
'ClaimItem.SaveClaimItem

'Save draft
ExpenseReport.SaveDraftExpenseReport

'Search back expense using ref no
Parameter("bResult1") = ExpenseReport.ValidateExpenseReportDraft(strRefNo)



