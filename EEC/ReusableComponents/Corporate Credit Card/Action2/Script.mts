Option Explicit

'Declarations
Dim Initiate : Set Initiate = Init()
Dim Homepage : Set Homepage = Home_Page()
Dim ClaimItem : Set ClaimItem = ClaimItem_Page()
Dim ExpenseReport : Set ExpenseReport = ExpenseReport_Page()
Dim strSearch, strRefNo, strType, strExpectedStatus
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
'If strType <> "" Then
'	ClaimItem.SelectExpenseCategory strType
'End If
'
'ClaimItem.UploadFile Initiate.GetTestDataGlobal & "\ReceiptToUpload\TestPurpose.xlsx"
'
''Save and exit claim
'ClaimItem.SaveClaimItem

'Check
ExpenseReport.SetExpenseReportReceiptCertified

'SUbmit draft
ExpenseReport.SubmitExpenseReport

'Search back expense using ref no and validate expense report submitted status
Parameter("bResult1") = ExpenseReport.ValidateExpenseReportSubmitted(strRefNo)

'Open the expense report
ExpenseReport.EditSpecificExpense(strRefNo)

'Validate audit log also submitted
Parameter("bResult2") = ExpenseReport.ValidateAuditLogSubmittedStatus()

