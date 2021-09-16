Option Explicit

Dim Test : Set Test = Init()
Dim Home : Set Home = Home_Page
Dim ClaimItem : Set ClaimItem = ClaimItem_Page
Dim strPath, strCategory, strReceiptAmount, arrExpenseData, arrCategoryList
Dim i, k, bVal
arrExpenseData = Split(Parameter("strExpenseData"), ",")
strPath = arrExpenseData(0)
strCategory = arrExpenseData(1)
strReceiptAmount = arrExpenseData(4)

'Navigate to Claim Item screen
Home.NavigateToClaimItem

'Click on Create Claim btn
ClaimItem.CreateClaimItem

'Create claim
ClaimItem.SelectExpenseCategory strCategory

'Upload proof of payment
ClaimItem.UploadFile Test.GetTestDataGlobal & strPath

'Step 2: Enter information into the mandatory fields
If ClaimItem.FillUpClaimDetails(strCategory, arrExpenseData) Then
	Reporter.ReportEvent micPass, "Validation 5: Fill up claim details successful", "Pass"
Else
	Reporter.ReportEvent micFail, "Validation 5: Fill up claim details successful", "Fail"
	Parameter("bResult") = False
	ExitAction
End If

'Step 3: Save claim item
ClaimItem.SaveClaimItem
If ClaimItem.ValidateClaimItemCreated(strCategory, strReceiptAmount) and ClaimItem.ValidateClaimItemListPage Then
	Reporter.ReportEvent micPass, "Claim Item are saved and created", "Pass"
	Parameter("bResult") = True
Else
	Reporter.ReportEvent micFail, "Claim Item are saved and created", "Fail"
	Parameter("bResult") = False
End If

