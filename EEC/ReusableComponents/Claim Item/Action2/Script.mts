Option Explicit

Dim Test : Set Test = Init()
Dim Home : Set Home = Home_Page
Dim ClaimItem : Set ClaimItem = ClaimItem_Page
Dim strPath, strCategory, arrExpenseData, strReceiptAmount
arrExpenseData = Split(Parameter("strExpenseData"), ",")
strPath = arrExpenseData(0)
strCategory = arrExpenseData(1)
strReceiptAmount = arrExpenseData(4)

'Step 1: Navigate to Claim item screen
Home.NavigateToClaimItem

'Step 1: Search and find Open up a specific claim item based on claim category
ClaimItem.SearchClaimItem strCategory
ClaimItem.EditClaimItem

'Delete existing proof of payment
If ClaimItem.RemoveAttachment Then
	Reporter.ReportEvent micPass, "Proof of payment is deleted", "Pass"
Else
	Reporter.ReportEvent micFail, "Proof of payment is deleted", "Fail"
End If

'Upload new attachmenet
If ClaimItem.UploadFile(Test.GetTestDataGlobal & strPath) Then
	Reporter.ReportEvent micPass, "Proof of payment is uploadd", "Pass"
Else
	Reporter.ReportEvent micFail, "Proof of payment is uploadd", "Fail"
End If

'make changes on mandatory fields
If ClaimItem.FillUpClaimDetails(strCategory, arrExpenseData) Then
	Reporter.ReportEvent micPass, "Fill up claim details successful", "Pass"
Else
	Reporter.ReportEvent micFail, "Fill up claim details successful", "Fail"
	Parameter("bResult") = False
	ExitAction
End If

'Step 3: Save claim item
ClaimItem.SaveClaimItem
If ClaimItem.ValidateClaimItemCreated(strCategory, strReceiptAmount) Then
	Reporter.ReportEvent micPass, "Claim Item are saved and created", "Pass"
Else
	Reporter.ReportEvent micFail, "Claim Item are saved and created", "Fail"
End If

'Validate
Home.NavigateToClaimItem
ClaimItem.SearchClaimItem strCategory
ClaimItem.EditClaimItem
Wait(3)

RunAction "Verify Claim Data", oneIteration, Parameter("strExpenseData"), Parameter("bResult")

If Parameter("bResult") Then
	Reporter.ReportEvent micPass, "Data modified successfully", "Pass"
Else
	Reporter.ReportEvent micFail, "Data modified successfully", "Fail"
End If
