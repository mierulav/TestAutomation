Option Explicit

Dim Test : Set Test = Init()
Dim Home : Set Home = Home_Page
Dim ClaimItem : Set ClaimItem = ClaimItem_Page
Dim strPath, strCategory, arrExpenseData, bVal
arrExpenseData = Split(Parameter("strExpenseData"), ",")
strPath = arrExpenseData(0)
strCategory = arrExpenseData(1)

'Step 1: Navigate to Claim item screen
Home.NavigateToClaimItem

'Step 1: Search and find Open up a specific claim item based on claim category
ClaimItem.SearchClaimItemByClaimCategory strCategory
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


'Step 4: Cancel by click cross icon at the top-right form
ClaimItem.ExitClaimItem

'Validate 
Home.NavigateToClaimItem
ClaimItem.SearchClaimItemByClaimCategory strCategory
ClaimItem.EditClaimItem

Wait(3)

RunAction "Verify Claim Data", oneIteration, Parameter("strExpenseData"), bVal

If bVal = False Then
	Reporter.ReportEvent micPass, "Data not modified", "Pass"
	Parameter("bResult") = True
Else
	Reporter.ReportEvent micFail, "Data not modified", "Fail"
	Parameter("bResult") = False
End If

