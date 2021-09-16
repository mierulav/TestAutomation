Option Explicit

Dim Home : Set Home = Home_Page()
Dim QuickAdd : Set QuickAdd = QuickAdd_Page()
Dim ClaimItem : Set ClaimItem = ClaimItem_Page()
Dim Test : Set Test = Init()

Parameter("bResult") = True
Dim arrExpenseData : arrExpenseData = Split(Parameter("strExpenseData"), ",")
Dim strProof : strProof = arrExpenseData(0)
Dim strExpenseCategory : strExpenseCategory = arrExpenseData(1)

'Step 1: Navigate to Quick Add page
Home.NavigateToQuickAdd

'Handle tutorial pop up
'Home.HandleQuickTip

'Validation Step 1
If Home.GetLanguageUsed = "English" Then
	Reporter.ReportEvent micPass, "Validation 1: English Language", "Pass"
Else
	Reporter.ReportEvent micFail, "Validation 1: English Language", "Fail"
	Parameter("bResult") = False
End If

If QuickAdd.ValidateQuickAddPage Then
	Reporter.ReportEvent micPass, "Validation 1: Quick Add screen successfully displayed", "Pass"
Else
	Reporter.ReportEvent micFail, "Validation 1: Quick Add screen successfully displayed", "Fail"
	Parameter("bResult") = False
End If

'Step 2: Upload the proof of payment attachment.
QuickAdd.UploadFile Test.GetTestDataGlobal & strProof

'Validation Step 2
If QuickAdd.ValidateProofFileUploaded Then
	Reporter.ReportEvent micPass, "Validation 2: Payment proof less than 2Mb successfully uploaded", "Pass"
Else
	Reporter.ReportEvent micFail, "Validation 2: Payment proof less than 2Mb successfully uploaded", "Fail"
	Parameter("bResult") = False
End If

'Step 3: Select the expense category.
If QuickAdd.SelectExpensesCategory(strExpenseCategory) = False Then
	Reporter.ReportEvent micFail, "Validation 3: Claim Category is Not Found", "Fail"
	ExitAction
End If

'Step 4: Click on Save button
QuickAdd.SaveClaim

Wait(2)

'Validation Step 4
If QuickAdd.SaveCompleteAlert Then
	Reporter.ReportEvent micPass, "Validation 4: Prompt message displayed", "Pass"
Else
	Reporter.ReportEvent micFail, "Validation 4: Prompt message displayed", "Fail"
	Parameter("bResult") = False
End If

'wait(5)
'
''handle dialog box - currently in JP country for Ad & Promo EC
'If Browser("EEC").DialogExists Then
'	Browser("EEC").HandleDialog micOK
'End If

If ClaimItem.ValidateClaimItemDetailPage Then
	Reporter.ReportEvent micPass, "Validation 4: Claim Item details page displayed", "Pass"
Else
	Reporter.ReportEvent micFail, "Validation 4: Claim Item details page displayed", "Fail"
	Parameter("bResult") = False
End If

Wait(3)

 'Step 5: Enter information into the mandatory fields
If ClaimItem.FillUpClaimDetails(strExpenseCategory, arrExpenseData) Then
	Reporter.ReportEvent micPass, "Validation 5: Fill up claim details successful", "Pass"
Else
	Reporter.ReportEvent micFail, "Validation 5: Fill up claim details successful", "Fail"
	Parameter("bResult") = False
	ExitAction
End If

'get Amount
Parameter("strAmount") = ClaimItem.GetReceiptAmount

If ClaimItem.ValidateClaimItemTaxFields Then
	Reporter.ReportEvent micFail, "Validation 5: Claim Item Tax fields are displayed", "Fail"
	Parameter("bResult") = False
Else
	Reporter.ReportEvent micPass, "Validation 5: Claim Item Tax fields are displayed", "Pass"
End If

'Step 6: Click on the Save and exit button
ClaimItem.SaveClaimItem

If ClaimItem.ValidateClaimItemListPage Then
	Reporter.ReportEvent micPass, "Validation 6: Claim Item list page displayed", "Pass"
Else
	Reporter.ReportEvent micFail, "Validation 6: Claim Item list page displayed", "Fail"
End If

 Set Home = Nothing
 Set QuickAdd = Nothing
 Set ClaimItem = Nothing

