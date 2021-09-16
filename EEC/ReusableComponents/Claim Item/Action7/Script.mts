Option Explicit

Dim Test : Set Test = Init()
Dim Home : Set Home = Home_Page
Dim ClaimItem : Set ClaimItem = ClaimItem_Page
Dim strPath, strCategory, strReceiptAmount, arrExpenseData, arrCategoryList, arrRes
Dim i, k, bVal
arrExpenseData = Split(Parameter("strExpenseData"), ",")
strPath = arrExpenseData(0)
strCategory = arrExpenseData(1)
strReceiptAmount = arrExpenseData(4)
arrRes = Array(False, False, False)

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

'Validate that exchange rate and amount in local currency field is displayed and is not edittable.
If ClaimItem.ValidateExchangeCurrencyFields Then
	Reporter.ReportEvent micPass, "Exchange currency fields are displayed", "Pass"
	arrRes(0) = True
Else
	Reporter.ReportEvent micFail, "Exchange currency fields are displayed", "Fail"
End If

'Validate local amount is calculated based on the conversion rate and invoice value
If ClaimItem.ValidateLocalCurrencyConvertedAmount Then
	Reporter.ReportEvent micPass, "Local conversion amount is correct", "Pass"
	arrRes(1) = True
Else
	Reporter.ReportEvent micFail, "Local conversion amount is correct", "Fail"
End If

'Save claim item
ClaimItem.SaveClaimItem
If ClaimItem.ValidateClaimItemCreated(strCategory, strReceiptAmount) Then
	Reporter.ReportEvent micPass, "Claim Item are saved and created", "Pass"
	arrRes(2) = True
Else
	Reporter.ReportEvent micFail, "Claim Item are saved and created", "Fail"
End If

For i = 1 To Ubound(arrRes)
	If arrRes(i) = False Then
		Parameter("bResult") = False
		Exit For
	End If
	Parameter("bResult") = True
Next

 @@ script infofile_;_ZIP::ssf2.xml_;_
