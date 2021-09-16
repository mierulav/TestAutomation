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

'Get all categories in the dropdown
wait(3)
arrCategoryList = Split(ClaimItem.GetExpenseCategoryListed, ";")

'Validation: Validate the expense category fields populated for each category are correct
ReDim arrbVal(Ubound(arrCategoryList))
For i = 0 To Ubound(arrCategoryList)

	ClaimItem.SelectExpenseCategory arrCategoryList(i)

	Select Case lcase(arrCategoryList(i))
		
		Case ""
			arrbVal(i) = True
			
		Case "accommodation"
			arrbVal(i) = ClaimItem.ValidateAccommodationObjects
			
		Case "advertising & promotion"
			arrbVal(i) =  ClaimItem.ValidateAdvertisingPromotionObjects
			
		Case "dental"
			arrbVal(i) = ClaimItem.ValidateGeneralClaimObjects
			
		Case "flights", "flight"
			arrbVal(i) = ClaimItem.ValidateFlightsObjects
			
		Case "gifts"
			arrbVal(i) = ClaimItem.ValidateGiftsObjects
			
		Case "internet"
			arrbVal(i) = ClaimItem.ValidateGeneralClaimObjects
			
		Case "meals"
			arrbVal(i) = ClaimItem.ValidateMealsObjects
			
		Case "other employee benefits"
			arrbVal(i) = ClaimItem.ValidateGeneralClaimObjects
			
		Case "others"
			arrbVal(i) = ClaimItem.ValidateGeneralClaimObjects
			
		Case "phone"
			arrbVal(i) = ClaimItem.ValidateGeneralClaimObjects
			
		Case "professional body subscription fee"
			arrbVal(i) = ClaimItem.ValidateGeneralClaimObjects
			
		Case "transport"
			arrbVal(i) = ClaimItem.ValidateTransportObjects
			
		Case "business/professional journal subscription fee"
			arrbVal(i) = ClaimItem.ValidateGeneralClaimObjects
			
		Case "courier"
			arrbVal(i) = ClaimItem.ValidateGeneralClaimObjects
			
		Case "employee education sponsorship/assistance"
			arrbVal(i) = ClaimItem.ValidateGeneralClaimObjects
			
		Case "other travel"
			arrbVal(i) = ClaimItem.ValidateOtherTravelObjects
			
		Case "per diem"
			arrbVal(i) = ClaimItem.ValidatePerDiemObjects
		
		Case "revenue stamps"
			arrbVal(i) = ClaimItem.ValidateGeneralClaimObjects
				
		Case "service parts"
			arrbVal(i) = ClaimItem.ValidateGeneralClaimObjects
			
		Case "other entertainment"
			arrbVal(i) = ClaimItem.ValidateOtherEntertainmentObjects
			
		Case "medical checkup/annual health checks"
			arrbVal(i) = ClaimItem.ValidateGeneralClaimObjects
			
		Case "medical expenses outpatient/inpatient"
			arrbVal(i) = ClaimItem.ValidateGeneralClaimObjects
		
		Case "employee children education fee"
			arrbVal(i) = ClaimItem.ValidateGeneralClaimObjects
		
		Case "fitness/wellness/lifestyle membership fee"
			arrbVal(i) = ClaimItem.ValidateGeneralClaimObjects
		
		Case "optical"
			arrbVal(i) = ClaimItem.ValidateGeneralClaimObjects
			
		Case "it consumable"
			arrbVal(i) = ClaimItem.ValidateGeneralClaimObjects
			
		Case Else
			Reporter.ReportEvent micFail, "Select Expense Category", "Fail"
			arrbVal(i) = False
			
	End Select
Next

For k = 0 To Ubound(arrbVal)
	If arrbVal(k) = False Then
		bVal = False
		Exit For
	End If
	bVal = True
Next

If bVal Then
	Reporter.ReportEvent micPass, "Validate Claim form fields for every category", "Pass"
Else
	Reporter.ReportEvent micFail, "Validate Claim form fields for every category", "Fail"
End If


'Create claim
ClaimItem.SelectExpenseCategory strCategory

'Upload proof of payment
ClaimItem.UploadFile Test.GetTestDataGlobal & strPath

'Step 2: Enter information into the mandatory fields
If ClaimItem.FillUpClaimDetails(strExpenseCategory, arrExpenseData) Then
	Reporter.ReportEvent micPass, "Validation 5: Fill up claim details successful", "Pass"
Else
	Reporter.ReportEvent micFail, "Validation 5: Fill up claim details successful", "Fail"
	Parameter("bResult") = False
	ExitAction
End If

ClaimItem.SetApplySingleTaxSelection(Parameter("strTaxPercentage"))

'Step 3: Save claim item
ClaimItem.SaveClaimItem
If ClaimItem.ValidateClaimItemCreated(strCategory, strReceiptAmount) and ClaimItem.ValidateClaimItemListPage Then
	Reporter.ReportEvent micPass, "Claim Item are saved and created", "Pass"
	Parameter("bResult") = True
Else
	Reporter.ReportEvent micFail, "Claim Item are saved and created", "Fail"
	Parameter("bResult") = False
End If

