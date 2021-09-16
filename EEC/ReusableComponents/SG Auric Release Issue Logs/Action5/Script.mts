Option Explicit

Dim Initiate : Set Initiate = Init

Dim Home : Set Home = Home_Page
Dim QuickAdd : Set QuickAdd = QuickAdd_Page
Dim ClaimItem : Set ClaimItem = ClaimItem_Page
Dim i, j, x

DataTable.ImportSheet Initiate.GetTestCaseData & "\ExpenseCategoryData.xls", Parameter("strVal"), "Local"

ReDim tempResult(DataTable.GetSheet("Local").GetRowCount)
ReDim arrExpenseData(DataTable.GetSheet("Local").GetParameterCount-2)
For x = 1 To Datatable.GetSheet("Local").GetRowCount
	Datatable.GetSheet("Local").SetCurrentRow(x)
	If DataTable.Value("ToTest", "Local") = "Y" Then
		For i = 0 To Ubound(arrExpenseData)
			arrExpenseData(i) = DataTable.GetSheet("Local").GetParameter(2+i)
		Next
		
		'Step 1: Navigate to claim item page
		Home.NavigateToQuickAdd
		
		'Step 2: Upload proof of payment
		QuickAdd.UploadFile arrExpenseData(0)
		
		'Step 4: Select expense category
		QuickAdd.SelectExpensesCategory arrExpenseData(1)
		QuickAdd.SaveClaim
		QuickAdd.SaveCompleteAlert
		
		Wait(2)
		
		Select Case lcase(arrExpenseData(1))
			
			Case "accommodation"
				ClaimItem.FillAccommodationForm arrExpenseData
				
			Case "advertising & promotion"
				ClaimItem.FillAdvertisingPromotionForm arrExpenseData
				
			Case "dental"
				ClaimItem.FillDentalClaimForm arrExpenseData
				
			Case "flights"
				ClaimItem.FillFlightForm arrExpenseData
				
			Case "gifts"
				ClaimItem.FillGiftForm arrExpenseData
				
			Case "internet"
				ClaimItem.FillInternetClaimForm arrExpenseData
				
			Case "meals"
				ClaimItem.FillMealForm arrExpenseData
				
			Case "other employee benefits"
				ClaimItem.FillOEBClaimForm arrExpenseData
				
			Case "others"
				ClaimItem.FillOthersClaimForm arrExpenseData
				
			Case "phone"
				ClaimItem.FillPhoneForm arrExpenseData
				
			Case "professional body subscription fee"
				ClaimItem.FillPBSClaimForm arrExpenseData
				
			Case "transport"
				ClaimItem.FillTransportForm arrExpenseData
				
			Case "business/professional journal subscription fee"
				ClaimItem.FillBJSClaimForm arrExpenseData
				
			Case "courier"
				ClaimItem.FillCourierClaimForm arrExpenseData
				
			Case "employee education sponsorship/assistance"
				ClaimItem.FillEESClaimForm arrExpenseData
				
			Case "other travel"
				ClaimItem.FillOtherTravelClaimForm arrExpenseData
				
			Case "per diem"
				ClaimItem.FillPerDiemClaimForm arrExpenseData
			
			Case "revenue stamps"
				ClaimItem.FillRevenueStampsClaimForm arrExpenseData
					
			Case "service parts"
				ClaimItem.FillServicePartsClaimForm arrExpenseData
				
			Case "other entertainment"
				ClaimItem.FillOtherEntertainmentClaimForm arrExpenseData
				
			Case "medical checkup/annual health checks"
				ClaimItem.FillMedicalCheckupClaimForm arrExpenseData
				
			Case "medical expenses outpatient/inpatient"
				ClaimItem.FillMedicalExpensesClaimForm arrExpenseData
			
			Case "employee children education fee"
				ClaimItem.FillChildrenEducationFeeClaimForm arrExpenseData
			
			Case "fitness/wellness/lifestyle membership fee"
				ClaimItem.FillWellnessClaimForm arrExpenseData
			
			Case "optical"
				ClaimItem.FillOpticalClaimForm arrExpenseData
				
			Case Else
				Reporter.ReportEvent micFail, "Select Expense Category", "Fail"
				Parameter("bResult") = False
				
		End Select
		
		'Step 5: Save claim item
		ClaimItem.SaveClaimItem

		'Step 6: Select existing claim
		ClaimItem.EditClaimItem
		
		Wait(2)
		'Step 7: Remove cost center
		ClaimItem.RemoveCostCenter
		
		'Step 8: Save Claim
		ClaimItem.SaveClaimItem
		
		Wait(2)
		'Validate: Validate Save successful
		tempResult(x-1) = ClaimItem.ValidateClaimItemListPage

	End If
Next

For j = 0 To Ubound(tempResult)
	If tempResult(j) = False Then
		Parameter("bResult") = False
		ExitAction
	End If	
Next

Parameter("bResult") = True

