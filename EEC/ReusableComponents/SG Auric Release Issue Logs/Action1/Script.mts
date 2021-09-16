Option Explicit @@ hightlight id_;_5575750_;_script infofile_;_ZIP::ssf2.xml_;_

Dim Home : Set Home = Home_Page
Dim ClaimItem : Set ClaimItem = ClaimItem_Page
Dim strExpenseCategory : strExpenseCategory = Parameter("strExpenseCategory")
	
'Step 1: Navigate to claim item page
Home.NavigateToClaimItem

'Step 2: Click on create claim item link
ClaimItem.CreateClaimItem

Wait(5)

'Step 3: Select expense category
ClaimItem.SelectExpenseCategory strExpenseCategory

'Validation: Validate the expense category fields populated for each category are correct
Select Case lcase(strExpenseCategory)
	
	Case "accommodation"
		Parameter("bResult") = ClaimItem.ValidateAccommodationObjects
		
	Case "advertising & promotion"
		Parameter("bResult") =  ClaimItem.ValidateAdvertisingPromotionObjects
		
	Case "dental"
		Parameter("bResult") = ClaimItem.ValidateGeneralClaimObjects
		
	Case "flights"
		Parameter("bResult") = ClaimItem.ValidateFlightsObjects
		
	Case "gifts"
		Parameter("bResult") = ClaimItem.ValidateGiftsObjects
		
	Case "internet"
		Parameter("bResult") = ClaimItem.ValidateGeneralClaimObjects
		
	Case "meals"
		Parameter("bResult") = ClaimItem.ValidateMealsObjects
		
	Case "other employee benefits"
		Parameter("bResult") = ClaimItem.ValidateGeneralClaimObjects
		
	Case "others"
		Parameter("bResult") = ClaimItem.ValidateGeneralClaimObjects
		
	Case "phone"
		Parameter("bResult") = ClaimItem.ValidateGeneralClaimObjects
		
	Case "professional body subscription fee"
		Parameter("bResult") = ClaimItem.ValidateGeneralClaimObjects
		
	Case "transport"
		Parameter("bResult") = ClaimItem.ValidateTransportObjects
		
	Case "business/professional journal subscription fee"
		Parameter("bResult") = ClaimItem.ValidateGeneralClaimObjects
		
	Case "courier"
		Parameter("bResult") = ClaimItem.ValidateGeneralClaimObjects
		
	Case "employee education sponsorship/assistance"
		Parameter("bResult") = ClaimItem.ValidateGeneralClaimObjects
		
	Case "other travel"
		Parameter("bResult") = ClaimItem.ValidateOtherTravelObjects
		
	Case "per diem"
		Parameter("bResult") = ClaimItem.ValidatePerDiemObjects
	
	Case "revenue stamps"
		Parameter("bResult") = ClaimItem.ValidateGeneralClaimObjects
			
	Case "service parts"
		Parameter("bResult") = ClaimItem.ValidateGeneralClaimObjects
		
	Case "other entertainment"
		Parameter("bResult") = ClaimItem.ValidateOtherEntertainmentObjects
		
	Case "medical checkup/annual health checks"
		Parameter("bResult") = ClaimItem.ValidateGeneralClaimObjects
		
	Case "medical expenses outpatient/inpatient"
		Parameter("bResult") = ClaimItem.ValidateGeneralClaimObjects
	
	Case "employee children education fee"
		Parameter("bResult") = ClaimItem.ValidateGeneralClaimObjects
	
	Case "fitness/wellness/lifestyle membership fee"
		Parameter("bResult") = ClaimItem.ValidateGeneralClaimObjects
	
	Case "optical"
		Parameter("bResult") = ClaimItem.ValidateGeneralClaimObjects
		
	Case Else
		Reporter.ReportEvent micFail, "Select Expense Category", "Fail"
		Parameter("bResult") = False
		
End Select

Set Home = nothing
Set ClaimItem = nothing
