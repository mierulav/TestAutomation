Option Explicit

Dim Home : Set Home = Home_Page
Dim ClaimItem : Set ClaimItem = ClaimItem_Page

'Step 1: Navigate to claim item page
Home.NavigateToClaimItem

'Step 2: Click on create claim item link
ClaimItem.CreateClaimItem

wait(2)

'Step 3: Select expense category
ClaimItem.SelectExpenseCategory Parameter("strExpenseCategory")

Select Case Lcase(Parameter("strExpenseCategory"))
	
	Case "meals"
		Parameter("bResult") = ClaimItem.ValidateMealsAutopopulatedField
	
	Case "gifts"
		Parameter("bResult") = ClaimItem.ValidateGiftsAutopopulatedField
		
	Case Else
		Parameter("bResult") =  False
		
End Select

Set Home = Nothing
Set ClaimItem = Nothing
