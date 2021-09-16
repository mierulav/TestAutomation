Option Explicit

Dim Home : Set Home = Home_Page
Dim ClaimItem : Set ClaimItem = ClaimItem_Page
Dim strExpenseCategory : strExpenseCategory = Parameter("strExpenseCategory")	
Dim arrPVal, arrTVal
			
'Step 1: Navigate to claim item page
Home.NavigateToClaimItem

'Step 2: Click on create claim item link
ClaimItem.CreateClaimItem

Wait(5)

'Step 3: Select expense category
If Instr(strExpenseCategory, "Phone") Then
	arrPVal = Split(strExpenseCategory, "-")
	ClaimItem.SelectExpenseCategory arrPVal(0)
	ClaimItem.SelectCategoryType arrPVal
ElseIf Instr(strExpenseCategory, "Transport") Then 
	arrTVal = Split(strExpenseCategory, "-")
	ClaimItem.SelectExpenseCategory arrTVal(0)
	ClaimItem.SelectCategoryType arrTVal
Else
	ClaimItem.SelectExpenseCategory strExpenseCategory
End If

Wait(5)

'Step 4: Select default date (today's date)
ClaimItem.SetReceiptDate FormatDate(Date)

''Validation: Validate expense category with tax is correct
Parameter("bResult") = ClaimItem.ValidateClaimItemTaxFields

Set Home = nothing
Set ClaimItem = nothing
