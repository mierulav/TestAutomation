Option Explicit

Dim Home : Set Home = Home_Page
Dim ClaimItem : Set ClaimItem = ClaimItem_Page
Dim defaultVal : defaultVal	= "NO"
Dim arrExpenseCategory : arrExpenseCategory = Split(Parameter("strExpenseCategory"), ",")
Dim i, x, tempResult
ReDim arrResult(Ubound(arrExpenseCategory))

'Step 1: Navigate to claim item page
Home.NavigateToClaimItem

'Step 2: Click on create claim item link
ClaimItem.CreateClaimItem

For i = 0 To Ubound(arrExpenseCategory)

	Wait(1)
	'Step 3: Select expense category
	ClaimItem.SelectExpenseCategory arrExpenseCategory(i)
	
	''Validation: Validate expense category with tax is correct
	If Ucase(ClaimItem.GetRecoverableValue) = defaultVal Then
		tempResult = True
	Else
		tempResult = False
	End If
	
	arrResult(i) = tempResult
	
Next

For x = 0 To Ubound(arrResult)
	If arrResult(x) = False Then
		Parameter("bResult") = False
		ExitAction
	End If
Next

Parameter("bResult") = True
Set Home = nothing
Set ClaimItem = nothing
