Option Explicit

Dim Home : Set Home = Home_Page
Dim ClaimItem : Set ClaimItem = ClaimItem_Page
Dim strClaim : strClaim = Parameter("strClaimEntitlement")
Dim arrClaim, ClaimList

'Step 1: Navigate to claim item page
Home.NavigateToClaimItem

'Step 2: Click on create claim item link
ClaimItem.CreateClaimItem

wait(5)

'Step 3: Select expense category
ClaimList = ClaimItem.GetExpenseCategoryListed

If Instr(strClaim, "Phone") > 0 Then
	arrClaim =  Split(strClaim, "-")
	strClaim = empty
	strClaim = "Phone"
End If

If Instr(strClaim, "Transport") > 0 Then
	arrClaim =  Split(strClaim, "-")	
	strClaim = empty
	strClaim = "Transport"	
End If

Select Case strClaim
	
	Case "Phone"
		If Instr(ClaimList, strClaim) = 0  Then
			Parameter("bResult") =  False
			ExitAction
		Else
			Browser("EEC").Page("EEC | Claim Items").WebList("drpClaimCategory").Select "Phone"
			
			If Instr(ClaimItem.GetPhoneExpenseCategoryListed, arrClaim(1)) = 0 Then
				Parameter("bResult") = False
				ExitAction
			Else
				Parameter("bResult") = True
			End If
		End If
		
	Case "Transport"
		If Instr(ClaimItem.GetTransportExpenseCategoryListed, strClaim) = 0  Then
			Parameter("bResult") =  False
			ExitAction
		Else
			Browser("EEC").Page("EEC | Claim Items").WebList("drpClaimCategory").Select "Transport"
			
			If Instr(ClaimList, arrClaim(1)) = 0 Then
				Parameter("bResult") = False
				ExitAction
			Else
				Parameter("bResult") = True
			End If
		End If
	
	Case Else
		If Instr(ClaimList, strClaim) = 0  Then
			Parameter("bResult") = False
		Else
			Parameter("bResult") = True
		End If

End Select

Set Home = nothing
Set ClaimItem = nothing
