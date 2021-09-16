Option Explicit

Dim Home : Set Home = Home_Page()
Dim ClaimItem : Set ClaimItem = ClaimItem_Page()
Dim intRowInitial, strCategory, intIndex, intRowAfter
intIndex = 1
strCategory = Parameter("strCategory")

'Navigate to Claim Item
Home.NavigateToClaimItem

'Get initial rowcount
intRowInitial = ClaimItem.GetClaimItemCount

'Search for specific item to delete, by Claim category, and delete
If ClaimItem.SearchClaimItem(strCategory) Then
	ClaimItem.DeleteClaimItem
Else
	Reporter.ReportEvent micDone, "No claim item with searched category", "Done"
	ExitAction
End If 

'Get rowcount after
intRowAfter = ClaimItem.GetClaimItemCount

'Validate
If intRowAfter >= intRowInitial Then
	Reporter.ReportEvent micFail, "Delete claim item successful", "Fail"
	Parameter("bResult") = False
Else
	Reporter.ReportEvent micPass, "Delete claim item successful", "Pass"
	Parameter("bResult") = True
End If

