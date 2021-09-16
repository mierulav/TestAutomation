Option Explicit

Dim Home : Set Home = Home_Page
Dim ClaimItem : Set ClaimItem = ClaimItem_Page
Dim arrVal, bVal, bVal2
Dim i, j, k, l, x
Dim arrInput : arrInput = Split(Parameter("strExpenseData"), ",")
Dim strPath : strPath = arrInput(0)

'Validate proof of payment (for only 1 proof upload)
If ClaimItem.ValidateAttachmentExist Then
	Reporter.ReportEvent micPass, "Validate proof of payment upload is correct", "Pass"
	bVal2 = True
Else
	Reporter.ReportEvent micPass, "Validate proof of payment upload is correct", "Fail"
	bVal2 = False
End If

'Verify all data is exactly as what is filled in during claim creation
arrVal = ClaimItem.GetClaimFieldsValue(arrInput(1))	

Wait(2)
 print "Verify Claim Data: Checking data changes"
ReDim arrResult(Ubound(arrVal))
For i = 1 To Ubound(arrInput)
	For j = 0 To Ubound(arrVal)
		If arrInput(i) = "" Then
			arrResult(j) = True
			Exit For
		End If
		If arrInput(i) = arrVal(j) Then
			arrResult(j) = True
			Exit For
		End If
		If arrInput(i) <> arrVal(j) and j = Ubound(arrVal) Then
			arrResult(j) = False
		End If
	Next
Next

print "Verify Claim Data: finalizing result"
For k = 0 To Ubound(arrResult)
	If arrResult(k) = Empty Then
			arrResult(k) = True
	End If
	If arrResult(k) = False Then
		bVal = False
		Exit For
	End If
	bVal = True
Next

If bVal2 Then
	If bVal Then
		Parameter("bResult") = True
	Else
		Parameter("bResult") = False
	End If
Else
	ExitAction
End If




