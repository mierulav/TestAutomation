Option Explicit

' 1. Test Data Path
Dim TestList : TestList = TestDataDir + "\Validations\" & Environment.Value("TestName") & ".xls"
Datatable.ImportSheet TestList, "TestData", "Global"

'Get Test Data
Dim i
For i = 1 To Datatable.GetSheet("Global").GetRowCount
	Datatable.GetSheet("Global").SetCurrentRow(i)
	If UCase(Datatable.Value("ToTest", "Global")) = "Y" Then 
		ProjectName = Datatable.Value("Market", "Global")
		LoginIntoConnect
		ReorderSubmission	
		LogoutAndCloseBrowser
	End If
Next

Sub LoginIntoConnect()

	Dim ConnectURL
	'1. Launch the Connect Market URL 
	If UCase(ProjectName) <> "AUTEC" Then
		ConnectURL = SystemURL + LCase(ProjectName) + "/en"
	Else
		ConnectURL = SystemURL + "connect/en"
	End If
	
	If Not Browser("Creationtime:=0").Exist Then
		SystemUtil.Run DefaultBrowser, ConnectURL
	End If
	
	'2. Login as an existing account member, and select shipto
	Login Datatable.Value("Username", "Global"), Datatable.Value("Password", "Global")
	SelectShipToDefault
	
End Sub

Sub ReorderSubmission()
	
	'3. Navigate to track and order
	OpenTrackOrderPage
	
	'4. Check user should be able to view previously placed orders.
	AssertObjects ProjectName & ": Track & Order - Order Card Objects", CheckOrderCardObjects
	
	'5. User clicks on show more button on order listing page.
	SearchOrder(Datatable.Value("SalesOrderNumber", "Global")) 
	ShowOrderDetails
	
	'6. User is able to select products for orders with multiple products from original order.
	Dim arrProductCodes : arrProductCodes = GetAllProductsCodeInOrderDetails
	ReorderAllProducts
	Dim i
	For i = 0 To Ubound(arrProductCodes)
		Dim blnRes : blnRes = CheckSpecificProductCode(arrProductCodes(i))
		Assert ProjectName & ": Validate that all products successfully imported into cart", blnRes
		If blnRes = False Then
			AssertExitRun ProjectName & ": Import product into cart", "Unsuccessful import product into cart !"
		End If
	Next
	
	If GetProductUnitPrice = 0 Then
		Assert ProjectName & " - Product price is 0 (Either OOS or Pricing issue)", False
		Exit Sub
	End If
	
	'8.User clicks on Proceed to check out.
	ProceedForCheckout
	
	'Step 8.1: For MYHEC to cater for unmapped payer code
	If ProjectName = "MYHEC" Then
		COUseThisPayer.Click
	End If
	
	'10. User makes changes to delivery instruction and P.O number in check out page.
	'11. User clicks on Place order button.
	SetDeliveryInstruction ProjectName & ": This is Reorder for " & ProjectName
	
	Select Case ProjectName
		Case "VNHEC"
		
		Case "MMHEC"
			''SetPONumber strPONumber
			SelectOOSProceedingAgreement "agree"
			
		Case Else
			SetPONumber "AutomationTest"
	End Select
		
	SubmitOrder
	
	'Validate Order Confirmation Page
	If CheckSalesOrderConfirmed Then
		Dim strOrderNumber : strOrderNumber = GetOrderNumber
	Else
		AssertExitRun ProjectName & ": Order Submission", "Unsuccessful order submission"
	End If 
	
	'Validate Order sent to ERP
	SAPEasyAccessScreen
	If  Instr(GetDeliveryInstruction(strOrderNumber, ProjectName), "This is Reorder for " & ProjectName) > 0 Then
		Dim tempRes : tempRes = true
	End If
	Assert ProjectName & ": Check Order Created in SAP", tempRes
	
End Sub
