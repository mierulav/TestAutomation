Option Explicit

' 1. Test Data Path
Dim TestList : TestList = TestDataDir + "\Validations\" & Environment.Value("TestName") & ".xls"
Datatable.ImportSheet TestList, "TestData", "Global"
ProjectName = "TWHEC"

'Get Test Data
Dim i
For i = 1 To Datatable.GetSheet("Global").GetRowCount
	Datatable.GetSheet("Global").SetCurrentRow(i)
	If UCase(Datatable.Value("ToTest", "Global")) = "Y" and UCase( Datatable.Value("Market", "Global")) = ProjectName Then
		ReorderSubmission	
		'''LogoutAndCloseBrowser
	End If
Next


Sub ReorderSubmission()
	'1. Launch the Connect Market URL 
	If UCase(ProjectName) <> "AUTEC" Then
		SystemURL = SystemURL + LCase(ProjectName) + "/en"
	Else
		SystemURL = SystemURL + "connect/en"
	End If
	
	If Not Browser("Creationtime:=0").Exist Then
		SystemUtil.Run DefaultBrowser, SystemURL
	End If
	
	'2. Login as an existing account member, and select shipto
	Login Datatable.Value("Username", "Global"), Datatable.Value("Password", "Global")
	SelectShipToDefault
	
	'3. Navigate to track and order
	OpenTrackOrderPage
	
	'4. Check user should be able to view previously placed orders.
	AssertObjects "Track & Order - Order Card Objects", CheckOrderCardObjects
	
	'5. User clicks on show more button on order listing page.
	SearchOrder(Datatable.Value("SalesOrderNumber", "Global")) 
	ShowOrderDetails
	
	'6. User is able to select products for orders with multiple products from original order.
	Dim arrProductCodes : arrProductCodes = GetAllProductsCodeInOrderDetails
	ReorderAllProducts
	Dim i
	For i = 0 To Ubound(arrProductCodes)
		Assert "Validate that all products successfully imported into cart", CheckSpecificProductCode(arrProductCodes(i))
		print arrProductCodes(i)
	Next
	
	'8.User clicks on Proceed to check out.
	ProceedForCheckout
	
	'Step 8.1: For MYHEC to cater for unmapped payer code
	If ProjectName = "MYHEC" Then
		COUseThisPayer.Click
	End If
	
	'10. User makes changes to delivery instruction and P.O number in check out page.
	'11. User clicks on Place order button.
	SetDeliveryInstruction "This is Reorder for " & ProjectName
	
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
		AssertExitRun "Order Submission", "Unsuccessful order submission"
	End If 
	
	'Validate Order sent to ERP
	SAPEasyAccessScreen
	If  GetDeliveryInstruction(strOrderNumber, ProjectName) =  "This is Reorder for " & ProjectName Then
		Dim tempRes : tempRes = true
	End If
	Assert "Check Order Created in SAP", tempRes
	
End Sub
