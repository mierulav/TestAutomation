OPTION EXPLICIT @@ script infofile_;_ZIP::ssf3.xml_;_

' 1. Test Data Path
Dim TestList : TestList = TestDataDir + "\Validations\" & Environment("TestName") & ".xls"

' 2. Test Data Information
Dim TestDataCX : TestDataCX = "TestDataCX"
Dim TestDataSAP : TestDataSAP = "TestDataSAP"
Datatable.AddSheet TestDataCX
Datatable.AddSheet	TestDataSAP
Datatable.ImportSheet TestList, "CX", TestDataCX
Datatable.ImportSheet TestList, "SAP", TestDataSAP

'3. Initialization
Dim x, y, i
Dim CXDataSheet : Set CXDataSheet = Datatable.GetSheet(TestDataCX)
Dim SAPDataSheet : Set SAPDataSheet = Datatable.GetSheet(TestDataSAP)
Dim strOrderNumber, ConnectURL
Dim strRole, strUsername, strPassword, strSoldToCode, strShipToCode, strShipToAddress, strProductCode, strProductName, _
strProductQuantity, strDeliveryInstruction, strPONumber, intMinimimumPurchase
Dim strSAPPONumber, strSAPDI, strDeliveryOrderNumber, strInvoiceNumber, strShipmentNumber, strCustomerReceiptDate, _
DateAddDiff, ShipmentType, TransportPlanningPt, WarehouseNo, ConfirmedReceipt, Shipment, Invoicing, _
PGI, TransferOrder, ItemPicking, DeliverOrder, Release, FwdAgent

'CX Test Data
For i = 1 To CXDataSheet.GetRowCount
	CXDataSheet.SetCurrentRow(i)
	If UCase(Datatable.Value("ToTest", TestDataCX)) = "Y" Then'and UCase(Datatable.Value("Market", TestDataCX)) = strMarket Then
		
		'ConnectURL = SystemURL
		ProjectName = Datatable.Value("Market", TestDataCX)
		strRole = "" 'if this is a test for secondary normal account, please input "normal" 
		strUsername = Datatable.Value("Username", TestDataCX)
		strPassword = Datatable.Value("Password", TestDataCX)
		strSoldToCode = Datatable.Value("SoldToCode", TestDataCX)
		strShipToCode = Datatable.Value("SoldToCode", TestDataCX)
		strShipToAddress = Datatable.Value("ShipToAddress", TestDataCX)
		strProductCode = Datatable.Value("SKU", TestDataCX)
		strProductName = Datatable.Value("ProductName", TestDataCX)
		strProductQuantity = Datatable.Value("ProductQuantity", TestDataCX)
		strDeliveryInstruction = Datatable.Value("DeliveryInstruction", TestDataCX)
		strPONumber = Datatable.Value("PONumber", TestDataCX)
		intMinimimumPurchase = Datatable.Value("MinimumPurchase", TestDataCX)
	
	End If
Next

'SAP Test Data
For i = 1 To SAPDataSheet.GetRowCount
	SAPDataSheet.SetCurrentRow(i)
	If UCase(Datatable.Value("ToTest", TestDataSAP)) = "Y" and UCase(Datatable.Value("Market", TestDataSAP)) = ProjectName Then
	
		Release = Datatable.Value("ReleaseCredit", TestDataSAP)
		DeliverOrder = Datatable.Value("DeliverOrder", TestDataSAP)
		ItemPicking = Datatable.Value("ItemPicking", TestDataSAP)
		TransferOrder = Datatable.Value("TransferOrder", TestDataSAP)
		PGI = Datatable.Value("PostGoodsIssue", TestDataSAP)
		Invoicing = Datatable.Value("Invoicing", TestDataSAP)
		Shipment = Datatable.Value("Shipment", TestDataSAP)
		ConfirmedReceipt = Datatable.Value("ConfirmedReceipt", TestDataSAP)
		WarehouseNo = Datatable.Value("WarehouseNo", TestDataSAP)
		TransportPlanningPt = Datatable.Value("TransportPlanningPt", TestDataSAP)
		ShipmentType = Datatable.Value("ShipmentType", TestDataSAP)
		FwdAgent = Datatable.Value("ForwardingAgent", TestDataSAP)
		DateAddDiff = Datatable.Value("DateAdd", TestDataSAP)
	
	End If
Next

'Step 1: Order creation
LoginIntoConnect
strOrderNumber = OrderSubmission
Select Case strOrderNumber
	Case "0"
		AssertExitRun "Unable to proceed", "Order Submission is unsuccessfull"
		ExitAction
	 
	 Case "1"
	 	AssertExitRun "Unable to proceed", "Cart Error"
	 	ExitAction
	 Case Else
		OrderTracking "Order Received", strOrderNumber
End Select

'Step 2: Order Fulfilment
OrderFulfilment

'Step 3: Logout
LogoutAndCloseBrowser

'5 Export test into testresult
Datatable.ExportSheet TestResultDir & "\" & GetStringDate & "Order submissions.xls", TestDataCX, "E2E" & ProjectName

'Login subs
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
	
	'Validate Login screen
	AssertObjects ProjectName & " - Login Objects", LoginObjects
	
	'2. Login as an existing account member
	Login strUsername, strPassword
	
	'Incase
	Browser("DKSH Connect").Navigate ConnectURL
	
End Sub

'Functions Operations
Function OrderSubmission()
	
	'Step 3: Select ShipToId
	SelectShipToAddress(strShipToCode)
	
	'Validate landing page - Header objects
	AssertObjects ProjectName & "- Landing page - Header", CheckHeaderObjects
	
	'Validate landing page - Footer objects
	AssertObjects ProjectName & "-Landing page - Footer", CheckFooterObjects
	
	'Validate landing page - Navigation Menu objects
	AssertObjects ProjectName & "- Landing page - Navigation Menu", CheckNavigationMenuObjects
	
	'Validate landing page - User Menu objects
	AssertObjects ProjectName & "- Landing page - User Menu", CheckUserMenuList(strRole)
	
	'Step 4: Navigate to All Products
	OpenAllProductPage
	
	'Validate PLP screen
	AssertObjects ProjectName & "- PLP page", CheckPLPObjects
	
	'Validate PLP item objects
	AssertObjects ProjectName & "- PLP page - Item objects", CheckPLPProductObjects
	
	'Step 5: Search products
	SearchProduct strProductCode
	
	'Step 6: Open product's PDP
	OpenProductPDP
	
	'Validate PDP screen
	AssertObjects ProjectName & "- PDP page", CheckPDPObjects
	
	'Validate Produt details
	Assert ProjectName & "- Product Code & Product Name", CheckPDPSelectedProduct(strProductCode, strProductName)
	
	'Step 7: Add product to Cart
	AddProductAndGoToCart
	
	'Error Handling for product that is not able to put into cart
	If CheckSpecificProductCode(strProductCode) = False Then
		OrderSubmission = "1"
		AssertExitRun ProjectName & "-  Step 7: Add Product to Cart", "Product did not succesfully added to cart" 
		ExitAction
	End If
	
	'Validate Cart oage
	AssertObjects ProjectName & "- Shopping Cats - Objects", CheckCartsObjects
	
	'Validate Carts calculation summary objects
	AssertObjects ProjectName & "- Shopping Carts - Calculation Summary Objects", CheckCartsCalculationObjects
	
	'Validate Carts product details
	Assert ProjectName & "- Shopping Carts - Added Product's details", CheckProductDetails(strProductCode, strProductName)
	
	'Validate Carts summary calculation (based on 1 item added calculation)
	Assert ProjectName & "- Shopping Carts - Calculation Summary", CheckCartSummaryCalculation
	
	'Validate Minimum Purchase alert
	Assert ProjectName & "- Shopping Carts - Minimum Purchase Alert", MinimumPurchaseAlert
	
	'Validate Carts product updated alert
	SetProductQuantity "10"
	Assert ProjectName & "- Shopping Carts - Product Quantity Alert", ProductQuantityUpdatedAlert
	
	'Validate Total Product Total Price (Quantity * Price per unit)
	SetProductQuantity strProductQuantity
	Assert ProjectName & "- Shopping Carts - Product Total Price", CheckProductTotalPrice
	
	'Step 8: Proceed Checkout
	ProceedForCheckout
	
	'Step 8.1: For MYHEC to cater for unmapped payer code
	If ProjectName = "MYHEC" Then
		COUseThisPayer.Click
	End If
	
	'Validate Checkout objects
	AssertObjects ProjectName & "- Checkout - Layout", CheckCheckoutObjects
	
	'Validate Checkout Calculation Summary objects
	AssertObjects ProjectName & "- Checkout - Calculation Summary Objects", CheckCheckoutCalculationSummaryObjects
	
	'Validate Product details 
	Assert ProjectName & "- Checkout - Product Details", CheckOrderDetails(strProductCode, strProductName, Trim(Replace(strShipToAddress, ",", "")))
	
	'Validate Calculation Summary
	Assert ProjectName & "- Checkout - Calculation Summary", Checkout_CheckCalculationSummary
	
	'Step 9: Submit order
	SetDeliveryInstruction strDeliveryInstruction
	
	Select Case ProjectName
		Case "VNHEC"
		
		Case "MMHEC"
			SetPONumber strPONumber
			SelectOOSProceedingAgreement "agree"
			
		Case Else
			SetPONumber strPONumber
	End Select
		
	SubmitOrder
	
	'Validate Order Confirmation Page
	If CheckSalesOrderConfirmed and GetOrderNumber <> False Then
		OrderSubmission = GetOrderNumber
	Else
		OrderSubmission = "0"
	End If 
	
End Function

Function OrderFulfilment()

	'Back to initial SAP screen
	SAPEasyAccessScreen
	
	'Release credit block if any
	If UCase(Release) = "Y" Then
		ReleaseCredit strOrderNumber
		SAPGuiSession("Session").SAPGuiWindow("SAP Easy Access  -  User").Maximize
	End If
	
	'Validations PO Number
	Assert ProjectName & "- Validate PO Number", IsEqual(GetPONumber(strOrderNumber), strPONumber)
	
	'Validation Delivery Instruction
	Assert ProjectName & "- Validate Delivery Instructions", IsEqual(GetDeliveryInstruction(strOrderNumber, ProjectName), strDeliveryInstruction)
		
	'Step 4: Create Delivery order
	If UCase(DeliverOrder)= "Y" Then
		strDeliveryOrderNumber = CreateDeliveryOrder(strOrderNumber)
		'Check order status = Order in Process for Delivery in Track page, and Delivery order number generated
		OrderTracking "Order in Process - Delivery", strDeliveryOrderNumber
	End If
	
	'Step 5: Transfer Order
	If UCase(TransferOrder) = "Y" Then
		CreateConfirmTransferOrder strDeliveryOrderNumber
	End If	
	
	'Step 5(1): Item Piking
	If UCase(ItemPicking) = "Y" Then
		Picking strDeliveryOrderNumber
	End If
	
	'Step 5(2): Post Goods Issue
	If UCase(PGI) = "Y" Then
		PostGoodsIssue
	End If
				
	'Step 6: Create Invoice 
	If UCase(Invoicing) = "Y" Then
		strInvoiceNumber = CreateInvoice(strDeliveryOrderNumber)
		'Check order status = Order in Process for Invoice in Track page, and Delivery order number generated
		OrderTracking "Order in Process - Invoice", strInvoiceNumber
	End If
	
	'Step 7: Create Shipment
	If UCase(Shipment) = "Y" Then
		strShipmentNumber = CreateShipment(strDeliveryOrderNumber, TransportPlanningPt, ShipmentType, FwdAgent)
		'Step 5.5: Check order status = Deliver in Transit in Track page, and Shipment number generated
		OrderTracking "Deliver in Transit", strShipmentNumber
	End If
	
	'Step 8: Create Customer Receipt
	If UCase(ConfirmedReceipt) = "Y" Then
		strCustomerReceiptDate = CreateCustomerConfirmationReceipt(strDeliveryOrderNumber)
		'Get acual date for specific country, add or minus minutes
		strCustomerReceiptDate = DateAdd("n", CSng(DateAddDiff), Replace(strCustomerReceiptDate, ".", "/"))
		'Check order status = Customer Confirmed Receipt and Date posted
		OrderTracking "Customer Confirmed Receipt", strCustomerReceiptDate
	End If
	
End Function

Sub OrderTracking(strOrderStatus, strOrderTrackingNumber)

	'Step 1: Open Track & Order page
	OpenTrackOrderPage
	
	'Validate Track & Order Page
	AssertObjects ProjectName & "- Track & Order Layouts", CheckOrderObjects
	
	'Validate Track & Order Page - Order Card
	AssertObjects ProjectName & "- Track & Order - Order Card Objects", CheckOrderCardObjects
	
	'Step 2: Search & Validate Order Number
	Assert ProjectName & "- Track & Order - Search Sales Order Number", SearchOrder(strOrderNumber) 
	
	'Open Order details
	ShowOrderDetails
	
	'Validate Track & Order details layout
	AssertObjects ProjectName & "- Track & Order Details - Layout", CheckOrderDetailsObjects
	
	'Validate Track & Order details layout
	AssertObjects ProjectName & "- Track & Order Details - Calculation Summary Objects", CheckOrderCalculationSummaryObjects
	
	'Validate Order Status
	Assert ProjectName & "- Track & Order Details - Sales Order Status", CheckOrderStatus(strOrderStatus)
	
	'Validate Order Tracking
	Assert ProjectName & "- Track & Order Details - Track Sales Order Status " & strOrderStatus, CheckOrderTracking(strOrderStatus, strOrderTrackingNumber)
	
	'Validate Order Tracking Calculation Summary
	Assert ProjectName & "- Track & Order Details - Calculation Summary", Track_CheckCalculationSummary
	
End Sub @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf27.xml_;_

 @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf57.xml_;_
