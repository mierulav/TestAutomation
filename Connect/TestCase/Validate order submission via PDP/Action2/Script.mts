OPTION EXPLICIT

' 1. Procedural data
Dim TestList : TestList = TestDataDir + "\Validations\Order submissions.xls"
Dim TestListSAP : TestListSAP = TestDataDir + "\Validations\Order fulfilments.xls"

' 2. CX Test Data Information
Datatable.AddSheet "TestList"
Datatable.ImportSheet TestList, "TestData", "TestList"
Dim x, y, i
Dim strOrderNumber
Dim strRole, strCustomerId, strPassword,strShipToName, strShipToAddress, strProductCode, strProductName, strProductQuantity, strDeliveryInstruction, strPONumber, intMinimimumPurchase, ConnectURL

'2.1 ERP Test Data Information
Datatable.AddSheet "TestListSAP"
Datatable.ImportSheet TestListSAP, "TestDataSAP", "TestListSAP"
Dim k, j, z, strSAPPONumber, strSAPDI, strDeliveryOrderNumber, strInvoiceNumber, strShipmentNumber, strCustomerReceiptDate

' 3. Order submission
For i = 1 To Datatable.GetSheet("TestList").GetRowCount
	Datatable.GetSheet("TestList").SetCurrentRow(i)
	If UCase(Datatable.Value("ProjectName", "TestList")) = ProjectName Then
		
		ProjectName = Datatable.Value("ProjectName", "TestList")
		ConnectURL = SystemURL
		strRole = "" 'if this is a test for secondary normal account, please input "normal" 
		strCustomerId = Datatable.Value("Username", "TestList")
		strPassword = Datatable.Value("Password", "TestList")
		strShipToName = Datatable.Value("ShipToName", "TestList")
		strShipToAddress = Datatable.Value("ShipToAddress", "TestList")
		strProductCode = Datatable.Value("ProductCode", "TestList")
		strProductName = Datatable.Value("ProductName", "TestList")
		strProductQuantity = Datatable.Value("ProductQuantity", "TestList")
		strDeliveryInstruction = Datatable.Value("DeliveryInstruction", "TestList")
		strPONumber = Datatable.Value("PONumber", "TestList")
		intMinimimumPurchase = Datatable.Value("MinimumPurchase", "TestList")
		
		strOrderNumber = OrderSubmission	
		
		Select Case strOrderNumber
			Case "0"
				Datatable.Value("SOCreated", "TestList") = "Submission error"
			 
			 Case "1"
			 	Datatable.Value("SOCreated", "TestList") = "Cart error"
			 	
			 Case Else
			 	Datatable.Value("SOCreated", "TestList") = strOrderNumber
				OrderTracking "Order Received", strOrderNumber
		End Select	
		
	End If
Next

'4. Order Fulfilment
For k = 1 To Datatable.GetSheet("TestListSAP").GetRowCount
	Datatable.GetSheet("TestListSAP").SetCurrentRow(i)
	
	If UCase(Datatable.Value("ProjectName", "TestList")) = ProjectName Then
		ProjectName = Datatable.Value("ProjectName", "TestList")
		OrderFulfilment	
	End If	
Next

'5 Export test into testresult
Datatable.ExportSheet TestResultDir + "\" + GetStringDate + "_PDP_" + "Order submissions.xls", "TestList", "TestList"

'Functions Operations
Function OrderSubmission()

	'Step 1: Navigate to the system
	If ProjectName <> "AUTEC" Then
		ConnectURL = ConnectURL + LCase(ProjectName) + "/en"
	Else
		ConnectURL = ConnectURL + "connect/en"
	End If
	
	SystemUtil.Run DefaultBrowser, ConnectURL
	
	'Validate Login screen
	Assert "Login Objects", LoginObjects
	
	'Step 2: Login
	Login strCustomerId, strPassword
	Browser("DKSH Connect").Navigate SystemURL
	
	'Step 3: Select ShipToId
	If Datatable.Value("MultipleShipToAddress", "TestList") = "Y" Then
		SelectShipToAddress(strShipToName)
	End If
	
	'Validate landing page - Header objects
	Assert "Landing page - Header", CheckHeaderObjects
	
	'Validate landing page - Footer objects
	Assert "Landing page - Footer", CheckFooterObjects
	
	'Validate landing page - Navigation Menu objects
	Assert "Landing page - Navigation Menu", CheckNavigationMenuObjects
	
	'Validate landing page - User Menu objects
	Assert "Landing page - User Menu", CheckUserMenuList(strRole)
	
	'Step 4: Navigate to All Products
	OpenAllProductPage
	
	'Validate PLP screen
	Assert "PLP page", CheckPLPObjects
	
	'Validate PLP item objects
	Assert "PLP page - Item objects", CheckPLPProductObjects
	
	'Step 5: Search products
	SearchProduct strProductCode
	
	'Step 6: Open product's PDP
	OpenProductPDP
	
	'Validate PDP screen
	Assert "PDP page", CheckPDPObjects
	
	'Validate Produt details
	Assert "Product Code & Product Name", CheckPDPSelectedProduct(strProductCode, strProductName)
	
	'Step 7: Add product to Cart
	AddProductAndGoToCart
	
	'Error Handling for product that is not able to put into cart
	If CheckSpecificProductCode(strProductCode) = False Then
		OrderSubmission = "1"
		Exit Function
	End If
	
	'Validate Cart oage
	Assert "Shopping Cats - Objects", CheckCartsObjects
	
	'Validate Carts calculation summary objects
	Assert "Shopping Carts - Calculation Summary Objects", CheckCartsCalculationObjects
	
	'Validate Carts product details
	Assert "Shopping Carts - Added Product's details", CheckProductDetails(strProductCode, strProductName)
	
	'Validate Carts summary calculation (based on 1 item added calculation)
	Assert "Shopping Carts - Calculation Summary", CheckCartSummaryCalculation
	
	'Validate Minimum Purchase alert
	Assert "Shopping Carts - Minimum Purchase Alert", MinimumPurchaseAlert
	
	'Validate Carts product updated alert
	SetProductQuantity "10"
	Assert "Shopping Carts - Product Quantity Alert", ProductQuantityUpdatedAlert
	
	'Validate Total Product Total Price (Quantity * Price per unit)
	SetProductQuantity strProductQuantity
	Assert "Shopping Carts - Product Total Price", CheckProductTotalPrice
	
	'Step 8: Proceed Checkout
	ProceedForCheckout
	
	'Step 8.1: For MYHEC to cater for unmapped payer code
	If ProjectName = "MYHEC" Then
		COUseThisPayer.Click
	End If
	
	'Validate Checkout objects
	Assert "Checkout - Layout", CheckCheckoutObjects
	
	'Validate Checkout Calculation Summary objects
	Assert "Checkout - Calculation Summary Objects", CheckCheckoutCalculationSummaryObjects
	
	'Validate Product details 
	Assert "Checkout - Product Details", CheckOrderDetails(strProductCode, strProductName, strShipToAddress)
	
	'Validate Calculation Summary
	Assert "Checkout - Calculation Summary", Checkout_CheckCalculationSummary
	
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
	If CheckSalesOrderConfirmed Then
		OrderSubmission = GetOrderNumber
	Else
		OrderSubmission = "0"
	End If 
	
End Function

Function OrderFulfilment()

	'Validations PO Number
	strSAPPONumber = GetPONumber(strOrderNumber)
	Assert "Validate PO Number", IsEqual(strSAPPONumber, Datatable.Value("PONumber", "TestList"))
	
	'Validation Delivery Instruction
	strSAPDI = GetDeliveryInstruction(strOrderNumber)
	Assert "Validate Delivery Instructions", IsEqual(strSAPDI, Datatable.Value("DeliveryInstruction", "TestList"))
	
	'Release credit block if any
	If Datatable.Value("ReleaseCredit", "TestList") = "Y" Then
		ReleaseCredit strOrderNumber
		SAPGuiSession("Session").SAPGuiWindow("SAP Easy Access  -  User").Maximize
	End If
	
	'Step 4: Create Delivery order
	If Datatable.Value("DeliverOrder", "TestList") = "Y" Then
		strDeliveryOrderNumber = CreateDeliveryOrder(strOrderNumber)
		'Check order status = Order in Process for Delivery in Track page, and Delivery order number generated
		OrderTracking "Order in Process - Delivery", strDeliveryOrderNumber
	End If
	
	'Validations PO Number
	strSAPPONumber = GetPONumber(strOrderNumber)
	Assert "Validate PO Number", IsEqual(strSAPPONumber, Datatable.Value("PONumber", "TestList"))
	
	'Validation Delivery Instruction
	strSAPDI = GetDeliveryInstruction(strOrderNumber)
	Assert "Validate Delivery Instructions", IsEqual(strSAPDI, Datatable.Value("DeliveryInstruction", "TestList"))
	
	'Step 5: Transfer Order
	If Datatable.Value("TransferOrder", "TestList") = "Y" Then
		CreateConfirmTransferOrder strDeliveryOrderNumber
	End If	
	
	'Step 5(1): Item Piking
	If Datatable.value("ItemPicking", "TestList") = "Y" Then
		Picking strDeliveryOrderNumber
	End If
	
	'Step 5(2): Post Goods Issue
	If Datatable.Value("PostGoodsIssue", "TestList") = "Y" Then
		PostGoodsIssue
	End If
				
	'Step 6: Create Invoice 
	If Datatable.Value("Invoicing", "TestList") = "Y" Then
		strInvoiceNumber = CreateInvoice(strDeliveryOrderNumber)
		'Check order status = Order in Process for Invoice in Track page, and Delivery order number generated
		OrderTracking "Order in Process - Invoice", strInvoiceNumber
	End If
	
	'Step 7: Create Shipment
	If Datatable.Value("Shipment", "TestList") = "Y" Then
		strShipmentNumber = CreateShipment(strDeliveryOrderNumber, Datatable.Value("TransportPlanningPt", "TestList"), Datatable.Value("ShipmentType", "TestList"))
		'Step 5.5: Check order status = Deliver in Transit in Track page, and Shipment number generated
		OrderTracking "Deliver in Transit", strShipmentNumber
	End If
	
	'Step 8: Create Customer Receipt
	If Datatable.Value("ConfirmedReceipt", "TestList") = "Y" Then
		strCustomerReceiptDate = CreateCustomerConfirmationReceipt(strDeliveryOrderNumber)
		'Get acual date for specific country, add or minus minutes
		strCustomerReceiptDate = DateAdd("n", CSng(Datatable.Value("DateAdd", "TestList")), Replace(strCustomerReceiptDate, ".", "/"))
		'Check order status = Customer Confirmed Receipt and Date posted
		OrderTracking "Customer Confirmed Receipt", strCustomerReceiptDate
	End If
	
	
End Function

Sub OrderTracking(strOrderStatus, strOrderTrackingNumber)

	'Step 1: Open Track & Order page
	OpenTrackOrderPage
	
	'Validate Track & Order Page
	Assert "Track & Order Layouts", CheckOrderObjects
	
	'Validate Track & Order Page - Order Card
	Assert "Track & Order - Order Card Objects", CheckOrderCardObjects
	
	'Step 2: Search & Validate Order Number
	Assert "Track & Order - Search Sales Order Number", SearchOrder(strOrderNumber) 
	
	'Open Order details
	ShowOrderDetails
	
	'Validate Track & Order details layout
	Assert "Track & Order Details - Layout", CheckOrderDetailsObjects
	
	'Validate Track & Order details layout
	Assert "Track & Order Details - Calculation Summary Objects", CheckOrderCalculationSummaryObjects
	
	'Validate Order Status
	Assert "Track & Order Details - Sales Order Status", CheckOrderStatus(strOrderStatus)
	
	'Validate Order Tracking
	Assert "Track & Order Details - Track Sales Order Status " & strOrderStatus, CheckOrderTracking(strOrderStatus, strOrderTrackingNumber)
	
	'Validate Order Tracking Calculation Summary
	Assert "Track & Order Details - Calculation Summary", Track_CheckCalculationSummary
	
End Sub

