OPTION EXPLICIT @@ script infofile_;_ZIP::ssf3.xml_;_

' 1. Procedural data
Dim TestList : TestList = TestDataDir + "\Validations\Order submissions.xls"

' 2. Test Data Information
Datatable.AddSheet "TestList"
Datatable.ImportSheet TestList, "TestData", "TestList"
Dim x, y, i
Dim strOrderNumber
Dim strRole, strCustomerId, strPassword,strShipToName, strShipToAddress, strProductCode, strProductName, strProductQuantity, strDeliveryInstruction, strPONumber, intMinimimumPurchase, ConnectURL

' 2. Order submission
For i = 1 To Datatable.GetSheet("TestList").GetRowCount
	Datatable.GetSheet("TestList").SetCurrentRow(i)
	If UCase(Datatable.Value("ToTest", "TestList")) = "Y" Then
		
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
		
		LogoutAndCloseBrowser		
		
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
	
	If Not Browser("Creationtime:=0").Exist Then
		SystemUtil.Run DefaultBrowser, ConnectURL
	End If

	'Validate Login screen
	AssertObjects "Login Objects", LoginObjects
	
	'Step 2: Login
	Login strCustomerId, strPassword
	Browser("DKSH Connect").Navigate ConnectURL
	
	'Step 3: Select ShipToId
	If Datatable.Value("MultipleShipToAddress", "TestList") = "Y" Then
		SelectShipToAddress(strShipToName)
	End If
	
	'Validate landing page - Header objects
	AssertObjects "Landing page - Header", CheckHeaderObjects
	
	'Validate landing page - Footer objects
	AssertObjects "Landing page - Footer", CheckFooterObjects
	
	'Validate landing page - Navigation Menu objects
	AssertObjects "Landing page - Navigation Menu", CheckNavigationMenuObjects
	
	'Validate landing page - User Menu objects
	AssertObjects "Landing page - User Menu", CheckUserMenuList(strRole)
	
	'Step 4: Navigate to All Products
	OpenAllProductPage
	
	'Validate PLP screen
	AssertObjects "PLP page", CheckPLPObjects
	
	'Validate PLP item objects
	AssertObjects "PLP page - Item objects", CheckPLPProductObjects
	
	'Step 5: Search products
	SearchProduct strProductCode
	
	'Step 6: Open product's PDP
	OpenProductPDP
	
	'Validate PDP screen
	AssertObjects "PDP page", CheckPDPObjects
	
	'Validate Produt details
	Assert "Product Code & Product Name", CheckPDPSelectedProduct(strProductCode, strProductName)
	
	'Step 7: Add product to Cart
	AddProductAndGoToCart
	
	'Error Handling for product that is not able to put into cart
	If CheckSpecificProductCode(strProductCode) = False Then
		OrderSubmission = "1"
		AssertExitRun "Step 7: Add Product to Cart", "Product did not succesfully added to cart" 
		Exit Function
	End If
	
	'Validate Cart oage
	AssertObjects "Shopping Cats - Objects", CheckCartsObjects
	
	'Validate Carts calculation summary objects
	AssertObjects "Shopping Carts - Calculation Summary Objects", CheckCartsCalculationObjects
	
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
	AssertObjects "Checkout - Layout", CheckCheckoutObjects
	
	'Validate Checkout Calculation Summary objects
	AssertObjects "Checkout - Calculation Summary Objects", CheckCheckoutCalculationSummaryObjects
	
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
			'SelectOOSProceedingAgreement "agree"
			
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
