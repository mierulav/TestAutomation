OPTION EXPLICIT @@ script infofile_;_ZIP::ssf3.xml_;_

' 1. Procedural data
Dim TestList : TestList = TestDataDir + "\Validations\Order fulfilments.xls"

' 2. Test Data Information
Datatable.AddSheet "TestList"
Datatable.ImportSheet TestList, "TestDataSAP", "TestList"
Dim x, y, i, ConnectURL, strOrderNumber, strSAPPONumber, strSAPDI, strDeliveryOrderNumber, strInvoiceNumber, strShipmentNumber, strCustomerReceiptDate


' 2. Order fulfilment
For i = 1 To Datatable.GetSheet("TestList").GetRowCount
	Datatable.GetSheet("TestList").SetCurrentRow(i)
	ConnectURL = SystemURL
	
	If UCase(Datatable.Value("ToTest", "TestList")) = "Y" Then
		ProjectName = Datatable.Value("ProjectName", "TestList")
		
		If ProjectName <> "AUTEC" Then
			ConnectURL = ConnectURL + LCase(ProjectName) + "/en"
		Else
			ConnectURL = ConnectURL + "connect/en"
		End If
		
		If Not Browser("Creationtime:=0").Exist Then
			SystemUtil.Run DefaultBrowser, ConnectURL
		Else
			Browser("Creationtime:=0").Navigate ConnectURL
		End If
		
		Login Datatable.Value("CustomerID", "TestList"), Datatable.Value("Password", "TestList")
		Browser("DKSH Connect").Navigate ConnectURL
		
		If Datatable.Value("MultipleAddress", "TestList") = "Y" Then
			SelectShipToAddress(Datatable.Value("ShipToName", "TestList"))
		End If
		
		OpenTrackOrderPage
		
		strOrderNumber = Datatable.Value("SONumber", "TestList") 
		OrderFulfilment	
		LogoutAndCloseBrowser
		
	End If	
Next

'5 Export test into testresult
Datatable.ExportSheet TestResultDir + "\" + GetStringDate + "_" + "Order fulfilments.xls", "TestList", "TestList"

'Functions Operations
Function OrderFulfilment()

	'Back to initial SAP screen
	SAPEasyAccessScreen
	
	'Release credit block if any
	If Datatable.Value("ReleaseCredit", "TestList") = "Y" Then
		ReleaseCredit strOrderNumber
		SAPGuiSession("Session").SAPGuiWindow("SAP Easy Access  -  User").Maximize @@ hightlight id_;_0_;_script infofile_;_ZIP::ssf17.xml_;_
	End If
	
	'Validations PO Number
	'strSAPPONumber = GetPONumber(strOrderNumber)
	Assert "Validate PO Number", IsEqual(GetPONumber(strOrderNumber), Datatable.Value("PONumber", "TestList"))
	
	'Validation Delivery Instruction
	'strSAPDI = GetDeliveryInstruction(strOrderNumber)
	Assert "Validate Delivery Instructions", IsEqual(GetDeliveryInstruction(strOrderNumber, ProjectName), Datatable.Value("DeliveryInstruction", "TestList"))
		
	'Step 4: Create Delivery order
	If Datatable.Value("DeliverOrder", "TestList") = "Y" Then
		strDeliveryOrderNumber = CreateDeliveryOrder(strOrderNumber)
		'Check order status = Order in Process for Delivery in Track page, and Delivery order number generated
		OrderTracking "Order in Process - Delivery", strDeliveryOrderNumber
	End If
	
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
	
	'Step 2: Search & Validate Order Number
	Assert "Track & Order - Search Sales Order Number", SearchOrder(strOrderNumber) 
	
	'Open Order details
	ShowOrderDetails
	
	'Validate Track & Order details layout
	Assert "Track & Order Details - Calculation Summary Objects", CheckOrderCalculationSummaryObjects
	
	'Validate Order Status
	Assert "Track & Order Details - Sales Order Status", CheckOrderStatus(strOrderStatus)
	
	'Validate Order Tracking
	Assert "Track & Order Details - Track Sales Order Status " & strOrderStatus, CheckOrderTracking(strOrderStatus, strOrderTrackingNumber)
	
	'Validate Order Tracking Calculation Summary
	Assert "Track & Order Details - Calculation Summary", Track_CheckCalculationSummary
	
End Sub


