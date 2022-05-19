
'==============================================================
' Main script 
'==============================================================

option explicit

'global env parameter
Dim strDiscoUrl : strDiscoUrl = Environment("DiscoURL")
Dim strYopMailUrl : strYopMailUrl = Environment("YopMailURL")

'test data import
Datatable.ImportSheet  Environment.Value("ProjectFolder") & "\TestData\" & Environment.Value("TestName") & ".xls", "sheet1", "Local"

Dim i
For i = 1 To Datatable.GetSheet("Local").GetRowCount
	Datatable.GetSheet("Local").SetCurrentRow(i)
	If Ucase(Datatable.Value("ToTest", "Local")) = "Y" Then
		'Precondition
		SystemUtil.Run Environment.Value("Browser") &".exe", strDiscoUrl		
		
		Dim strUser : strUser = Datatable.Value("DiscoUsername", "Local")
		Dim strPassword : strPassword = Datatable.Value("DiscoPassword", "Local")
		'Step 1: Login discover
		signInToDiscover()
		loginDiscover strUser, strPassword @@ script infofile_;_ZIP::ssf6.xml_;_
		
		'Step 2: 'Navigate to My Profile and check blue tick
		checkCustomerIsSAPVerified()
				
		'Step 3: Go to Track and Order and get order information (code, name, quantity, package)
		getOrderHistoryInformation()
		
		'Step 4: Reorder
		clickReorderOnOrderPage()
		
		'Step 5: Check product info in cart before checkout -- validations on cart info
		checkCartInformationandProceedCheckout()
			
		'Step 5.1: Check product info in final review checkout -- validations on checkout info
		checkCheckoutFinalReview()
		
		Dim strSalesOrderNumber
		'Step 6: Set payment info and place order
		finalizeOrderPlacement()
		
		Dim strEmail : strEmail = strUser
		'Step 7: Verify information in Track & Order, Email, Order created in SAP -- validation
		Browser("CreationTime:=0").OpenNewTab
		Browser("CreationTime:=0").OpenNewTab
		Browser("CreationTime:=2").Navigate(strYopMailUrl)
		verifyOrderPlacementInSAPandEmail()
		
		Dim strOrderStatus : strOrderStatus = "Order Received"
		'Step 8: verifyOrderTrackingDetails()
		verifyOrderTrackingDetails()		
		
		Browser("CreationTime:=0").CloseAllTabs
	End If
Next



'==============================================================
' Main script operation subs
'==============================================================

'Navigate to My Profile and check blue tick

Sub checkCustomerIsSAPVerified()
	
	navigateToMyProfile()
	Dim blRes : blRes = checkCustomerVerified()
	Assert blRes, "Customer is not SAP Verified !"
	If blRes = False Then
		ExitAction
	End If
	
End Sub

'Check product info in cart before checkout -- validations on cart info
Sub checkCartInformationandProceedCheckout()
	
	'Assert checkCartIngredientInformation(strProductName), "Shopping cart product name is not correct !"
	Assert checkCartIngredientInformation(strProductCode), "Shopping cart product code is not correct !"
	Assert checkCartPackagingInformation(strPackagingType), "Shopping cart packaging type is not correct !"
	Assert checkCartQuantityInformation(strQuantity), "Shopping cart quantity is not correct !"
	clickCheckoutToFinalReview()
	
End Sub

'Check product info in final review checkout -- validations on checkout info
Sub checkCheckoutFinalReview()
	
	'Assert checkCheckoutProductName(strProductName), "Checkout product name is not correct !"
	Assert checkCheckoutPackage(strPackagingType), "Checkout packaging type is not correct !"
	Assert checkCheckoutQuantity(strQuantity), "Checkout quantity is not correct !"
	
End Sub

'Set payment info and place order
Sub finalizeOrderPlacement()
	
	setPONumber "Automation1234"
	setDeliveryInstructions "Automation1234 Reorder for Discovery+"
	attachPODocument()
	placeOrder()
	Assert checkOrderSubmission(), "Order submission is taking longer time or it is not successful !"
	strSalesOrderNumber = getSalesOrderNumber
	
End Sub

'Verify information in Track & Order, Email, Order created in SAP -- validation
Sub verifyOrderPlacementInSAPandEmail()
	
	Assert CheckSalesOrderNumberInSAP(strSalesOrderNumber), "Order is not successfully generated in SAP !"
	checkYopmail strEmail
	Assert checkOrderConfirmationEmail(), "Order confirmation is not received !"
	Assert checkPOAttachment("000-po-attachment.xlsx"), "PO  attachment is not found !"
	
End Sub

'Go to Track and Order and get order information (code, name, quantity, package)
Sub getOrderHistoryInformation()
	
	navigateToOrderTracking()
	clickViewDetailsOnOrderPage()
	Dim strProductCode : strProductCode = getTODisplayedProductCode()
'	Dim strProductName : strProductName = getTODisplayedProductName()
	Dim strPackagingType : strPackagingType = getTODisplayedPackageType()
	Dim strQuantity : strQuantity = getToDisplayedQuantity()
	
End Sub

'check tracking order information
Sub verifyOrderTrackingDetails()
	
	navigateToOrderTracking()
	searchOrder strSalesOrderNumber
	If Not checkSearchResultFirstItem(strSalesOrderNumber) Then
		Assert False, "Sales order " & strSalesOrderNumber & " is not found !"
		ExitAction
	End If
	clickViewDetailsOnOrderPage()
	Assert checkOrderStatusOnOrderDetailPage(strOrderStatus), "Sales order status is not " & strOrderStatus
	
End Sub @@ script infofile_;_ZIP::ssf139.xml_;_
