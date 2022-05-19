
'==============================================================
' Main script 
'==============================================================

option explicit

'global env parameter
Dim strDiscoUrl : strDiscoUrl = Environment("DiscoURL")
Dim strSFUrl : strSFUrl = Environment("SFURL")
Dim strYopMailUrl : strYopMailUrl = Environment("YopMailURL")
Dim strSFUser : strSFUser =  Environment("SFUsername")
Dim strSFPwd : strSFPwd =  Environment("SFPassword")

'test data import
Datatable.ImportSheet  Environment.Value("ProjectFolder") & "\TestData\" & Environment.Value("TestName") & ".xls", "sheet1", "Local"

Dim i
For i = 1 To Datatable.GetSheet("Local").GetRowCount
	Datatable.GetSheet("Local").SetCurrentRow(i)
	If Ucase(Datatable.Value("ToTest", "Local")) = "Y" Then
		'Precondition
		SystemUtil.Run Environment.Value("Browser") &".exe", strDiscoUrl
		Browser("CreationTime:=0").OpenNewTab
		Browser("CreationTime:=1").Navigate(strSFUrl)
		Browser("CreationTime:=0").OpenNewTab
		Browser("CreationTime:=2").Navigate(strYopMailUrl)		
		
		Dim strUser : strUser = Datatable.Value("DiscoUsername", "Local")
		Dim strPassword : strPassword = Datatable.Value("DiscoPassword", "Local")
		'Step 1: Login discover
		signInToDiscover()
		loginDiscover strUser, strPassword @@ script infofile_;_ZIP::ssf6.xml_;_
		
		Dim strPIMCode : strPIMCode = Datatable.Value("PIMCode", "Local")
		Dim strAnnualAmt : strAnnualAmt = "1000"
		Dim strComments : strComments = "Please give your best price"
		Dim strQuoteNumber
		'Step 2: Submit quote to order
		SubmitQuoteFromDiscovery()
		
		'Step 3: 'Login Salesforce and process quote
		loginSalesforce strSFUser, strSFPwd
		
		Dim strValidityPeriod : strValidityPeriod = Datatable.Value("ValidityPeriod", "Local")
		Dim strDKSHComment : strDKSHComment =  Datatable.Value("DKSHComments", "Local")
		Dim strPayment : strPayment = Datatable.Value("PaymentTerms", "Local")
		Dim strInco : strInco = Datatable.Value("IncoTerms", "Local")
		Dim strQuotePrice : strQuotePrice = Datatable.Value("QuotePrice", "Local")
		Dim strQuotePackage : strQuotePackage = Datatable.Value("QuotePackage", "Local")
		Dim strQuoteProject : strQuoteProject = Datatable.Value("QuoteProject", "Local")
		'Step 4: Process inital quote
		processInitialQuoteInSalesforce()
		
		Dim strEmail : strEmail = strUser
		'Step 4.1: Check email for submitted quote order
		Browser("CreationTime:=2").Navigate(strYopMailUrl)
		checkYopmail strEmail
		Assert checkQuoteSentEmail, "Quote sent email is not found !"
			
		'Step 5: 'Check quote replied from dksh
		checkQuoteResponseAndResubmittedQuote()
		
		Dim strRevisePrice : strRevisePrice = Datatable.Value("RevisePrice", "Local")
		'Step 6: Proces again quote - Check status sfdc
		processResubmitedQuoteInSalesforce()
		
		'Step 6.1: Check email for submitted quote order
		Browser("CreationTime:=2").Navigate(strYopMailUrl)
		checkYopmail strEmail
		Assert checkQuoteResubmittedEmail, "Quote resubmitted email is not found !"
		
		'Step 7: Check quote replied from dksh
		checkQuoteResponse()
		
		Dim strSalesOrderNumber
		'Step 8: Place order
		fromQuoteOrderToPlaceOrder()
		
		'Step 8.1: Check email for submitted quote order
		Browser("CreationTime:=2").Navigate(strYopMailUrl)
		checkYopmail strEmail
		Assert checkOrderConfirmationEmail, "Order confirmation email is not found !"
		
		Dim strOrderStatus : strOrderStatus = "Order Received"
		'Step 9: Check track and order information
		checkOrderTrackingDetails()
		
		
		Browser("CreationTime:=0").CloseAllTabs
	End If
Next



'==============================================================
' Main script operation subs
'==============================================================

'Submit quote to order
Sub SubmitQuoteFromDiscovery()
	searchProduct strPIMCode
	toQuoteRequest()
	fillInQuoteDetails strAnnualAmt, strComments
	submitQuote()
	checkQuoteSubmission()
	strQuoteNumber = getQuoteNumber()
End Sub

'Process inital quote
Sub processInitialQuoteInSalesforce()
	Wait(60) 'Waiting for things to sync-up in SFDC
	searchItem strQuoteNumber
	clickIntoSearchedQuote()
	checkSearchedQuote()
	setQutoationValidity strValidityPeriod
	setDKSHComment strDKSHComment
	setPaymentAndIncoTerms strPayment, strInco
	saveQuotationHeader()
	goToQuotationProductDetails()
	setOneOffQuotationType()
	setQuoationPrice strQuotePrice
	setProductPackage strQuotePackage
	setQuotatationProject strQuoteProject
	saveQuotationProductDetails()
	goBackToQuoationHeader()
	Assert checkQuotationHeaderStatus("Quotation Sent"), "Saleforce status check for Quotation Requested failed !"
	
End Sub

'Check quote replied from dksh
Sub checkQuoteResponseAndResubmittedQuote()

	navigateToMyQuoteRequest()
	searchQuoteAndViewDetails strQuoteNumber
	Assert checkQuoteStatus("Vendor Quote"), 	"Discovery status check for Vendor Quote failed !"
	Assert checkQuoteValidityDate(strValidityPeriod), "Quote validity period in Discover is not correct ! "
	Assert checkQuotePrice(strQuotePrice), "Quote price displayed in Discover is not correct !"
	checkDKSHComment
	checkPaymentAndIncoTerms
	editQuotation
	setCustomerComments("please revise again")
	editShipping
	submitQuote
	checkQuoteSubmission
	navigateToMyQuoteRequest
	searchQuoteAndViewDetails strQuoteNumber
	Assert checkQuoteStatus("Resubmitted"), "Discovery status check for Resubmitted failed !"
	
End Sub

'Proces again quote - Check status sfdc
Sub processResubmitedQuoteInSalesforce()
	
	goToQuotation()
	searchQuotationNumber strQuoteNumber
	clickIntoQuotationDetails strQuoteNumber
	goToQuotationProductDetails()
	setQuoationPrice strRevisePrice
	saveQuotationProductDetails()
	goBackToQuoationHeader()
	Assert checkQuotationHeaderStatus("Revised Price"), "Saleforce status check for Revised Price failed !"
	
End Sub
 @@ script infofile_;_ZIP::ssf85.xml_;_
'Check quote replied from dksh
Sub checkQuoteResponse()

	navigateToMyQuoteRequest
	searchQuoteAndViewDetails strQuoteNumber
	Assert checkQuoteStatus("Vendor Quote"), 	"Discovery status check for Vendor Quote failed !"
	Assert checkQuoteValidityDate(strValidityPeriod), "Quote validity period in Discover is not correct ! "
	Assert checkQuotePrice(strRevisePrice), "Revise price displayed in Discover is not correct !"
	checkDKSHComment
	checkPaymentAndIncoTerms	
	
End Sub

'Place order
Sub fromQuoteOrderToPlaceOrder()

	navigateToMyQuoteRequest
	searchQuoteAndViewDetails strQuoteNumber
	placeQuotedOrder
	placeOrder
	Assert checkOrderSubmission, "Order submission is not successful !"
	strSalesOrderNumber = getSalesOrderNumber
	navigateToMyQuoteRequest
	searchQuoteAndViewDetails strQuoteNumber
	Assert checkQuoteStatus("Ordered"), "Discovery status check for Ordered failed !"
	
End Sub @@ script infofile_;_ZIP::ssf130.xml_;_

'check tracking order information
Sub checkOrderTrackingDetails()
	
	navigateToOrderTracking()
	searchOrder strSalesOrderNumber
	If Not checkSearchResultFirstItem(strSalesOrderNumber) Then
		Assert False, "Sales order " & strSalesOrderNumber & " is not found !"
		ExitAction
	End If
	clickViewDetailsOnOrderPage()
	Assert checkOrderStatusOnOrderDetailPage(strOrderStatus), "Sales order status is not " & strOrderStatus
	
End Sub @@ script infofile_;_ZIP::ssf133.xml_;_
