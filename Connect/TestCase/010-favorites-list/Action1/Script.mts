Option Explicit

' 1. Test Data Path
Dim TestList : TestList = TestDataDir + "\Validations\" & Environment.Value("TestName") & ".xls"
Datatable.ImportSheet TestList, "TestData", "Global"

'Get Test Data
Dim i
For i = 1 To Datatable.GetSheet("Global").GetRowCount
	Datatable.GetSheet("Global").SetCurrentRow(i)
	If UCase(Datatable.Value("ToTest", "Global")) = "Y" Then 
		'List for test
		ProjectName = Datatable.Value("Market", "Global")
		Dim strDate : strDate = GetStringDate
		Dim strListName : strListName = strDate & "-DNU-Automation"
		Dim strListName2 : strListName2 = strDate & "-DNU-Automation2"
	
		LoginIntoConnect
		CreateNewFavoritesList
		AddProductIntoExistingFavoritesList
		AddProductIntoNewFavoritesList
		AddToCartAllProductsInFavoritesList
		OrderSubmission
		DeleteAllCreatedFavoritesLists(strListName)
		DeleteAllCreatedFavoritesLists(strListName2)
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

Sub OrderSubmission()
	
	If GetProductUnitPrice = 0 Then
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
	SetDeliveryInstruction "This is order placement created from favorites list " & ProjectName
	
	Select Case ProjectName
		Case "VNHEC"
		
		Case "MMHEC"
			SetPONumber "AutomationTest"
			SelectOOSProceedingAgreement "agree"
			
		Case Else
			SetPONumber "AutomationTest"
	End Select
		
	SubmitOrder
	
	'Validate Order Confirmation Page
	If CheckSalesOrderConfirmed and GetOrderNumber <> False Then
		Dim strOrderNumber : strOrderNumber = GetOrderNumber
	Else
		AssertExitRun "Order Submission for " & ProjectName, "Unsuccessful order submission"
		Exit Sub
	End If 
	
	'Validate Order sent to ERP
	SAPEasyAccessScreen
	If  GetDeliveryInstruction(strOrderNumber, ProjectName) =  "This is order placement created from favorites list " & ProjectName Then
		Dim tempRes : tempRes = true
	End If
	
	Assert "Check Order Created in SAP", tempRes
	
End Sub

'Create new list in the Favorites list module
Sub CreateNewFavoritesList()
	
	OpenFavoritesList
	ClickNewFavoritesList
	FLSetFavoritesListName strListName
	FLSaveFavoritesList
	
	'Check newly created list
	Assert ProjectName & "- Create New Favourites List in Favorites List module", CheckFavoritesList(strListName)
	
End Sub

'Add product into the newly created Favorites list 
Sub AddProductIntoExistingFavoritesList()
	
	'Search for product
	OpenAllProductPage
	SearchProduct(Datatable.Value("ProductCode", "Global"))
	OpenProductPDP
	'Add as favorites
	FavoriteAProduct
	SelectExistingFavoritesList strListName
	SaveProductAsFavorite
	'Check product saved alert
	Assert ProjectName & " - Alert Succesful Added to Favorites List ", CheckSuccesfulAlert
	CloseFavoritesListBox @@ script infofile_;_ZIP::ssf66.xml_;_
	OpenFavoritesList
	SelectFavoritesList strListName
	Assert  ProjectName & " - Add Product into Existing Favorites List", CheckProductInFavoritesList(Datatable.Value("ProductCode", "Global"))
	
End Sub

'Add and create new Favorites list on-the-go
Sub AddProductIntoNewFavoritesList()
	
	'Search for product
	OpenAllProductPage
	SearchProduct(Datatable.Value("ProductCode", "Global"))
	OpenProductPDP
	'Add as favorites
	FavoriteAProduct
	PDPSetFavoritesListName strListName2
	SaveProductAsFavorite
	'Check product saved alert
	Assert  ProjectName & " - Alert Succesful Added to Favorites List ", CheckSuccesfulAlert
	CloseFavoritesListBox @@ script infofile_;_ZIP::ssf66.xml_;_
	OpenFavoritesList
	SelectFavoritesList strListName2
	Assert  ProjectName & " - Add Product into Created Favorites List from PDP", CheckProductInFavoritesList(Datatable.Value("ProductCode", "Global"))
	
End Sub

'Submit all products in the Favorites lity
Sub AddToCartAllProductsInFavoritesList()
	
	OpenFavoritesList
	SelectFavoritesList strListName
	Dim arrProductCodes : arrProductCodes  = GetAllItemsInFavoritesList
	AddAllItemsToCart
	Dim i
	For i = 0 To Ubound(arrProductCodes)-1
		Dim blnRes : blnRes = CheckSpecificProductCode(arrProductCodes(i))
		Assert  ProjectName & " - Check Product " & arrProductCodes(i) & " is In The Cart",  blnRes
		If blnRes = False Then
			AssertExitRun ProjectName & ": Import product into cart", "Unsuccessful import product into cart !"
		End If
	Next
	If GetProductUnitPrice = 0 Then
		Assert ProjectName & " - Product price is 0 (Either OOS or Pricing issue)", False
		Exit Sub
	End If
	SetProductQuantityBasedOnMOV
	
End Sub

Sub DeleteAllCreatedFavoritesLists(strList)
	OpenFavoritesList
	SelectFavoritesList(strList)
	DeleteFavoritesList
End Sub


