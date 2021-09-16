OPTION EXPLICIT @@ script infofile_;_ZIP::ssf3.xml_;_

' 1. Procedural data
Dim TestProducts : TestProducts = TestDataProduct +  "\ProductInformation.xls"
Dim TestList : TestList = TestDataValidation + "\Split order.xls"

' 2. Test Data Information
Dim x, y, i, strProductQuantity, strProductCode1, intProductPrice1, strProductCode2, intProductPrice2, bV1, strSalesOrg1, strSalesOrg2

Datatable.AddSheet "TestListSplitOrder"
Datatable.ImportSheet TestList, "SplitOrder", "TestListSplitOrder"
For x = 1 To Datatable.GetSheet("TestListSplitOrder").GetRowCount
	Datatable.GetSheet("TestListSplitOrder").SetCurrentRow(x)
	strSalesOrg1 = Datatable.Value("SalesOrg1", "TestListSplitOrder")
	strSalesOrg2 = Datatable.Value("SalesOrg2", "TestListSplitOrder")
Next

Datatable.AddSheet "ProductInfo"
Datatable.ImportSheet TestProducts, ProjectName, "ProductInfo"
For i = 1 To Datatable.GetSheet("ProductInfo").GetRowCount
	Datatable.GetSheet("ProductInfo").SetCurrentRow(i)
	'Select item to use
	If Ucase(Datatable.Value("SalesOrg", "ProductInfo")) = strSalesOrg1 Then
		strProductCode1 = Datatable.Value("ProductCode", "ProductInfo")
		intProductPrice1 = CSng(Datatable.Value("Price", "ProductInfo"))		
	End If
	If Ucase(Datatable.Value("SalesOrg", "ProductInfo")) = strSalesOrg2 Then
		strProductCode2 = Datatable.Value("ProductCode", "ProductInfo")
		intProductPrice2 = CSng(Datatable.Value("Price", "ProductInfo"))
	End If
	If i = Datatable.GetSheet("ProductInfo").GetRowCount and IsEmpty(strProductCode1) and IsEmpty(strProductCode2)  Then
		ExitAction False
		Exit For
	End If
Next

'3. Customer with mapped payer codes
ValidateSplitOrder

'4. Export test into testresults
Datatable.ExportSheet TestResultDir + "\" + TestCaseName, "TestListSplitOrder" , "SplitOrder"


'Subs Operations
Sub ValidateSplitOrder()

	'Navigate to the system
	SystemUtil.Run DefaultBrowser, SystemURL
	
	'Login
	Login GlobalUsername, GlobalPassword
	
	'Select ShipToID
	SelectShipToDefault
	
	'Precalculation to get minimum quantity to proceed checkout
	strProductQuantity = GetMinimumProductQuantity(intProductPrice1, CSng(MinimumPurchase))
	
	'Search for product to add to cart @@ script infofile_;_ZIP::ssf8.xml_;_
	SearchProductAndAddToCart strProductCode1, strProductQuantity @@ script infofile_;_ZIP::ssf13.xml_;_
	
	'Precalculation to get minimum quantity to proceed checkout
	strProductQuantity = GetMinimumProductQuantity(intProductPrice2, CSng(MinimumPurchase))
	
	'Search for product to add to cart @@ script infofile_;_ZIP::ssf8.xml_;_
	SearchProductAndAddToCart strProductCode2, strProductQuantity
	
	'Checkout
	Checkout
	
	'Place Order
	PlaceOrderWithoutDeliveryInstructions
	
	'Get Split Orders Number
	Dim arrSOs : arrSOs = GetSplitSalesOrderNumbers
	
	If Ubound(arrSOs) > 1 Then
		For y = 0 To Ubound(arrSOs)
			If IsNumeric(arrSOs(y)) Then
				bV1 = True
			Else
				bV1 = False
				Exit For
			End If
		Next
	Else
		bV1 = False
	End If
	
	'Logout	
	LogoutAndCloseBrowser
	
	'Stamp results
	Datatable.Value("Result1", "TestListSplitOrder") = bV1
	Datatable.Value("Remarks", "TestListSplitOrder") = "Tested"
	
End Sub

'Sub to immediate log out
Sub ImmediateLogout()
	'Logout	
	LogoutAndCloseBrowser
End Sub



