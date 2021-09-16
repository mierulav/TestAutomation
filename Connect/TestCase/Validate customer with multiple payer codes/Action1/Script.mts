OPTION EXPLICIT @@ script infofile_;_ZIP::ssf3.xml_;_

' 1. Procedural data
Dim TestProducts : TestProducts = TestDataProduct +  "\ProductInformation.xls"
Dim TestList : TestList = TestDataDir + "\Validations\Customer with multiple payer codes.xls"

' 2. Test Data Information
Dim x, y, i, strCustomerID, strShipToID, strPayerCode, strProductQuantity, strProductCode, intProductPrice, bV1, bV2, bV3, bFlag 
Datatable.AddSheet "ProductInfo"
Datatable.ImportSheet TestProducts, ProjectName, "ProductInfo"

For i = 1 To Datatable.GetSheet("ProductInfo").GetRowCount
	Datatable.GetSheet("ProductInfo").SetCurrentRow(i)
	'Select item to use
	If Ucase(Datatable.Value("ToUse", "ProductInfo")) = "Y" Then
		strProductCode = Datatable.Value("ProductCode", "ProductInfo")
		intProductPrice = CSng(Datatable.Value("Price", "ProductInfo"))
		Exit For		
	End If
Next

' 3. Customer with mapped payer codes
ValidateMappedPayerCode

' 4. Customer with unmapped payer codes
ValidateUnMappedPayerCode

'5 Export test into testresult
Datatable.ExportSheet TestResultDir + "\" + TestCaseName + ".xls", "TestListMapped", "MappedPayerCode"
Datatable.ExportSheet TestResultDir + "\" + TestCaseName + ".xls", "TestListUnMapped", "UnMappedPayerCode"


'Subs Operations
Sub ValidateMappedPayerCode()

Datatable.AddSheet "TestListMapped"
Datatable.ImportSheet TestList, "PayerCodeMapped", "TestListMapped"
For x = 1 To Datatable.GetSheet("TestListMapped").GetRowCount : Do
	bFlag = False
	Datatable.GetSheet("TestListMapped").SetCurrentRow(x)
	strCustomerID = Datatable.Value("CustomerID", "TestListMapped")
	strShipToID = Datatable.Value("ShipToID", "TestListMapped")
	strPayerCode = Datatable.Value("PayerCode", "TestListMapped")
	
	'Navigate to the system
	SystemUtil.Run DefaultBrowser, SystemURL
	
	'Login
	Login strCustomerID, "12341234"
	
	'Select ShipToID
	SelectShipTo(strShipToID)
	
	'Precalculation to get minimum quantity to proceed checkout
	strProductQuantity = GetMinimumProductQuantity(intProductPrice, CSng(MinimumPurchase))
	
	'Search for product to add to cart @@ script infofile_;_ZIP::ssf8.xml_;_
	SearchProductAndAddToCart strProductCode, strProductQuantity @@ script infofile_;_ZIP::ssf13.xml_;_
	
	'Check block code if exist and go to next loop
	If CheckBlockCode1 Then
		'Logout	
		LogoutAndCloseBrowser
		
		'Stamp results
		Datatable.Value("Result1", "TestListMapped") = "NA"
		Datatable.Value("Result2", "TestListMapped") = "NA"
		Datatable.Value("Remarks", "TestListMapped") = "Block Code Exist"
		Exit Do
	End If
	
	'Checkout
	Checkout
	
	'Get the defaulted payer code address displayed
	Browser("DKSH Connect").Page("Checkout").Sync
	Dim strDefaultPayerDetails : strDefaultPayerDetails = Trim(Browser("DKSH Connect").Page("Checkout").WebElement("PayerDetails").GetROProperty("innertext"))
	
	'Go to Payer Book
	Browser("DKSH Connect").Page("Checkout").WebButton("Change payer").Click
	
	Dim objDesc : Set objDesc = Description.Create
	objDesc("class").Value = "btn btn-primary btn-block js-payer-select"	
	Dim objC : Set objC = Browser("creationtime:=0").Page("title:=Checkout.*").ChildObjects(objDesc)
	
	'Validation #1: multiple payer code available to select
	If objC.Count > 1 Then
		bV1 = True
	Else
		bV1 = False
	End If
	
	'Select payer
	For i = 0 To objC.Count-1
		If GetNumber(objC(i).GetROProperty("outerhtml")) = strShipToID Then
			objC(i).Click
		End If 
	Next
	
	'Get payer details displayed
	Dim strPayerDetails : strPayerDetails = Trim(Browser("DKSH Connect").Page("Checkout").WebElement("PayerDetails").GetROProperty("innertext"))
	
	'Validation #2: Defaulted payer code must be same as ShipTo Code
	If strPayerDetails = strDefaultPayerDetails Then
		bV2 = True
	Else
		bV2 = False
	End If
	
	'Logout	
	LogoutAndCloseBrowser
	
	'Stamp results
	Datatable.Value("Result1", "TestListMapped") = bV1
	Datatable.Value("Result2", "TestListMapped") = bV2
	Datatable.Value("Remarks", "TestListMapped") = "Tested"
	
Loop While bFlag <> False : Next
	
End Sub

Sub ValidateUnMappedPayerCode()

Datatable.AddSheet "TestListUnMapped"
Datatable.ImportSheet TestList, "PayerCodeNotMapped", "TestListUnMapped"
For y = 1 To Datatable.GetSheet("TestListUnMapped").GetRowCount : Do
	bFlag = False
	Datatable.GetSheet("TestListUnMapped").SetCurrentRow(y)
	strCustomerID = Datatable.Value("CustomerID", "TestListUnMapped")
	strShipToID = Datatable.Value("ShipToID", "TestListUnMapped")
	strPayerCode = Datatable.Value("PayerCode", "TestListUnMapped")
	
	'Navigate to the system
	SystemUtil.Run DefaultBrowser, SystemURL
	
	'Login
	Login strCustomerID, "12341234"
	
	'Select ShipToID
	SelectShipTo(strShipToID)
	
	'Precalculation to get minimum quantity to proceed checkout
	strProductQuantity = GetMinimumProductQuantity(intProductPrice, CSng(MinimumPurchase))
	
	'Search for product to add to cart @@ script infofile_;_ZIP::ssf8.xml_;_
	SearchProductAndAddToCart strProductCode, strProductQuantity
	
	'Check block code if exist and go to next loop
	If CheckBlockCode1 Then
		'Logout	
		LogoutAndCloseBrowser
		
		'Stamp results
		Datatable.Value("Result1", "TestListUnMapped") = "NA"
		Datatable.Value("Result2", "TestListUnMapped") = "NA"
		Datatable.Value("Result3", "TestListUnMapped") = "NA"
		Datatable.Value("Remarks", "TestListUnMapped") = "Block Code Exist"
		Exit Do
	End If
	
	'Checkout
	Checkout
	
	'Validation #1: Payer Book pop-up
	If Browser("DKSH Connect").Page("Checkout").WebElement("PayerBook").Exist Then
		bV1 = True
	Else
		bV1 = False
	End If 
	
	'Validation #2: Multiple payer codes available to select
	Dim objDesc : Set objDesc = Description.Create
	objDesc("class").Value = "btn btn-primary btn-block js-payer-select"	
	Dim objC : Set objC = Browser("creationtime:=0").Page("title:=Checkout.*").ChildObjects(objDesc)
	
	If objC.Count > 1 Then
		bV2 = True
	Else
		bV2 = False
	End If
	
	'Select payer
	For i = 0 To objC.Count-1
		If GetNumber(objC(i).GetROProperty("outerhtml")) = strShipToID Then
			objC(i).Click
		End If 
	Next
	
	'Validation #3: Able to select payer
	Browser("DKSH Connect").Page("Checkout").WebElement("PayerDetails").WaitProperty "Visible", True
	If Trim(Browser("DKSH Connect").Page("Checkout").WebElement("PayerDetails").GetROProperty("innertext")) = "" Then
		bV3 = False
	Else
		bV3 = True
	End If
	
	'Logout	
	LogoutAndCloseBrowser
	
	'Stamp results
	Datatable.Value("Result1", "TestListUnMapped") = bV1
	Datatable.Value("Result2", "TestListUnMapped") = bV2
	Datatable.Value("Result3", "TestListUnMapped") = bV3
	Datatable.Value("Remarks", "TestListUnMapped") = "Tested"
	
Loop While bFlag = True : Next
	
End Sub


