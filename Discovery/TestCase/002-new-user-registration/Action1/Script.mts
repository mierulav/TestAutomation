
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
		Dim strCompany : strCompany = Datatable.Value("Company", "Local")
		Dim strIndustry : strIndustry = Datatable.Value("Industry", "Local")
		Dim strCountry	: strCountry = Datatable.Value("Country", "Local")
		Dim strFname : strFname = GetInitialEachStrings(strIndustry) & "-" & Lcase(Left(strCountry, 3)) & "-" & GetStringDate
		Dim strLname : strLname = strCompany
		Dim strEmail : strEmail = strFname & "@yopmail.com"
		Dim strPassword : strPassword = "Test@123"
		
		'Precondition
		SystemUtil.Run Environment.Value("Browser") &".exe", strDiscoUrl
		Browser("CreationTime:=0").OpenNewTab
		Browser("CreationTime:=1").Navigate(strSFUrl)
		Browser("CreationTime:=0").OpenNewTab
		Browser("CreationTime:=2").Navigate(strYopMailUrl)
				
		'Step 1: registration on discover page
		Browser("CreationTime:=0").Sync
		goToUserRegistration()
		registerNewUser()
		
		'Step 2: check email notification
		Browser("CreationTime:=2").Sync
		checkYopmail strEmail
		
		'Step 3: login salesforce
		Browser("CreationTime:=1").Sync
		loginSalesforce strSFUser, strSFPwd
		
		'Step 4: convert lead in salesforce
		convertLeadInSFDC()
		
		'Step 5: check email now user registered and login
		Browser("CreationTime:=0").Close()
		checkAccountCreated
		
		Browser("CreationTime:=1").CloseAllTabs
	End If
Next

'==============================================================
' Main script operation subs
'==============================================================

'registration on discover page
sub registerNewUser()
	
	Browser("DKSH Discover | Performance").Page("Register").Check CheckPoint("Register")
	Browser("DKSH Discover | Performance").Page("Register").WebList("titleCode").Select "Dr."
	Browser("DKSH Discover | Performance").Page("Register").WebEdit("firstName").Set strFname
	Browser("DKSH Discover | Performance").Page("Register").WebEdit("lastName").Set strLname
	Browser("DKSH Discover | Performance").Page("Register").WebEdit("email").Set strEmail
	Browser("DKSH Discover | Performance").Page("Register").WebEdit("pwd").Set strPassword
	Browser("DKSH Discover | Performance").Page("Register").WebEdit("checkPwd").Set strPassword
	Browser("DKSH Discover | Performance").Page("Register").WebEdit("company").Set strCompany
	Browser("DKSH Discover | Performance").Page("Register").WebList("industry").Select strIndustry
	Browser("DKSH Discover | Performance").Page("Register").WebList("department").Select "Accounting & Finance"
	Browser("DKSH Discover | Performance").Page("Register").WebList("position").Select "Assistant Manager"
	Browser("DKSH Discover | Performance").Page("Register").WebList("country").Select strCountry
	Browser("DKSH Discover | Performance").Page("Register").WebEdit("city").Set "city field here"
	Browser("DKSH Discover | Performance").Page("Register").WebEdit("postCode").Set "10110"
	Browser("DKSH Discover | Performance").Page("Register").WebEdit("addresses").Set "address field here "
	Browser("DKSH Discover | Performance").Page("Register").WebEdit("phone").Set "12341234"
	Browser("DKSH Discover | Performance").Page("Register").WebElement("WebElement").Click
	Browser("DKSH Discover | Performance").Page("Register").WebButton("Register").Click
	Browser("DKSH Discover | Performance").Page("Register Thankyou Page").Image("PIM-Register-Thankyou-1400x400").Check CheckPoint("PIM-Register-Thankyou-1400x400")
	
end sub

'convert lead in salesforce
Sub convertLeadInSFDC()
	Wait(5)
	searchItem strEmail
	Dim blResult : blResult = checkSearchedAccount()
	Assert blResult, "Account is not created in SFDC !"
	If Not blResult Then
		ExitAction
		Browser("CreationTime:=0").CloseAllTabs
	End If
	clickIntoSearchedAccount()
	convertLead()
End Sub

'check email now user registered and login
Sub checkAccountCreated()
	Browser("CreationTime:=2").Sync
	checkYopmail strEmail
	checkActivationEmailLink()
	Browser("CreationTime:=2").Sync
	assert loginDiscover(strEmail, strPassword), "Login failed !"
End Sub


 @@ hightlight id_;_4523944_;_script infofile_;_ZIP::ssf65.xml_;_
 @@ script infofile_;_ZIP::ssf5.xml_;_
