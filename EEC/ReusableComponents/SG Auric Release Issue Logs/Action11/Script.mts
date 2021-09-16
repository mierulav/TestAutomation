Option Explicit

Dim Home, Fx, arrTestData, strExtraVal, strVal, strSearchCriteria
Set Home = Home_Page
Set Fx = Function_Page
arrTestData = Split(Parameter("strTestData"), ",")
strSearchCriteria = arrTestData(0)
strVal =  arrTestData(1)
strExtraVal = arrTestData(2) ' this value is for Employee ID and Email address validation

'Step 1: Navigate to Function screen
Home.NavigateToFunction

'Step 2:
Select Case Lcase(strSearchCriteria)
	
	Case "employee id"
	Parameter("bResult") = Fx.ValidateSearchByEmployeeID(strVal, strExtraVal)
	
	Case "employee name"
	Parameter("bResult") = Fx.ValidateSearchByEmployeeName(strVal)
	
	Case "email address"
	Parameter("bResult") = Fx.ValidateSearchByEmailAddress(strVal, strExtraVal)
	
	Case "reference no"
	Parameter("bResult") = Fx.ValidateSearchByReferenceNo(strVal)

	Case "title"
	Parameter("bResult") = Fx.ValidateSearchByTitle(strVal)
	
	Case Else
	Parameter("bResult") = False
	ExitAction
	
End Select

