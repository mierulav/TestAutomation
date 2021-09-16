OPTION EXPLICIT @@ script infofile_;_ZIP::ssf3.xml_;_

' 1. Procedural data
Dim TestList : TestList = TestDataDir + "\Validations\Track and trace order.xls"

' 2. Test Data Information
Datatable.AddSheet "TestList"
Dim i, strSearchVal, strTrackBy, bTempResult, bTemp, strOrderStatus
Datatable.ImportSheet TestList, "TestData", "TestList"

'Open Connect Site
SystemUtil.Run DefaultBrowser, SystemURL + "common/en/login"

'Start Test
 For i = 1 To Datatable.GetSheet("TestList").GetRowCount
 	Datatable.GetSheet("TestList").SetCurrentRow(i)
 		If UCase(Datatable.Value("ToTest", "TestList")) = "Y" Then
 			strSearchVal = Datatable.Value("SearchValue", "TestList")
 			strTrackBy = Datatable.Value("TrackBy", "TestList")
 			
 			'Execution Track & Trace
 			bTempResult = ExecuteTrackAndTrace(strSearchVal, strTrackBy)
 			
			'Insert Result into Test Result
			Datatable.Value("Result", "TestList") = bTempResult
			
 		End If
 	
 Next

'5 Export test into testresult
Datatable.ExportSheet TestResultDir + "\" + GetStringDate + "Track and trace order.xls", "TestList", "TestList"

'Functions Operations
Function ExecuteTrackAndTrace(strSearchValue, strTrackBy)

	'Go to login page
	Browser("DKSH Connect").Navigate SystemURL + "common/en/login"
	
	'Go To Track & Trace module
	GoToTrackAndTrace

	'Validate Trace Landing page
	Assert "Trace Landing Page", CheckTraceObjects
	
	'Search Trackby
	bTemp = TrackYourOrder(strTrackBy, strSearchValue)
	
	'Exit Function if no result
	If bTemp = "False" Then
		
		'denotes execution undone
		ExecuteTrackAndTrace = False
		Exit Function
		
	End If
	
	'Validation not for Sales Order Number tracking
	If strTrackBy <> "SalesOrderNumber" Then
		
		'Validate Trace Order List Result
		Assert "Trace Order List Result", CheckTraceOrderListObjects
		
	End If
	
	'Validation General Results objects
	Assert "Trace Order Result Object", CheckTraceResultObjects
	
	'Get Order Status first line item
	strOrderStatus = Trace_GetOrderStatus
	
	Select Case LCase(strOrderStatus)
		
		Case "order received"
			bTemp  = CheckOrderReceivedStatusObjects
			
		Case "order on hold"
			bTemp = CheckOrderOnHoldStatusObjects
			
		Case "order in process"
			bTemp = CheckOrderInProcessStatusObjects
		
		Case "deliver in transit"
			bTemp = CheckDeliveryInTransitStatusObjects
		
		Case "customer confirmed receipt"
			bTemp = CheckCCStatusObjects
			
		Case Else
			bTemp = False
			
	End Select
	
	'Validate the status specific objects
	Assert "Trace Status - " & strOrderStatus, bTemp
	
	'denotes execution done
	ExecuteTrackAndTrace = True
		 
End Function





