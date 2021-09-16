Dim objQTP, objXL, objWb, objSh, i, NumRows

ParentFilePath = "D:\Microsoft\OneDrive\OneDrive - DKSH\Documents\automation\Connect"
TestCasePath = ParentFilePath + "\TestCase"
ExecutionListPath = ParentFilePath + "\TestDriver\ExecutionList.xlsx"

Set objXL = CreateObject("Excel.Application")
Set objWb = objXL.Workbooks.Open(ExecutionListPath)
Set objSh = objWb.Worksheets(1)
Set objQTP = CreateObject("QuickTest.Application")
Set objResult = CreateOBject("QuickTest.RunResultsOptions")

'Shows Apps
If objQTP.Launched = False Then

	objQTP.Launch
	objQTP.Visible = True
	
End If

'Get rowNums
NumRows = objSh.UsedRange.Rows.Count

'Establish loop
For i = 2 To NumRows Step 1

	If objSh(i, "A") = "Y" Then

		objResult.ResultsLocation =  TestCasePath + "\" + objSh(i, "C")
		objQTP.Open(objSh(i, "B"))
		objSh.Cells(i, "D") = Now
		objQTP.Test.Run objResult
		objSh.Cells(i, "E") = Now
		objQTP.Test.Close
	
	End If
	
Next
	
objQTP.Quit
objWb.Save

Set objQTP = Nothing
Set objSh = Nothing
Set objWb = Nothing
Set objXL = Nothing
