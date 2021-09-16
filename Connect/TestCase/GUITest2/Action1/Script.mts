'Dim objExcel, objFSO, objQTP
'
'Set objExcel = CreateObject("Excel.Application")
'Set objFSO = CreateObject("Scripting.FileSytemObject")
'Set objQTP = CreateObject("QuickTest.Application")
'Set objWShell = CreateObject("Wscript.Shell")
'Set objResult = CreateObject("QuickTest.RunResultsOptions")
'
'objResult.ResultsLocation = E
'objQTP.Test.Run 
'Set objExcel = Nothing
'
'
Dim objQTP, objXL, objWb, objSh

ParentFilePath = "D:\Microsoft\OneDrive\OneDrive - DKSH\Documents\automation\Connect"
TestCasePath = ParentFilePath + "\TestCase"
ExecutionListPath = ParentFilePath + "\TestDriver\ExecutionList.xlsx"

Set objXL = CreateObject("Excel.Application")
Set objWb = objXL.Workbooks.Open(FilePath)
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
For i = 2 To NumRows

	objResult.ResultsLocation =  TestCasePath + "\" + objSh(i, "C")
	objQTP.Open(objSh(i, "B"))
	objSh.Cells(i, "D") = Now
	objQTP.Test.Run objResult
	objSh.Cells(i, "E") = Now
	objQTP.Test.Close
	
Next
	
objQTP.Quit
objWb.Save

Set objQTP = Nothing
Set objSh = Nothing
Set objWb = Nothing
Set objXL = Nothing

