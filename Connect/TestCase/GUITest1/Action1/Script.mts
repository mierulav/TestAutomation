Dim qtApp, qtTest, qtResultsOpt, objExcel, FilePath

'Create the QTP Application object
Set qtApp = CreateObject("QuickTest.Application")
'Create Excel Application
Set objExcel = CreateObject("Excel.Application")
FilePath = ""

'Open Workbook
objExcel.Visible = True
objExcel.Workbooks.Open FilePath
objExcel.

'If QTP is notopen then open it
If  qtApp.launched <> True then
	qtApp.Launch
End If

'Make the QuickTest application visible - This is optional
qtApp.Visible = True



'Open the test in read-only mode
qtApp.Open "D:\Microsoft\OneDrive\OneDrive - DKSH\Documents\automation\Connect\TestCase\Validate order submission via PDP", True

'set run settings for the test
Set qtTest = qtApp.Test

'Run the test
qtTest.Run



' Close the test
qtTest.Close

'Close QTP
qtApp.quit

'Release Object
Set qtTest = Nothing
Set qtApp = Nothing


Set qtResultsOpt = CreateObject("QuickTest.RunResultsOptions")  'Create the Run Results Options object
qtResultsOpt.ResultsLocation = strTestResPath 'Specify the location to save the test results.
qtTest.Run qtResultsOpt,True 'Run the test and wait until end of the test run
