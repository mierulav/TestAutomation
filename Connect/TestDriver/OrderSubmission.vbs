Dim qtApp, qtTest, qtResultsOpt

'Create the QTP Application object
Set qtApp = CreateObject("QuickTest.Application")

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
qtApp.Quit

'Release Object
Set qtTest = Nothing
Set qtApp = Nothing

