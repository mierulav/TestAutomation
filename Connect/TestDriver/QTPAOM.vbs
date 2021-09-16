QTP Automatomation Object Model

'Initiate QTP object
Set qtpObj = CreateObject(QuickTest.Application)

'Launch QTP and make it visible
qtpObj.Launch
qtpObj.Visible = True

'Open test path
qtpObj.Open("path")

'Run test
qtpObj.Test.Run

'Close test
qtpObj.Test.Close

'Close QTP
qtpObj.Quit

'Release obj
Set qtpObj = Nothing

