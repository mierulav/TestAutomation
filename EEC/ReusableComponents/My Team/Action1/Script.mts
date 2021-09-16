Option Explicit

Dim Homepage : Set Homepage = Home_Page()
Dim MyTeam : Set MyTeam = MyTeam_Page()

'Navigate to myteam
Homepage.NavigateToMyTeam

'Search for expense report
MyTeam.SearchExpenseReport Parameter("strSearch")

'Edit expense report
MyTeam.EditExpenseReport

'Revise
MyTeam.ReviseExpenseReport Parameter("strReason")

'Validate no expense report is in the expense report list
If MyTeam.SearchExpenseReport(Parameter("strSearch")) Then
	Parameter("bResult1") = False
Else
	Parameter("bResult1") = True
End If


