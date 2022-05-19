
Option Explicit

'Setup test environment
Dim ExecutionList, NumRows, i, tempRes
Dim cTest : Set cTest = TestEnv()
ExecutionList = cTest.GetTestModule

'Create excel object
If Not objExcel.Visible Then
	objExcel.Visible = True
End If
Set objELWorkbook = objExcel.Workbooks.Open(ExecutionList)
Set objELSheet = objELWorkbook.Worksheets("Sheet1")
NumRows = objELSheet.UsedRange.Rows.Count

'Set columns order
Dim ToTest :  ToTest = 1
Dim TestCaseName : TestCaseName = 2
Dim ExecutionStatus : ExecutionStatus = 3
Dim Timestamp : Timestamp = 4

For i = 2 To NumRows Step 1
	If UCase(objELSheet.Cells(i, ToTest).Value) = "Y" Then
		objELSheet.Cells(i, ExecutionStatus) = cTest.ReadTest(objELSheet.Cells(i, TestCaseName).Value)
		objELSheet.Cells(i, Timestamp).Value = Now
	End If 
	Next

objExcel.Quit

