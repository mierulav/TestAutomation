Option Explicit

Dim Test : Set Test = Init()
Dim x, i, j 
Dim strOutputs, arrOutputs, strOVendorCode, strOGLCodeExpense, strOGLCodeTax, strOTotalAmount, strOExpensesAmount, strOTaxAmount, bResult, strInitialVal, strToAddVal1, strToAddVal2, strToAddVal3, strToAddVal4, strToAddVal5

'Get Test Data Processed status expenses and GLCode
Datatable.ImportSheet Test.GetTestDataFile, Test.GetCountry, "Local"
'Datatable.ImportSheet Test.GetTestDataGlobal & "\GLCode\" & Test.GetCountry & ".xls", Test.GetCompanyCode, "Global"
Datatable.ImportSheet Test.GetTestDataGlobal & "\GLCode\GLCode.xls", "Sheet1", "Global"

'Run SAP Posting Validation 
For i = 1 To Datatable.GetSheet("Local").GetRowCount
	Datatable.GetSheet("Local").SetCurrentRow(i)
	Select Case Datatable.Value("EECTax", "Local")
		Case "Y"
			RunAction "SAP Posting Validation For Tax Expense [SAP]", oneIteration, _
				Datatable.Value("VendorCode", "Local"), Datatable.Value("CompanyCode", "Local"), Trim(Datatable.Value("EECRefNo", "Local")), strOutputs, bResult
				Datatable.GetSheet("Local").SetCurrentRow(i)
				If bResult = False Then
					Reporter.ReportEvent micFail, "Not able to find SAP Posting for RefNo " & Trim(Datatable.Value("EECRefNo", "Local")) & " , Category " &  Datatable.Value("EECExpenseCategory", "Local"), "Fail"
				Else
					arrOutputs = Split(strOutputs, ",")
					strOVendorCode = arrOutputs(0)
					strOGLCodeExpense = arrOutputs(1)
					strOGLCodeTax = arrOutputs(2)
					strOTotalAmount = arrOutputs(3)
					strOExpensesAmount = arrOutputs(4)
					strOTaxAmount = arrOutputs(5) @@ hightlight id_;_0_;_script infofile_;_ZIP::ssf1.xml_;_
				End If
		Case Else
			RunAction "SAP Posting Validation for Non Tax Expense [SAP]", oneIteration, _
				Datatable.Value("VendorCode", "Local"), Datatable.Value("CompanyCode", "Local"), Trim(Datatable.Value("EECRefNo", "Local")), strOutputs, bResult	
				Datatable.GetSheet("Local").SetCurrentRow(i)
				If bResult = False Then
					Reporter.ReportEvent micFail, "Not able to find SAP Posting for RefNo " & Trim(Datatable.Value("EECRefNo", "Local")) & " , Category " &  Datatable.Value("EECExpenseCategory", "Local"), "Fail"
				Else
					arrOutputs = Split(strOutputs, ",")
					strOVendorCode = arrOutputs(0)
					strOGLCodeExpense = arrOutputs(1)
					strOTotalAmount = arrOutputs(2)
					strOExpensesAmount = arrOutputs(3)
				End If
	End Select
	
	'Get GL Code for specific Expense category and validate the SAP posting
	For x = 1 To Datatable.GetSheet("Global").GetRowCount
		Datatable.GetSheet("Global").SetCurrentRow(x)
		Datatable.GetSheet("Local").SetCurrentRow(i)
		
		'High level result to be export to excel
		If bResult Then
			Datatable.Value("Results", "Local") = "OK"
		Else
			Datatable.Value("Results", "Local") = "No SAP documents found"
		End If
		
		'Result field value initial
		strInitialVal = Datatable.Value("Results", "Local")
		
	
		'detail results
		If bResult _
			And LCase(Datatable.Value("ExpenseCategory", "Global")) = LCase(Datatable.Value("EECExpenseCategory", "Local")) Then
				Reporter.ReportEvent micDone, "RefNo: " &  Trim(Datatable.Value("EECRefNo", "Local")), "Expense Category: " & Datatable.Value("EECExpenseCategory", "Local")
			
			'Validate Vendor Code
			If Datatable.Value("VendorCode", "Local") = strOVendorCode Then
				Reporter.ReportEvent micPass, "SAP Posting Validation: Validate Vendor Code is " & Datatable.Value("VendorCode", "Local"), "Actual : " & strOVendorCode
			Else
				Reporter.ReportEvent micFail, "SAP Posting Validation: Validate Vendor Code is " & Datatable.Value("VendorCode", "Local"), "Actual : " & strOVendorCode
				strToAddVal1 = "Vendor Code Not Tally, Actual: " & strOVendorCode
			End If
			
			'Validate Total Amount (tax inclusive)
			If CSng(Datatable.Value("EECAmount", "Local")) = CSng(strOTotalAmount) Then
				Reporter.ReportEvent micPass, "SAP Posting Validation: Validate Amount is " & Datatable.Value("EECAmount", "Local"), "Actual : " & strOTotalAmount
			Else
				Reporter.ReportEvent micFail, "SAP Posting Validation: Validate Amount is " & Datatable.Value("EECAmount", "Local"), "Actual : " & strOTotalAmount
				strToAddVal2 = "Total Amount Not Tally, Actual: " & strOTotalAmount
			End If
			
			'Validate GLCode expense
			If Datatable.Value("GLCodeExpense", "Global") = strOGLCodeExpense Then
				Reporter.ReportEvent micPass, "SAP Posting Validation: Validate GL Account for Expense type is " & Datatable.Value("GLCodeExpense", "Global"), "Actual : " & strOGLCodeExpense
			Else
				Reporter.ReportEvent micFail, "SAP Posting Validation: Validate GL Account for Expense type is " & Datatable.Value("GLCodeExpense", "Global"), "Actual : " & strOGLCodeExpense
				strToAddVal3 = "GL Code Not Tally, Actual: " & strOGLCodeExpense
			End If
			
			'For Tax Expenses
			If Datatable.Value("EECTax", "Local") = "Y" Then
				
				'Validate Tax Amount			
				If Csng(Datatable.Value("EECTaxAmount", "Local")) = CSng(strOTaxAmount) Then
					Reporter.ReportEvent micPass, "SAP Posting Validation: Validate Amount is " & Datatable.Value("EECTaxAmount", "Local"), "Actual : " & strOTaxAmount
				Else
					Reporter.ReportEvent micFail, "SAP Posting Validation: Validate Amount is " & Datatable.Value("EECTaxAmount", "Local"), "Actual : " & strOTaxAmount
					strToAddVal4 = "Tax Amount Not Tally, Actual: " & strOTaxAmount
				End If	
				
				'Validate GLCode tax
				If Datatable.Value("GLCodeTax", "Global") = strOGLCodeTax Then
					Reporter.ReportEvent micPass, "SAP Posting Validation: Validate GL Account for Tax is " & Datatable.Value("GLCodeTax", "Global"), "Actual : " & strOGLCodeTax
				Else
					Reporter.ReportEvent micFail, "SAP Posting Validation: Validate GL Account for Tax is " & Datatable.Value("GLCodeTax", "Global"), "Actual : " & strOGLCodeTax
					strToAddVal5 = "Tax Amount Not Tally, Actual: " & strOGLCodeTax
				End If

			End If
		
			Datatable.Value("Results", "Local") = strInitialVal & ", " & strToAddVal1 & ", " & strToAddVal2 & ", " & strToAddVal3 & ", " & strToAddVal4 & ", " & strToAddVal5
		
		End If
	Next
Next

Datatable.ExportSheet Environment.Value("ProjectDir") & "\TestResult\SAPPosting\" & Init.GetCompanyCode & ".xls", "Local", "Sheet1"

			

