Option Explicit 

Dim strVendorCode : strVendorCode = Parameter("strVendorCode")
Dim strCompanyCode : strCompanyCode = Parameter("strCompanyCode")
Dim strRefNo : strRefNo = Parameter("strRefNo")
Dim SAPVendorLineItemDisplay : Set SAPVendorLineItemDisplay = SAPGuiSession("Session").SAPGuiWindow("Vendor Line Item Display")

SAP.OKCode.Set "fbl1n"
SAP.Enter.Click

'Enable search for all type of SAP doc
SAPGuiSession("Session").SAPGuiWindow("Vendor Line Item Display").SAPGuiCheckBox("Special G/L transactions").Set "ON" @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf1.xml_;_
SAPGuiSession("Session").SAPGuiWindow("Vendor Line Item Display").SAPGuiCheckBox("Noted items").Set "ON" @@ hightlight id_;_2_;_script infofile_;_ZIP::ssf1.xml_;_
SAPGuiSession("Session").SAPGuiWindow("Vendor Line Item Display").SAPGuiCheckBox("Parked items").Set "ON" @@ hightlight id_;_3_;_script infofile_;_ZIP::ssf1.xml_;_
SAPGuiSession("Session").SAPGuiWindow("Vendor Line Item Display").SAPGuiCheckBox("Customer items").Set "ON" @@ hightlight id_;_4_;_script infofile_;_ZIP::ssf1.xml_;_
 @@ hightlight id_;_0_;_script infofile_;_ZIP::ssf1.xml_;_
If SAPVendorLineItemDisplay.SAPGuiEdit("Vendor account").GetROProperty("Value") = strVendorCode Then
	SAPVendorLineItemDisplay.SAPGuiButton("Dynamic selections   (Shift+F4").Click
Else
	SAPVendorLineItemDisplay.SAPGuiEdit("Vendor account").Set strVendorCode
	SAPVendorLineItemDisplay.SAPGuiEdit("Company code").Set strCompanyCode
	SAPVendorLineItemDisplay.SAPGuiEdit("Company code").SetFocus
	SAPVendorLineItemDisplay.SAPGuiButton("Dynamic selections   (Shift+F4").Click
End If
SAPVendorLineItemDisplay.SAPGuiTree("TableTreeControl").ActivateNode "Document;Reference"
SAPVendorLineItemDisplay.SAPGuiEdit("Reference").Set strRefNo
SAPVendorLineItemDisplay.SAPGuiEdit("Reference").SetFocus
SAPVendorLineItemDisplay.SAPGuiButton("Execute   (F8)").Click
SAP.StatusBar.Sync
If Instr(Trim(SAP.StatusBar.GetROProperty("Text")), "1 items displayed") = 0  Then
	SAP.OKCode.Set "/ns000"
	SAP.Enter.Click
	Parameter("bResult") = False
	ExitAction
End If

If SAPGuiSession("Session").SAPGuiWindow("Vendor Line Item Display_2").SAPGuiLabel("VendorName").Exist Then
	SAPGuiSession("Session").SAPGuiWindow("Vendor Line Item Display_2").SAPGuiLabel("VendorName").SetFocus
	SAPGuiSession("Session").SAPGuiWindow("Vendor Line Item Display_2").SAPGuiLabel("VendorName").SetCaretPos 4
	SAPGuiSession("Session").SAPGuiWindow("Vendor Line Item Display_2").SendKey F2
ElseIf SAPGuiSession("Session").SAPGuiWindow("Vendor Line Item Display_2").SAPGuiLabel("EECRefNum2").Exist Then
	SAPGuiSession("Session").SAPGuiWindow("Vendor Line Item Display_2").SAPGuiLabel("EECRefNum2").SetFocus
	SAPGuiSession("Session").SAPGuiWindow("Vendor Line Item Display_2").SAPGuiLabel("EECRefNum2").SetCaretPos 13
	SAPGuiSession("Session").SAPGuiWindow("Vendor Line Item Display_2").SendKey F2
Else
	SAPGuiSession("Session").SAPGuiWindow("Vendor Line Item Display_2").SAPGuiLabel("EECRefNum").SetFocus
	SAPGuiSession("Session").SAPGuiWindow("Vendor Line Item Display_2").SAPGuiLabel("EECRefNum").SetCaretPos 13
	SAPGuiSession("Session").SAPGuiWindow("Vendor Line Item Display_2").SendKey F2
End If

SAP.CallUpDocumentOverview.Click

Dim GetVendorCode, GetGLCodeExpense, GetGLCodeTax

If SAPGuiSession("Session").SAPGuiWindow("Document Overview - Display Vendor Invoice").SAPGuiLabel("nVendorCode").Exist Then
	GetVendorCode = Trim(SAPGuiSession("Session").SAPGuiWindow("Document Overview - Display Vendor Invoice").SAPGuiLabel("nVendorCode").GetROProperty("Content"))
Else
	GetVendorCode = Trim(SAPGuiSession("Session").SAPGuiWindow("Document Overview - Display Vendor Invoice").SAPGuiLabel("VendorCode").GetROProperty("Content"))
End If

If SAPGuiSession("Session").SAPGuiWindow("Document Overview - Display Vendor Invoice").SAPGuiLabel("nGLCodeExpense").Exist Then
	GetGLCodeExpense = Trim(SAPGuiSession("Session").SAPGuiWindow("Document Overview - Display Vendor Invoice").SAPGuiLabel("nGLCodeExpense").GetROProperty("Content"))
Else
	GetGLCodeExpense = Trim(SAPGuiSession("Session").SAPGuiWindow("Document Overview - Display Vendor Invoice").SAPGuiLabel("GLCodeExpense").GetROProperty("Content"))
End If

If SAPGuiSession("Session").SAPGuiWindow("Document Overview - Display Vendor Invoice").SAPGuiLabel("nGLCodeTax").Exist Then
	GetGLCodeTax = Trim(SAPGuiSession("Session").SAPGuiWindow("Document Overview - Display Vendor Invoice").SAPGuiLabel("nGLCodeTax").GetROProperty("Content"))
Else
	GetGLCodeTax = Trim(SAPGuiSession("Session").SAPGuiWindow("Document Overview - Display Vendor Invoice").SAPGuiLabel("GLCodeTax").GetROProperty("Content"))
End If

Dim TotalAmount
If SAPGuiSession("Session").SAPGuiWindow("Document Overview - Display Vendor Invoice").SAPGuiLabel("n2Amount-Total").Exist Then
	TotalAmount = Replace(Trim(SAPGuiSession("Session").SAPGuiWindow("Document Overview - Display Vendor Invoice").SAPGuiLabel("n2Amount-Total").GetROProperty("Content")), ",", "")
ElseIf SAPGuiSession("Session").SAPGuiWindow("Document Overview - Display Vendor Invoice").SAPGuiLabel("n3Amount-Total").Exist Then
		TotalAmount = Replace(Trim(SAPGuiSession("Session").SAPGuiWindow("Document Overview - Display Vendor Invoice").SAPGuiLabel("n3Amount-Total").GetROProperty("Content")), ",", "")
ElseIf SAPGuiSession("Session").SAPGuiWindow("Document Overview - Display Vendor Invoice").SAPGuiLabel("n4Amount-Total").Exist Then
		TotalAmount = Replace(Trim(SAPGuiSession("Session").SAPGuiWindow("Document Overview - Display Vendor Invoice").SAPGuiLabel("n4Amount-Total").GetROProperty("Content")), ",", "")
Else
	TotalAmount = Replace(Trim(SAPGuiSession("Session").SAPGuiWindow("Document Overview - Display Vendor Invoice").SAPGuiLabel("nAmount-Total").GetROProperty("Content")), ",", "")
End If

Dim ExpenseAmount 
If SAPGuiSession("Session").SAPGuiWindow("Document Overview - Display Vendor Invoice").SAPGuiLabel("n2Amount-Expense").Exist Then
	ExpenseAmount = Replace(Trim(SAPGuiSession("Session").SAPGuiWindow("Document Overview - Display Vendor Invoice").SAPGuiLabel("n2Amount-Expense").GetROProperty("Content")), ",", "")
'ElseIf SAPGuiSession("Session").SAPGuiWindow("Document Overview - Display Vendor Invoice").SAPGuiLabel("n3Amount-Expense").Exist Then
'		ExpenseAmount = Replace(Trim(SAPGuiSession("Session").SAPGuiWindow("Document Overview - Display Vendor Invoice").SAPGuiLabel("n3Amount-Expense").GetROProperty("Content")), ",", "")
ElseIf SAPGuiSession("Session").SAPGuiWindow("Document Overview - Display Vendor Invoice").SAPGuiLabel("n4Amount-Expense").Exist Then
		ExpenseAmount = Replace(Trim(SAPGuiSession("Session").SAPGuiWindow("Document Overview - Display Vendor Invoice").SAPGuiLabel("n4Amount-Expense").GetROProperty("Content")), ",", "")
Else
	ExpenseAmount = Replace(Trim(SAPGuiSession("Session").SAPGuiWindow("Document Overview - Display Vendor Invoice").SAPGuiLabel("nAmount-Expense").GetROProperty("Content")), ",", "")
End If

Dim TaxAmount
If SAPGuiSession("Session").SAPGuiWindow("Document Overview - Display Vendor Invoice").SAPGuiLabel("n2Amount-Tax").Exist Then
	TaxAmount = Replace(Trim(SAPGuiSession("Session").SAPGuiWindow("Document Overview - Display Vendor Invoice").SAPGuiLabel("n2Amount-Tax").GetROProperty("Content")), ",", "")
'ElseIf SAPGuiSession("Session").SAPGuiWindow("Document Overview - Display Vendor Invoice").SAPGuiLabel("n3Amount-Tax").Exist Then
'		TaxAmount = Replace(Trim(SAPGuiSession("Session").SAPGuiWindow("Document Overview - Display Vendor Invoice").SAPGuiLabel("n3Amount-Tax").GetROProperty("Content")), ",", "")
ElseIf SAPGuiSession("Session").SAPGuiWindow("Document Overview - Display Vendor Invoice").SAPGuiLabel("n4Amount-Tax").Exist Then
		TaxAmount = Replace(Trim(SAPGuiSession("Session").SAPGuiWindow("Document Overview - Display Vendor Invoice").SAPGuiLabel("n4Amount-Tax").GetROProperty("Content")), ",", "")
Else
	TaxAmount = Replace(Trim(SAPGuiSession("Session").SAPGuiWindow("Document Overview - Display Vendor Invoice").SAPGuiLabel("nAmount-Tax").GetROProperty("Content")), ",", "")
End If

SAP.OKCode.Set "/ns000"
SAP.Enter.Click

'Outputs
Parameter("strOutputs") = GetVendorCode & "," & GetGLCodeExpense & "," & GetGLCodeTax & "," & TotalAmount & "," & ExpenseAmount & "," & TaxAmount
Parameter("bResult") = True
