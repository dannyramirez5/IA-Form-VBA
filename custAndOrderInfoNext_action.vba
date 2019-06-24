Option Explicit
Dim customerID As String
Dim customerName As String
Dim orderNumber As String
Dim invoice As String
Dim entered As String
Dim requestor As String
Dim fse As String

Private Sub adjustmentPreparedByLbl_Click()
   helpAdjustmentPreparedBy.Show
End Sub

Private Sub customerAndOrderInformationBackCmndBx_Click()
   customerAndOrderInformationForm.Hide
   adjustmentTypeForm.Show
End Sub

Private Sub customerAndOrderInformationNextCmndBx_Click()
   Sheets("Completed Form").unprotect "dantheman5"
   
   customerID = customerAndOrderInformationForm.customerIDTxtBx
   customerName = customerAndOrderInformationForm.customerNameTxtBx
   orderNumber = customerAndOrderInformationForm.orderNumberTxtBx
   invoice = customerAndOrderInformationForm.invoiceNumberTxtBx
   entered = customerAndOrderInformationForm.orderEnteredByTxtBx
   requestor = customerAndOrderInformationForm.adjustmentPreparedByTxtBx
   fse = customerAndOrderInformationForm.fsetxtbx
   
   If customerID = "" Then
   MsgBox "Please enter the Customer's ID.", vbCritical, "Missing Information"
   Exit Sub
   ElseIf customerName = "" Then
   MsgBox "Please enter the Customer's Name.", vbCritical, "Missing Information"
   Exit Sub
   ElseIf orderNumber = "" Then
   MsgBox "Please enter the Order Number.", vbCritical, "Missing Information"
   Exit Sub
   ElseIf invoice = "" Then
   MsgBox "Please enter the Invoice Number.", vbCritical, "Missing Information"
   Exit Sub
   ElseIf requestor = "" Then
   MsgBox "Please enter the name of who is preparing this Adjustment.", vbCritical, "Missing Information"
   Exit Sub
   End If

   Cells(8, 1) = customerID
   Cells(9, 1) = customerName
   Cells(8, 5) = orderNumber
   Cells(9, 5) = invoice
   Cells(8, 9) = entered
   Cells(9, 9) = requestor
   Cells(10, 9) = fse
  
   compForm.customerInformationEditBtn.Visible = True
  
   customerAndOrderInformationForm.Hide
   formPreTaxAmount.Show
   
   
   Sheets("Completed Form").Protect "dantheman5"
End Sub

Private Sub customerAndOrderInformationSaveAndCloseCmndBx_Click()
   Sheets("Completed Form").unprotect "dantheman5"
   
   customerID = customerAndOrderInformationForm.customerIDTxtBx
   customerName = customerAndOrderInformationForm.customerNameTxtBx
   orderNumber = customerAndOrderInformationForm.orderNumberTxtBx
   invoice = customerAndOrderInformationForm.invoiceNumberTxtBx
   entered = customerAndOrderInformationForm.orderEnteredByTxtBx
   requestor = customerAndOrderInformationForm.adjustmentPreparedByTxtBx
   fse = customerAndOrderInformationForm.fsetxtbx
   
   If customerID = "" Then
   MsgBox "Please enter the Customer's ID.", vbCritical, "Missing Information"
   Exit Sub
   ElseIf customerName = "" Then
   MsgBox "Please enter the Customer's Name.", vbCritical, "Missing Information"
   Exit Sub
   ElseIf orderNumber = "" Then
   MsgBox "Please enter the Order Number.", vbCritical, "Missing Information"
   Exit Sub
   ElseIf invoice = "" Then
   MsgBox "Please enter the Invoice Number.", vbCritical, "Missing Information"
   Exit Sub
   ElseIf requestor = "" Then
   MsgBox "Please enter the name of who is preparing this Adjustment.", vbCritical, "Missing Information"
   Exit Sub
   End If

   Cells(8, 1) = customerID
   Cells(9, 1) = customerName
   Cells(8, 5) = orderNumber
   Cells(9, 5) = invoice
   Cells(8, 9) = entered
   Cells(9, 9) = requestor
   Cells(10, 9) = fse
   
   compForm.customerInformationEditBtn.Visible = True

   customerAndOrderInformationForm.Hide
   
   Sheets("Completed Form").Protect "dantheman5"
End Sub

Private Sub customerIDLbl_Click()
   helpCustomerID.Show
End Sub

Private Sub customerNameLbl_Click()
   helpCustomerName.Show
End Sub

Private Sub invoiceNumberLbl_Click()
   helpInvoiceNumber.Show
End Sub

Private Sub Label1_Click()
    fseHelp.Show
End Sub

Private Sub orderEnteredByLbl_Click()
   helpOrderEnteredBy.Show
End Sub

Private Sub orderNumberlbl_Click()
   helpOrderNumber.Show
End Sub
