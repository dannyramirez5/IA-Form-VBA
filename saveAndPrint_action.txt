Private Sub CommandButton1_Click()
    emailorsaveform.Hide
    formEmail.Show
End Sub

Private Sub CommandButton2_Click()

    Sheets("Completed Form").unprotect "password"

    Dim PdfFile As String, Title As String
    Dim i As Long

    Title = Range("E8")
 
    PdfFile = Title & " Invoice Adjustment"
    i = InStrRev(PdfFile, ".")
    If i > 1 Then PdfFile = Left(PdfFile, i - 1)
    PdfFile = PdfFile & ".pdf"
    
    With compForm
        .ExportAsFixedFormat Type:=xlTypePDF, Filename:="V:\IA Forms\" & PdfFile, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    End With

    Sheets("Completed Form").Protect "password"
    
    emailorsaveform.Hide
    
    MsgBox "Your invoice adjustment has been created and saved in the Invoice Adjustment folder (V:\IA Forms)."

End Sub
