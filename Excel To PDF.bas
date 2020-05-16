Sub SheetsToPDFs()

    Dim strPath As String
    Dim wks As Worksheet

    strPath = ActiveWorkbook.Path & "\"
    
    For Each wks In ActiveWorkbook.Worksheets
        If wks.Name = "List of Defects" Or wks.Name = "PDF OUT" Then
        Else
        wks.ExportAsFixedFormat xlTypePDF, strPath & Left(ActiveWorkbook.Name, (InStrRev(ActiveWorkbook.Name, ".", -1, vbTextCompare) - 4)) & wks.Name & ".pdf"
        End If
    Next wks
    MsgBox "Tüm pdfler baþarýlý bir " & vbNewLine & "þekilde çýkarýlmýþtýr."
        
End Sub
