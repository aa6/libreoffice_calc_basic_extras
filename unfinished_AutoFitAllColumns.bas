Sub AutoFitAllColumnsOnSheet(Optional sheet As com.sun.star.sheet.XSpreadsheet)
    If IsMissing(sheet) Then
        sheet = ThisComponent.getCurrentSelection.Spreadsheet
    End If
    ecols = sheet.Columns.createEnumeration()
    while ecols.hasMoreElements()
        c = ecols.nextElement()
        c.OptimalWidth = TRUE
    wend
End Sub