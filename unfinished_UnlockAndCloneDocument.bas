
Function CloneNewUnlockedDocument(Optional document As Object)
    ' Create new document
    If IsMissing(document) Then
        document = ThisComponent
    End If
    
    new_doc = StarDesktop.loadComponentFromURL("private:factory/scalc","_blank",0,Array())
    ' Delete all sheets (you can't delete all but you can left at least one).
    For Each sheetname In new_doc.Sheets.ElementNames
        If new_doc.Sheets.getByName(sheetname).RangeAddress.Sheet <> 0 Then ' If sheet index isn't 0.
            new_doc.Sheets.RemoveByName(sheetname)
        End If
    Next sheetname
    ' Rename first sheet and insert the rest.
    For Each sheetname In document.Sheets.ElementNames
        If document.Sheets.getByName(sheetname).RangeAddress.Sheet = 0 Then ' Rename
            new_doc.Sheets.getByIndex(0).Name = sheetname
        Else
            new_doc.Sheets.insertNewByName(sheetname,new_doc.Sheets.Count)
        End If      
    Next sheetname
    ' Copy data from protected sheets to new unprotected document.
    'xray document.Sheets(0)
    dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
    'document_frame = document.CurrentController.getFrame()
    For Each sheetname In document.Sheets.ElementNames
        document.CurrentController.Select(document.Sheets.getByName(sheetname))
        dispatcher.executeDispatch(document.CurrentController.Frame,".uno:SelectAll","",0,Array())
        dispatcher.executeDispatch(document.CurrentController.Frame,".uno:Copy","",0,Array())
        new_doc.CurrentController.Select(new_doc.Sheets.getByName(sheetname))
        dispatcher.executeDispatch(new_doc.CurrentController.Frame,".uno:Paste","",0,Array())
        'usedarea = GetSheetUsedArea(sheet)
        'usedareaname = "R" + (usedarea.StartRow+1) + "C" + (usedarea.StartColumn+1) + ":" + "R" + (usedarea.EndRow+1) + "C" + (usedarea.EndColumn+1)
        'msgbox usedareaname
        'xray sheet.getCellRangeByName(usedareaname)
        'data = sheet.getCellRangeByName(usedareaname).getDataArray()
        'new_doc.Sheets.getByName(sheetname).getCellRangeByName(usedareaname).setDataArray(data)
    Next sheetname
    'For Each sheetname In doc_remnants.Sheets.ElementNames
    '   sheet = document.Sheets.getByName(sheetname)
    '   new_doc.Sheets.insertNewByName(sheetname,0)
    'Next sheetname
    
    CloneNewUnlockedDocument = new_doc
End Function
