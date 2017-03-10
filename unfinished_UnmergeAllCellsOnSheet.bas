Sub UnmergeAllCellsOnSheet(sheet As com.sun.star.sheet.XSpreadsheet)
  usedarea = GetSheetUsedArea(sheet)
  For row = usedarea.StartRow To usedarea.EndRow Step 1
    For col = usedarea.StartColumn To usedarea.EndColumn Step 1
        cell = sheet.getCellByPosition(col,row)
        If cell.IsMerged Then cell.merge(FALSE)'UnMergeAndNormalizeCell(cell)       
      Next col
    Next row
End Sub