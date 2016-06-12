'--------------------------------------------------------------------------------------------------'
' UnMergeAndNormalizeCell                                                                          '
'--------------------------------------------------------------------------------------------------'
' Unmerges merged cell and fill the unmerged cell with merged cell value or formula.               '
'                                                                                                  '
' Parameters:                                                                                      '
'                                                                                                  '
'   Cell As Variant                                                                                '
'     Reference to a range or cell (com.sun.star.table.XCellRange / com.sun.star.table.XCell) or a '
'     string name ("B5","R1C1", etc).                                                              '
'                                                                                                  '
'   Optional FillHorizontally As Boolean <Default = TRUE>                                          '
'     If set to FALSE, unmerged cells will NOT be filled with a merged cell value horizontally.    '
'                                                                                                  '
'   Optional FillVertically As Boolean <Default = TRUE>                                            '
'     If set to FALSE, unmerged cells will NOT be filled with a merged cell value vertically.      '
'                                                                                                  '
' Examples:                                                                                        '
'--------------------------------------------------------------------------------------------------'
'                                                                                                  '
'     UnMergeAndNormalizeCell("B5")                                                                '
'     UnMergeAndNormalizeCell("R1C1")                                                              '
'     UnMergeAndNormalizeCell("$'Sheet.name.with.dots'.$G$9")                                      '
'     UnMergeAndNormalizeCell(ThisComponent.getCurrentSelection())                                 '
'     UnMergeAndNormalizeCell(ThisComponent.Sheets.getByIndex(0).getCellByPosition(6,4))           '
'                                                                                                  '
'--------------------------------------------------------------------------------------------------'
'                                                                                                  '
'     UnMergeAndNormalizeCell("A1")                                                                '
'                                                                                                  '
' Expected results:                                                                                '
'                                                                                                  '
'   +===+===================+===+===+===+===+                                                      '
'   |   |         A         | B | C | D | E |                                                      '
'   +===+===================+===+===+===+===+                                                      '
'   | 1 |                           |   |   |                                                      '
'   +---+                           +---+---+                                                      '
'   | 2 |     Merged cell value     |   |   |                                                      '
'   +---+                           +---+---+                                                      '
'   | 3 |                           |   |   |                                                      '
'   +---+-------------------+---+---+---+---+                                                      '
'   | 4 |                   |   |   |   |   |                                                      '
'   +---+-------------------+---+---+---+---+                                                      '
'                                                                                                  '
'   +===+===================+===================+===================+===+===+                      '
'   |   |         A         |         B         |         C         | D | E |                      '
'   +===+===================+===================+===================+===+===+                      '
'   | 1 | Merged cell value | Merged cell value | Merged cell value |   |   |                      '
'   +---+-------------------+-------------------+-------------------+---+---+                      '
'   | 2 | Merged cell value | Merged cell value | Merged cell value |   |   |                      '
'   +---+-------------------+-------------------+-------------------+---+---+                      '
'   | 3 | Merged cell value | Merged cell value | Merged cell value |   |   |                      '
'   +---+-------------------+-------------------+-------------------+---+---+                      '
'   | 4 |                   |                   |                   |   |   |                      '
'   +---+-------------------+-------------------+-------------------+---+---+                      '
'                                                                                                  '
'--------------------------------------------------------------------------------------------------'
'                                                                                                  '
'     UnMergeAndNormalizeCell("A1",FALSE)                                                          '
'                                                                                                  '
' Expected results:                                                                                '
'                                                                                                  '
'   +===+===================+===+===+===+===+                                                      '
'   |   |         A         | B | C | D | E |                                                      '
'   +===+===================+===+===+===+===+                                                      '
'   | 1 |                           |   |   |                                                      '
'   +---+                           +---+---+                                                      '
'   | 2 |     Merged cell value     |   |   |                                                      '
'   +---+                           +---+---+                                                      '
'   | 3 |                           |   |   |                                                      '
'   +---+-------------------+---+---+---+---+                                                      '
'   | 4 |                   |   |   |   |   |                                                      '
'   +---+-------------------+---+---+---+---+                                                      '
'                                                                                                  '
'   +===+===================+===+===+===+===+                                                      '
'   |   |         A         | B | C | D | E |                                                      '
'   +===+===================+===+===+===+===+                                                      '
'   | 1 | Merged cell value |   |   |   |   |                                                      '
'   +---+-------------------+---+---+---+---+                                                      '
'   | 2 | Merged cell value |   |   |   |   |                                                      '
'   +---+-------------------+---+---+---+---+                                                      '
'   | 3 | Merged cell value |   |   |   |   |                                                      '
'   +---+-------------------+---+---+---+---+                                                      '
'   | 4 |                   |   |   |   |   |                                                      '
'   +---+-------------------+---+---+---+---+                                                      '
'                                                                                                  '
'--------------------------------------------------------------------------------------------------'
'                                                                                                  '
'     UnMergeAndNormalizeCell("A1",TRUE,FALSE)                                                     '
'                                                                                                  '
' Expected results:                                                                                '
'                                                                                                  '
'   +===+===================+===+===+===+===+                                                      '
'   |   |         A         | B | C | D | E |                                                      '
'   +===+===================+===+===+===+===+                                                      '
'   | 1 |                           |   |   |                                                      '
'   +---+                           +---+---+                                                      '
'   | 2 |     Merged cell value     |   |   |                                                      '
'   +---+                           +---+---+                                                      '
'   | 3 |                           |   |   |                                                      '
'   +---+-------------------+---+---+---+---+                                                      '
'   | 4 |                   |   |   |   |   |                                                      '
'   +---+-------------------+---+---+---+---+                                                      '
'                                                                                                  '
'   +===+===================+===================+===================+===+===+                      '
'   |   |         A         |         B         |         C         | D | E |                      '
'   +===+===================+===================+===================+===+===+                      '
'   | 1 | Merged cell value | Merged cell value | Merged cell value |   |   |                      '
'   +---+-------------------+-------------------+-------------------+---+---+                      '
'   | 2 |                   |                   |                   |   |   |                      '
'   +---+-------------------+-------------------+-------------------+---+---+                      '
'   | 3 |                   |                   |                   |   |   |                      '
'   +---+-------------------+-------------------+-------------------+---+---+                      '
'   | 4 |                   |                   |                   |   |   |                      '
'   +---+-------------------+-------------------+-------------------+---+---+                      '
'                                                                                                  '
'--------------------------------------------------------------------------------------------------'
'                                                                                                  '
'     UnMergeAndNormalizeCell("A1",FALSE,FALSE)                                                    '
'                                                                                                  '
' Expected results:                                                                                '
'                                                                                                  '
'   +===+===================+===+===+===+===+                                                      '
'   |   |         A         | B | C | D | E |                                                      '
'   +===+===================+===+===+===+===+                                                      '
'   | 1 |                           |   |   |                                                      '
'   +---+                           +---+---+                                                      '
'   | 2 |     Merged cell value     |   |   |                                                      '
'   +---+                           +---+---+                                                      '
'   | 3 |                           |   |   |                                                      '
'   +---+-------------------+---+---+---+---+                                                      '
'   | 4 |                   |   |   |   |   |                                                      '
'   +---+-------------------+---+---+---+---+                                                      '
'                                                                                                  '
'   +===+===================+===+===+===+===+                                                      '
'   |   |         A         | B | C | D | E |                                                      '
'   +===+===================+===+===+===+===+                                                      '
'   | 1 | Merged cell value |   |   |   |   |                                                      '
'   +---+-------------------+---+---+---+---+                                                      '
'   | 2 |                   |   |   |   |   |                                                      '
'   +---+-------------------+---+---+---+---+                                                      '
'   | 3 |                   |   |   |   |   |                                                      '
'   +---+-------------------+---+---+---+---+                                                      '
'   | 4 |                   |   |   |   |   |                                                      '
'   +---+-------------------+---+---+---+---+                                                      '
'                                                                                                  '
'--------------------------------------------------------------------------------------------------'
' Feedback & Issues:                                                                               '
'   https://github.com/aa6/libreoffice_calc_basic_extras/issues                                    '
'--------------------------------------------------------------------------------------------------'
Sub UnMergeAndNormalizeCell(Cell As Variant, Optional FillHorizontally As Boolean, Optional FillVertically As Boolean)
    
    Dim col As Integer
    Dim row As Integer
    Dim args0() As Object
    Dim args1(0) As New com.sun.star.beans.PropertyValue    
    Dim args2(1) As New com.sun.star.beans.PropertyValue    
    Dim args3(0) As New com.sun.star.beans.PropertyValue    
    Dim rowspan As Integer
    Dim colspan As Integer
    Dim dispatcher As Object
    Dim target_cell As Object
    Dim previous_selection As Object

    previous_selection = ThisComponent.getCurrentSelection()
    dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
    
    args1(0).Name = "ToPoint"
    If TypeName(Cell) = "String" Then
        args1(0).Value = Cell
    Else 
        args1(0).Value = IIf(TRUE,Cell,Cell).AbsoluteName ' `Object variable not set.` workaround. '
    End If
    args2(0).Name = "By"
    args2(0).Value = 1
    args2(1).Name = "Sel"
    args2(1).Value = false

    ' Get rowspan. '
    dispatcher.executeDispatch(ThisComponent.CurrentController.Frame, ".uno:GoToCell", "", 0, args1())
    target_cell = ThisComponent.getCurrentSelection()
    dispatcher.executeDispatch(ThisComponent.CurrentController.Frame, ".uno:GoDown", "", 0, args2())
    rowspan = ThisComponent.getCurrentSelection().CellAddress.Row - target_cell.CellAddress.Row

    ' Get colspan. '
    dispatcher.executeDispatch(ThisComponent.CurrentController.Frame, ".uno:GoToCell", "", 0, args1())
    dispatcher.executeDispatch(ThisComponent.CurrentController.Frame, ".uno:GoRight", "", 0, args2())
    colspan = ThisComponent.getCurrentSelection().CellAddress.Column - target_cell.CellAddress.Column

    ' Set cursor back to its original position. '
    args3(0).Name = "ToPoint"
    args3(0).Value = previous_selection.AbsoluteName
    dispatcher.executeDispatch(ThisComponent.CurrentController.Frame, ".uno:GoToCell", "", 0, args3())

    ' Unmerge cell and fill with values. '
    target_cell.merge(FALSE)
    For row = 0 To IIf(FillVertically = FALSE, 0 ,rowspan - 1)
        For col = 0 To IIf(FillHorizontally = FALSE, 0, colspan - 1)
            target_cell.Spreadsheet.getCellByPosition(      _
                target_cell.RangeAddress.StartColumn + col, _
                target_cell.RangeAddress.StartRow + row     _
            ).setFormulaArray(target_cell.formulaArray)
        Next col
    Next row
  
End Sub