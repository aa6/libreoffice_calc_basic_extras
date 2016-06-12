'--------------------------------------------------------------------------------------------------'
' MergedCellGetColSpanCount                                                                        '
'--------------------------------------------------------------------------------------------------'
' Returns a merged cell column span count.                                                         '
'                                                                                                  '
' Parameters:                                                                                      '
'                                                                                                  '
'   Cell As Variant                                                                                '
'     Reference to a range or cell (com.sun.star.table.XCellRange / com.sun.star.table.XCell) or a '
'     string name ("B5","R1C1", etc).                                                              '
'                                                                                                  '
'   Optional FailIfNotMerged As Boolean <Default = FALSE>                                          '
'     If set to TRUE, the function will return -1 if given Cell is not merged.                     '
'                                                                                                  '
' Examples:                                                                                        '
'--------------------------------------------------------------------------------------------------'
'                                                                                                  '
'     colspan = MergedCellGetColSpanCount("B5")                                                    '
'     colspan = MergedCellGetColSpanCount("R1C1")                                                  '
'     colspan = MergedCellGetColSpanCount("$'Sheet.name.with.dots'.$G$9")                          '
'     colspan = MergedCellGetColSpanCount(ThisComponent.getCurrentSelection())                     '
' or                                                                                               '
'     cell = ThisComponent.Sheets.getByIndex(0).getCellByPosition(6,4)                             '
'     colspan = MergedCellGetColSpanCount(cell)                                                    '
'                                                                                                  '
'--------------------------------------------------------------------------------------------------'
'                                                                                                  '
'     colspan = MergedCellGetColSpanCount("B5",TRUE)                                               '
'                                                                                                  '
' Will return -1 if B5 is not a merged cell.                                                       '
'--------------------------------------------------------------------------------------------------'
' Feedback & Issues:                                                                               '
'   https://github.com/aa6/libreoffice_calc_basic_extras/issues                                    '
'--------------------------------------------------------------------------------------------------'
Function MergedCellGetColSpanCount(Cell As Variant, Optional FailIfNotMerged As Boolean)
    
    Dim args1(0) As New com.sun.star.beans.PropertyValue    
    Dim args2(1) As New com.sun.star.beans.PropertyValue    
    Dim args3(0) As New com.sun.star.beans.PropertyValue    
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
    dispatcher.executeDispatch(ThisComponent.CurrentController.Frame, ".uno:GoToCell", "", 0, args1())
    target_cell = ThisComponent.getCurrentSelection()
    
    If FailIfNotMerged = TRUE AND NOT target_cell.IsMerged Then
        MergedCellGetColSpanCount = -1
    Else
        args2(0).Name = "By"
        args2(0).Value = 1
        args2(1).Name = "Sel"
        args2(1).Value = false
        dispatcher.executeDispatch(ThisComponent.CurrentController.Frame, ".uno:GoRight", "", 0, args2())
        MergedCellGetColSpanCount = ThisComponent.getCurrentSelection().CellAddress.Column - target_cell.CellAddress.Column
    End If
        
    args3(0).Name = "ToPoint"
    args3(0).Value = previous_selection.AbsoluteName
    dispatcher.executeDispatch(ThisComponent.CurrentController.Frame, ".uno:GoToCell", "", 0, args3())
  
End Function