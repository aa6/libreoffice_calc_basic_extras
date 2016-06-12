'--------------------------------------------------------------------------------------------------'
' SetCursorToCell                                                                                  '
'--------------------------------------------------------------------------------------------------'
' Sets cursor to a cell specified by a cell name or reference object.                              '
' GoToCell would've been a shorter name for this function but I found `Go` a bit confusing because '
' it has unclear meaning.                                                                          '
'                                                                                                  '
' Parameters:                                                                                      '
'                                                                                                  '
'   Cell As Variant                                                                                '
'     Reference to a range or cell (com.sun.star.table.XCellRange / com.sun.star.table.XCell) or a '
'     string name ("B5","R1C1", etc).                                                              '
'                                                                                                  '
' Examples:                                                                                        '
'--------------------------------------------------------------------------------------------------'
'                                                                                                  '
'     SetCursorToCell("B5")                                                                        '
'     SetCursorToCell("R1C1")                                                                      '
'     SetCursorToCell("$'Sheet.name.with.dots'.$G$9")                                              '
'     SetCursorToCell(ThisComponent.Sheets.getByIndex(0).getCellByPosition(6,4))                   '
'                                                                                                  '
'--------------------------------------------------------------------------------------------------'
' Feedback & Issues:                                                                               '
'   https://github.com/aa6/libreoffice_calc_basic_extras/issues                                    '
'--------------------------------------------------------------------------------------------------'
Sub SetCursorToCell(Cell As Variant)

    Dim args(0) As New com.sun.star.beans.PropertyValue
    Dim dispatcher As Object

    args(0).Name = "ToPoint"
    If TypeName(Cell) = "String" Then
        args(0).Value = Cell
    Else 
        args(0).Value = IIf(TRUE,Cell,Cell).AbsoluteName ' `Object variable not set.` workaround. '
    End If

    dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
    dispatcher.executeDispatch(ThisComponent.CurrentController.Frame, ".uno:GoToCell", "", 0, args())

End Sub