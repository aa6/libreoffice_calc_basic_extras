' Sets cursor to a cell specified by name. GoToCell would've been a shorter name for this function '
' but I found `Go` a bit confusing because it has unclear meaning.                                 '
' Examples:                                                                                        '
'   SetCursorToCell("B5")                                                                          '
'   SetCursorToCell("R1C1")                                                                        '
'   SetCursorToCell("$'Sheet.name.with.dots'.$G$9")                                                '
Sub SetCursorToCell(CellName As String)

    Dim args(0) As New com.sun.star.beans.PropertyValue
    Dim dispatcher As Object

    args(0).Name = "ToPoint"
    args(0).Value = CellName
    dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
    dispatcher.executeDispatch(ThisComponent.CurrentController.Frame, ".uno:GoToCell", "", 0, args())

End Sub