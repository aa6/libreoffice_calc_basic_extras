' Returns sheet used area. Return value has the next properties:                                   '
'   IDL class: com.sun.star.table.CellRangeAddress                                                 '
'   .Sheet       -- Sheet index                                                                    '
'   .StartColumn -- Start column index                                                             '
'   .StartRow    -- Start row index                                                                '
'   .EndColumn   -- End column index                                                               '
'   .EndRow      -- End row index                                                                  '
' See also:                                                                                        '
'   https://wiki.documentfoundation.org/Macros/Calc/001/fr                                         '
'   https://openoffice.org/api/docs/common/ref/com/sun/star/sheet/XUsedAreaCursor.html             '
'   http://api.libreoffice.org/docs/idl/ref/structcom_1_1sun_1_1star_1_1table_1_1CellRangeAddress.html '
'   http://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1sheet_1_1XUsedAreaCursor.html '
Function GetSheetUsedArea(Sheet As Object) As Object
    
    Dim cursor As Object
    
    cursor = Sheet.CreateCursor()
    cursor.gotoStartOfUsedArea(FALSE) ' FALSE sets cursor size to a 1x1 cell. '
    cursor.gotoEndOfUsedArea(TRUE)    ' TRUE expands cursor range.            '
    GetSheetUsedArea = cursor.getRangeAddress()
    
End Function