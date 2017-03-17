'--------------------------------------------------------------------------------------------------'
' GetSheetUsedArea                                                                                 '
'--------------------------------------------------------------------------------------------------'
' Returns sheet used area.                                                                         '
'                                                                                                  '
' Parameters:                                                                                      '
'                                                                                                  '
'   Optional TargetSheet As Object <Default = ThisComponent.CurrentController.ActiveSheet>         '
'     Reference to a sheet (com.sun.star.sheet.XSpreadsheet).                                      '
'                                                                                                  '
' Examples:                                                                                        '
'--------------------------------------------------------------------------------------------------'
'                                                                                                  '
'     used_area = GetSheetUsedArea()                                                               '
' or                                                                                               '
'     used_area = GetSheetUsedArea(0)                                                              '
' or                                                                                               '
'     used_area = GetSheetUsedArea("Sheet1")                                                       '
' or                                                                                               '
'     used_area = GetSheetUsedArea(ThisComponent.Sheets.getByIndex(0))                             '
'                                                                                                  '
' Expected values:                                                                                 '
'   used_area:                                                                                     '
'     com.sun.star.table.CellRangeAddress                                                          '
'       .Sheet       integer Sheet index                                                           '
'       .StartColumn long    Start column index                                                    '
'       .StartRow    long    Start row index                                                       '
'       .EndColumn   long    End column index                                                      '
'       .EndRow      long    End row index                                                         '
'--------------------------------------------------------------------------------------------------'
' See also:                                                                                        '
'   https://wiki.documentfoundation.org/Macros/Calc/001/fr                                         '
'   https://openoffice.org/api/docs/common/ref/com/sun/star/sheet/XUsedAreaCursor.html             '
'   http://api.libreoffice.org/docs/idl/ref/structcom_1_1sun_1_1star_1_1table_1_1CellRangeAddress.html   '
'   http://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1sheet_1_1XUsedAreaCursor.html '
'--------------------------------------------------------------------------------------------------'
' Feedback & Issues:                                                                               '
'   https://github.com/aa6/libreoffice_calc_basic_extras/issues                                    '
'--------------------------------------------------------------------------------------------------'
Function GetSheetUsedArea(Optional TargetSheet As Variant) As Object
    
    Dim sheet As Object
    Dim cursor As Object

    Select Case TRUE
        Case IsMissing(TargetSheet)
            sheet = ThisComponent.CurrentController.ActiveSheet
        Case TypeName(TargetSheet) = "String"
            sheet = ThisComponent.Sheets.getByName(TargetSheet)
        Case TypeName(TargetSheet) = "Number"
            sheet = ThisComponent.Sheets.getByIndex(TargetSheet)
        Case Else
            sheet = TargetSheet
    End Select
    
    cursor = sheet.CreateCursor()
    cursor.gotoStartOfUsedArea(FALSE) ' FALSE sets cursor size to a 1x1 cell. '
    cursor.gotoEndOfUsedArea(TRUE)    ' TRUE expands cursor range.            '
    GetSheetUsedArea = cursor.getRangeAddress()
    
End Function