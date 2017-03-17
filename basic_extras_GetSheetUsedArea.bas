'--------------------------------------------------------------------------------------------------'
' GetSheetUsedArea                                                                                 '
'--------------------------------------------------------------------------------------------------'
' Returns sheet used area.                                                                         '
'                                                                                                  '
' Parameters:                                                                                      '
'                                                                                                  '
'   Optional TargetSheet As Object <Default = ThisComponent.CurrentController.ActiveSheet>         '
'       Reference to a sheet (com.sun.star.sheet.XSpreadsheet).                                    '
'                                                                                                  '
' Return value:                                                                                    '
'                                                                                                  '
'   GetSheetUsedAreaResultObject                                                                   '
'                                                                                                  '
'       .Sheet As com.sun.star.sheet.XSpreadsheet                                                  '
'           Target sheet.                                                                          '
'                                                                                                  '
'       .SheetIndex As Integer                                                                     '
'       .SheetNumber As Integer                                                                    '
'           Target sheet index.                                                                    '
'                                                                                                  '
'       .StartRow As Long                                                                          '
'           Used area start row index.                                                             '
'                                                                                                  '
'       .StartCol As Long                                                                          '
'       .StartColumn As Long                                                                       '
'           Used area start column index.                                                          '
'                                                                                                  '
'       .EndRow As Long                                                                            '
'           Used area end row index.                                                               '
'                                                                                                  '
'       .EndCol As Long                                                                            '
'       .EndColumn As Long                                                                         '
'           Used area end column index.                                                            '
'                                                                                                  '
'       .RangeAddress As com.sun.star.table.CellRangeAddress                                       '
'       .CellRangeAddress As com.sun.star.table.CellRangeAddress                                   '
'                                                                                                  '
'           .Sheet As Integer                                                                      '
'               Target sheet index.                                                                '
'                                                                                                  '
'           .StartColumn As Long                                                                   '
'               Used area start column index.                                                      '
'                                                                                                  '
'           .StartRow As Long                                                                      '
'               Used area start row index.                                                         '
'                                                                                                  '
'           .EndColumn As Long                                                                     '
'               Used area end column index.                                                        '
'                                                                                                  '
'           .EndRow As Long                                                                        '
'               Used area end row index.                                                           '
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
Type GetSheetUsedAreaResultObject

    Sheet As Object
    SheetIndex As Integer
    SheetNumber As Integer
    StartRow As Long
    StartCol As Long
    StartColumn As Long
    EndRow As Long
    EndCol As Long
    EndColumn As Long
    RangeAddress As com.sun.star.table.CellRangeAddress
    CellRangeAddress As com.sun.star.table.CellRangeAddress
  
End Type
Function GetSheetUsedArea(Optional TargetSheet As Variant) As GetSheetUsedAreaResultObject
    
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
    GetSheetUsedArea = New GetSheetUsedAreaResultObject
    GetSheetUsedArea.RangeAddress = cursor.getRangeAddress()
    GetSheetUsedArea.CellRangeAddress = GetSheetUsedArea.RangeAddress
    GetSheetUsedArea.Sheet = sheet
    GetSheetUsedArea.SheetIndex = GetSheetUsedArea.CellRangeAddress.Sheet
    GetSheetUsedArea.SheetNumber = GetSheetUsedArea.SheetIndex
    GetSheetUsedArea.StartRow = GetSheetUsedArea.CellRangeAddress.StartRow
    GetSheetUsedArea.StartCol = GetSheetUsedArea.CellRangeAddress.StartColumn
    GetSheetUsedArea.StartColumn = GetSheetUsedArea.CellRangeAddress.StartColumn
    GetSheetUsedArea.EndRow = GetSheetUsedArea.CellRangeAddress.EndRow
    GetSheetUsedArea.EndCol = GetSheetUsedArea.CellRangeAddress.EndColumn
    GetSheetUsedArea.EndColumn = GetSheetUsedArea.CellRangeAddress.EndColumn
   
End Function