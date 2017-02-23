'--------------------------------------------------------------------------------------------------'
' DuplicateRowsByCells                                                                             '
'--------------------------------------------------------------------------------------------------'
' Pumpurumpum.                                                                                     '
'                                                                                                  '
' Parameters:                                                                                      '
'                                                                                                  '
'   Arr As Variant                                                                                 '
'     The input array.                                                                             '
'                                                                                                  '
' Examples:                                                                                        '
'--------------------------------------------------------------------------------------------------'
'                                                                                                  '
' Input:                                                                                           '
'                                                                                                  '
'   +===+=========+=========+=========+=========+                                                  '
'   |   |    A    |    B    |    C    |    D    |                                                  '
'   +===+=========+=========+=========+=========+                                                  '
'   | 1 | Title 5 | Title 6 | Title 7 | Title 8 |                                                  '
'   +---+---------+---------+---------+---------+                                                  '
'   | 2 | Val 251 | Val 261 | Val 271 | Val 281 |                                                  '
'   +---+---------+---------+---------+---------+                                                  '
'   | 3 | Val 252 | Ral 262 | Val 272 | Val 282 |                                                  '
'   +---+---------+---------+---------+---------+                                                  '
'                                                                                                  '
' Code:                                                                                            '
'                                                                                                  '
'     duplicated_cells = DuplicateRowsByCells(Array(0,1),2)                                        '
'     For counter = 0 To UBound(duplicated_cells)                                                  '
'         cell = duplicated_cells(counter)                                                         '
'         cell.setString(cell.getString() + " " + counter)                                         '
'     Next counter                                                                                 '
'                                                                                                  '
' Output:                                                                                          '
'                                                                                                  '
'   +===+===========+=========+=========+=========+                                                '
'   |   |     A     |    B    |    C    |    D    |                                                '
'   +===+===========+=========+=========+=========+                                                '
'   | 1 | Title 5 2 | Title 6 | Title 7 | Title 8 |                                                '
'   +---+-----------+---------+---------+---------+                                                '
'   | 2 | Title 5 3 | Title 6 | Title 7 | Title 8 |                                                '
'   +---+-----------+---------+---------+---------+                                                '
'   | 3 | Val 251 0 | Val 261 | Val 271 | Val 281 |                                                '
'   +---+-----------+---------+---------+---------+                                                '
'   | 4 | Val 251 1 | Val 261 | Val 271 | Val 281 |                                                '
'   +---+-----------+---------+---------+---------+                                                '
'   | 5 | Val 252   | Ral 262 | Val 272 | Val 282 |                                                '
'   +---+-----------+---------+---------+---------+                                                '
'                                                                                                  '
'--------------------------------------------------------------------------------------------------'
'                                                                                                  '
' Input:                                                                                           '
'                                                                                                  '
'   +===+=========+=========+=========+=========+                                                  '
'   |   |    A    |    B    |    C    |    D    |                                                  '
'   +===+=========+=========+=========+=========+                                                  '
'   | 1 | Title 5 | Title 6 | Title 7 | Title 8 |                                                  '
'   +---+---------+---------+---------+---------+                                                  '
'   | 2 | Val 251 | Val 261 | Val 271 | Val 281 |                                                  '
'   +---+---------+---------+---------+---------+                                                  '
'   | 3 | Val 252 | Ral 262 | Val 272 | Val 282 |                                                  '
'   +---+---------+---------+---------+---------+                                                  '
'                                                                                                  '
' Code:                                                                                            '
'                                                                                                  '
'     srch = FindStr(Array("Val 261","Ral"),ThisComponent,Array("substr","prepared","reversed"))   '
'     While FindStr(srch).HasResults                                                               '
'         duplicated_cells = DuplicateRowsByCells(srch.cell,3)                                     '
'         For counter = 0 To UBound(duplicated_cells)                                              '
'             cell = duplicated_cells(counter)                                                     '
'             cell.setString(cell.getString() + " " + counter)                                     '
'         Next counter                                                                             '
'     Wend                                                                                         '
'                                                                                                  '
' Output:                                                                                          '
'                                                                                                  '
'   +===+=========+===========+=========+=========+                                                '
'   |   |    A    |     B     |    C    |    D    |                                                '
'   +===+=========+===========+=========+=========+                                                '
'   | 1 | Title 5 | Title 6   | Title 7 | Title 8 |                                                '
'   +---+---------+-----------+---------+---------+                                                '
'   | 2 | Val 251 | Val 261 0 | Val 271 | Val 281 |                                                '
'   +---+---------+-----------+---------+---------+                                                '
'   | 3 | Val 251 | Val 261 1 | Val 271 | Val 281 |                                                '
'   +---+---------+-----------+---------+---------+                                                '
'   | 4 | Val 251 | Val 261 2 | Val 271 | Val 281 |                                                '
'   +---+---------+-----------+---------+---------+                                                '
'   | 5 | Val 252 | Ral 262 0 | Val 272 | Val 282 |                                                '
'   +---+---------+-----------+---------+---------+                                                '
'   | 6 | Val 252 | Ral 262 1 | Val 272 | Val 282 |                                                '
'   +---+---------+-----------+---------+---------+                                                '
'   | 7 | Val 252 | Ral 262 2 | Val 272 | Val 282 |                                                '
'   +---+---------+-----------+---------+---------+                                                '
'                                                                                                  '
'--------------------------------------------------------------------------------------------------'
'                                                                                                  '
' Input:                                                                                           '
'                                                                                                  '
'   +===+=========+=========+=========+=========+                                                  '
'   |   |    A    |    B    |    C    |    D    |                                                  '
'   +===+=========+=========+=========+=========+                                                  '
'   | 1 | Title 5 | Title 6 | Title 7 | Title 8 |                                                  '
'   +---+---------+---------+---------+---------+                                                  '
'   | 2 | Val 251 | Val 261 | Val 271 | Val 281 |                                                  '
'   +---+---------+---------+---------+---------+                                                  '
'   | 3 | Val 252 | Ral 262 | Val 272 | Val 282 |                                                  '
'   +---+---------+---------+---------+---------+                                                  '
'                                                                                                  '
' Code:                                                                                            '
'                                                                                                  '
'     Dim cells() As Variant                                                                       '
'     srch = FindStr(Array("Val 261","Ral"), ThisComponent, Array("substr","prepared","reversed")) '
'     While FindStr(srch).HasResults                                                               '
'         Redim Preserve cells(UBound(cells) + 1) As Variant                                       '
'         cells(UBound(cells)) = srch.Cell                                                         '
'     Wend                                                                                         '
'     duplicated_cells = DuplicateRowsByCells(cells,3)                                             '
'     For counter = 0 To UBound(duplicated_cells)                                                  '
'         cell = duplicated_cells(counter)                                                         '
'         cell.setString(cell.getString() + " " + counter)                                         '
'     Next counter                                                                                 '
'                                                                                                  '
' Output:                                                                                          '
'                                                                                                  '
'   +===+=========+===========+=========+=========+                                                '
'   |   |    A    |     B     |    C    |    D    |                                                '
'   +===+=========+===========+=========+=========+                                                '
'   | 1 | Title 5 | Title 6   | Title 7 | Title 8 |                                                '
'   +---+---------+-----------+---------+---------+                                                '
'   | 2 | Val 251 | Val 261 3 | Val 271 | Val 281 |                                                '
'   +---+---------+-----------+---------+---------+                                                '
'   | 3 | Val 251 | Val 261 4 | Val 271 | Val 281 |                                                '
'   +---+---------+-----------+---------+---------+                                                '
'   | 4 | Val 251 | Val 261 5 | Val 271 | Val 281 |                                                '
'   +---+---------+-----------+---------+---------+                                                '
'   | 5 | Val 252 | Ral 262 0 | Val 272 | Val 282 |                                                '
'   +---+---------+-----------+---------+---------+                                                '
'   | 6 | Val 252 | Ral 262 1 | Val 272 | Val 282 |                                                '
'   +---+---------+-----------+---------+---------+                                                '
'   | 7 | Val 252 | Ral 262 2 | Val 272 | Val 282 |                                                '
'   +---+---------+-----------+---------+---------+                                                '
'                                                                                                  '
' Code (cells are not reversed but the function reverse them by itself):                           '
'                                                                                                  '
'     Dim cells() As Variant                                                                       '
'     srch = FindStr(Array("Val 261","Ral"), ThisComponent, Array("substr","prepared"))            '
'     While FindStr(srch).HasResults                                                               '
'         Redim Preserve cells(UBound(cells) + 1) As Variant                                       '
'         cells(UBound(cells)) = srch.Cell                                                         '
'     Wend                                                                                         '
'     duplicated_cells = DuplicateRowsByCells(cells,3)                                             '
'     For counter = 0 To UBound(duplicated_cells)                                                  '
'         cell = duplicated_cells(counter)                                                         '
'         cell.setString(cell.getString() + " " + counter)                                         '
'     Next counter                                                                                 '
'                                                                                                  '
' Output (the same):                                                                               '
'                                                                                                  '
'   +===+=========+===========+=========+=========+                                                '
'   |   |    A    |     B     |    C    |    D    |                                                '
'   +===+=========+===========+=========+=========+                                                '
'   | 1 | Title 5 | Title 6   | Title 7 | Title 8 |                                                '
'   +---+---------+-----------+---------+---------+                                                '
'   | 2 | Val 251 | Val 261 3 | Val 271 | Val 281 |                                                '
'   +---+---------+-----------+---------+---------+                                                '
'   | 3 | Val 251 | Val 261 4 | Val 271 | Val 281 |                                                '
'   +---+---------+-----------+---------+---------+                                                '
'   | 4 | Val 251 | Val 261 5 | Val 271 | Val 281 |                                                '
'   +---+---------+-----------+---------+---------+                                                '
'   | 5 | Val 252 | Ral 262 0 | Val 272 | Val 282 |                                                '
'   +---+---------+-----------+---------+---------+                                                '
'   | 6 | Val 252 | Ral 262 1 | Val 272 | Val 282 |                                                '
'   +---+---------+-----------+---------+---------+                                                '
'   | 7 | Val 252 | Ral 262 2 | Val 272 | Val 282 |                                                '
'   +---+---------+-----------+---------+---------+                                                '
'                                                                                                  '
'--------------------------------------------------------------------------------------------------'
' See also:                                                                                        '
'   ASCII tables generator.                                                                        '
'     https://ozh.github.io/ascii-tables/                                                          '
'--------------------------------------------------------------------------------------------------'
' Feedback & Issues:                                                                               '
'   https://github.com/aa6/libreoffice_calc_basic_extras/issues                                    '
'--------------------------------------------------------------------------------------------------'
Function DuplicateRowsByCells(RowCells as Variant, ResultRowsCount as Integer, Optional ActiveSheetForIntegerRowCells as Variant)
  Dim cells() As Variant
  Dim result_cells() As Variant 
  Dim prepared_cells() As Variant
  If IsMissing(ActiveSheetForIntegerRowCells) Then
    ActiveSheetForIntegerRowCells = ThisComponent.CurrentController.ActiveSheet
  End If
  If InStr(TypeName(RowCells),"(") > 0 Then
      cells = RowCells
  Else
      Redim Preserve cells(UBound(cells) + 1) As Variant
      cells(UBound(cells)) = RowCells
  End If
  ' Prepare rows and sort backwards by row index to prevent duplicating the same rows. '
  Redim prepared_cells(UBound(cells)) As Variant
  For index = 0 To UBound(cells)
    If TypeName(cells(index)) = "Integer" Then
      prepared_cells(index) = Array(cells(index), ActiveSheetForIntegerRowCells.getCellByPosition(0,cells(index)))
    Else
      prepared_cells(index) = Array(cells(index).CellAddress.Row, cells(index))
    End If
  Next index
    For i = 0 To Ubound(prepared_cells)
        For j = 0 To Ubound(prepared_cells)
            If  prepared_cells(i)(0) > prepared_cells(j)(0) Then
                swap =  prepared_cells(i)
                prepared_cells(i) = prepared_cells(j)
                prepared_cells(j) = swap
            End If
        Next j
    Next i
  For Each cell In prepared_cells
    cell = cell(1)
    Redim Preserve result_cells(UBound(result_cells) + 1) As Variant
    result_cells(UBound(result_cells)) = cell
    If ResultRowsCount > 1 Then
      cell.Spreadsheet.Rows.insertByIndex(cell.CellAddress.Row + 1, ResultRowsCount - 1)
      src_range = cell.spreadsheet.getRangeAddress()
      src_range.StartRow = cell.CellAddress.Row
      src_range.EndRow = cell.CellAddress.Row
      dst_cell = cell.getCellAddress()
      dst_cell.Column = 0
      For count = 2 To ResultRowsCount
        dst_cell.Row = dst_cell.Row + 1
        cell.Spreadsheet.copyRange(dst_cell,src_range)
        Redim Preserve result_cells(UBound(result_cells) + 1) As Variant
        result_cells(UBound(result_cells)) = cell.Spreadsheet.getCellByPosition(cell.CellAddress.Column,dst_cell.Row)
      Next count
    End If
  Next cell
  DuplicateRowsByCells = result_cells
End Function