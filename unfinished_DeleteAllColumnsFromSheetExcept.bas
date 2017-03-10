Function DeleteAllColumnsFromSheetExcept(headers As Variant, Optional sheet as Object, Optional header_rows_length As Integer)

  ' Setting default values.                                                 '
  If IsMissing(header_rows_length) Then
    header_rows_length = 48
  End If
  If IsMissing(sheet) Then
    'Xray ThisComponent.getCurrentSelection
    sheet = ThisComponent.getCurrentSelection.Spreadsheet
  End If
  
  ' Getting used area.                                                      '
  Dim cursor As Object
  Dim sheet_used_range As Object
    
  cursor = sheet.CreateCursor()
  cursor.gotoStartOfUsedArea(FALSE) ' FALSE sets cursor size to a 1x1 cell. '
  cursor.gotoEndOfUsedArea(TRUE)    ' TRUE expands cursor range.            '
  sheet_used_range = cursor.getRangeAddress()
  
  ' Cycling headers that need to stay.                                      '
    Dim col As Long
    Dim row As Long
  Dim col_indexes() As Variant
  For Each header In headers
    ' If already a column index is given.                                 '
    If TypeName(header) = "Integer" OR TypeName(header) = "Long" Then
      Redim Preserve col_indexes(UBound(col_indexes) + 1) As Variant
      col_indexes(UBound(col_indexes)) = header
    End If
    ' We need to find a column index by a cell text.                      '
    If TypeName(header) = "String" Then
        For col = sheet_used_range.StartColumn To sheet_used_range.EndColumn Step 1
          For row = 0 To header_rows_length Step 1
                If LCase(Trim(sheet.getCellByPosition(col,row).getString())) = LCase(header) Then
            Redim Preserve col_indexes(UBound(col_indexes) + 1) As Variant
            col_indexes(UBound(col_indexes)) = col
                End If
            Next row
        Next col
    End If
  Next header
  
  ' Ensure that column indexes are unique                                    '
  Dim unique_col_indexes() As Variant
  For Each col In col_indexes
    For Each uniq_col In unique_col_indexes
      If uniq_col = col Then
        Goto NextNonUniqueColumn
      End If
    Next uniq_col
    Redim Preserve unique_col_indexes(UBound(unique_col_indexes) + 1) As Variant
    unique_col_indexes(UBound(unique_col_indexes)) = col
    NextNonUniqueColumn:
  Next col
  
  ' Sort columns descending                                                   '
  Dim rsorted_unique_col_indexes(UBound(unique_col_indexes)) As Variant
  Dim temp As Variant
  For i = 0 To Ubound(unique_col_indexes)
    rsorted_unique_col_indexes(i) = unique_col_indexes(i)
  Next i
  For i = 0 To Ubound(rsorted_unique_col_indexes)
    For j = 0 To Ubound(rsorted_unique_col_indexes)
      'If CInt(rsorted_unique_col_indexes(i)) > CInt(rsorted_unique_col_indexes(j)) Then
      If CLng(rsorted_unique_col_indexes(i)) > CLng(rsorted_unique_col_indexes(j)) Then
        temp = rsorted_unique_col_indexes(i)
        rsorted_unique_col_indexes(i) = rsorted_unique_col_indexes(j)
        rsorted_unique_col_indexes(j) = temp
      End If
    Next j
  Next i
  
  ' Delete the columns
  For col = sheet_used_range.EndColumn To sheet_used_range.StartColumn Step -1
    For Each index In rsorted_unique_col_indexes
      If col = index Then
        Goto NextColumn
      End If
    Next index
    sheet.Columns.removeByIndex(col,1)
    NextColumn:
  Next col
  
End Function