Function MergeColumnsByTitles(ColumnTitles As Variant, Optional Glue As String, Optional SearchOptions As Variant, Optional TargetSheet As Variant)
    Dim titles() As Variant
    Dim column_indexes() As Integer
    Dim search_options() As Variant
    Dim sheet_used_range As Variant
    Dim column_indexes_unique() As Integer
    Glue = IIF(IsMissing(Glue), " ", Glue)
    TargetSheet = IIF(IsMissing(TargetSheet), ThisComponent.CurrentController.ActiveSheet, TargetSheet)
    SearchTrimmed = FALSE
    SearchSubstring = FALSE
    SearchCaseSensitive = FALSE
    If InStr(TypeName(ColumnTitles),"(") > 0 Then
        titles = ColumnTitles
    Else
        Redim Preserve titles(UBound(titles) + 1) As Variant
        titles(UBound(titles)) = ColumnTitles
    End If
    If NOT IsMissing(SearchOptions) Then
      If InStr(TypeName(SearchOptions),"(") > 0 Then
          search_options = SearchOptions
      Else
          Redim Preserve search_options(UBound(search_options) + 1) As Variant
          search_options(UBound(search_options)) = SearchOptions
      End If
      For Each item In search_options
          If TypeName(item) = "String" Then
              Select Case LCase(Trim(item))
                  Case "exact"
                      SearchTrimmed = FALSE
                      SearchSubstring = FALSE
                  Case "trim", "trimmed"
                      SearchTrimmed = TRUE
                  Case "notrim", "notrimmed", "nontrim", "nontrimmed"
                      SearchTrimmed = FALSE
                  Case "substr", "substring", "substring-search"
                      SearchSubstring = TRUE
                  Case "nosubstr"
                      SearchSubstring = FALSE
                  Case "case-sensitive", "casesensitive"
                      SearchCaseSensitive = TRUE
                  Case "case-insensitive", "caseinsensitive", "notcasesensitive"
                      SearchCaseSensitive = FALSE
                  Case Else
                      Err.Raise("UNSUPPORTED FINDSTR OPTION: " & LCase(Trim(item)))
              End Select
          Else
              Err.Raise("FINDSTR OPTION MUST BE A STRING")
          End If
      Next item
    End If
    ' Determine sheet used range to search in.
    cursor = TargetSheet.CreateCursor()
    cursor.gotoStartOfUsedArea(FALSE) ' FALSE sets cursor size to a 1x1 cell. '
    cursor.gotoEndOfUsedArea(TRUE)    ' TRUE expands cursor range.            '
    sheet_used_range = cursor.getRangeAddress()
    ' Transform titles into column_indexes.
    For Each title In titles
        If TypeName(title) = "Number" OR TypeName(title) = "Long" Then
            Redim Preserve column_indexes(UBound(column_indexes) + 1) As Integer
            column_indexes(UBound(column_indexes)) = title
        Else
            title = IIF(SearchTrimmed, Trim(title), title)
            title = IIF(SearchCaseSensitive, title, LCase(title))
            For row = sheet_used_range.StartRow To sheet_used_range.EndRow
                For col = sheet_used_range.StartColumn To sheet_used_range.EndColumn
                    cellvalue = TargetSheet.getCellByPosition(col,row).getString()
                    cellvalue = IIF(SearchTrimmed, Trim(cellvalue), cellvalue)
                    cellvalue = IIF(SearchCaseSensitive, cellvalue, LCase(cellvalue))
                    If SearchSubstring Then
                        If InStr(cellvalue, title) Then
                          Redim Preserve column_indexes(UBound(column_indexes) + 1) As Integer
                          column_indexes(UBound(column_indexes)) = col
                          Goto NextTitle
                        End If
                    Else
                        If cellvalue = title Then
                          Redim Preserve column_indexes(UBound(column_indexes) + 1) As Integer
                          column_indexes(UBound(column_indexes)) = col
                          Goto NextTitle
                        End If
                    End If
                Next col
            Next row
        End If
        NextTitle:
    Next title
    ' Remove duplicates but preserve column order.
    For Each item In column_indexes
        For Each entry In column_indexes_unique
            If item = entry Then
                Goto ArrayUniqueNextItem
            End If
        Next entry
        Redim Preserve column_indexes_unique(UBound(column_indexes_unique) + 1) As Integer
        column_indexes_unique(UBound(column_indexes_unique)) = item
        ArrayUniqueNextItem:
    Next item
    ' Exit if one column or less.
    If UBound(column_indexes_unique) < 1 Then
        Exit Function
    End if
    ' Merge columns contents.
    firstcol = column_indexes_unique(LBound(column_indexes_unique))
    ReDim Preserve column_indexes_unique(LBound(column_indexes_unique) + 1 To UBound(column_indexes_unique))
    For row = sheet_used_range.StartRow To sheet_used_range.EndRow
        For Each col in column_indexes_unique
            TargetSheet.getCellByPosition(firstcol,row).setString(          _
                TargetSheet.getCellByPosition(firstcol,row).getString() + Glue +  _
                TargetSheet.getCellByPosition(col,row).getString()          _
            )
        Next col
    Next row
    ' Sort columns in descending.
    For i = 1 To Ubound(column_indexes_unique)
        For j = 1 To Ubound(column_indexes_unique)
            If column_indexes_unique(i) > column_indexes_unique(j) Then
                swap = column_indexes_unique(i)
                column_indexes_unique(i) = column_indexes_unique(j)
                column_indexes_unique(j) = swap
            End If
        Next j
    Next i
    ' Delete rsorted columns.
    For Each col In column_indexes_unique
        TargetSheet.Columns.removeByIndex(col,1)
    Next col
End Function