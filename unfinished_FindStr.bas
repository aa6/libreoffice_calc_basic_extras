'--------------------------------------------------------------------------------------------------'
' FindStr                                                                                          '
'--------------------------------------------------------------------------------------------------'
' Searches cells by string contents.                                                               '
'                                                                                                  '
' Parameters:                                                                                      '
'                                                                                                  '
'   SearchSubject As Variant                                                                       '
'     String or array of strings to search for.                                                    '
'                                                                                                  '
'   Optional SearchAreas As Variant <Default = ThisComponent.CurrentController.ActiveSheet>        '
'     Place or array of places to search.                                                          '
'                                                                                                  '
'     Possible types:                                                                              '
'       com.sun.star.sheet.Spreadsheet                                                             '
'       com.sun.star.sheet.SpreadsheetDocument                                                     '
'                                                                                                  '
'   Optional SearchOptions As Variant                                                              '
'     String with option name or an array of option name strings.                                  '
'                                                                                                  '
'     Default options (options that are on by default):                                            '
'       "case-insensitive", "non-substring", "notrim", "non-prepared"                              '
'                                                                                                  '
'     Possible options:                                                                            '
'       "trim", "trimmed"                                                                          '
'         Enable trimming of SearchSubject and cell values when search.                            '
'       "notrim", "notrimmed", "non-trim", "non-trimmed"                                           '
'         Disable trimming.                                                                        '
'       "substr", "substring", "substring-search"                                                  '
'         Enable matches of SearchSubject as a substring of cell values.                           '
'       "nosubstr", "non-substring", "non-substring-search"                                        '
'         Disable substring search.                                                                '
'       "case-insensitive", "caseinsensitive", "notcasesensitive"                                  '
'         Enable case insensitive seach.                                                           '
'       "case-sensitive", "casesensitive"                                                          '
'         Disable case insensitive search.                                                         '
'       "exact"                                                                                    '
'         Non-trimmed case-sensitive and non-substring search.                                     '
'       "prep", "prepare", "prepared"                                                              '
'         Prepared search. First call of FindStr will not search for values and would just return  '
'         search context for further search. It is useful for wrapping search code in a cycle.     '
'       "noprep", "noprepare", "non-prepared"                                                      '
'         Disables prepared search. First call of FindStr will search for values and would return  '
'         a search context including information about first search match.                         '
'       "reverse", "reversed", "reverse-search", "reversed-search"                                 '
'                                                                                                  '
' Examples:                                                                                        '
'--------------------------------------------------------------------------------------------------'
'                                                                                                  '
'     sheet = ThisComponent.Sheets.getByName("Sheet1")                                             '
'     srch = FindStr(Array("Title 1","Title 3","Title 7"), sheet , "prepared")                     '
'     While FindStr(srch).HasResults                                                               '
'       MsgBox "Row: " + srch.Row + " Col: " + srch.Col + " Value: " + srch.Cell.getString()       '
'     Wend                                                                                         '
'                                                                                                  '
' Expected results:                                                                                '
'                                                                                                  '
'   +===+=========+=========+=========+=========+   +===+=========+=========+=========+=========+  '
'   |   |    A    |    B    |    C    |    D    |   |   |    A    |    B    |    C    |    D    |  '
'   +===+=========+=========+=========+=========+   +===+=========+=========+=========+=========+  '
'   | 1 | Title 1 | Title 2 | Title 3 | Title 4 |   | 1 | Title 5 | Title 6 | Title 7 | Title 8 |  '
'   +---+---------+---------+---------+---------+   +---+---------+---------+---------+---------+  '
'   | 2 | Val 111 | Val 121 | Val 131 | Val 141 |   | 2 | Val 251 | Val 261 | Val 271 | Val 281 |  '
'   +---+---------+---------+---------+---------+   +---+---------+---------+---------+---------+  '
'   | 3 | Val 112 | Val 122 | Val 132 | Val 142 |   | 3 | Val 252 | Val 262 | Val 272 | Val 282 |  '
'   +---+---------+---------+---------+---------+   +---+---------+---------+---------+---------+  '
'   | Sheet1 |                                      | Sheet2 |                                     '
'   +========+                                      +========+                                     '
'                                                                                                  '
' Output:                                                                                          '
'                                                                                                  '
'   Row: 0 Col: 0 Value: Title 1                                                                   '
'   Row: 0 Col: 2 Value: Title 3                                                                   '
'                                                                                                  '
'--------------------------------------------------------------------------------------------------'
'                                                                                                  '
'     srch = FindStr(Array("Title 1","Title 3","Title 7"), , "prepared")                           '
'     While FindStr(srch).HasResults                                                               '
'       MsgBox "Row: " + srch.Row + " Col: " + srch.Col + " Value: " + srch.Cell.getString()       '
'     Wend                                                                                         '
'                                                                                                  '
' Expected results:                                                                                '
'                                                                                                  '
'   +===+=========+=========+=========+=========+   +===+=========+=========+=========+=========+  '
'   |   |    A    |    B    |    C    |    D    |   |   |    A    |    B    |    C    |    D    |  '
'   +===+=========+=========+=========+=========+   +===+=========+=========+=========+=========+  '
'   | 1 | Title 1 | Title 2 | Title 3 | Title 4 |   | 1 | Title 5 | Title 6 | Title 7 | Title 8 |  '
'   +---+---------+---------+---------+---------+   +---+---------+---------+---------+---------+  '
'   | 2 | Val 111 | Val 121 | Val 131 | Val 141 |   | 2 | Val 251 | Val 261 | Val 271 | Val 281 |  '
'   +---+---------+---------+---------+---------+   +---+---------+---------+---------+---------+  '
'   | 3 | Val 112 | Val 122 | Val 132 | Val 142 |   | 3 | Val 252 | Val 262 | Val 272 | Val 282 |  '
'   +---+---------+---------+---------+---------+   +---+---------+---------+---------+---------+  '
'   | Sheet1 |                                      | Sheet2 |                                     '
'   +========+                                      +========+                                     '
'                                                                                                  '
' Output (Sheet1 is active):                                                                       '
'                                                                                                  '
'   Row: 0 Col: 0 Value: Title 1                                                                   '
'   Row: 0 Col: 2 Value: Title 3                                                                   '
'                                                                                                  '
' Output (Sheet2 is active):                                                                       '
'                                                                                                  '
'   Row: 0 Col: 2 Value: Title 7                                                                   '
'                                                                                                  '
'--------------------------------------------------------------------------------------------------'
'                                                                                                  '
'     sheet1 = ThisComponent.Sheets.getByName("Sheet1")                                            '
'     sheet2 = ThisComponent.Sheets.getByName("Sheet2")                                            '
'     srch = FindStr(Array("Title 1","Title 3","Title 7"), Array(sheet1,sheet2) , "prepared")      '
'     While FindStr(srch).HasResults                                                               '
'       MsgBox "Row: " + srch.Row + " Col: " + srch.Col + " Value: " + srch.Cell.getString()       '
'     Wend                                                                                         '
'                                                                                                  '
' Expected results:                                                                                '
'                                                                                                  '
'   +===+=========+=========+=========+=========+   +===+=========+=========+=========+=========+  '
'   |   |    A    |    B    |    C    |    D    |   |   |    A    |    B    |    C    |    D    |  '
'   +===+=========+=========+=========+=========+   +===+=========+=========+=========+=========+  '
'   | 1 | Title 1 | Title 2 | Title 3 | Title 4 |   | 1 | Title 5 | Title 6 | Title 7 | Title 8 |  '
'   +---+---------+---------+---------+---------+   +---+---------+---------+---------+---------+  '
'   | 2 | Val 111 | Val 121 | Val 131 | Val 141 |   | 2 | Val 251 | Val 261 | Val 271 | Val 281 |  '
'   +---+---------+---------+---------+---------+   +---+---------+---------+---------+---------+  '
'   | 3 | Val 112 | Val 122 | Val 132 | Val 142 |   | 3 | Val 252 | Val 262 | Val 272 | Val 282 |  '
'   +---+---------+---------+---------+---------+   +---+---------+---------+---------+---------+  '
'   | Sheet1 |                                      | Sheet2 |                                     '
'   +========+                                      +========+                                     '
'                                                                                                  '
' Output:                                                                                          '
'                                                                                                  '
'   Row: 0 Col: 0 Value: Title 1                                                                   '
'   Row: 0 Col: 2 Value: Title 3                                                                   '
'   Row: 0 Col: 2 Value: Title 7                                                                   '
'                                                                                                  '
'--------------------------------------------------------------------------------------------------'
'                                                                                                  '
'     search_words = Array("Title 1","Title 3","Title 7")                                          '
'     sheet1 = ThisComponent.Sheets.getByName("Sheet1")                                            '
'     sheet2 = ThisComponent.Sheets.getByName("Sheet2")                                            '
'     srch = FindStr(search_words, Array(sheet1,sheet2) , Array("prepared","reverse"))             '
'     While FindStr(srch).HasResults                                                               '
'       MsgBox "Row: " + srch.Row + " Col: " + srch.Col + " Value: " + srch.Cell.getString()       '
'     Wend                                                                                         '
'                                                                                                  '
' Expected results:                                                                                '
'                                                                                                  '
'   +===+=========+=========+=========+=========+   +===+=========+=========+=========+=========+  '
'   |   |    A    |    B    |    C    |    D    |   |   |    A    |    B    |    C    |    D    |  '
'   +===+=========+=========+=========+=========+   +===+=========+=========+=========+=========+  '
'   | 1 | Title 1 | Title 2 | Title 3 | Title 4 |   | 1 | Title 5 | Title 6 | Title 7 | Title 8 |  '
'   +---+---------+---------+---------+---------+   +---+---------+---------+---------+---------+  '
'   | 2 | Val 111 | Val 121 | Val 131 | Val 141 |   | 2 | Val 251 | Val 261 | Val 271 | Val 281 |  '
'   +---+---------+---------+---------+---------+   +---+---------+---------+---------+---------+  '
'   | 3 | Val 112 | Val 122 | Val 132 | Val 142 |   | 3 | Val 252 | Val 262 | Val 272 | Val 282 |  '
'   +---+---------+---------+---------+---------+   +---+---------+---------+---------+---------+  '
'   | Sheet1 |                                      | Sheet2 |                                     '
'   +========+                                      +========+                                     '
'                                                                                                  '
' Output:                                                                                          '
'                                                                                                  '
'   Row: 0 Col: 2 Value: Title 7                                                                   '
'   Row: 0 Col: 2 Value: Title 3                                                                   '
'   Row: 0 Col: 0 Value: Title 1                                                                   '
'                                                                                                  '
'--------------------------------------------------------------------------------------------------'
'                                                                                                  '
'     srch = FindStr("Title 7", ThisComponent)                                                     '
'     If srch.Found Then                                                                           '
'       MsgBox "Row: " + srch.Row + " Col: " + srch.Col + " Value: " + srch.Cell.getString()       '
'     EndIf                                                                                        '
'                                                                                                  '
' Expected results:                                                                                '
'                                                                                                  '
'   +===+=========+=========+=========+=========+   +===+=========+=========+=========+=========+  '
'   |   |    A    |    B    |    C    |    D    |   |   |    A    |    B    |    C    |    D    |  '
'   +===+=========+=========+=========+=========+   +===+=========+=========+=========+=========+  '
'   | 1 | Title 1 | Title 2 | Title 3 | Title 4 |   | 1 | Title 5 | Title 6 | Title 7 | Title 8 |  '
'   +---+---------+---------+---------+---------+   +---+---------+---------+---------+---------+  '
'   | 2 | Val 111 | Val 121 | Val 131 | Val 141 |   | 2 | Val 251 | Val 261 | Val 271 | Val 281 |  '
'   +---+---------+---------+---------+---------+   +---+---------+---------+---------+---------+  '
'   | 3 | Val 112 | Val 122 | Val 132 | Val 142 |   | 3 | Val 252 | Val 262 | Val 272 | Val 282 |  '
'   +---+---------+---------+---------+---------+   +---+---------+---------+---------+---------+  '
'   | Sheet1 |                                      | Sheet2 |                                     '
'   +========+                                      +========+                                     '
'                                                                                                  '
' Output (Sheet1 is active):                                                                       '
'                                                                                                  '
'   Row: 0 Col: 2 Value: Title 7                                                                   '
'                                                                                                  '
'--------------------------------------------------------------------------------------------------'
' See also:                                                                                        '
'   ASCII tables generator.                                                                        '
'     https://ozh.github.io/ascii-tables/                                                          '
'--------------------------------------------------------------------------------------------------'
' Feedback & Issues:                                                                               '
'   https://github.com/aa6/libreoffice_calc_basic_extras/issues                                    '
'--------------------------------------------------------------------------------------------------'
' Options: Array("Case-Insensitive","Regex","etc.")
' article_col = FindStr("test", Array(searcharea,"case-insensitive","trimmed"))
Type SearchContextOptionsOfFindStrFunction

    CaseSensitive As Boolean
    TrimmedSearch As Boolean
    PreparedSearch As Boolean
    ReversedSearch As Boolean
    SubstringSearch As Boolean
  
End Type
Type SearchContextOfFindStrFunction

    Row As Long
    Col As Long
    Cell As com.sun.star.table.XCell
    Sheet As com.sun.star.sheet.XSpreadsheet
    Column As Long
    Empty As Boolean
    Found As Boolean
    HasResults As Boolean
    HasNoResults As Boolean
    NotEmpty As Boolean
    NotFound As Boolean
    IsEmpty As Boolean
    IsFound As Boolean
    IsNotEmpty As Boolean
    IsNotFound As Boolean
    SheetName As String
    SheetIndex As Long
    ' For internal usage: '
    ContextRow As Long
    ContextCol As Long
    ContextArea As Long
    ContextAreas() As Variant
    ContextWords() As Variant
    ContextOptions As SearchContextOptionsOfFindStrFunction
  
End Type
Function FindStr_Object_Is_RangeAddress(TestObject as Variant)
    On Local Error Goto Label_FindStr_Object_Is_Not_RangeAddress
    test = TestObject.Sheet
    test = TestObject.StartRow
    test = TestObject.StartColumn
    test = TestObject.EndRow
    test = TestObject.EndColumn
    FindStr_Object_Is_RangeAddress = TRUE
    Exit Function
    Label_FindStr_Object_Is_Not_RangeAddress:
    FindStr_Object_Is_RangeAddress = FALSE
End Function
Function FindStr(SearchSubject as Variant, Optional SearchAreas As Variant, Optional SearchOptions As Variant) As Object
    
    Dim col As Long
    Dim row As Long
    Dim item As Variant
    Dim word As Variant
    Dim sheet As Variant
    Dim cellvalue As String
    Dim usedrange As Object
    Dim search_word As Variant
    Dim search_area As Variant
    Dim search_result As Boolean
    Dim search_context As SearchContextOfFindStrFunction 
    Dim search_areas() As Variant ' Normalized search_area is an Array(Sheet,RangeAddress). '
    Dim search_words() As Variant
    Dim search_options() As Variant
    Dim search_area_index As Variant
    Dim normalized_search_areas() As Variant
    Dim normalized_search_words() As Variant
    
    ' Processing SearchItems. '
    If TypeName(SearchSubject) <> "Object" Then
        ' Initializing new search context. '
        search_context = New SearchContextOfFindStrFunction
        search_context.ContextOptions = New SearchContextOptionsOfFindStrFunction
        ' Preparing SearchOptions. '
        search_context.ContextOptions.CaseSensitive = FALSE
        search_context.ContextOptions.TrimmedSearch = FALSE
        search_context.ContextOptions.ReversedSearch = FALSE
        search_context.ContextOptions.PreparedSearch = FALSE
        search_context.ContextOptions.SubstringSearch = FALSE
        If NOT IsMissing(SearchOptions) Then
            If InStr(TypeName(SearchOptions),"(") > 0 Then
                search_options = SearchOptions
            Else
                Redim Preserve search_options(UBound(search_options) + 1) As Variant
                search_options(UBound(search_options)) = SearchOptions
            End If
        End If
        For Each item In search_options
            If TypeName(item) = "String" Then
                Select Case LCase(Trim(item))
                    Case "exact"
                        search_context.ContextOptions.TrimmedSearch = FALSE
                        search_context.ContextOptions.SubstringSearch = FALSE
                    Case "trim", "trimmed"
                        search_context.ContextOptions.TrimmedSearch = TRUE
                    Case "prep", "prepare", "prepared"
                        search_context.ContextOptions.PreparedSearch = TRUE
                    Case "notrim", "notrimmed", "nontrim", "nontrimmed"
                        search_context.ContextOptions.TrimmedSearch = FALSE
                    Case "substr", "substring", "substring-search"
                        search_context.ContextOptions.SubstringSearch = TRUE
                    Case "nosubstr"
                        search_context.ContextOptions.SubstringSearch = FALSE
                    Case "case-sensitive", "casesensitive"
                        search_context.ContextOptions.CaseSensitive = TRUE
                    Case "case-insensitive", "caseinsensitive", "notcasesensitive"
                        search_context.ContextOptions.CaseSensitive = FALSE
                    Case "reverse", "reversed", "reverse-search", "reversed-search"
                        search_context.ContextOptions.ReversedSearch = TRUE
                    Case Else
                        Err.Raise("UNSUPPORTED FINDSTR OPTION: " & LCase(Trim(item)))
                End Select
            Else
                Err.Raise("FINDSTR OPTION MUST BE A STRING")
            End If
        Next item
        ' Preparing search_words. '
        If InStr(TypeName(SearchSubject),"(") > 0 Then
            search_words = SearchSubject
        Else
            Redim Preserve search_words(UBound(search_words) + 1) As Variant
            search_words(UBound(search_words)) = CStr(SearchSubject)
        End If
        ' Normalizing raw search_words. '
        For Each search_word In search_words
            search_word = IIf(search_context.ContextOptions.TrimmedSearch, Trim(search_word), search_word)
            search_word = IIf(search_context.ContextOptions.CaseSensitive, search_word, LCase(search_word))
            Redim Preserve normalized_search_words(UBound(normalized_search_words) + 1) As Variant
            normalized_search_words(UBound(normalized_search_words)) = search_word
        Next search_word
        search_context.ContextWords = normalized_search_words ' Object property redim workaround. '
        ' Preparing SearchAreas. '
        If NOT IsMissing(SearchAreas) Then
            If InStr(TypeName(SearchAreas),"(") > 0 Then
                search_areas = SearchAreas
            Else
                Redim Preserve search_areas(UBound(search_areas) + 1) As Variant
                search_areas(UBound(search_areas)) = SearchAreas
            End If
        Else
            Redim Preserve search_areas(UBound(search_areas) + 1) As Variant
            search_areas(UBound(search_areas)) = ThisComponent.CurrentController.ActiveSheet
        End If
        ' Normalizing raw search_areas. '
        For Each search_area In search_areas 
            Select Case TRUE
                Case InStr(TypeName(search_area),"(") > 0
                    If NOT search_area(0).SupportsService("com.sun.star.sheet.Spreadsheet") Then
                        Err.Raise("search_area(0) must be a com.sun.star.sheet.Spreadsheet")
                    End If
                    If NOT FindStr_Object_Is_RangeAddress(search_area(1)) Then
                        Err.Raise("search_area(1) must be a RangeAddress")
                    End If
                    Redim Preserve normalized_search_areas(UBound(normalized_search_areas) + 1) As Variant
                    normalized_search_areas(UBound(normalized_search_areas)) = search_area
                Case FindStr_Object_Is_RangeAddress(search_area)
                    'Xray ThisComponent
                    sheet = ThisComponent.Sheets.getByIndex(search_area.Sheet)
                    Redim Preserve normalized_search_areas(UBound(normalized_search_areas) + 1) As Variant
                    normalized_search_areas(UBound(normalized_search_areas)) = Array(sheet,search_area)
                Case search_area.SupportsService("com.sun.star.sheet.Spreadsheet")
                    item = search_area.CreateCursor()
                    item.gotoStartOfUsedArea(FALSE) ' FALSE sets cursor size to a 1x1 cell. '
                    item.gotoEndOfUsedArea(TRUE)    ' TRUE expands cursor range. '
                    Redim Preserve normalized_search_areas(UBound(normalized_search_areas) + 1) As Variant
                    normalized_search_areas(UBound(normalized_search_areas)) = Array(search_area,item.getRangeAddress())
                Case search_area.SupportsService("com.sun.star.sheet.SpreadsheetDocument")
                    For Each item In search_area.Sheets.ElementNames
                        sheet = search_area.Sheets.getByName(item)
                        item = sheet.CreateCursor()
                        item.gotoStartOfUsedArea(FALSE) ' FALSE sets cursor size to a 1x1 cell. '
                        item.gotoEndOfUsedArea(TRUE)    ' TRUE expands cursor range. '
                        Redim Preserve normalized_search_areas(UBound(normalized_search_areas) + 1) As Variant
                        normalized_search_areas(UBound(normalized_search_areas)) = Array(sheet,item.getRangeAddress())            
                    Next item
                Case Else
                    Err.Raise("UNSUPPORTED TYPE OF FINDSTR SEARCH RANGE") ' Sorry. '
            End Select
        Next search_area
        search_context.ContextAreas = normalized_search_areas
        ' Finishing context initialization. '
        If search_context.ContextOptions.ReversedSearch = TRUE Then
            search_context.ContextArea = UBound(search_context.ContextAreas)
            search_context.ContextRow = -2
            search_context.ContextCol = -2
        Else
            search_context.ContextRow = 0
            search_context.ContextCol = 0
            search_context.ContextArea = 0
        End If
        If search_context.ContextOptions.PreparedSearch = TRUE Then
            FindStr = search_context
            Exit Function
        End If
    Else
        ' Implying that each object passed as SearchSubject would be a search context. '
        search_context = SearchSubject
    End If

    ' Searching. '    
    FindStr = search_context
    search_context.Empty = TRUE
    search_context.Found = FALSE
    search_context.HasResults = FALSE
    search_context.HasNoResults = TRUE
    search_context.IsEmpty = TRUE
    search_context.IsFound = FALSE
    search_context.NotEmpty = FALSE
    search_context.NotFound = TRUE
    search_context.IsNotEmpty = FALSE
    search_context.IsNotFound = TRUE

    If search_context.ContextOptions.ReversedSearch = TRUE Then
        search_area_step = -1
        search_area_index_target = 0
    Else
        search_area_step = 1
        search_area_index_target = UBound(search_context.ContextAreas)
    End If

    For search_area_index = search_context.ContextArea To search_area_index_target Step search_area_step
        sheet = search_context.ContextAreas(search_area_index)(0)
        usedrange = search_context.ContextAreas(search_area_index)(1)
        If search_context.ContextOptions.ReversedSearch = TRUE Then
            search_area_row_target = usedrange.StartRow
            search_area_col_target = usedrange.StartColumn
        Else
            search_area_row_target = usedrange.EndRow
            search_area_col_target = usedrange.EndColumn
        End If 
        If search_context.ContextRow = -2 Then
            search_context.ContextRow = usedrange.EndRow - usedrange.StartRow
        End if
        For row = usedrange.StartRow + search_context.ContextRow To search_area_row_target Step search_area_step
            If search_context.ContextCol = -2 Then
                search_context.ContextCol = usedrange.EndColumn - usedrange.StartColumn
            End if
            For col = usedrange.StartColumn + search_context.ContextCol To search_area_col_target Step search_area_step
                cellvalue = sheet.getCellByPosition(col,row).getString()
                cellvalue = IIf(search_context.ContextOptions.TrimmedSearch, Trim(cellvalue), cellvalue)
                cellvalue = IIf(search_context.ContextOptions.CaseSensitive, cellvalue, LCase(cellvalue))
                For Each word In search_context.ContextWords
                    If search_context.ContextOptions.SubstringSearch Then
                        search_result = IIf(InStr(1,cellvalue,word,0) > 0, TRUE, FALSE)
                    Else
                        search_result = IIf(cellvalue = word, TRUE, FALSE)
                    End If
                    If search_result = TRUE Then
                        search_context.Empty = FALSE
                        search_context.Found = TRUE
                        search_context.HasResults = TRUE
                        search_context.HasNoResults = FALSE
                        search_context.NotEmpty = TRUE
                        search_context.NotFound = FALSE
                        search_context.IsEmpty = FALSE
                        search_context.IsFound = TRUE
                        search_context.IsNotEmpty = TRUE
                        search_context.IsNotFound = FALSE
                        search_context.Sheet = sheet
                        search_context.Cell = sheet.getCellByPosition(col,row)
                        search_context.Row = search_context.Cell.CellAddress.Row
                        search_context.Col = search_context.Cell.CellAddress.Column
                        search_context.Column = search_context.Cell.CellAddress.Column
                        search_context.SheetName = sheet.Name
                        search_context.SheetIndex = sheet.RangeAddress.Sheet
                        search_context.ContextRow = row - usedrange.StartRow
                        search_context.ContextCol = col - usedrange.StartColumn + search_area_step
                        search_context.ContextArea = search_area_index
                        Exit Function
                    End If
                Next word
            Next col
            search_context.ContextCol = IIF(search_context.ContextOptions.ReversedSearch, -2, 0)
        Next row
        search_context.ContextRow = IIF(search_context.ContextOptions.ReversedSearch, -2, 0)
    Next search_area_index
    ' Search finished
    Erase search_context.Sheet
    Erase search_context.Cell
    Erase search_context.Row
    Erase search_context.Col
    Erase search_context.Column
    Erase search_context.SheetName
    Erase search_context.SheetIndex
    search_context.ContextArea = search_area_index

End Function