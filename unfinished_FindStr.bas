Option Explicit
' Options: Array("Case-Insensitive","Regex","etc.")
' article_col = FindStr("test", Array(searcharea,"case-insensitive","trimmed"))
Type SearchContextOptionsOfFindStrFunction

    CaseSensitive As Boolean
    TrimmedSearch As Boolean
    PreparedSearch As Boolean
    SubstringSearch As Boolean
  
End Type
Type SearchContextOfFindStrFunction

    Row As Long
    Col As Long
    Cell As com.sun.star.table.XCell
    Sheet As com.sun.star.sheet.XSpreadsheet
    Column As Long
    Empty As Boolean
    NotEmpty As Boolean
    IsEmpty As Boolean
    IsNotEmpty As Boolean
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
        search_context.ContextRow = 0
        search_context.ContextCol = 0
        search_context.ContextArea = 0
        search_context.ContextOptions = New SearchContextOptionsOfFindStrFunction
        ' Preparing SearchOptions. '
        search_context.ContextOptions.CaseSensitive = FALSE
        search_context.ContextOptions.TrimmedSearch = FALSE
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
    search_context.IsEmpty = TRUE
    search_context.NotEmpty = FALSE
    search_context.IsNotEmpty = FALSE
    For search_area_index = search_context.ContextArea To UBound(search_context.ContextAreas) Step 1
        sheet = search_context.ContextAreas(search_area_index)(0)
        usedrange = search_context.ContextAreas(search_area_index)(1)
        For row = usedrange.StartRow + search_context.ContextRow To usedrange.EndRow
            For col = usedrange.StartColumn + search_context.ContextCol To usedrange.EndColumn
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
                        search_context.NotEmpty = TRUE
                        search_context.IsEmpty = FALSE
                        search_context.IsNotEmpty = TRUE
                        search_context.Sheet = sheet
                        search_context.Cell = sheet.getCellByPosition(col,row)
                        search_context.Row = search_context.Cell.CellAddress.Row
                        search_context.Col = search_context.Cell.CellAddress.Column
                        search_context.Column = search_context.Cell.CellAddress.Column
                        search_context.SheetName = sheet.Name
                        search_context.SheetIndex = sheet.RangeAddress.Sheet
                        search_context.ContextRow = row - usedrange.StartRow
                        search_context.ContextCol = col - usedrange.StartColumn + 1
                        search_context.ContextArea = search_area_index
                        Exit Function
                    End If
                Next word
            Next col
            search_context.ContextCol = 0
        Next row
        search_context.ContextRow = 0
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