'--------------------------------------------------------------------------------------------------'
' ParseFirstFloat                                                                                  '
'--------------------------------------------------------------------------------------------------'
' Returns first not separated by non-whitespace chars float nubmers.                               '
'                                                                                                  '
' Parameters:                                                                                      '
'                                                                                                  '
'   Str As String                                                                                  '
'     String to parse.                                                                             '
'                                                                                                  '
'   Optional CheckNegatives As Boolean <Default = TRUE>                                            '
'     Ignore the minus sign and parse only numbers.                                                '
'                                                                                                  '
' Examples:                                                                                        '
'--------------------------------------------------------------------------------------------------'
'                                                                                                  '
'     numbers = ParseFirstFloat(str)                                                               '
'                                                                                                  '
' Expected values:                                                                                 '
'   str: "Price is 2425 , 93 abcd 4223"     numbers: "2425,93"                                     '
'   str: "Price is 2425. abcd 93 efgh 4223" numbers: "2425"                                        '
'   str: ".2. abcd 93 efgh 4223"            numbers: "2"                                           '
'   str: "Price. 0.45"                      numbers: "0.45"                                        '
'   str: "Price. 0.45 3"                    numbers: "0.453"                                       '
'   str: "Price. 0.-45"                     numbers: "0"                                           '
'   str: "12.34.56.78"                      numbers: "12.34"                                       '
'   str: "44."                              numbers: "44"                                          '
'   str: "1 .45"                            numbers: "1.45"                                        '
'   str: ".45"                              numbers: "45"                                          '
'   str: "-.45"                             numbers: "-45"                                         '
'   str: ".-45"                             numbers: "-45"                                         '
'   str: ".4-5"                             numbers: "4"                                           '
'   str: "Price. -45"                       numbers: "-45"                                         '
'--------------------------------------------------------------------------------------------------'
'                                                                                                  '
'     numbers = ParseFirstFloat(str, FALSE)                                                        '
'                                                                                                  '
' Expected values:                                                                                 '
'   str: "Price. -45"                       numbers: "45"                                          '
'--------------------------------------------------------------------------------------------------'
Function ParseFirstFloat(Str As String, Optional CheckNegatives As Boolean) As String

    Dim pos As Long
    Dim poschar As String
    Dim decimal_separators As String
    Dim numbers_are_negative As Boolean
    
    If IsMissing(CheckNegatives) Then
        CheckNegatives = TRUE
    End If
    ParseFirstFloat = ""
    decimal_separators = ".,"
    numbers_are_negative = FALSE
    
    For pos = 1 To Len(Str)
        poschar = Mid(Str, pos, 1)
        If CheckNegatives AND poschar = "-" AND ParseFirstFloat = "" Then
            numbers_are_negative = TRUE
            GoTo ParseFirstFloatNextPos
        End If
        If Instr(decimal_separators, poschar) AND ParseFirstFloat <> "" Then
            ParseFirstFloat = ParseFirstFloat & poschar
            decimal_separators = ""
            GoTo ParseFirstFloatNextPos
        End If
        If Instr("0123456789", poschar) <> 0 Then
            ParseFirstFloat = ParseFirstFloat & poschar
        Else
            If NOT (Instr(" ", poschar) <> 0) AND NOT (Len(ParseFirstFloat) = 0) Then
                GoTo ParseFirstFloatExitPos
            End If
        End If
        ParseFirstFloatNextPos:
    Next pos
    ParseFirstFloatExitPos:

    If Len(ParseFirstFloat) = 0 Then
        Exit Function
    End If

    ' Trim trailing decimal_separators. "45." --> "45" '
    If Instr("0123456789", Mid(ParseFirstFloat, Len(ParseFirstFloat), 1)) = 0 Then
        ParseFirstFloat = Mid(ParseFirstFloat, 1, Len(ParseFirstFloat) - 1)
    End If

    If numbers_are_negative Then
        ParseFirstFloat = "-" & ParseFirstFloat
    End If

End Function