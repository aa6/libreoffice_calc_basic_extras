' Returns first not separated by non-whitespace chars float nubmers.                               '
' Examples:                                                                                        '
'   ParseFirstNumbers("Price is 2425 , 93 snsn 4223") REM Returns "2425,93"                        '
'   ParseFirstNumbers("Price is 2425. znzn 93 snsn 4223") REM Returns "2425"                       '
'   ParseFirstNumbers("Price is .2. znzn 93 snsn 4223") REM Returns "2"                            '
'   ParseFirstNumbers("Welcome. 0.45") REM Returns "0.45"                                          '
'   ParseFirstNumbers("Welcome. 0.45 3") REM Returns "0.453"                                       '
'   ParseFirstNumbers("Welcome. 0.-45") REM Returns "0"                                            '
'   ParseFirstNumbers("Welcome. -45") REM Returns "-45"                                            '
'   ParseFirstNumbers("Welcome. -45", FALSE) REM Returns "45"                                      '
'   ParseFirstNumbers("12.34.56.78") REM Returns "12.34"                                           '
'   ParseFirstNumbers("44.") REM Returns "44"                                                      '
'   ParseFirstNumbers("1 .45") REM Returns "1.45"                                                  '
'   ParseFirstNumbers(".45") REM Returns "45"                                                      '
'   ParseFirstNumbers("-.45") REM Returns "-45"                                                    '
'   ParseFirstNumbers(".-45") REM Returns "-45"                                                    '
'   ParseFirstNumbers(".4-5") REM Returns "4"                                                      '
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