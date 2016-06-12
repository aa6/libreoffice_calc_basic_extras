'--------------------------------------------------------------------------------------------------'
' ParseAllNumbers                                                                                  '
'--------------------------------------------------------------------------------------------------'
' Returns all numbers in string.                                                                   '
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
'     numbers = ParseFirstNumbers(str)                                                             '
'                                                                                                  '
' Expected values:                                                                                 '
'   str: "abcdef"     numbers: ""                                                                  '
'   str: "a2.5bcdef"  numbers: "25"                                                                '
'   str: "a1bc2d3ef"  numbers: "123"                                                               '
'   str: "a1b-c2d3ef" numbers: "123"                                                               '
'   str: "-a1bc2d3ef" numbers: "-123"                                                              '
'--------------------------------------------------------------------------------------------------'
'                                                                                                  '
'     numbers = ParseFirstNumbers(str, FALSE)                                                      '
'                                                                                                  '
' Expected values:                                                                                 '
'   str: "a1bc2d3ef"  numbers: "123"                                                               '
'   str: "a1b-c2d3ef" numbers: "123"                                                               '
'   str: "-a1bc2d3ef" numbers: "123"                                                               '
'--------------------------------------------------------------------------------------------------'
' Feedback & Issues:                                                                               '
'   https://github.com/aa6/libreoffice_calc_basic_extras/issues                                    '
'--------------------------------------------------------------------------------------------------'
Function ParseAllNumbers(Str As String, Optional CheckNegatives As Boolean) As String
   
    Dim pos As Long
    Dim poschar As String
    Dim numbers_are_negative As Boolean

    If IsMissing(CheckNegatives) Then
        CheckNegatives = TRUE
    End If
    ParseAllNumbers = ""
    numbers_are_negative = FALSE

    For pos = 1 to Len(Str)
        poschar = Mid(Str,pos,1)
        If CheckNegatives AND poschar = "-" AND ParseAllNumbers = "" Then
            numbers_are_negative = TRUE
        End If
        If Instr("0123456789", poschar) <> 0 Then
            ParseAllNumbers = ParseAllNumbers & poschar
        End If
    Next pos

    If numbers_are_negative Then
        ParseAllNumbers = "-" & ParseAllNumbers
    End If

End Function