' Returns all numbers in string.                                                                   '
' Examples:                                                                                        '
'   ParseFirstNumbers("price is 2425 , 93 snsn 4223") REM Returns "2425934223"                     '
'   ParseFirstNumbers("price is 2-425 , 93 snsn 4223") REM Returns "2425934223"                    '
'   ParseFirstNumbers("-price is 2425 , 93 snsn 4223") REM Returns "-2425934223"                   '
'   ParseFirstNumbers("-price is 2425 , 93 snsn 4223", FALSE) REM Returns "2425934223"             '
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