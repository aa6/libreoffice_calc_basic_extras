Function ReplaceMultipleSpaces(Str As String)
    white_flag = FALSE
    ReplaceMultipleSpaces = ""
    For pos = 1 To Len(Str)
        char = Mid(Str,pos,1)
        If char = " " Then
            If white_flag = FALSE Then
                ReplaceMultipleSpaces = ReplaceMultipleSpaces & char
                white_flag = TRUE
            End If
        Else
            ReplaceMultipleSpaces = ReplaceMultipleSpaces & char
            white_flag = FALSE
        End If 
    Next pos
End Function