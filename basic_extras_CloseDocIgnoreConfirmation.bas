'--------------------------------------------------------------------------------------------------'
' CloseDocIgnoreConfirmation                                                                       '
'--------------------------------------------------------------------------------------------------'
' Closes the document ignoring possible confirmation dialogues.                                    '
' The main purpose of this function is to make code more readable because `Document.Close(TRUE)`   '
' don't really make any sense if you don't know what does it mean exactly.                         '
'                                                                                                  '
' Parameters:                                                                                      '
'                                                                                                  '
'   Optional Document as Variant <Default = ThisComponent>                                         '
'     Reference to a document (com.sun.star.uno.XInterface).                                       '
'                                                                                                  '
' Examples:                                                                                        '
'--------------------------------------------------------------------------------------------------'
'                                                                                                  '
'     CloseDocIgnoreConfirmation()                                                                 '
' or                                                                                               '
'     CloseDocIgnoreConfirmation(docref)                                                           '
'                                                                                                  '
'--------------------------------------------------------------------------------------------------'
' Feedback & Issues:                                                                               '
'   https://github.com/aa6/libreoffice_calc_basic_extras/issues                                    '
'--------------------------------------------------------------------------------------------------'
Function CloseDocIgnoreConfirmation(Optional Document As Object)

    If IsMissing(Document) Then
        Document = ThisComponent
    End If

    Document.Close(TRUE)
  
End Function