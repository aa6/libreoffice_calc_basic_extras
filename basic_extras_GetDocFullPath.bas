' Returns document full path if document has a path. Returns empty string if document has no path  '
' and IgnoreNoPathError flag is set to TRUE.                                                       '
' GetDocFullPath has variable amount of arguments. You can omit Document argument and pass the     '
' IgnoreNoPathError as a first argument.                                                           '
'                                                                                                  '
' Example #1:                                                                                      '
' Will return current opened document path or raises and error if document is not saved.           '
'                                                                                                  '
'     GetDocFullPath()                                                                             '
'                                                                                                  '
' Example #2:                                                                                      '
' Will return current opened document path or "" if document is not saved.                         '
'                                                                                                  '
'     GetDocFullPath(TRUE)                                                                         '
'                                                                                                  '
'                                                                                                  '
'   ParseFirstNumbers("-price is 2425 , 93 snsn 4223") REM Returns "-2425934223"                   '
'   ParseFirstNumbers("-price is 2425 , 93 snsn 4223", FALSE) REM Returns "2425934223"             '
Function GetDocFullPath(Optional Document as Variant, Optional IgnoreNoPathError as Boolean) As String

  If IsMissing(Document) Then
    Document = ThisComponent
  ElseIf TypeName(Document) = "Boolean" Then
    IgnoreNoPathError = Document
    Document = ThisComponent
  End If
    
    ' Default behavior is to return empty string if document location is empty.'
    ' This can happen if document is new and not saved thus do not have a path.'
    If NOT Document.hasLocation() AND IgnoreNoPathError <> TRUE Then
        ' Err.Raise is not valid statement but will generate error anyway.     '
        Err.Raise("Document has no path. Probably because it is not saved.")
    End If
    
    GetDocFullPath = ConvertFromURL(Document.getLocation())
    
End Function