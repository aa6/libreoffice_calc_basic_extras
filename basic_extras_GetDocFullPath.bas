'--------------------------------------------------------------------------------------------------'
' GetDocFullPath                                                                                   '
'--------------------------------------------------------------------------------------------------'
' Returns document full path if document has a path. Returns empty string if document has no path  '
' and IgnoreNoPathError flag is set to TRUE.                                                       '
'                                                                                                  '
' Parameters:                                                                                      '
'                                                                                                  '
'   Optional Document as Variant <Default = ThisComponent>                                         '
'     Reference to a document (com.sun.star.uno.XInterface).                                       '
'                                                                                                  '
'   Optional IgnoreNoPathError as Boolean <Default = FALSE>                                        '
'     Will return empty string instead of raising an error if document has no path (is not saved). '
'                                                                                                  '
' Examples:                                                                                        '
'--------------------------------------------------------------------------------------------------'
'                                                                                                  '
'     docpath = GetDocFullPath()                                                                   '
'                                                                                                  '
' Expected values:                                                                                 '
'   docpath: "/home/user/current_opened_document.ods"                                              '
'                                                                                                  '
' Will raise an error if document is not saved.                                                    '
'--------------------------------------------------------------------------------------------------'
'                                                                                                  '
'     docpath = GetDocFullPath(docref)                                                             '
'                                                                                                  '
' Expected values:                                                                                 '
'   docpath: "/home/user/docref_document.ods"                                                      '
'                                                                                                  '
' Will raise an error if document is not saved.                                                    '
'--------------------------------------------------------------------------------------------------'
'                                                                                                  '
'     docpath = GetDocFullPath(docref, TRUE)                                                       '
'                                                                                                  '
' Expected values:                                                                                 '
'   docpath: "/home/user/docref_document.ods"                                                      '
'                                                                                                  '
' Will NOT raise an error if document is not saved.                                                '
'--------------------------------------------------------------------------------------------------'
' GetDocFullPath has variable amount of parameters. You can omit Document parameter and pass the   '
' IgnoreNoPathError as a first parameter.                                                          '
'                                                                                                  '
'     docpath = GetDocFullPath(TRUE)                                                               '
'                                                                                                  '
' Expected values:                                                                                 '
'   docpath: "/home/user/current_opened_document.ods"                                              '
'                                                                                                  '
' Will NOT raise an error if document is not saved.                                                '
'--------------------------------------------------------------------------------------------------'
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