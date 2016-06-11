' Returns the filename of the given path.                                                          '
' For example "/home/user/document.ods" input will return "document.ods".                          '
' ByVal keyword prevents FullPath from modification because by default arguments are passed by     '
' reference.                                                                                       '
Function GetFileBaseName(ByVal FullPath As String) As String

    Dim pos As Long
    Dim pathlen As Long
    Dim pathurl As String
    ' Fetching file base name from FullPath.                                                       '
    ' Converting to URL for Linux/Windows compatibility.                                           '
    '   URL notation does not allow certain special characters to be used. These are either        '
    '   replaced by other characters or encoded. A slash (/) is used as a path separator. For      '
    '   example, a file referred to as C:\My File.sxw on the local host in "Windows notation"      '
    '   becomes file:///C|/My%20File.sxw in URL notation.                                          '
    ' https://help.libreoffice.org/Basic/Basic_Glossary                                            '
    pathurl = ConvertToURL(FullPath)
    ' FullPath could be mistakenly converted to http. For example:                                 '
    ' ConvertToURL("many.dots.in.file.name.ods") will be misinterpreted.                           '
    If Left(pathurl,7) <> "file://" Then 
        pathurl = ConvertToURL("/" + FullPath)
    End If
    pathlen = Len(pathurl)
    For pos = pathlen To 1 Step -1
        If Mid(pathurl, pos, 1) = "/" Then
            GetFileBaseName = ConvertFromURL(Right(pathurl, pathlen - pos))
            Exit For
        End If
    Next pos

End Function