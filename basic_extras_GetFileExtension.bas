'--------------------------------------------------------------------------------------------------'
' GetFileExtension                                                                                 '
'--------------------------------------------------------------------------------------------------'
' Returns parsed file extension of a given path.                                                   '
' Notice that files starting with dot (.htaccess) are considered to have no extension.             '
'                                                                                                  '
' Parameters:                                                                                      '
'                                                                                                  '
'   ByVal FullPath As String                                                                       '
'     Path to parse. ByVal keyword prevents it from modification because by default arguments are  '
'     passed by reference.                                                                         '
'                                                                                                  '
' Examples:                                                                                        '
'--------------------------------------------------------------------------------------------------'
'                                                                                                  '
'     extension = GetFileExtension(path)                                                           '
'                                                                                                  '
' Expected values:                                                                                 '
'   path: "/home/user/.htaccess"                          extension: ""                            '
'   path: "/home/user/document.ods"                       extension: "ods"                         '
'   path: "user/document.ods"                             extension: "ods"                         '
'   path: "C:\User\Admin\Рабочий стол\document.name.ods"  extension: "ods"                         '
'--------------------------------------------------------------------------------------------------'
Function GetFileExtension(ByVal FullPath As String) As String

    Dim pos As Long
    Dim pathlen As Long
    Dim pathurl As String
    Dim basename As String
    Dim basenamelen As Long
    Dim basenamelastdotindex As Long
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
    If Left(pathurl, 7) <> "file://" Then 
        pathurl = ConvertToURL("/" + FullPath)
    End If
    pathlen = Len(pathurl)
    For pos = pathlen To 1 Step -1
        If Mid(pathurl, pos, 1) = "/" Then
            basename = ConvertFromURL(Right(pathurl, pathlen - pos))
            Exit For
        End If
    Next pos
    
    ' Finding last occurence of "." in the file name. First symbol is ignored because filenames    '
    ' starting with dot (.htaccess) are considered to have no extension.                           '
    basenamelen = Len(basename)
    basenamelastdotindex = basenamelen
    For pos = basenamelen To 2 Step -1
        If Mid(basename, pos, 1) = "." Then
            basenamelastdotindex = pos
            Exit For
        End If
    Next pos
    
    GetFileExtension = Right(basename,basenamelen - basenamelastdotindex)

End Function