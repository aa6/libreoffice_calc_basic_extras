'--------------------------------------------------------------------------------------------------'
' GetBaseDirectory                                                                                 '
'--------------------------------------------------------------------------------------------------'
' Returns parsed base directory of a given path.                                                   '
'                                                                                                  '
' Parameters:                                                                                      '
'                                                                                                  '
'   ByVal FullPath As String                                                                       '
'     Path to parse. ByVal keyword prevents it from modification because by default arguments are  '
'     passed by reference.                                                                         '
'                                                                                                  '
'   Optional DropTrailingSlash As Boolean <Default = FALSE>                                        '
'     Drops trailing slash at the end of parsed base directory.                                    '
'                                                                                                  '
' Examples:                                                                                        '
'--------------------------------------------------------------------------------------------------'
'                                                                                                  '
'     basedir = GetBaseDirectory(path)                                                             '
'                                                                                                  '
' Expected values:                                                                                 '
'   path: "/home/user/document.ods"    basedir: "/home/user/"                                      '
'   path: "/home/user/"                basedir: "/home/user/"                                      '
'   path: "/home/user"                 basedir: "/home/"                                           '
'   path: "C:\User\Admin\Рабочий стол" basedir: "C:\User\Admin\"                                   '
'   path: "C:\User\Admin\"             basedir: "C:\User\Admin\"                                   '
'   path: "C:\User\Admin"              basedir: "C:\User\"                                         '
'--------------------------------------------------------------------------------------------------'
'                                                                                                  '
'     basedir = GetBaseDirectory(path, TRUE)                                                       '
'                                                                                                  '
' Expected values:                                                                                 '
'   path: "/home/user/document.ods"    basedir: "/home/user"                                       '
'   path: "/home/user/"                basedir: "/home/user"                                       '
'   path: "/home/user"                 basedir: "/home"                                            '
'   path: "C:\User\Admin\Рабочий стол" basedir: "C:\User\Admin"                                    '
'   path: "C:\User\Admin\"             basedir: "C:\User\Admin"                                    '
'   path: "C:\User\Admin"              basedir: "C:\User"                                          '
'--------------------------------------------------------------------------------------------------'
Function GetBaseDirectory(ByVal FullPath As String, Optional DropTrailingSlash As Boolean) As String

    Dim i As Long
    Dim pathlen As Long
    Dim pathurl As String
    Dim lendiff As Long
    Dim basename As String
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
    For i = pathlen To 1 Step -1
        If Mid(pathurl,i,1) = "/" Then
            basename = ConvertFromURL(Right(pathurl,pathlen - i))
            Exit For
        End If
    Next i
    
    If DropTrailingSlash = TRUE Then
        basename = "/" + basename
    End If
    lendiff = Len(FullPath) - Len(basename)
    If lendiff < 0 Then
        lendiff = 0
    End If
    
    GetBaseDirectory = Left(FullPath,lendiff)

End Function