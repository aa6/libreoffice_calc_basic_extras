'--------------------------------------------------------------------------------------------------'
' ParseFilePath                                                                                    '
'--------------------------------------------------------------------------------------------------'
' Returns object filled with parsed path parts.                                                    '
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
'     parsed_path = ParseFilePath(path)                                                            '
'     MsgBox(parsed_path.FileNameNoExtension)                                                      '
'                                                                                                  '
' Expected values:                                                                                 '
'                                                                                                  '
'   path: "/home/user/.htaccess"                                                                   '
'   parsed_path:                                                                                   '
'     FilePathParsedByParseFilePathFunction                                                        '
'       .FileDir              string "/home/user/"                                                 '
'       .FileName             string ".htaccess"                                                   '
'       .FileDirName          string "user"                                                        '
'       .FileFullPath         string "/home/user/.htaccess"                                        '
'       .FileExtension        string ""                                                            '
'       .FileDirNoSlash       string "/home/user"                                                  '
'       .FileNameNoExtension  string ".htaccess"                                                   '
'                                                                                                  '
'   path: "/home/user/document.ods"                                                                '
'   parsed_path:                                                                                   '
'     FilePathParsedByParseFilePathFunction                                                        '
'       .FileDir              string "/home/user/"                                                 '
'       .FileName             string "document.ods"                                                '
'       .FileDirName          string "user"                                                        '
'       .FileFullPath         string "/home/user/document.ods"                                     '
'       .FileExtension        string "ods"                                                         '
'       .FileDirNoSlash       string "/home/user"                                                  '
'       .FileNameNoExtension  string "document"                                                    '
'                                                                                                  '
'   path: "user/document.ods"                                                                      '
'   parsed_path:                                                                                   '
'     FilePathParsedByParseFilePathFunction                                                        '
'       .FileDir              string "user/"                                                       '
'       .FileName             string "document.ods"                                                '
'       .FileDirName          string "user"                                                        '
'       .FileFullPath         string "user/document.ods"                                           '
'       .FileExtension        string "ods"                                                         '
'       .FileDirNoSlash       string "user"                                                        '
'       .FileNameNoExtension  string "document"                                                    '
'                                                                                                  '
'   path: "C:\Users\Admin\Рабочий стол\document.name.ods"                                          '
'   parsed_path:                                                                                   '
'     FilePathParsedByParseFilePathFunction                                                        '
'       .FileDir              string "C:\Users\Admin\Рабочий стол\"                                '
'       .FileName             string "document.name.ods"                                           '
'       .FileDirName          string "Рабочий стол"                                                '
'       .FileFullPath         string "C:\Users\Admin\Рабочий стол\document.name.ods"               '
'       .FileExtension        string "ods"                                                         '
'       .FileDirNoSlash       string "C:\Users\Admin\Рабочий стол"                                 '
'       .FileNameNoExtension  string "document.name"                                               '
'--------------------------------------------------------------------------------------------------'
' Feedback & Issues:                                                                               '
'   https://github.com/aa6/libreoffice_calc_basic_extras/issues                                    '
'--------------------------------------------------------------------------------------------------'
Type FilePathParsedByParseFilePathFunction
  
    FileDir As String
    FileName As String
    FileDirName As String
    FileFullPath As String
    FileExtension As String
    FileDirNoSlash As String
    FileNameNoExtension As String
  
End Type
                                                                                  '
Function ParseFilePath(ByVal FullPath As String) As FilePathParsedByParseFilePathFunction
    
    Dim pos As Long
    Dim pathlen As Long
    Dim pathurl As String
    Dim dirlendiff As Long
    Dim basenamelen As Long
    Dim basedirectory As String
    Dim basenameextdiff As Long
    Dim basenamelastdotindex As Long
    ' Fetching file base name from FullPath.                                                       '
    ' Converting to URL for Linux/Windows compatibility.                                           '
    '   URL notation does not allow certain special characters to be used. These are either        '
    '   replaced by other characters or encoded. A slash (/) is used as a path separator. For      '
    '   example, a file referred to as C:\My File.sxw on the local host in "Windows notation"      '
    '   becomes file:///C|/My%20File.sxw in URL notation.                                          '
    ' https://help.libreoffice.org/Basic/Basic_Glossary                                            '
    pathurl = ConvertToURL(FullPath)
    ParseFilePath = CreateObject("FilePathParsedByParseFilePathFunction")
    ParseFilePath.FileFullPath = FullPath
    ' FullPath could be mistakenly converted to http. For example:                                 '
    ' ConvertToURL("many.dots.in.file.name.ods") will be misinterpreted.                           ' 
    If Left(pathurl, 7) <> "file://" Then 
        pathurl = ConvertToURL("/" + FullPath)
    End If
    pathlen = Len(pathurl)
    For pos = pathlen To 1 Step -1
        If Mid(pathurl, pos, 1) = "/" Then
            ParseFilePath.FileName = ConvertFromURL(Right(pathurl, pathlen - pos))
            Exit For
        End If
    Next pos
    ' Finding last occurence of "." in the file name. First symbol is ignored because filenames    '
    ' starting with dot (.htaccess) are considered to have no extension.                           '
    basenamelen = Len(ParseFilePath.FileName)
    basenamelastdotindex = basenamelen
    For pos = basenamelen To 2 Step -1
        If Mid(ParseFilePath.FileName, pos, 1) = "." Then
            basenamelastdotindex = pos
            Exit For
        End If
    Next pos
    ParseFilePath.FileExtension = Right(ParseFilePath.FileName, basenamelen - basenamelastdotindex)
    basenameextdiff = basenamelen - Len(ParseFilePath.FileExtension)
    If Len(ParseFilePath.FileExtension) > 0 Then
        basenameextdiff = basenameextdiff - 1 ' Dot separator. '
    End If
    ParseFilePath.FileNameNoExtension = Left(ParseFilePath.FileName, basenameextdiff)
    ' Getting directory name with slash and without. '
    dirlendiff = Len(FullPath) - Len(ParseFilePath.FileName)
    ParseFilePath.FileDir = Left(FullPath, dirlendiff)
    dirlendiff = dirlendiff - 1
    If dirlendiff < 0 Then
        dirlendiff = 0
    End If
    ParseFilePath.FileDirNoSlash = Left(FullPath, dirlendiff)
    ' Getting file directory name. '
    pathurl = ConvertToURL(ParseFilePath.FileDirNoSlash)
    If Left(pathurl, 7) <> "file://" Then 
        pathurl = ConvertToURL("/" + ParseFilePath.FileDirNoSlash)
    End If
    pathlen = Len(pathurl)
    For pos = pathlen To 1 Step -1
        If Mid(pathurl, pos, 1) = "/" Then
            ParseFilePath.FileDirName = ConvertFromURL(Right(pathurl, pathlen - pos))
            Exit For
        End If
    Next pos
    
End Function