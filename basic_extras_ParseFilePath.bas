' Structure that used to store data about a parsed path.                                           '
Type FilePathParsedByParseFilePathFunction
  
    FileDir As String
    FileName As String
    FileDirName As String
    FileFullPath As String
    FileExtension As String
    FileDirNoSlash As String
    FileNameNoExtension As String
  
End Type

' Returns FilePathParsedByParseFilePathFunction object filled with parsed values.                  '
' ByVal keyword prevents FullPath from modification because by default arguments are passed by     '
' reference.                                                                                       '
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
    basenameextdiff = basenamelen - Len(ParseFilePath.FileExtension) - 1
    If basenameextdiff < 0 Then
        basenameextdiff = 0
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