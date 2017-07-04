'--------------------------------------------------------------------------------------------------'
' SaveDocAsODS                                                                                    '
'--------------------------------------------------------------------------------------------------'
' Saves document in OpenDocument Spreadsheet format (ODS).                                         '
'                                                                                                  '
' Parameters:                                                                                      '
'                                                                                                  '
'   ByVal FilePath As String                                                                       '
'     Path to save. ByVal keyword prevents it from modification because by default arguments are   '
'     passed by reference.                                                                         '
'                                                                                                  '
' Examples:                                                                                        '
'--------------------------------------------------------------------------------------------------'
'                                                                                                  '
'     SaveDocAsODS("/home/user/document.ods")                                                      '
' or                                                                                               '
'     SaveDocAsODS("C:\Users\Admin\Рабочий стол\document.ods")                                     '
'                                                                                                  '
'--------------------------------------------------------------------------------------------------'
' See also:                                                                                        '
'   FilterName types can be found in "API Name" column at:                                         '
'     https://wiki.openoffice.org/wiki/Framework/Article/Filter/FilterList_OOo_2_1                 '
'   Manual to the ThisComponent.storeAsURL last seen at:                                           '
'     https://wiki.openoffice.org/wiki/Documentation/BASIC_Guide/StarDesktop                       '
'--------------------------------------------------------------------------------------------------'
' Feedback & Issues:                                                                               '
'   https://github.com/aa6/libreoffice_calc_basic_extras/issues                                    '
'--------------------------------------------------------------------------------------------------'
Function SaveDocAsODS(Optional FilePath As String, Optional Document As Object)
    
    Dim SaveParams(1) As New com.sun.star.beans.PropertyValue
    SaveParams(0).Name = "Overwrite"
    SaveParams(0).Value = TRUE
    SaveParams(1).Name = "FilterName"
    SaveParams(1).Value = "calc8"

    If IsMissing(Document) Then
        Document = ThisComponent
    End If
    If IsMissing(FilePath) Then
        If Document.hasLocation() Then
            FilePath = Document.getLocation()
        Else
            Err.Raise("Document has no default path to save as.")
        End If
    End If
    FilePath = ConvertToURL(FilePath)
    If FileExists(FilePath) Then
        For Iteration = 1 To 1000
            TemporaryPath = FilePath + "." + CStr(CLng(999999 * Rnd)) + ".tmp"
            If NOT FileExists(TemporaryPath) Then
                Document.storeAsURL(TemporaryPath,SaveParams())
                Kill(FilePath)
                Document.storeAsURL(FilePath,SaveParams())
                Kill(TemporaryPath)
                Exit Function
            End If
        Next Iteration
        Err.Raise("Can't generate temporary path to save file.")
    Else
        Document.storeAsURL(FilePath,SaveParams())
    End If
    
End Function