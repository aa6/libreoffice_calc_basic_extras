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
Function SaveDocAsODS(ByVal FilePath As String)
    
    Dim SaveParams(1) As New com.sun.star.beans.PropertyValue
    SaveParams(0).Name = "Overwrite"
    SaveParams(0).Value = TRUE
    SaveParams(1).Name = "FilterName"
    SaveParams(1).Value = "calc8"
    ThisComponent.storeAsURL(ConvertToURL(FilePath),SaveParams())
    
End Function