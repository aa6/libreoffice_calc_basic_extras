' Saves document in OpenDocument Spreadsheet Flat XML format (FODS).                               '
' FilterName types can be found in "API Name" column at:                                           '
' https://wiki.openoffice.org/wiki/Framework/Article/Filter/FilterList_OOo_2_1                     '
' Manual to the ThisComponent.storeAsURL last seen at:                                             '
' https://wiki.openoffice.org/wiki/Documentation/BASIC_Guide/StarDesktop                           '
' ByVal keyword prevents FilePath from modification because by default arguments are passed by     '
' reference.                                                                                       '
Function SaveDocAsFODS(ByVal FilePath As String)
    
    Dim SaveParams(1) As New com.sun.star.beans.PropertyValue
    SaveParams(0).Name = "Overwrite"
    SaveParams(0).Value = TRUE
    SaveParams(1).Name = "FilterName"
    SaveParams(1).Value = "OpenDocument Spreadsheet Flat XML"
    ThisComponent.storeAsURL(ConvertToURL(FilePath),SaveParams())
    
End Function