' Saves document in Microsoft Excel 97/2000/XP XLS format.                                         '
' FilterName types can be found in "API Name" column at:                                           '
' https://wiki.openoffice.org/wiki/Framework/Article/Filter/FilterList_OOo_2_1                     '
' Manual to the ThisComponent.storeAsURL last seen at:                                             '
' https://wiki.openoffice.org/wiki/Documentation/BASIC_Guide/StarDesktop                           '
' ByVal keyword prevents FilePath from modification because by default arguments are passed by     '
' reference.                                                                                       '
Function SaveDocAsXLS(ByVal FilePath As String)
    
    Dim SaveParams(1) As New com.sun.star.beans.PropertyValue
    SaveParams(0).Name = "Overwrite"
    SaveParams(0).Value = TRUE
    SaveParams(1).Name = "FilterName"
    SaveParams(1).Value = "MS Excel 97"
    ThisComponent.storeAsURL(ConvertToURL(FilePath),SaveParams())
    
End Function