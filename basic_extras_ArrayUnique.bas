'--------------------------------------------------------------------------------------------------'
' ArrayUnique                                                                                      '
'--------------------------------------------------------------------------------------------------'
' Returns an array containing elements of input array without duplicate values.                    '
'                                                                                                  '
' Parameters:                                                                                      '
'                                                                                                  '
'   Arr As Variant                                                                                 '
'     The input array.                                                                             '
'                                                                                                  '
' Examples:                                                                                        '
'--------------------------------------------------------------------------------------------------'
'                                                                                                  '
'     result_array = ParseFirstNumbers(input_array)                                                '
'                                                                                                  '
' Expected values:                                                                                 '
'                                                                                                  '
'   input_array:                                                                                   '
'     Array(10,"vodka","10","beer","water",12,"beer","applejuice",12)                              '
'   result_array:                                                                                  '
'     Array(10,"vodka","10","beer","water",12,"applejuice")                                        '
'--------------------------------------------------------------------------------------------------'
'                                                                                                  '
'     Sub TestArrayUnique                                                                          '
'         Dim item As Variant                                                                      '
'         Dim result As String                                                                     '
'         Dim uniqarr As Variant                                                                   '
'         result = ""                                                                              '
'         uniqarr = ArrayUnique(Array(10,"vodka","10","beer","water",12,"beer","applejuice",12))   '
'         For Each item In uniqarr                                                                 '
'             result = result + item + ","                                                         '
'         Next item                                                                                '
'         MsgBox(result)                                                                           '
'     End Sub                                                                                      '
'                                                                                                  '
'--------------------------------------------------------------------------------------------------'
Function ArrayUnique(Arr As Variant)

    Dim item As Variant
    Dim entry As Variant
    Dim result() As Variant

    For Each item In Arr
        For Each entry In result
            If item = entry Then
                Goto ArrayUniqueNextArrItem
            End If
        Next entry
        Redim Preserve result(UBound(result) + 1) As Variant
        result(UBound(result)) = item
        ArrayUniqueNextArrItem:
    Next item
    ArrayUnique = result

End Function