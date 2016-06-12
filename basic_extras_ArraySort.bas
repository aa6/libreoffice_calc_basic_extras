'--------------------------------------------------------------------------------------------------'
' ArraySort                                                                                        '
'--------------------------------------------------------------------------------------------------'
' Returns ascending sorted copy of the input array.                                                '
'                                                                                                  '
' Parameters:                                                                                      '
'                                                                                                  '
'   Arr As Variant                                                                                 '
'     The input array.                                                                             '
'                                                                                                  '
' Examples:                                                                                        '
'--------------------------------------------------------------------------------------------------'
'                                                                                                  '
'     sorted_array = ArraySort(input_array)                                                        '
'                                                                                                  '
' Expected values:                                                                                 '
'                                                                                                  '
'   input_array:                                                                                   '
'     Array("12",12,10,"10","",2,20,"beer",1,"water","soda","beer","applejuice",12)                '
'   sorted_array:                                                                                  '
'     Array("",1,10,"10","12",12,12,2,20,"applejuice","beer","beer","soda","water")                '
'--------------------------------------------------------------------------------------------------'
'                                                                                                  '
'     Sub TestArraySort                                                                            '
'         Dim item As Variant                                                                      '
'         Dim result As String                                                                     '
'         Dim inputarr As Variant                                                                  '
'         Dim resultarr As Variant                                                                 '
'         result = ""                                                                              '
'         inputarr = Array("12",12,10,"10","",2,20,"beer",1,"water","soda","beer","applejuice",12) '
'         resultarr = ArraySort(inputarr)                                                          '
'         For Each item In resultarr                                                               '
'             result = result + IIf(TypeName(item) = "String", """" + item + """", item) + ","     '
'         Next item                                                                                '
'         MsgBox(result)                                                                           '
'     End Sub                                                                                      '
'                                                                                                  '
'--------------------------------------------------------------------------------------------------'
' Feedback & Issues:                                                                               '
'   https://github.com/aa6/libreoffice_calc_basic_extras/issues                                    '
'--------------------------------------------------------------------------------------------------'
Function ArraySort(Arr As Variant)

    Dim i As Long
    Dim j As Long
    Dim swap As Variant
    Dim result(UBound(Arr)) As Variant

    For i = 0 To Ubound(Arr)
        result(i) = Arr(i)
    Next i
    For i = 0 To Ubound(result)
        For j = 0 To Ubound(result)
            If CStr(result(i)) < CStr(result(j)) Then
                swap = result(i)
                result(i) = result(j)
                result(j) = swap
            End If
        Next j
    Next i
    ArraySort = result

End Function