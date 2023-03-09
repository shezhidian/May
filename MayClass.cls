VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MayClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum EnumByRowOrColumn
    ByColumn
    ByRow
End Enum

Public Enum EnumFromFirstOrLast
    FromFirst
    FromLast
End Enum

Public Enum EnumGetLeftOrRight
    GetLeft
    GetRight
End Enum
Public Function ArraySortBy(sourceData As Variant, columnData1 As Variant, rule1 As Variant, Optional columnData2 As Variant = Empty, Optional rule2 As Variant = Empty, Optional columnData3 As Variant = Empty, Optional rule3 As Variant = Empty, Optional columnData4 As Variant = Empty, Optional rule4 As Variant = Empty, Optional columnData5 As Variant = Empty, Optional rule5 As Variant = Empty, Optional columnData6 As Variant = Empty, Optional rule6 As Variant = Empty)
10        On Error GoTo ArraySortBy_Error
          Dim res As Variant
20        If IsEmpty(columnData2) Or IsEmpty(rule2) Then
30            res = WorksheetFunction.SortBy(sourceData, columnData1, rule1)
40        ElseIf IsEmpty(columnData3) Or IsEmpty(rule3) Then
50            res = WorksheetFunction.SortBy(sourceData, columnData1, rule1, columnData2, rule2)
60        ElseIf IsEmpty(columnData4) Or IsEmpty(rule4) Then
70            res = WorksheetFunction.SortBy(sourceData, columnData1, rule1, columnData2, rule2, columnData3, rule3)
80        ElseIf IsEmpty(columnData5) Or IsEmpty(rule5) Then
90            res = WorksheetFunction.SortBy(sourceData, columnData1, rule1, columnData2, rule2, columnData3, rule3, columnData4, rule4)
100       ElseIf IsEmpty(columnData6) Or IsEmpty(rule6) Then
110           res = WorksheetFunction.SortBy(sourceData, columnData1, rule1, columnData2, rule2, columnData3, rule3, columnData4, rule4, columnData5, rule5)
120       Else
130           res = WorksheetFunction.SortBy(sourceData, columnData1, rule1, columnData2, rule2, columnData3, rule3, columnData4, rule4, columnData5, rule5, columnData6, rule6)
140       End If
          
150       ArraySortBy = res
160       Exit Function
ArraySortBy_Error:
170       ArraySortBy = "#ArraySortBy:line(" & CStr(Erl) & ")" & Err.Description & "!"

End Function


Public Function CountChar(str_input As String, targetChar As String)
10        On Error GoTo countChar_Error
          Dim res As Variant
          Dim i As Long
          Dim j As Long
20        res = 0
30        For i = 1 To Len(str_input)
40           If Mid(str_input, i, 1) = targetChar Then
50            res = res + 1
60           End If
70        Next i
          
80        CountChar = res
90        Exit Function
countChar_Error:
100       CountChar = "#countChar:line(" & CStr(Erl) & ")" & Err.Description & "!"

End Function

Public Function CompareTwoArrays(data1 As Variant, data2 As Variant)
10        On Error GoTo CompareTwoArrays_Error
          Dim res As Variant
          Dim res_title As Variant
          Dim item As Variant
          Dim num_InBoth As Long
          Dim num_In1 As Long
          Dim num_In2 As Long
          Dim max_num As Long
          Dim itemCountData1 As Long
          Dim itemCountData2 As Long
          Dim po As New MayClassPosition
          
20        If IsEmpty(data1) Or IsEmpty(data2) Then
30            res = "#data1 or data2 is empty!"
40            GoTo CompareTwoArrays_end
50        End If
          
60        res_title = ArrayTranspose(To2DArray(Array("InBoth", "In1NotIn2", "In2NotIn1")))
70        num_InBoth = 0
80        num_In1 = 0
90        num_In2 = 0
          
          
100       itemCountData1 = ArrayRowsCount(data1) * ArrayColumnsCount(data1)
110       itemCountData2 = ArrayRowsCount(data2) * ArrayColumnsCount(data2)
120       If itemCountData1 >= itemCountData2 Then
130           res = ArrayCreate(itemCountData1, 3)
140       Else
150           res = ArrayCreate(itemCountData2, 3)
160       End If
          
170       For Each item In data1
180           ArrayFindFirstPosition data2, item, po, ByColumn
190           If po.columnNum > 0 Or po.rowNum > 0 Then
200               num_InBoth = num_InBoth + 1
210               res(num_InBoth, 1) = item
220           Else
230               num_In1 = num_In1 + 1
240               res(num_In1, 2) = item
250           End If
260       Next item
          
270       For Each item In data2
280           ArrayFindFirstPosition data1, item, po, ByColumn
290           If po.columnNum = 0 And po.rowNum = 0 Then
300               num_In2 = num_In2 + 1
310               res(num_In2, 3) = item
320           End If
330       Next item
          
          
340       If num_In1 = 0 And num_In2 = 0 And num_InBoth = 0 Then
350           res = "#no data in two arrays!"
360           GoTo CompareTwoArrays_end
370       End If
          
380       max_num = ArrayMax(To2DArray(Array(num_In1, num_In2, num_InBoth)))
390       res = SubArray(res, 1, 1, max_num)
400       res = ArrayStack(res_title, res)
          
CompareTwoArrays_end:
410       CompareTwoArrays = res
420       Exit Function
CompareTwoArrays_Error:
430       CompareTwoArrays = "#CompareTwoArrays:line(" & CStr(Erl) & ")" & Err.Description & "!"

End Function

Public Function GetOneLable(twoColumnData As Variant, label As Variant)
10        On Error GoTo getOneLable_Error
          Dim res As Variant
          Dim i As Long

20        res = Empty
30        For i = 1 To ArrayRowsCount(twoColumnData)
40            If twoColumnData(i, 1) = label Then
50                res = twoColumnData(i, 2)
60                Exit For
70            End If
80        Next i
          
90        If IsEmpty(res) Then
100           res = "#can't find label:" & label & "!"
110       End If
          
120       GetOneLable = res
130       Exit Function
getOneLable_Error:
140       GetOneLable = "#getOneLable:line(" & CStr(Erl) & ")" & Err.Description & "!"

End Function


Public Function ArrayRowAnd(chooser1 As Variant, chooser2 As Variant)
10        On Error GoTo ArrayRowAnd_Error
          Dim res As Variant
          Dim i As Long
          Dim j As Long

20        If ArrayRowsCount(chooser1) <> ArrayRowsCount(chooser2) Then
30            res = "#diffrent rows count for chooser1 and chooser2!"
40        Else
50            res = ArrayCreate(ArrayRowsCount(chooser1), 1, False)
60            For i = 1 To ArrayRowsCount(chooser1)
70                If chooser1(i, 1) = True And chooser2(i, 1) = True Then
80                    res(i, 1) = True
90                End If
100           Next i
110       End If
          
120       ArrayRowAnd = res
130       Exit Function
ArrayRowAnd_Error:
140       ArrayRowAnd = "#ArrayRowAnd:line(" & CStr(Erl) & ")" & Err.Description & "!"

End Function

Public Function ArrayRowOr(chooser1 As Variant, chooser2 As Variant)
10        On Error GoTo ArrayRowOr_Error
          Dim res As Variant
          Dim i As Long
          Dim j As Long
          
20        If ArrayRowsCount(chooser1) <> ArrayRowsCount(chooser2) Then
30            res = "#diffrent rows count for chooser1 and chooser2!"
40        Else
50            res = ArrayCreate(ArrayRowsCount(chooser1), 1, False)
60            For i = 1 To ArrayRowsCount(chooser1)
70                If chooser1(i, 1) = True Or chooser2(i, 1) = True Then
80                    res(i, 1) = True
90                End If
100           Next i
110       End If
          
120       ArrayRowOr = res
130       Exit Function
ArrayRowOr_Error:
140       ArrayRowOr = "#ArrayRowOr:line(" & CStr(Erl) & ")" & Err.Description & "!"

End Function

Public Function ArrayRound(data As Variant, decimalNum As Long)
10        On Error GoTo ArrayRound_Error
          Dim i As Long
          Dim j As Long

20        For i = 1 To ArrayRowsCount(data)
30            For j = 1 To ArrayColumnsCount(data)
40                If IsNumeric(data(i, j)) And Not IsEmpty(data(i, j)) Then
50                    data(i, j) = Round(data(i, j), decimalNum)
60                End If
70            Next j
80        Next i
          
90        ArrayRound = data
100       Exit Function
ArrayRound_Error:
110       ArrayRound = "#ArrayRound:line(" & CStr(Erl) & ")" & Err.Description & "!"

End Function

Public Function ArrayMin(data As Variant)
10        On Error GoTo ArrayMin_Error
          Dim res As Variant
          Dim item As Variant
20        res = Empty
30        For Each item In data
40            If IsNumeric(item) And Not IsEmpty(item) Then
50                If IsEmpty(res) Then
60                    res = item
70                Else
80                    If item < res Then
90                        res = item
100                   End If
110               End If
120           End If
130       Next item
          
140       If IsEmpty(res) Then
150           res = "#No number in array!"
160       End If

170       ArrayMin = res
180       Exit Function
ArrayMin_Error:
190       ArrayMin = "#ArrayMin:line(" & CStr(Erl) & ")" & Err.Description & "!"

End Function

Public Function ArrayAvg(data As Variant, Optional IfOnlyCountNumericCells As Boolean = True)
10        On Error GoTo ArrayAvg_Error
          Dim res As Variant
          Dim i As Long
          Dim item As Variant
20        res = Empty
30        i = 0
40        For Each item In data
50            If IsNumeric(item) And Not IsEmpty(item) Then
60                i = i + 1
70                If IsEmpty(res) Then
80                    res = item
90                Else
100                   res = res + item
110               End If
120           ElseIf IfOnlyCountNumericCells = False Then
130               i = i + 1
140           End If
150       Next item
          
160       If IsEmpty(res) Then
170           res = "#No number in array!"
180       Else
190           res = res / i
200       End If
          
210       ArrayAvg = res
220       Exit Function
ArrayAvg_Error:
230       ArrayAvg = "#ArrayAvg:line(" & CStr(Erl) & ")" & Err.Description & "!"

End Function

Public Function ArraySum(data As Variant)
10        On Error GoTo ArraySum_Error
          Dim res As Variant
          Dim item As Variant
20        res = Empty
30        For Each item In data
40            If IsNumeric(item) And Not IsEmpty(item) Then
50                If IsEmpty(res) Then
60                    res = item
70                Else
80                    res = res + item
90                End If
100           End If
110       Next item
          
120       If IsEmpty(res) Then
130           res = "#No number in array!"
140       End If
          
150       ArraySum = res
160       Exit Function
ArraySum_Error:
170       ArraySum = "#ArraySum:line(" & CStr(Erl) & ")" & Err.Description & "!"

End Function

Public Function ArrayMax(data As Variant)
10        On Error GoTo ArrayMax_Error
          Dim res As Variant
          Dim item As Variant
20        res = Empty
30        For Each item In data
40            If IsNumeric(item) And Not IsEmpty(item) Then
50                If IsEmpty(res) Then
60                    res = item
70                Else
80                    If item > res Then
90                        res = item
100                   End If
110               End If
120           End If
130       Next item
          
140       If IsEmpty(res) Then
150           res = "#No number in array!"
160       End If
170       ArrayMax = res
180       Exit Function
ArrayMax_Error:
190       ArrayMax = "#ArrayMax:line(" & CStr(Erl) & ")" & Err.Description & "!"

End Function

Public Function ArrayLCase(data As Variant)
10        On Error GoTo ArrayLCase_Error
          Dim i As Long
          Dim j As Long

20        For i = 1 To ArrayRowsCount(data)
30            For j = 1 To ArrayColumnsCount(data)
40                If TypeName(data(i, j)) = "String" Then
50                    data(i, j) = LCase(data(i, j))
60                End If
70            Next j
80        Next i

90        ArrayLCase = data
100       Exit Function
ArrayLCase_Error:
110       ArrayLCase = "#ArrayLCase:line(" & CStr(Erl) & ")" & Err.Description & "!"

End Function

Public Function ArrayUCase(data As Variant)
10        On Error GoTo ArrayUCase_Error
          Dim i As Long
          Dim j As Long

20        For i = 1 To ArrayRowsCount(data)
30            For j = 1 To ArrayColumnsCount(data)
40                If TypeName(data(i, j)) = "String" Then
50                    data(i, j) = UCase(data(i, j))
60                End If
70            Next j
80        Next i

90        ArrayUCase = data
          
100       Exit Function
ArrayUCase_Error:
110       ArrayUCase = "#ArrayUCase:line(" & CStr(Erl) & ")" & Err.Description & "!"

End Function

Public Function ArrayCStr(data As Variant, Optional IfAddStrPrefix As Boolean = False, Optional IfOnlyAddStrPrefixForNumbers As Boolean = True)
10        On Error GoTo ArrayCStr_Error
          Dim i As Long
          Dim j As Long

20        For i = 1 To ArrayRowsCount(data)
30            For j = 1 To ArrayColumnsCount(data)
40                If TypeName(data(i, j)) <> "String" And Not IsEmpty(data(i, j)) Then
50                    If IfAddStrPrefix = True Then
60                        If IfOnlyAddStrPrefixForNumbers = True Then
70                            If IsNumeric(data(i, j)) Then
80                                data(i, j) = "'" & CStr(data(i, j))
90                            End If
100                       Else
110                           data(i, j) = "'" & CStr(data(i, j))
120                       End If
130                   Else
140                       data(i, j) = CStr(data(i, j))
150                   End If
160               End If
170           Next j
180       Next i
          
190       ArrayCStr = data
200       Exit Function
ArrayCStr_Error:
210       ArrayCStr = "#ArrayCStr:line(" & CStr(Erl) & ")" & Err.Description & "!"

End Function

Public Function ArrayInstr(columnData As Variant, StrToMatch As String)
10        On Error GoTo ArrayInstr_Error
          Dim res As Variant
          Dim i As Long
20        res = ArrayCreate(ArrayRowsCount(columnData), 1)
30        For i = 1 To ArrayRowsCount(columnData)
40            res(i, 1) = InStr(columnData(i, 1), StrToMatch)
50        Next i
          
60        ArrayInstr = res
70        Exit Function
ArrayInstr_Error:
80        ArrayInstr = "#ArrayInstr:line(" & CStr(Erl) & ")" & Err.Description & "!"

End Function

Public Function ArrayReplace(data As Variant, valueToMatch As Variant, valueToSet As Variant)
10        On Error GoTo ArrayReplace_Error
          Dim i As Long
          Dim j As Long
          
20        For i = 1 To ArrayRowsCount(data)
30            For j = 1 To ArrayColumnsCount(data)
40                If data(i, j) = valueToMatch Then
50                    data(i, j) = valueToSet
60                End If
70            Next j
80        Next i

90        ArrayReplace = data
100       Exit Function
ArrayReplace_Error:
110       ArrayReplace = "#ArrayReplace:line(" & CStr(Erl) & ")" & Err.Description & "!"


End Function

Public Function ArrayNot(data As Variant)
10        On Error GoTo ArrayNot_Error
          Dim i As Long
          Dim j As Long
          
20        For i = 1 To ArrayRowsCount(data)
30            For j = 1 To ArrayColumnsCount(data)
40                If TypeName(data(i, j)) = "Boolean" Then
50                    data(i, j) = Not data(i, j)
60                End If
70            Next j
80        Next i

90        ArrayNot = data
100       Exit Function
ArrayNot_Error:
110       ArrayNot = "#ArrayNot:line(" & CStr(Erl) & ")" & Err.Description & "!"

End Function

Public Function ArrayColumnOr(columnData As Variant)
10        On Error GoTo ArrayColumnOr_Error
          Dim res As Variant
          Dim i As Long
20        res = False
30        For i = 1 To ArrayRowsCount(columnData)
40            If columnData(i, 1) = True Then
50                res = True
60                Exit For
70            End If
80        Next i
          
90        ArrayColumnOr = res
100       Exit Function
ArrayColumnOr_Error:
110       ArrayColumnOr = "#ArrayColumnOr:line(" & CStr(Erl) & ")" & Err.Description & "!"

End Function

Public Function ArrayColumnAnd(columnData As Variant)
10        On Error GoTo ArrayColumnAnd_Error
          Dim res As Variant
          Dim i As Long
20        res = True
30        For i = 1 To ArrayRowsCount(columnData)
40            If columnData(i, 1) = False Then
50                res = False
60                Exit For
70            End If
80        Next i

90        ArrayColumnAnd = res
100       Exit Function
ArrayColumnAnd_Error:
110       ArrayColumnAnd = "#ArrayColumnAnd:line(" & CStr(Erl) & ")" & Err.Description & "!"

End Function

Public Function ArrayStringJoin(columnData As Variant, delimiter As String)
10        On Error GoTo ArrayStringJoin_Error
          Dim res As Variant
          Dim i As Long

20        If ArrayRowsCount(columnData) <= 1 Then
30            res = columnData
40        Else
50            For i = 1 To ArrayRowsCount(columnData)
60                If i = ArrayRowsCount(columnData) Then
70                    res = res & columnData(i, 1)
80                Else
90                    res = res & columnData(i, 1) & delimiter
100               End If
110           Next i
120       End If
          
130       ArrayStringJoin = res
140       Exit Function
ArrayStringJoin_Error:
150       ArrayStringJoin = "#ArrayStringJoin:line(" & CStr(Erl) & ")" & Err.Description & "!"

End Function

Public Function ArrayTrim(data As Variant)
10        On Error GoTo ArrayTrim_Error
          Dim i As Long
          Dim j As Long
20        For i = 1 To UBound(data, 1)
30            For j = 1 To UBound(data, 2)
40                If TypeName(data(i, j)) = "String" And Not IsEmpty(data(i, j)) Then
50                    data(i, j) = Trim(data(i, j))
60                End If
70            Next j
80        Next i

90        ArrayTrim = data
100       Exit Function
ArrayTrim_Error:
110       ArrayTrim = "#ArrayTrim:line(" & CStr(Erl) & ")" & Err.Description & "!"

End Function

Public Function ArrayVLookup(sourceArray As Variant, targetColumnData As Variant, columnToSearch As Long, Optional columnsToTake As Variant = Empty, Optional IfIncludeTitle As Boolean = False)
10        On Error GoTo ArrayVLookup_Error
          Dim res As Variant
          Dim resCount As Variant
          Dim chooser As Variant
          Dim sourceColumnDataToSearch As Variant
          Dim i As Long
          Dim po As New MayClassPosition

          
20        sourceColumnDataToSearch = ArrayColumnTakeByNum(sourceArray, columnToSearch, IfIncludeTitle)
30        chooser = ArrayCreate(ArrayRowsCount(sourceColumnDataToSearch), 1, False)
40        For i = 1 To ArrayRowsCount(sourceColumnDataToSearch)
50            po.columnNum = 0
60            po.rowNum = 0
70            ArrayFindFirstPosition targetColumnData, sourceColumnDataToSearch(i, 1), po, ByColumn
80            If po.columnNum > 0 Or po.rowNum > 0 Then
90                chooser(i, 1) = True
100           End If
110       Next i
          
120       res = ArrayChooseTrue(SubArray(sourceArray, 2, 1), chooser)
130       If Not IsEmpty(columnsToTake) Then
140         res = ArrayColumnTakeByNums(res, columnsToTake, True)
150       End If

160       ArrayVLookup = res
170       Exit Function
ArrayVLookup_Error:
180       ArrayVLookup = "#ArrayVLookup:line(" & CStr(Erl) & ")" & Err.Description & "!"

End Function

Public Function ArrayStringBetweenString(data As Variant, delimiter As String, firstOrLast As EnumFromFirstOrLast, leftOrRight As EnumGetLeftOrRight)
10        On Error GoTo ArrayStringBetweenString_Error
          Dim i As Long
          Dim j As Long
          Dim num As Long
          Dim item As Variant
20        For i = 1 To ArrayRowsCount(data)
30            For j = 1 To ArrayColumnsCount(data)
40                item = data(i, j)
50                If TypeName(item) = "String" Then
60                    num = InStr(item, delimiter)
70                    If num > 0 Then
80                        If firstOrLast = FromLast Then
90                            num = InStrRev(item, delimiter)
100                       End If
                          
110                       If leftOrRight = GetLeft Then
120                           data(i, j) = Left(item, num - 1)
130                       Else
140                           data(i, j) = Right(item, Len(item) - num)
150                       End If
                          
160                   End If
170               End If
180           Next j
190       Next i
          
          
200       ArrayStringBetweenString = data
210       Exit Function
ArrayStringBetweenString_Error:
220       ArrayStringBetweenString = "#ArrayStringBetweenString:line(" & CStr(Erl) & ")" & Err.Description & "!"

End Function

Public Function StringSplitToArray(sourceString As String, delimiter As String)
10        On Error GoTo StringSplitToArray_Error
          Dim res As Variant

20        res = Split(sourceString, delimiter)

30        StringSplitToArray = To2DArray(res)
40        Exit Function
StringSplitToArray_Error:
50        StringSplitToArray = "#StringSplitToArray:line(" & CStr(Erl) & ")" & Err.Description & "!"

End Function

Public Function ArrayEqual(columnData As Variant, target As Variant)
10        On Error GoTo ArrayEqual_Error
          Dim res As Variant
          Dim i As Long

20        res = ArrayCreate(ArrayRowsCount(columnData), 1, False)
30        For i = 1 To ArrayRowsCount(columnData)
40            If columnData(i, 1) = target Then
50                res(i, 1) = True
60            End If
70        Next i

80        ArrayEqual = res
90        Exit Function
ArrayEqual_Error:
100       ArrayEqual = "#ArrayEqual:line(" & CStr(Erl) & ")" & Err.Description & "!"

End Function

Public Function ArrayIsEqualOrGreaterThan(columnData As Variant, target As Variant)
10        On Error GoTo ArrayIsEqualOrGreaterThan_Error
          Dim res As Variant
          Dim i As Long

20        res = ArrayCreate(UBound(columnData, 1), 1, False)
30        For i = 1 To UBound(columnData, 1)
40            If IsNumeric(columnData(i, 1)) And Not IsEmpty(columnData(i, 1)) Then
50                If columnData(i, 1) >= target Then
60                    res(i, 1) = True
70                End If
80            End If
90        Next i

100       ArrayIsEqualOrGreaterThan = res
110       Exit Function
ArrayIsEqualOrGreaterThan_Error:
120       ArrayIsEqualOrGreaterThan = "#ArrayIsEqualOrGreaterThan:line(" & CStr(Erl) & ")" & Err.Description & "!"

End Function

Public Function ArrayIsEqualOrLessThan(columnData As Variant, target As Variant)
10        On Error GoTo ArrayIsEqualOrLessThan_Error
          Dim res As Variant
          Dim i As Long

20        res = ArrayCreate(UBound(columnData, 1), 1, False)
          
30        For i = 1 To UBound(columnData, 1)
40            If IsNumeric(columnData(i, 1)) And Not IsEmpty(columnData(i, 1)) Then
50                If columnData(i, 1) <= target Then
60                    res(i, 1) = True
70                End If
80            End If
90        Next i

100       ArrayIsEqualOrLessThan = res
110       Exit Function
ArrayIsEqualOrLessThan_Error:
120       ArrayIsEqualOrLessThan = "#ArrayIsEqualOrLessThan:line(" & CStr(Erl) & ")" & Err.Description & "!"

End Function

Public Function ArrayIsLessThan(columnData As Variant, target As Variant)
10        On Error GoTo ArrayIsLessThan_Error
          Dim res As Variant
          Dim i As Long

20        res = ArrayCreate(UBound(columnData, 1), 1, False)
30        For i = 1 To UBound(columnData, 1)
40            If IsNumeric(columnData(i, 1)) And Not IsEmpty(columnData(i, 1)) Then
50                If columnData(i, 1) < target Then
60                    res(i, 1) = True
70                End If
80            End If
90        Next i

100       ArrayIsLessThan = res
110       Exit Function
ArrayIsLessThan_Error:
120       ArrayIsLessThan = "#ArrayIsLessThan:line(" & CStr(Erl) & ")" & Err.Description & "!"

End Function

Public Function ArrayIsGreaterThan(columnData As Variant, target As Variant)
10        On Error GoTo ArrayIsGreaterThan_Error
          Dim res As Variant
          Dim i As Long

20        res = ArrayCreate(ArrayRowsCount(columnData), 1, False)
30        For i = 1 To ArrayRowsCount(columnData)
40            If IsNumeric(columnData(i, 1)) And Not IsEmpty(columnData(i, 1)) Then
50                If columnData(i, 1) > target Then
60                    res(i, 1) = True
70                End If
80            End If
90        Next i

100       ArrayIsGreaterThan = res
110       Exit Function
ArrayIsGreaterThan_Error:
120       ArrayIsGreaterThan = "#ArrayIsGreaterThan:line(" & CStr(Erl) & ")" & Err.Description & "!"

End Function


Public Function ArrayDuplicated(columnData As Variant)
10        On Error GoTo ArrayDuplicated_Error
          Dim resUnique As Variant
          Dim resCount As Variant
          Dim res As Variant
          Dim chooser As Variant

          Dim item As Variant
          Dim currentRow As Long
          Dim IfFound As Boolean
          Dim i As Long
20        currentRow = 2
          
30        resUnique = ArrayCreate(UBound(columnData, 1), 1)
40        resCount = ArrayCreate(UBound(columnData, 1), 1, 1)
50        For Each item In columnData
60            IfFound = False
70            For i = 1 To currentRow - 1
80                If item = resUnique(i, 1) Then
90                    IfFound = True
100                   resCount(i, 1) = resCount(i, 1) + 1
110               End If
120           Next i
130           If IfFound = False Then
140               resUnique(currentRow - 1, 1) = item
150               currentRow = currentRow + 1
160           End If
170       Next item
          
          
180       chooser = ArrayIsGreaterThan(resCount, 1)
190       res = ArrayChooseTrue(resUnique, chooser)

200       ArrayDuplicated = res
210       Exit Function
ArrayDuplicated_Error:
220       ArrayDuplicated = "#ArrayDuplicated:line(" & CStr(Erl) & ")" & Err.Description & "!"

End Function

Public Function ArrayUnique(columnData As Variant)
10        On Error GoTo ArrayUnique_Error
          Dim res As Variant
          Dim item As Variant
          Dim currentRow As Long
          Dim IfFound As Boolean
          Dim i As Long
20        currentRow = 2
          
30        res = ArrayCreate(UBound(columnData, 1), 1)
40        For Each item In columnData
50            IfFound = False
60            For i = 1 To currentRow - 1
70                If item = res(i, 1) Then
80                    IfFound = True
90                End If
100           Next i
110           If IfFound = False Then
120               res(currentRow - 1, 1) = item
130               currentRow = currentRow + 1
140           End If
150       Next item
160       If currentRow = 2 Then
170           res = ArrayCreate(0, 0)
180       Else
190           res = SubArray(res, 1, 1, currentRow - 2)
200       End If
          
210       ArrayUnique = res
220       Exit Function
ArrayUnique_Error:
230       ArrayUnique = "#ArrayUnique:line(" & CStr(Erl) & ")" & Err.Description & "!"

End Function

Public Function ArrayColumnsCount(data As Variant)
10        On Error GoTo ArrayColumnsCount_Error
          Dim res As Long

20        If ArrayDimension(data) <> 2 Then
30            data = To2DArray(data)
40        End If
50        If LBound(data, 2) = 0 And UBound(data, 2) = 0 Then
60            res = 0
70        Else
80            res = UBound(data, 2) - LBound(data, 2) + 1
90        End If

100       ArrayColumnsCount = res
110       Exit Function
ArrayColumnsCount_Error:
120       ArrayColumnsCount = "#ArrayColumnsCount:line(" & CStr(Erl) & ")" & Err.Description & "!"
End Function

Public Function ArrayRowsCount(data As Variant)
10        On Error GoTo ArrayRowsCount_Error
          Dim res As Variant
20        If ArrayDimension(data) <> 2 Then
30            data = To2DArray(data)
40        End If
50        If LBound(data, 1) = 0 And UBound(data, 1) = 0 Then
60            res = 0
70        Else
80            res = UBound(data, 1) - LBound(data, 1) + 1
90        End If
          
100       ArrayRowsCount = res
110       Exit Function
ArrayRowsCount_Error:
120       ArrayRowsCount = "#ArrayRowsCount:line(" & CStr(Erl) & ")" & Err.Description & "!"

End Function

Public Function ArrayDimension(data As Variant)

10        On Error GoTo ArrayDimension_Error
          Dim res As Variant
20        If IsErrorString(ArrayDimension_1(data)) Then
30            ArrayDimension = 0
40        ElseIf IsErrorString(ArrayDimension_2(data)) Then
50            ArrayDimension = 1
60        Else
70            ArrayDimension = 2
80        End If
          
90        Exit Function
ArrayDimension_Error:
100       ArrayDimension = "#ArrayDimension:line(" & CStr(Erl) & ")" & Err.Description & "!"
End Function
Private Function ArrayDimension_1(data As Variant)
10        On Error GoTo ArrayDimension_Error
          Dim n As Long
20        n = UBound(data, 1)
30        Exit Function
ArrayDimension_Error:
40        ArrayDimension_1 = "#ArrayDimension_1:line(" & CStr(Erl) & ")" & Err.Description & "!"

End Function
Private Function ArrayDimension_2(data As Variant)
10        On Error GoTo ArrayDimension_Error
          Dim n As Long
20        n = UBound(data, 2)
30        Exit Function
ArrayDimension_Error:
40        ArrayDimension_2 = "#ArrayDimension_2:line(" & CStr(Erl) & ")" & Err.Description & "!"

End Function


Public Function To2DArray(data As Variant)
10        On Error GoTo To2DArray_Error
          Dim res As Variant
          Dim dimension As Long
          Dim i As Long
20        dimension = ArrayDimension(data)
30        If dimension = 0 Then
40            res = ArrayCreate(1, 1, data)
50        ElseIf dimension = 1 Then
60            res = ArrayCreate(UBound(data, 1) - LBound(data, 1) + 1, 1)
70            For i = LBound(data, 1) To UBound(data, 1)
80                res(i - LBound(data, 1) + 1, 1) = data(i)
90            Next i
100       Else
110           To2DArray = data
120           Exit Function
130       End If

140       To2DArray = res
150       Exit Function
To2DArray_Error:
160       To2DArray = "#To2DArray:line(" & CStr(Erl) & ")" & Err.Description & "!"

End Function

Public Function ArrayRemoveRowAt(data As Variant, rowNum As Long)
10        On Error GoTo ArrayRemoveRowAt_Error
          Dim res As Variant

20        If rowNum = 1 And ArrayRowsCount(data) = 1 Then
30            res = ArrayCreate(0, 0)
40        ElseIf rowNum = 1 Then
50            res = SubArray(data, 2, 1)
60        ElseIf rowNum = ArrayRowsCount(data) Then
70            res = SubArray(data, 1, 1, ArrayRowsCount(data) - 1)
80        Else
90            res = ArrayStack(SubArray(data, 1, 1, rowNum - 1), SubArray(data, rowNum + 1, 1))
100       End If

110       ArrayRemoveRowAt = res
120       Exit Function
ArrayRemoveRowAt_Error:
130       ArrayRemoveRowAt = "#ArrayRemoveRowAt:line(" & CStr(Erl) & ")" & Err.Description & "!"

End Function

Public Function ArrayRemoveColumnAt(data As Variant, columnNum As Long)

10        On Error GoTo ArrayRemoveColumnAt_Error
          Dim res As Variant
20        If columnNum = 1 And ArrayColumnsCount(data) = 1 Then
30            res = ArrayCreate(0, 0)
40        ElseIf columnNum = 1 Then
50            res = SubArray(data, 1, 2)
60        ElseIf columnNum = ArrayColumnsCount(data) Then
70            res = SubArray(data, 1, 1, , UBound(data, 2) - 1)
80        Else
90            res = ArrayRange(SubArray(data, 1, 1, , columnNum - 1), SubArray(data, 1, columnNum + 1))
100       End If
          
          
110       ArrayRemoveColumnAt = res
120       Exit Function
ArrayRemoveColumnAt_Error:
130       ArrayRemoveColumnAt = "#ArrayRemoveColumnAt:line(" & CStr(Erl) & ")" & Err.Description & "!"
End Function

Public Function ArrayChooseTrue(data As Variant, chooser As Variant)

10        On Error GoTo ArrayChooseTrue_Error
          Dim res As Variant
          Dim rowsCount As Long
          Dim i As Long
          Dim j As Long
          Dim currentRow As Long
20        rowsCount = Count(chooser, True)
30        If rowsCount = 0 Then
40            res = ArrayCreate(0, 0)
50        Else
60            res = ArrayCreate(rowsCount, ArrayColumnsCount(data))
70            currentRow = 1
80            For i = 1 To ArrayRowsCount(data)
90                If chooser(i, 1) = True Then
100                   For j = 1 To ArrayColumnsCount(data)
110                       res(currentRow, j) = data(i, j)
120                   Next j
130                   currentRow = currentRow + 1
140               End If
150           Next i
160       End If
          
170       ArrayChooseTrue = res
180       Exit Function
ArrayChooseTrue_Error:
190       ArrayChooseTrue = "#ArrayChooseTrue:line(" & CStr(Erl) & ")" & Err.Description & "!"
End Function


Public Function ArrayCount_New(sourceArray As Variant, targetColumnData As Variant)
10        On Error GoTo ArrayCount_New_Error
          Dim res As Variant
          Dim item As Variant
          Dim i As Long
          Dim targetRowCount As Long

20        targetRowCount = ArrayRowsCount(targetColumnData)
30        res = ArrayCreate(targetRowCount, 1, 0)
40        For Each item In sourceArray
50            For i = 1 To targetRowCount
60                If targetColumnData(i, 1) = item Then
70                    res(i, 1) = res(i, 1) + 1
80                    Exit For
90                End If
100           Next i
110       Next item
         
120       ArrayCount_New = res
130       Exit Function
ArrayCount_New_Error:
140       ArrayCount_New = "#ArrayCount_New:line(" & CStr(Erl) & ")" & Err.Description & "!"

End Function


Public Function ArrayCount_Old(sourceArray As Variant, targetColumnData As Variant)
10        On Error GoTo ArrayCount_Old_Error
          Dim res As Variant
          Dim i As Long

20        res = ArrayCreate(UBound(targetColumnData, 1), 1)
30        For i = 1 To UBound(targetColumnData, 1)
40            res(i, 1) = Count(sourceArray, targetColumnData(i, 1))
50        Next i

60        ArrayCount_Old = res
70        Exit Function
ArrayCount_Old_Error:
80        ArrayCount_Old = "#ArrayCount_Old:line(" & CStr(Erl) & ")" & Err.Description & "!"

End Function

Public Function Count(sourceArray As Variant, target As Variant)

10        On Error GoTo Count_Error
          Dim total As Long
          Dim s As Variant
20        total = 0

30        For Each s In sourceArray
40            If s = target Then
50                total = total + 1
60            End If
70        Next s

80        Count = total
90        Exit Function
Count_Error:
100       Count = "#Count:line(" & CStr(Erl) & ")" & Err.Description & "!"
End Function

Public Function ArrayTranspose(data As Variant)

10        On Error GoTo ArrayTranspose_Error
          Dim res As Variant
          Dim i, j As Long
          Dim rowCount As Long
          Dim columnCount As Long
20        rowCount = ArrayRowsCount(data)
30        columnCount = ArrayColumnsCount(data)
40        res = ArrayCreate(columnCount, rowCount)
50        For i = 1 To rowCount
60            For j = 1 To columnCount
70                res(j, i) = data(i, j)
80            Next j
90        Next i
100       ArrayTranspose = res
110       Exit Function
ArrayTranspose_Error:
120       ArrayTranspose = "#ArrayTranspose:line(" & CStr(Erl) & ")" & Err.Description & "!"
End Function

Public Function ArrayContains(data As Variant, target As Variant)
10        On Error GoTo ArrayContains_Error
          Dim po As New MayClassPosition
20        ArrayFindFirstPosition data, target, po
30        If po.rowNum <> 0 Then
40            ArrayContains = True
50        Else
60            ArrayContains = False
70        End If
80        Exit Function
ArrayContains_Error:
90        ArrayContains = "#ArrayContains:line(" & CStr(Erl) & ")" & Err.Description & "!"

End Function

Public Function ArrayFindFirstPosition(sourceArray As Variant, targetData As Variant, ByRef po As MayClassPosition, Optional byRowOrColumn As EnumByRowOrColumn = ByRow)

10        On Error GoTo ArrayFindFirstPosition_Error
          Dim i As Long
          Dim j As Long
20        po.rowNum = 0
30        po.columnNum = 0
40        If byRowOrColumn = ByRow Then
50            For i = 1 To ArrayRowsCount(sourceArray)
60                For j = 1 To ArrayColumnsCount(sourceArray)
70                    If sourceArray(i, j) = targetData Then
80                        po.rowNum = i
90                        po.columnNum = j
100                       GoTo ArrayFindFirstPosition_end
110                   End If
120               Next j
130           Next i
140       Else
150            For i = 1 To ArrayColumnsCount(sourceArray)
160               For j = 1 To ArrayRowsCount(sourceArray)
170                   If sourceArray(j, i) = targetData Then
180                       po.rowNum = j
190                       po.columnNum = i
200                       GoTo ArrayFindFirstPosition_end
210                   End If
220               Next j
230           Next i
240       End If
ArrayFindFirstPosition_end:
250       Exit Function
ArrayFindFirstPosition_Error:
260       ArrayFindFirstPosition = "#ArrayFindFirstPosition:line(" & CStr(Erl) & ")" & Err.Description & "!"
End Function

Public Function ArrayColumnTakeByTitles(data As Variant, columnsToTake As Variant, IfIncludeTitle As Boolean)
10        On Error GoTo ArrayColumnTakeByTitles_Error
          Dim res As Variant
          Dim i As Long
          
20        For i = 1 To UBound(columnsToTake, 1)
30            If i = 1 Then
40                res = ArrayColumnTakeByTitle(data, CStr(columnsToTake(i, 1)), IfIncludeTitle)
50            Else
60                res = ArrayRange(res, ArrayColumnTakeByTitle(data, CStr(columnsToTake(i, 1)), IfIncludeTitle))
70            End If
80        Next i
          
90        ArrayColumnTakeByTitles = res
100       Exit Function
ArrayColumnTakeByTitles_Error:
110       ArrayColumnTakeByTitles = "#ArrayColumnTakeByTitles:line(" & CStr(Erl) & ")" & Err.Description & "!"

End Function

Public Function ArrayColumnTakeByTitle(data As Variant, columnToTake As String, IfIncludeTitle As Boolean)

10        On Error GoTo ArrayColumnTakeByTitle_Error
          Dim po As New MayClassPosition
20        ArrayFindFirstPosition SubArray(data, 1, 1, 1), columnToTake, po
30        If po.columnNum = 0 Then
40            ArrayColumnTakeByTitle = "#ArrayColumnTakeByTitle:line(" & CStr(Erl) & ")Title can't be found:" & columnToTake & "!"
50            Exit Function
60        Else
70            ArrayColumnTakeByTitle = ArrayColumnTakeByNum(data, po.columnNum, IfIncludeTitle)
80        End If
90        Exit Function
ArrayColumnTakeByTitle_Error:
100       ArrayColumnTakeByTitle = "#ArrayColumnTakeByTitle:line(" & CStr(Erl) & ")" & Err.Description & "!"
End Function

Public Function ArrayColumnTakeByNums(data As Variant, columnsToTake As Variant, IfIncludeTitle As Boolean)
10        On Error GoTo ArrayColumnTakeByNums_Error
          Dim res As Variant
          Dim i As Long
20        For i = 1 To UBound(columnsToTake, 1)
30            If i = 1 Then
40                res = ArrayColumnTakeByNum(data, CLng(columnsToTake(i, 1)), IfIncludeTitle)
50            Else
60                res = ArrayRange(res, ArrayColumnTakeByNum(data, CLng(columnsToTake(i, 1)), IfIncludeTitle))
70            End If
80        Next i

90        ArrayColumnTakeByNums = res
100       Exit Function
ArrayColumnTakeByNums_Error:
110       ArrayColumnTakeByNums = "#ArrayColumnTakeByNums:line(" & CStr(Erl) & ")" & Err.Description & "!"

End Function

Public Function ArrayColumnTakeByNum(data As Variant, columnToTake As Long, IfIncludeTitle As Boolean)

10        On Error GoTo ArrayColumnTakeByNum_Error
          Dim rowStart As Long
          Dim height As Long
          
20        If IfIncludeTitle = True Then
30            rowStart = 1
40            height = ArrayRowsCount(data)
50        Else
60            rowStart = 2
70            height = ArrayRowsCount(data) - 1
80        End If
          
90        ArrayColumnTakeByNum = SubArray(data, rowStart, columnToTake, height, 1)
          
100       Exit Function
ArrayColumnTakeByNum_Error:
110       ArrayColumnTakeByNum = "#ArrayColumnTakeByNum:line(" & CStr(Erl) & ")" & Err.Description & "!"
End Function

Public Function ArrayStack(sourceData As Variant, targetData As Variant)

10        On Error GoTo ArrayStack_Error
          Dim res As Variant
          Dim resSourceData As Variant
          Dim resTargetData As Variant
          Dim sourceColumnCount  As Long
          Dim targetColumnCount As Long
          
20        sourceColumnCount = UBound(sourceData, 2)
30        targetColumnCount = UBound(targetData, 2)
40        If sourceColumnCount >= targetColumnCount Then
50            resSourceData = sourceData
60            resTargetData = ArrayReshape(targetData, UBound(targetData, 1), sourceColumnCount)
70        Else
80            resSourceData = ArrayReshape(sourceData, UBound(sourceData, 1), targetColumnCount)
90            resTargetData = targetData
100       End If
          
          Dim columnCountNew As Long
          Dim rows1 As Long
          Dim rows2 As Long
          Dim i As Long
          Dim j As Long
110       columnCountNew = UBound(resSourceData, 2)
120       rows1 = UBound(resSourceData, 1)
130       rows2 = UBound(resTargetData, 1)
140       res = ArrayCreate(rows1 + rows2, columnCountNew)
150       For i = 1 To rows1 + rows2
160           For j = 1 To columnCountNew
170               If i <= rows1 Then
180                   res(i, j) = resSourceData(i, j)
190               Else
200                   res(i, j) = resTargetData(i - rows1, j)
210               End If
220           Next j
230       Next i
          
240       ArrayStack = res
250       Exit Function
ArrayStack_Error:
260       ArrayStack = "#ArrayStack:line(" & CStr(Erl) & ")" & Err.Description & "!"
End Function

Public Function ArrayRange(sourceData As Variant, targetData As Variant)

10        On Error GoTo ArrayRange_Error
           Dim res As Variant
          Dim resSourceData As Variant
          Dim resTargetData As Variant
          Dim sourceRowCount  As Long
          Dim targetRowCount As Long
          
20        sourceRowCount = ArrayRowsCount(sourceData)
30        targetRowCount = ArrayRowsCount(targetData)
40        If sourceRowCount >= targetRowCount Then
50            resSourceData = sourceData
60            resTargetData = ArrayReshape(targetData, sourceRowCount, ArrayColumnsCount(targetData))
70        Else
80            resSourceData = ArrayReshape(sourceData, targetRowCount, ArrayColumnsCount(sourceData))
90            resTargetData = targetData
100       End If
          
          Dim rowCountNew As Long
          Dim columns1 As Long
          Dim columns2 As Long
          Dim i As Long
          Dim j As Long
110       rowCountNew = ArrayRowsCount(resSourceData)
120       columns1 = ArrayColumnsCount(resSourceData)
130       columns2 = ArrayColumnsCount(resTargetData)
140       res = ArrayCreate(rowCountNew, columns1 + columns2)
150       For i = 1 To rowCountNew
160           For j = 1 To columns1 + columns2
170               If j <= columns1 Then
180                   res(i, j) = resSourceData(i, j)
190               Else
200                   res(i, j) = resTargetData(i, j - columns1)
210               End If
220           Next j
230       Next i
          
240       ArrayRange = res
250       Exit Function
ArrayRange_Error:
260       ArrayRange = "#ArrayRange:line(" & CStr(Erl) & ")" & Err.Description & "!"
End Function


Public Function SubArray(data As Variant, rowStart As Long, columnStart As Long, Optional height As Long = 0, Optional width As Long = 0)
10        On Error GoTo SubArray_Error
          Dim i As Long
          Dim j As Long
          Dim m As Long
          Dim n As Long

20        If height = 0 Then
30            height = UBound(data, 1) - rowStart + 1
40        End If
          
50        If width = 0 Then
60            width = UBound(data, 2) - columnStart + 1
70        End If

          Dim res() As Variant
80        ReDim res(1 To height, 1 To width)
90        For i = rowStart To rowStart + height - 1
100           For j = columnStart To columnStart + width - 1
110               res(i - rowStart + 1, j - columnStart + 1) = data(i, j)
120           Next j
130       Next i
140       SubArray = res
150       Exit Function
SubArray_Error:
160       SubArray = "#SubArray:line(" & CStr(Erl) & ")" & Err.Description & "!"
End Function

Public Function ArrayReshape(data As Variant, height As Long, width As Long)

10        On Error GoTo ArrayReshape_Error
          Dim oldRowCOunt, oldColumnCount As Long
          Dim i, j As Long
          Dim res As Variant
20        oldRowCOunt = ArrayRowsCount(data)
30        oldColumnCount = ArrayColumnsCount(data)
          
40        res = ArrayCreate(height, width)
          
50        For i = 1 To oldRowCOunt
60            If i > height Then
70                Exit For
80            End If
90            For j = 1 To oldColumnCount
100               If j > width Then
110                   Exit For
120               End If
130               res(i, j) = data(i, j)
140           Next j
150       Next i

160       ArrayReshape = res
170       Exit Function
ArrayReshape_Error:
180       ArrayReshape = "#ArrayReshape:line(" & CStr(Erl) & ")" & Err.Description & "!"
End Function

Public Function ArrayCreate(rowCount As Long, columnCount As Long, Optional data As Variant = Empty)

10        On Error GoTo ArrayCreate_Error
          Dim res() As Variant
20        If rowCount = 0 Or columnCount = 0 Then
30            ReDim res(0, 0)
40            GoTo ArrayCreate_end
50        End If
60        ReDim res(1 To rowCount, 1 To columnCount)
70        If Not IsEmpty(data) Then
              Dim i As Long
              Dim j As Long
80            For i = 1 To rowCount
90                For j = 1 To columnCount
100                   res(i, j) = data
110               Next j
120           Next i
130       End If
ArrayCreate_end:
140       ArrayCreate = res
150       Exit Function
ArrayCreate_Error:
160       ArrayCreate = "#ArrayCreate:line(" & CStr(Erl) & ")" & Err.Description & "!"
End Function

Public Function ArrayNumberCreate(rowCount As Long, columnCount As Long, Optional startNum As Long = 1, Optional step As Long = 1, Optional byRowOrColumn As EnumByRowOrColumn = ByColumn)

10        On Error GoTo ArrayNumberCreate_Error
          Dim res() As Variant
          Dim currentNum As Long
          Dim i As Long
          Dim j As Long
20        If rowCount = 0 Or columnCount = 0 Then
30            ReDim res(0, 0)
40            GoTo ArrayNumberCreate_end
50        End If
60        ReDim res(1 To rowCount, 1 To columnCount)
70        currentNum = startNum
80        If byRowOrColumn = ByRow Then
90            For i = 1 To rowCount
100               For j = 1 To columnCount
110                   res(i, j) = currentNum
120                   currentNum = currentNum + step
130               Next j
140           Next i
150       Else
160           For i = 1 To columnCount
170               For j = 1 To rowCount
180                   res(j, i) = currentNum
190                   currentNum = currentNum + step
200               Next j
210           Next i
220       End If
ArrayNumberCreate_end:
230       ArrayNumberCreate = res
240       Exit Function
ArrayNumberCreate_Error:
250       ArrayNumberCreate = "#ArrayNumberCreate:line(" & CStr(Erl) & ")" & Err.Description & "!"
End Function

Public Function IsErrorString(data As Variant)
10    On Error Resume Next
20    IsErrorString = False
30    If TypeName(data) = "String" Then
40        If InStr(data, "#") = 1 And InStr(data, "!") = Len(data) Then
50            IsErrorString = True
60        End If
70    End If
End Function

Public Function IsEmptyOrEmptyString(data As Variant)
10        On Error GoTo IsEmptyOrEmptyString_Error
          Dim res As Variant
20        res = False
30        If IsEmpty(data) Then
40            res = True
50        ElseIf TypeName(data) = "String" Then
60            If Len(data) = 0 Then
70                res = True
80            End If
90        End If
          
100       IsEmptyOrEmptyString = res
110       Exit Function
IsEmptyOrEmptyString_Error:
120       IsEmptyOrEmptyString = "#IsEmptyOrEmptyString:line(" & CStr(Erl) & ")" & Err.Description & "!"

End Function

Public Function IsEmptyOrErrorString(data As Variant)
10        On Error GoTo isEmptyOrErrorString_Error
          Dim res As Variant
          
20        If IsEmptyOrEmptyString(data) Or IsErrorString(data) Then
30            res = True
40        Else
50            res = False
60        End If
          
70        IsEmptyOrErrorString = res
80        Exit Function
isEmptyOrErrorString_Error:
90        IsEmptyOrErrorString = "#isEmptyOrErrorString:line(" & CStr(Erl) & ")" & Err.Description & "!"

End Function

Public Function ShowData(data As Variant)

10        On Error GoTo ShowData_Error
          Dim res As Variant
          Dim wbk As Workbook
          Dim sh As Worksheet
          Dim num_row As Long
          Dim num_column As Long
20        Set wbk = Workbooks.Add
30        Set sh = wbk.Sheets(1)
40        res = To2DArray(data)
50        sh.Range(sh.Cells(1, 1), sh.Cells(ArrayRowsCount(res), ArrayColumnsCount(res))) = res
          
60        wbk.Activate
70        Exit Function
ShowData_Error:
80        ShowData = "#ShowData:line(" & CStr(Erl) & ")" & Err.Description & "!"
End Function

Public Function GetNumberOfLastRow(ByRef sh As Worksheet)

          Dim i As Long
          Dim numLastRowOfWholeRange As Long
          Dim numLastRow As Long
10       On Error GoTo GetNumberOfLastRow_Error
20        numLastRow = 1
30        numLastRowOfWholeRange = sh.Cells.Rows.Count

40        For i = 1 To sh.UsedRange.Columns.Count + sh.UsedRange.Rows.Column - 1
              Dim numLastRowOfCurrentColumn As Long
50            numLastRowOfCurrentColumn = sh.Cells(numLastRowOfWholeRange, i).End(xlUp).Row
60            If numLastRowOfCurrentColumn > numLastRow Then
70                numLastRow = numLastRowOfCurrentColumn
80            End If
90        Next i
100       GetNumberOfLastRow = numLastRow
110      Exit Function

GetNumberOfLastRow_Error:
120       GetNumberOfLastRow = "#GetNumberOfLastRow:line(" & CStr(Erl) & ")" & Err.Description & "!"
End Function

Public Function GetNumberOfLastColumn(ByRef sh As Worksheet)
          Dim i As Long
          Dim numLastColumnOfWholeRange As Long
          Dim numLastColumn As Long
          
10       On Error GoTo GetNumberOfLastColumn_Error
          
20        numLastColumn = 1
30        numLastColumnOfWholeRange = sh.Cells.Columns.Count
         
40        For i = 1 To sh.UsedRange.Rows.Count + sh.UsedRange.Rows.Row - 1
              Dim numLastColumnOfCurrentColumn As Long
50            numLastColumnOfCurrentColumn = sh.Cells(i, numLastColumnOfWholeRange).End(xlToLeft).Column
60            If numLastColumnOfCurrentColumn > numLastColumn Then
70                numLastColumn = numLastColumnOfCurrentColumn
80            End If
90        Next i
100       GetNumberOfLastColumn = numLastColumn

110      Exit Function

GetNumberOfLastColumn_Error:
120       GetNumberOfLastColumn = "#GetNumberOfLastColumn:line(" & CStr(Erl) & ")" & Err.Description & "!"
End Function

Public Function GetUsedRange(ByRef sh As Worksheet, Optional IfFromFirstCell As Boolean = False) As Range

          Dim j As Long
          
10       On Error GoTo GetUsedRange_Error
          
          Dim numLastRow As Long
          Dim numLastColumn As Long
20        numLastRow = GetNumberOfLastRow(sh)
30        numLastColumn = GetNumberOfLastColumn(sh)
40        If IfFromFirstCell = True Then
50            Set GetUsedRange = sh.Range(sh.Cells(1, 1), sh.Cells(numLastRow, numLastColumn))
60        Else
70            Set GetUsedRange = sh.Range(sh.Cells(sh.UsedRange.Rows.Row, sh.UsedRange.Rows.Column), sh.Cells(numLastRow, numLastColumn))
80        End If
90       Exit Function

GetUsedRange_Error:
100       GetUsedRange = "#GetUsedRange:line(" & CStr(Erl) & ")" & Err.Description & "!"
End Function

Public Function wbkReopen(wbkPath As String) As Workbook
10        On Error Resume Next
          
          Dim tmpWbk As Workbook
          Dim wbkName As String
          
20        wbkName = GetName(wbkPath)
          
30        Set tmpWbk = Workbooks(wbkName)
          
40        If tmpWbk Is Nothing Then
              'Not opened, so Open it
50            Err.Clear
60            Workbooks.Open fileName:=wbkPath
70            If Err.Number <> 0 Then
80                wbkReopen = Nothing
90                Exit Function
100           End If
110           Set wbkReopen = ActiveWorkbook
120       Else
              'Opened, so return the handle of the workbook
130           Set wbkReopen = tmpWbk
140       End If
End Function
Public Function GetFolder(pathText As String) As String
10    On Error Resume Next
          Dim fileName As String
20        fileName = GetName(pathText)
30        GetFolder = Left(pathText, Len(pathText) - Len(fileName) - 1)
End Function
Public Function GetShortName(name As String) As String
10    On Error Resume Next
          Dim tmpList As Variant
20        tmpList = Split(name, ".")
30        GetShortName = tmpList(LBound(tmpList))
End Function
Public Function GetSuffix(name As String) As String
10    On Error Resume Next
          Dim tmpList As Variant
20        tmpList = Split(name, ".")
30        GetSuffix = tmpList(UBound(tmpList))
End Function
Public Function GetName(pathTxt) As String
10    On Error Resume Next
          Dim tmpList As Variant
20        tmpList = Split(pathTxt, "\")
30        GetName = tmpList(UBound(tmpList))
End Function
Public Function GetFileListInFolderStep1(folderPath As String)
10        On Error GoTo getFileListInFolderStep1_Error
          Dim res As Variant
          Dim fso As New FileSystemObject
          Dim fldr As Folder
          Dim MyFiles As Files
          Dim File As File
          Dim MyFolders As Folders
          Dim Folder As Folder
          Dim currentRow As Long
20        Set fldr = fso.GetFolder(folderPath)
30        Set MyFiles = fldr.Files
40        Set MyFolders = fldr.SubFolders
50        currentRow = 0
60        If MyFiles.Count > 0 Then
70            res = ArrayCreate(MyFiles.Count, 1)
80            For Each File In MyFiles
90                res(currentRow + 1, 1) = File.Path
100               currentRow = currentRow + 1
110           Next
120       Else
130           res = ArrayCreate(0, 0)
140       End If
150       GetFileListInFolderStep1 = res
160       Exit Function
getFileListInFolderStep1_Error:
170       GetFileListInFolderStep1 = "#getFileListInFolderStep1:line(" & CStr(Erl) & ")" & Err.Description & "!"

End Function

Public Function GetFileListInFolderAndSub(folderPath As String)
10        On Error GoTo getFileListInFolderAndSub_Error
          Dim res As Variant
          Dim fso As New FileSystemObject
          Dim fldr As Folder
          Dim MyFiles As Files
          Dim File As File
          Dim MyFolders As Folders
          Dim Folder As Folder
          Dim currentRow As Long
20        Set fldr = fso.GetFolder(folderPath)
30        Set MyFiles = fldr.Files
40        Set MyFolders = fldr.SubFolders
50        currentRow = 0
60        If MyFiles.Count > 0 Then
70            res = ArrayCreate(MyFiles.Count, 1)
80            For Each File In MyFiles
90                res(currentRow + 1, 1) = File.Path
100               currentRow = currentRow + 1
110           Next
120       Else
130           res = ArrayCreate(0, 0)
140       End If
150       If MyFolders.Count > 0 Then
160           For Each Folder In MyFolders
170              res = ArrayStack(res, GetFileListInFolderAndSub(Folder.Path))
180           Next Folder
190       End If
          

200       GetFileListInFolderAndSub = res
210       Exit Function
getFileListInFolderAndSub_Error:
220       GetFileListInFolderAndSub = "#getFileListInFolderAndSub:line(" & CStr(Erl) & ")" & Err.Description & "!"

End Function

Public Function GetFolderListInFolderStep1(folderPath As String)
10        On Error GoTo getFolderListInFolderStep1_Error
          Dim res As Variant
          Dim fso As New FileSystemObject
          Dim fldr As Folder
          Dim MyFiles As Files
          Dim File As File
          Dim MyFolders As Folders
          Dim Folder As Folder
          Dim currentRow As Long
20        Set fldr = fso.GetFolder(folderPath)
30        Set MyFiles = fldr.Files
40        Set MyFolders = fldr.SubFolders
50        currentRow = 0
60        If MyFolders.Count > 0 Then
70            res = ArrayCreate(MyFolders.Count, 1)
80            For Each Folder In MyFolders
90                res(currentRow + 1, 1) = Folder.Path
100               currentRow = currentRow + 1
110           Next
120       Else
130           res = ArrayCreate(0, 0)
140       End If
          
150       GetFolderListInFolderStep1 = res
160       Exit Function
getFolderListInFolderStep1_Error:
170       GetFolderListInFolderStep1 = "#getFolderListInFolderStep1:line(" & CStr(Erl) & ")" & Err.Description & "!"

End Function

Public Function GetFolderListInFolderAndSub(folderPath As String)
10        On Error GoTo getFolderListInFolderAndSub_Error
          Dim res As Variant
          Dim fso As New FileSystemObject
          Dim fldr As Folder
          Dim MyFiles As Files
          Dim File As File
          Dim MyFolders As Folders
          Dim Folder As Folder
          Dim currentRow As Long
20        Set fldr = fso.GetFolder(folderPath)
30        Set MyFiles = fldr.Files
40        Set MyFolders = fldr.SubFolders
50        currentRow = 0
60        If MyFolders.Count > 0 Then
70            res = ArrayCreate(MyFolders.Count, 1)
80            For Each Folder In MyFolders
90                res(currentRow + 1, 1) = Folder.Path
100               currentRow = currentRow + 1
110           Next
120           For Each Folder In MyFolders
130               res = ArrayStack(res, GetFolderListInFolderAndSub(Folder.Path))
140           Next
150       Else
160           res = ArrayCreate(0, 0)
170       End If
          

180       GetFolderListInFolderAndSub = res
190       Exit Function
getFolderListInFolderAndSub_Error:
200       GetFolderListInFolderAndSub = "#getFolderListInFolderAndSub:line(" & CStr(Erl) & ")" & Err.Description & "!"

End Function

Public Function CopyOnlyFormats(sourceRange As Range, targetRange As Range)
10        On Error GoTo CopyOnlyFormats_Error
          
20        sourceRange.Copy
30        targetRange.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
40        Application.CutCopyMode = False
          
50        Exit Function
CopyOnlyFormats_Error:
60        CopyOnlyFormats = "#CopyOnlyFormats:line(" & CStr(Erl) & ")" & Err.Description & "!"

End Function

Public Function AboutMay()
10    AboutMay = "Author: shezhidian@163.com"
End Function
