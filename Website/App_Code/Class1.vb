Imports Microsoft.VisualBasic

Public Class Class1

    Public Shared Function theFix(ByVal theVariable)
        theFix = Replace(theVariable, "'", "''")
        ' this replaces a single apostrophe ( ' ) with two single apostrophes ( ''). Note: two single apostrophes, not one quotation mark
    End Function

    Public Shared Function theDateFix(ByVal theVariable)
        theDateFix = Replace(theVariable, "'", "")
        ' this replaces a single apostrophe ( ' ) with nothing . 
    End Function

    Public Shared Function trapnulls(ByVal theVariable)
        If IsDBNull(theVariable) = False Then
            trapnulls = theVariable
        Else
            trapnulls = 0
        End If

        ' this replaces a single apostrophe ( ' ) with two single apostrophes ( ''). Note: two single apostrophes, not one quotation mark
    End Function

    Public Shared Function FindBinLocation(ByRef TheArray As Object) As String
        'This function gives max value of int array without sorting an array
        Dim i As Integer
        Dim MaxIntegersIndex As Integer

        Dim MyArrayMaxFits(6) As Object
        Dim MaxFit As Integer
        'Dim MyArrayBinChoices(1, 0) As Object
        Dim MyArrayBinChoices(1, UBound(TheArray)) As Object
        Dim c As Integer
        c = 0




        MaxIntegersIndex = 0

        For i = 0 To UBound(TheArray)
            'If TheArray(i) > TheArray(MaxIntegersIndex) Then
            '    MaxIntegersIndex = i
            'End If
            MyArrayMaxFits(1) = Math.Floor(TheArray(1, i) / TheArray(6, i)) * Math.Floor(TheArray(2, i) / TheArray(5, i)) * Math.Floor(TheArray(3, i) / TheArray(4, i))
            MyArrayMaxFits(2) = Math.Floor(TheArray(1, i) / TheArray(6, i)) * Math.Floor(TheArray(2, i) / TheArray(4, i)) * Math.Floor(TheArray(3, i) / TheArray(5, i))
            MyArrayMaxFits(3) = Math.Floor(TheArray(2, i) / TheArray(6, i)) * Math.Floor(TheArray(1, i) / TheArray(5, i)) * Math.Floor(TheArray(3, i) / TheArray(4, i))
            MyArrayMaxFits(4) = Math.Floor(TheArray(2, i) / TheArray(6, i)) * Math.Floor(TheArray(3, i) / TheArray(5, i)) * Math.Floor(TheArray(1, i) / TheArray(4, i))
            MyArrayMaxFits(5) = Math.Floor(TheArray(3, i) / TheArray(6, i)) * Math.Floor(TheArray(1, i) / TheArray(5, i)) * Math.Floor(TheArray(2, i) / TheArray(4, i))
            MyArrayMaxFits(6) = Math.Floor(TheArray(3, i) / TheArray(6, i)) * Math.Floor(TheArray(2, i) / TheArray(5, i)) * Math.Floor(TheArray(1, i) / TheArray(4, i))
            MaxFit = MaxValOfIntArray(MyArrayMaxFits)

            'If MaxFit > 0 Then
            ' If MaxFit > TheArray(7, i) Then
            ' ReDim Preserve MyArrayBinChoices(1, c)
            MyArrayBinChoices(0, c) = TheArray(0, i)
            MyArrayBinChoices(1, c) = Int(MaxFit) - Int(TheArray(7, i))
            c = c + 1
            '  End If
            '  End If
        Next

        'Sort the array
        Dim p
        'Dim k
        Dim j
        Dim s
        Dim m
        'Dim MyArray(1, c - 1) As Object

        'For p = 0 To c - 1 Step 1
        '    For j = 0 To c - 2 Step 1
        '        If MyArrayBinChoices(1, j) < MyArrayBinChoices(1, j + 1) Then
        '            For m = 0 To 1
        '                s = MyArrayBinChoices(m, j + 1)
        '                MyArrayBinChoices(m, j + 1) = MyArrayBinChoices(m, j)
        '                MyArrayBinChoices(m, j) = s
        '            Next
        '        End If
        '    Next

        '    '  c = 0

        'Next

        'Sort the array
        For p = c - 1 To 0 Step -1
            For j = c - 2 To 0 Step -1
                If MyArrayBinChoices(1, j) < MyArrayBinChoices(1, j + 1) Then
                    For m = 0 To 1
                        s = MyArrayBinChoices(m, j + 1)
                        MyArrayBinChoices(m, j + 1) = MyArrayBinChoices(m, j)
                        MyArrayBinChoices(m, j) = s
                    Next
                End If
            Next
        Next


        'index of max value is MaxValOfIntArray
        'FindBinLocation = MyArrayBinChoices(0, 0)
        FindBinLocation = MyArrayBinChoices(0, c - 1)
        'FindBinLocation = TheArray(7, 0)
    End Function

    Public Shared Function MaxValOfIntArray(ByRef TheArray As Object) As Decimal
        'This function gives max value of int array without sorting an array
        Dim i As Decimal
        Dim MaxIntegersIndex As Decimal
        MaxIntegersIndex = 0

        For i = 1 To UBound(TheArray)
            If TheArray(i) > TheArray(MaxIntegersIndex) Then
                MaxIntegersIndex = i
            End If
        Next
        'index of max value is MaxValOfIntArray
        MaxValOfIntArray = TheArray(MaxIntegersIndex)
    End Function




End Class
