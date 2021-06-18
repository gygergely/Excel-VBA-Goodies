Attribute VB_Name = "Module1"
Option Explicit

Sub test_comparison()

    Dim condition_list() As Variant
    Dim compare_tbl As ListObject
    Dim column_list() As Variant
    Dim rng As Range
    Dim i As Long, x As Long

    'assign value to table related variables - tbl the table itself | rng the table range
    Set compare_tbl = sh_test.ListObjects("compare_tabl")
    Set rng = Range("compare_tabl")

    'store the values to be compared in an array
    condition_list = Range("comp_condition")

    'define which columns should be compared to the condition list
    ReDim column_list(1 To 3)
    column_list(1) = 1
    column_list(2) = 2
    column_list(3) = 3

    ' iterate through all table rows
    For i = 1 To rng.Rows.Count
        
        'perform the comparison
        If f_comparing_results(compare_tbl, condition_list, column_list, i) Then
            
            'if all good color the relevant table cells
            For x = LBound(column_list) To UBound(column_list)
                compare_tbl.DataBodyRange(i, x).Interior.Color = RGB(100, 150, 200)
            Next x
        
        End If
        
    Next i

End Sub


Function f_comparing_results(tbl As ListObject, conditions_list As Variant, column_list As Variant, row_to_check As Long) As Boolean

    Dim i As Long

    f_comparing_results = True

    For i = LBound(column_list) To UBound(column_list)
        If conditions_list(i, 1) <> tbl.DataBodyRange(row_to_check, i) Then
            f_comparing_results = False
            Exit For
        End If
    Next i

End Function

Sub clear_table_format()

    Dim compare_tbl As ListObject

    Set compare_tbl = sh_test.ListObjects("compare_tabl")

    compare_tbl.DataBodyRange.ClearFormats

End Sub
