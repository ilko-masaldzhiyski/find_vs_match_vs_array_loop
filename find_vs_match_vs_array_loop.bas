Option Explicit

Sub compare_search()
    Dim rng_test_data As Range
    Dim int_trial_counter As Integer
    Dim arr_results(1 To 1, 1 To 5) As Variant, arr_thresholds As Variant, _
        arr_test_data As Variant, var_threshold As Variant

    arr_thresholds = Array(0.99, 0.9, 0.85, 0.8, 0.75, 0.7, 0.65, 0.6, 0.55, 0.5, 0.45, 0.4, 0.35, 0.3, _
                           0.25, 0.2, 0.15, 0.1, 0.05, 0.001)
    With ActiveWorkbook.Sheets(2)
        .Cells(1, 1).Resize(1, UBound(arr_results, 2)).Value2 = Array("Threshold", "Find", "Match", "Array", "Elements")
    End With
    For Each var_threshold In arr_thresholds
        Call generate_test_data(var_threshold)

        Set rng_test_data = ActiveWorkbook.Sheets(1).UsedRange
        arr_test_data = rng_test_data.Value2
        For int_trial_counter = LBound(arr_results, 1) To UBound(arr_results, 1)
            arr_results(int_trial_counter, 1) = var_threshold
            arr_results(int_trial_counter, 2) = time_find(rng_test_data)
            arr_results(int_trial_counter, 3) = time_match(rng_test_data)
            arr_results(int_trial_counter, 4) = time_array(arr_test_data)
            arr_results(int_trial_counter, 5) = time_array_items(arr_test_data)
        Next int_trial_counter
        With ActiveWorkbook.Sheets(2)
            .Cells(.UsedRange.Rows.Count + 1, 1).Resize(UBound(arr_results, 1), UBound(arr_results, 2)).Value2 = arr_results
        End With
    Next var_threshold
End Sub

Function generate_test_data(ByVal dbl_threshold As Double)
    Dim arr_test_data(1 To 100000, 1 To 2) As Variant
    Dim lng_counter As Long
    Rnd -1652 'Random seed

    For lng_counter = LBound(arr_test_data) To UBound(arr_test_data, 1)
        If Rnd > dbl_threshold Then arr_test_data(lng_counter, 1) = "foo"
        If Rnd > dbl_threshold Then arr_test_data(lng_counter, 2) = "bar"
    Next lng_counter
    With ActiveWorkbook.Sheets(1)
        .UsedRange.ClearContents
        .Range("A1").Resize(UBound(arr_test_data, 1), UBound(arr_test_data, 2)).Value2 = arr_test_data
    End With
End Function

Function time_find(ByVal rng_test_data As Range) As Double
    Dim lng_result_array As Long
    Dim dbl_start_time As Double, dbl_end_time As Double
    Dim rng_lookup_column As Range, rng_found_result As Range
    Dim str_first_address As String

    dbl_start_time = Timer
    Set rng_lookup_column = rng_test_data.Resize(rng_test_data.Rows.Count, 1)
    With rng_lookup_column
        Set rng_found_result = .Find("foo", After:=.Cells(.Rows.Count, .Columns.Count), _
                                     LookIn:=xlValues, SearchDirection:=xlNext, MatchCase:=False)
        str_first_address = rng_found_result.Address
        Do
            Set rng_found_result = .FindNext(rng_found_result)
            If rng_found_result.Offset(0, 1) = "bar" Then
                lng_result_array = lng_result_array + 1
            End If
        Loop While Not rng_found_result Is Nothing And rng_found_result.Address <> str_first_address
    End With
    dbl_end_time = Timer
    time_find = dbl_end_time - dbl_start_time
End Function

Function time_match(ByVal rng_test_data As Range) As Double
    Dim lng_match_position As Long, lng_result_array As Long
    Dim dbl_start_time As Double, dbl_end_time As Double
    Dim rng_lookup_column As Range

    dbl_start_time = Timer
    Set rng_lookup_column = rng_test_data.Resize(rng_test_data.Rows.Count, 1)
    On Error GoTo Finish
    Do
        lng_match_position = Application.Match("foo", rng_lookup_column, False)
        If rng_lookup_column(lng_match_position, 2) = "bar" Then
            lng_result_array = lng_result_array + 1
        End If
        Set rng_lookup_column = rng_lookup_column.Resize(rng_lookup_column.Rows.Count - lng_match_position, 1).Offset(lng_match_position, 0)
    Loop
Finish:
    dbl_end_time = Timer
    time_match = dbl_end_time - dbl_start_time
End Function

Function time_array(ByVal arr_test_data As Variant) As Double
    Dim lng_array_counter As Long, lng_result_array As Long
    Dim dbl_start_time As Double, dbl_end_time As Double

    dbl_start_time = Timer
    For lng_array_counter = LBound(arr_test_data, 1) To UBound(arr_test_data, 1)
        If arr_test_data(lng_array_counter, 1) = "foo" And _
           arr_test_data(lng_array_counter, 2) = "bar" Then
            lng_result_array = lng_result_array + 1
        End If
    Next lng_array_counter
    dbl_end_time = Timer
    time_array = dbl_end_time - dbl_start_time
End Function

Function time_array_items(ByVal arr_test_data As Variant) As Long
    Dim lng_array_counter As Long, lng_result_array As Long

    For lng_array_counter = LBound(arr_test_data, 1) To UBound(arr_test_data, 1)
        If arr_test_data(lng_array_counter, 1) = "foo" And _
           arr_test_data(lng_array_counter, 2) = "bar" Then
            lng_result_array = lng_result_array + 1
        End If
    Next lng_array_counter
    time_array_items = lng_result_array
End Function
