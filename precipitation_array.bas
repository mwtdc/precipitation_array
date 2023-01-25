Sub precip_array()

't = Timer

'ExcelEvents (False)
'Application.Calculation = xlManual

Dim sh_precip, sh_frcst As Worksheet
Dim arr, arr2, arr3, arr4(1 To 6, 1 To 24) As Variant
Dim i, j, x, k, n, p As Integer

Set sh_precip = ThisWorkbook.Sheets("Прогноз погоды")
Set sh_frcst = ThisWorkbook.Sheets("Прогнозирование")

search_date = CDate(sh_frcst.Range("O1").value)
search_locality = sh_frcst.Range("T1").value

Row_beg_locality = sh_precip.Range("A:A").Find(What:=search_locality, After:=sh_precip.Range("A1"), SearchOrder:=xlByRows, SearchDirection:=xlNext, LookAt:=xlWhole).row
Row_end_locality = sh_precip.Range("A:A").Find(What:=search_locality, After:=sh_precip.Range("A1"), SearchOrder:=xlByRows, SearchDirection:=xlPrevious, LookAt:=xlWhole).row
Row_beg_date = sh_precip.Range("E:E").Find(What:=search_date, After:=sh_precip.Range("E" & Row_beg_locality - 1), SearchOrder:=xlByRows, SearchDirection:=xlNext, LookAt:=xlWhole).row
Row_end_date = sh_precip.Range("E:E").Find(What:=search_date, After:=sh_precip.Range("E" & Row_end_locality + 1), SearchOrder:=xlByRows, SearchDirection:=xlPrevious, LookAt:=xlWhole).row

arr = sh_precip.Range("I" & Row_beg_date & ":K" & Row_end_date).value
arr2 = sh_precip.Range("G" & Row_beg_date & ":G" & Row_end_date).value
arr3 = sh_precip.Range("L" & Row_beg_date & ":L" & Row_end_date).value

If search_date = Date + 1 Then

    If UBound(arr) = 18 Then
    
    'Заполняем c 0 по 15 часы
    For k = 1 To 16
        For n = 1 To 3
        p = 4 - n
        If arr(k, p) = "" Then arr(k, p) = 0
        arr4(n, k) = arr(k, p)
        Next n
    arr4(4, k) = WorksheetFunction.Round((1.7 * arr4(1, k) + 0.8 * arr4(2, k) + 0.5 * arr4(3, k)) / 3, 0)
    If arr4(4, k) > 100 Then arr4(4, k) = 100
    arr4(5, k) = arr2(k, 1)
    arr4(6, k) = arr3(k, 1)
    Next k
    'Заполняем 18,21 часы
    j = 16
    For k = 19 To 23 Step 3
    j = j + 1
        For n = 1 To 3
        p = 4 - n
        If arr(j, p) = "" Then arr(j, p) = 0
        arr4(n, k) = arr(j, p)
        Next n
    arr4(4, k) = WorksheetFunction.Round((1.7 * arr4(1, k) + 0.8 * arr4(2, k) + 0.5 * arr4(3, k)) / 3, 0)
    If arr4(4, k) > 100 Then arr4(4, k) = 100
    arr4(5, k) = arr2(j, 1)
    arr4(6, k) = arr3(j, 1)
    Next k
    'Заполняем 16,19 часы
    For k = 17 To 20 Step 3
        arr4(1, k) = WorksheetFunction.Round((2 * arr4(1, k - 1) + arr4(1, k + 2)) / 3, 0)
        arr4(2, k) = WorksheetFunction.Round((2 * arr4(2, k - 1) + arr4(2, k + 2)) / 3, 0)
        arr4(3, k) = WorksheetFunction.Round((2 * arr4(3, k - 1) + arr4(3, k + 2)) / 3, 0)
        arr4(4, k) = WorksheetFunction.Round((2 * arr4(4, k - 1) + arr4(4, k + 2)) / 3, 0)
        arr4(5, k) = arr4(5, k - 1)
        arr4(6, k) = arr4(6, k - 1)
    Next k
    'Заполняем 17,20 часы
    For k = 18 To 21 Step 3
        arr4(1, k) = WorksheetFunction.Round((arr4(1, k - 2) + 2 * arr4(1, k + 1)) / 3, 0)
        arr4(2, k) = WorksheetFunction.Round((arr4(2, k - 2) + 2 * arr4(2, k + 1)) / 3, 0)
        arr4(3, k) = WorksheetFunction.Round((arr4(3, k - 2) + 2 * arr4(3, k + 1)) / 3, 0)
        arr4(4, k) = WorksheetFunction.Round((arr4(4, k - 2) + 2 * arr4(4, k + 1)) / 3, 0)
        arr4(5, k) = arr4(5, k - 1)
        arr4(6, k) = arr4(6, k - 1)
    Next k
    'Заполняем 22,23 часы
    For k = 23 To 24
        arr4(1, k) = arr4(1, 22)
        arr4(2, k) = arr4(2, 22)
        arr4(3, k) = arr4(3, 22)
        arr4(4, k) = arr4(4, 22)
        arr4(5, k) = arr4(5, k - 1)
        arr4(6, k) = arr4(6, k - 1)
    Next k
    
    End If
    
    If UBound(arr) = 24 Then
    
    For k = 1 To 24
        For n = 1 To 3
        p = 4 - n
        If arr(k, p) = "" Then arr(k, p) = 0
        arr4(n, k) = arr(k, p)
        Next n
    arr4(4, k) = WorksheetFunction.Round((1.7 * arr4(1, k) + 0.8 * arr4(2, k) + 0.5 * arr4(3, k)) / 3, 0)
    If arr4(4, k) > 100 Then arr4(4, k) = 100
    arr4(5, k) = arr2(k, 1)
    arr4(6, k) = arr3(k, 1)
    Next k
    
    End If
    
    If UBound(arr) = 22 Then
    
    For k = 1 To 22
        For n = 1 To 3
        p = 4 - n
        If arr(k, p) = "" Then arr(k, p) = 0
        arr4(n, k) = arr(k, p)
        Next n
    arr4(4, k) = WorksheetFunction.Round((1.7 * arr4(1, k) + 0.8 * arr4(2, k) + 0.5 * arr4(3, k)) / 3, 0)
    If arr4(4, k) > 100 Then arr4(4, k) = 100
    arr4(5, k) = arr2(k, 1)
    arr4(6, k) = arr3(k, 1)
    Next k
    
    'Заполняем 22,23 часы
    For k = 23 To 24
        arr4(1, k) = arr4(1, 22)
        arr4(2, k) = arr4(2, 22)
        arr4(3, k) = arr4(3, 22)
        arr4(4, k) = arr4(4, 22)
        arr4(5, k) = arr4(5, k - 1)
        arr4(6, k) = arr4(6, k - 1)
    Next k
    
    End If
    
End If

If search_date = Date + 2 Then

    If UBound(arr) = 8 Then
    
    'Забираем с 0 по 21 час с шагом 3 для итогового массива, но с шагом 1 по массиву с прогнозом
    j = 0
    For k = 1 To 22 Step 3
    j = j + 1
        For n = 1 To 3
        p = 4 - n
        If arr(j, p) = "" Then arr(j, p) = 0
        arr4(n, k) = arr(j, p)
        arr4(5, k) = arr2(j, 1)
        arr4(6, k) = arr3(j, 1)
        Next n
    Next k
    'Заполняем 1,4,7,10,13,16,19 часы
    For k = 2 To 20 Step 3
        arr4(1, k) = WorksheetFunction.Round((2 * arr4(1, k - 1) + arr4(1, k + 2)) / 3, 0)
        arr4(2, k) = WorksheetFunction.Round((2 * arr4(2, k - 1) + arr4(2, k + 2)) / 3, 0)
        arr4(3, k) = WorksheetFunction.Round((2 * arr4(3, k - 1) + arr4(3, k + 2)) / 3, 0)
        arr4(5, k) = ""
        arr4(6, k) = ""
    Next k
    'Заполняем 2,5,8,11,14,17,20 часы
    For k = 3 To 21 Step 3
        arr4(1, k) = WorksheetFunction.Round((arr4(1, k - 2) + 2 * arr4(1, k + 1)) / 3, 0)
        arr4(2, k) = WorksheetFunction.Round((arr4(2, k - 2) + 2 * arr4(2, k + 1)) / 3, 0)
        arr4(3, k) = WorksheetFunction.Round((arr4(3, k - 2) + 2 * arr4(3, k + 1)) / 3, 0)
        arr4(5, k) = ""
        arr4(6, k) = ""
    Next k
    'Заполняем 22,23 часы
    For k = 23 To 24
        arr4(1, k) = arr4(1, 22)
        arr4(2, k) = arr4(2, 22)
        arr4(3, k) = arr4(3, 22)
        arr4(5, k) = ""
        arr4(6, k) = ""
    Next k
    
    End If
    
    If UBound(arr) = 10 Then
    
    'Забираем 0 и 3 час с шагом 3
    For k = 1 To 4 Step 3
        For n = 1 To 3
        p = 4 - n
        If arr(k, p) = "" Then arr(k, p) = 0
        arr4(n, k) = arr(k, p)
        arr4(5, k) = arr2(k, 1)
        arr4(6, k) = arr3(k, 1)
        Next n
    Next k
    'Забираем с 6 по 21 час с шагом 3 для итогового массива, но с шагом 1 по массиву с прогнозом
    j = 4
    For k = 7 To 22 Step 3
    j = j + 1
        For n = 1 To 3
        p = 4 - n
        If arr(j, p) = "" Then arr(j, p) = 0
        arr4(n, k) = arr(j, p)
        arr4(5, k) = arr2(j, 1)
        arr4(6, k) = arr3(j, 1)
        Next n
    Next k
    'Заполняем 1,4,7,10,13,16,19 часы
    For k = 2 To 20 Step 3
        arr4(1, k) = WorksheetFunction.Round((2 * arr4(1, k - 1) + arr4(1, k + 2)) / 3, 0)
        arr4(2, k) = WorksheetFunction.Round((2 * arr4(2, k - 1) + arr4(2, k + 2)) / 3, 0)
        arr4(3, k) = WorksheetFunction.Round((2 * arr4(3, k - 1) + arr4(3, k + 2)) / 3, 0)
        arr4(5, k) = ""
        arr4(6, k) = ""
    Next k
    'Заполняем 2,5,8,11,14,17,20 часы
    For k = 3 To 21 Step 3
        arr4(1, k) = WorksheetFunction.Round((arr4(1, k - 2) + 2 * arr4(1, k + 1)) / 3, 0)
        arr4(2, k) = WorksheetFunction.Round((arr4(2, k - 2) + 2 * arr4(2, k + 1)) / 3, 0)
        arr4(3, k) = WorksheetFunction.Round((arr4(3, k - 2) + 2 * arr4(3, k + 1)) / 3, 0)
        arr4(5, k) = ""
        arr4(6, k) = ""
    Next k
    'Заполняем 22,23 часы
    For k = 23 To 24
        arr4(1, k) = arr4(1, 22)
        arr4(2, k) = arr4(2, 22)
        arr4(3, k) = arr4(3, 22)
        arr4(5, k) = ""
        arr4(6, k) = ""
    Next k
    
    End If
    
    If UBound(arr) = 14 Then
    
    'Заполняем c 0 по 9 часы
    For k = 1 To 10
        For n = 1 To 3
        p = 4 - n
        If arr(k, p) = "" Then arr(k, p) = 0
        arr4(n, k) = arr(k, p)
        Next n
    arr4(4, k) = WorksheetFunction.Round((1.7 * arr4(1, k) + 0.8 * arr4(2, k) + 0.5 * arr4(3, k)) / 3, 0)
    If arr4(4, k) > 100 Then arr4(4, k) = 100
    arr4(5, k) = arr2(k, 1)
    arr4(6, k) = arr3(k, 1)
    Next k
    'Заполняем 12,15,18,21 часы
    j = 10
    For k = 13 To 22 Step 3
    j = j + 1
        For n = 1 To 3
        p = 4 - n
        If arr(j, p) = "" Then arr(j, p) = 0
        arr4(n, k) = arr(j, p)
        Next n
    arr4(4, k) = WorksheetFunction.Round((1.7 * arr4(1, k) + 0.8 * arr4(2, k) + 0.5 * arr4(3, k)) / 3, 0)
    If arr4(4, k) > 100 Then arr4(4, k) = 100
    arr4(5, k) = arr2(j, 1)
    arr4(6, k) = arr3(j, 1)
    Next k
    'Заполняем 10,13,16,19 часы
    For k = 11 To 20 Step 3
        arr4(1, k) = WorksheetFunction.Round((2 * arr4(1, k - 1) + arr4(1, k + 2)) / 3, 0)
        arr4(2, k) = WorksheetFunction.Round((2 * arr4(2, k - 1) + arr4(2, k + 2)) / 3, 0)
        arr4(3, k) = WorksheetFunction.Round((2 * arr4(3, k - 1) + arr4(3, k + 2)) / 3, 0)
        arr4(4, k) = WorksheetFunction.Round((2 * arr4(4, k - 1) + arr4(4, k + 2)) / 3, 0)
        arr4(5, k) = arr4(5, k - 1)
        arr4(6, k) = arr4(6, k - 1)
    Next k
    'Заполняем 11,14,17,20 часы
    For k = 12 To 21 Step 3
        arr4(1, k) = WorksheetFunction.Round((arr4(1, k - 2) + 2 * arr4(1, k + 1)) / 3, 0)
        arr4(2, k) = WorksheetFunction.Round((arr4(2, k - 2) + 2 * arr4(2, k + 1)) / 3, 0)
        arr4(3, k) = WorksheetFunction.Round((arr4(3, k - 2) + 2 * arr4(3, k + 1)) / 3, 0)
        arr4(4, k) = WorksheetFunction.Round((arr4(4, k - 2) + 2 * arr4(4, k + 1)) / 3, 0)
        arr4(5, k) = arr4(5, k - 1)
        arr4(6, k) = arr4(6, k - 1)
    Next k
    'Заполняем 22,23 часы
    For k = 23 To 24
        arr4(1, k) = arr4(1, 22)
        arr4(2, k) = arr4(2, 22)
        arr4(3, k) = arr4(3, 22)
        arr4(4, k) = arr4(4, 22)
        arr4(5, k) = arr4(5, k - 1)
        arr4(6, k) = arr4(6, k - 1)
    Next k
    
    End If
    
End If

'Средняя облачность через коэффициенты по ярусам
For k = 1 To 24
arr4(4, k) = WorksheetFunction.Round((1.7 * arr4(1, k) + 0.8 * arr4(2, k) + 0.5 * arr4(3, k)) / 3, 0)
If arr4(4, k) > 100 Then arr4(4, k) = 100
Next k

'sh_frcst.Range("B3:Y8").ClearContents
sh_frcst.Range("B3:Y8").value = arr4
Erase arr4

End Sub