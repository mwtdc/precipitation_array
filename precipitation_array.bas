Attribute VB_Name = "Module1"
Sub precip_array()

't = Timer

'ExcelEvents (False)
'Application.Calculation = xlManual

Dim sh_precip, sh_frcst As Worksheet
Dim arr, arr2, arr3, arr4(1 To 6, 1 To 24) As Variant
Dim i, j, x, k, n, p As Integer

Set sh_precip = ThisWorkbook.Sheets("������� ������")
Set sh_frcst = ThisWorkbook.Sheets("���������������")

search_date = CDate(sh_frcst.Range("O1").Value)
search_locality = sh_frcst.Range("T1").Value

Row_beg_locality = sh_precip.Range("A:A").Find(What:=search_locality, After:=sh_precip.Range("A1"), SearchOrder:=xlByRows, SearchDirection:=xlNext, LookAt:=xlWhole).Row
Row_end_locality = sh_precip.Range("A:A").Find(What:=search_locality, After:=sh_precip.Range("A1"), SearchOrder:=xlByRows, SearchDirection:=xlPrevious, LookAt:=xlWhole).Row
Row_beg_date = sh_precip.Range("E:E").Find(What:=search_date, After:=sh_precip.Range("E" & Row_beg_locality - 1), SearchOrder:=xlByRows, SearchDirection:=xlNext, LookAt:=xlWhole).Row
Row_end_date = sh_precip.Range("E:E").Find(What:=search_date, After:=sh_precip.Range("E" & Row_end_locality + 1), SearchOrder:=xlByRows, SearchDirection:=xlPrevious, LookAt:=xlWhole).Row

arr = sh_precip.Range("I" & Row_beg_date & ":K" & Row_end_date).Value
arr2 = sh_precip.Range("G" & Row_beg_date & ":G" & Row_end_date).Value
arr3 = sh_precip.Range("L" & Row_beg_date & ":L" & Row_end_date).Value

If search_date = Date + 1 Then

    If UBound(arr) = 18 Then
    
    '��������� c 0 �� 15 ����
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
    '��������� 18,21 ����
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
    '��������� 16,19 ����
    For k = 17 To 20 Step 3
        arr4(1, k) = WorksheetFunction.Round((2 * arr4(1, k - 1) + arr4(1, k + 2)) / 3, 0)
        arr4(2, k) = WorksheetFunction.Round((2 * arr4(2, k - 1) + arr4(2, k + 2)) / 3, 0)
        arr4(3, k) = WorksheetFunction.Round((2 * arr4(3, k - 1) + arr4(3, k + 2)) / 3, 0)
        arr4(4, k) = WorksheetFunction.Round((2 * arr4(4, k - 1) + arr4(4, k + 2)) / 3, 0)
        arr4(5, k) = arr4(5, k - 1)
        arr4(6, k) = arr4(6, k - 1)
    Next k
    '��������� 17,20 ����
    For k = 18 To 21 Step 3
        arr4(1, k) = WorksheetFunction.Round((arr4(1, k - 2) + 2 * arr4(1, k + 1)) / 3, 0)
        arr4(2, k) = WorksheetFunction.Round((arr4(2, k - 2) + 2 * arr4(2, k + 1)) / 3, 0)
        arr4(3, k) = WorksheetFunction.Round((arr4(3, k - 2) + 2 * arr4(3, k + 1)) / 3, 0)
        arr4(4, k) = WorksheetFunction.Round((arr4(4, k - 2) + 2 * arr4(4, k + 1)) / 3, 0)
        arr4(5, k) = arr4(5, k - 1)
        arr4(6, k) = arr4(6, k - 1)
    Next k
    '��������� 22,23 ����
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
    
    '��������� 22,23 ����
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
    
    '�������� � 0 �� 21 ��� � ����� 3 ��� ��������� �������, �� � ����� 1 �� ������� � ���������
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
    '��������� 1,4,7,10,13,16,19 ����
    For k = 2 To 20 Step 3
        arr4(1, k) = WorksheetFunction.Round((2 * arr4(1, k - 1) + arr4(1, k + 2)) / 3, 0)
        arr4(2, k) = WorksheetFunction.Round((2 * arr4(2, k - 1) + arr4(2, k + 2)) / 3, 0)
        arr4(3, k) = WorksheetFunction.Round((2 * arr4(3, k - 1) + arr4(3, k + 2)) / 3, 0)
        arr4(5, k) = ""
        arr4(6, k) = ""
    Next k
    '��������� 2,5,8,11,14,17,20 ����
    For k = 3 To 21 Step 3
        arr4(1, k) = WorksheetFunction.Round((arr4(1, k - 2) + 2 * arr4(1, k + 1)) / 3, 0)
        arr4(2, k) = WorksheetFunction.Round((arr4(2, k - 2) + 2 * arr4(2, k + 1)) / 3, 0)
        arr4(3, k) = WorksheetFunction.Round((arr4(3, k - 2) + 2 * arr4(3, k + 1)) / 3, 0)
        arr4(5, k) = ""
        arr4(6, k) = ""
    Next k
    '��������� 22,23 ����
    For k = 23 To 24
        arr4(1, k) = arr4(1, 22)
        arr4(2, k) = arr4(2, 22)
        arr4(3, k) = arr4(3, 22)
        arr4(5, k) = ""
        arr4(6, k) = ""
    Next k
    
    End If
    
    If UBound(arr) = 10 Then
    
    '�������� 0 � 3 ��� � ����� 3
    For k = 1 To 4 Step 3
        For n = 1 To 3
        p = 4 - n
        If arr(k, p) = "" Then arr(k, p) = 0
        arr4(n, k) = arr(k, p)
        arr4(5, k) = arr2(k, 1)
        arr4(6, k) = arr3(k, 1)
        Next n
    Next k
    '�������� � 6 �� 21 ��� � ����� 3 ��� ��������� �������, �� � ����� 1 �� ������� � ���������
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
    '��������� 1,4,7,10,13,16,19 ����
    For k = 2 To 20 Step 3
        arr4(1, k) = WorksheetFunction.Round((2 * arr4(1, k - 1) + arr4(1, k + 2)) / 3, 0)
        arr4(2, k) = WorksheetFunction.Round((2 * arr4(2, k - 1) + arr4(2, k + 2)) / 3, 0)
        arr4(3, k) = WorksheetFunction.Round((2 * arr4(3, k - 1) + arr4(3, k + 2)) / 3, 0)
        arr4(5, k) = ""
        arr4(6, k) = ""
    Next k
    '��������� 2,5,8,11,14,17,20 ����
    For k = 3 To 21 Step 3
        arr4(1, k) = WorksheetFunction.Round((arr4(1, k - 2) + 2 * arr4(1, k + 1)) / 3, 0)
        arr4(2, k) = WorksheetFunction.Round((arr4(2, k - 2) + 2 * arr4(2, k + 1)) / 3, 0)
        arr4(3, k) = WorksheetFunction.Round((arr4(3, k - 2) + 2 * arr4(3, k + 1)) / 3, 0)
        arr4(5, k) = ""
        arr4(6, k) = ""
    Next k
    '��������� 22,23 ����
    For k = 23 To 24
        arr4(1, k) = arr4(1, 22)
        arr4(2, k) = arr4(2, 22)
        arr4(3, k) = arr4(3, 22)
        arr4(5, k) = ""
        arr4(6, k) = ""
    Next k
    
    End If
    
    If UBound(arr) = 14 Then
    
    '��������� c 0 �� 9 ����
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
    '��������� 12,15,18,21 ����
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
    '��������� 10,13,16,19 ����
    For k = 11 To 20 Step 3
        arr4(1, k) = WorksheetFunction.Round((2 * arr4(1, k - 1) + arr4(1, k + 2)) / 3, 0)
        arr4(2, k) = WorksheetFunction.Round((2 * arr4(2, k - 1) + arr4(2, k + 2)) / 3, 0)
        arr4(3, k) = WorksheetFunction.Round((2 * arr4(3, k - 1) + arr4(3, k + 2)) / 3, 0)
        arr4(4, k) = WorksheetFunction.Round((2 * arr4(4, k - 1) + arr4(4, k + 2)) / 3, 0)
        arr4(5, k) = arr4(5, k - 1)
        arr4(6, k) = arr4(6, k - 1)
    Next k
    '��������� 11,14,17,20 ����
    For k = 12 To 21 Step 3
        arr4(1, k) = WorksheetFunction.Round((arr4(1, k - 2) + 2 * arr4(1, k + 1)) / 3, 0)
        arr4(2, k) = WorksheetFunction.Round((arr4(2, k - 2) + 2 * arr4(2, k + 1)) / 3, 0)
        arr4(3, k) = WorksheetFunction.Round((arr4(3, k - 2) + 2 * arr4(3, k + 1)) / 3, 0)
        arr4(4, k) = WorksheetFunction.Round((arr4(4, k - 2) + 2 * arr4(4, k + 1)) / 3, 0)
        arr4(5, k) = arr4(5, k - 1)
        arr4(6, k) = arr4(6, k - 1)
    Next k
    '��������� 22,23 ����
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

'������� ���������� ����� ������������ �� ������
For k = 1 To 24
arr4(4, k) = WorksheetFunction.Round((1.7 * arr4(1, k) + 0.8 * arr4(2, k) + 0.5 * arr4(3, k)) / 3, 0)
If arr4(4, k) > 100 Then arr4(4, k) = 100
Next k

'sh_frcst.Range("B3:Y8").ClearContents
sh_frcst.Range("B3:Y8").Value = arr4
Erase arr4

End Sub
