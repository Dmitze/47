Attribute VB_Name = "add_1_one_Troop"

'===========3.9.13===========
Sub CreateAddOneForOneTroop()

    Dim wbAllData As Workbook
    Dim wsAllData As Worksheet
    Dim wsReportingMove As Worksheet
    Dim tbAllData As ListObject
    Dim tbReportingMove As ListObject

    Set wbAllData = ThisWorkbook
    Set wsAllData = wbAllData.Worksheets("База")
    Set tbAllData = wsAllData.ListObjects("AllData")
    Set wsReportingMove = wbAllData.Worksheets("Додаток 1")
    Set tbReportingMove = wsReportingMove.ListObjects("tbAddOne")

    Dim frm As UserForm1
    Set frm = New UserForm1

    ' Формуємо словник унікальних підрозділів
    Dim dictUnits As Object
    Set dictUnits = CreateObject("Scripting.Dictionary")

    Dim cell As Range
    For Each cell In tbAllData.ListColumns("Звідки").DataBodyRange
        If Not dictUnits.exists(cell.value) And Not IsEmpty(cell.value) Then
            dictUnits.Add cell.value, Nothing
        End If
    Next cell

    For Each cell In tbAllData.ListColumns("Куди").DataBodyRange
        If Not dictUnits.exists(cell.value) And Not IsEmpty(cell.value) Then
            dictUnits.Add cell.value, Nothing
        End If
    Next cell

    Dim wsTroops As Worksheet
Set wsTroops = wbAllData.Worksheets("Підрозділи")

Dim k As Long
k = 1
wsTroops.Range("A1:A" & wsTroops.Rows.count).ClearContents ' Очистити перед вставкою

Dim unitKey As Variant
For Each unitKey In dictUnits.Keys
    wsTroops.Cells(k, 1).value = unitKey
    k = k + 1
Next unitKey

' Показуємо форму для вибору
Dim lastRowT As Long
lastRowT = wsTroops.Cells(wsTroops.Rows.count, 1).End(xlUp).Row


frm.Caption = "Виберіть підрозділ"
frm.ListData = wsTroops.Range("A1:A" & lastRowT)
    frm.show

    Dim selectedUnit As String
    If frm.ListBox1.ListIndex <> -1 Then
        selectedUnit = frm.ListBox1.value
    Else
        MsgBox "Підрозділ не вибрано!", vbExclamation
        Exit Sub
    End If

    wsReportingMove.Range("AddOneNameTroop") = selectedUnit 'записуємо підрозділ в додаток 1


    ' Зчитуємо діапазон дат
    Dim startDate As Date, endDate As Date
    Dim AddOnePeriodFrom As Variant
    Dim AddOnePeriodTo As Variant
    
    AddOnePeriodFrom = wsReportingMove.Range("AddOnePeriodFrom").value
    AddOnePeriodTo = wsReportingMove.Range("AddOnePeriodTo").value
    
    startDate = CDate(InputBox("Введіть дату початку періоду у форматі дд.мм.рррр", "Дата З", AddOnePeriodFrom))
    endDate = CDate(InputBox("Введіть дату кінця періоду у форматі дд.мм.рррр", "Дата ПО", AddOnePeriodTo))
    
    wsReportingMove.Range("AddOnePeriod") = startDate & "-" & endDate
    wsReportingMove.Range("AddOnePeriodFrom") = startDate
    wsReportingMove.Range("AddOnePeriodTo") = endDate
    
    

' Очистити всі рядки таблиці
    
        Dim i As Long
        For i = tbReportingMove.ListRows.count To 1 Step -1
            tbReportingMove.ListRows(i).Delete
        Next i
    
    
'    Dim docList As String
'docList = ""

    ' Підготовка словника результатів
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")

    For i = 1 To tbAllData.ListRows.count
        Dim rowDate As Variant
        rowDate = tbAllData.DataBodyRange(i, tbAllData.ListColumns("Дата документа").Index).value
        If IsDate(rowDate) Then
            Dim unitFrom As String, unitTo As String
            unitFrom = tbAllData.DataBodyRange(i, tbAllData.ListColumns("Звідки").Index).value
            unitTo = tbAllData.DataBodyRange(i, tbAllData.ListColumns("Куди").Index).value

            If (unitFrom = selectedUnit Or unitTo = selectedUnit) Then
                Dim key As String
                key = tbAllData.DataBodyRange(i, tbAllData.ListColumns("Найменування").Index).value & "||" & _
                      tbAllData.DataBodyRange(i, tbAllData.ListColumns("Ціна").Index).value & "||" & _
                      tbAllData.DataBodyRange(i, tbAllData.ListColumns("Одиниця виміру").Index).value

                Dim arr() As Variant
                If dict.exists(key) Then
                    arr = dict(key)
                Else
                    ReDim arr(1 To 4)
                End If

                Dim typeOp As String

                Dim qInKg As Double 'надійшло
                Dim qOutKg As Double 'вибуло
                
'                qInL = tbAllData.DataBodyRange(i, tbAllData.ListColumns("У літрах надійшло").Index).value
                qInKg = tbAllData.DataBodyRange(i, tbAllData.ListColumns("Надійшло").Index).value
'                qOutL = tbAllData.DataBodyRange(i, tbAllData.ListColumns("У літрах вибуло").Index).value
                qOutKg = tbAllData.DataBodyRange(i, tbAllData.ListColumns("Вибуло").Index).value
'залишки

'*трохи костиль, може не коректно працювати якщо є операції де отримувач і відправник один і той же підрозділ....
                If rowDate < startDate Then
                    If unitTo = selectedUnit Then
'                        arr(1) = arr(1) + qInL
                        arr(1) = arr(1) + qInKg + qOutKg
                    ElseIf unitFrom = selectedUnit Then
'                    End If
'                    If unitFrom = selectedUnit Then
'                        arr(1) = arr(1) - qOutL
                        arr(1) = arr(1) - qOutKg - qInKg
                    End If
                    GoTo nexti
                End If


                If rowDate >= startDate And rowDate <= endDate Then
'                Надійшло
                    If unitTo = selectedUnit Then
                        arr(2) = arr(2) + qOutKg + qInKg
'                        вибуло
                    ElseIf unitFrom = selectedUnit Then
                        arr(3) = arr(3) + qOutKg + qInKg
                    End If
'                        Вибуло
'                    If unitFrom = selectedUnit Then

'                        arr(3) = arr(3) + qOutKg + qInKg
'                    End If

                End If
'
                
nexti:
'                Залишок на кінець
                arr(4) = arr(1) - arr(3) + arr(2)
                dict(key) = arr
            End If
        End If
        Call ProgressBar.Progress(i / tbAllData.ListRows.count * 100)
        
    Next i

    ' Запис у таблицю
    Dim counter As Long
    counter = 0
    Dim arrParts() As String, itemKey As Variant
    For Each itemKey In dict.Keys
    
    
        Dim a
        a = dict(itemKey)
        If a(1) <> 0 Or a(2) <> 0 Or a(3) <> 0 Then 'перевірка, якщо позиція 0 і не було руху
            arrParts = Split(itemKey, "||")
              counter = counter + 1
               With tbReportingMove.ListRows.Add
                .Range(1, 1).value = counter
                .Range(1, 2).value = arrParts(0) ' Найменування
                .Range(1, 3).value = arrParts(2) ' Одиниця виміру
                .Range(1, 4).value = a(2) ' Надійшло
                .Range(1, 5).value = a(3) ' вибуло
                .Range(1, 6).value = a(4) ' Залишок на кінець
                .Range(1, 7).value = CDbl(arrParts(1)) ' Ціна
                .Range(1, 8).value = a(1) ' Залишок на початок

            End With
           
        End If
    Next itemKey
'    wsReportingMove.Range("rngAdditions") = docList

    MsgBox "За період з " & startDate & " по " & endDate & " створено ", vbInformation, "Додаток 1: " & selectedUnit

    
    
End Sub
