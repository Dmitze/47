Attribute VB_Name = "add_1_one_Troop"
Option Explicit

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

    ' Словник унікальних підрозділів
    Dim dictUnits As Object
    Set dictUnits = CreateObject("Scripting.Dictionary")

    Dim cell As Range
    For Each cell In tbAllData.ListColumns("Звідки").DataBodyRange
        If Not dictUnits.exists(cell.Value) And Len(Trim(CStr(cell.Value))) > 0 Then
            dictUnits.Add cell.Value, Nothing
        End If
    Next cell

    For Each cell In tbAllData.ListColumns("Куди").DataBodyRange
        If Not dictUnits.exists(cell.Value) And Len(Trim(CStr(cell.Value))) > 0 Then
            dictUnits.Add cell.Value, Nothing
        End If
    Next cell

    Dim wsTroops As Worksheet
    Set wsTroops = wbAllData.Worksheets("Підрозділи")

    Dim k As Long
    k = 1
    wsTroops.Range("A1:A" & wsTroops.Rows.Count).ClearContents

    Dim unitKey As Variant
    For Each unitKey In dictUnits.Keys
        wsTroops.Cells(k, 1).Value = unitKey
        k = k + 1
    Next unitKey

    Dim lastRowT As Long
    lastRowT = wsTroops.Cells(wsTroops.Rows.Count, 1).End(xlUp).Row

    frm.Caption = "Виберіть підрозділ"
    frm.ListData = wsTroops.Range("A1:A" & lastRowT)
    frm.show

    Dim selectedUnit As String
    If frm.ListBox1.ListIndex <> -1 Then
        selectedUnit = frm.ListBox1.Value
    Else
        MsgBox "Підрозділ не вибрано!", vbExclamation
        Exit Sub
    End If

    wsReportingMove.Range("AddOneNameTroop") = selectedUnit

    ' Зчитуємо діапазон дат
    Dim startDate As Date, endDate As Date
    Dim AddOnePeriodFrom As Variant
    Dim AddOnePeriodTo As Variant
    
    AddOnePeriodFrom = wsReportingMove.Range("AddOnePeriodFrom").Value
    AddOnePeriodTo = wsReportingMove.Range("AddOnePeriodTo").Value
    
    startDate = CDate(InputBox("Введіть дату початку періоду у форматі дд.мм.рррр", "Дата З", AddOnePeriodFrom))
    endDate = CDate(InputBox("Введіть дату кінця періоду у форматі дд.мм.рррр", "Дата ПО", AddOnePeriodTo))
    
    wsReportingMove.Range("AddOnePeriod") = startDate & "-" & endDate
    wsReportingMove.Range("AddOnePeriodFrom") = startDate
    wsReportingMove.Range("AddOnePeriodTo") = endDate

    ' Очистити всі рядки таблиці
    Dim i As Long
    For i = tbReportingMove.ListRows.Count To 1 Step -1
        tbReportingMove.ListRows(i).Delete
    Next i

    ' Підготовка словника результатів
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")

    Dim cDateDoc As Long, cFrom As Long, cTo As Long
    Dim cName As Long, cPrice As Long, cUnit As Long
    Dim cIn As Long, cOut As Long, cOpType As Long

    With tbAllData
        cDateDoc = .ListColumns("Дата документа").Index
        cFrom = .ListColumns("Звідки").Index
        cTo = .ListColumns("Куди").Index
        cName = .ListColumns("Найменування").Index
        cPrice = .ListColumns("Ціна").Index
        cUnit = .ListColumns("Одиниця виміру").Index
        cIn = .ListColumns("Надійшло").Index
        cOut = .ListColumns("Вибуло").Index
        cOpType = .ListColumns("Запасне поле2").Index
    End With

    Dim dataRows As Long
    dataRows = tbAllData.ListRows.Count

    For i = 1 To dataRows
        Dim rowDateVal As Variant
        rowDateVal = tbAllData.DataBodyRange(i, cDateDoc).Value
        If IsDate(rowDateVal) Then
            Dim rDate As Date
            rDate = CDate(rowDateVal)
            
            Dim unitFrom As String, unitTo As String
            unitFrom = Trim(CStr(tbAllData.DataBodyRange(i, cFrom).Value))
            unitTo = Trim(CStr(tbAllData.DataBodyRange(i, cTo).Value))

            If (unitFrom = selectedUnit Or unitTo = selectedUnit) And rDate <= endDate Then
                Dim key As String
                key = CStr(tbAllData.DataBodyRange(i, cName).Value) & "||" & _
                      CStr(tbAllData.DataBodyRange(i, cPrice).Value) & "||" & _
                      CStr(tbAllData.DataBodyRange(i, cUnit).Value)

                Dim arr() As Variant
                If dict.exists(key) Then
                    arr = dict(key)
                Else
                    ReDim arr(1 To 4)
                    arr(1) = 0# ' Залишок на початок
                    arr(2) = 0# ' Надійшло за період
                    arr(3) = 0# ' Вибуло за період
                    arr(4) = 0# ' Залишок на кінець
                End If

                Dim opType As String
                opType = LCase(Trim(CStr(tbAllData.DataBodyRange(i, cOpType).Value)))

                Dim qIn As Double, qOut As Double, movQty As Double
                qIn = 0#: qOut = 0#
                If IsNumeric(tbAllData.DataBodyRange(i, cIn).Value) Then qIn = CDbl(tbAllData.DataBodyRange(i, cIn).Value)
                If IsNumeric(tbAllData.DataBodyRange(i, cOut).Value) Then qOut = CDbl(tbAllData.DataBodyRange(i, cOut).Value)

                Select Case opType
                    Case "надходження"
                        movQty = qIn
                        If unitTo = selectedUnit And movQty <> 0 Then
                            If rDate < startDate Then
                                arr(1) = arr(1) + movQty
                            Else
                                arr(2) = arr(2) + movQty
                            End If
                        End If

                    Case "вибуття"
                        movQty = qOut
                        If unitFrom = selectedUnit And movQty <> 0 Then
                            If rDate < startDate Then
                                arr(1) = arr(1) - movQty
                            Else
                                arr(3) = arr(3) + movQty
                            End If
                        End If

                    Case "внутрішнє переміщення"
                        movQty = qOut
                        If movQty = 0 Then movQty = qIn

                        If movQty <> 0 Then
                            If unitTo = selectedUnit Then
                                If rDate < startDate Then
                                    arr(1) = arr(1) + movQty
                                Else
                                    arr(2) = arr(2) + movQty
                                End If
                            End If
                            If unitFrom = selectedUnit Then
                                If rDate < startDate Then
                                    arr(1) = arr(1) - movQty
                                Else
                                    arr(3) = arr(3) + movQty
                                End If
                            End If
                        End If
                End Select

                arr(4) = arr(1) + arr(2) - arr(3)
                dict(key) = arr
            End If
        End If
        If dataRows > 0 Then Call ProgressBar.Progress(i / dataRows * 100)
    Next i

    ' Запис у таблицю
    Dim counter As Long
    counter = 0
    Dim arrParts() As String, itemKey As Variant
    For Each itemKey In dict.Keys
        Dim a As Variant
        a = dict(itemKey)
        ' Якщо є рух або ненульові залишки
        If Round(a(1), 6) <> 0 Or Round(a(2), 6) <> 0 Or Round(a(3), 6) <> 0 Or Round(a(4), 6) <> 0 Then
            arrParts = Split(itemKey, "||")
            counter = counter + 1
            With tbReportingMove.ListRows.Add
                .Range(1, 1).Value = counter
                .Range(1, 2).Value = arrParts(0) ' Найменування
                .Range(1, 3).Value = arrParts(2) ' Одиниця виміру
                .Range(1, 4).Value = a(2)        ' Надійшло за період
                .Range(1, 5).Value = a(3)        ' Вибуло за період
                .Range(1, 6).Value = a(4)        ' Залишок на кінець
                .Range(1, 7).Value = CDbl(arrParts(1)) ' Ціна
                .Range(1, 8).Value = a(1)        ' Залишок на початок
            End With
        End If
    Next itemKey

    MsgBox "За період з " & startDate & " по " & endDate & " сформовано рядків: " & counter, vbInformation, "Додаток 1: " & selectedUnit

End Sub
