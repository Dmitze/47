Attribute VB_Name = "ReportingMove"

Sub ReportingMoveCreate() 'Донесення 440_18(підрозділ)

    Dim wbAllData As Workbook
    Dim wsAllData As Worksheet
    Dim wsReportingMove As Worksheet
    Dim tbAllData As ListObject
    Dim tbReportingMove As ListObject

    Set wbAllData = ThisWorkbook
    Set wsAllData = wbAllData.Worksheets("База")
    Set tbAllData = wsAllData.ListObjects("AllData")
    Set wsReportingMove = wbAllData.Worksheets("Донесення 440_18(підрозділ)")
    Set tbReportingMove = wsReportingMove.ListObjects("tbReportingMove")

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

    wsReportingMove.Range("rngSelectedUnit") = selectedUnit 'записуємо підрозділ в донесення


    ' Зчитуємо діапазон дат
    Dim startDate As Date, endDate As Date
    startDate = CDate(InputBox("Введіть дату початку періоду у форматі дд.мм.рррр", "Дата З", Date))
    endDate = CDate(InputBox("Введіть дату кінця періоду у форматі дд.мм.рррр", "Дата ПО", Date))
    
    wsReportingMove.Range("rngStartDate") = startDate
    wsReportingMove.Range("rngEndDate") = endDate

    ' Очистити всі рядки таблиці, крім першого
    If tbReportingMove.ListRows.count > 1 Then
        Dim i As Long
        For i = tbReportingMove.ListRows.count To 2 Step -1
            tbReportingMove.ListRows(i).Delete
        Next i
    End If
 
    

    ' Очистити клітинки першого рядка, крім колонок з формулами
    Dim firstRow As ListRow
    Set firstRow = tbReportingMove.ListRows(1)
    Dim j As Long
    For j = 1 To tbReportingMove.ListColumns.count
        Select Case j
            Case 1, 10, 11, 15, 16, 17
                ' Пропустити колонки з формулами
            Case Else
                firstRow.Range.Cells(1, j).ClearContents
        End Select
    Next j
    
Dim docList As String
docList = ""

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
                key = tbAllData.DataBodyRange(i, tbAllData.ListColumns("Код номенклатури").Index).value & "||" & _
                      tbAllData.DataBodyRange(i, tbAllData.ListColumns("Найменування").Index).value & "||" & _
                      tbAllData.DataBodyRange(i, tbAllData.ListColumns("Ціна").Index).value & "||" & _
                      tbAllData.DataBodyRange(i, tbAllData.ListColumns("Одиниця виміру").Index).value
'                      & "||" & _
'                      tbAllData.DataBodyRange(i, tbAllData.ListColumns("Примітка").Index).value & "||" & _
'                      tbAllData.DataBodyRange(i, tbAllData.ListColumns("Поле 11").Index).value

                Dim arr() As Variant
                If dict.exists(key) Then
                    arr = dict(key)
                Else
                    ReDim arr(1 To 14)
                End If

                Dim typeOp As String
                typeOp = Trim(LCase(tbAllData.DataBodyRange(i, tbAllData.ListColumns("Поле 10").Index).value))
                Dim DateDoc As Date
                Dim Doc As String
                Dim NumberDoc As String
                
                DateDoc = Trim(tbAllData.DataBodyRange(i, tbAllData.ListColumns("Дата документа").Index).value)
                Doc = Trim(tbAllData.DataBodyRange(i, tbAllData.ListColumns("Тип документа").Index).value)
                NumberDoc = Trim(tbAllData.DataBodyRange(i, tbAllData.ListColumns("№ документа").Index).value)
                
'                Dim AddMensure As String
'                AddMensure = Trim(tbAllData.DataBodyRange(i, tbAllData.ListColumns("Поле 11").Index).value)
'                arr(15) = DateDoc
'                arr(16) = NumberDoc
'                arr(17) = Doc
'                arr(18) = AddMensure
                
    
'                Dim qInL As Double 'надійшло в літрах
                Dim qInKg As Double 'надійшло в кілограмах
'                Dim qOutL As Double 'вибуло в літрах
                Dim qOutKg As Double 'вибуло в кілограмах
                
'                qInL = tbAllData.DataBodyRange(i, tbAllData.ListColumns("У літрах надійшло").Index).value
                qInKg = tbAllData.DataBodyRange(i, tbAllData.ListColumns("Надійшло").Index).value
'                qOutL = tbAllData.DataBodyRange(i, tbAllData.ListColumns("У літрах вибуло").Index).value
                qOutKg = tbAllData.DataBodyRange(i, tbAllData.ListColumns("Вибуло").Index).value

                If rowDate < startDate Then
                    If unitTo = selectedUnit Then
'                        arr(1) = arr(1) + qInL
                        arr(2) = arr(2) + qInKg
                    End If
                    If unitFrom = selectedUnit Then
'                        arr(1) = arr(1) - qOutL
                        arr(2) = arr(2) - qOutKg
                    End If
                    GoTo nexti
                End If

                If rowDate >= startDate And rowDate <= endDate Then
                    Select Case typeOp
                        Case "надійшло зі складу в/ч"
'                            arr(3) = arr(3) + qInL
                            arr(4) = arr(4) + qInKg
                        Case "надійшло з інших в/ч (підрозділів)"
'                            arr(5) = arr(5) + qInL
                            arr(6) = arr(6) + qInKg
                        Case "інші надходження"
'                            arr(7) = arr(7) + qInL
                            arr(8) = arr(8) + qInKg
                        Case "витрата"
'                            arr(9) = arr(9) + qOutL
                            arr(10) = arr(10) + qOutKg
                        Case "передано до інших в/ч (підрозділів)"
'                            arr(11) = arr(11) + qOutL
                            arr(12) = arr(12) + qOutKg
                        Case "інші вибуття"
'                            arr(13) = arr(13) + qOutL
                            arr(14) = arr(14) + qOutKg
                    End Select
                    
  
                     Dim newDoc As String
                    
                     ' Новий документ, який хочемо додати
                     newDoc = Doc & " №" & NumberDoc & " від " & DateDoc
                     
                     ' Перевіряємо, чи вже є такий документ у списку
                     If Not IsDocumentInList(docList, newDoc) Then
                         If Len(docList) > 0 Then
                             docList = docList & "; " & newDoc
                         Else
                             docList = newDoc
                         End If
                     End If
                End If
nexti:
                dict(key) = arr
            End If
        End If
        Call ProgressBar.Progress(i / tbAllData.ListRows.count * 100)
        
    Next i
'================================= тут закінчив, перепризначити залишки і рух, додати додатковий тип операції вкожен документ
    ' Запис у таблицю
    Dim arrParts() As String, itemKey As Variant
    For Each itemKey In dict.Keys
        Dim a
        a = dict(itemKey)
'        If a(1) <> 0 Or a(3) <> 0 Then
            arrParts = Split(itemKey, "||")
            If tbReportingMove.DataBodyRange.Rows(1).Cells(1, 2) = "" Then
                With tbReportingMove.DataBodyRange.Rows
                .Cells(1, 2).value = arrParts(1) ' Найменування
                .Cells(1, 3).value = arrParts(0) ' Код номенклатури
'                .Cells(1, 4).value = arrParts(5)
                .Cells(1, 4).value = arrParts(3) ' Одиниця виміру
                .Cells(1, 5).value = CDbl(arrParts(2)) ' Ціна
'                .Cells(1, 32).value = arrParts(4) ' Примітка
'                .Cells(1, 7).value = a(1)
                .Cells(1, 6).value = a(2) 'залишок на початок
'                .Cells(1, 9).value = a(3)
                .Cells(1, 7).value = a(4) 'надійшло зі складу частини
'                .Cells(1, 11).value = a(5)
                .Cells(1, 8).value = a(6) 'надійшло з інших в/ч, підрозділів
'                .Cells(1, 13).value = a(7)
                .Cells(1, 9).value = a(8) 'інше надходження
'                .Cells(1, 18).value = a(9)
                .Cells(1, 12).value = a(10) 'витрати
'                .Cells(1, 20).value = a(11)
                .Cells(1, 13).value = a(12) 'передано до інших в/ч, підрозділів
'                .Cells(1, 22).value = a(13)
                .Cells(1, 14).value = a(14) 'інше вибуття
'                .Range(1, tbReportingMove.ListColumns("Код номенклатури").Index).value = arrParts(0)
            End With
            
            Else
            With tbReportingMove.ListRows.Add
                .Range(1, 2).value = arrParts(1) ' Найменування
                .Range(1, 3).value = arrParts(0) ' Код номенклатури
'                .Cells(1, 4).value = arrParts(5)
                .Range(1, 4).value = arrParts(3) ' Одиниця виміру
                .Range(1, 5).value = CDbl(arrParts(2)) ' Ціна
'                .Cells(1, 32).value = arrParts(4) ' Примітка
'                .Cells(1, 7).value = a(1)
                .Range(1, 6).value = a(2) 'залишок на початок
'                .Cells(1, 9).value = a(3)
                .Range(1, 7).value = a(4) 'надійшло зі складу частини
'                .Cells(1, 11).value = a(5)
                .Range(1, 8).value = a(6) 'надійшло з інших в/ч, підрозділів
'                .Cells(1, 13).value = a(7)
                .Range(1, 9).value = a(8) 'інше надходження
'                .Cells(1, 18).value = a(9)
                .Range(1, 12).value = a(10) 'витрати
'                .Cells(1, 20).value = a(11)
                .Range(1, 13).value = a(12) 'передано до інших в/ч, підрозділів
'                .Cells(1, 22).value = a(13)
                .Range(1, 14).value = a(14) 'інше вибуття
'                .Range(1, tbReportingMove.ListColumns("Код номенклатури").Index).value = arrParts(0)
            End With
            End If
'        End If
    Next itemKey
    wsReportingMove.Range("rngAdditions") = docList

    MsgBox "За період з " & startDate & " по " & endDate & " створене ", vbInformation, "Донесення про рух майна по підрозділу: " & selectedUnit

End Sub


