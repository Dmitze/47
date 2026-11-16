Attribute VB_Name = "Add14OneTroop"
Sub CreateAdd14OneTroop()
    Dim wbAllData As Workbook
    Dim wsAllData As Worksheet
    Dim tbAllData As ListObject
    Set wbAllData = ThisWorkbook
    Set wsAllData = wbAllData.Worksheets("База")
    Set tbAllData = wsAllData.ListObjects("AllData") 'таблиця з базою даних

    Dim wsAdd13 As Worksheet
    Set wsAdd13 = wbAllData.Worksheets("Книга обліку д14")

    Dim frm As UserForm1
    Set frm = New UserForm1

    Dim colFrom As Range, colTo As Range, celll As Range
    Dim dictUnits As Object
    Set dictUnits = CreateObject("Scripting.Dictionary")

    Set colFrom = tbAllData.ListColumns("Звідки").DataBodyRange
    Set colTo = tbAllData.ListColumns("Куди").DataBodyRange

    ' Формуємо словник унікальних підрозділів
    For Each celll In colFrom
        If Not dictUnits.exists(celll.Value) And Not IsEmpty(celll.Value) Then
            dictUnits.Add celll.Value, Nothing
        End If
    Next celll

    For Each celll In colTo
        If Not dictUnits.exists(celll.Value) And Not IsEmpty(celll.Value) Then
            dictUnits.Add celll.Value, Nothing
        End If
    Next celll

    Dim wsTroops As Worksheet
    Set wsTroops = wbAllData.Worksheets("Підрозділи")

    Dim k As Long
    k = 1
    wsTroops.Range("A1:A" & wsTroops.Rows.Count).ClearContents ' Очистити перед вставкою

    Dim unitKey As Variant
    For Each unitKey In dictUnits.Keys
        wsTroops.Cells(k, 1).Value = unitKey
        k = k + 1
    Next unitKey

    ' Зчитати назад у масив
    Dim lastRowT As Long
    lastRowT = wsTroops.Cells(wsTroops.Rows.Count, 1).End(xlUp).Row

    frm.Caption = "Виберіть підрозділ"
    frm.ListData = wsTroops.Range("A1:A" & lastRowT)

    ' Показати форму
    frm.Show

    Dim selectedUnit As String
    ' Отримати вибране
    If frm.ListBox1.ListIndex <> -1 Then
        selectedUnit = frm.ListBox1.Value
    Else
        MsgBox "Підрозділ не вибрано!", vbExclamation
        Exit Sub
    End If

    '===========================================================================
    ' Очищення попередніх даних на аркуші "Книга обліку д14"
    '===========================================================================
    Dim lastColumnStart As Long
    lastColumnStart = wsAdd13.Cells(22, wsAdd13.Columns.Count).End(xlToLeft).Column + 1

    Dim LastRowStart As Long
    LastRowStart = wsAdd13.Cells.Find(What:="*", After:=wsAdd13.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
                                SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row

    On Error Resume Next
    If lastColumnStart > 8 Then
        wsAdd13.Range(wsAdd13.Cells(1, 9), wsAdd13.Cells(LastRowStart, lastColumnStart)).EntireColumn.Delete
    End If
    If wsAdd13.Rows.Count >= 24 Then
        wsAdd13.Rows("24:" & wsAdd13.Rows.Count).Delete
    End If
    On Error GoTo 0

    Dim RngToSearchData As Range
    Set RngToSearchData = wsAdd13.Range("12:12")

    Dim filteredData() As Variant

    Dim criteriaRange14OneTroop As Range
    Set criteriaRange14OneTroop = wsAdd13.Range("criteriaRange14OneTroop")

    criteriaRange14OneTroop.Cells(2, 1) = selectedUnit
    criteriaRange14OneTroop.Cells(3, 2) = selectedUnit

    ' Знімаємо фільтри, якщо були встановлені
    If wsAllData.FilterMode Then
        wsAllData.ShowAllData
    End If

    ' Фільтруємо по заданих критеріях
    tbAllData.Range.AdvancedFilter Action:=xlFilterInPlace, criteriaRange:=criteriaRange14OneTroop, Unique:=False

    Dim rowCount As Long, colCount As Long
    Dim visibleRange As Range
    Dim cell As Range
    Dim area As Range

    rowCount = 0
    colCount = tbAllData.ListColumns.Count
    On Error Resume Next
    Set visibleRange = tbAllData.DataBodyRange.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    If Not visibleRange Is Nothing Then
        For Each area In visibleRange.Areas
            rowCount = rowCount + area.Rows.Count
        Next area

        ReDim filteredData(1 To rowCount + 1, 1 To colCount)

        For k = 1 To colCount
            filteredData(1, k) = tbAllData.HeaderRowRange.Cells(1, k).Value
        Next k

        Dim iRow As Long, jCol As Long
        iRow = 2
        For Each area In visibleRange.Areas
            For Each cell In area.Rows
                For jCol = 1 To colCount
                    filteredData(iRow, jCol) = cell.Cells(1, jCol).Value
                Next jCol
                iRow = iRow + 1
            Next cell
        Next area
    Else
        MsgBox "Немає операцій для обраного підрозділу: " & selectedUnit, vbInformation
        Exit Sub
    End If

    Dim lastRow13 As Long
    Dim CopyRange1 As Range
    Dim foundCell As Range

    ' Шаблон 3-х колонок майна: [надійшло], [вибуло], [становить]
    Set CopyRange1 = wsAdd13.Range(wsAdd13.Cells(13, 6), wsAdd13.Cells(22, 8))

    Dim n As Long
    For n = LBound(filteredData) + 1 To UBound(filteredData)

        Dim Document As String
        Dim DateDocument As Date
        Dim NumberDocument As String
        Dim FromTo As String
        Dim FromUnit As String
        Dim toUnit As String

        FromUnit = Trim(CStr(filteredData(n, 8))) ' підрозділ звідки
        toUnit = Trim(CStr(filteredData(n, 9)))   ' підрозділ куди

        Document = Trim(CStr(filteredData(n, 25)))
        NumberDocument = Trim(CStr(filteredData(n, 14)))
        FromTo = FromUnit & Chr(10) & "(" & toUnit & ")"

        On Error Resume Next
        DateDocument = CDate(filteredData(n, 13))
        On Error GoTo 0

        Dim typeOperation As String
        typeOperation = Trim(CStr(filteredData(n, 26)))

        '===========================================================================
        ' ФІЛЬТРАЦІЯ ОПЕРАЦІЙ (Вимога 5):
        ' 1. Відображаємо ВИКЛЮЧНО операції за Накладними (видача та повернення).
        ' 2. Знищене / списане майно (Акти списання, Рапорти, вибуття, "Списане")
        '    НЕ відображається у Книзі обліку д14 (лише тимчасовий рух майна).
        '===========================================================================
        If InStr(1, Document, "накладн", vbTextCompare) = 0 Then GoTo NextRow
        If toUnit = "Списане" Or toUnit = "Списання" Or FromUnit = "Списане" Or FromUnit = "Списання" Then GoTo NextRow
        If LCase(typeOperation) = "вибуття" Then GoTo NextRow
        If FromUnit = toUnit Then GoTo NextRow

        ' Визначаємо напрямок відносно обраного підрозділу
        Dim isIncoming As Boolean, isOutgoing As Boolean
        isIncoming = (toUnit = selectedUnit And FromUnit <> selectedUnit)
        isOutgoing = (FromUnit = selectedUnit And toUnit <> selectedUnit)

        If Not isIncoming And Not isOutgoing Then GoTo NextRow

        '===========================================================================
        ' ВИЗНАЧЕННЯ КІЛЬКОСТІ ОПЕРАЦІЇ:
        ' - При видачі (зі складу в підрозділ) кількість зазвичай у колонці G (Вибуло).
        ' - При поверненні (з підрозділу на склад) кількість зазвичай у колонці F (Надійшло).
        ' - Беремо ненульове значення з QtyOut або QtyIn.
        '===========================================================================
        Dim qIn As Double, qOut As Double, opQty As Double
        qIn = 0#: qOut = 0#
        If IsNumeric(filteredData(n, 6)) Then qIn = CDbl(filteredData(n, 6))
        If IsNumeric(filteredData(n, 7)) Then qOut = CDbl(filteredData(n, 7))

        If qOut > 0 Then
            opQty = qOut
        ElseIf qIn > 0 Then
            opQty = qIn
        Else
            opQty = 0#
        End If

        If opQty = 0# Then GoTo NextRow

        ' Дані про товар
        Dim name As String, Nomenclature As String, unit As String, price As Double
        name = Trim(CStr(filteredData(n, 2)))
        Nomenclature = Trim(CStr(filteredData(n, 18)))
        unit = Trim(CStr(filteredData(n, 4)))
        price = 0#
        If IsNumeric(filteredData(n, 20)) Then price = CDbl(filteredData(n, 20))

        Dim generalInfo As String
        generalInfo = name & Nomenclature & unit & price

        lastRow13 = wsAdd13.Cells(wsAdd13.Rows.Count, "A").End(xlUp).Row + 1

        Dim LastColumn As Long
        LastColumn = wsAdd13.Cells.Find(What:="*", After:=wsAdd13.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

        Dim lastNumber As Long
        lastNumber = 0
        If IsNumeric(wsAdd13.Cells(22, LastColumn).Value) Then
            lastNumber = CLng(wsAdd13.Cells(22, LastColumn).Value)
        End If

        ' Перевірка, чи є вже рядок для цього документа
        Dim t As Long, DocExist As Boolean, DocRow As Long
        DocExist = False
        DocRow = 0

        For t = 24 To lastRow13 - 1
            If wsAdd13.Cells(t, 2).Value = Document And _
               CStr(wsAdd13.Cells(t, 3).Value) = CStr(NumberDocument) And _
               wsAdd13.Cells(t, 4).Value = DateDocument And _
               wsAdd13.Cells(t, 5).Value = FromTo Then
                DocExist = True
                DocRow = t
                Exit For
            End If
        Next t

        ' Якщо документа ще немає — додаємо новий рядок
        If Not DocExist Then
            With wsAdd13
                .Cells(lastRow13, 1).Value = filteredData(n, 1) ' дата внесення
                .Cells(lastRow13, 2).Value = Document           ' найменування документа
                .Cells(lastRow13, 3).NumberFormat = "@"
                .Cells(lastRow13, 3).Value = NumberDocument     ' номер документа
                On Error GoTo WrongDate
                .Cells(lastRow13, 4).Value = DateDocument       ' дата документа
                On Error GoTo 0
                .Cells(lastRow13, 5).Value = FromTo             ' підрозділ звідки (підрозділ куди)
            End With

            ' Переносимо залишок ("становить") з попереднього рядка для всіх існуючих майн
            Dim cellTotal As Range
            For Each cellTotal In wsAdd13.Rows(21).Cells
                If cellTotal.Column > LastColumn Then Exit For
                If InStr(1, cellTotal.Value, "становить", vbTextCompare) > 0 Then
                    wsAdd13.Cells(lastRow13, cellTotal.Column).Value = wsAdd13.Cells(lastRow13 - 1, cellTotal.Column).Value
                End If
            Next cellTotal
        End If

        ' Цільовий рядок для запису операції
        Dim rowTarget As Long
        rowTarget = IIf(DocExist, DocRow, lastRow13)

        ' Перевіряємо, чи є вже стовпчики для цього майна
        Set foundCell = RngToSearchData.Find(generalInfo, LookIn:=xlValues, LookAt:=xlWhole)
        Dim nameExist As Boolean
        nameExist = Not (foundCell Is Nothing)

        Dim colIn As Long, colOut As Long, colBal As Long

        If Not nameExist Then
            ' Створюємо новий блок колонок [надійшло], [вибуло], [становить]
            CopyRange1.Copy
            wsAdd13.Cells(13, LastColumn + 1).PasteSpecial xlPasteAll

            wsAdd13.Cells(12, LastColumn + 1).Value = generalInfo
            wsAdd13.Cells(14, LastColumn + 1).Value = name
            wsAdd13.Cells(16, LastColumn + 1).Value = Nomenclature
            wsAdd13.Cells(18, LastColumn + 1).Value = unit
            wsAdd13.Cells(20, LastColumn + 1).Value = price

            wsAdd13.Cells(22, LastColumn + 1).Value = lastNumber + 1
            wsAdd13.Cells(22, LastColumn + 2).Value = lastNumber + 2
            wsAdd13.Cells(22, LastColumn + 3).Value = lastNumber + 3

            colIn = LastColumn + 1
            colOut = LastColumn + 2
            colBal = LastColumn + 3

            Set foundCell = wsAdd13.Cells(12, colIn)
            nameExist = True
        Else
            colIn = foundCell.Column
            colOut = foundCell.Column + 1
            colBal = foundCell.Column + 2
        End If

        '===========================================================================
        ' РОЗРАХУНОК ТА ЗАПИС КІЛЬКОСТІ І ЗАЛИШКУ ДЛЯ ПІДРОЗДІЛУ
        '===========================================================================
        Dim prevBal As Double, currBal As Double
        prevBal = 0#
        If rowTarget > 24 Then
            If IsNumeric(wsAdd13.Cells(rowTarget - 1, colBal).Value) Then
                prevBal = CDbl(wsAdd13.Cells(rowTarget - 1, colBal).Value)
            End If
        End If

        If isIncoming Then
            ' Сценарій А: НАДХОДЖЕННЯ В ПІДРОЗДІЛ (Видача зі складу)
            ' Записуємо у колонку 1 ("надійшло")
            wsAdd13.Cells(rowTarget, colIn).Value = wsAdd13.Cells(rowTarget, colIn).Value + opQty
        ElseIf isOutgoing Then
            ' Сценарій Б: ВИБУТТЯ З ПІДРОЗДІЛУ (Повернення на склад)
            ' Записуємо у колонку 2 ("вибуло")
            wsAdd13.Cells(rowTarget, colOut).Value = wsAdd13.Cells(rowTarget, colOut).Value + opQty
        End If

        Dim valIn As Double, valOut As Double
        valIn = 0#: valOut = 0#
        If IsNumeric(wsAdd13.Cells(rowTarget, colIn).Value) Then valIn = CDbl(wsAdd13.Cells(rowTarget, colIn).Value)
        If IsNumeric(wsAdd13.Cells(rowTarget, colOut).Value) Then valOut = CDbl(wsAdd13.Cells(rowTarget, colOut).Value)

        currBal = prevBal + valIn - valOut
        wsAdd13.Cells(rowTarget, colBal).Value = currBal

        ' Якщо операція додана в існуючий рядок документа (DocExist = True),
        ' каскадно оновлюємо залишок для всіх наступних рядків
        If DocExist Then
            Dim rCas As Long, casIn As Double, casOut As Double, casBal As Double
            casBal = currBal
            For rCas = rowTarget + 1 To lastRow13
                casIn = 0#: casOut = 0#
                If IsNumeric(wsAdd13.Cells(rCas, colIn).Value) Then casIn = CDbl(wsAdd13.Cells(rCas, colIn).Value)
                If IsNumeric(wsAdd13.Cells(rCas, colOut).Value) Then casOut = CDbl(wsAdd13.Cells(rCas, colOut).Value)
                casBal = casBal + casIn - casOut
                wsAdd13.Cells(rCas, colBal).Value = casBal
            Next rCas
        End If

        ' Копіюємо стилі загальної інформації для нового рядка
        If Not DocExist Then
            wsAdd13.Range(wsAdd13.Cells(23, 1), wsAdd13.Cells(23, 11)).Copy
            wsAdd13.Range(wsAdd13.Cells(lastRow13, 1), wsAdd13.Cells(lastRow13, 11)).PasteSpecial Paste:=xlPasteFormats
        End If

NextRow:
        Set foundCell = Nothing
        nameExist = False
        DocExist = False
        Call ProgressBar.Progress(n / UBound(filteredData) * 100)

    Next n

    ' Знаходимо останню колонку після всіх маніпуляцій
    Dim LastColumnFinal As Long
    LastColumnFinal = wsAdd13.Cells.Find(What:="*", After:=wsAdd13.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

    ' Копіюємо стилі основної інформації (для колонок показників)
    If LastColumnFinal >= 9 And lastRow13 >= 24 Then
        wsAdd13.Range(wsAdd13.Cells(23, 6), wsAdd13.Cells(23, 8)).Copy
        wsAdd13.Range(wsAdd13.Cells(23, 9), wsAdd13.Cells(lastRow13, LastColumnFinal)).PasteSpecial Paste:=xlPasteFormats
    End If

    ' Скидаємо фільтри
    On Error Resume Next
    wsAllData.ShowAllData
    On Error GoTo 0
    Exit Sub

WrongDate:
    MsgBox "Ймовірно, неправильний формат дати в " & Document & " № " & NumberDocument & ". Виправте дату і спробуйте сформувати додаток ще раз!", vbCritical, "Неправильна дата"
    wsAllData.ShowAllData
    Exit Sub
End Sub