Attribute VB_Name = "Add14"

Option Explicit

' ============================================================
'  Додаток 13 (Книга обліку д14) - виправлена версія
'  Головна зміна: ПЕРЕД заповненням перевіряються всі дані бази
'  (дата, ціна, кількість, тип операції) і показується точний
'  рядок аркуша "База", де формат неправильний.
'  Потрібні в книзі: функція CleanFileName, форма ProgressBar.
'  Функція FindValue більше не потрібна.
'
'  ЗМІНИ (v2):
'   1. Підтягуються ТІЛЬКИ накладні: видача зі складу на підрозділ і
'      повернення з підрозділу на склад. Інше (акти, списання, ...) пропускається.
'   2. НАПРЯМОК у книзі: видача на підрозділ = НАДІЙШЛО, повернення на склад = ВИБУЛО
'      (раніше було навпаки: видача йшла у "вибуло", а повернення - у "надійшло").
'   3. Сортування: рік накладної -> номер накладної (за числом) -> дата -> порядок бази.
'   4. Залишок "становить" рахується від попереднього рядка; якщо стає від'ємним
'      (повернули більше, ніж видали) - у кінці показується попередження.
' ============================================================

' Документи цього типу вважаються накладними (видача / повернення)
Private Const NUMBERED_DOC_WORD As String = "накладн"
' Слово в назві складу: визначає напрямок (Звідки = склад -> видача, Куди = склад -> повернення)
Private Const WAREHOUSE_WORD As String = "склад"

' ключі сортування (заповнюються в Add13SortRows)
Private m13Year() As Long
Private m13Num() As Double
Private m13Date() As Double

' ---------- допоміжні функції ----------

Private Function SafeStr(ByVal v As Variant) As String
    If IsError(v) Then
        SafeStr = ""
    ElseIf IsNull(v) Or IsEmpty(v) Then
        SafeStr = ""
    Else
        SafeStr = CStr(v)
    End If
End Function

Private Function ToNum(ByVal v As Variant) As Double
    If IsError(v) Then
        ToNum = 0
    ElseIf IsEmpty(v) Or IsNull(v) Then
        ToNum = 0
    ElseIf IsNumeric(v) Then
        ToNum = CDbl(v)
    Else
        ToNum = 0
    End If
End Function

' Порожня комірка = 0 (ок). Текст, який не є числом, або помилка = неправильно
Private Function IsGoodNumber(ByVal v As Variant) As Boolean
    If IsError(v) Then
        IsGoodNumber = False
    ElseIf IsNull(v) Then
        IsGoodNumber = False
    ElseIf IsEmpty(v) Then
        IsGoodNumber = True
    Else
        IsGoodNumber = IsNumeric(v)
    End If
End Function

Private Function TryGetDate(ByVal v As Variant, ByRef d As Date) As Boolean
    TryGetDate = False
    If IsError(v) Then Exit Function
    If IsNull(v) Or IsEmpty(v) Then Exit Function
    Select Case VarType(v)
        Case vbDate
            d = v
            TryGetDate = True
        Case vbString
            If IsDate(v) Then
                d = CDate(v)
                TryGetDate = True
            End If
        Case vbDouble, vbLong, vbInteger, vbSingle, vbCurrency
            If v > 0 And v < 2958466 Then
                d = CDate(v)
                TryGetDate = True
            End If
    End Select
End Function

Private Function GetLastNumber(ByVal ws As Worksheet, ByVal lastCol As Long) As Long
    Dim v As Variant
    v = ws.Cells(22, lastCol).value
    If Not IsError(v) And Not IsEmpty(v) And IsNumeric(v) Then
        GetLastNumber = CLng(v)
    Else
        GetLastNumber = CLng(Application.WorksheetFunction.Max(ws.Rows(22)))
    End If
End Function

Private Function FolderExists(ByVal p As String) As Boolean
    FolderExists = False
    If Len(p) = 0 Then Exit Function
    On Error Resume Next
    FolderExists = (Dir(p, vbDirectory) <> vbNullString)
    On Error GoTo 0
End Function

' ---------- накладні: вид операції, фільтр, сортування ----------

' 1 = видача зі складу на підрозділ (в книзі НАДІЙШЛО)
' 2 = повернення з підрозділу на склад (в книзі ВИБУЛО)
' 0 = не підходить (не накладна, акт, списання, переміщення між підрозділами ...)
Private Function Add13Kind(ByVal docType As String, ByVal opType As String, _
                           ByVal FromUnit As String, ByVal toUnit As String) As Long
    Dim fromW As Boolean, toW As Boolean

    If InStr(1, docType, NUMBERED_DOC_WORD, vbTextCompare) = 0 Then Exit Function
    opType = LCase$(Trim$(opType))
    If opType <> "внутрішнє переміщення" And opType <> "надходження" Then Exit Function

    fromW = (InStr(1, FromUnit, WAREHOUSE_WORD, vbTextCompare) > 0)
    toW = (InStr(1, toUnit, WAREHOUSE_WORD, vbTextCompare) > 0)

    If fromW And Not toW Then
        Add13Kind = 1
    ElseIf toW And Not fromW Then
        Add13Kind = 2
    End If
End Function

' Залишає в масиві лише накладні з видачею / поверненням. skipped = скільки рядків відкинуто.
Private Sub Add13KeepInvoices(ByRef data() As Variant, ByRef srcRow() As Long, ByRef skipped As Long)
    Dim n As Long, cols As Long
    Dim r As Long, c As Long, k As Long
    Dim keep() As Boolean
    Dim kept As Long
    Dim tmp() As Variant
    Dim tmpSrc() As Long

    skipped = 0
    n = UBound(data, 1) - 1
    cols = UBound(data, 2)
    ReDim keep(2 To n + 1)

    For r = 2 To n + 1
        keep(r) = (Add13Kind(SafeStr(data(r, 25)), SafeStr(data(r, 26)), SafeStr(data(r, 8)), SafeStr(data(r, 9))) > 0)
        If keep(r) Then kept = kept + 1
    Next r

    skipped = n - kept
    If skipped = 0 Then Exit Sub

    ReDim tmp(1 To kept + 1, 1 To cols)
    ReDim tmpSrc(1 To kept + 1)
    For c = 1 To cols
        tmp(1, c) = data(1, c)
    Next c
    k = 1
    For r = 2 To n + 1
        If keep(r) Then
            k = k + 1
            tmpSrc(k) = srcRow(r)
            For c = 1 To cols
                tmp(k, c) = data(r, c)
            Next c
        End If
    Next r
    data = tmp
    srcRow = tmpSrc
End Sub

Private Function Add13Sgn(ByVal x As Double, ByVal y As Double) As Long
    If x < y Then
        Add13Sgn = -1
    ElseIf x > y Then
        Add13Sgn = 1
    End If
End Function

' рік -> номер -> дата. Рівні рядки лишаються в порядку бази (стабільне сортування).
Private Function Add13Cmp(ByVal a As Long, ByVal b As Long) As Long
    Dim s As Long
    s = Add13Sgn(m13Year(a), m13Year(b))
    If s <> 0 Then
        Add13Cmp = s
        Exit Function
    End If
    s = Add13Sgn(m13Num(a), m13Num(b))
    If s <> 0 Then
        Add13Cmp = s
        Exit Function
    End If
    Add13Cmp = Add13Sgn(m13Date(a), m13Date(b))
End Function

' "112/п/ІНЖ" -> 112 ; "14565-в" -> 14565 ; без числа на початку -> в кінець року
Private Function Add13LeadingNumber(ByVal s As String) As Double
    Dim i As Long
    s = Trim$(s)
    i = 1
    Do While i <= Len(s)
        If Not (Mid$(s, i, 1) Like "#") Then Exit Do
        i = i + 1
    Loop
    If i > 1 Then
        Add13LeadingNumber = val(Left$(s, i - 1))
    Else
        Add13LeadingNumber = 1E+300
    End If
End Function

Private Sub Add13SortRows(ByRef data() As Variant, ByRef srcRow() As Long)
    Dim n As Long, cols As Long
    Dim i As Long, j As Long, c As Long
    Dim idx() As Long
    Dim cur As Long
    Dim dt As Date
    Dim tmp() As Variant
    Dim tmpSrc() As Long

    n = UBound(data, 1) - 1
    If n < 2 Then Exit Sub
    cols = UBound(data, 2)

    ReDim m13Year(1 To n)
    ReDim m13Num(1 To n)
    ReDim m13Date(1 To n)
    ReDim idx(1 To n)

    For i = 1 To n
        If TryGetDate(data(i + 1, 13), dt) Then
            m13Date(i) = CDbl(dt)
            m13Year(i) = Year(dt)
        Else
            m13Date(i) = 0
            m13Year(i) = 9999
        End If
        m13Num(i) = Add13LeadingNumber(SafeStr(data(i + 1, 14)))
        idx(i) = i
    Next i

    ' сортування вставками (стабільне)
    For i = 2 To n
        cur = idx(i)
        j = i - 1
        Do While j >= 1
            If Add13Cmp(idx(j), cur) <= 0 Then Exit Do
            idx(j + 1) = idx(j)
            j = j - 1
        Loop
        idx(j + 1) = cur
    Next i

    ReDim tmp(1 To n + 1, 1 To cols)
    ReDim tmpSrc(1 To n + 1)
    For c = 1 To cols
        tmp(1, c) = data(1, c)
    Next c
    For i = 1 To n
        tmpSrc(i + 1) = srcRow(idx(i) + 1)
        For c = 1 To cols
            tmp(i + 1, c) = data(idx(i) + 1, c)
        Next c
    Next i
    data = tmp
    srcRow = tmpSrc
End Sub

' ---------- формування додатка ----------

Sub CreateAdd13()
    Dim wbAllData As Workbook
    Dim wsAllData As Worksheet
    Dim tbAllData As ListObject
    Dim wsAdd13 As Worksheet
    Dim criteriaRange13 As Range
    Dim RngToSearchData As Range
    Dim CopyRange1 As Range
    Dim foundCell As Range
    Dim visibleRange As Range
    Dim area As Range
    Dim rowCell As Range

    Dim filteredData() As Variant
    Dim srcRow() As Long
    Dim rowCount As Long, colCount As Long
    Dim i As Long, j As Long, k As Long, n As Long, t As Long, c As Long
    Dim curN As Long

    Dim lastColumnStart As Long, LastRowStart As Long
    Dim lastRow13 As Long, LastColumn As Long, LastColumnFinal As Long, lastRowFinal As Long
    Dim lastNumber As Long
    Dim DocRow As Long
    Dim DocExist As Boolean, nameExist As Boolean

    Dim docName As String, NumberDocument As String, FromTo As String
    Dim DateDocument As Date, dtTmp As Date
    Dim typeOperation As String
    Dim itemName As String, Nomenclature As String, itemUnit As String
    Dim price As Double
    Dim generalInfo As String
    Dim quantity As Double, previousQuantyty As Double, actualQuantyty As Double
    Dim v21 As Variant
    Dim kind As Long, dirSign As Long, tgtRow As Long, colIn As Long
    Dim curBal As Double
    Dim skippedCnt As Long, negCount As Long
    Dim negText As String
    Dim msgRow As String, problems As String, problemCount As Long

    On Error GoTo ErrHandler

    Set wbAllData = ThisWorkbook
    Set wsAllData = wbAllData.Worksheets("База")
    Set tbAllData = wsAllData.ListObjects("AllData")
    Set wsAdd13 = wbAllData.Worksheets("Книга обліку д14")

    colCount = tbAllData.ListColumns.count
    If colCount < 26 Then
        MsgBox "У таблиці AllData лише " & colCount & " колонок, а код очікує щонайменше 26.", vbCritical, "Неправильна структура бази"
        GoTo CleanUp
    End If

    Application.ScreenUpdating = False
    wsAdd13.Activate

    ' --- Видаляємо попередні дані ---
    lastColumnStart = wsAdd13.Cells(22, wsAdd13.Columns.count).End(xlToLeft).Column + 1
    LastRowStart = wsAdd13.Cells.Find(What:="*", After:=wsAdd13.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
                    SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    On Error Resume Next
    If lastColumnStart >= 9 Then
        wsAdd13.Range(wsAdd13.Cells(1, 9), wsAdd13.Cells(LastRowStart, lastColumnStart)).EntireColumn.Delete
    End If
    wsAdd13.Rows("24:" & wsAdd13.Rows.count).Delete
    On Error GoTo ErrHandler

    Set RngToSearchData = wsAdd13.Range("12:12")
    Set criteriaRange13 = wsAdd13.Range("criteriaRange13")

    ' --- Знімаємо старі фільтри і фільтруємо за критеріями ---
    If wsAllData.FilterMode Then wsAllData.ShowAllData
    tbAllData.Range.AdvancedFilter Action:=xlFilterInPlace, criteriaRange:=criteriaRange13, Unique:=False

    On Error Resume Next
    Set visibleRange = tbAllData.DataBodyRange.SpecialCells(xlCellTypeVisible)
    On Error GoTo ErrHandler

    If visibleRange Is Nothing Then
        MsgBox "За заданими критеріями у базі не знайдено жодного запису.", vbInformation, "Немає даних"
        GoTo CleanUp
    End If

    ' --- Читаємо відфільтровані дані в масив ---
    rowCount = 0
    For Each area In visibleRange.Areas
        rowCount = rowCount + area.Rows.count
    Next area

    ReDim filteredData(1 To rowCount + 1, 1 To colCount)
    ReDim srcRow(1 To rowCount + 1)

    For k = 1 To colCount
        filteredData(1, k) = tbAllData.HeaderRowRange.Cells(1, k).value
    Next k

    i = 2
    For Each area In visibleRange.Areas
        For Each rowCell In area.Rows
            srcRow(i) = rowCell.Row
            For j = 1 To colCount
                filteredData(i, j) = rowCell.Cells(1, j).value
            Next j
            i = i + 1
        Next rowCell
    Next area

    ' --- залишаємо лише НАКЛАДНІ: видача на підрозділ і повернення на склад ---
    Add13KeepInvoices filteredData, srcRow, skippedCnt
    rowCount = UBound(filteredData, 1) - 1
    If rowCount < 1 Then
        MsgBox "Серед відібраних записів немає накладних з видачею майна на підрозділ або поверненням на склад.", _
               vbInformation, "Немає даних"
        GoTo CleanUp
    End If

    ' --- ПЕРЕВІРКА ДАНИХ перед заповненням ---
    problems = ""
    problemCount = 0
    For n = 2 To UBound(filteredData, 1)
        msgRow = ""
        If Not TryGetDate(filteredData(n, 13), dtTmp) Then
            msgRow = msgRow & " дата документа [" & SafeStr(filteredData(n, 13)) & "];"
        End If
        If Not IsGoodNumber(filteredData(n, 20)) Then
            msgRow = msgRow & " ціна [" & SafeStr(filteredData(n, 20)) & "];"
        End If
        Select Case LCase(Trim(SafeStr(filteredData(n, 26))))
            Case "надходження"
                If Not IsGoodNumber(filteredData(n, 6)) Then
                    msgRow = msgRow & " кількість надходження [" & SafeStr(filteredData(n, 6)) & "];"
                End If
            Case "вибуття"
                If Not IsGoodNumber(filteredData(n, 7)) Then
                    msgRow = msgRow & " кількість вибуття [" & SafeStr(filteredData(n, 7)) & "];"
                End If
            Case "внутрішнє переміщення"
                If Not IsGoodNumber(filteredData(n, 6)) And Not IsGoodNumber(filteredData(n, 7)) Then
                    msgRow = msgRow & " кількість переміщення [" & SafeStr(filteredData(n, 6)) & " / " & SafeStr(filteredData(n, 7)) & "];"
                End If
            Case Else
                msgRow = msgRow & " тип операції [" & SafeStr(filteredData(n, 26)) & "];"
        End Select

        If Len(msgRow) > 0 Then
            problemCount = problemCount + 1
            If problemCount <= 15 Then
                problems = problems & "Рядок " & srcRow(n) & " (" & SafeStr(filteredData(n, 25)) & _
                           " № " & SafeStr(filteredData(n, 14)) & "):" & msgRow & vbCrLf
            End If
        End If
    Next n

    If problemCount > 0 Then
        If problemCount > 15 Then problems = problems & "... і ще рядків: " & (problemCount - 15) & vbCrLf
        MsgBox "Знайдено записи з неправильним форматом на аркуші ""База""." & vbCrLf & _
               "Виправте їх і запустіть формування ще раз." & vbCrLf & vbCrLf & problems, _
               vbCritical, "Неправильні дані в базі"
        GoTo CleanUp
    End If

    ' --- сортування: рік накладної -> номер -> дата ---
    Add13SortRows filteredData, srcRow

    ' --- Основний цикл ---
    ' заголовок з назвою майна і категоріями
    Set CopyRange1 = wsAdd13.Range(wsAdd13.Cells(13, 6), wsAdd13.Cells(22, 8))

    For n = 2 To UBound(filteredData, 1)
        curN = n

        docName = SafeStr(filteredData(n, 25))
        Call TryGetDate(filteredData(n, 13), DateDocument)
        NumberDocument = SafeStr(filteredData(n, 14))
        FromTo = SafeStr(filteredData(n, 8)) & Chr(10) & "(" & SafeStr(filteredData(n, 9)) & ")"

        typeOperation = LCase(Trim(SafeStr(filteredData(n, 26))))
        itemName = SafeStr(filteredData(n, 2))
        Nomenclature = SafeStr(filteredData(n, 18))
        itemUnit = SafeStr(filteredData(n, 4))
        price = ToNum(filteredData(n, 20))

        generalInfo = itemName & Nomenclature & itemUnit & price

        ' чи є вже таке майно (колонки) на аркуші
        Set foundCell = RngToSearchData.Find(What:=generalInfo, LookIn:=xlValues, LookAt:=xlWhole)
        nameExist = Not (foundCell Is Nothing)

        lastRow13 = wsAdd13.Cells(wsAdd13.Rows.count, "A").End(xlUp).Row + 1
        LastColumn = wsAdd13.Cells.Find(What:="*", After:=wsAdd13.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
                        SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
        lastNumber = GetLastNumber(wsAdd13, LastColumn)

        ' чи є вже такий документ
        DocExist = False
        DocRow = 0
        For t = 24 To lastRow13 - 1
            If SafeStr(wsAdd13.Cells(t, 2).value) = docName _
               And SafeStr(wsAdd13.Cells(t, 3).value) = NumberDocument _
               And SafeStr(wsAdd13.Cells(t, 4).value) = CStr(DateDocument) _
               And SafeStr(wsAdd13.Cells(t, 5).value) = FromTo Then
                DocExist = True
                DocRow = t
                Exit For
            End If
        Next t

        If Not DocExist Then
            ' додаємо інформацію про документ
            With wsAdd13
                .Cells(lastRow13, 1).value = filteredData(n, 1)      'дата внесення
                .Cells(lastRow13, 2).value = docName                 'найменування документа
                .Cells(lastRow13, 3).NumberFormat = "@"
                .Cells(lastRow13, 3).value = NumberDocument          'номер документа
                .Cells(lastRow13, 4).value = DateDocument            'дата документа
                .Cells(lastRow13, 5).value = FromTo                  'підрозділ звідки (підрозділ куди)
            End With

            ' дублюємо всі колонки "становить" з попереднього рядка
            For c = 1 To LastColumn
                v21 = wsAdd13.Cells(21, c).value
                If Not IsError(v21) Then
                    If InStr(CStr(v21), "становить") > 0 Then
                        wsAdd13.Cells(lastRow13, c).value = wsAdd13.Cells(lastRow13 - 1, c).value
                    End If
                End If
            Next c
        End If

        ' якщо такого майна ще немає - додаємо 3 нові колонки
        If Not nameExist Then
            CopyRange1.Copy
            wsAdd13.Cells(13, LastColumn + 1).PasteSpecial xlPasteAll
            wsAdd13.Cells(12, LastColumn + 1).NumberFormat = "@"
            wsAdd13.Cells(12, LastColumn + 1).value = generalInfo
            wsAdd13.Cells(14, LastColumn + 1).value = itemName
            wsAdd13.Cells(16, LastColumn + 1).value = Nomenclature
            wsAdd13.Cells(18, LastColumn + 1).value = itemUnit
            wsAdd13.Cells(20, LastColumn + 1).value = price

            ' нумерація колонок
            wsAdd13.Cells(22, LastColumn + 1).value = lastNumber + 1
            wsAdd13.Cells(22, LastColumn + 2).value = lastNumber + 2
            wsAdd13.Cells(22, LastColumn + 3).value = lastNumber + 3
        End If

        ' --- напрямок: ВИДАЧА на підрозділ = НАДІЙШЛО, ПОВЕРНЕННЯ на склад = ВИБУЛО ---
        kind = Add13Kind(docName, typeOperation, SafeStr(filteredData(n, 8)), SafeStr(filteredData(n, 9)))
        If kind = 1 Then
            quantity = ToNum(filteredData(n, 7))                     ' видача: кількість у "Вибуло"
            If quantity = 0 Then quantity = ToNum(filteredData(n, 6))
            dirSign = 1
        Else
            quantity = ToNum(filteredData(n, 6))                     ' повернення: кількість у "Надійшло"
            If quantity = 0 Then quantity = ToNum(filteredData(n, 7))
            dirSign = -1
        End If

        If DocExist Then
            tgtRow = DocRow
        Else
            tgtRow = lastRow13
        End If

        If Not nameExist Then
            ' нове майно: колонки створені вище (LastColumn+1 надійшло, +2 вибуло, +3 становить)
            If dirSign > 0 Then
                wsAdd13.Cells(tgtRow, LastColumn + 1).value = quantity
            Else
                wsAdd13.Cells(tgtRow, LastColumn + 2).value = quantity
            End If
            curBal = dirSign * quantity
            wsAdd13.Cells(tgtRow, LastColumn + 3).value = curBal
        Else
            ' майно вже є: залишок = попередній залишок (він уже перенесений у цей рядок) +/- кількість
            colIn = foundCell.Column
            curBal = ToNum(wsAdd13.Cells(tgtRow, colIn + 2).value) + dirSign * quantity
            wsAdd13.Cells(tgtRow, colIn + 2).value = curBal
            If dirSign > 0 Then
                wsAdd13.Cells(tgtRow, colIn).value = ToNum(wsAdd13.Cells(tgtRow, colIn).value) + quantity
            Else
                wsAdd13.Cells(tgtRow, colIn + 1).value = ToNum(wsAdd13.Cells(tgtRow, colIn + 1).value) + quantity
            End If
        End If

        ' повернули більше, ніж видали -> залишок від'ємний
        If curBal < -0.000001 Then
            negCount = negCount + 1
            If negCount <= 10 Then
                negText = negText & itemName & ": накладна " & NumberDocument & " від " & Format$(DateDocument, "dd.mm.yyyy") & _
                          ", залишок " & curBal & vbCrLf
            End If
        End If

        ' копіюємо стилі загальної інформації для нового рядка
        If Not DocExist Then
            wsAdd13.Range(wsAdd13.Cells(23, 1), wsAdd13.Cells(23, 8)).Copy
            wsAdd13.Range(wsAdd13.Cells(lastRow13, 1), wsAdd13.Cells(lastRow13, 8)).PasteSpecial Paste:=xlPasteFormats
        End If

        Set foundCell = Nothing
        nameExist = False
        DocExist = False

        Call ProgressBar.Progress(n / UBound(filteredData, 1) * 100)
    Next n
    curN = 0

    ' --- Стилі показників підрозділів ---
    LastColumnFinal = wsAdd13.Cells.Find(What:="*", After:=wsAdd13.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
                        SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    lastRowFinal = wsAdd13.Cells(wsAdd13.Rows.count, "A").End(xlUp).Row
    If LastColumnFinal >= 9 And lastRowFinal >= 24 Then
        wsAdd13.Range(wsAdd13.Cells(23, 6), wsAdd13.Cells(23, 8)).Copy
        wsAdd13.Range(wsAdd13.Cells(23, 9), wsAdd13.Cells(lastRowFinal, LastColumnFinal)).PasteSpecial Paste:=xlPasteFormats
    End If

    msgRow = "Додаток сформовано." & vbCrLf & "Оброблено накладних (видача / повернення): " & rowCount
    If skippedCnt > 0 Then msgRow = msgRow & vbCrLf & "Не підтягнуто (не накладні або інші операції): " & skippedCnt
    If negCount > 0 Then
        msgRow = msgRow & vbCrLf & vbCrLf & "УВАГА: від'ємний залишок (повернено більше, ніж видано?): " & negCount & vbCrLf & negText
        MsgBox msgRow, vbExclamation, "Готово, є попередження"
    Else
        MsgBox msgRow, vbInformation, "Готово"
    End If

CleanUp:
    On Error Resume Next
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    If Not wsAllData Is Nothing Then
        If wsAllData.FilterMode Then wsAllData.ShowAllData
    End If
    Exit Sub

ErrHandler:
    Dim errText As String
    errText = "Помилка " & Err.Number & ": " & Err.Description
    If curN > 0 Then
        errText = errText & vbCrLf & vbCrLf & "Запис: рядок " & srcRow(curN) & " аркуша ""База""" & vbCrLf & _
                  SafeStr(filteredData(curN, 25)) & " № " & SafeStr(filteredData(curN, 14))
    End If
    MsgBox errText, vbCritical, "Помилка формування додатка"
    Resume CleanUp
End Sub

' ---------- збереження додатка в окремий файл ----------

Sub SaveAdd13()
    Dim wsSrc As Worksheet
    Dim WbNewWO As Workbook
    Dim i As Long
    Dim SavePath As String
    Dim DateAdd13 As String
    Dim fileName As String

    Set wsSrc = ThisWorkbook.Worksheets("Книга обліку д14")

    wsSrc.Copy                      ' копія аркуша в нову книгу
    Set WbNewWO = ActiveWorkbook

    ' видаляємо кнопки та інші фігури (у зворотному порядку)
    For i = WbNewWO.Worksheets(1).Shapes.count To 1 Step -1
        WbNewWO.Worksheets(1).Shapes(i).Delete
    Next i

    ' шлях до папки збереження
    SavePath = Trim(SafeStr(wsSrc.Range("PathToSaveAdd13").value))
    If Right(SavePath, 1) = Application.PathSeparator Then SavePath = Left(SavePath, Len(SavePath) - 1)

    ' якщо папки не існує - діалог вибору
    If Not FolderExists(SavePath) Then
        With Application.FileDialog(msoFileDialogFolderPicker)
            .Title = "Виберіть папку для збереження"
            .AllowMultiSelect = False
            If .show = -1 Then
                SavePath = .SelectedItems(1)
                wsSrc.Range("PathToSaveAdd13").value = SavePath
                If Right(SavePath, 1) = Application.PathSeparator Then SavePath = Left(SavePath, Len(SavePath) - 1)
            Else
                MsgBox "Збереження скасовано", vbExclamation
                Application.DisplayAlerts = False
                WbNewWO.Close SaveChanges:=False
                Application.DisplayAlerts = True
                Exit Sub
            End If
        End With
    End If

    DateAdd13 = Format(Date, "dd.mm.yyyy")
    fileName = SavePath & Application.PathSeparator & CleanFileName(DateAdd13 & " д13 " & ".xlsx")

    On Error GoTo WrongPath
    WbNewWO.SaveAs fileName:=fileName, FileFormat:=xlOpenXMLWorkbook
    On Error GoTo 0
    MsgBox "Файл збережено:" & vbCrLf & fileName, vbInformation, "Готово"
    Exit Sub

WrongPath:
    MsgBox "Не вдалося зберегти файл (" & Err.Description & ")." & vbCrLf & _
           "Перевірте адресу збереження (на зразок ""D:\База обліку майна ЗСУ 2024"", без косої лінії в кінці) і спробуйте ще раз.", _
           vbCritical, "Неправильний шлях збереження"
End Sub



