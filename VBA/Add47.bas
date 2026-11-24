Attribute VB_Name = "Add47"
Option Explicit

' =====================================================================
'  Додаток 47 - формування (CreateAdd47) та збереження (SaveAdd47)
'  ВИПРАВЛЕННЯ відносно попередньої версії:
'   1. Помилка на  quantity = filteredData(n, 7)  (Type mismatch):
'      кількість тепер читається як число (.Value2), а не як текст (.Text),
'      і перетворюється безпечною функцією Add47TryQty. Якщо в базі
'      некоректне значення - показується повідомлення з номером документа.
'   2. Додано Option Explicit, усі змінні оголошені.
'   3. Усі Cells(...) тепер із префіксом wsAdd47. (раніше працювали з активним аркушем).
'   4. Шість однакових блоків "знайти підрозділ / додати колонки / знайти категорію"
'      винесені в одну функцію Add47PostQty.
'   6. НОВЕ: не підтягуються рядки, де Z = вибуття і I = Списане (див. налаштування вгорі).
'   7. НОВЕ: сортування: рік -> накладні за номером, акти лише за датою (приходи перед видачею).
'   8. НОВЕ: повернення з підрозділу на склад ('надходження' зі Звідки = підрозділ) віднімається в підрозділу.
'   5. SaveAdd47: WbNew -> WbNewWO, оголошено savePath, DateAdd47 - рядок,
'      безпечна перевірка папки, DisplayAlerts, видалення фігур у зворотному порядку.
' =====================================================================

Private mCategoryError As String   ' текст помилки категорії (для повідомлення)

' ------------------------- НАЛАШТУВАННЯ ------------------------------
' True = НЕ підтягувати в Додаток 47 рядки, де в базі
'        "Запасне поле2" (колонка Z) містить слово "вибуття"  І
'        "Куди" (колонка I) містить слово "Списане".
Private Const SKIP_WRITE_OFF_ROWS As Boolean = True
' (запасний режим, за замовчуванням вимкнений) True = додатково пропускати УСІ рядки
'        серійного номера ("Примітка"), якщо для нього є списання.
Private Const SKIP_WRITTEN_OFF_SERIAL_ITEMS As Boolean = False
' ---------------------------------------------------------------------

' Документи цього типу сортуються за НОМЕРОМ (накладні мають наскрізну нумерацію).
' Усі інші (акти приймання-передачі, наряди тощо) - лише за датою.
Private Const NUMBERED_DOC_WORD As String = "накладн"
' True = повернення з підрозділу на склад додається ще й у колонку "Надійшло" (F), як і раніше.
Private Const RETURN_COUNTS_AS_RECEIVED As Boolean = True

' підрозділи, які мають майно (отримували видачу / списували) - для розпізнавання повернень
Private mInternal As Object

' ключі сортування (заповнюються в Add47SortRows)
Private mkYear() As Long
Private mkNum() As Double
Private mkRest() As String
Private mkDate() As Double
Private mkEff() As Double
Private mkCls() As Long

Sub CreateAdd47()

    Dim wbAllData As Workbook
    Dim wsAllData As Worksheet
    Dim tbAllData As ListObject
    Dim wsAdd47 As Worksheet
    Dim wsRef As Worksheet
    Dim RngToSearchData As Range

    Dim lastColumnStart As Long
    Dim LastRowStart As Long

    Dim filteredData() As Variant   ' масив з відфільтрованими даними

    Dim frm As UserForm1
    Dim NameOfThings As String
    Dim rngWithName As Range
    Dim ListGoods As Range

    Dim goodName As String
    Dim Nomenclature As String
    Dim unit As String
    Dim GoodPrice As Double
    Dim criteriaRange47 As Range

    Dim rowCount As Long
    Dim colCount As Long
    Dim visibleRange As Range
    Dim cell As Range
    Dim area As Range
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim n As Long

    Dim lastRow47 As Long
    Dim LastColumn As Long
    Dim LastColumnFinal As Long
    Dim CopyRange1 As Range

    Dim CategoryFrom As String
    Dim CategoryTo As String
    Dim typeOperation As String
    Dim quantity As Double
    Dim FromUnit As String
    Dim toUnit As String

    Dim cellAll As Range
    Dim cellWithSum As Range
    Dim sumRange As Range
    Dim criteriaRange As Range
    Dim criteriaValue As Variant
    Dim result As Double
    Dim removedCnt As Long

    Set wbAllData = ThisWorkbook
    Set wsAllData = wbAllData.Worksheets("База")
    Set tbAllData = wsAllData.ListObjects("AllData")
    Set wsAdd47 = wbAllData.Worksheets("Додаток 47")
    Set wsRef = wbAllData.Worksheets("Довідник д47")

    Set RngToSearchData = wsAdd47.Range("16:16")

    ' оновити таблицю з переліком МЦ перед вибором
    wsRef.PivotTables("ptReduses47").RefreshTable

    ' ---------------- Видаляємо попередні дані ----------------
    lastColumnStart = wsAdd47.Cells(19, wsAdd47.Columns.count).End(xlToLeft).Column + 1
    LastRowStart = wsAdd47.Cells.Find(What:="*", After:=wsAdd47.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
                                      SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    On Error Resume Next
    wsAdd47.Range(wsAdd47.Cells(1, 20), wsAdd47.Cells(LastRowStart, lastColumnStart)).EntireColumn.Delete
    wsAdd47.Rows("21:" & wsAdd47.Rows.count).Delete
    On Error GoTo 0

    ' ---------------- Форма вибору майна ----------------
    Set frm = UserForms.Add(UserForm1.name)
    frm.ListData = wsRef.Range("Goods47")
    frm.show

    NameOfThings = frm.TextBox1.value
    If Len(NameOfThings) = 0 Then GoTo DontSelect

    Set ListGoods = wsRef.Range("Goods47")
    Set rngWithName = ListGoods.Find(NameOfThings, LookIn:=xlValues, LookAt:=xlWhole)
    If rngWithName Is Nothing Then GoTo DontSelect

    ' значення за замовчуванням
    goodName = rngWithName.Offset(, -12).value

    Nomenclature = rngWithName.Offset(, -10).value
    If Nomenclature = "(пусто)" Then
        MsgBox "Ймовірно, ви не вказали код номенклатури вибраного майна, вкажіть код номенклатури (якщо не відомо, поставте б/н) і спробуйте створити додаток ще раз!", vbCritical, "Код номенклатури не вказаний"
        Exit Sub
    End If

    unit = rngWithName.Offset(, -11).value
    If unit = "(пусто)" Then
        MsgBox "Ймовірно, ви не вказали одиницю виміру вибраного майна, вкажіть одиницю виміру і спробуйте створити додаток ще раз!", vbCritical, "Одиниця виміру не вказана"
        Exit Sub
    End If

    On Error GoTo WrongPrice
    GoodPrice = rngWithName.Offset(, -9).value
    On Error GoTo 0

    ' ---------------- Критерії для фільтрації бази ----------------
    Set criteriaRange47 = wsAdd47.Range("criteriaRange47")
    With criteriaRange47
        .Cells(2, 1) = "'=" & goodName       ' найменування
        .Cells(2, 2) = "'=" & unit           ' одиниця виміру
        .Cells(2, 3) = "'=" & Nomenclature   ' код номенклатури
        .Cells(2, 4) = GoodPrice             ' ціна
    End With

    ' знімаємо фільтри, якщо були поставлені
    If wsAllData.FilterMode Then wsAllData.ShowAllData

    ' фільтруємо по заданих критеріях
    tbAllData.Range.AdvancedFilter Action:=xlFilterInPlace, criteriaRange:=criteriaRange47, Unique:=False

    ' ---------------- Копіюємо відфільтровані рядки в масив ----------------
    colCount = tbAllData.ListColumns.count
    On Error Resume Next
    Set visibleRange = tbAllData.DataBodyRange.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    If visibleRange Is Nothing Then GoTo DontSelect   ' нічого не знайдено в базі

    rowCount = 0
    For Each area In visibleRange.Areas
        rowCount = rowCount + area.Rows.count
    Next area

    ReDim filteredData(1 To rowCount + 1, 1 To colCount)

    ' заголовки таблиці
    For k = 1 To colCount
        filteredData(1, k) = tbAllData.HeaderRowRange.Cells(1, k).value
    Next k

    ' дані
    i = 2
    For Each area In visibleRange.Areas
        For Each cell In area.Rows
            For j = 1 To colCount
                If j = 6 Or j = 7 Then
                    ' [ВИПРАВЛЕНО] кількість (Надійшло / Вибуло) беремо як значення, а не як текст:
                    ' .Text давав "####" у вузьких колонках і ламав перетворення в Double
                    filteredData(i, j) = cell.Cells(1, j).Value2
                Else
                    filteredData(i, j) = CStr(cell.Cells(1, j).text)
                End If
            Next j
            i = i + 1
        Next cell
    Next area

    ' ---- не враховуємо вже знищене (списане) майно ----
    removedCnt = 0
    If SKIP_WRITTEN_OFF_SERIAL_ITEMS Or SKIP_WRITE_OFF_ROWS Then
        Add47RemoveWrittenOff filteredData, removedCnt
    End If
    If UBound(filteredData, 1) < 2 Then GoTo AllWrittenOff

    ' ---- сортування: рік документа -> номер документа -> дата документа ----
    Add47SortRows filteredData
    Add47BuildInternalUnits filteredData

    ' ---------------- Формуємо додаток ----------------
    wsAdd47.Activate   ' потрібно для Copy / PasteSpecial

    For n = LBound(filteredData, 1) + 1 To UBound(filteredData, 1)

        lastRow47 = wsAdd47.Cells(wsAdd47.Rows.count, "A").End(xlUp).Row + 1

        LastColumn = wsAdd47.Cells.Find(What:="*", After:=wsAdd47.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
                                        SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

        ' Якщо той самий документ і ті ж підрозділи - новий рядок не додаємо
        If wsAdd47.Cells(lastRow47 - 1, 1).value = filteredData(n, 1) And _
           wsAdd47.Cells(lastRow47 - 1, 2).value = filteredData(n, 25) And _
           wsAdd47.Cells(lastRow47 - 1, 3).value = filteredData(n, 14) And _
           wsAdd47.Cells(lastRow47 - 1, 4).value = filteredData(n, 13) And _
           wsAdd47.Cells(lastRow47 - 1, 5).value = filteredData(n, 8) & Chr(10) & "(" & filteredData(n, 9) & ")" Then

            lastRow47 = lastRow47 - 1

        Else
            wsAdd47.Range(wsAdd47.Cells(lastRow47 - 1, 14), wsAdd47.Cells(lastRow47 - 1, LastColumn)).Copy
            wsAdd47.Range(wsAdd47.Cells(lastRow47, 14), wsAdd47.Cells(lastRow47, LastColumn)).PasteSpecial xlPasteAll
        End If

        ' категорії відправника і отримувача з бази даних
        CategoryFrom = filteredData(n, 5)
        CategoryTo = filteredData(n, 27)
        typeOperation = filteredData(n, 26)

        ' заголовок з назвою підрозділу і категоріями (шаблон для нового підрозділу)
        Set CopyRange1 = wsAdd47.Range(wsAdd47.Cells(16, 14), wsAdd47.Cells(20, 19))

        Select Case typeOperation

            Case "надходження"
                If Not Add47TryQty(filteredData(n, 6), filteredData(n, 7), quantity) Then GoTo BadQuantity
                toUnit = filteredData(n, 9)
                FromUnit = filteredData(n, 8)

                If Add47IsReturn(FromUnit, toUnit, RngToSearchData) Then
                    ' ПОВЕРНЕННЯ з підрозділу на склад: у підрозділу віднімаємо, на складі додаємо
                    If RETURN_COUNTS_AS_RECEIVED Then Add47AddToCell wsAdd47, lastRow47, 6, quantity
                    If Not Add47PostQty(wsAdd47, RngToSearchData, CopyRange1, FromUnit, CategoryFrom, -quantity, lastRow47) Then GoTo BadCategory
                    If Not Add47PostQty(wsAdd47, RngToSearchData, CopyRange1, toUnit, CategoryTo, quantity, lastRow47) Then GoTo BadCategory
                Else
                    ' ОТРИМАННЯ на склад від постачальника / вищої частини
                    Add47AddToCell wsAdd47, lastRow47, 6, quantity
                    If Not Add47PostQty(wsAdd47, RngToSearchData, CopyRange1, toUnit, CategoryTo, quantity, lastRow47) Then GoTo BadCategory
                End If

            Case "вибуття"
                If Not Add47TryQty(filteredData(n, 7), filteredData(n, 6), quantity) Then GoTo BadQuantity
                FromUnit = filteredData(n, 8)

                Add47AddToCell wsAdd47, lastRow47, 7, quantity
                If Not Add47PostQty(wsAdd47, RngToSearchData, CopyRange1, FromUnit, CategoryFrom, -quantity, lastRow47) Then GoTo BadCategory

            Case "внутрішнє переміщення"
                If Not Add47TryQty(filteredData(n, 7), filteredData(n, 6), quantity) Then GoTo BadQuantity
                FromUnit = filteredData(n, 8)
                toUnit = filteredData(n, 9)

                ' підрозділ, з якого вибуває
                If Not Add47PostQty(wsAdd47, RngToSearchData, CopyRange1, FromUnit, CategoryFrom, -quantity, lastRow47) Then GoTo BadCategory
                ' підрозділ, до якого надходить
                If Not Add47PostQty(wsAdd47, RngToSearchData, CopyRange1, toUnit, CategoryTo, quantity, lastRow47) Then GoTo BadCategory

        End Select

        With wsAdd47
            .Cells(lastRow47, 1).value = filteredData(n, 1)    ' дата внесення
            .Cells(lastRow47, 2).value = filteredData(n, 25)   ' найменування документа
            .Cells(lastRow47, 3).NumberFormat = "@"
            .Cells(lastRow47, 3).value = filteredData(n, 14)   ' номер документа
            .Cells(lastRow47, 4).value = filteredData(n, 13)   ' дата документа
            .Cells(lastRow47, 5).value = filteredData(n, 8) & Chr(10) & "(" & filteredData(n, 9) & ")"  ' звідки (куди)
        End With

        ' копіюємо стилі загальної інформації
        wsAdd47.Range(wsAdd47.Cells(20, 1), wsAdd47.Cells(20, 19)).Copy
        wsAdd47.Range(wsAdd47.Cells(lastRow47, 1), wsAdd47.Cells(lastRow47, 19)).PasteSpecial Paste:=xlPasteFormats

        Call ProgressBar.Progress(n / UBound(filteredData) * 100)

    Next n

    ' ---------------- Завершальні дії ----------------
    ' остання колонка після всіх маніпуляцій
    LastColumnFinal = wsAdd47.Cells.Find(What:="*", After:=wsAdd47.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
                                         SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

    ' стилі основної інформації (для показників підрозділів)
    wsAdd47.Range(wsAdd47.Cells(20, 14), wsAdd47.Cells(20, 19)).Copy
    wsAdd47.Range(wsAdd47.Cells(21, 14), wsAdd47.Cells(lastRow47, LastColumnFinal)).PasteSpecial Paste:=xlPasteFormats

    ' проставляємо суму (стовпці "Усього")
    For Each cellAll In wsAdd47.Range(wsAdd47.Cells(17, 1), wsAdd47.Cells(17, LastColumnFinal)).Cells
        If Not IsError(cellAll.value) Then
            If cellAll.value = "Усього" Then
                Set cellWithSum = cellAll.Offset(2, 0)
                cellWithSum.AutoFill destination:=wsAdd47.Range(wsAdd47.Cells(cellWithSum.Row, cellWithSum.Column), _
                                                                wsAdd47.Cells(lastRow47, cellWithSum.Column))
            End If
        End If
    Next cellAll

    ' підсумки по категоріях
    For j = 9 To 13
        For i = 21 To lastRow47
            Set sumRange = wsAdd47.Range(wsAdd47.Cells(i, 14), wsAdd47.Cells(i, LastColumnFinal))
            Set criteriaRange = wsAdd47.Range(wsAdd47.Cells(18, 14), wsAdd47.Cells(18, LastColumnFinal))
            criteriaValue = wsAdd47.Cells(18, j).value

            result = Application.WorksheetFunction.SumIfs(sumRange, criteriaRange, criteriaValue)
            wsAdd47.Cells(i, j).value = result
        Next i
    Next j

    ' скидаємо область друку
    wsAdd47.PageSetup.PrintArea = ""

    ' скидаємо фільтри
    On Error Resume Next
    If wsAllData.FilterMode Then wsAllData.ShowAllData
    On Error GoTo 0
    Application.CutCopyMode = False
    If removedCnt > 0 Then
        MsgBox "Додаток сформовано." & vbCrLf & "Не підтягнуто рядків списання (вибуття -> Списане): " & removedCnt, _
               vbInformation, "Додаток 47"
    End If
    Exit Sub

' ===================== ПОМИЛКИ =====================
DontSelect:
    On Error Resume Next
    If wsAllData.FilterMode Then wsAllData.ShowAllData
    On Error GoTo 0
    MsgBox "Ймовірно, майно відсутнє в базі, або Ви не підтвердили вибір майна: вибране майно повинне ""перестрибнути"" з нижнього у верхнє поле вікна вибору майна, спробуйте сформувати додаток ще раз!", vbCritical, "Майно не знайдено"
    Exit Sub

WrongPrice:
    MsgBox "Ймовірно, ви не вказали ціну вибраного майна, вкажіть ціну (якщо не відома, поставте 1 грн.) і спробуйте створити додаток ще раз!", vbCritical, "Ціна не вказана"
    Exit Sub

BadQuantity:
    On Error Resume Next
    If wsAllData.FilterMode Then wsAllData.ShowAllData
    On Error GoTo 0
    MsgBox "Некоректна або порожня кількість у базі даних." & vbCrLf & vbCrLf & _
           "Тип операції: " & filteredData(n, 26) & vbCrLf & _
           "Документ: " & filteredData(n, 25) & " № " & filteredData(n, 14) & " від " & filteredData(n, 13) & vbCrLf & _
           "Дата внесення: " & filteredData(n, 1) & vbCrLf & vbCrLf & _
           "Перевірте колонки ""Надійшло"" / ""Вибуло"" у цьому рядку (має бути число), виправте і сформуйте додаток ще раз.", _
           vbCritical, "Помилка кількості"
    Exit Sub

BadCategory:
    On Error Resume Next
    If wsAllData.FilterMode Then wsAllData.ShowAllData
    On Error GoTo 0
    MsgBox "Не знайдено колонку категорії." & vbCrLf & mCategoryError & vbCrLf & vbCrLf & _
           "У базі неправильно вказана категорія: повинно бути 1, 2, 3, 4 або 5 " & _
           "у текстовому форматі (зліва зверху в клітинці з категорією зелений трикутник). " & _
           "Переконайтеся, що правильно проставлені категорії відправника та одержувача.", _
           vbCritical, "Категорія не знайдена"
    Exit Sub

AllWrittenOff:
    On Error Resume Next
    If wsAllData.FilterMode Then wsAllData.ShowAllData
    On Error GoTo 0
    MsgBox "За вибраним майном є лише записи списання (вибуття -> Списане), тому додаток не сформовано." & vbCrLf & vbCrLf & _
           "Щоб підтягувати їх, змініть SKIP_WRITE_OFF_ROWS на False у верхній частині модуля.", _
           vbInformation, "Майно вже списане"
    Exit Sub

End Sub


' =====================================================================
'  Допоміжні процедури
' =====================================================================

' Читає кількість: спочатку основна колонка, якщо вона порожня - запасна.
' Повертає False, якщо значення некоректне або обидві колонки порожні.
Private Function Add47TryQty(ByVal primary As Variant, ByVal fallback As Variant, ByRef result As Double) As Boolean
    result = 0
    If Not Add47IsBlank(primary) Then
        Add47TryQty = Add47ToDouble(primary, result)
    ElseIf Not Add47IsBlank(fallback) Then
        Add47TryQty = Add47ToDouble(fallback, result)
    End If
End Function

Private Function Add47IsBlank(ByVal v As Variant) As Boolean
    If IsError(v) Then Exit Function                ' помилка - це не "порожньо"
    If IsEmpty(v) Then
        Add47IsBlank = True
        Exit Function
    End If
    Add47IsBlank = (Len(Trim$(Replace(CStr(v), ChrW(160), ""))) = 0)
End Function

Private Function Add47ToDouble(ByVal v As Variant, ByRef result As Double) As Boolean
    Dim s As String
    Dim dec As String

    If IsError(v) Then Exit Function

    Select Case VarType(v)
        Case vbDouble, vbSingle, vbInteger, vbLong, vbCurrency, vbDecimal, vbByte
            result = CDbl(v)
            Add47ToDouble = True
            Exit Function
    End Select

    s = CStr(v)
    s = Replace(s, ChrW(160), "")      ' нерозривний пробіл
    s = Replace(s, " ", "")
    dec = Mid$(CStr(0.5), 2, 1)        ' десятковий роздільник системи
    s = Replace(s, ",", dec)
    s = Replace(s, ".", dec)

    If Len(s) = 0 Then Exit Function
    If Not IsNumeric(s) Then Exit Function

    result = CDbl(s)
    Add47ToDouble = True
End Function

' Додає delta до числа в клітинці (порожню клітинку вважає нулем)
Private Sub Add47AddToCell(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long, ByVal delta As Double)
    Dim cur As Double
    Dim v As Variant
    v = ws.Cells(r, c).value
    If Not IsError(v) Then
        If IsNumeric(v) Then cur = CDbl(v)
    End If
    ws.Cells(r, c).value = cur + delta
End Sub

' Знаходить (або створює) блок підрозділу в рядку 16, знаходить колонку потрібної
' категорії і додає delta в поточний рядок. Повертає False, якщо категорію не знайдено.
Private Function Add47PostQty(ByVal ws As Worksheet, ByVal rngUnits As Range, ByVal blockSrc As Range, _
                              ByVal unitName As String, ByVal category As String, _
                              ByVal delta As Double, ByVal lastRow As Long) As Boolean
    Dim lastCol As Long
    Dim afterCol As Long
    Dim foundCell As Range
    Dim catCell As Range

    If FindValue(unitName, rngUnits) Then
        ' підрозділ уже є
        Set foundCell = rngUnits.Find(unitName, LookIn:=xlValues, LookAt:=xlWhole)
        If foundCell Is Nothing Then
            Add47PostQty = True
            Exit Function
        End If
        afterCol = foundCell.Column
    Else
        ' підрозділу немає - додаємо новий блок колонок праворуч
        lastCol = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
                                SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
        ws.Activate
        blockSrc.Copy
        ws.Cells(16, lastCol + 1).PasteSpecial xlPasteAll
        ws.Cells(16, lastCol + 1).value = unitName
        afterCol = lastCol
    End If

    Set catCell = ws.Cells.Find(What:=category, After:=ws.Cells(18, afterCol), LookIn:=xlValues, LookAt:=xlWhole, _
                                SearchOrder:=xlByColumns, SearchDirection:=xlNext)
    If catCell Is Nothing Then
        mCategoryError = "Підрозділ: " & unitName & "; категорія: """ & category & """"
        Exit Function
    End If

    Add47AddToCell ws, lastRow, catCell.Column, delta
    Add47PostQty = True
End Function

Private Function Add47FolderExists(ByVal p As String) As Boolean
    If Len(p) = 0 Then Exit Function
    On Error Resume Next
    Add47FolderExists = ((GetAttr(p) And vbDirectory) = vbDirectory)
    On Error GoTo 0
End Function


' =====================================================================
'  Пропуск знищеного (списаного) майна
' =====================================================================

' Видаляє з масиву рядки знищеного майна. removed = кількість видалених рядків.
Private Sub Add47RemoveWrittenOff(ByRef arr() As Variant, ByRef removed As Long)
    Dim n As Long, cols As Long
    Dim r As Long, c As Long, k As Long
    Dim cNote As Long, cTo As Long, cType As Long
    Dim dead As Object
    Dim key As String
    Dim keep() As Boolean
    Dim kept As Long
    Dim tmp() As Variant

    removed = 0
    n = UBound(arr, 1) - 1
    If n < 1 Then Exit Sub
    cols = UBound(arr, 2)
    cNote = Add47HeaderIdx(arr, "Примітка", 16)   ' серійний номер
    cTo = Add47HeaderIdx(arr, "Куди", 9)
    cType = Add47HeaderIdx(arr, "Запасне поле2", 26)   ' тип операції

    Set dead = CreateObject("Scripting.Dictionary")
    ReDim keep(2 To n + 1)

    ' серійні номери, які вже списані
    If SKIP_WRITTEN_OFF_SERIAL_ITEMS Then
        For r = 2 To n + 1
            If Add47IsWrittenOff(CStr(arr(r, cTo))) Then
                key = Add47SerialKey(CStr(arr(r, cNote)))
                If Len(key) > 0 Then
                    If Not dead.exists(key) Then dead.Add key, 1
                End If
            End If
        Next r
    End If

    For r = 2 To n + 1
        keep(r) = True
        If SKIP_WRITTEN_OFF_SERIAL_ITEMS Then
            key = Add47SerialKey(CStr(arr(r, cNote)))
            If Len(key) > 0 Then
                If dead.exists(key) Then keep(r) = False
            End If
        End If
        If SKIP_WRITE_OFF_ROWS Then
            If Add47IsWriteOffRow(CStr(arr(r, cType)), CStr(arr(r, cTo))) Then keep(r) = False
        End If
        If keep(r) Then kept = kept + 1
    Next r

    removed = n - kept
    If removed = 0 Then Exit Sub

    ReDim tmp(1 To kept + 1, 1 To cols)
    For c = 1 To cols
        tmp(1, c) = arr(1, c)
    Next c
    k = 1
    For r = 2 To n + 1
        If keep(r) Then
            k = k + 1
            For c = 1 To cols
                tmp(k, c) = arr(r, c)
            Next c
        End If
    Next r
    arr = tmp
End Sub

' Рядок списання: тип операції містить "вибуття" і "Куди" містить "Списане"
Private Function Add47IsWriteOffRow(ByVal typeOp As String, ByVal toUnit As String) As Boolean
    If InStr(1, typeOp, "вибуття", vbTextCompare) > 0 Then
        If InStr(1, toUnit, "списане", vbTextCompare) > 0 Then Add47IsWriteOffRow = True
    End If
End Function

' "Куди" = Списане / Списання / Знищено ... (для режиму за серійними номерами)
Private Function Add47IsWrittenOff(ByVal s As String) As Boolean
    Dim t As String
    t = LCase$(Trim$(Replace(s, ChrW(160), " ")))
    Add47IsWrittenOff = (Left$(t, 6) = "списан") Or (Left$(t, 6) = "знищен")
End Function

' Ключ серійного номера: лише літери й цифри, нижній регістр.
' Порожній ключ = це не серійний номер (немає цифр, надто коротко/довго - ймовірно коментар).
Private Function Add47SerialKey(ByVal s As String) As String
    Dim i As Long, ch As Long
    Dim res As String
    Dim hasDigit As Boolean

    s = LCase$(s)
    For i = 1 To Len(s)
        ch = AscW(Mid$(s, i, 1))
        If ch < 0 Then ch = ch + 65536
        If ch >= 48 And ch <= 57 Then
            res = res & ChrW(ch)
            hasDigit = True
        ElseIf (ch >= 97 And ch <= 122) Or (ch >= 1072 And ch <= 1103) Or ch = 1110 Or ch = 1111 Or ch = 1108 Or ch = 1169 Then
            res = res & ChrW(ch)
        End If
    Next i
    If hasDigit And Len(res) >= 4 And Len(res) <= 30 Then Add47SerialKey = res
End Function

Private Function Add47HeaderIdx(ByRef arr() As Variant, ByVal nm As String, ByVal fallback As Long) As Long
    Dim c As Long
    For c = 1 To UBound(arr, 2)
        If StrComp(Trim$(CStr(arr(1, c))), nm, vbTextCompare) = 0 Then
            Add47HeaderIdx = c
            Exit Function
        End If
    Next c
    Add47HeaderIdx = fallback
End Function


' =====================================================================
'  Сортування: рік -> накладні за номером, акти лише за датою
' =====================================================================

' Правила (колонки: 13 = "Дата документа", 14 = "№ документа", 25 = "Тип документа"):
'  1) рік дати документа (2025 перед 2026);
'  2) накладні - за номером (далі за датою); акти та інші документи - лише за датою;
'  3) в один день акти (приходи) йдуть перед накладними.
' Щоб змішати акти (за датою) і накладні (за номером), кожній накладній призначається
' "ефективна дата" = найбільша дата серед накладних цього року з меншим або рівним номером.
' Тоді накладні зберігають порядок номерів, а акти стають на своє місце за датою.
Private Sub Add47SortRows(ByRef arr() As Variant)
    Dim n As Long, cols As Long
    Dim i As Long, j As Long, c As Long
    Dim idx() As Long
    Dim nb() As Long
    Dim nbCnt As Long
    Dim cur As Long
    Dim d As Double
    Dim cTypeDoc As Long
    Dim curYear As Long
    Dim runMax As Double
    Dim docType As String
    Dim tmp() As Variant

    n = UBound(arr, 1) - 1
    If n < 2 Then Exit Sub
    cols = UBound(arr, 2)
    cTypeDoc = Add47HeaderIdx(arr, "Тип документа", 25)

    ReDim mkYear(1 To n)
    ReDim mkNum(1 To n)
    ReDim mkRest(1 To n)
    ReDim mkDate(1 To n)
    ReDim mkEff(1 To n)
    ReDim mkCls(1 To n)
    ReDim idx(1 To n)
    ReDim nb(1 To n)

    For i = 1 To n
        d = Add47ParseDate(arr(i + 1, 13))
        mkDate(i) = d
        If d = 0 Then
            mkYear(i) = 9999            ' без дати - в кінець
        Else
            mkYear(i) = Year(d)
        End If
        Add47SplitDocNo CStr(arr(i + 1, 14)), mkNum(i), mkRest(i)

        docType = LCase$(CStr(arr(i + 1, cTypeDoc)))
        If InStr(docType, NUMBERED_DOC_WORD) > 0 And mkNum(i) < 1E+299 Then
            mkCls(i) = 1                ' накладна з номером
            nbCnt = nbCnt + 1
            nb(nbCnt) = i
        Else
            mkCls(i) = 0                ' акт / інший документ - лише за датою
        End If
        mkEff(i) = d
        idx(i) = i
    Next i

    ' --- етап 1: накладні в порядку номерів (рік, номер, решта номера, дата) ---
    For i = 2 To nbCnt
        cur = nb(i)
        j = i - 1
        Do While j >= 1
            If Add47CmpInvoice(nb(j), cur) <= 0 Then Exit Do
            nb(j + 1) = nb(j)
            j = j - 1
        Loop
        nb(j + 1) = cur
    Next i

    ' --- ефективна дата накладної = найбільша дата серед попередніх за номером (в межах року) ---
    curYear = -1
    runMax = 0
    For i = 1 To nbCnt
        If mkYear(nb(i)) <> curYear Then
            curYear = mkYear(nb(i))
            runMax = 0
        End If
        If mkDate(nb(i)) > runMax Then runMax = mkDate(nb(i))
        mkEff(nb(i)) = runMax
    Next i

    ' --- етап 2: загальне сортування (стабільне: рівні рядки лишаються в порядку бази) ---
    For i = 2 To n
        cur = idx(i)
        j = i - 1
        Do While j >= 1
            If Add47Cmp(idx(j), cur) <= 0 Then Exit Do
            idx(j + 1) = idx(j)
            j = j - 1
        Loop
        idx(j + 1) = cur
    Next i

    ReDim tmp(1 To n + 1, 1 To cols)
    For c = 1 To cols
        tmp(1, c) = arr(1, c)
    Next c
    For i = 1 To n
        For c = 1 To cols
            tmp(i + 1, c) = arr(idx(i) + 1, c)
        Next c
    Next i
    arr = tmp
End Sub

Private Function Add47Sgn(ByVal x As Double, ByVal y As Double) As Long
    If x < y Then
        Add47Sgn = -1
    ElseIf x > y Then
        Add47Sgn = 1
    End If
End Function

' порядок накладних між собою: рік -> номер -> решта номера -> дата
Private Function Add47CmpInvoice(ByVal a As Long, ByVal b As Long) As Long
    Dim s As Long

    s = Add47Sgn(mkYear(a), mkYear(b))
    If s <> 0 Then
        Add47CmpInvoice = s
        Exit Function
    End If
    s = Add47Sgn(mkNum(a), mkNum(b))
    If s <> 0 Then
        Add47CmpInvoice = s
        Exit Function
    End If
    s = StrComp(mkRest(a), mkRest(b), vbTextCompare)
    If s <> 0 Then
        Add47CmpInvoice = s
        Exit Function
    End If
    Add47CmpInvoice = Add47Sgn(mkDate(a), mkDate(b))
End Function

' загальний порядок: рік -> ефективна дата -> (акти перед накладними) -> номер (лише накладні) -> дата
Private Function Add47Cmp(ByVal a As Long, ByVal b As Long) As Long
    Dim s As Long

    s = Add47Sgn(mkYear(a), mkYear(b))
    If s <> 0 Then
        Add47Cmp = s
        Exit Function
    End If
    s = Add47Sgn(mkEff(a), mkEff(b))
    If s <> 0 Then
        Add47Cmp = s
        Exit Function
    End If
    s = Add47Sgn(mkCls(a), mkCls(b))
    If s <> 0 Then
        Add47Cmp = s
        Exit Function
    End If

    If mkCls(a) = 1 Then
        s = Add47Sgn(mkNum(a), mkNum(b))
        If s <> 0 Then
            Add47Cmp = s
            Exit Function
        End If
        s = StrComp(mkRest(a), mkRest(b), vbTextCompare)
        If s <> 0 Then
            Add47Cmp = s
            Exit Function
        End If
    End If

    Add47Cmp = Add47Sgn(mkDate(a), mkDate(b))
End Function

' "362/п/РХБЗ" -> num = 362, rest = "/п/рхбз". Без числа на початку - в кінець року.
Private Sub Add47SplitDocNo(ByVal s As String, ByRef num As Double, ByRef rest As String)
    Dim i As Long
    s = Trim$(s)
    i = 1
    Do While i <= Len(s)
        If Not (Mid$(s, i, 1) Like "#") Then Exit Do
        i = i + 1
    Loop
    If i > 1 Then
        num = val(Left$(s, i - 1))
        rest = LCase$(Mid$(s, i))
    Else
        num = 1E+300
        rest = LCase$(s)
    End If
End Sub

' Дата з тексту дд.мм.рррр або рррр-мм-дд (незалежно від регіональних налаштувань). 0 = не розпізнано.
Private Function Add47ParseDate(ByVal v As Variant) As Double
    Dim s As String
    Dim p() As String
    Dim res As Double

    If IsError(v) Then Exit Function
    If VarType(v) = vbDate Then
        Add47ParseDate = CDbl(v)
        Exit Function
    End If
    s = Trim$(CStr(v))
    If Len(s) = 0 Then Exit Function
    If InStr(s, " ") > 0 Then s = Left$(s, InStr(s, " ") - 1)

    If InStr(s, ".") > 0 Then
        p = Split(s, ".")
        If UBound(p) = 2 Then
            If IsNumeric(p(0)) And IsNumeric(p(1)) And IsNumeric(p(2)) Then
                res = Add47MakeDate(CLng(p(2)), CLng(p(1)), CLng(p(0)))
                If res > 0 Then
                    Add47ParseDate = res
                    Exit Function
                End If
            End If
        End If
    ElseIf InStr(s, "-") > 0 Then
        p = Split(s, "-")
        If UBound(p) = 2 Then
            If IsNumeric(p(0)) And IsNumeric(p(1)) And IsNumeric(p(2)) Then
                res = Add47MakeDate(CLng(p(0)), CLng(p(1)), CLng(p(2)))
                If res > 0 Then
                    Add47ParseDate = res
                    Exit Function
                End If
            End If
        End If
    End If

    If IsDate(s) Then Add47ParseDate = CDbl(CDate(s))
End Function

Private Function Add47MakeDate(ByVal y As Long, ByVal m As Long, ByVal d As Long) As Double
    If y < 100 Then y = y + 2000
    If y >= 1900 And y <= 2200 And m >= 1 And m <= 12 And d >= 1 And d <= 31 Then
        Add47MakeDate = CDbl(DateSerial(y, m, d))
    End If
End Function


' =====================================================================
'  Розпізнавання повернення з підрозділу на склад
' =====================================================================

' Ключ назви підрозділу: нижній регістр, без пробілів
Private Function Add47UnitKey(ByVal s As String) As String
    s = Replace(s, ChrW(160), "")
    s = Replace(s, " ", "")
    Add47UnitKey = LCase$(s)
End Function

' Збирає підрозділи, які мали майно: кому видавали (внутрішнє переміщення) і звідки вибували.
Private Sub Add47BuildInternalUnits(ByRef arr() As Variant)
    Dim r As Long
    Dim cType As Long, cFrom As Long, cTo As Long
    Dim t As String, k As String

    Set mInternal = CreateObject("Scripting.Dictionary")
    cType = Add47HeaderIdx(arr, "Запасне поле2", 26)
    cFrom = Add47HeaderIdx(arr, "Звідки", 8)
    cTo = Add47HeaderIdx(arr, "Куди", 9)

    For r = 2 To UBound(arr, 1)
        t = LCase$(CStr(arr(r, cType)))
        k = ""
        If InStr(t, "внутр") > 0 Then
            k = Add47UnitKey(CStr(arr(r, cTo)))
        ElseIf InStr(t, "вибуття") > 0 Then
            k = Add47UnitKey(CStr(arr(r, cFrom)))
        End If
        If Len(k) > 0 And Left$(k, 5) <> "склад" Then
            If Not mInternal.exists(k) Then mInternal.Add k, 1
        End If
    Next r
End Sub

' "надходження", де "Звідки" - це підрозділ (а не постачальник) = повернення на склад
Private Function Add47IsReturn(ByVal FromUnit As String, ByVal toUnit As String, ByVal rngUnits As Range) As Boolean
    Dim k As String

    k = Add47UnitKey(FromUnit)
    If Len(k) = 0 Then Exit Function
    If k = Add47UnitKey(toUnit) Then Exit Function
    If Left$(k, 5) = "склад" Then Exit Function

    If Not mInternal Is Nothing Then
        If mInternal.exists(k) Then
            Add47IsReturn = True
            Exit Function
        End If
    End If
    If FindValue(FromUnit, rngUnits) Then Add47IsReturn = True   ' блок підрозділу вже є на аркуші
End Function


' =====================================================================
'  Збереження додатка 47
' =====================================================================
Sub SaveAdd47()

    Dim WbNewWO As Workbook
    Dim nameAdd47 As String
    Dim nomenclature47 As String
    Dim unit47 As String
    Dim price47 As Double
    Dim formattedPrice As String
    Dim s As Long
    Dim SavePath As String
    Dim DateAdd47 As String
    Dim fileName As String
    Dim errText As String

    ActiveSheet.Copy                    ' зберігаємо як нову книгу
    Set WbNewWO = ActiveWorkbook

    nameAdd47 = WbNewWO.Worksheets(1).Range("nameAdd47")
    nomenclature47 = WbNewWO.Worksheets(1).Range("nomenclature47")
    unit47 = WbNewWO.Worksheets(1).Range("unit47")
    price47 = WbNewWO.Worksheets(1).Range("price47")

    formattedPrice = Format(price47, "#,##0.0000")

    ' видаляємо кнопки/фігури (у зворотному порядку, інакше частина пропускається)
    For s = WbNewWO.Worksheets(1).Shapes.count To 1 Step -1
        WbNewWO.Worksheets(1).Shapes(s).Delete
    Next s

    ' шлях до файла зберігання
    SavePath = WbNewWO.Worksheets(1).Range("PathToSaveAdd47")
    If Right$(SavePath, 1) = "\" Or Right$(SavePath, 1) = "/" Then
        SavePath = Left$(SavePath, Len(SavePath) - 1)
    End If

    DateAdd47 = Format(Date, "dd.mm.yyyy")

    If Not Add47FolderExists(SavePath) Then
        With Application.FileDialog(msoFileDialogFolderPicker)
            .Title = "Виберіть папку для збереження"
            .AllowMultiSelect = False
            If .show = -1 Then
                SavePath = .SelectedItems(1)
                ' записуємо новий шлях у клітинку
                ThisWorkbook.Worksheets("Додаток 47").Range("PathToSaveAdd47").value = SavePath
            Else
                MsgBox "Збереження скасовано", vbExclamation
                Application.DisplayAlerts = False
                WbNewWO.Close SaveChanges:=False
                Application.DisplayAlerts = True
                Exit Sub
            End If
        End With
    End If

    ' складаємо назву файла
    fileName = SavePath & Application.PathSeparator & _
               CleanFileName(DateAdd47 & " д47 " & nameAdd47 & " " & nomenclature47 & " " & unit47 & " " & formattedPrice & ".xlsx")

    ' зберігаємо файл
    Application.DisplayAlerts = False
    On Error GoTo WrongPath
    WbNewWO.SaveAs fileName:=fileName, FileFormat:=xlOpenXMLWorkbook
    On Error GoTo 0
    Application.DisplayAlerts = True
    Exit Sub

WrongPath:
    errText = Err.Description
    Application.DisplayAlerts = True
    MsgBox "Ймовірно, не правильно вказана адреса збереження файла, виправте адресу збереження файлу (на зразок ""D:\База обліку майна ЗСУ 2024"" *без косої лінії вкінці) і спробуйте сформувати додаток ще раз!" & _
           vbCrLf & vbCrLf & "Деталі: " & errText, vbCritical, "Не правильний шлях збереження"
End Sub


