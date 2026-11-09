# 05 — Модулі розрахунку залишків ТМЦ та аналітичних звітів

У цьому документі наведено вичерпний опис та вихідний код 6 модулів розрахунку залишків майна:
1. **Residuse.bas** — зведені залишки по всіх підрозділах (PivotTable + SUMIFS)
2. **Reiduse2.bas** — оптимізований детальний розрахунок залишків (Dictionary в ОЗП)
3. **ResiduseOneTroop.bas** — залишки конкретного підрозділу (версія 1.0)
4. **ResiduseOneTroop2.bas** — розширені залишки підрозділу з фільтром за датами (версія 2.0)
5. **ResiduseOneGood.bas** — залишки одного обраного товару по всіх підрозділах
6. **ResiduseOnePeople.bas** — залишки майна за матеріально-відповідальною особою (МВО)

## Порівняльна таблиця модулів залишків

| Модуль | Аркуш виводу | Метод розрахунку | Фільтр дат | Фільтр МВО | Збереження у файл |
|--------|--------------|------------------|------------|------------|-------------------|
| `Residuse.bas` | «Залишки по підрозділах» | PivotTable + SUMIFS формули | Ні | Ні | Ні |
| `Reiduse2.bas` | «Залишки по підрозділах2»| В пам'яті через Dictionary | Ні | Ні | Ні |
| `ResiduseOneTroop.bas` | «Залишок по одному» | Dictionary (по підрозділу) | Ні | Ні | Так (`SaveResiduseOneUnit`) |
| `ResiduseOneTroop2.bas`| «Залишок по одному 2.0»| Dictionary + Date Interval | Так | Так | Так (масовий експорт) |
| `ResiduseOneGood.bas` | «Залишки по товару» | Перебір бази AllData | Ні | Ні | Ні |
| `ResiduseOnePeople.bas`| «Залишок по особі» | Dictionary (по МВО) | Ні | Так | Так (`SaveResiduseOnePeople`) |

---

## 1. Residuse.bas — Зведений звіт залишків (v1)
### 1.1. Опис та алгоритм
Формує зведену матрицю на аркуші «Залишки по підрозділах». Спочатку оновлює зведену таблицю `ptUnitResiduse`, а потім динамічно генерує формули `SUMIFS` для обчислення чистих залишків (`Прихід - Вибуття`) по кожному підрозділу. Для візуалізації прогресу використовується форма `ProgressBar.frm`.

### 1.2. Вихідний код Residuse.bas
```vba
Attribute VB_Name = "Residuse"
Sub UpdateResiduse()

Dim wbAllData As Workbook
Dim wsAllData As Worksheet
Dim wsResiduse As Worksheet
Dim tbAllData As ListObject

    Set wbAllData = ThisWorkbook
    Set wsAllData = wbAllData.Worksheets("База")
    Set tbAllData = wsAllData.ListObjects("AllData")
    Set wsResiduse = wbAllData.Worksheets("Залишки по підрозділах")
 
    Dim dict As Object
    Dim cell As Range
    Dim colFrom As Range, colTo As Range
    Dim i As Long
    
   wsResiduse.Range("D6", wsResiduse.Cells(wsResiduse.Rows.count, wsResiduse.Columns.count)).ClearContents
    
    ' Встановіть посилання на стовпці таблиці
    Set colFrom = tbAllData.ListColumns(8).DataBodyRange ' Підрозділ звідки
    Set colTo = tbAllData.ListColumns(9).DataBodyRange 'підрозділ куди
    
    ' Створіть словник для зберігання унікальних значень
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Витягніть унікальні значення з першого стовпця
    For Each cell In colFrom
        If Not dict.exists(cell.value) Then
            dict.Add cell.value, Nothing
        End If
    Next cell
    
    ' Витягніть унікальні значення з другого стовпця
    For Each cell In colTo
        If Not dict.exists(cell.value) Then
            dict.Add cell.value, Nothing
        End If
    Next cell
    
    Dim arr() As Variant
    Dim j As Long, temp As Variant
    
    
     ' Перетворіть ключі словника на масив
    arr = dict.Keys
    
    ' Виконайте сортування масиву
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) > arr(j) Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
    Dim k  As Long
    ' Запишіть унікальні значення в рядок на цільовому листі
    For i = LBound(arr) To UBound(arr)
   k = i + 3
    
            wsResiduse.Cells(6, k + 1).value = arr(i)
    Next i
    
    Dim ptUnitResiduse As PivotTable
    Set ptUnitResiduse = wsResiduse.PivotTables("ptUnitResiduse")
    ptUnitResiduse.RefreshTable
    
    
     Dim totalSum As Double
    Dim unitName As String
    Dim destination As String
    
      Dim s As Long, l As Long
    Dim lastRow As Long
    Dim lastCol As Long
    
    lastRow = wsResiduse.Cells(wsResiduse.Rows.count, "A").End(xlUp).Row
    lastCol = wsResiduse.Cells(6, wsResiduse.Columns.count).End(xlToLeft).Column
    
    For s = 8 To lastRow
        unitName = wsResiduse.Cells(s, 1).value ' Значення з колонки A
        
        For l = 4 To lastCol ' Починаємо з  (стовпець 4)
            destination = wsResiduse.Cells(6, l).value ' Значення з рядка 6
            
            ' Розрахунок підсумкової суми для комбінації unitName і destination
            totalSum = WorksheetFunction.SumIfs(tbAllData.ListColumns("Надійшло").DataBodyRange, _
                                                tbAllData.ListColumns("Куди").DataBodyRange, destination, _
                                                tbAllData.ListColumns("Найменування").DataBodyRange, unitName) _
                      + WorksheetFunction.SumIfs(tbAllData.ListColumns("Вибуло").DataBodyRange, _
                                                tbAllData.ListColumns("Куди").DataBodyRange, destination, _
                                                tbAllData.ListColumns("Найменування").DataBodyRange, unitName) _
                      - WorksheetFunction.SumIfs(tbAllData.ListColumns("Вибуло").DataBodyRange, _
                                                tbAllData.ListColumns("Звідки").DataBodyRange, destination, _
                                                tbAllData.ListColumns("Найменування").DataBodyRange, unitName)
            
            ' Виведення результату у відповідну комірку на перетині рядка і колонки
            wsResiduse.Cells(s, l).value = totalSum
        Next l
        Call ProgressBar.Progress(s / lastRow * 100)
        
    Next s
    


End Sub

```

---

## 2. Reiduse2.bas — Високопродуктивний зведений розрахунок (v2)
### 2.1. Опис та алгоритм
Модуль вирішує проблему повільної роботи Excel при великих обсягах даних. Замість формул `SUMIFS`, процедура `UpdateResiduseDetailedAll`:
1. Зчитує всю таблицю `AllData` в оперативну пам'ять (2D-масив).
2. Будує вкладені словники `Scripting.Dictionary`: `dictUnits(Підрозділ) -> dictGoods(Товар)`.
3. Миттєво виводить готовий результат масивом на аркуш «Залишки по підрозділах2» та накладає автофільтр.

### 2.2. Вихідний код Reiduse2.bas
```vba
Attribute VB_Name = "Reiduse2"
Option Explicit

'==============================================================
' Допоміжні функції
'==============================================================
Private Function SafeNum(ByVal v As Variant) As Double
    If IsError(v) Then Exit Function
    If IsNumeric(v) Then SafeNum = CDbl(v)
End Function

Private Function SafeStr(ByVal v As Variant) As String
    If IsError(v) Then Exit Function
    If IsNull(v) Then Exit Function
    SafeStr = CStr(v)
End Function

'==============================================================
' Основна процедура
'==============================================================
Sub UpdateResiduseDetailedAll()
    Dim wb As Workbook
    Dim wsData As Worksheet, wsOutput As Worksheet
    Dim tbAllData As ListObject
    Dim dictUnits As Object, d As Object
    Dim data As Variant
    Dim dataCount As Long, i As Long, k As Long, n As Long
    Dim unitKey As Variant, allKeys As Variant
    Dim parts() As String
    Dim out() As Variant
    Dim estimatedMinutes As Double
    Dim confirm As VbMsgBoxResult
    Dim lastRowRes As Long

    ' індекси колонок
    Dim cFrom As Long, cTo As Long, cIn As Long, cOut As Long
    Dim cCode As Long, cName As Long, cPrice As Long, cUnit As Long
    Dim cCat As Long, cCat3 As Long, cNote As Long, c11 As Long, c12 As Long

    ' змінні для рядка
    Dim fromVal As String, toVal As String, cat As String
    Dim prefix As String, suffix As String, key As String
    Dim vIn As Double, vOut As Double, v As Double

    Set wb = ThisWorkbook
    Set wsData = wb.Worksheets("База")
    Set wsOutput = wb.Worksheets("Залишки по підрозділах2")
    Set tbAllData = wsData.ListObjects("AllData")

    If tbAllData.DataBodyRange Is Nothing Then
        MsgBox "У таблиці AllData немає даних.", vbExclamation
        Exit Sub
    End If

    dataCount = tbAllData.ListRows.count

    '--- ПРОГНОЗ ЧАСУ ---
    estimatedMinutes = dataCount * 0.035 / 60
    If estimatedMinutes < 1 Then
        confirm = MsgBox("Оновлення залишків займе: до 1 хвилини." & vbCrLf & _
                         "Бажаєте продовжити?", vbYesNo + vbQuestion)
    ElseIf estimatedMinutes > 60 Then
        confirm = MsgBox("Оновлення залишків займе: більше години." & vbCrLf & _
                         "Бажаєте продовжити?", vbYesNo + vbQuestion)
    Else
        confirm = MsgBox("Оновлення залишків займе: приблизно " & _
                         WorksheetFunction.RoundUp(estimatedMinutes, 0) & " хвилин(и)." & vbCrLf & _
                         "Бажаєте продовжити?", vbYesNo + vbQuestion)
    End If
    If confirm = vbNo Then Exit Sub

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    '--- Індекси колонок (один раз) ---
    With tbAllData
        cFrom = .ListColumns("Звідки").Index
        cTo = .ListColumns("Куди").Index
        cIn = .ListColumns("Надійшло").Index
        cOut = .ListColumns("Вибуло").Index
        cCode = .ListColumns("Код номенклатури").Index
        cName = .ListColumns("Найменування").Index
        cPrice = .ListColumns("Ціна").Index
        cUnit = .ListColumns("Одиниця виміру").Index
        cCat = .ListColumns("Категорія").Index
        cCat3 = .ListColumns("Запасне поле3").Index
        cNote = .ListColumns("Примітка").Index
        c11 = .ListColumns("Поле 11").Index
        c12 = .ListColumns("Поле 12").Index
    End With

    '--- Читаємо всю базу в масив ---
    data = tbAllData.DataBodyRange.value
    If dataCount = 1 Then
        ' для одного рядка .Value повертає 2D масив теж, але перестрахуємось
        If Not IsArray(data) Then
            MsgBox "Не вдалося прочитати дані таблиці.", vbExclamation
            GoTo CleanExit
        End If
    End If

    '--- Список підрозділів (спочатку "Звідки", потім "Куди") ---
    Set dictUnits = CreateObject("Scripting.Dictionary")
    For i = 1 To dataCount
        fromVal = SafeStr(data(i, cFrom))
        If fromVal <> "" Then
            If Not dictUnits.exists(fromVal) Then dictUnits.Add fromVal, CreateObject("Scripting.Dictionary")
        End If
    Next i
    For i = 1 To dataCount
        toVal = SafeStr(data(i, cTo))
        If toVal <> "" Then
            If Not dictUnits.exists(toVal) Then dictUnits.Add toVal, CreateObject("Scripting.Dictionary")
        End If
    Next i

    '--- Один прохід по базі ---
    For i = 1 To dataCount
        fromVal = SafeStr(data(i, cFrom))
        toVal = SafeStr(data(i, cTo))

        If fromVal <> "" Or toVal <> "" Then
            vIn = SafeNum(data(i, cIn))
            vOut = SafeNum(data(i, cOut))

            prefix = SafeStr(data(i, cCode)) & "||" & _
                     SafeStr(data(i, cName)) & "||" & _
                     SafeStr(data(i, cPrice)) & "||" & _
                     SafeStr(data(i, cUnit)) & "||"
            suffix = "||" & SafeStr(data(i, cNote)) & "||" & _
                     SafeStr(data(i, c11)) & "||" & _
                     SafeStr(data(i, c12))

            ' Надходження в підрозділ "Куди"
            If toVal <> "" Then
                cat = SafeStr(data(i, cCat3))
                key = prefix & cat & suffix
                Set d = dictUnits(toVal)
                If Not d.exists(key) Then d.Add key, 0#
                d(key) = d(key) + vIn
                d(key) = d(key) + vOut
            End If

            ' Вибуття з підрозділу "Звідки"
            If fromVal <> "" Then
                If fromVal = toVal Then
                    cat = SafeStr(data(i, cCat3))
                Else
                    cat = SafeStr(data(i, cCat))
                End If
                key = prefix & cat & suffix
                Set d = dictUnits(fromVal)
                If Not d.exists(key) Then d.Add key, 0#
                d(key) = d(key) - vOut
            End If
        End If

        If i Mod 500 = 0 Then
            Call ProgressBar.Progress(i / dataCount * 90)
        End If
    Next i

    '--- Формуємо масив виводу ---
    ReDim out(1 To 2 * dataCount + 10, 1 To 14)
    n = 0

    For Each unitKey In dictUnits.Keys
        Set d = dictUnits(unitKey)
        allKeys = d.Keys
        For k = 0 To d.count - 1
            v = d(allKeys(k))
            If Round(v, 6) <> 0 Then
                parts = Split(allKeys(k), "||")
                n = n + 1
                out(n, 1) = n
                out(n, 2) = parts(1)
                out(n, 3) = SafeNum(parts(2))
                out(n, 4) = v
                out(n, 5) = parts(5)
                out(n, 6) = parts(4)
                out(n, 7) = parts(3)
                out(n, 8) = parts(0)
                ' колонка 9 (Поле 10) - не використовується
                out(n, 10) = parts(6)
                out(n, 11) = parts(7)
                out(n, 12) = CStr(unitKey)
                out(n, 13) = SafeNum(parts(2)) * v
                out(n, 14) = parts(1) & ", ц. " & parts(2) & ", зал: " & _
                             v & " " & parts(3) & " " & unitKey & ", пр " & parts(5) & _
                             ", кат " & parts(4) & ", к.н " & parts(0)
            End If
        Next k
    Next unitKey

    '--- Запис на аркуш ---
    wsOutput.Cells.ClearContents

    With wsOutput
        .Cells(2, 1).value = "№"
        .Cells(2, 2).value = "Найменування"
        .Cells(2, 3).value = "Ціна"
        .Cells(2, 4).value = "Залишок"
        .Cells(2, 5).value = "Примітка"
        .Cells(2, 6).value = "Категорія"
        .Cells(2, 7).value = "Одиниця виміру"
        .Cells(2, 8).value = "Код номенклатури"
        .Cells(2, 9).value = "Поле 10"
        .Cells(2, 10).value = "Поле 11"
        .Cells(2, 11).value = "Поле 12"
        .Cells(2, 12).value = "Підрозділ"
        .Cells(2, 13).value = "Сума (ціна х кількість)"
    End With

    If n > 0 Then
        wsOutput.Range("A3").Resize(n, 14).value = out
    End If

    '--- Фільтри ---
    If wsOutput.AutoFilterMode Then wsOutput.AutoFilterMode = False
    wsOutput.Range("A2").CurrentRegion.AutoFilter

    '--- Підсумок ---
    lastRowRes = wsOutput.Cells.Find(What:="*", _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlPrevious).Row

    On Error Resume Next
    wsOutput.Range("RngSumRes").Formula = "=SUBTOTAL(9,M3:M" & lastRowRes & ")"
    wsOutput.Range("RngTotalRes").value = "Всього:"
    If Err.Number <> 0 Then
        MsgBox "Іменовані діапазони RngSumRes / RngTotalRes не знайдено — підсумок пропущено.", vbExclamation
        Err.Clear
    End If
    On Error GoTo ErrHandler

    Call ProgressBar.Progress(100)

CleanExit:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    If Err.Number = 0 Then MsgBox "Оновлення завершено!"
    Exit Sub

ErrHandler:
    Dim msg As String
    msg = "Помилка " & Err.Number & ": " & Err.Description
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox msg, vbCritical, "UpdateResiduseDetailedAll"
End Sub


```

---

## 3. ResiduseOneTroop.bas — Залишки одного підрозділу (v1)
### 3.1. Опис та алгоритм
Дозволяє вибрати підрозділ через пошукову форму `UserForm1`, розрахувати його поточні залишки на аркуші «Залишок по одному», а процедура `SaveResiduseOneUnit` створює чистий файл без макросів у поточному каталозі.

### 3.2. Вихідний код ResiduseOneTroop.bas
```vba
Attribute VB_Name = "ResiduseOneTroop"
Sub UpdateResiduseDetailed()

Dim wbAllData As Workbook
Dim wsAllData As Worksheet
Dim wsResiduse As Worksheet
Dim tbAllData As ListObject

Set wbAllData = ThisWorkbook
Set wsAllData = wbAllData.Worksheets("База")
Set tbAllData = wsAllData.ListObjects("AllData")
Set wsResiduse = wbAllData.Worksheets("Довідка по підрозділу") ' ОНОВЛЕНО


Dim frm As UserForm1
Set frm = New UserForm1

Dim colFrom As Range, colTo As Range, cell As Range
Dim dictUnits As Object
Set dictUnits = CreateObject("Scripting.Dictionary")

Set colFrom = tbAllData.ListColumns("Звідки").DataBodyRange
Set colTo = tbAllData.ListColumns("Куди").DataBodyRange

' Формуємо словник унікальних підрозділів
For Each cell In colFrom
    If Not dictUnits.exists(cell.value) And Not IsEmpty(cell.value) Then
        dictUnits.Add cell.value, Nothing
    End If
Next cell

For Each cell In colTo
    If Not dictUnits.exists(cell.value) And Not IsEmpty(cell.value) Then
        dictUnits.Add cell.value, Nothing
    End If
Next cell

' Виводимо список підрозділів у колонку A, починаючи з A7
Dim n As Long
n = 7
wsResiduse.Range("A7:A" & wsResiduse.Rows.count).ClearContents ' Очистити перед вставкою

Dim unitKey As Variant
For Each unitKey In dictUnits.Keys
    wsResiduse.Cells(n, 1).value = unitKey
    n = n + 1
Next unitKey

' Зчитати назад у масив
Dim lastRow As Long
lastRow = wsResiduse.Cells(wsResiduse.Rows.count, 1).End(xlUp).Row

'Dim unitArray As Variant
'unitArray = wsResiduse.Range("A7:A" & lastRow).value

' Передати в ListBox
'If IsArray(unitArray) Then
    frm.Caption = "Виберіть підрозділ"
    frm.ListData = wsResiduse.Range("A7:A" & lastRow)
'Else
'    frm.ListBox1.AddItem unitArray ' якщо один елемент
'End If

' Показати форму
frm.show

' Отримати вибране
If frm.ListBox1.ListIndex <> -1 Then
    selectedUnit = frm.ListBox1.value
Else
    MsgBox "Підрозділ не вибрано!", vbExclamation
    Exit Sub
End If



 wsResiduse.Range("selectedUnit").value = selectedUnit ' Вибраний підрозділ
 wsResiduse.Range("A7:A" & lastRow).ClearContents

Dim dict As Object
Set dict = CreateObject("Scripting.Dictionary")

Dim i As Long
Dim key As Variant
Dim result As Double

' Очистити старі дані
wsResiduse.Range("b8:i" & wsResiduse.Rows.count).ClearContents

Dim dataCount As Long
dataCount = tbAllData.ListRows.count

' Формуємо унікальні ключі на основі декількох полів
For i = 1 To dataCount
    If tbAllData.DataBodyRange(i, tbAllData.ListColumns("Куди").Index).value = selectedUnit _
    Or tbAllData.DataBodyRange(i, tbAllData.ListColumns("Звідки").Index).value = selectedUnit Then
    
      Dim dynamicCategory As String

' Вибираємо значення для категорії в залежності від напрямку
If tbAllData.DataBodyRange(i, tbAllData.ListColumns("Куди").Index).value = selectedUnit Then
    dynamicCategory = tbAllData.DataBodyRange(i, tbAllData.ListColumns("Запасне поле3").Index).value
Else
    dynamicCategory = tbAllData.DataBodyRange(i, tbAllData.ListColumns("Категорія").Index).value
End If

key = tbAllData.DataBodyRange(i, tbAllData.ListColumns("Код номенклатури").Index).value & "||" & _
      tbAllData.DataBodyRange(i, tbAllData.ListColumns("Найменування").Index).value & "||" & _
      tbAllData.DataBodyRange(i, tbAllData.ListColumns("Ціна").Index).value & "||" & _
      tbAllData.DataBodyRange(i, tbAllData.ListColumns("Одиниця виміру").Index).value & "||" & _
      dynamicCategory & "||" & _
      tbAllData.DataBodyRange(i, tbAllData.ListColumns("Примітка").Index).value


        If Not dict.exists(key) Then
            dict.Add key, 0
        End If
        
        ' Додаємо/віднімаємо залишок
        If tbAllData.DataBodyRange(i, tbAllData.ListColumns("Куди").Index).value = selectedUnit Then
            dict(key) = dict(key) + tbAllData.DataBodyRange(i, tbAllData.ListColumns("Надійшло").Index).value
            dict(key) = dict(key) + tbAllData.DataBodyRange(i, tbAllData.ListColumns("Вибуло").Index).value
        End If
        
        If tbAllData.DataBodyRange(i, tbAllData.ListColumns("Звідки").Index).value = selectedUnit Then
            dict(key) = dict(key) - tbAllData.DataBodyRange(i, tbAllData.ListColumns("Вибуло").Index).value
        End If
    End If
Next i

' Виводимо результати
Dim rowIndex As Long
rowIndex = 8
Dim parts() As String

Dim k As Long
Dim totalKeys As Long
Dim allKeys As Variant

allKeys = dict.Keys
totalKeys = dict.count
d = 0
For k = 1 To totalKeys
    If dict(allKeys(k - 1)) <> 0 Then
        parts = Split(allKeys(k - 1), "||")
        d = d + 1
        wsResiduse.Cells(rowIndex, 2).value = d
        wsResiduse.Cells(rowIndex, 3).value = parts(1) ' Найменування
        wsResiduse.Cells(rowIndex, 4).value = CDbl(parts(2)) ' Ціна
        wsResiduse.Cells(rowIndex, 5).value = CDbl(dict(allKeys(k - 1))) ' Залишок
        wsResiduse.Cells(rowIndex, 6).value = parts(5) ' Примітка
        wsResiduse.Cells(rowIndex, 7).value = parts(4) ' Категорія
        wsResiduse.Cells(rowIndex, 8).value = parts(3) ' Одиниця виміру
        wsResiduse.Cells(rowIndex, 9).value = parts(0) ' Код номенклатури
        rowIndex = rowIndex + 1
    End If
    
    Call ProgressBar.Progress(k / totalKeys * 100)
Next k

If wsResiduse.AutoFilterMode Then wsResiduse.AutoFilterMode = False
    wsResiduse.Range("B7").CurrentRegion.AutoFilter


End Sub




Sub SaveResiduseOneUnit()

        ThisWorkbook.Sheets("Довідка по підрозділу").Select 'копіюємо перший лист з нарядом
        ThisWorkbook.Sheets("Довідка по підрозділу").Copy 'зберігаємо як нову книгу
   
     Dim WbNew As Workbook
     
        Set WbNew = ActiveWorkbook
   
    Dim shp As Shape
    For Each shp In WbNew.Worksheets(1).Shapes
        shp.Delete
    Next shp
    
'    WbNew.Worksheets(1).Columns("N").ClearContents
        
      Dim fileName As String
     fileName = ", залишки МЦ станом на " & Date
        'складаємо назву файла
        
        WbNew.Worksheets(1).Range("A3") = fileName
        
        'зберігаємо файл
'        On Error GoTo WrongPath
         WbNew.SaveAs fileName:=ThisWorkbook.Path & Application.PathSeparator & _
         CleanFileName(WbNew.Worksheets(1).Range("selectedUnit").value & _
         fileName & ".xlsx"), FileFormat:=xlOpenXMLWorkbook
'     On Error GoTo 0
'      Exit Sub
'WrongPath:
     MsgBox "Файл з залишками матеріальних цінностей збережений і тій же папці, де файл бази даних", vbInformation, ThisWorkbook.Path
'     Exit Sub
     
End Sub



```

---

## 4. ResiduseOneTroop2.bas — Розширені залишки підрозділу (v2.0)
### 4.1. Опис та алгоритм
Найбільш функціональний звіт системи. Підтримує вибір часового інтервалу (початкова та кінцева дати через функцію `AskResidusePeriod`), облік розбивки по матеріально-відповідальних особах (МВО), а також режим масового формування залишків одразу для всіх підрозділів частини (`buildAllUnits`).

### 4.2. Вихідний код ResiduseOneTroop2.bas
```vba
Attribute VB_Name = "ResiduseOneTroop2"
Option Explicit

Private Const RES_SEP As String = "||"

'===============================================================================
' МОДУЛЬ: ResiduseOneTroop
'
' ЛОГІКА ТИПІВ ОПЕРАЦІЙ БЕРЕТЬСЯ З КОЛОНКИ "Запасне поле2":
'
' 1. "надходження"
'    - впливає ТІЛЬКИ на підрозділ у полі "Куди";
'    - кількість береться з "Надійшло".
'
' 2. "вибуття"
'    - впливає ТІЛЬКИ на підрозділ у полі "Звідки";
'    - кількість береться з "Вибуло".
'
' 3. "внутрішнє переміщення"
'    - віднімається у підрозділу "Звідки";
'    - додається підрозділу "Куди";
'    - у поточній структурі бази кількість зберігається у "Вибуло";
'    - для сумісності зі старими даними, якщо "Вибуло" = 0,
'      використовується "Надійшло".
'
' Це відновлює логіку першого UpdateResiduseDetailed:
' звичайне "надходження" не повинно автоматично списувати майно з поля "Звідки".
'===============================================================================


'===============================================================================
' ОСНОВНА КНОПКА "ОНОВИТИ" НА АРКУШІ "Довідка по підрозділу 2.0"
'
' Формує довідку:
' B  №
' C  Підрозділ
' D  МВО
' E  Найменування
' F  Ціна
' G  Надійшло за період
' H  Вибуло за період
' I  Залишок на кінець періоду
' J  Примітка
' K  Категорія
' L  Одиниця виміру
' M  Код номенклатури
' N  Залишок на початок періоду (службова прихована колонка)
'
' Якщо форму вибору підрозділу закрити без вибору, пропонується
' сформувати довідку одразу за всіма підрозділами.
'===============================================================================
Public Sub UpdateResiduseDetailed2()

    Dim wb As Workbook
    Dim wsData As Worksheet
    Dim tbAllData As ListObject
    Dim frm As UserForm1
    Dim dictUnits As Object
    Dim unitArray() As Variant
    Dim cell As Range
    Dim unitKey As Variant
    Dim selectedUnit As String
    Dim buildAllUnits As Boolean
    Dim answer As VbMsgBoxResult
    Dim startDate As Date
    Dim endDate As Date
    Dim i As Long

    Set wb = ThisWorkbook
    Set wsData = wb.Worksheets("База")
    Set tbAllData = wsData.ListObjects("AllData")

    Set dictUnits = CreateObject("Scripting.Dictionary")

    For Each cell In tbAllData.ListColumns("Звідки").DataBodyRange
        If Len(Trim$(CStr(cell.value))) > 0 Then
            If Not dictUnits.exists(CStr(cell.value)) Then
                dictUnits.Add CStr(cell.value), Nothing
            End If
        End If
    Next cell

    For Each cell In tbAllData.ListColumns("Куди").DataBodyRange
        If Len(Trim$(CStr(cell.value))) > 0 Then
            If Not dictUnits.exists(CStr(cell.value)) Then
                dictUnits.Add CStr(cell.value), Nothing
            End If
        End If
    Next cell

    If dictUnits.count = 0 Then
        MsgBox "У базі не знайдено підрозділів.", vbExclamation
        Exit Sub
    End If

    ReDim unitArray(1 To dictUnits.count, 1 To 1)

    i = 1
    For Each unitKey In dictUnits.Keys
        unitArray(i, 1) = unitKey
        i = i + 1
    Next unitKey

    Set frm = New UserForm1
    frm.Caption = "Виберіть підрозділ"
    frm.ListData = unitArray
    frm.show

    If frm.ListBox1.ListIndex <> -1 Then
        selectedUnit = CStr(frm.ListBox1.value)
        buildAllUnits = False
        Unload frm
    Else
        Unload frm

        answer = MsgBox( _
            "Підрозділ не вибрано." & vbCrLf & vbCrLf & _
            "Сформувати довідку за ВСІМА підрозділами?" & vbCrLf & vbCrLf & _
            "Так — сформувати за всіма підрозділами." & vbCrLf & _
            "Ні — закрити без формування.", _
            vbYesNo + vbQuestion, _
            "Формування довідки")

        If answer = vbNo Then Exit Sub

        buildAllUnits = True
        selectedUnit = vbNullString
    End If

    If Not AskResidusePeriod(startDate, endDate) Then Exit Sub

    BuildResiduseDetailedReport selectedUnit, startDate, endDate, buildAllUnits

    If buildAllUnits Then
        MsgBox "Довідку сформовано за всіма підрозділами.", vbInformation
    Else
        MsgBox "Довідку сформовано за підрозділом: " & selectedUnit, vbInformation
    End If

End Sub


'===============================================================================
' ПУБЛІЧНИЙ РОБОЧИЙ МЕТОД.
'
' ЙОГО ВИКОРИСТОВУЄ CreateAddOneForOneTroop / CreateAddOneForAllTroops.
'
' buildAllUnits = False:
'   формує тільки selectedUnit.
'
' buildAllUnits = True:
'   формує всі підрозділи ОДНИМ проходом по AllData.
'===============================================================================
Public Sub BuildResiduseDetailedReport( _
    ByVal selectedUnit As String, _
    ByVal startDate As Date, _
    ByVal endDate As Date, _
    ByVal buildAllUnits As Boolean)

    Dim wb As Workbook
    Dim wsData As Worksheet
    Dim wsOut As Worksheet
    Dim tbAllData As ListObject

    Dim dict As Object
    Dim i As Long
    Dim dataCount As Long
    Dim rowIndex As Long
    Dim counter As Long

    Dim rowDateValue As Variant
    Dim rowDate As Date

    Dim fromVal As String
    Dim toVal As String
    Dim operationType As String

    Dim mvoFrom As String
    Dim mvoTo As String

    Dim catFrom As String
    Dim catTo As String

    Dim itemName As String
    Dim itemCode As String
    Dim itemUnit As String
    Dim itemNote As String
    Dim itemPrice As Variant

    Dim qIn As Double
    Dim qOut As Double
    Dim movementQty As Double

    Dim key As String
    Dim itemKey As Variant
    Dim a As Variant
    Dim parts() As String
    Dim endBalance As Double

    If endDate < startDate Then
        Err.Raise vbObjectError + 701, "BuildResiduseDetailedReport", _
                  "Дата кінця періоду не може бути меншою за дату початку."
    End If

    Set wb = ThisWorkbook
    Set wsData = wb.Worksheets("База")
    Set wsOut = wb.Worksheets("Довідка по підрозділу 2.0")
    Set tbAllData = wsData.ListObjects("AllData")
    Set dict = CreateObject("Scripting.Dictionary")

    Application.ScreenUpdating = False

    '---------------------------------------------------------------------------
    ' Очищення і структура довідки
    '---------------------------------------------------------------------------
    wsOut.Range("B8:N" & wsOut.Rows.count).ClearContents

    With wsOut
        .Cells(7, 2).value = "№"
        .Cells(7, 3).value = "Підрозділ"
        .Cells(7, 4).value = "МВО"
        .Cells(7, 5).value = "Найменування"
        .Cells(7, 6).value = "Ціна"
        .Cells(7, 7).value = "Надійшло"
        .Cells(7, 8).value = "Вибуло"
        .Cells(7, 9).value = "Залишок"
        .Cells(7, 10).value = "Примітка"
        .Cells(7, 11).value = "Категорія"
        .Cells(7, 12).value = "Одиниця виміру"
        .Cells(7, 13).value = "Код номенклатури"
        .Cells(7, 14).value = "Залишок на початок"
        .Columns("N").Hidden = True
    End With

    On Error Resume Next
    If buildAllUnits Then
        wsOut.Range("selectedUnit").value = "Всі підрозділи"
    Else
        wsOut.Range("selectedUnit").value = selectedUnit
    End If
    On Error GoTo 0

    dataCount = tbAllData.ListRows.count

    '---------------------------------------------------------------------------
    ' ОДИН ПРОХІД ПО БАЗІ.
    '
    ' КРИТИЧНО:
    ' напрямок руху визначається НЕ лише по "Звідки"/"Куди",
    ' а спочатку по "Запасне поле2" = тип операції.
    '---------------------------------------------------------------------------
    For i = 1 To dataCount

        rowDateValue = tbAllData.DataBodyRange(i, _
                       tbAllData.ListColumns("Дата документа").Index).value

        If IsDate(rowDateValue) Then

            rowDate = CDate(rowDateValue)

            'Операції після кінця звітного періоду не впливають на звіт.
            If rowDate <= endDate Then

                fromVal = Trim$(CStr(tbAllData.DataBodyRange(i, _
                          tbAllData.ListColumns("Звідки").Index).value))

                toVal = Trim$(CStr(tbAllData.DataBodyRange(i, _
                        tbAllData.ListColumns("Куди").Index).value))

                operationType = LCase$(Trim$(CStr(tbAllData.DataBodyRange(i, _
                                tbAllData.ListColumns("Запасне поле2").Index).value)))

                qIn = SafeDoubleResiduse(tbAllData.DataBodyRange(i, _
                      tbAllData.ListColumns("Надійшло").Index).value)

                qOut = SafeDoubleResiduse(tbAllData.DataBodyRange(i, _
                       tbAllData.ListColumns("Вибуло").Index).value)

                itemName = CStr(tbAllData.DataBodyRange(i, _
                           tbAllData.ListColumns("Найменування").Index).value)

                itemCode = CStr(tbAllData.DataBodyRange(i, _
                           tbAllData.ListColumns("Код номенклатури").Index).value)

                itemUnit = CStr(tbAllData.DataBodyRange(i, _
                           tbAllData.ListColumns("Одиниця виміру").Index).value)

                itemPrice = tbAllData.DataBodyRange(i, _
                            tbAllData.ListColumns("Ціна").Index).value

                itemNote = CStr(tbAllData.DataBodyRange(i, _
                           tbAllData.ListColumns("Примітка").Index).value)

                Select Case operationType

                    '===========================================================
                    ' 1. НАДХОДЖЕННЯ
                    '
                    ' Додаємо ТІЛЬКИ підрозділу "Куди".
                    ' Поле "Звідки" у такому рядку НЕ є підставою для списання.
                    '===========================================================
                    Case "надходження"

                        movementQty = qIn

                        If movementQty <> 0 And Len(toVal) > 0 Then

                            If buildAllUnits Or toVal = selectedUnit Then

                                mvoTo = CStr(tbAllData.DataBodyRange(i, _
                                        tbAllData.ListColumns("МВО куди").Index).value)

                                catTo = CStr(tbAllData.DataBodyRange(i, _
                                        tbAllData.ListColumns("Запасне поле3").Index).value)

                                key = MakeResiduseKey( _
                                        toVal, mvoTo, itemName, itemPrice, itemUnit, _
                                        catTo, itemNote, itemCode)

                                UpdateResiduseDictionary _
                                    dict, key, rowDate, startDate, endDate, _
                                    movementQty, True

                            End If

                        End If


                    '===========================================================
                    ' 2. ВИБУТТЯ
                    '
                    ' Віднімаємо ТІЛЬКИ у підрозділу "Звідки".
                    ' Поле "Куди" у такому рядку НЕ є підставою для надходження.
                    '===========================================================
                    Case "вибуття"

                        movementQty = qOut

                        If movementQty <> 0 And Len(fromVal) > 0 Then

                            If buildAllUnits Or fromVal = selectedUnit Then

                                mvoFrom = CStr(tbAllData.DataBodyRange(i, _
                                          tbAllData.ListColumns("МВО звідки").Index).value)

                                catFrom = CStr(tbAllData.DataBodyRange(i, _
                                          tbAllData.ListColumns("Категорія").Index).value)

                                key = MakeResiduseKey( _
                                        fromVal, mvoFrom, itemName, itemPrice, itemUnit, _
                                        catFrom, itemNote, itemCode)

                                UpdateResiduseDictionary _
                                    dict, key, rowDate, startDate, endDate, _
                                    movementQty, False

                            End If

                        End If


                    '===========================================================
                    ' 3. ВНУТРІШНЄ ПЕРЕМІЩЕННЯ
                    '
                    ' Одна і та сама кількість:
                    '   - віднімається у "Звідки";
                    '   - додається у "Куди".
                    '
                    ' У поточній базі така кількість записується у "Вибуло".
                    ' Якщо у старому рядку вона записана у "Надійшло",
                    ' використовуємо qIn як запасний варіант.
                    '===========================================================
                    Case "внутрішнє переміщення"

                        movementQty = qOut
                        If movementQty = 0 Then movementQty = qIn

                        If movementQty <> 0 Then

                            '---------------------------------------------------
                            ' ДОДАЄМО ОТРИМУВАЧУ
                            '---------------------------------------------------
                            If Len(toVal) > 0 Then

                                If buildAllUnits Or toVal = selectedUnit Then

                                    mvoTo = CStr(tbAllData.DataBodyRange(i, _
                                            tbAllData.ListColumns("МВО куди").Index).value)

                                    catTo = CStr(tbAllData.DataBodyRange(i, _
                                            tbAllData.ListColumns("Запасне поле3").Index).value)

                                    key = MakeResiduseKey( _
                                            toVal, mvoTo, itemName, itemPrice, itemUnit, _
                                            catTo, itemNote, itemCode)

                                    UpdateResiduseDictionary _
                                        dict, key, rowDate, startDate, endDate, _
                                        movementQty, True

                                End If

                            End If

                            '---------------------------------------------------
                            ' ВІДНІМАЄМО У ВІДПРАВНИКА
                            '---------------------------------------------------
                            If Len(fromVal) > 0 Then

                                If buildAllUnits Or fromVal = selectedUnit Then

                                    mvoFrom = CStr(tbAllData.DataBodyRange(i, _
                                              tbAllData.ListColumns("МВО звідки").Index).value)

                                    catFrom = CStr(tbAllData.DataBodyRange(i, _
                                              tbAllData.ListColumns("Категорія").Index).value)

                                    key = MakeResiduseKey( _
                                            fromVal, mvoFrom, itemName, itemPrice, itemUnit, _
                                            catFrom, itemNote, itemCode)

                                    UpdateResiduseDictionary _
                                        dict, key, rowDate, startDate, endDate, _
                                        movementQty, False

                                End If

                            End If

                        End If


                    '===========================================================
                    ' НЕВІДОМИЙ / ПОРОЖНІЙ ТИП ОПЕРАЦІЇ
                    '
                    ' Нічого не рахуємо, щоб не створювати помилковий рух.
                    '===========================================================
                    Case Else
                        'Навмисно пропускаємо.

                End Select

            End If
        End If

        If dataCount > 0 Then
            Call ProgressBar.Progress(i / dataCount * 100)
        End If

    Next i

    '---------------------------------------------------------------------------
    ' Виведення готової довідки
    '
    ' a(1) = залишок на початок
    ' a(2) = надійшло за період
    ' a(3) = вибуло за період
    '
    ' Залишок на кінець = a(1) + a(2) - a(3)
    '---------------------------------------------------------------------------
    rowIndex = 8
    counter = 0

    For Each itemKey In dict.Keys

        a = dict(itemKey)
        endBalance = CDbl(a(1)) + CDbl(a(2)) - CDbl(a(3))

        'Показуємо і позиції з нульовим кінцевим залишком,
        'якщо за період був рух.
        If CDbl(a(1)) <> 0 Or CDbl(a(2)) <> 0 Or CDbl(a(3)) <> 0 Or endBalance <> 0 Then

            parts = Split(CStr(itemKey), RES_SEP)
            counter = counter + 1

            wsOut.Cells(rowIndex, 2).value = counter
            wsOut.Cells(rowIndex, 3).value = parts(0)               'Підрозділ
            wsOut.Cells(rowIndex, 4).value = parts(1)               'МВО
            wsOut.Cells(rowIndex, 5).value = parts(2)               'Найменування
            wsOut.Cells(rowIndex, 6).value = parts(3)               'Ціна
            wsOut.Cells(rowIndex, 7).value = CDbl(a(2))             'Надійшло
            wsOut.Cells(rowIndex, 8).value = CDbl(a(3))             'Вибуло
            wsOut.Cells(rowIndex, 9).value = endBalance             'Залишок
            wsOut.Cells(rowIndex, 10).value = parts(6)              'Примітка
            wsOut.Cells(rowIndex, 11).value = parts(5)              'Категорія
            wsOut.Cells(rowIndex, 12).value = parts(4)              'Од. вим.
            wsOut.Cells(rowIndex, 13).value = parts(7)              'Код
            wsOut.Cells(rowIndex, 14).value = CDbl(a(1))            'Залишок на початок

            rowIndex = rowIndex + 1

        End If
    Next itemKey

    If wsOut.AutoFilterMode Then wsOut.AutoFilterMode = False

    If rowIndex > 8 Then
        wsOut.Range("B7:N" & rowIndex - 1).AutoFilter
    Else
        wsOut.Range("B7:N7").AutoFilter
    End If

    Application.ScreenUpdating = True

End Sub


'===============================================================================
' Додає одну операцію у словник залишків.
'
' a(1) = залишок на початок періоду
' a(2) = надходження за період
' a(3) = вибуття за період
'===============================================================================
Private Sub UpdateResiduseDictionary( _
    ByRef dict As Object, _
    ByVal key As String, _
    ByVal rowDate As Date, _
    ByVal startDate As Date, _
    ByVal endDate As Date, _
    ByVal qty As Double, _
    ByVal isIncoming As Boolean)

    Dim a As Variant

    If dict.exists(key) Then
        a = dict(key)
    Else
        ReDim a(1 To 3)
    End If

    If rowDate < startDate Then

        If isIncoming Then
            a(1) = CDbl(a(1)) + qty
        Else
            a(1) = CDbl(a(1)) - qty
        End If

    ElseIf rowDate <= endDate Then

        If isIncoming Then
            a(2) = CDbl(a(2)) + qty
        Else
            a(3) = CDbl(a(3)) + qty
        End If

    End If

    dict(key) = a

End Sub


Private Function MakeResiduseKey( _
    ByVal unitName As String, _
    ByVal mvoName As String, _
    ByVal itemName As String, _
    ByVal itemPrice As Variant, _
    ByVal itemUnit As String, _
    ByVal itemCategory As String, _
    ByVal itemNote As String, _
    ByVal itemCode As String) As String

    MakeResiduseKey = _
        unitName & RES_SEP & _
        mvoName & RES_SEP & _
        itemName & RES_SEP & _
        CStr(itemPrice) & RES_SEP & _
        itemUnit & RES_SEP & _
        itemCategory & RES_SEP & _
        itemNote & RES_SEP & _
        itemCode

End Function


Private Function SafeDoubleResiduse(ByVal v As Variant) As Double

    If IsError(v) Or IsEmpty(v) Or Len(Trim$(CStr(v))) = 0 Then
        SafeDoubleResiduse = 0
    ElseIf IsNumeric(v) Then
        SafeDoubleResiduse = CDbl(v)
    Else
        SafeDoubleResiduse = 0
    End If

End Function


'===============================================================================
' Вибір періоду при ручному формуванні "Довідка по підрозділу 2.0".
' За замовчуванням використовує дати з "Додаток 1".
'===============================================================================
Private Function AskResidusePeriod(ByRef startDate As Date, ByRef endDate As Date) As Boolean

    Dim wsAddOne As Worksheet
    Dim defaultFrom As Variant
    Dim defaultTo As Variant
    Dim v As Variant

    Set wsAddOne = ThisWorkbook.Worksheets("Додаток 1")

    On Error Resume Next
    defaultFrom = wsAddOne.Range("AddOnePeriodFrom").value
    defaultTo = wsAddOne.Range("AddOnePeriodTo").value
    On Error GoTo 0

    If Not IsDate(defaultFrom) Then defaultFrom = DateSerial(Year(Date), 1, 1)
    If Not IsDate(defaultTo) Then defaultTo = Date

    v = Application.InputBox( _
        Prompt:="Введіть дату початку періоду у форматі дд.мм.рррр", _
        Title:="Дата З", _
        Default:=Format$(CDate(defaultFrom), "dd.mm.yyyy"), _
        Type:=2)

    If VarType(v) = vbBoolean And v = False Then Exit Function

    If Not IsDate(v) Then
        MsgBox "Некоректна дата початку періоду.", vbExclamation
        Exit Function
    End If

    startDate = CDate(v)

    v = Application.InputBox( _
        Prompt:="Введіть дату кінця періоду у форматі дд.мм.рррр", _
        Title:="Дата ПО", _
        Default:=Format$(CDate(defaultTo), "dd.mm.yyyy"), _
        Type:=2)

    If VarType(v) = vbBoolean And v = False Then Exit Function

    If Not IsDate(v) Then
        MsgBox "Некоректна дата кінця періоду.", vbExclamation
        Exit Function
    End If

    endDate = CDate(v)

    If endDate < startDate Then
        MsgBox "Дата кінця періоду не може бути раніше дати початку.", vbExclamation
        Exit Function
    End If

    AskResidusePeriod = True

End Function


'===============================================================================
' Збереження довідки — залишено сумісним з попередньою логікою.
'===============================================================================
Public Sub SaveResiduseOneUnit2()

    ThisWorkbook.Sheets("Довідка по підрозділу 2.0").Copy

    Dim WbNew As Workbook
    Dim shp As Shape
    Dim fileName As String

    Set WbNew = ActiveWorkbook

    For Each shp In WbNew.Worksheets(1).Shapes
        shp.Delete
    Next shp

    fileName = ", залишки МЦ станом на " & Date
    WbNew.Worksheets(1).Range("A3").value = fileName

    WbNew.SaveAs _
        fileName:=ThisWorkbook.Path & Application.PathSeparator & _
                  CleanFileName(WbNew.Worksheets(1).Range("selectedUnit").value & _
                  fileName & ".xlsx"), _
        FileFormat:=xlOpenXMLWorkbook

    MsgBox "Файл з залишками матеріальних цінностей збережений в тій же папці, де файл бази даних", _
           vbInformation, ThisWorkbook.Path

End Sub


```

---

## 5. ResiduseOneGood.bas — Залишки обраного товару по підрозділах
### 5.1. Опис та алгоритм
Дозволяє проконтролювати, де саме в частині знаходиться конкретна номенклатурна позиція. Користувач обирає найменування зі списку `Goods`, і макрос виводить перелік підрозділів, де цей товар є в наявності, із зазначенням кількості, ціни та суми на аркуші «Залишки по товару».

### 5.2. Вихідний код ResiduseOneGood.bas
```vba
Attribute VB_Name = "ResiduseOneGood"
Sub ResiduseForGood()
Dim wbAllData As Workbook
Dim wsAllData As Worksheet
Dim wsResiduse As Worksheet
Dim tbAllData As ListObject

Set wbAllData = ThisWorkbook
Set wsAllData = wbAllData.Worksheets("База")
Set tbAllData = wsAllData.ListObjects("AllData")
Set wsResiduse = wbAllData.Worksheets("Залишки по майну")

Dim tbGoodCriteria As ListObject
Set tbGoodCriteria = wsResiduse.ListObjects("tbGoodCriteria")

'очищаємо дані з попереднього запиту
tbGoodCriteria.DataBodyRange.ClearContents

Dim frm As UserForm1
Set frm = New UserForm1

'Dim colFrom As Range, colTo As Range, cell As Range
'Dim dictUnits As Object
'Set dictUnits = CreateObject("Scripting.Dictionary")
'
'Set colFrom = tbAllData.ListColumns("Звідки").DataBodyRange
'Set colTo = tbAllData.ListColumns("Куди").DataBodyRange
'
'' Формуємо словник унікальних підрозділів
'For Each cell In colFrom
'    If Not dictUnits.exists(cell.value) And Not IsEmpty(cell.value) Then
'        dictUnits.Add cell.value, Nothing
'    End If
'Next cell
'
'For Each cell In colTo
'    If Not dictUnits.exists(cell.value) And Not IsEmpty(cell.value) Then
'        dictUnits.Add cell.value, Nothing
'    End If
'Next cell

' Виводимо список підрозділів у колонку A, починаючи з A7
'Dim n As Long
'n = 7
'wsResiduse.Range("A7:A" & wsResiduse.Rows.count).ClearContents ' Очистити перед вставкою
'
'Dim unitKey As Variant
'For Each unitKey In dictUnits.Keys
'    wsResiduse.Cells(n, 1).value = unitKey
'    n = n + 1
'Next unitKey
'
'' Зчитати назад у масив
'Dim lastRow As Long
'lastRow = wsResiduse.Cells(wsResiduse.Rows.count, 1).End(xlUp).Row

'Dim unitArray As Variant
'unitArray = wsResiduse.Range("A7:A" & lastRow).value

' Передати в ListBox
'If IsArray(unitArray) Then
    frm.Caption = "Виберіть майно"
    frm.ListData = wsResiduse2.Range("Goods")
'Else
'    frm.ListBox1.AddItem unitArray ' якщо один елемент
'End If

' Показати форму
frm.show

Dim NameOfThings As String
Dim ListGoods As Range
        Set ListGoods = wsResiduse2.Range("Goods")

' Отримати вибране
If frm.ListBox1.ListIndex <> -1 Then
    NameOfThings = frm.TextBox1.value
Else
    MsgBox "Майно не вибрано!", vbExclamation
    Exit Sub
End If

Set rngWithName = ListGoods.Find(NameOfThings, LookIn:=xlValues, LookAt:=xlWhole)
'================================закінчив тут, треба відібрати в верхню таблицю-фільтр критерії, якщо вони в полях проставлено "так"
'відфільтрувати базу даних по цих катергоріях
'відібрати унікальні підрозділи
'якщо в них нема залишків не виводити
'якщо є залишки - вивести в таблицю
'може зробити mvp тільки по найменуванню?
 If Not rngWithName Is Nothing Then
            Dim name As String
            Dim Nomenclature As String
            Dim unit As String
            Dim Category As String
            Dim NoteKt As String
            Dim price As Double
            Dim VAT As Double
            Dim QuantityRest As Double
            Dim NoteDt As String
            
'            Dim Place As String
            Dim CategoryFrom As String
            Dim CategoryTo As String
            Dim Note1 As String
            Dim Note2 As String
            Dim Note3 As String
            'Визначаємо значення за замовчуванням
       'рядок, де прописуємо "так", якщо хочемо щоб цей параметр включився в пошук
       Dim rowYes As Long
       rowYes = 8
       'найменування
       If wsResiduse.Cells(rowYes, 1) = "так" Then
            name = rngWithName.Offset(, -12).value
       End If
       'ціна
       If wsResiduse.Cells(rowYes, 2) = "так" Then
            price = rngWithName.Offset(, -11).value
       End If
       'одиниця виміру
       If wsResiduse.Cells(rowYes, 3) = "так" Then
            unit = rngWithName.Offset(, -7).value
       End If
       'категорія
       If wsResiduse.Cells(rowYes, 4) = "так" Then
            Category = rngWithName.Offset(, -8).value
       End If
       'примітка
       If wsResiduse.Cells(rowYes, 5) = "так" Then
            Note = rngWithName.Offset(, -9).value
       End If
       'код номенклатури
       If wsResiduse.Cells(rowYes, 6) = "так" Then
            Nomenclature = rngWithName.Offset(, -6).value
       End If
       'додаткове поле 2
       If wsResiduse.Cells(rowYes, 7) = "так" Then
           Note2 = rngWithName.Offset(, -4).value
       End If
       'додаткове поле 3
       If wsResiduse.Cells(rowYes, 8) = "так" Then
            Note3 = rngWithName.Offset(, -3).value
       End If
            
        
            
                With tbGoodCriteria.DataBodyRange.Rows
                    .Cells(1, 1) = name
                    .Cells(1, 2) = price
                    .Cells(1, 3) = unit
                    .Cells(1, 4) = Category
                    .Cells(1, 5) = Note
                    .Cells(1, 6) = Nomenclature
                    .Cells(1, 7) = Note2
                    .Cells(1, 8) = Note3
                    
'                     .Cells(1, 7) = CDbl(quantity)
'
'                    .Cells(1, 11) = CategoryTo
                    
'                    .Cells(1, 12) = Note1
                    
                End With
           End If



'
' wsResiduse.Range("selectedUnit").value = selectedUnit ' Вибраний підрозділ
' wsResiduse.Range("A7:A" & lastRow).ClearContents

Dim dict As Object
Set dict = CreateObject("Scripting.Dictionary")

Dim i As Long
Dim key As Variant
Dim result As Double

' Очистити старі дані
wsResiduse.Range("b8:i" & wsResiduse.Rows.count).ClearContents

Dim dataCount As Long
dataCount = tbAllData.ListRows.count

' Формуємо унікальні ключі на основі декількох полів
For i = 1 To dataCount
    If tbAllData.DataBodyRange(i, tbAllData.ListColumns("Куди").Index).value = selectedUnit _
    Or tbAllData.DataBodyRange(i, tbAllData.ListColumns("Звідки").Index).value = selectedUnit Then
    
      Dim dynamicCategory As String

' Вибираємо значення для категорії в залежності від напрямку
If tbAllData.DataBodyRange(i, tbAllData.ListColumns("Куди").Index).value = selectedUnit Then
    dynamicCategory = tbAllData.DataBodyRange(i, tbAllData.ListColumns("Запасне поле3").Index).value
Else
    dynamicCategory = tbAllData.DataBodyRange(i, tbAllData.ListColumns("Категорія").Index).value
End If

key = tbAllData.DataBodyRange(i, tbAllData.ListColumns("Код номенклатури").Index).value & "||" & _
      tbAllData.DataBodyRange(i, tbAllData.ListColumns("Найменування").Index).value & "||" & _
      tbAllData.DataBodyRange(i, tbAllData.ListColumns("Ціна").Index).value & "||" & _
      tbAllData.DataBodyRange(i, tbAllData.ListColumns("Одиниця виміру").Index).value & "||" & _
      dynamicCategory & "||" & _
      tbAllData.DataBodyRange(i, tbAllData.ListColumns("Примітка").Index).value


        If Not dict.exists(key) Then
            dict.Add key, 0
        End If
        
        ' Додаємо/віднімаємо залишок
        If tbAllData.DataBodyRange(i, tbAllData.ListColumns("Куди").Index).value = selectedUnit Then
            dict(key) = dict(key) + tbAllData.DataBodyRange(i, tbAllData.ListColumns("Надійшло").Index).value
            dict(key) = dict(key) + tbAllData.DataBodyRange(i, tbAllData.ListColumns("Вибуло").Index).value
        End If
        
        If tbAllData.DataBodyRange(i, tbAllData.ListColumns("Звідки").Index).value = selectedUnit Then
            dict(key) = dict(key) - tbAllData.DataBodyRange(i, tbAllData.ListColumns("Вибуло").Index).value
        End If
    End If
Next i

' Виводимо результати
Dim rowIndex As Long
rowIndex = 8
Dim parts() As String

Dim k As Long
Dim totalKeys As Long
Dim allKeys As Variant

allKeys = dict.Keys
totalKeys = dict.count
d = 0
For k = 1 To totalKeys
    If dict(allKeys(k - 1)) <> 0 Then
        parts = Split(allKeys(k - 1), "||")
        d = d + 1
        wsResiduse.Cells(rowIndex, 2).value = d
        wsResiduse.Cells(rowIndex, 3).value = parts(1) ' Найменування
        wsResiduse.Cells(rowIndex, 4).value = CDbl(parts(2)) ' Ціна
        wsResiduse.Cells(rowIndex, 5).value = CDbl(dict(allKeys(k - 1))) ' Залишок
        wsResiduse.Cells(rowIndex, 6).value = parts(5) ' Примітка
        wsResiduse.Cells(rowIndex, 7).value = parts(4) ' Категорія
        wsResiduse.Cells(rowIndex, 8).value = parts(3) ' Одиниця виміру
        wsResiduse.Cells(rowIndex, 9).value = parts(0) ' Код номенклатури
        rowIndex = rowIndex + 1
    End If
    
    Call ProgressBar.Progress(k / totalKeys * 100)
Next k
End Sub


```

---

## 6. ResiduseOnePeople.bas — Залишки майна по МВО
### 6.1. Опис та алгоритм
Формує відомість матеріальних цінностей, що обліковуються за конкретною матеріально-відповідальною особою. Обхід колонок `fromMRP` (Звідки) та `toMRP` (Куди) дозволяє точно порахувати підзвітний залишок майна за військовослужбовцем та зберегти звіт у файл.

### 6.2. Вихідний код ResiduseOnePeople.bas
```vba
Attribute VB_Name = "ResiduseOnePeople"
'=========3.9.14=========



Sub UpdateResiduseOnePeople()

Dim wbAllData As Workbook
Dim wsAllData As Worksheet
Dim wsResiduse As Worksheet
Dim tbAllData As ListObject

Set wbAllData = ThisWorkbook
Set wsAllData = wbAllData.Worksheets("База")
Set tbAllData = wsAllData.ListObjects("AllData")
Set wsResiduse = wbAllData.Worksheets("Довідка по людині") ' ОНОВЛЕНО


Dim frm As UserForm1
Set frm = New UserForm1

Dim colFrom As Range, colTo As Range, cell As Range
Dim dictUnits As Object
Set dictUnits = CreateObject("Scripting.Dictionary")

Set colFrom = tbAllData.ListColumns("МВО звідки").DataBodyRange
Set colTo = tbAllData.ListColumns("МВО куди").DataBodyRange

' Формуємо словник унікальних людей
For Each cell In colFrom
    If Not dictUnits.exists(cell.value) And Not IsEmpty(cell.value) Then
        dictUnits.Add cell.value, Nothing
    End If
Next cell

For Each cell In colTo
    If Not dictUnits.exists(cell.value) And Not IsEmpty(cell.value) Then
        dictUnits.Add cell.value, Nothing
    End If
Next cell

' Виводимо список людей
Dim wsTroops As Worksheet
Set wsTroops = wbAllData.Worksheets("Підрозділи")

Dim k As Long
k = 1
wsTroops.Range("B1:B" & wsTroops.Rows.count).ClearContents ' Очистити перед вставкою

Dim unitKey As Variant
For Each unitKey In dictUnits.Keys
    wsTroops.Cells(k, 2).value = unitKey
    k = k + 1
Next unitKey


Dim lastRowT As Long
lastRowT = wsTroops.Cells(wsTroops.Rows.count, 2).End(xlUp).Row


' Показуємо форму для вибору
frm.Caption = "Виберіть людину"
frm.ListData = wsTroops.Range("B1:B" & lastRowT)
    frm.show

    Dim selectedUnit As String
    If frm.ListBox1.ListIndex <> -1 Then
        selectedUnit = frm.ListBox1.value
    Else
        MsgBox "Людину не вибрано! Спробуйте ще раз!", vbExclamation
        Exit Sub
    End If

 wsResiduse.Range("selectedPeople").value = selectedUnit ' Вибраний підрозділ


Dim dict As Object
Set dict = CreateObject("Scripting.Dictionary")

Dim i As Long
Dim key As Variant
Dim result As Double

' Очистити старі дані
wsResiduse.Range("A8:I" & wsResiduse.Rows.count).ClearContents

Dim dataCount As Long
dataCount = tbAllData.ListRows.count

' Формуємо унікальні ключі на основі декількох полів
For i = 1 To dataCount
    If tbAllData.DataBodyRange(i, tbAllData.ListColumns("МВО куди").Index).value = selectedUnit _
    Or tbAllData.DataBodyRange(i, tbAllData.ListColumns("МВО звідки").Index).value = selectedUnit Then
    
      Dim dynamicCategory As String

' Вибираємо значення для категорії в залежності від напрямку
If tbAllData.DataBodyRange(i, tbAllData.ListColumns("МВО куди").Index).value = selectedUnit Then
    dynamicCategory = tbAllData.DataBodyRange(i, tbAllData.ListColumns("Запасне поле3").Index).value
Else
    dynamicCategory = tbAllData.DataBodyRange(i, tbAllData.ListColumns("Категорія").Index).value
End If

key = tbAllData.DataBodyRange(i, tbAllData.ListColumns("Код номенклатури").Index).value & "||" & _
      tbAllData.DataBodyRange(i, tbAllData.ListColumns("Найменування").Index).value & "||" & _
      tbAllData.DataBodyRange(i, tbAllData.ListColumns("Ціна").Index).value & "||" & _
      tbAllData.DataBodyRange(i, tbAllData.ListColumns("Одиниця виміру").Index).value & "||" & _
      dynamicCategory & "||" & _
      tbAllData.DataBodyRange(i, tbAllData.ListColumns("Примітка").Index).value
'      & "||" & _
'      tbAllData.DataBodyRange(i, tbAllData.ListColumns("Куди").Index).value

        If Not dict.exists(key) Then
            dict.Add key, 0
        End If
        
        ' Додаємо/віднімаємо залишок
        If tbAllData.DataBodyRange(i, tbAllData.ListColumns("МВО куди").Index).value = selectedUnit Then
            dict(key) = dict(key) + tbAllData.DataBodyRange(i, tbAllData.ListColumns("Надійшло").Index).value
            dict(key) = dict(key) + tbAllData.DataBodyRange(i, tbAllData.ListColumns("Вибуло").Index).value
        End If
        
        If tbAllData.DataBodyRange(i, tbAllData.ListColumns("МВО звідки").Index).value = selectedUnit Then
            dict(key) = dict(key) - tbAllData.DataBodyRange(i, tbAllData.ListColumns("Вибуло").Index).value
        End If
    End If
Next i

' Виводимо результати
Dim rowIndex As Long
rowIndex = 8
Dim parts() As String

Dim m As Long
Dim totalKeys As Long
Dim allKeys As Variant

allKeys = dict.Keys
totalKeys = dict.count
d = 0
For m = 1 To totalKeys
    If dict(allKeys(m - 1)) <> 0 Then
        parts = Split(allKeys(m - 1), "||")
        d = d + 1
        wsResiduse.Cells(rowIndex, 1).value = d
        wsResiduse.Cells(rowIndex, 2).value = parts(1) ' Найменування
        wsResiduse.Cells(rowIndex, 3).value = CDbl(parts(2)) ' Ціна
        wsResiduse.Cells(rowIndex, 4).value = CDbl(dict(allKeys(m - 1))) ' Залишок
        wsResiduse.Cells(rowIndex, 5).value = parts(5) ' Примітка
        wsResiduse.Cells(rowIndex, 6).value = parts(4) ' Категорія
        wsResiduse.Cells(rowIndex, 7).value = parts(3) ' Одиниця виміру
        wsResiduse.Cells(rowIndex, 8).value = parts(0) ' Код номенклатури
'        wsResiduse.Cells(rowIndex, 9).value = parts(6) ' підрозділ
        rowIndex = rowIndex + 1
    End If
    
    Call ProgressBar.Progress(m / totalKeys * 100)
Next m
'ставимо фільтри для зручності
    If wsResiduse.AutoFilterMode Then wsResiduse.AutoFilterMode = False
    wsResiduse.Range("A7").CurrentRegion.AutoFilter
    

End Sub




Sub SaveResiduseOnePeople()

        ThisWorkbook.Sheets("Довідка по людині").Select 'копіюємо перший лист з нарядом
        ThisWorkbook.Sheets("Довідка по людині").Copy 'зберігаємо як нову книгу
   
     Dim WbNew As Workbook
     
        Set WbNew = ActiveWorkbook
   
    Dim shp As Shape
    For Each shp In WbNew.Worksheets(1).Shapes
        shp.Delete
    Next shp
    
'    WbNew.Worksheets(1).Columns("N").ClearContents
        
      Dim fileName As String
     fileName = ", залишки МЦ станом на " & Date
        'складаємо назву файла
        
        WbNew.Worksheets(1).Range("A3") = fileName
        
        'зберігаємо файл
'        On Error GoTo WrongPath
         WbNew.SaveAs fileName:=ThisWorkbook.Path & Application.PathSeparator & _
         CleanFileName(WbNew.Worksheets(1).Range("selectedPeople").value & _
         fileName & ".xlsx"), FileFormat:=xlOpenXMLWorkbook
'     On Error GoTo 0
'      Exit Sub
'WrongPath:
     MsgBox "Файл з залишками матеріальних цінностей збережений і тій же папці, де файл бази даних", vbInformation, ThisWorkbook.Path
'     Exit Sub
     
End Sub





```
