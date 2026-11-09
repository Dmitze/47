# 03 — Базові та допоміжні модулі VBA

У цьому документі наведено повну документацію та вихідний код базових бібліотек проекту:
1. **functions.bas** — системні та облікові утиліти
2. **Module1.bas** — службовий модуль експорту компонентів VBA
3. **RibbonCallbacks.bas** — обробники подій стрічки Ribbon
4. **checkErrors.bas** — модуль валідації та пошуку помилок у базі даних
5. **ManualAdd.bas** — модуль ручного інтерактивного внесення операцій
6. **MoveForUnit.bas** — звіт руху майна по обраному підрозділу
7. **SumInWords.bas** — бібліотека генерації сум та чисел прописом українською мовою

---

## 1. functions.bas — Бібліотека загальних функцій

### 1.1. Призначення та взаємодія
Модуль містить центральні службові функції, які використовуються майже всіма іншими модулями системи:
- `CleanFileName`: очищає рядок від символів, заборонених у файловій системі Windows (\ / : * ? " < > |), замінюючи їх на тире. Забезпечує безпомилкове збереження документів.
- `FindValue`: швидка перевірка наявності значення у заданому діапазоні Excel Range.
- `IsDocumentInList`: перевіряє, чи міститься номер або назва документа у списку, розділеному крапкою з комою.
- `ClearUnitFilter_Residues`: скидає встановлені фільтри на аркуші залишків.
- `GetResidusePool`: критично важлива функція для формування документів. Запитує користувача, чи фільтрувати залишки лише за підрозділом-відправником, чи показувати всі залишки частини.
- `GetSelectedRowsArrayCell`: зчитує номери виділених користувачем рядків у таблиці Excel (використовується при переносі рядків у документи).
- `RemoveCustomContextMenu`: видаляє тимчасовий пункт кастомного контекстного меню.

### 1.2. Повний вихідний код functions.bas
```vba
Attribute VB_Name = "functions"
Function CleanFileName(str As String) As String
    Dim invalidChars As String
    Dim i As Integer
    Dim char As String
    
    ' Задаємо всі недопустимі символи, включаючи перенесення рядка
    invalidChars = "/\:*?""<>|" & Chr(10) & Chr(13)
    
    ' Проходимо через кожен символ в рядку
    For i = 1 To Len(str)
        char = Mid(str, i, 1)
        ' Якщо символ є недопустимим, замінюємо його на дефіс
        If InStr(invalidChars, char) > 0 Then
            CleanFileName = CleanFileName & "-"
        Else
            CleanFileName = CleanFileName & char
        End If
    Next i
End Function
Function FindValue(searchValue As String, searchRange As Range) As Boolean
    Dim foundCell As Range

    'Шукаємо задане значення в діапазоні
    Set foundCell = searchRange.Find(searchValue, LookIn:=xlValues, LookAt:=xlWhole)

    'Перевіряємо, чи знайдено задане значення
    If Not foundCell Is Nothing Then
        FindValue = True
    Else
        FindValue = False
    End If
End Function
'Пошук документа в рядку (повторюваного значення)
Function IsDocumentInList(docList As String, docName As String) As Boolean
    Dim docs() As String
    Dim i As Long
    
    docs = Split(docList, ";")
    For i = LBound(docs) To UBound(docs)
        If Trim(docs(i)) = Trim(docName) Then
            IsDocumentInList = True
            Exit Function
        End If
    Next i
    IsDocumentInList = False
End Function

'очистка фільтрів на сторінці
Sub ClearUnitFilter_Residues(ByVal ws As Worksheet)
    On Error Resume Next
    If ws.FilterMode Then ws.ShowAllData
    On Error GoTo 0
End Sub

'функція для пошуку матеріалів тільки в конкретному підрозділі
Function GetResidusePool(ByVal wsResiduse2 As Worksheet, ByVal FromUnit As String) As Range

    Dim lastRow As Long
    lastRow = wsResiduse2.Cells(wsResiduse2.Rows.count, "N").End(xlUp).Row
    If lastRow < 3 Then
        MsgBox "У залишках немає даних.", vbExclamation
        Exit Function
    End If

    Dim AllOrOneUnit As VbMsgBoxResult
    AllOrOneUnit = MsgBox( _
        "Шукати лише в підрозділі:" & vbNewLine & _
        FromUnit & " - так" & vbNewLine & _
        "шукати по всій базі - ні", _
        vbYesNo, "Де шукати потрібне майно?" _
    )

    Dim ResiduseCheet As Range
    Set ResiduseCheet = wsResiduse2.Range("N3:N" & lastRow) 'костиль для масиву

    If AllOrOneUnit = vbYes Then

        Dim hdr As Range
        Set hdr = wsResiduse2.Range("A2").CurrentRegion

        On Error Resume Next
        If wsResiduse2.FilterMode Then wsResiduse2.ShowAllData
        On Error GoTo 0

        hdr.AutoFilter Field:=12, Criteria1:=FromUnit

        On Error Resume Next
        Set GetResidusePool = ResiduseCheet.SpecialCells(xlCellTypeVisible)
        On Error GoTo 0

        If GetResidusePool Is Nothing Then
'
' Application.WorksheetFunction.CountA(GetResidusePool) = 0
            MsgBox "По підрозділу '" & FromUnit & "' немає позицій.", vbExclamation
            Call ClearUnitFilter_Residues(wsResiduse2)
            Exit Function
        End If

        Call ClearUnitFilter_Residues(wsResiduse2)

    Else
        Set GetResidusePool = ResiduseCheet
    End If

End Function

'===========3.9.16=============
'Function GetSelectedRowsArray()
'
'    Dim dict As Object
'    Dim cell As Range
'    Dim lo As ListObject
'    Dim rowIndex As Long
'
'    Set dict = CreateObject("Scripting.Dictionary")
'    Set lo = Worksheets("База").ListObjects("AllData")
'
'    For Each cell In Selection
'
'        If Not Intersect(cell, lo.DataBodyRange) Is Nothing Then
'
'            If cell.EntireRow.Hidden = False Then
'
'                rowIndex = cell.Row - lo.DataBodyRange.Row + 1
'
'                If Not dict.exists(rowIndex) Then
'                    dict.Add rowIndex, rowIndex
'                End If
'
'            End If
'
'        End If
'
'    Next cell
'
'    GetSelectedRowsArray = dict.Keys
'
'End Function

'===========3.9.16=============
'===========3.9.17=============

Function GetSelectedRowsArrayCell()

    Dim dict As Object
    Dim cell As Range
    
    Set dict = CreateObject("Scripting.Dictionary")

    For Each cell In Selection
        
        If cell.EntireRow.Hidden = False Then
            
            If Not dict.exists(cell.Row) Then
                dict.Add cell.Row, cell.Row
            End If
            
        End If
        
    Next cell
    
    GetSelectedRowsArrayCell = dict.Keys

End Function


Sub RemoveCustomContextMenu()

    Dim cb As CommandBar
    
    On Error Resume Next
    Set cb = Application.CommandBars("Cell")
    cb.Controls("Додати в документ").Delete
    On Error GoTo 0

End Sub
'===========3.9.17=============

```

---

## 2. Module1.bas — Службовий модуль експорту компонентів VBA

### 2.1. Призначення
Службовий модуль розробника. Дозволяє в один клік вивантажити весь вихідний код проекту (модулі .bas, класи .cls та форми .frm/.frx) у підпапку `VBA_Export\` у каталозі поточної книги.

### 2.2. Повний вихідний код Module1.bas
```vba
Attribute VB_Name = "Module1"
Sub ExportAllVBAElements()
    Dim vbaProject As Object
    Dim vbaComponent As Object
    Dim exportPath As String
    Dim extension As String
    
    ' Путь к папке, куда будут сохранены файлы (в данном случае - папка с текущим файлом Excel)
    exportPath = ActiveWorkbook.Path & "\VBA_Export\"
    
    ' Создаем папку, если она не существует
    If Dir(exportPath, vbDirectory) = "" Then
        MkDir exportPath
    End If
    
    Set vbaProject = ActiveWorkbook.VBProject
    
    ' Перебираем все компоненты в проекте
    For Each vbaComponent In vbaProject.VBComponents
        ' Определяем расширение файла в зависимости от типа компонента
        Select Case vbaComponent.Type
            Case 1 ' Стандартный модуль
                extension = ".bas"
            Case 2 ' Модуль класса
                extension = ".cls"
            Case 3 ' Пользовательская форма (UserForm)
                extension = ".frm"
            Case 100 ' Документы (Листы, ЭтаКнига)
                extension = ".cls"
            Case Else
                extension = ".txt"
        End Select
        
        ' Экспортируем компонент, если в нем есть код
        If vbaComponent.CodeModule.CountOfLines > 0 Then
            vbaComponent.Export exportPath & vbaComponent.name & extension
        End If
    Next vbaComponent
    
    MsgBox "Все VBA-скрипты успешно выгружены в папку:" & vbCrLf & exportPath, vbInformation, "Готово"
End Sub


```

---

## 3. RibbonCallbacks.bas — Обробники стрічки Ribbon

### 3.1. Призначення
Містить макроси зворотного виклику для інтерфейсу користувача Custom UI Ribbon XML. Зв'язує кнопки на панелі швидкого доступу Excel із головними процедурами системи.

### 3.2. Повний вихідний код RibbonCallbacks.bas
```vba
Attribute VB_Name = "RibbonCallbacks"
Option Explicit

Private Sub CallImport(control As IRibbonControl)
   Call UnivImport
End Sub

Private Sub CallManualAdd(control As IRibbonControl)
    Call ManualAdding
End Sub

```

---

## 4. checkErrors.bas — Перевірка коректності бази даних

### 4.1. Призначення та алгоритм перевірки
Процедура `HighlightRowsWithAllChecks` виконує комплексний аудит таблиці `AllData`:
- Перевіряє наявність обов'язкових колонок: Категорія, Запасне поле2, Запасне поле3, Тип документа, Ціна, Код номенклатури, № документа, Дата документа, Куди, Звідки, Одиниця виміру, Найменування, Надійшло, Вибуло.
- Видаляє попередні примітки та очищає жовте підсвічування.
- Валідує: чи є значення категорій допустимими (1-5, прид., не прид., потр. рем.), чи заповнені обов'язкові текстові реквізити, чи коректно вказані кількості та ціни.
- При знаходженні помилки підсвічує рядок жовтим кольором (`ColorIndex = 6`) та додає клітинці примітку з детальним поясненням помилки.

### 4.2. Повний вихідний код checkErrors.bas
```vba
Attribute VB_Name = "checkErrors"
Sub HighlightRowsWithAllChecks()
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim rw As ListRow
    Dim fields As Variant
    Dim colIndex As Object
    Dim colName As Variant
    Dim cell As Range
    Dim allowedValues As Variant
    Dim allowedZap2 As Variant
    Dim val As String
    Dim commentText As String
    Dim hasIssue As Boolean
    
    Set ws = ThisWorkbook.Worksheets("База")
    Set lo = ws.ListObjects("AllData")
    
    ' Колонки для перевірки
    fields = Array("Категорія", "Запасне поле2", "Запасне поле3", _
                   "Тип документа", "Ціна", "Код номенклатури", _
                   "№ документа", "Дата документа", "Куди", _
                   "Звідки", "Одиниця виміру", "Найменування", _
                   "Надійшло", "Вибуло")
    
    ' Індекс колонок
    Set colIndex = CreateObject("Scripting.Dictionary")
    For Each colName In fields
        On Error Resume Next
        colIndex(colName) = lo.ListColumns(colName).Index
        On Error GoTo 0
        If colIndex(colName) = 0 Then
            MsgBox "Не знайдено колонку: " & colName, vbExclamation
            Exit Sub
        End If
    Next colName
    
    ' Очистка попередніх підсвіток і приміток
    lo.DataBodyRange.Interior.ColorIndex = xlNone
    For Each colName In fields
        For Each cell In lo.ListColumns(colIndex(colName)).DataBodyRange
            If Not cell.Comment Is Nothing Then cell.Comment.Delete
        Next cell
    Next colName
    
    ' Дозволені значення
    allowedValues = Array("1", "2", "3", "4", "5", "пр.", "не пр.", "пот. рем.")
    allowedZap2 = Array("надходження", "вибуття", "внутрішнє переміщення")
    
    ' Перебір рядків
    For Each rw In lo.ListRows
        hasIssue = False
        
        ' Перевірка окремих колонок
        For Each colName In fields
            Set cell = rw.Range.Cells(1, colIndex(colName))
            commentText = ""
            val = Trim(cell.value)
            
            Select Case colName
                Case "Категорія", "Запасне поле3"
                    If Len(val) = 0 Or IsError(Application.Match(val, allowedValues, 0)) Then
                        commentText = "Поле не повинно бути пустим" & vbNewLine & _
                                      "Допустимі значення для поля: " & Join(allowedValues, ", ")
                        hasIssue = True
                    End If
                Case "Запасне поле2"
                    If Len(val) = 0 Or IsError(Application.Match(val, allowedZap2, 0)) Then
                        commentText = "Поле не повинно бути пустим" & vbNewLine & _
                                      "Допустимі значення для поля: " & Join(allowedZap2, ", ")
                        hasIssue = True
                    End If
                Case "Дата документа"
                    If Len(val) = 0 Or Not IsDate(val) Then
                        commentText = "Поле не повинно бути пустим" & vbNewLine & "Значення має бути датою"
                        hasIssue = True
                    End If
                Case "Ціна"
                    If Len(val) = 0 Or Not IsValidNumber(val) Then
                        commentText = "Поле не повинно бути пустим" & vbNewLine & "Значення має бути числом"
                        hasIssue = True
                    End If
                Case "Найменування"
                    If Len(val) = 0 Then
                        commentText = "Поле не повинно бути пустим"
                        hasIssue = True
                    ElseIf InStr(val, "~") > 0 Or InStr(val, Chr(10)) > 0 Or InStr(val, Chr(13)) > 0 Then
                        commentText = "Не допускаються перенос рядків та символ ~"
                        hasIssue = True
                    End If
                Case "Надійшло", "Вибуло"
                    If Len(val) > 0 And Not IsValidNumber(val) Then
                        commentText = "Значення має бути числом"
                        hasIssue = True
                    End If
                Case Else
                    If Len(val) = 0 Then
                        commentText = "Поле не повинно бути пустим"
                        hasIssue = True
                    End If
            End Select
            
            ' Додаємо примітку
            If Len(commentText) > 0 Then cell.AddComment commentText
        Next colName
        
        ' Перевірка одночасного заповнення Надійшло і Вибуло
        Dim valNad As String, valVyb As String
        valNad = Trim(rw.Range.Cells(1, colIndex("Надійшло")).value)
        valVyb = Trim(rw.Range.Cells(1, colIndex("Вибуло")).value)
        If Len(valNad) > 0 And Len(valVyb) > 0 Then
            rw.Range.Cells(1, colIndex("Надійшло")).AddComment _
                "Кількість має бути або в Надійшло або у Вибуло. Кожна операція окремий рядок"
            rw.Range.Cells(1, colIndex("Вибуло")).AddComment _
                "Кількість має бути або в Надійшло або у Вибуло. Кожна операція окремий рядок"
            hasIssue = True
        End If
        
        ' Перевірка узгодженості Надійшло - Запасне поле2
        Dim valZap2 As String
        valZap2 = Trim(rw.Range.Cells(1, colIndex("Запасне поле2")).value)
        If Len(valNad) > 0 And LCase(valZap2) <> "надходження" Then
            rw.Range.Cells(1, colIndex("Запасне поле2")).AddComment _
                "Змініть тип операції або перемістіть кількість з Надійшло у Вибуло"
            hasIssue = True
        End If
        If Len(valNad) = 0 And LCase(valZap2) = "надходження" Then
            rw.Range.Cells(1, colIndex("Запасне поле2")).AddComment _
                "Змініть тип операції або перемістіть кількість з Надійшло у Вибуло"
            hasIssue = True
        End If
        
        ' Підсвітка рядка
        If hasIssue Then rw.Range.Interior.Color = vbYellow
    Next rw
    MsgBox "Жовтим кольором підсвічені рядки з помилками. Читайте примітки в клітинках з помилками.", vbInformation, "Перевірку завершено"
End Sub

'------------------------------------------
Function IsValidNumber(s As String) As Boolean
    Dim i As Long, dotCount As Long, ch As String
    s = Trim(s)
    If Len(s) = 0 Then IsValidNumber = False: Exit Function
    
    dotCount = 0
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        Select Case ch
            Case "0" To "9"
            Case ","
                dotCount = dotCount + 1
                If dotCount > 1 Then IsValidNumber = False: Exit Function
            Case Else
                IsValidNumber = False
                Exit Function
        End Select
    Next i
    IsValidNumber = True
End Function



```

---

## 5. ManualAdd.bas — Інтерактивне ручне внесення проводок

### 5.1. Призначення та алгоритм
Процедура `ManualAdding` реалізує майстер внесення операції через послідовні вікна `InputBox`:
- Запитує номер документа та перевіряє, чи не існує вже такий документ у таблиці `AllData` (захист від повторного внесення).
- По черзі зчитує 27 реквізитів: підрозділ-відправник, одержувач, підставу, службу, дату, номенклатуру, кількість, одиницю, категорію, ціну, серійний номер тощо.
- Формує 2D-масив `ArrInv(1 To 1, 1 To 27)` і вставляє новий рядок у кінець таблиці `AllData` одним блоком.

### 5.2. Повний вихідний код ManualAdd.bas
```vba
Attribute VB_Name = "ManualAdd"
Sub ManualAdding()

Dim AddManual

AddManual = MsgBox("Внести дані вручну?", vbQuestion + vbOKCancel, "Внести вручну?")
 
If AddManual = 2 Then
 Exit Sub
End If

Dim wbAllData As Workbook
Dim wsAllData As Worksheet
Dim tbAllData As ListObject
    Set wbAllData = ThisWorkbook
    Set wsAllData = wbAllData.Worksheets("База")
    Set tbAllData = wsAllData.ListObjects("AllData")
    
'Dim wsAccountingBook As Worksheet
'Dim tbAccountingBook As ListObject
'
'    Set wsAccountingBook = wbAllData.Worksheets("Книга обліку д13")
'    Set tbAccountingBook = wsAccountingBook.ListObjects("AccountingBook")
'
''для перевірки чи є в книзі обліку мц, яке розносимо в БД
'Dim RngToSearchData As Range
'Dim DataThing As String

            Dim Executor As String  '+
            Dim FromUnit As String '+
            Dim fromMRP As String '+
            Dim toMRP As String '+
            Dim toUnit As String '+
            Dim NumberInv As String '+
'            Dim NumberInvClear As String
            Dim DateInv As Date '+
            Dim NumberDemand As String '+
             
            Dim NameThings As String '+
            Dim MensureThings As String '+
            Dim QuantityThings As Double '+
            Dim СategoryThings As String '+
            Dim СategoryThingsTo As String
            Dim ItemNumber As String '+
            Dim Note As String '+
            Dim price As Double '+
            
            Dim Field10 '+
            Dim Field11 '+
            Dim Field12 '+
'            Dim TimeInv
            
            Dim Service As String '+
            Dim Content As String '+
            
            Dim InputDate As Date
            Dim typeDoc As String '+


InputDate = Date




'перевірка, чи є такий документ в базі
NumberInv = InputBox("Введіть номер документа:", "№ документа") '+

Dim foundCell As Range
Dim firstColumn As Range
        
            Set firstColumn = tbAllData.ListColumns(14).DataBodyRange
            
            Set foundCell = firstColumn.Find(What:=NumberInv, LookIn:=xlValues, LookAt:=xlWhole)
    
            If Not foundCell Is Nothing Then
            Dim EnterAnyway
                EnterAnyway = MsgBox("Документ з номером " & NumberInv & " вже є в базі." & vbNewLine & _
                "Все одно додати в базу?", vbYesNo, "Документ вже є в базі")
                If EnterAnyway = 7 Then
                    Exit Sub
                End If
            
            End If

typeDoc = InputBox("Введіть назву(акт, накладна і т.д.) документа:", "Тип документа", "Акт") '+

DateInv = InputBox("Введіть дату документа " & typeDoc & " №" & NumberInv & " в форматі 27.04.2023 :", "Дата документа") '+


FromUnit = InputBox("Підрозділ (організація, база, в/ч), з якого надійшло (який передає) майно", "Постачальник")
toUnit = InputBox("Підрозділ (організація, база, в/ч), отримувач майна", "Отримувач")

NumberDemand = InputBox("Назва, номер, і дата документа підстави операції (наряд, рапорт, заявка і т.д.)", "Підстава операцї")

Service = InputBox("До номенклатури якої служби забезпечення відноситься майно (зв'язок, медичне, прод, КЕС і т.д.)?", "Служба забезпечення")
Content = InputBox("Виберіть тип операції:" & vbNewLine & _
                    "надходження - 1" & vbNewLine & _
                    "вибуття - 2" & vbNewLine & _
                    "внутрішнє переміщення - 3", "Тип операції", 1)
'
        If Content = 2 Then
            Content = "вибуття"
        ElseIf Content = 1 Then
            Content = "надходження"
        Else
            Content = "внутрішнє переміщення"
        End If

Executor = InputBox("Прізвище, хто здійснює занесення документа в базу", "Виконавець")
toMRP = InputBox("Прізвище мателіально відповідальної особи, яка отримує", "МВО отримувача")
fromMRP = InputBox("Прізвище мателіально відповідальної особи, яка передає", "МВО відправника")

NextPosition:

NameThings = InputBox("Введіть найменування майна", "Найменування")

ItemNumber = InputBox("Код номенклатури (за наявності) " & NameThings, "Код", "б/н")

MensureThings = InputBox("Одиниця виміру: " & NameThings, "Одиниця виміру", "к-т")

QuantityThings = InputBox("Кількість: " & NameThings, "Кількість", 1)

СategoryThings = InputBox("Категорія відправника: " & NameThings, "Категорія", "прид.")

СategoryThingsTo = InputBox("Категорія одержувача: " & NameThings, "Категорія", "прид.")

Dim txtPrice As String

txtPrice = InputBox("Ціна за одиницю: " & NameThings, "Ціна")


If txtPrice <> "" Then
    txtPrice = Replace(txtPrice, ".", Application.DecimalSeparator)
    txtPrice = Replace(txtPrice, ",", Application.DecimalSeparator)

    If IsNumeric(txtPrice) Then
        price = CDbl(txtPrice)
    Else
        MsgBox "Некоректне число"
    End If
End If


Note = InputBox("Примітка: " & NameThings, "Примітка")

Field10 = InputBox("Поле 10 (додаткова інформація, яку ви бажаєте додати про майно) для " & NameThings, "Поле 10")
Field11 = InputBox("Поле 11 (додаткова інформація, яку ви бажаєте додати про майно) для " & NameThings, "Поле 11")
Field12 = InputBox("Поле 12 (додаткова інформація, яку ви бажаєте додати про майно) для " & NameThings, "Поле 12")


wsAllData.Activate 'активуємо лист з базою

            Dim ArrInv()
            Dim ArrSizeX
            Dim ArrSizeY
                ArrSizeX = tbAllData.ListColumns.count
                ArrSizeY = 1
                
                            
                            ReDim ArrInv(1 To ArrSizeY, 1 To ArrSizeX)
                            arrRow = 1
                           
                           ArrInv(arrRow, 1) = InputDate
                           ArrInv(arrRow, 2) = NameThings
                           ArrInv(arrRow, 4) = MensureThings
                           ArrInv(arrRow, 5) = СategoryThings
                           
                           If Content = "вибуття" Or Content = "внутрішнє переміщення" Then
                                ArrInv(arrRow, 7) = QuantityThings
                           Else
                                ArrInv(arrRow, 6) = QuantityThings
                           End If
                                                                                 
                           ArrInv(arrRow, 8) = FromUnit
                           ArrInv(arrRow, 9) = toUnit
                           ArrInv(arrRow, 10) = fromMRP
                           ArrInv(arrRow, 11) = toMRP
                           ArrInv(arrRow, 12) = NumberDemand
                           ArrInv(arrRow, 13) = DateInv
                           ArrInv(arrRow, 14) = NumberInv
                           ArrInv(arrRow, 15) = NumberInvClear
                           ArrInv(arrRow, 16) = Note
                           
                           ArrInv(arrRow, 17) = Executor
                           ArrInv(arrRow, 18) = ItemNumber
                           
'                           ArrInv(arrRow, 19) = TimeInv
                           ArrInv(arrRow, 20) = price
                           ArrInv(arrRow, 21) = Service
                                                      
                           ArrInv(arrRow, 22) = Field10
                           ArrInv(arrRow, 23) = Field11
                           ArrInv(arrRow, 24) = Field12
                           ArrInv(arrRow, 25) = typeDoc
                           ArrInv(arrRow, 26) = Content
                           ArrInv(arrRow, 27) = СategoryThingsTo


Dim lastRow As Long
    lastRow = wsAllData.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row + 1
Dim arrTarget
wsAllData.Activate
    Set arrTarget = wsAllData.Range(Cells(lastRow, 1), Cells(lastRow + UBound(ArrInv()) - 1, ArrSizeX))
    arrTarget.value = ArrInv
    

'  'Якщо нема мц в книзі обліку - додаємо колонки з такими назвами
'
'            DataThing = NameThings & ItemNumber & MensureThings & Price
'            Set RngToSearchData = wsAccountingBook.Range("16:16")
'            ThingExist = FindValue(DataThing, RngToSearchData)
'
'                    If Not ThingExist Then
''                        MsgBox "Працюємо"
'                        wsAccountingBook.Activate
'                        Dim LastColumn As Long
'                        LastColumn = tbAccountingBook.Range.Columns.count
'
'                        Dim CopyRange1 As Range
'                        Dim CopyRange2 As Range
'
'                        'копіюємо стовпчики
'                        Set CopyRange1 = wsAccountingBook.Range(Cells(15, LastColumn - 2), Cells(16, LastColumn))
'
'                            CopyRange1.Copy
'                            wsAccountingBook.Cells(15, LastColumn + 1).PasteSpecial xlPasteAll
'
'                        Set CopyRange2 = wsAccountingBook.Range(Cells(7, LastColumn - 2), Cells(14, LastColumn))
'                            CopyRange2.Copy
'                            wsAccountingBook.Cells(7, LastColumn + 1).PasteSpecial xlPasteFormats
'
'                        tbAccountingBook.ListColumns(LastColumn - 2).DataBodyRange.Copy
'                        wsAccountingBook.Cells(18, LastColumn + 1).PasteSpecial xlPasteAll
'
'                        tbAccountingBook.ListColumns(LastColumn - 1).DataBodyRange.Copy
'                        wsAccountingBook.Cells(18, LastColumn + 2).PasteSpecial xlPasteAll
'
'                        tbAccountingBook.ListColumns(LastColumn).DataBodyRange.Copy
'                        wsAccountingBook.Cells(18, LastColumn + 3).PasteSpecial xlPasteAll
'
'                        'додаємо позицію в книгу
'                        wsAccountingBook.Cells(8, LastColumn + 1) = NameThings
'                        wsAccountingBook.Cells(10, LastColumn + 1) = ItemNumber
'                        wsAccountingBook.Cells(12, LastColumn + 1) = MensureThings
'                        wsAccountingBook.Cells(14, LastColumn + 1) = Price
'
'
'                    End If
'


Dim WhatIsNext

WhatIsNext = MsgBox("Позиція згідно " & vbNewLine & typeDoc & " №" & NumberInv & " від " & DateInv & _
vbNewLine & NameThings & vbNewLine & _
"номенклатурний номер: " & ItemNumber & vbNewLine & "Одиниця виміру: " & MensureThings & vbNewLine & "Кількість:" & _
QuantityThings & vbNewLine & "Ціна: " & price & vbNewLine & "Примітка" & Note _
& vbNewLine & "рознесена. Перейти до наступної позиції?", vbYesNo, "Рознести наступну позицію?")

If WhatIsNext = 6 Then
    GoTo NextPosition
End If
        
 
    
'
'            'додаємо накладну в книгу обліку
'        Set rowAccBook = tbAccountingBook.ListRows.Add
'
'        rowAccBook.Range(1) = InputDate
'        rowAccBook.Range(2) = typeDoc
'        rowAccBook.Range(3) = NumberInv
'        rowAccBook.Range(4) = DateInv
'
'        If Content = "вибуття" Or Content = "внутрішнє переміщення" Then
'             rowAccBook.Range(5) = toUnit
'        Else
'             rowAccBook.Range(5) = fromUnit
'        End If



MsgBox "Документ рознесений", , NumberInv

            
End Sub





```

---

## 6. MoveForUnit.bas — Рух майна по підрозділу

### 6.1. Призначення
Модуль призначений для формування вибірки всіх операцій руху майна для конкретного підрозділу за визначений проміжок часу за допомогою механізму `AdvancedFilter`:
- `MoveUnitGoods`: запитує підрозділ і дати, фільтрує таблицю `AllData` і вивантажує результат на аркуш «Рух по підрозділу».
- `SaveMoveUnitGoods`: експортує сформований звіт у незалежний файл .xlsx.

### 6.2. Повний вихідний код MoveForUnit.bas
```vba
Attribute VB_Name = "MoveForUnit"
Option Explicit
Sub MoveUnitGoods() 'рух по підрозділу

    Dim wbAllData As Workbook
    Dim wsAllData As Worksheet
    Dim wsMove As Worksheet
    Dim tbAllData As ListObject

    Set wbAllData = ThisWorkbook
    Set wsAllData = wbAllData.Worksheets("База")
    Set tbAllData = wsAllData.ListObjects("AllData")
    Set wsMove = wbAllData.Worksheets("Рух по підрозділу")

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
    wsMove.Range("A7:A" & wsMove.Rows.count).ClearContents ' Очистити перед вставкою

    Dim unitKey As Variant
    For Each unitKey In dictUnits.Keys
        wsMove.Cells(n, 1).value = unitKey
        n = n + 1
    Next unitKey

    ' Зчитати назад у масив
    Dim lastRow As Long
    lastRow = wsMove.Cells(wsMove.Rows.count, 1).End(xlUp).Row

    frm.Caption = "Виберіть підрозділ"
    frm.ListData = wsMove.Range("A7:A" & lastRow)
    frm.show

    Dim selectedUnit As String
    If frm.ListBox1.ListIndex <> -1 Then
        selectedUnit = frm.ListBox1.value
    Else
        MsgBox "Підрозділ не вибрано!", vbExclamation
        Exit Sub
    End If

wsMove.Range("moveUnit") = selectedUnit
wsMove.Range("CreateDate") = Date


    ' Очистити тимчасові значення
    wsMove.Range("A7:A" & lastRow).ClearContents

    ' Очистити попередні фільтри
    If Not tbAllData.AutoFilter Is Nothing Then
    If tbAllData.AutoFilter.FilterMode Then
        tbAllData.AutoFilter.ShowAllData
    End If
End If

    ' Очистити старі дані на листі "Рух по підрозділу"
    wsMove.Rows("5:" & wsMove.Rows.count).ClearContents

    ' Вставити заголовки у рядок 4
    'tbAllData.HeaderRowRange.Copy destination:=wsMove.Range("b4")

' Зчитуємо діапазон дат
    Dim startDate As Date, endDate As Date
    On Error Resume Next
    startDate = CDate(InputBox("Введіть дату початку періоду у форматі дд.мм.рррр", "Дата З (включно)", Date))
    endDate = CDate(InputBox("Введіть дату кінця періоду у форматі дд.мм.рррр", "Дата ПО (включно)", Date))
    On Error GoTo 0


'==========3.9.25==============

Dim periodMove As String
    ' Створити тимчасові критерії у комірках
    With wsMove
        .Range("c1:d3").ClearContents
        .Range("c1").value = "Звідки"
        .Range("d1").value = "Куди"
        .Range("c2").value = selectedUnit
        .Range("d3").value = selectedUnit
        
        .Range("e1").value = "Дата документа"
        .Range("f1").value = "Дата документа"
        
    If startDate = 0 Then
    .Range("e2").value = ""
    .Range("e3").value = ""
     periodMove = "з (не застосовувався)"
    Else
    .Range("e2").value = ">=" & CLng(startDate)
    .Range("e3").value = ">=" & CLng(startDate)
    periodMove = "з " & startDate
    End If
          
    If endDate = 0 Then
    .Range("f2").value = ""
    .Range("f3").value = ""
    periodMove = periodMove & " по(не застосовувався)"
    Else
    .Range("f2").value = "<=" & CLng(endDate)
    .Range("f3").value = "<=" & CLng(endDate)
    periodMove = periodMove & " по " & endDate
    End If
    
    End With
wsMove.Range("periodMove") = periodMove


'==========3.9.25==============


    ' Застосувати розширений фільтр
    tbAllData.Range.AdvancedFilter _
        Action:=xlFilterCopy, _
        criteriaRange:=wsMove.Range("c1:f3"), _
        CopyToRange:=wsMove.Range("a5"), _
        Unique:=False
        
'CopyToRange:=wsMove.Range("b5"), _
'     Очистити тимчасові критерії
    wsMove.Range("c1:f3").ClearContents

    ' Скинути фільтр у джерельній таблиці
    If Not tbAllData.AutoFilter Is Nothing Then
    If tbAllData.AutoFilter.FilterMode Then
        tbAllData.AutoFilter.ShowAllData
    End If
End If

End Sub

Sub SaveMoveUnitGoods()

        ThisWorkbook.Sheets("Рух по підрозділу").Select 'копіюємо перший лист з нарядом
        ThisWorkbook.Sheets("Рух по підрозділу").Copy 'зберігаємо як нову книгу
   
     Dim WbNew As Workbook
     
        Set WbNew = ActiveWorkbook
   
    Dim shp As Shape
    For Each shp In WbNew.Worksheets(1).Shapes
        shp.Delete
    Next shp
    
'    WbNew.Worksheets(1).Columns("N").ClearContents
        
      Dim fileName As String
     fileName = ", рух майна станом на " & WbNew.Worksheets(1).Range("CreateDate")
        'складаємо назву файла
        
'        WbNew.Worksheets(1).Range("A3") = fileName
        
        'зберігаємо файл
'        On Error GoTo WrongPath
         WbNew.SaveAs fileName:=ThisWorkbook.Path & Application.PathSeparator & _
         CleanFileName(WbNew.Worksheets(1).Range("moveUnit").value & _
         fileName & ".xlsx"), FileFormat:=xlOpenXMLWorkbook
'     On Error GoTo 0
'      Exit Sub
'WrongPath:
     MsgBox "Файл з рухом матеріальних цінностей збережений і тій же папці, де файл бази даних", vbInformation, ThisWorkbook.Path
'     Exit Sub
     
End Sub

```

---

## 7. SumInWords.bas — Сума та числа прописом українською мовою

### 7.1. Призначення та ключові функції
Модуль містить повну реалізацію конвертації чисел у текстовий формат з урахуванням граматичних родів, відмінків та числівників української мови:
- `СумаПрописом1(суміi)`: формує рядок виду "Дванадцять тисяч триста сорок п'ять гривень 00 копійок".
- `ЧислоПрописом11(число)`: повертає кількість предметів або порядкових номерів прописом.
- Підтримує розряди одиниць, тисяч, мільйонів, мільярдів з правильними закінченнями.

### 7.2. Повний вихідний код SumInWords.bas
```vba
Attribute VB_Name = "SumInWords"
'Option Explicit
' Функции для вычисления суммы прописью по
' числовому значению от 0 до 999999999999
'
' Вспомогательные переменные
Dim Тысячи, Миллионы As Boolean
Dim Миллиарды, ВторойДесяток As Boolean
' Массмв составных частей
Dim Часть(32)  As String
' Логические константы
Const Истина   As Boolean = True
Const Ложь     As Boolean = False
'
' Функция возвращает сумму прописью
'
Function СумаПрописом1(Гривнi)
    ' Считаем копейки
    Переменная = (Гривнi - Fix(Гривнi)) * 100
    If (Переменная - Fix(Переменная)) >= 0.5 Then
        If Переменная >= 99.5 Then
            Переменная = 0
            Гривнi = Гривнi + 1
        Else
            Переменная = Fix(Переменная) + 1
        End If
    Else
        Переменная = Fix(Переменная)
    End If
    Копейки = CStr(Переменная)
    ' Вызов функции для получения числа прописью
    Число = CStr(Fix(Гривнi))
    СумаПрописом1 = ЧислоПрописомЖР(Число)
    ' Строку с заглавной буквы
    СумаПрописом1 = UCase(Mid(СумаПрописом1, 1, 1)) + _
                    Mid(СумаПрописом1, 2)
    ' Вычислить длину исходного числа
    Длина = Len(Число)
    ' Если число только из одной цифры, добавить
    ' до двух (для единообразия алгоритма)
    If Длина = 1 Then
        Число = "0" & Число
        Длина = Длина + 1
    End If
    ' Добавление нужного окончания строки
    '
    ' Для чисел, оканчивающихся на 10, 11, 12, 13,
    ' 14, 15, 16, 17, 18, 19 добавляем "гривень"
    If Mid(Число, Длина - 1, 1) = 1 Then
        СумаПрописом1 = СумаПрописом1 + "гривень"
        ' Для всех остальных случаев
    Else
        Select Case Mid(Число, Длина)
            ' Для чисел, оканчивающихся на 1 добавляем "гривня"
            Case 1
                СумаПрописом1 = СумаПрописом1 + "гривня"
                ' Для чисел, оканчивающихся на 2, 3, 4
                ' добавляем "гривнi "
            Case 2, 3, 4
                СумаПрописом1 = СумаПрописом1 + "гривнi"
                ' Для чисел, оканчивающихся на 5, 6, 7, 8,
                ' 9, 0 добавляем "гривень"
            Case Else
                СумаПрописом1 = СумаПрописом1 + "гривень"
        End Select
    End If
    ' Окончательно формируем результат, добавляя копейки
    If Len(Копейки) = 1 Then
        Копейки = "0" & Копейки
    End If
    СумаПрописом1 = СумаПрописом1 + " " + Копейки + " "
    ' Для чисел, оканчивающихся на 10, 11, 12, 13,
    ' 14, 15, 16, 17, 18, 19 добавляем "копiйок"
    If Mid(Копейки, 1, 1) = 1 Then
        СумаПрописом1 = СумаПрописом1 + "копiйок"
        ' Для всех остальных случаев
    Else
        Select Case Mid(Копейки, 2)
            ' Для чисел, оканчивающихся на 1 добавляем "копiйка"
            Case 1
                СумаПрописом1 = СумаПрописом1 + "копiйка"
                ' Для чисел, оканчивающихся на 2, 3, 4
                ' добавляем "копiйки"
            Case 2, 3, 4
                СумаПрописом1 = СумаПрописом1 + "копiйки"
                ' Для чисел, оканчивающихся на 5, 6, 7, 8,
                ' 9, 0 добавляем "копiйок"
            Case Else
                СумаПрописом1 = СумаПрописом1 + "копiйок"
        End Select
    End If
End Function
'
' функция возвращает число прописью
'
Function ЧислоПрописом11(Число)
    ' Присвоение значений массиву частей
    Часть(1) = "оди":       Часть(2) = "дві"
    Часть(3) = "три":       Часть(4) = "чотир"
    Часть(5) = "п'ят":       Часть(6) = "шiст"
    Часть(7) = "сiм":       Часть(8) = "вiсiм"
    Часть(9) = "дев'ят":     Часть(10) = "н"
    Часть(11) = "и":        Часть(12) = "ь"
    Часть(13) = "надцять":  Часть(14) = "дцять"
    Часть(15) = "сорок":    Часть(16) = "дев'яно"
    Часть(17) = "сто":      Часть(18) = "два"
    Часть(19) = "стi":      Часть(20) = "сот"
    Часть(21) = "одна":     Часть(22) = "тисяч"
    Часть(23) = "а":        Часть(24) = "i"
    Часть(25) = "мiльйон":  Часть(26) = "iв"
    Часть(27) = " ":        Часть(28) = "":
    Часть(29) = "десят":    Часть(30) = "ста"
    Часть(31) = "мільярд": Часть(32) = "нуль "
    ' Временные переменные вначале сбрасываются
    Тысячи = Ложь:      Миллионы = Ложь
    Миллиарды = Ложь:   ВторойДесяток = Ложь
    ' Отбрасываем дробную часть, если она есть
    Число = Fix(Число)
    ' Определяем длину исходного числа
    Длина = Len(Число)
    ' Цикл по всем цифрам числа, начиная с крайней
    ' левой до крайней правой
    For Позиция = Длина To 1 Step -1
        ' Добавляются очередные слова, описывающие
        ' текущую цифру
        ЧислоПрописом11 = ЧислоПрописом11 + _
                        ЦифраСтрокой(Mid(Число, _
                        Длина - Позиция + 1, 1), _
                        Позиция)
    Next Позиция
    ' Алгоритм возвращает пустую строку при
    ' нулевом аргументе. Исправим это
    If ЧислоПрописом11 = "" Then
        ЧислоПрописом = Часть(32)
    End If
End Function
'
' Составление слов из частей по очередной
' цифре числа и по предистории работы
'
' Функция доступна только в текущем модуле
'
Private Function ЦифраСтрокой(Цифра, Место) As String
    ' Если сотни или десятки миллиардов, то
    ' запомнить об этом для будущего
    If (Цифра <> 0) And ((Место = 11) Or _
        (Место = 12)) Then
        Миллиарды = Истина
    End If
    ' Если сотни или десятки миллионов, то
    ' запомнить об этом для будущего
    If (Цифра <> 0) And ((Место = 8) Or _
        (Место = 9)) Then
        Миллионы = Истина
    End If
    ' Если сотни или десятки тысяч, то
    ' запомнить об этом для будущего
    If (Цифра <> 0) And ((Место = 5) Or _
        (Место = 6)) Then
        Тысячи = Истина
    End If
    ' Если предыдущая цифра была единица
    ' в поле десятков, то выбираем
    If ВторойДесяток Then
        Select Case Цифра
            ' пишем "десять "
            Case 0
                ЦифраСтрокой = Часть(29) + Часть(12) + _
                               Часть(27)
                ' пишем "дванадцять "
            Case 2
                ЦифраСтрокой = Часть(18) + Часть(13) + _
                               Часть(27)
                ' в остальных случаях пишем название цифры
                ' плюс "надцять "
            Case Else
                ЦифраСтрокой = Часть(Цифра) + Часть(13) + _
                               Часть(27)
        End Select
        ' Добавляем название разрядов
        Select Case Место
            Case 4
                ' добавляем "тисяч "
                ЦифраСтрокой = ЦифраСтрокой + Часть(22) + _
                               Часть(27)
                ' добавляем "мiльйонiв "
            Case 7
                ЦифраСтрокой = ЦифраСтрокой + Часть(25) + _
                               Часть(26) + Часть(27)
                ' добавляем "мiлiардiв "
            Case 10
                ЦифраСтрокой = ЦифраСтрокой + Часть(31) + _
                               Часть(26) + Часть(27)
        End Select
        ' Сбрасываем значения, так как переходим к
        ' предыдущим разрядам
        ВторойДесяток = Ложь:    Миллионы = Ложь
        Миллиарды = Ложь:        Тысячи = Ложь
        ' Во всех остальных случаях, то есть
        ' не для описания чисел второго десятка
    Else
        ' Определяем название десятков
        If (Место = 2) Or (Место = 5) Or _
            (Место = 8) Or (Место = 11) Then
            Select Case Цифра
                ' Запоминаем про второй десяток для
                ' подстановки при следующем входе
                Case 1
                    ВторойДесяток = Истина
                    ' пишем "двадцять "
                Case 2
                    ЦифраСтрокой = Часть(18) + Часть(14) + _
                                   Часть(27)
                    ' пишем "тридцять "
                Case 3
                    ЦифраСтрокой = Часть(Цифра) + Часть(14) + _
                                   Часть(27)
                    ' пишем "сорок "
                Case 4
                    ЦифраСтрокой = Часть(15) + Часть(27)
                    ' пишем "дев'яносто "
                Case 9
                    ЦифраСтрокой = Часть(16) + Часть(17) + _
                                   Часть(27)
                    ' в остальных случаях пишем название цифры
                    ' плюс "десят "
                Case 5, 6, 7, 8
                    ЦифраСтрокой = Часть(Цифра) + Часть(29) + _
                                   Часть(27)
            End Select
        End If
        ' Определяем названия сотен
        If (Место = 3) Or (Место = 6) Or _
            (Место = 9) Or (Место = 12) Then
            Select Case Цифра
                ' пишем "сто "
                Case 1
                    ЦифраСтрокой = Часть(17) + Часть(27)
                    ' пишем "двісті "
                Case 2
                    ЦифраСтрокой = Часть(2) + Часть(19) + _
                                   Часть(27)
                    ' пишем "триста "
                Case 3
                    ЦифраСтрокой = Часть(3) + Часть(30) + _
                                   Часть(27)
                    ' пишем "чотириста "
                Case 4
                    ЦифраСтрокой = Часть(4) + Часть(11) + _
                                   Часть(30) + Часть(27)
                    ' в остальных случаях пишем название цифры
                    ' плюс "сот "
                Case 5, 6, 7, 8, 9
                    ЦифраСтрокой = Часть(Цифра) + Часть(20) + _
                                   Часть(27)
            End Select
        End If
        ' Определяем названия единиц
        If (Место = 1) Or (Место = 4) Or _
            (Место = 7) Or (Место = 10) Then
            Select Case Цифра
                ' пишем  "один " для миллионов и миллиардов, одиниць и "одна " для остальных вариантов
                Case 1
                    If (Место = 7) Or (Место = 10) Or (Место = 1) Then
                        ЦифраСтрокой = Часть(1) + Часть(10) + Часть(27)
                    Else
                        ЦифраСтрокой = Часть(21) + Часть(27)
                    End If
                    ' пишем "два " для миллионов и миллиардов, одиниць и "двi " для остальных вариантов
                Case 2
                    If (Место = 7) Or (Место = 10) Or (Место = 1) Then
                        ЦифраСтрокой = Часть(18) + Часть(27)
                    Else
                        ЦифраСтрокой = Часть(2) + Часть(27)
                    End If
                    ' пишем "три "
                Case 3
                    ЦифраСтрокой = Часть(Цифра) + Часть(27)
                    ' пишем "чотири "
                Case 4
                    ЦифраСтрокой = Часть(4) + Часть(11) + _
                                   Часть(27)
                    ' пишем "сiм" или "вiсiм"
                Case 7, 8
                    ЦифраСтрокой = Часть(Цифра) + Часть(27)
                    ' в остальных случаях пишем название цифры
                Case 5, 6, 9
                    ЦифраСтрокой = Часть(Цифра) + Часть(12) + _
                                   Часть(27)
            End Select
            ' Определяем названия тысяч
            If Место = 4 Then
                Select Case Цифра
                    ' пишем "тисяч" только в том случае, если
                    ' хотя бы в одном разряде тысяч есть не нулевое
                    ' значение
                    Case 0
                        If Тысячи Then
                            ЦифраСтрокой = Часть(22) + Часть(27)
                        End If
                        ' пишем "одна тисяча "
                    Case 1
                        ЦифраСтрокой = Часть(21) + Часть(27) + _
                                       Часть(22) + Часть(23) + Часть(27)
                        ' пишем "дві тисячі "
                    Case 2
                        ЦифраСтрокой = Часть(2) + Часть(27) + _
                                       Часть(22) + Часть(24) + Часть(27)
                        ' добавляем "тисячі "
                    Case 3, 4
                        ЦифраСтрокой = ЦифраСтрокой + Часть(22) + _
                                       Часть(24) + Часть(27)
                        ' в остальных случаях добавляем "тисяч "
                    Case 5, 6, 7, 8, 9
                        ЦифраСтрокой = ЦифраСтрокой + Часть(22) + _
                                       Часть(27)
                End Select
                ' Сбрасываем значения тысяч, так как
                ' переходим к предыдущим разрядам
                Тысячи = Ложь
            End If
            ' Определяем названия миллионов
            If Место = 7 Then
                Select Case Цифра
                    ' пишем "мiльйонiв " только в том случае,
                    ' если хотя бы в одном разряде миллионов
                    ' есть не нулевое значение
                    Case 0
                        If Миллионы Then
                            ЦифраСтрокой = Часть(25) + Часть(26) + _
                                           Часть(27)
                        End If
                        ' добавляем "мiльйон "
                    Case 1
                        ЦифраСтрокой = ЦифраСтрокой + Часть(25) + _
                                       Часть(27)
                        ' добавляем "мiльйони "
                    Case 2, 3, 4
                        ЦифраСтрокой = ЦифраСтрокой + Часть(25) + _
                                       Часть(11) + Часть(27)
                        ' добавляем "мiльйонiв "
                    Case 5, 6, 7, 8, 9
                        ЦифраСтрокой = ЦифраСтрокой + Часть(25) + _
                                       Часть(26) + Часть(27)
                End Select
                ' Сбрасываем значения миллионов, так как
                ' переходим к предыдущим разрядам
                Миллионы = Ложь
            End If
            ' Определяем названия миллиардов
            If Место = 10 Then
                Select Case Цифра
                    ' пишем "мiлiардiв " только в том случае,
                    ' если хотя бы в одном разряде миллиардов
                    ' есть не нулевое значение
                    Case 0
                        If Миллиарды Then
                            ЦифраСтрокой = Часть(31) + Часть(26) + _
                                           Часть(27)
                        End If
                        ' добавляем "мiлiард "
                    Case 1
                        ЦифраСтрокой = ЦифраСтрокой + Часть(31) + _
                                       Часть(27)
                        ' добавляем "мiлiарди "
                    Case 2, 3, 4
                        ЦифраСтрокой = ЦифраСтрокой + Часть(31) + _
                                       Часть(11) + Часть(27)
                        ' добавляем "мiлiардiв "
                    Case 5, 6, 7, 8, 9
                        ЦифраСтрокой = ЦифраСтрокой + Часть(31) + _
                                       Часть(26) + Часть(27)
                End Select
                ' Сбрасываем значения миллиардов, так как
                ' переходим к предыдущим разрядам
                Миллиарды = Ложь
            End If
        End If
    End If
End Function

' функция возвращает число прописью
'
Function ЧислоПрописомЖР(Число)
    ' Присвоение значений массиву частей
    Часть(1) = "оди":       Часть(2) = "дві"
    Часть(3) = "три":       Часть(4) = "чотир"
    Часть(5) = "п'ят":       Часть(6) = "шiст"
    Часть(7) = "сiм":       Часть(8) = "вiсiм"
    Часть(9) = "дев'ят":     Часть(10) = "н"
    Часть(11) = "и":        Часть(12) = "ь"
    Часть(13) = "надцять":  Часть(14) = "дцять"
    Часть(15) = "сорок":    Часть(16) = "дев'яно"
    Часть(17) = "сто":      Часть(18) = "два"
    Часть(19) = "стi":      Часть(20) = "сот"
    Часть(21) = "одна":     Часть(22) = "тисяч"
    Часть(23) = "а":        Часть(24) = "i"
    Часть(25) = "мiльйон":  Часть(26) = "iв"
    Часть(27) = " ":        Часть(28) = "":
    Часть(29) = "десят":    Часть(30) = "ста"
    Часть(31) = "мільярд": Часть(32) = "нуль "
    ' Временные переменные вначале сбрасываются
    Тысячи = Ложь:      Миллионы = Ложь
    Миллиарды = Ложь:   ВторойДесяток = Ложь
    ' Отбрасываем дробную часть, если она есть
    Число = Fix(Число)
    ' Определяем длину исходного числа
    Длина = Len(Число)
    ' Цикл по всем цифрам числа, начиная с крайней
    ' левой до крайней правой
    For Позиция = Длина To 1 Step -1
        ' Добавляются очередные слова, описывающие
        ' текущую цифру
        ЧислоПрописомЖР = ЧислоПрописомЖР + _
                         ЦифраСтрокойЖР(Mid(Число, _
                         Длина - Позиция + 1, 1), _
                         Позиция)
    Next Позиция
    ' Алгоритм возвращает пустую строку при
    ' нулевом аргументе. Исправим это
    If ЧислоПрописомЖР = "" Then
        ЧислоПрописомЖР = Часть(32)
    End If
End Function
'
' Составление слов из частей по очередной
' цифре числа и по предистории работы
'
' Функция доступна только в текущем модуле
'
Private Function ЦифраСтрокойЖР(Цифра, Место) As String
    ' Если сотни или десятки миллиардов, то
    ' запомнить об этом для будущего
    If (Цифра <> 0) And ((Место = 11) Or _
        (Место = 12)) Then
        Миллиарды = Истина
    End If
    ' Если сотни или десятки миллионов, то
    ' запомнить об этом для будущего
    If (Цифра <> 0) And ((Место = 8) Or _
        (Место = 9)) Then
        Миллионы = Истина
    End If
    ' Если сотни или десятки тысяч, то
    ' запомнить об этом для будущего
    If (Цифра <> 0) And ((Место = 5) Or _
        (Место = 6)) Then
        Тысячи = Истина
    End If
    ' Если предыдущая цифра была единица
    ' в поле десятков, то выбираем
    If ВторойДесяток Then
        Select Case Цифра
            ' пишем "десять "
            Case 0
                ЦифраСтрокойЖР = Часть(29) + Часть(12) + _
                                Часть(27)
                ' пишем "дванадцять "
            Case 2
                ЦифраСтрокойЖР = Часть(18) + Часть(13) + _
                                Часть(27)
                ' в остальных случаях пишем название цифры
                ' плюс "надцять "
            Case Else
                ЦифраСтрокойЖР = Часть(Цифра) + Часть(13) + _
                                Часть(27)
        End Select
        ' Добавляем название разрядов
        Select Case Место
            Case 4
                ' добавляем "тисяч "
                ЦифраСтрокойЖР = ЦифраСтрокойЖР + Часть(22) + _
                                Часть(27)
                ' добавляем "мiльйонiв "
            Case 7
                ЦифраСтрокойЖР = ЦифраСтрокойЖР + Часть(25) + _
                                Часть(26) + Часть(27)
                ' добавляем "мiлiардiв "
            Case 10
                ЦифраСтрокойЖР = ЦифраСтрокойЖР + Часть(31) + _
                                Часть(26) + Часть(27)
        End Select
        ' Сбрасываем значения, так как переходим к
        ' предыдущим разрядам
        ВторойДесяток = Ложь:    Миллионы = Ложь
        Миллиарды = Ложь:        Тысячи = Ложь
        ' Во всех остальных случаях, то есть
        ' не для описания чисел второго десятка
    Else
        ' Определяем название десятков
        If (Место = 2) Or (Место = 5) Or _
            (Место = 8) Or (Место = 11) Then
            Select Case Цифра
                ' Запоминаем про второй десяток для
                ' подстановки при следующем входе
                Case 1
                    ВторойДесяток = Истина
                    ' пишем "двадцять "
                Case 2
                    ЦифраСтрокойЖР = Часть(18) + Часть(14) + _
                                    Часть(27)
                    ' пишем "тридцять "
                Case 3
                    ЦифраСтрокойЖР = Часть(Цифра) + Часть(14) + _
                                    Часть(27)
                    ' пишем "сорок "
                Case 4
                    ЦифраСтрокойЖР = Часть(15) + Часть(27)
                    ' пишем "дев'яносто "
                Case 9
                    ЦифраСтрокойЖР = Часть(16) + Часть(17) + _
                                    Часть(27)
                    ' в остальных случаях пишем название цифры
                    ' плюс "десят "
                Case 5, 6, 7, 8
                    ЦифраСтрокойЖР = Часть(Цифра) + Часть(29) + _
                                    Часть(27)
            End Select
        End If
        ' Определяем названия сотен
        If (Место = 3) Or (Место = 6) Or _
            (Место = 9) Or (Место = 12) Then
            Select Case Цифра
                ' пишем "сто "
                Case 1
                    ЦифраСтрокойЖР = Часть(17) + Часть(27)
                    ' пишем "двісті "
                Case 2
                    ЦифраСтрокойЖР = Часть(2) + Часть(19) + _
                                    Часть(27)
                    ' пишем "триста "
                Case 3
                    ЦифраСтрокойЖР = Часть(3) + Часть(30) + _
                                    Часть(27)
                    ' пишем "чотириста "
                Case 4
                    ЦифраСтрокойЖР = Часть(4) + Часть(11) + _
                                    Часть(30) + Часть(27)
                    ' в остальных случаях пишем название цифры
                    ' плюс "сот "
                Case 5, 6, 7, 8, 9
                    ЦифраСтрокойЖР = Часть(Цифра) + Часть(20) + _
                                    Часть(27)
            End Select
        End If
        ' Определяем названия единиц
        If (Место = 1) Or (Место = 4) Or _
            (Место = 7) Or (Место = 10) Then
            Select Case Цифра
                ' пишем  "один " для миллионов и миллиардов и "одна " для остальных вариантов
                Case 1
                    If (Место = 7) Or (Место = 10) Then
                        ЦифраСтрокойЖР = Часть(1) + Часть(10) + Часть(27)
                    Else
                        ЦифраСтрокойЖР = Часть(21) + Часть(27)
                    End If
                    ' пишем "два " для миллионов и миллиардов и "двi " для остальных вариантов
                Case 2
                    If (Место = 7) Or (Место = 10) Then
                        ЦифраСтрокойЖР = Часть(18) + Часть(27)
                    Else
                        ЦифраСтрокойЖР = Часть(2) + Часть(27)
                    End If
                    ' пишем "три "
                Case 3
                    ЦифраСтрокойЖР = Часть(Цифра) + Часть(27)
                    ' пишем "чотири "
                Case 4
                    ЦифраСтрокойЖР = Часть(4) + Часть(11) + _
                                    Часть(27)
                    ' пишем "сiм" или "вiсiм"
                Case 7, 8
                    ЦифраСтрокойЖР = Часть(Цифра) + Часть(27)
                    ' в остальных случаях пишем название цифры
                Case 5, 6, 9
                    ЦифраСтрокойЖР = Часть(Цифра) + Часть(12) + _
                                    Часть(27)
            End Select
            ' Определяем названия тысяч
            If Место = 4 Then
                Select Case Цифра
                    ' пишем "тисяч" только в том случае, если
                    ' хотя бы в одном разряде тысяч есть не нулевое
                    ' значение
                    Case 0
                        If Тысячи Then
                            ЦифраСтрокойЖР = Часть(22) + Часть(27)
                        End If
                        ' пишем "одна тисяча "
                    Case 1
                        ЦифраСтрокойЖР = Часть(21) + Часть(27) + _
                                        Часть(22) + Часть(23) + Часть(27)
                        ' пишем "дві тисячі "
                    Case 2
                        ЦифраСтрокойЖР = Часть(2) + Часть(27) + _
                                        Часть(22) + Часть(24) + Часть(27)
                        ' добавляем "тисячі "
                    Case 3, 4
                        ЦифраСтрокойЖР = ЦифраСтрокойЖР + Часть(22) + _
                                        Часть(24) + Часть(27)
                        ' в остальных случаях добавляем "тисяч "
                    Case 5, 6, 7, 8, 9
                        ЦифраСтрокойЖР = ЦифраСтрокойЖР + Часть(22) + _
                                        Часть(27)
                End Select
                ' Сбрасываем значения тысяч, так как
                ' переходим к предыдущим разрядам
                Тысячи = Ложь
            End If
            ' Определяем названия миллионов
            If Место = 7 Then
                Select Case Цифра
                    ' пишем "мiльйонiв " только в том случае,
                    ' если хотя бы в одном разряде миллионов
                    ' есть не нулевое значение
                    Case 0
                        If Миллионы Then
                            ЦифраСтрокойЖР = Часть(25) + Часть(26) + _
                                            Часть(27)
                        End If
                        ' добавляем "мiльйон "
                    Case 1
                        ЦифраСтрокойЖР = ЦифраСтрокойЖР + Часть(25) + _
                                        Часть(27)
                        ' добавляем "мiльйони "
                    Case 2, 3, 4
                        ЦифраСтрокойЖР = ЦифраСтрокойЖР + Часть(25) + _
                                        Часть(11) + Часть(27)
                        ' добавляем "мiльйонiв "
                    Case 5, 6, 7, 8, 9
                        ЦифраСтрокойЖР = ЦифраСтрокойЖР + Часть(25) + _
                                        Часть(26) + Часть(27)
                End Select
                ' Сбрасываем значения миллионов, так как
                ' переходим к предыдущим разрядам
                Миллионы = Ложь
            End If
            ' Определяем названия миллиардов
            If Место = 10 Then
                Select Case Цифра
                    ' пишем "мiлiардiв " только в том случае,
                    ' если хотя бы в одном разряде миллиардов
                    ' есть не нулевое значение
                    Case 0
                        If Миллиарды Then
                            ЦифраСтрокойЖР = Часть(31) + Часть(26) + _
                                            Часть(27)
                        End If
                        ' добавляем "мiлiард "
                    Case 1
                        ЦифраСтрокойЖР = ЦифраСтрокойЖР + Часть(31) + _
                                        Часть(27)
                        ' добавляем "мiлiарди "
                    Case 2, 3, 4
                        ЦифраСтрокойЖР = ЦифраСтрокойЖР + Часть(31) + _
                                        Часть(11) + Часть(27)
                        ' добавляем "мiлiардiв "
                    Case 5, 6, 7, 8, 9
                        ЦифраСтрокойЖР = ЦифраСтрокойЖР + Часть(31) + _
                                        Часть(26) + Часть(27)
                End Select
                ' Сбрасываем значения миллиардов, так как
                ' переходим к предыдущим разрядам
                Миллиарды = Ложь
            End If
        End If
    End If
End Function

Function СумаПрописомАнгл(ByVal MyNumber, ByVal CurrencyCode) As String

    Const USD = "USD"
    Const UAH = "UAH"
    Const EUR = "EUR"
    Const GBP = "GBP"
    Dim Dollars As String, Cents As String, temp As String, CurrencyCodeName As String
    Dim DecimalPlace As Integer, count As Integer
    Dim Place(9) As String
    
    Place(2) = " Thousand "
    Place(3) = " Million "
    Place(4) = " Billion "
    Place(5) = " Trillion "
    
    MyNumber = Trim(str(MyNumber))
    DecimalPlace = InStr(MyNumber, ".")
    If DecimalPlace > 0 Then
        '        Cents = GetTens(Left(Mid(MyNumber, DecimalPlace + 1) & "00", 2))
        Cents = Left(Mid(MyNumber, DecimalPlace + 1) & "00", 2)
        
        MyNumber = Trim(Left(MyNumber, DecimalPlace - 1))
    End If
    
    count = 1
    Do While MyNumber <> ""
        temp = GetHundreds(Right(MyNumber, 3))
        If temp <> "" Then Dollars = temp & Place(count) & Dollars
        If Len(MyNumber) > 3 Then
            MyNumber = Left(MyNumber, Len(MyNumber) - 3)
        Else
            MyNumber = ""
        End If
        count = count + 1
    Loop
    
    Select Case CurrencyCode
        Case USD
            Select Case Dollars
                Case ""
                    Dollars = "No Dollars"
                Case "One"
                    Dollars = "One Dollar"
                Case Else
                    Dollars = Dollars & " Dollars"
            End Select
            
            Select Case Cents
                Case ""
                    Cents = " and 00 Cents"
                    '                    Cents = " and No Cents"
                    '                Case "One"
                Case "01"
                    Cents = " and " & Cents & " Cent" '" and One Cent"
                Case Else
                    Cents = " and " & Cents & " Cents"
            End Select
            
        Case UAH
            Select Case Dollars
                Case ""
                    Dollars = "No Hryvnias"
                Case "One"
                    Dollars = "One Hryvnia"
                Case Else
                    Dollars = Dollars & " Hryvnias"
            End Select
            
            Select Case Cents
                Case ""
                    Cents = " and 00 Kopecks"
                    '                    Cents = " and No Kopecks"
                    '                Case "One"
                Case "01"
                    Cents = " and " & Cents & " Kopeck" '" and One Kopeck"
                Case Else
                    Cents = " and " & Cents & " Kopecks"
            End Select
            
        Case EUR
            Select Case Dollars
                Case ""
                    Dollars = "No Euros"
                Case "One"
                    Dollars = "One Euro"
                Case Else
                    Dollars = Dollars & " Euros"
            End Select
            
            Select Case Cents
                Case ""
                    Cents = " and 00 Eurocents"
                    '                    Cents = " and No Eurocents"
                    '                Case "One"
                Case "01"
                    Cents = " and " & Cents & " Eurocent" '" and One Eurocent"
                Case Else
                    Cents = " and " & Cents & " Eurocents"
            End Select
        Case GBP
            Select Case Dollars
                Case ""
                    Dollars = "No Pounds"
                Case "One"
                    Dollars = "One Pound"
                Case Else
                    Dollars = Dollars & " Pounds"
            End Select
            
            Select Case Cents
                Case ""
                    Cents = " and 00 Pence"
                    '                    Cents = " and No Pence"
                    '                Case "One"
                Case "01"
                    Cents = " and & Cents & Penny" '" and One Penny"
                Case Else
                    Cents = " and " & Cents & " Pence"
            End Select
            
    End Select
    '    СумаПрописомАнгл = Dollars & Cents
    СумаПрописомАнгл = Left$(Dollars & Cents, 1) & LCase(Mid$(Dollars & Cents, 2, Len(Dollars & Cents)))
End Function

Function GetHundreds(ByVal MyNumber)
    Dim result As String
    If val(MyNumber) = 0 Then Exit Function
    MyNumber = Right("000" & MyNumber, 3)
    ' Convert the hundreds place.
    If Mid(MyNumber, 1, 1) <> "0" Then
        result = GetDigit(Mid(MyNumber, 1, 1)) & " Hundred "
    End If
    ' Convert the tens and ones place.
    If Mid(MyNumber, 2, 1) <> "0" Then
        result = result & GetTens(Mid(MyNumber, 2))
    Else
        result = result & GetDigit(Mid(MyNumber, 3))
    End If
    GetHundreds = result
End Function

Function GetTens(TensText)
    Dim result As String
    result = ""                                   ' Null out the temporary function value.
    If val(Left(TensText, 1)) = 1 Then            ' If value between 10-19&hellip;
    Select Case val(TensText)
        Case 10: result = "Ten"
        Case 11: result = "Eleven"
        Case 12: result = "Twelve"
        Case 13: result = "Thirteen"
        Case 14: result = "Fourteen"
        Case 15: result = "Fifteen"
        Case 16: result = "Sixteen"
        Case 17: result = "Seventeen"
        Case 18: result = "Eighteen"
        Case 19: result = "Nineteen"
        Case Else
    End Select
Else                                              ' If value between 20-99&hellip;
    Select Case val(Left(TensText, 1))
        Case 2: result = "Twenty "
        Case 3: result = "Thirty "
        Case 4: result = "Forty "
        Case 5: result = "Fifty "
        Case 6: result = "Sixty "
        Case 7: result = "Seventy "
        Case 8: result = "Eighty "
        Case 9: result = "Ninety "
        Case Else
    End Select
    result = result & GetDigit _
    (Right(TensText, 1))                          ' Retrieve ones place.
End If
GetTens = result
End Function

Function GetDigit(Digit)
    Select Case val(Digit)
        Case 1: GetDigit = "One"
        Case 2: GetDigit = "Two"
        Case 3: GetDigit = "Three"
        Case 4: GetDigit = "Four"
        Case 5: GetDigit = "Five"
        Case 6: GetDigit = "Six"
        Case 7: GetDigit = "Seven"
        Case 8: GetDigit = "Eight"
        Case 9: GetDigit = "Nine"
        Case Else: GetDigit = ""
    End Select
End Function

Function GetDigit2(Digit)                         'включаючи 0
    Select Case val(Digit)
        Case 0: GetDigit2 = "Zero"
        Case 1: GetDigit2 = "One"
        Case 2: GetDigit2 = "Two"
        Case 3: GetDigit2 = "Three"
        Case 4: GetDigit2 = "Four"
        Case 5: GetDigit2 = "Five"
        Case 6: GetDigit2 = "Six"
        Case 7: GetDigit2 = "Seven"
        Case 8: GetDigit2 = "Eight"
        Case 9: GetDigit2 = "Nine"
        Case Else: GetDigit2 = ""
    End Select
End Function

Function ЧислоПрописомАнгл(ByVal numIn)
    Dim LSide, RSide, temp, DecPlace, count, oNum
    oNum = numIn
    ReDim Place(9) As String
    Place(2) = " Thousand "
    Place(3) = " Million "
    Place(4) = " Billion "
    Place(5) = " Trillion "
    numIn = Trim(str(numIn))                      'String representation of amount
    ' Edit 2.(0)/Internationalisation
    ' Don't change point sign here as the above assignment preserves the point!
    DecPlace = InStr(numIn, ".")                  'Pos of dec place 0 if none
    If DecPlace > 0 Then                          'Convert Right & set numIn
        RSide = GetTens(Left(Mid(numIn, DecPlace + 1) & "00", 2))
        numIn = Trim(Left(numIn, DecPlace - 1))
    End If
    RSide = numIn
    count = 1
    Do While numIn <> ""
        temp = GetHundreds(Right(numIn, 3))
        If temp <> "" Then LSide = temp & Place(count) & LSide
        If Len(numIn) > 3 Then
            numIn = Left(numIn, Len(numIn) - 3)
        Else
            numIn = ""
        End If
        count = count + 1
    Loop
    
    '    ЧислоПрописомАнгл = LSide
    ЧислоПрописомАнгл = Left$(LSide, 1) & LCase(Mid$(LSide, 2, Len(LSide)))
    
    If InStr(oNum, Application.DecimalSeparator) > 0 Then ' << Edit 2.(1)
        ЧислоПрописомАнгл = ЧислоПрописомАнгл & " point " & fractionWords(oNum)
        '        ЧислоПрописомАнгл = ЧислоПрописомАнгл & " point " & fractionWords(oNum)
        ЧислоПрописомАнгл = Left$(ЧислоПрописомАнгл, 1) & LCase(Mid$(ЧислоПрописомАнгл, 2, Len(ЧислоПрописомАнгл)))
        
    End If
    
End Function
Function fractionWords(n) As String
    Dim fraction As String, x As Long
    fraction = Split(n, Application.DecimalSeparator)(1) ' << Edit 2.(2)
    For x = 1 To Len(fraction)
        If fractionWords <> "" Then fractionWords = fractionWords & " "
        fractionWords = fractionWords & GetDigit2(Mid(fraction, x, 1))
    Next x
End Function

Sub RegisterUDF()
    Dim s      As String
    s = "Конвертує число в суму прописом АНГЛІЙСЬКОЮ мовою" & vbNewLine & _
        "Валюти:" & vbNewLine & _
        "Українська гривня - ""UAH""" & vbNewLine & _
        "Американський долар - ""USD""" & vbNewLine & _
        "Євро - ""EUR""" & vbNewLine & _
        "Англійський фунт - ""GBP""" & vbNewLine & _
        "Перший аргумент - число(або посилання на клітинку)" & vbNewLine & _
        "Другий аргумент - валюта"
    
    Application.MacroOptions Macro:="СумаПрописомАнгл", Description:=s, Category:=9
End Sub

Sub UnregisterUDF()
    Application.MacroOptions Macro:="СумаПрописомАнгл", Description:=Empty, Category:=Empty
End Sub


```
