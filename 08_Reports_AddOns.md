# 08 — Модулі звітів: Додатки та Відомості

У цьому документі наведено вичерпний опис та вихідний код модулів розширеної звітності:
1. **add_1_one_Troop.bas** — Додаток 1: Зведена відомість руху для підрозділу (v1)
2. **add_1_one_Troop2.bas** — Додаток 1: Розширена відомість з групуванням по МВО та масовим експортом (v2, 69 КБ)
3. **Add14.bas** — Відомість № 14: Розгорнутий звіт залишків по всіх субрахунках для всієї частини
4. **Add14OneTroop.bas** — Відомість № 14: Звіт по субрахунках для одного вибраного підрозділу
5. **Add47.bas** — Додаток 47: Відомість надходження та передачі конкретного найменування
6. **add47PNP.bas** — Додаток 47 (ПН-підрозділ): Форма обліку надходження від підрозділів
7. **ReportingMove.bas** — Звітна відомість ф. 18 (наказ МОУ № 440)

---

## 1. add_1_one_Troop.bas — Додаток 1 (версія 1.0)
### 1.1. Призначення
Формує класичний бухгалтерський звіт руху: «Залишок на початок періоду + Надійшло за період - Вибуло за період = Залишок на кінець періоду». Таблиця `tbAddOne` на аркуші «Додаток 1».

### 1.2. Вихідний код add_1_one_Troop.bas
```vba
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

```

---

## 2. add_1_one_Troop2.bas — Додаток 1 (версія 2.0, розширена)
### 2.1. Архітектура та можливості
Найбільший звітний модуль системи (69 КБ). Забезпечує:
- Розрахунок залишків на будь-який період часу;
- Автоматичне розбиття та групування залишків підрозділу по кожній матеріально-відповідальній особі (`Add1_BuildMVORowGroupsForUnit`);
- Масовий експорт окремих книг Excel для кожного МВО у підкаталог `А1_Відомос\` (`Add1_SaveOneUnitByMVO`);
- Масовий розрахунок одразу для ВСІХ підрозділів військової частини (`CreateAddOneForAllTroops`).

### 2.2. Вихідний код add_1_one_Troop2.bas
```vba
Attribute VB_Name = "add_1_one_Troop2"
Option Explicit

'===============================================================================
' МОДУЛЬ: ДОДАТОК 1
'
' ПРИЗНАЧЕННЯ
' 1. CreateAddOneForOneTroop2
'    - вибір одного підрозділу;
'    - вибір періоду;
'    - формування "Довідка по підрозділу 2.0";
'    - поділ довідки за МВО;
'    - окремий файл Додатка 1 на КОЖНОГО МВО;
'    - у кожному файлі тільки майно цього МВО;
'    - користувач вибирає папку збереження;
'    - створюється окрема папка підрозділу.
'
' 2. CreateAddOneForAllTroops
'    - період вводиться один раз;
'    - "Довідка по підрозділу 2.0" формується ОДИН раз за всіма підрозділами;
'    - кожен підрозділ ділиться за МВО;
'    - створюється окремий файл на кожного МВО;
'    - збереження:
'          Д1_експорт\
'              <батальйон>\
'                  Д1_<підрозділ>_<МВО>.xlsx
'
' ЗОВНІШНІ ЗАЛЕЖНОСТІ
' - UserForm1
' - Public Sub BuildResiduseDetailedReport(...) у модулі ResiduseOneTroop
'
' ОЧІКУВАНА СТРУКТУРА "Довідка по підрозділу 2.0"
' C = Підрозділ
' D = МВО
' E = Найменування
' F = Ціна
' G = Надійшло
' H = Вибуло
' I = Залишок
' J = Примітка
' K = Категорія
' L = Одиниця виміру
' M = Код номенклатури
' N = Залишок на початок (службова прихована колонка)
'===============================================================================

Private Const EXPORT_ROOT_FOLDER_NAME As String = "Д1_експорт"
Private Const KEY_SEP As String = "||"
Private Const NO_MVO_LABEL As String = "Без МВО"


'===============================================================================
' 1. СТВОРИТИ ДОДАТОК 1 ЗА ОДНИМ ПІДРОЗДІЛОМ
'    ОКРЕМИЙ ФАЙЛ НА КОЖНОГО МВО
'===============================================================================
Public Sub CreateAddOneForOneTroop2()


    'Перевіряємо наявність усіх 3 кнопок і створюємо відсутні.
    EnsureAddOneButtons
    Dim selectedUnit As String
    Dim startDate As Date
    Dim endDate As Date

    Dim wsResiduse As Worksheet
    Dim wsTemplate As Worksheet

    Dim dictGroups As Object
    Dim dictMvoNames As Object

    Dim answer As VbMsgBoxResult

'    On Error GoTo ErrHandler

    '---------------------------------------------------------------------------
    ' 1. Вибір підрозділу
    '---------------------------------------------------------------------------
    If Not Add1_SelectUnit(selectedUnit) Then Exit Sub

    '---------------------------------------------------------------------------
    ' 2. Вибір періоду
    '---------------------------------------------------------------------------
    If Not Add1_AskPeriod(startDate, endDate) Then Exit Sub

    Application.ScreenUpdating = False

    '---------------------------------------------------------------------------
    ' 3. Спочатку формуємо "Довідка по підрозділу 2.0"
    '---------------------------------------------------------------------------
    BuildResiduseDetailedReport selectedUnit, startDate, endDate, False

    Set wsResiduse = ThisWorkbook.Worksheets("Довідка по підрозділу 2.0")
    Set wsTemplate = ThisWorkbook.Worksheets("Додаток 1")

    '---------------------------------------------------------------------------
    ' 4. Розкладаємо конкретні рядки довідки за МВО
    '---------------------------------------------------------------------------
    Add1_BuildMVORowGroupsForUnit _
        wsResiduse, _
        selectedUnit, _
        dictGroups, _
        dictMvoNames

    Application.ScreenUpdating = True

    If dictGroups.count = 0 Then
        MsgBox _
            "У сформованій довідці немає майна для підрозділу:" & vbCrLf & _
            selectedUnit, _
            vbExclamation, _
            "Додаток 1"
        Exit Sub
    End If

    '---------------------------------------------------------------------------
    ' 5. Підтвердження
    '---------------------------------------------------------------------------
    answer = MsgBox( _
        "Підрозділ: " & selectedUnit & vbCrLf & vbCrLf & _
        "Знайдено МВО: " & dictGroups.count & vbCrLf & _
        "Буде створено " & dictGroups.count & " окремих файл(и/ів)." & vbCrLf & _
        "У кожному файлі буде тільки майно відповідного МВО." & vbCrLf & vbCrLf & _
        "Продовжити?", _
        vbYesNo + vbQuestion, _
        "Створення Додатків 1")

    If answer <> vbYes Then Exit Sub

    '---------------------------------------------------------------------------
    ' 6. Вибір папки та створення окремих файлів
    '---------------------------------------------------------------------------
    Add1_SaveOneUnitByMVO _
        wsResiduse, _
        wsTemplate, _
        selectedUnit, _
        startDate, _
        endDate, _
        dictGroups, _
        dictMvoNames

    Exit Sub

'ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

    MsgBox _
        "Помилка: " & Err.Description, _
        vbExclamation, _
        "CreateAddOneForOneTroop2"

End Sub


'===============================================================================
' 2. СТВОРИТИ ДОДАТКИ 1 ЗА ВСІМА ПІДРОЗДІЛАМИ
'
' Довідка формується ОДИН раз за всіма підрозділами.
' Далі кожен підрозділ автоматично ділиться за МВО.
'===============================================================================
Public Sub CreateAddOneForAllTroops()


    'Перевіряємо наявність усіх 3 кнопок і створюємо відсутні.
    EnsureAddOneButtons
    Dim wb As Workbook
    Dim wsResiduse As Worksheet
    Dim wsTemplate As Worksheet

    Dim startDate As Date
    Dim endDate As Date

    Dim dictUnits As Object
    Dim unitKey As Variant

    Dim dictGroups As Object
    Dim dictMvoNames As Object

    Dim oldCalc As XlCalculation
    Dim createdCount As Long
    Dim skippedUnits As Long

    On Error GoTo ErrHandler

    Set wb = ThisWorkbook

    If Len(wb.Path) = 0 Then
        MsgBox "Спочатку збережіть книгу.", vbExclamation
        Exit Sub
    End If

    'Період вводимо один раз.
    If Not Add1_AskPeriod(startDate, endDate) Then Exit Sub

    oldCalc = Application.Calculation

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Set wsResiduse = wb.Worksheets("Довідка по підрозділу 2.0")
    Set wsTemplate = wb.Worksheets("Додаток 1")

    '---------------------------------------------------------------------------
    ' Головна оптимізація:
    ' один прохід по AllData для всіх підрозділів.
    '---------------------------------------------------------------------------
    BuildResiduseDetailedReport vbNullString, startDate, endDate, True

    Set dictUnits = Add1_GetUnitsFromResiduse(wsResiduse)

    If dictUnits.count = 0 Then
        MsgBox "У сформованій довідці немає даних для експорту.", vbExclamation
        GoTo SafeExit
    End If

    For Each unitKey In dictUnits.Keys

        Add1_BuildMVORowGroupsForUnit _
            wsResiduse, _
            CStr(unitKey), _
            dictGroups, _
            dictMvoNames

        If dictGroups.count > 0 Then

            createdCount = createdCount + _
                Add1_SaveAllModeUnitByMVO( _
                    wsResiduse, _
                    wsTemplate, _
                    CStr(unitKey), _
                    startDate, _
                    endDate, _
                    dictGroups, _
                    dictMvoNames, _
                    wb.Path)

        Else
            skippedUnits = skippedUnits + 1
        End If

    Next unitKey

SafeExit:
    Application.Calculation = oldCalc
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox _
        "Готово." & vbCrLf & vbCrLf & _
        "Створено файлів: " & createdCount & vbCrLf & _
        "Підрозділів без даних: " & skippedUnits & vbCrLf & vbCrLf & _
        "Папка:" & vbCrLf & _
        wb.Path & Application.PathSeparator & EXPORT_ROOT_FOLDER_NAME, _
        vbInformation, _
        "Додаток 1"

    Exit Sub

ErrHandler:
    On Error Resume Next

    Application.Calculation = oldCalc
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox _
        "Помилка: " & Err.Description, _
        vbExclamation, _
        "CreateAddOneForAllTroops"

End Sub

Public Sub CreateAddOneFromCurrentResiduse()


    'Перевіряємо наявність усіх 3 кнопок і створюємо відсутні.
    EnsureAddOneButtons
    Dim wb As Workbook
    Dim wsResiduse As Worksheet
    Dim wsTemplate As Worksheet

    Dim startDate As Date
    Dim endDate As Date

    Dim dictUnits As Object
    Dim dictGroups As Object
    Dim dictMvoNames As Object

    Dim unitKey As Variant

    Dim totalUnits As Long
    Dim totalFiles As Long
    Dim createdCount As Long
    Dim skippedUnits As Long

    Dim answer As VbMsgBoxResult
    Dim oldCalc As XlCalculation

    On Error GoTo ErrHandler

    Set wb = ThisWorkbook

    If Len(wb.Path) = 0 Then
        MsgBox "Спочатку збережіть книгу.", vbExclamation
        Exit Sub
    End If

    Set wsResiduse = wb.Worksheets("Довідка по підрозділу 2.0")
    Set wsTemplate = wb.Worksheets("Додаток 1")

    '===========================================================================
    ' 1. БЕРЕМО ПІДРОЗДІЛИ ТІЛЬКИ З ПОТОЧНОЇ ДОВІДКИ
    '
    ' НІЧОГО НЕ ПЕРЕРАХОВУЄМО ІЗ "БАЗА".
    ' НІЧОГО НЕ ВІДНОВЛЮЄМО.
    '
    ' Якщо користувач видалив зайві підрозділи/рядки з
    ' "Довідка по підрозділу 2.0", вони у формування не потраплять.
    '===========================================================================
    Set dictUnits = Add1_GetUnitsFromResiduse(wsResiduse)

    If dictUnits.count = 0 Then
        MsgBox _
            "На аркуші ""Довідка по підрозділу 2.0"" немає даних для формування Додатків 1.", _
            vbExclamation, _
            "Додаток 1"
        Exit Sub
    End If

    '===========================================================================
    ' 2. РАХУЄМО, СКІЛЬКИ РЕАЛЬНО БУДЕ ФАЙЛІВ
    '
    ' Один підрозділ може мати декілька МВО.
    ' Один МВО = один окремий файл.
    '===========================================================================
    totalUnits = dictUnits.count
    totalFiles = 0

    For Each unitKey In dictUnits.Keys

        Add1_BuildMVORowGroupsForUnit _
            wsResiduse, _
            CStr(unitKey), _
            dictGroups, _
            dictMvoNames

        totalFiles = totalFiles + dictGroups.count

    Next unitKey

    If totalFiles = 0 Then
        MsgBox _
            "У поточній довідці не знайдено МВО/майна для формування.", _
            vbExclamation, _
            "Додаток 1"
        Exit Sub
    End If

    '===========================================================================
    ' 3. ПЕРІОД
    '
    ' ВАЖЛИВО:
    ' Дані "Надійшло / Вибуло / Залишок" уже розраховані в довідці.
    ' Ці дати потрібні для шапки Додатка 1.
    '
    ' Тому потрібно вказати ТОЙ САМИЙ період,
    ' за який була сформована поточна довідка.
    '===========================================================================
    If Not Add1_AskPeriod(startDate, endDate) Then Exit Sub

    '===========================================================================
    ' 4. ПІДТВЕРДЖЕННЯ
    '===========================================================================
    answer = MsgBox( _
        "Додатки 1 будуть сформовані ТІЛЬКИ з поточного вмісту аркуша" & vbCrLf & _
        """Довідка по підрозділу 2.0""." & vbCrLf & vbCrLf & _
        "Підрозділів у довідці: " & totalUnits & vbCrLf & _
        "Окремих файлів за МВО: " & totalFiles & vbCrLf & vbCrLf & _
        "Видалені вами рядки та підрозділи відновлюватися не будуть." & vbCrLf & vbCrLf & _
        "Продовжити?", _
        vbYesNo + vbQuestion, _
        "Створити з поточної довідки")

    If answer <> vbYes Then Exit Sub

    '===========================================================================
    ' 5. ЕКСПОРТ
    '
    ' Зберігаємо за існуючою структурою:
    '
    ' Д1_експорт\
    '     <батальйон>\
    '         Д1_<підрозділ>_<МВО>.xlsx
    '===========================================================================
    oldCalc = Application.Calculation

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    createdCount = 0
    skippedUnits = 0

    For Each unitKey In dictUnits.Keys

        'Знову формуємо групи МВО, але ВИКЛЮЧНО
        'з тих рядків, які зараз реально є в довідці.
        Add1_BuildMVORowGroupsForUnit _
            wsResiduse, _
            CStr(unitKey), _
            dictGroups, _
            dictMvoNames

        If dictGroups.count > 0 Then

            createdCount = createdCount + _
                Add1_SaveAllModeUnitByMVO( _
                    wsResiduse, _
                    wsTemplate, _
                    CStr(unitKey), _
                    startDate, _
                    endDate, _
                    dictGroups, _
                    dictMvoNames, _
                    wb.Path)

        Else
            skippedUnits = skippedUnits + 1
        End If

    Next unitKey

SafeExit:
    Application.Calculation = oldCalc
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox _
        "Готово." & vbCrLf & vbCrLf & _
        "Використано лише поточний вміст ""Довідка по підрозділу 2.0""." & vbCrLf & _
        "Підрозділів у довідці: " & totalUnits & vbCrLf & _
        "Створено окремих файлів за МВО: " & createdCount & vbCrLf & _
        "Підрозділів без даних: " & skippedUnits & vbCrLf & vbCrLf & _
        "Папка:" & vbCrLf & _
        wb.Path & Application.PathSeparator & EXPORT_ROOT_FOLDER_NAME, _
        vbInformation, _
        "Додаток 1"

    Exit Sub

ErrHandler:
    On Error Resume Next

    Application.Calculation = oldCalc
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox _
        "Помилка: " & Err.Description, _
        vbExclamation, _
        "CreateAddOneFromCurrentResiduse"

End Sub
'===============================================================================
' 3. ВИБІР ПІДРОЗДІЛУ
'===============================================================================
Private Function Add1_SelectUnit(ByRef selectedUnit As String) As Boolean

    Dim wsData As Worksheet
    Dim tbAllData As ListObject

    Dim dictUnits As Object
    Dim cell As Range
    Dim unitKey As Variant

    Dim unitArray() As Variant
    Dim frm As UserForm1
    Dim i As Long

    Set wsData = ThisWorkbook.Worksheets("База")
    Set tbAllData = wsData.ListObjects("AllData")
    Set dictUnits = CreateObject("Scripting.Dictionary")

    'Звідки
    For Each cell In tbAllData.ListColumns("Звідки").DataBodyRange
        If Len(Trim$(CStr(cell.value))) > 0 Then
            If Not dictUnits.exists(CStr(cell.value)) Then
                dictUnits.Add CStr(cell.value), Nothing
            End If
        End If
    Next cell

    'Куди
    For Each cell In tbAllData.ListColumns("Куди").DataBodyRange
        If Len(Trim$(CStr(cell.value))) > 0 Then
            If Not dictUnits.exists(CStr(cell.value)) Then
                dictUnits.Add CStr(cell.value), Nothing
            End If
        End If
    Next cell

    If dictUnits.count = 0 Then
        MsgBox "У базі не знайдено підрозділів.", vbExclamation
        Exit Function
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

    If frm.ListBox1.ListIndex = -1 Then
        Unload frm
        Exit Function
    End If

    selectedUnit = CStr(frm.ListBox1.value)

    Unload frm

    Add1_SelectUnit = True

End Function


'===============================================================================
' 4. ВИБІР ПЕРІОДУ
'===============================================================================
Private Function Add1_AskPeriod(ByRef startDate As Date, ByRef endDate As Date) As Boolean

    Dim ws As Worksheet
    Dim defaultFrom As Variant
    Dim defaultTo As Variant
    Dim v As Variant

    Set ws = ThisWorkbook.Worksheets("Додаток 1")

    On Error Resume Next
    defaultFrom = ws.Range("AddOnePeriodFrom").value
    defaultTo = ws.Range("AddOnePeriodTo").value
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

    'Запам'ятовуємо останній період.
    On Error Resume Next
    ws.Range("AddOnePeriodFrom").value = startDate
    ws.Range("AddOnePeriodTo").value = endDate
    ws.Range("AddOnePeriod").value = _
        Format$(startDate, "dd.mm.yyyy") & "-" & _
        Format$(endDate, "dd.mm.yyyy")
    On Error GoTo 0

    Add1_AskPeriod = True

End Function


'===============================================================================
' 5. РОЗКЛАСТИ КОНКРЕТНІ РЯДКИ ПІДРОЗДІЛУ ЗА МВО
'
' dictGroups:
'   key   = нормалізований МВО
'   item  = Collection номерів рядків "Довідка по підрозділу 2.0"
'
' dictMvoNames:
'   key   = нормалізований МВО
'   item  = оригінальний текст МВО для підпису та назви файла
'
' Завдяки Collection конкретних рядків майно різних МВО фізично не змішується.
'===============================================================================
Private Sub Add1_BuildMVORowGroupsForUnit( _
    ByVal wsResiduse As Worksheet, _
    ByVal selectedUnit As String, _
    ByRef dictGroups As Object, _
    ByRef dictMvoNames As Object)

    Dim lastRow As Long
    Dim r As Long

    Dim rowUnit As String
    Dim rawMVO As String
    Dim displayMVO As String
    Dim mvoKey As String

    Dim rowsForMVO As Collection

    Set dictGroups = CreateObject("Scripting.Dictionary")
    Set dictMvoNames = CreateObject("Scripting.Dictionary")

    lastRow = wsResiduse.Cells(wsResiduse.Rows.count, "C").End(xlUp).Row

    For r = 8 To lastRow

        rowUnit = Add1_NormalizeText(wsResiduse.Cells(r, "C").value)

        If StrComp( _
            rowUnit, _
            Add1_NormalizeText(selectedUnit), _
            vbTextCompare) = 0 Then

            rawMVO = Trim$(CStr(wsResiduse.Cells(r, "D").value))

            If Len(Add1_NormalizeText(rawMVO)) = 0 Then
                displayMVO = NO_MVO_LABEL
            Else
                displayMVO = rawMVO
            End If

            mvoKey = Add1_NormalizeText(displayMVO)

            If Not dictGroups.exists(mvoKey) Then

                Set rowsForMVO = New Collection

                dictGroups.Add mvoKey, rowsForMVO
                dictMvoNames.Add mvoKey, displayMVO

            Else

                Set rowsForMVO = dictGroups(mvoKey)

            End If

            'Критично: зберігаємо конкретний номер рядка саме цього МВО.
            rowsForMVO.Add r

        End If

    Next r

End Sub


'===============================================================================
' 6. ЗБЕРЕЖЕННЯ ОДНОГО ПІДРОЗДІЛУ
'
' Користувач один раз вибирає батьківську папку.
'
' Створюється:
' <обрана папка>\
'   Д1_<підрозділ>\
'       Д1_<підрозділ>_<МВО1>.xlsx
'       Д1_<підрозділ>_<МВО2>.xlsx
'===============================================================================
Private Sub Add1_SaveOneUnitByMVO( _
    ByVal wsResiduse As Worksheet, _
    ByVal wsTemplate As Worksheet, _
    ByVal selectedUnit As String, _
    ByVal startDate As Date, _
    ByVal endDate As Date, _
    ByVal dictGroups As Object, _
    ByVal dictMvoNames As Object)

    Dim fd As FileDialog

    Dim parentFolder As String
    Dim unitFolderPath As String

    Dim mvoKey As Variant
    Dim mvoDisplay As String
    Dim rowsForMVO As Collection

    Dim exportWb As Workbook
    Dim wsExport As Worksheet

    Dim fileName As String
    Dim filePath As String

    Dim answer As VbMsgBoxResult
    Dim createdCount As Long
    Dim skippedCount As Long

    Dim currentReportNumber As String
    Dim nextReportNumber As String

'    On Error GoTo ErrHandler

    Set fd = Application.FileDialog(msoFileDialogFolderPicker)

    With fd
        .Title = "Виберіть папку для збереження Додатків 1"
        .AllowMultiSelect = False

        If Len(ThisWorkbook.Path) > 0 Then
            .InitialFileName = ThisWorkbook.Path & Application.PathSeparator
        End If

        If .show <> -1 Then Exit Sub

        parentFolder = .SelectedItems(1)
    End With

    unitFolderPath = _
        parentFolder & Application.PathSeparator & _
        "Д1_" & Add1_SanitizePathName(selectedUnit)

    If Dir(unitFolderPath, vbDirectory) = vbNullString Then
        MkDir unitFolderPath
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    'Поточний номер беремо з Додаток 1!E7.
    'Кожен УСПІШНО створений окремий файл отримує наступний номер.
    currentReportNumber = Add1_GetCurrentReportNumber()

    For Each mvoKey In dictGroups.Keys

        mvoDisplay = CStr(dictMvoNames(mvoKey))
        Set rowsForMVO = dictGroups(mvoKey)

        'Наступний номер відомості для цього файла.
        nextReportNumber = Add1_IncrementReportNumber(currentReportNumber)

        'Кожен МВО починається з чистої копії шаблону.
        wsTemplate.Copy

        Set exportWb = ActiveWorkbook
        Set wsExport = exportWb.Worksheets(1)

        'Заповнюємо ТІЛЬКИ конкретними рядками цього МВО.
        Add1_FillFromExactRows _
            wsResiduse, _
            wsExport, _
            selectedUnit, _
            mvoDisplay, _
            rowsForMVO, _
            startDate, _
            endDate, _
            nextReportNumber

        On Error Resume Next
        wsExport.name = Left$( _
            "Д1_" & _
            Add1_SanitizePathName(selectedUnit) & "_" & _
            Add1_SanitizePathName(mvoDisplay), _
            31)
'        On Error GoTo ErrHandler

        fileName = _
            "Д1_" & _
            Add1_SanitizePathName(selectedUnit) & "_" & _
            Add1_SanitizePathName(mvoDisplay) & _
            ".xlsx"

        filePath = _
            unitFolderPath & _
            Application.PathSeparator & _
            fileName

        If Len(Dir(filePath)) > 0 Then

            Application.DisplayAlerts = True

            answer = MsgBox( _
                "Файл уже існує:" & vbCrLf & vbCrLf & _
                fileName & vbCrLf & vbCrLf & _
                "Замінити його?", _
                vbYesNo + vbQuestion, _
                "Збереження Додатка 1")

            Application.DisplayAlerts = False

            If answer = vbNo Then

                skippedCount = skippedCount + 1

                exportWb.Close SaveChanges:=False
                Set wsExport = Nothing
                Set exportWb = Nothing

                GoTo NextMVO

            End If

            Kill filePath
        End If

        exportWb.SaveAs _
            fileName:=filePath, _
            FileFormat:=xlOpenXMLWorkbook

        exportWb.Close SaveChanges:=False

        'Номер вважаємо використаним тільки після успішного SaveAs.
        currentReportNumber = nextReportNumber
        Add1_CommitReportNumber currentReportNumber

        createdCount = createdCount + 1

        Set wsExport = Nothing
        Set exportWb = Nothing

NextMVO:
    Next mvoKey

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox _
        "Збереження завершено." & vbCrLf & vbCrLf & _
        "Підрозділ: " & selectedUnit & vbCrLf & _
        "МВО: " & dictGroups.count & vbCrLf & _
        "Створено файлів: " & createdCount & vbCrLf & _
        "Пропущено: " & skippedCount & vbCrLf & vbCrLf & _
        "Папка:" & vbCrLf & unitFolderPath, _
        vbInformation, _
        "Додаток 1"

    Exit Sub

ErrHandler:
    On Error Resume Next

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    If Not exportWb Is Nothing Then
        exportWb.Close SaveChanges:=False
    End If

    Err.Raise _
        Err.Number, _
        "Add1_SaveOneUnitByMVO", _
        Err.Description

End Sub


'===============================================================================
' 7. ПАКЕТНЕ ЗБЕРЕЖЕННЯ ОДНОГО ПІДРОЗДІЛУ У РЕЖИМІ "ВСІ ПІДРОЗДІЛИ"
'
' Зберігаємо стару структуру:
' Д1_експорт\<батальйон>\
'
' Але тепер всередині батальйону окремий файл на кожного МВО:
' Д1_<підрозділ>_<МВО>.xlsx
'
' Повертає кількість створених файлів.
'===============================================================================
Private Function Add1_SaveAllModeUnitByMVO( _
    ByVal wsResiduse As Worksheet, _
    ByVal wsTemplate As Worksheet, _
    ByVal selectedUnit As String, _
    ByVal startDate As Date, _
    ByVal endDate As Date, _
    ByVal dictGroups As Object, _
    ByVal dictMvoNames As Object, _
    ByVal workbookFolder As String) As Long

    Dim exportRootPath As String
    Dim battalionFolderPath As String

    Dim mvoKey As Variant
    Dim mvoDisplay As String
    Dim rowsForMVO As Collection

    Dim exportWb As Workbook
    Dim wsExport As Worksheet

    Dim fileName As String
    Dim filePath As String

    Dim createdCount As Long

    Dim currentReportNumber As String
    Dim nextReportNumber As String

    On Error GoTo ErrHandler

    'Беремо останній використаний номер із Додаток 1!E7.
    'Оскільки E7 оновлюється після кожного успішно створеного файла,
    'нумерація продовжується між усіма підрозділами і МВО.
    currentReportNumber = Add1_GetCurrentReportNumber()

    exportRootPath = _
        workbookFolder & _
        Application.PathSeparator & _
        EXPORT_ROOT_FOLDER_NAME

    If Dir(exportRootPath, vbDirectory) = vbNullString Then
        MkDir exportRootPath
    End If

    battalionFolderPath = _
        exportRootPath & _
        Application.PathSeparator & _
        Add1_GetBattalionFolderName(selectedUnit)

    If Dir(battalionFolderPath, vbDirectory) = vbNullString Then
        MkDir battalionFolderPath
    End If

    For Each mvoKey In dictGroups.Keys

        mvoDisplay = CStr(dictMvoNames(mvoKey))
        Set rowsForMVO = dictGroups(mvoKey)

        'Наступний номер відомості для цього файла.
        nextReportNumber = Add1_IncrementReportNumber(currentReportNumber)

        wsTemplate.Copy

        Set exportWb = ActiveWorkbook
        Set wsExport = exportWb.Worksheets(1)

        Add1_FillFromExactRows _
            wsResiduse, _
            wsExport, _
            selectedUnit, _
            mvoDisplay, _
            rowsForMVO, _
            startDate, _
            endDate, _
            nextReportNumber

        On Error Resume Next
        wsExport.name = Left$( _
            "Д1_" & _
            Add1_SanitizePathName(selectedUnit) & "_" & _
            Add1_SanitizePathName(mvoDisplay), _
            31)
        On Error GoTo ErrHandler

        fileName = _
            "Д1_" & _
            Add1_SanitizePathName(selectedUnit) & "_" & _
            Add1_SanitizePathName(mvoDisplay) & _
            ".xlsx"

        filePath = _
            battalionFolderPath & _
            Application.PathSeparator & _
            fileName

        If Len(Dir(filePath)) > 0 Then
            Kill filePath
        End If

        exportWb.SaveAs _
            fileName:=filePath, _
            FileFormat:=xlOpenXMLWorkbook

        exportWb.Close SaveChanges:=False

        'Фіксуємо номер тільки після успішного створення файла.
        currentReportNumber = nextReportNumber
        Add1_CommitReportNumber currentReportNumber

        createdCount = createdCount + 1

        Set wsExport = Nothing
        Set exportWb = Nothing

    Next mvoKey

    Add1_SaveAllModeUnitByMVO = createdCount
    Exit Function

ErrHandler:
    On Error Resume Next

    If Not exportWb Is Nothing Then
        exportWb.Close SaveChanges:=False
    End If

    Err.Raise _
        Err.Number, _
        "Add1_SaveAllModeUnitByMVO", _
        Err.Description

End Function


'===============================================================================
' 8. ЗАПОВНЕННЯ ОДНОГО ДОДАТКА 1 З КОНКРЕТНОЇ COLLECTION РЯДКІВ МВО
'
' Ця процедура ВЗАГАЛІ НЕ читає інші рядки довідки.
' Тому майно іншого МВО не може потрапити у поточний файл.
'===============================================================================
Private Sub Add1_FillFromExactRows( _
    ByVal wsResiduse As Worksheet, _
    ByVal wsTarget As Worksheet, _
    ByVal selectedUnit As String, _
    ByVal mvoDisplay As String, _
    ByVal rowsForMVO As Collection, _
    ByVal startDate As Date, _
    ByVal endDate As Date, _
    ByVal reportNumber As String)

    Dim tbl As ListObject
    Dim dictItems As Object

    Dim rowNumber As Variant

    Dim itemName As String
    Dim itemUnit As String
    Dim itemPrice As Variant

    Dim qIn As Double
    Dim qOut As Double
    Dim qBalance As Double
    Dim qStart As Double

    Dim key As String
    Dim itemKey As Variant
    Dim a As Variant
    Dim parts() As String

    Dim newRow As ListRow
    Dim totalRow As ListRow
    Dim dataRange As Range

    Dim counter As Long

    Dim totalNadiyshlo As Double
    Dim totalVybylo As Double
    Dim totalZalyshky As Double

    Set dictItems = CreateObject("Scripting.Dictionary")

    '---------------------------------------------------------------------------
    ' Знаходимо таблицю
    '---------------------------------------------------------------------------
    On Error Resume Next
    Set tbl = wsTarget.ListObjects("tbAddOne")
    On Error GoTo 0

    If tbl Is Nothing Then

        If wsTarget.ListObjects.count = 0 Then
            Err.Raise _
                vbObjectError + 901, _
                "Add1_FillFromExactRows", _
                "На аркуші Додаток 1 не знайдено таблицю."
        End If

        Set tbl = wsTarget.ListObjects(1)
    End If

    '---------------------------------------------------------------------------
    ' Шапка
    '---------------------------------------------------------------------------
    Add1_FillHeader _
        wsTarget, _
        selectedUnit, _
        startDate, _
        endDate, _
        reportNumber

    '---------------------------------------------------------------------------
    ' Читаємо ТІЛЬКИ передані рядки конкретного МВО
    '---------------------------------------------------------------------------
    For Each rowNumber In rowsForMVO

        itemName = CStr(wsResiduse.Cells(CLng(rowNumber), "E").value)
        itemPrice = wsResiduse.Cells(CLng(rowNumber), "F").value
        itemUnit = CStr(wsResiduse.Cells(CLng(rowNumber), "L").value)

        qIn = Add1_SafeDouble(wsResiduse.Cells(CLng(rowNumber), "G").value)
        qOut = Add1_SafeDouble(wsResiduse.Cells(CLng(rowNumber), "H").value)
        qBalance = Add1_SafeDouble(wsResiduse.Cells(CLng(rowNumber), "I").value)
        qStart = Add1_SafeDouble(wsResiduse.Cells(CLng(rowNumber), "N").value)

        'У межах одного МВО групування:
        'Найменування + Ціна + Од. виміру
        key = _
            itemName & KEY_SEP & _
            CStr(itemPrice) & KEY_SEP & _
            itemUnit

        If dictItems.exists(key) Then
            a = dictItems(key)
        Else
            ReDim a(1 To 4)
        End If

        a(1) = CDbl(a(1)) + qIn
        a(2) = CDbl(a(2)) + qOut
        a(3) = CDbl(a(3)) + qBalance
        a(4) = CDbl(a(4)) + qStart

        dictItems(key) = a

    Next rowNumber

    '---------------------------------------------------------------------------
    ' Повністю очищаємо старі рядки таблиці
    '---------------------------------------------------------------------------
    Do While tbl.ListRows.count > 0
        tbl.ListRows(tbl.ListRows.count).Delete
    Loop

    '---------------------------------------------------------------------------
    ' Заповнюємо тільки майном поточного МВО
    '---------------------------------------------------------------------------
    counter = 0

    For Each itemKey In dictItems.Keys

        a = dictItems(itemKey)
        parts = Split(CStr(itemKey), KEY_SEP)

        counter = counter + 1

        Set newRow = tbl.ListRows.Add

        newRow.Range(1, 1).value = counter
        newRow.Range(1, 2).value = parts(0)                   'Найменування
        newRow.Range(1, 3).value = parts(2)                   'Од. виміру
        newRow.Range(1, 4).value = CDbl(a(1))                 'Надійшло
        newRow.Range(1, 5).value = CDbl(a(2))                 'Вибуло
        newRow.Range(1, 6).value = CDbl(a(3))                 'Залишок
        newRow.Range(1, 7).value = Add1_SafeDouble(parts(1))  'Ціна

        'Повертаємо початковий залишок у тому вигляді,
        'як він був розрахований при формуванні "Довідка по підрозділу 2.0".
        If tbl.ListColumns.count >= 8 Then

            newRow.Range(1, 8).value = CDbl(a(4))

        End If

    Next itemKey

    '---------------------------------------------------------------------------
    ' Форматування + "Всього"
    '---------------------------------------------------------------------------
    If Not tbl.DataBodyRange Is Nothing Then

        Set dataRange = tbl.DataBodyRange

        dataRange.EntireRow.AutoFit

        With dataRange.Columns(2)
            .VerticalAlignment = xlCenter
            .HorizontalAlignment = xlLeft
        End With

        If dataRange.Columns.count >= 6 Then
            With dataRange.Columns(3).Resize(, 4)
                .VerticalAlignment = xlCenter
                .HorizontalAlignment = xlRight
            End With
        End If

        totalNadiyshlo = WorksheetFunction.Round( _
            Application.WorksheetFunction.Sum(dataRange.Columns(4)), 3)

        totalVybylo = WorksheetFunction.Round( _
            Application.WorksheetFunction.Sum(dataRange.Columns(5)), 3)

        totalZalyshky = WorksheetFunction.Round( _
            Application.WorksheetFunction.Sum(dataRange.Columns(6)), 3)

        Set totalRow = tbl.ListRows.Add

        totalRow.Range(1, 1).ClearContents
        totalRow.Range(1, 2).value = "Всього"
        totalRow.Range(1, 3).ClearContents
        totalRow.Range(1, 4).value = totalNadiyshlo
        totalRow.Range(1, 5).value = totalVybylo
        totalRow.Range(1, 6).value = totalZalyshky

        With totalRow.Range
            .Font.Bold = True
            .EntireRow.AutoFit
            .VerticalAlignment = xlCenter
        End With

        totalRow.Range(1, 2).HorizontalAlignment = xlLeft

        With totalRow.Range.Cells(1, 4).Resize(1, 3)
            .NumberFormat = "0.000"
            .HorizontalAlignment = xlRight
            .VerticalAlignment = xlCenter
        End With

    End If

    '---------------------------------------------------------------------------
    ' Верхній підписант = конкретний МВО.
    ' Нижній підписний блок лишається порожнім для ручного внесення.
    '---------------------------------------------------------------------------
    Add1_WriteMVOSignature _
        wsTarget, _
        tbl, _
        mvoDisplay

End Sub


'===============================================================================
' 9. ШАПКА ДОДАТКА 1
'
' Підрозділ — об'єднаний A9:F9, запис через A9.
'
' Якщо є аркуш "Підписи" / "Таблиця_Підпис":
' - беремо повну назву підрозділу;
' - беремо № узагальнюючої відомості.
'
' Нижні підписанти з аркуша "Підписи" НЕ підставляються.
'===============================================================================
Private Sub Add1_FillHeader( _
    ByVal wsTarget As Worksheet, _
    ByVal selectedUnit As String, _
    ByVal startDate As Date, _
    ByVal endDate As Date, _
    ByVal reportNumber As String)

    Dim wsSign As Worksheet
    Dim tblSign As ListObject
    Dim rngSign As Range

    Dim fullUnitName As String

    'Підрозділ — у A9:F9 через A9.
    wsTarget.Range("A9").value = selectedUnit

    'Номер узагальнюючої відомості.
    'В E7 завжди записуємо як текст, щоб значення на кшталт
    '15/15/113, А-15/009-Б тощо Excel не перетворював у дату.
    With wsTarget.Range("E7")
        .NumberFormat = "@"
        .value = reportNumber
    End With

    On Error Resume Next
    wsTarget.Range("AddOneNameTroop").value = selectedUnit
    wsTarget.Range("AddOnePeriodFrom").value = startDate
    wsTarget.Range("AddOnePeriodTo").value = endDate
    wsTarget.Range("AddOnePeriod").value = _
        Format$(startDate, "dd.mm.yyyy") & "-" & _
        Format$(endDate, "dd.mm.yyyy")
    On Error GoTo 0

    'З аркуша "Підписи" беремо тільки повну назву підрозділу.
    'Номер відомості звідти більше НЕ беремо:
    'єдине джерело нумерації — Додаток 1!E7.
    If Add1_TryGetSignatureRow( _
        selectedUnit, _
        wsSign, _
        tblSign, _
        rngSign) Then

        fullUnitName = CStr(wsSign.Cells(rngSign.Row, 3).value)

        If Len(Trim$(fullUnitName)) = 0 Then
            fullUnitName = selectedUnit
        End If

        wsTarget.Range("A9").value = fullUnitName

    End If

End Sub


'===============================================================================
' 10. ВЕРХНІЙ ПІДПИСНИЙ БЛОК = МВО
'
' Рядок верхнього підпису:
'   одразу після таблиці.
'
' Нижній підписний блок:
'   signRow + 3
'
' Його значення справа очищається, щоб начальника служби / діловода /
' іншу особу можна було вписати вручну.
'
' Форматування/межі/об'єднання клітинок не змінюємо.
'===============================================================================
Private Sub Add1_WriteMVOSignature( _
    ByVal ws As Worksheet, _
    ByVal tbl As ListObject, _
    ByVal mvoDisplay As String)

    Dim firstCaptionRow As Long
    Dim secondCaptionRow As Long
    Dim oldWrongRow As Long

    'Знаходимо два підписні блоки за текстом під ними
    Add1_FindSignatureRows _
        ws, _
        tbl, _
        firstCaptionRow, _
        secondCaptionRow

    If firstCaptionRow = 0 Then
        Err.Raise vbObjectError + 950, _
                  "Add1_WriteMVOSignature", _
                  "Не знайдено верхній блок підпису МВО."
    End If

    '=========================================================
    ' Прибираємо МВО зі старого неправильного місця,
    ' куди його записувала попередня версія макросу
    '=========================================================
    oldWrongRow = tbl.Range.Row + tbl.Range.Rows.count

    If oldWrongRow <> firstCaptionRow - 1 Then
        Add1_ClearSignatureCell ws.Cells(oldWrongRow, "E")
    End If

    '=========================================================
    ' ВЕРХНІЙ БЛОК
    '
    ' Пояснювальний напис знаходиться в firstCaptionRow.
    ' Отже сам підпис / ПІБ МВО має бути рядком вище.
    '
    ' Записуємо СПРАВА — у F.
    '=========================================================
    Add1_SetSignatureCell _
        ws.Cells(firstCaptionRow - 1, "F"), _
        mvoDisplay

    '=========================================================
    ' НИЖНІЙ БЛОК
    '
    ' Залишаємо порожнім для ручного внесення:
    ' начальник служби / діловод / інша особа
    '=========================================================
    If secondCaptionRow > 0 Then
        Add1_ClearSignatureCell _
            ws.Cells(secondCaptionRow - 1, "F")
    End If

End Sub
Private Sub Add1_FindSignatureRows( _
    ByVal ws As Worksheet, _
    ByVal tbl As ListObject, _
    ByRef firstCaptionRow As Long, _
    ByRef secondCaptionRow As Long)

    Dim r As Long
    Dim c As Long
    Dim startRow As Long
    Dim lastRow As Long
    Dim txt As String

    firstCaptionRow = 0
    secondCaptionRow = 0

    startRow = tbl.Range.Row + tbl.Range.Rows.count

    'Шукаємо з запасом нижче таблиці
    lastRow = startRow + 30

    For r = startRow To lastRow

        For c = 1 To 6   'A:F

            txt = CStr(ws.Cells(r, c).value)

            If InStr(1, txt, _
                "посада, військове звання, підпис", _
                vbTextCompare) > 0 Then

                If firstCaptionRow = 0 Then

                    firstCaptionRow = r

                ElseIf r <> firstCaptionRow Then

                    secondCaptionRow = r
                    Exit Sub

                End If

            End If

        Next c

    Next r

End Sub
Private Sub Add1_SetSignatureCell( _
    ByVal targetCell As Range, _
    ByVal textValue As String)

    If targetCell.MergeCells Then

        With targetCell.MergeArea
            .Cells(1, 1).value = textValue
            .HorizontalAlignment = xlRight
            .VerticalAlignment = xlBottom
        End With

    Else

        With targetCell
            .value = textValue
            .HorizontalAlignment = xlRight
            .VerticalAlignment = xlBottom
        End With

    End If

End Sub
Private Sub Add1_ClearSignatureCell(ByVal targetCell As Range)

    If targetCell.MergeCells Then
        targetCell.MergeArea.ClearContents
    Else
        targetCell.ClearContents
    End If

End Sub
Private Sub Add1_SetSignatureRightValue( _
    ByVal targetCell As Range, _
    ByVal textValue As String)

    If targetCell.MergeCells Then

        targetCell.MergeArea.Cells(1, 1).value = textValue
        targetCell.MergeArea.HorizontalAlignment = xlCenter
        targetCell.MergeArea.VerticalAlignment = xlCenter

    Else

        targetCell.value = textValue
        targetCell.HorizontalAlignment = xlCenter
        targetCell.VerticalAlignment = xlCenter

    End If

End Sub


Private Sub Add1_ClearSignatureRightValue(ByVal targetCell As Range)

    If targetCell.MergeCells Then
        targetCell.MergeArea.Cells(1, 1).ClearContents
    Else
        targetCell.ClearContents
    End If

End Sub


'===============================================================================
' 11. ОТРИМАТИ УНІКАЛЬНІ ПІДРОЗДІЛИ З ГОТОВОЇ ДОВІДКИ
'===============================================================================
Private Function Add1_GetUnitsFromResiduse(ByVal ws As Worksheet) As Object

    Dim dict As Object
    Dim lastRow As Long
    Dim r As Long
    Dim unitName As String

    Set dict = CreateObject("Scripting.Dictionary")

    lastRow = ws.Cells(ws.Rows.count, "C").End(xlUp).Row

    For r = 8 To lastRow

        unitName = Trim$(CStr(ws.Cells(r, "C").value))

        If Len(unitName) > 0 Then
            If Not dict.exists(unitName) Then
                dict.Add unitName, Nothing
            End If
        End If

    Next r

    Set Add1_GetUnitsFromResiduse = dict

End Function


'===============================================================================
' 12. ПОШУК ПІДРОЗДІЛУ НА АРКУШІ "ПІДПИСИ"
'
' Використовуємо тільки:
' - № відомості;
' - повну назву підрозділу.
'
' Підписантів звідси автоматично не вставляємо.
'===============================================================================
Private Function Add1_TryGetSignatureRow( _
    ByVal unitKey As String, _
    ByRef wsSign As Worksheet, _
    ByRef tblSign As ListObject, _
    ByRef rngSign As Range) As Boolean

    On Error GoTo NoSignature

    Set wsSign = ThisWorkbook.Worksheets("Підписи")
    Set tblSign = wsSign.ListObjects("Таблиця_Підпис")

    If tblSign.DataBodyRange Is Nothing Then GoTo NoSignature

    Set rngSign = tblSign.ListColumns("Підрозділ").DataBodyRange.Find( _
        What:=unitKey, _
        LookIn:=xlValues, _
        LookAt:=xlWhole)

    If rngSign Is Nothing Then

        Set rngSign = tblSign.ListColumns("Підрозділ скорочено").DataBodyRange.Find( _
            What:=unitKey, _
            LookIn:=xlValues, _
            LookAt:=xlWhole)

    End If

    If rngSign Is Nothing Then GoTo NoSignature

    Add1_TryGetSignatureRow = True
    Exit Function

NoSignature:
    Set rngSign = Nothing
    Set tblSign = Nothing
    Set wsSign = Nothing

    Add1_TryGetSignatureRow = False

End Function


'===============================================================================
' 13. ВИЗНАЧЕННЯ ПАПКИ БАТАЛЬЙОНУ
'
' Залишена стара логіка:
' якщо назва "(рота_1БпС)" -> папка "1БпС"
'===============================================================================
Private Function Add1_GetBattalionFolderName(ByVal unitKey As String) As String

    Dim normalized As String
    Dim pos As Long
    Dim folderName As String

    normalized = Trim$(CStr(unitKey))

    If Len(normalized) >= 2 Then

        If Left$(normalized, 1) = "(" And _
           Right$(normalized, 1) = ")" Then

            normalized = Mid$( _
                normalized, _
                2, _
                Len(normalized) - 2)

        End If

    End If

    pos = InStrRev(normalized, "_")

    If pos > 0 Then
        folderName = Mid$(normalized, pos + 1)
    Else
        folderName = normalized
    End If

    folderName = Add1_SanitizePathName(folderName)

    If Len(folderName) = 0 Then
        folderName = "Без_батальйону"
    End If

    Add1_GetBattalionFolderName = folderName

End Function


'===============================================================================
' 14. НОРМАЛІЗАЦІЯ ТЕКСТУ
'
' "ПРОКОПЧУК А."
' "ПРОКОПЧУК А. "
' "ПРОКОПЧУК А." (нерозривний пробіл)
'
' сприймаються як один МВО.
'===============================================================================
Private Function Add1_NormalizeText(ByVal v As Variant) As String

    Dim s As String

    s = CStr(v)

    s = Replace(s, Chr$(160), " ")
    s = Replace(s, vbTab, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")

    s = Application.WorksheetFunction.Trim(s)

    Add1_NormalizeText = s

End Function


'===============================================================================
' 15. БЕЗПЕЧНЕ ПЕРЕТВОРЕННЯ В ЧИСЛО
'===============================================================================
Private Function Add1_SafeDouble(ByVal v As Variant) As Double

    If IsError(v) Or IsEmpty(v) Then

        Add1_SafeDouble = 0

    ElseIf Len(Trim$(CStr(v))) = 0 Then

        Add1_SafeDouble = 0

    ElseIf IsNumeric(v) Then

        Add1_SafeDouble = CDbl(v)

    Else

        Add1_SafeDouble = 0

    End If

End Function


'===============================================================================
' 16. БЕЗПЕЧНА НАЗВА ПАПКИ / ФАЙЛА
'===============================================================================
Private Function Add1_SanitizePathName(ByVal rawName As String) As String

    Dim result As String

    result = Trim$(CStr(rawName))

    result = Replace(result, "\", "_")
    result = Replace(result, "/", "_")
    result = Replace(result, ":", "_")
    result = Replace(result, "*", "_")
    result = Replace(result, "?", "_")
    result = Replace(result, Chr$(34), "_")
    result = Replace(result, "<", "_")
    result = Replace(result, ">", "_")
    result = Replace(result, "|", "_")
    result = Replace(result, "[", "_")
    result = Replace(result, "]", "_")

    Do While Len(result) > 0 And _
             (Right$(result, 1) = " " Or _
              Right$(result, 1) = ".")

        result = Left$(result, Len(result) - 1)

    Loop

    If Len(result) = 0 Then
        result = "Без_назви"
    End If

    Add1_SanitizePathName = result

End Function

 
'===============================================================================
' НУМЕРАЦІЯ УЗАГАЛЬНЮЮЧИХ ВІДОМОСТЕЙ — ДОДАТОК 1!E7
'
' Правило:
' - 15          -> 16
' - 15/15/112   -> 15/15/113
' - А-15/009-Б  -> А-15/010-Б
' - №7-А        -> №8-А
'
' Збільшується НА 1 ПРАВИЙ (ОСТАННІЙ) ЧИСЛОВИЙ БЛОК у рядку.
' Усі букви, "/", "-", пробіли та інші символи зберігаються.
'
' Кожен окремо збережений файл Додатка 1 отримує свій наступний номер.
' Після успішного SaveAs новий номер записується назад у
' ThisWorkbook.Worksheets("Додаток 1").Range("E7").
'===============================================================================

Private Function Add1_GetCurrentReportNumber() As String

    Dim s As String

    s = Trim$(CStr(ThisWorkbook.Worksheets("Додаток 1").Range("E7").Value2))

    If Len(s) = 0 Then
        Err.Raise _
            vbObjectError + 1101, _
            "Add1_GetCurrentReportNumber", _
            "У клітинці Додаток 1!E7 не вказано номер узагальнюючої відомості."
    End If

    If Not Add1_TextContainsDigit(s) Then
        Err.Raise _
            vbObjectError + 1102, _
            "Add1_GetCurrentReportNumber", _
            "Номер у Додаток 1!E7 не містить жодної цифри: " & s
    End If

    Add1_GetCurrentReportNumber = s

End Function


Private Function Add1_IncrementReportNumber(ByVal sourceNumber As String) As String

    Dim i As Long
    Dim digitEnd As Long
    Dim digitStart As Long

    Dim prefixPart As String
    Dim digitPart As String
    Dim suffixPart As String

    sourceNumber = CStr(sourceNumber)

    'Шукаємо справа останню цифру.
    For i = Len(sourceNumber) To 1 Step -1

        If Mid$(sourceNumber, i, 1) Like "#" Then
            digitEnd = i
            Exit For
        End If

    Next i

    If digitEnd = 0 Then
        Err.Raise _
            vbObjectError + 1103, _
            "Add1_IncrementReportNumber", _
            "Неможливо збільшити номер, оскільки він не містить цифр: " & sourceNumber
    End If

    'Визначаємо початок цього останнього числового блока.
    digitStart = digitEnd

    Do While digitStart > 1

        If Mid$(sourceNumber, digitStart - 1, 1) Like "#" Then
            digitStart = digitStart - 1
        Else
            Exit Do
        End If

    Loop

    prefixPart = Left$(sourceNumber, digitStart - 1)
    digitPart = Mid$(sourceNumber, digitStart, digitEnd - digitStart + 1)
    suffixPart = Mid$(sourceNumber, digitEnd + 1)

    Add1_IncrementReportNumber = _
        prefixPart & _
        Add1_IncrementDigitString(digitPart) & _
        suffixPart

End Function


'Збільшення цифрового блока без обмеження Long.
'Також зберігаються початкові нулі:
'009 -> 010
'099 -> 100
'999 -> 1000
Private Function Add1_IncrementDigitString(ByVal digits As String) As String

    Dim i As Long
    Dim n As Long
    Dim carry As Long
    Dim result As String

    result = digits
    carry = 1

    For i = Len(result) To 1 Step -1

        n = (Asc(Mid$(result, i, 1)) - 48) + carry

        If n >= 10 Then
            Mid$(result, i, 1) = CStr(n - 10)
            carry = 1
        Else
            Mid$(result, i, 1) = CStr(n)
            carry = 0
            Exit For
        End If

    Next i

    If carry = 1 Then
        result = "1" & result
    End If

    Add1_IncrementDigitString = result

End Function


Private Function Add1_TextContainsDigit(ByVal s As String) As Boolean

    Dim i As Long

    For i = 1 To Len(s)

        If Mid$(s, i, 1) Like "#" Then
            Add1_TextContainsDigit = True
            Exit Function
        End If

    Next i

End Function


Private Sub Add1_CommitReportNumber(ByVal reportNumber As String)

    With ThisWorkbook.Worksheets("Додаток 1").Range("E7")
        .NumberFormat = "@"
        .value = reportNumber
    End With

End Sub


'===============================================================================
' 17. АВТОМАТИЧНА ПЕРЕВІРКА 3 КНОПОК НА АРКУШІ "ДОДАТОК 1"
'
' Має бути:
'   1) CreateAddOneForOneTroop2
'   2) CreateAddOneForAllTroops
'   3) CreateAddOneFromCurrentResiduse
'
' ВАЖЛИВО:
' - НЕ використовується Shape.PrintObject, бо саме ця властивість
'   викликала в поточній книзі помилку:
'   "Object doesn't support this property or method".
' - кнопки 2 і 3 створюються НОВИМИ Shape, а не Duplicate існуючої;
' - існуюча кнопка 1 зберігається;
' - якщо кнопка 1 не знайдена за OnAction, додатково шукаємо її за ім'ям
'   та текстом;
' - кнопки 2 і 3 завжди розставляються праворуч від кнопки 1.
'===============================================================================
Public Sub EnsureAddOneButtons(Optional ByVal showMessage As Boolean = False)

    Dim ws As Worksheet
    Dim btnOne As Shape
    Dim btnAll As Shape
    Dim btnCurrent As Shape

    Dim createdCount As Long
    Dim gap As Double

    On Error GoTo ErrHandler

    Set ws = ThisWorkbook.Worksheets("Додаток 1")
    gap = 18

    '---------------------------------------------------------------------------
    ' 1. Знайти першу існуючу кнопку
    '---------------------------------------------------------------------------
    Set btnOne = Add1_FindOneUnitButton(ws)

    If btnOne Is Nothing Then

        Set btnOne = Add1_CreateNewButton( _
            ws:=ws, _
            buttonName:="btnAddOneOneUnit", _
            captionText:="Створити Додаток 1" & vbCrLf & "за підрозділ", _
            macroName:="CreateAddOneForOneTroop2", _
            leftPos:=ws.Range("J2").Left, _
            topPos:=ws.Range("J2").Top, _
            buttonWidth:=180, _
            buttonHeight:=95)

        createdCount = createdCount + 1

    Else
        'Навіть для старої кнопки відновлюємо правильний макрос.
        On Error Resume Next
        btnOne.OnAction = "CreateAddOneForOneTroop2"
        On Error GoTo ErrHandler
    End If

    'Прибираємо зайві дублікати першої кнопки від попередніх невдалих запусків.
    Add1_DeleteDuplicateOneUnitButtons ws, btnOne

    '---------------------------------------------------------------------------
    ' 2. Прибираємо старі помилкові дублікати кнопок 2 і 3,
    '    якщо вони не є нашими нормальними Shape.
    '---------------------------------------------------------------------------
    Set btnAll = Add1_GetButtonByName(ws, "btnAddOneAllUnits")
    If Not btnAll Is Nothing Then
        If btnAll.Type = msoGroup Then
            btnAll.Delete
            Set btnAll = Nothing
        End If
    End If

    Set btnCurrent = Add1_GetButtonByName(ws, "btnAddOneCurrentResiduse")
    If Not btnCurrent Is Nothing Then
        If btnCurrent.Type = msoGroup Then
            btnCurrent.Delete
            Set btnCurrent = Nothing
        End If
    End If

    '---------------------------------------------------------------------------
    ' 3. Кнопка "за всі підрозділи"
    '---------------------------------------------------------------------------
    If btnAll Is Nothing Then

        Set btnAll = Add1_FindShapeByMacro(ws, "CreateAddOneForAllTroops")

        If Not btnAll Is Nothing Then
            If btnAll.Type = msoGroup Then
                btnAll.Delete
                Set btnAll = Nothing
            End If
        End If

    End If

    If btnAll Is Nothing Then

        Set btnAll = Add1_CreateNewButton( _
            ws:=ws, _
            buttonName:="btnAddOneAllUnits", _
            captionText:="Створити Додаток 1" & vbCrLf & "за всі підрозділи", _
            macroName:="CreateAddOneForAllTroops", _
            leftPos:=btnOne.Left + btnOne.Width + gap, _
            topPos:=btnOne.Top, _
            buttonWidth:=btnOne.Width, _
            buttonHeight:=btnOne.Height)

        createdCount = createdCount + 1

    Else

        Add1_ConfigureGeneratedButton _
            btnAll, _
            "Створити Додаток 1" & vbCrLf & "за всі підрозділи", _
            "CreateAddOneForAllTroops"

    End If

    '---------------------------------------------------------------------------
    ' 4. Кнопка "за вибраними підрозділами"
    '---------------------------------------------------------------------------
    If btnCurrent Is Nothing Then

        Set btnCurrent = Add1_FindShapeByMacro(ws, "CreateAddOneFromCurrentResiduse")

        If Not btnCurrent Is Nothing Then
            If btnCurrent.Type = msoGroup Then
                btnCurrent.Delete
                Set btnCurrent = Nothing
            End If
        End If

    End If

    If btnCurrent Is Nothing Then

        Set btnCurrent = Add1_CreateNewButton( _
            ws:=ws, _
            buttonName:="btnAddOneCurrentResiduse", _
            captionText:="Створити Додаток 1" & vbCrLf & "за вибраними підрозділами", _
            macroName:="CreateAddOneFromCurrentResiduse", _
            leftPos:=btnAll.Left + btnAll.Width + gap, _
            topPos:=btnOne.Top, _
            buttonWidth:=btnOne.Width, _
            buttonHeight:=btnOne.Height)

        createdCount = createdCount + 1

    Else

        Add1_ConfigureGeneratedButton _
            btnCurrent, _
            "Створити Додаток 1" & vbCrLf & "за вибраними підрозділами", _
            "CreateAddOneFromCurrentResiduse"

    End If

    '---------------------------------------------------------------------------
    ' 5. Примусово розставити кнопки без накладання
    '---------------------------------------------------------------------------
    btnAll.Left = btnOne.Left + btnOne.Width + gap
    btnAll.Top = btnOne.Top
    btnAll.Width = btnOne.Width
    btnAll.Height = btnOne.Height

    btnCurrent.Left = btnAll.Left + btnAll.Width + gap
    btnCurrent.Top = btnOne.Top
    btnCurrent.Width = btnOne.Width
    btnCurrent.Height = btnOne.Height

    '---------------------------------------------------------------------------
    ' 6. Контроль призначених макросів
    '---------------------------------------------------------------------------
    btnOne.OnAction = "CreateAddOneForOneTroop2"
    btnAll.OnAction = "CreateAddOneForAllTroops"
    btnCurrent.OnAction = "CreateAddOneFromCurrentResiduse"

    If showMessage Then

        MsgBox _
            "Перевірку кнопок завершено." & vbCrLf & vbCrLf & _
            "Створено відсутніх: " & createdCount & vbCrLf & _
            "На аркуші мають бути 3 робочі кнопки.", _
            vbInformation, _
            "Кнопки Додатка 1"

    End If

    Exit Sub

ErrHandler:

    If showMessage Then
        MsgBox _
            "Не вдалося перевірити/створити кнопки." & vbCrLf & _
            "Помилка № " & Err.Number & ": " & Err.Description, _
            vbExclamation, _
            "Кнопки Додатка 1"
    End If

End Sub


'===============================================================================
' 18. РУЧНИЙ ЗАПУСК ПЕРЕВІРКИ КНОПОК
'===============================================================================
Public Sub InstallAddOneButtons()

    EnsureAddOneButtons True

End Sub


'===============================================================================
' 19. ПОШУК ПЕРШОЇ КНОПКИ
'
' Порядок:
' 1) за OnAction;
' 2) за ім'ям;
' 3) за текстом Shape.
'===============================================================================
Private Function Add1_FindOneUnitButton(ByVal ws As Worksheet) As Shape

    Dim shp As Shape
    Dim candidate As Shape
    Dim captionText As String

    '1. За макросом: якщо після попередніх спроб є дублікати,
    '   залишаємо найлівішу кнопку як основну.
    For Each shp In ws.Shapes

        If Add1_ShapeRunsMacro(shp, "CreateAddOneForOneTroop2") Then

            If candidate Is Nothing Then
                Set candidate = shp
            ElseIf shp.Left < candidate.Left Then
                Set candidate = shp
            End If

        End If

    Next shp

    If Not candidate Is Nothing Then
        Set Add1_FindOneUnitButton = candidate
        Exit Function
    End If

    '2. За ім'ям
    Set shp = Add1_GetButtonByName(ws, "btnAddOneOneUnit")
    If Not shp Is Nothing Then
        Set Add1_FindOneUnitButton = shp
        Exit Function
    End If

    '3. За текстом: теж беремо найлівішу відповідну кнопку.
    Set candidate = Nothing

    For Each shp In ws.Shapes

        captionText = Add1_GetShapeCaption(shp)

        If Len(captionText) > 0 Then

            If InStr(1, captionText, "Створити", vbTextCompare) > 0 _
            And InStr(1, captionText, "Додаток 1", vbTextCompare) > 0 _
            And InStr(1, captionText, "підрозділ", vbTextCompare) > 0 _
            And InStr(1, captionText, "всі", vbTextCompare) = 0 _
            And InStr(1, captionText, "вибран", vbTextCompare) = 0 Then

                If candidate Is Nothing Then
                    Set candidate = shp
                ElseIf shp.Left < candidate.Left Then
                    Set candidate = shp
                End If

            End If

        End If

    Next shp

    Set Add1_FindOneUnitButton = candidate

End Function


Private Function Add1_ShapeRunsMacro( _
    ByVal shp As Shape, _
    ByVal macroName As String) As Boolean

    Dim actionText As String

    On Error Resume Next
    actionText = CStr(shp.OnAction)
    On Error GoTo 0

    If Len(actionText) > 0 Then
        Add1_ShapeRunsMacro = _
            (InStr(1, actionText, macroName, vbTextCompare) > 0)
    End If

End Function


'===============================================================================
' 20. ВИДАЛИТИ ЗАЙВІ ДУБЛІКАТИ ПЕРШОЇ КНОПКИ
'===============================================================================
Private Sub Add1_DeleteDuplicateOneUnitButtons( _
    ByVal ws As Worksheet, _
    ByVal keepButton As Shape)

    Dim i As Long

    For i = ws.Shapes.count To 1 Step -1

        If ws.Shapes(i).name <> keepButton.name Then

            If Add1_ShapeRunsMacro( _
                ws.Shapes(i), _
                "CreateAddOneForOneTroop2") Then

                ws.Shapes(i).Delete

            End If

        End If

    Next i

End Sub


'===============================================================================
' 21. ЗНАЙТИ SHAPE ЗА ON ACTION
'===============================================================================
Private Function Add1_FindShapeByMacro( _
    ByVal ws As Worksheet, _
    ByVal macroName As String) As Shape

    Dim shp As Shape
    Dim actionText As String

    For Each shp In ws.Shapes

        actionText = vbNullString

        On Error Resume Next
        actionText = CStr(shp.OnAction)
        On Error GoTo 0

        If Len(actionText) > 0 Then

            If InStr(1, actionText, macroName, vbTextCompare) > 0 Then
                Set Add1_FindShapeByMacro = shp
                Exit Function
            End If

        End If

    Next shp

End Function


'===============================================================================
' 21. ОТРИМАТИ SHAPE ЗА ІМ'ЯМ
'===============================================================================
Private Function Add1_GetButtonByName( _
    ByVal ws As Worksheet, _
    ByVal buttonName As String) As Shape

    On Error Resume Next
    Set Add1_GetButtonByName = ws.Shapes(buttonName)
    On Error GoTo 0

End Function


'===============================================================================
' 22. ОТРИМАТИ ТЕКСТ SHAPE
'===============================================================================
Private Function Add1_GetShapeCaption(ByVal shp As Shape) As String

    Dim s As String

    On Error Resume Next

    s = shp.TextFrame2.TextRange.text

    If Len(s) = 0 Then
        s = shp.TextFrame.Characters.text
    End If

    On Error GoTo 0

    Add1_GetShapeCaption = s

End Function


'===============================================================================
' 23. СТВОРИТИ НОВУ КНОПКУ
'
' ВАЖЛИВО:
' Shape.PrintObject тут НЕ використовується.
'===============================================================================
Private Function Add1_CreateNewButton( _
    ByVal ws As Worksheet, _
    ByVal buttonName As String, _
    ByVal captionText As String, _
    ByVal macroName As String, _
    ByVal leftPos As Double, _
    ByVal topPos As Double, _
    ByVal buttonWidth As Double, _
    ByVal buttonHeight As Double) As Shape

    Dim shp As Shape

    Set shp = ws.Shapes.AddShape( _
        msoShapeRoundedRectangle, _
        leftPos, _
        topPos, _
        buttonWidth, _
        buttonHeight)

    On Error Resume Next
    shp.name = buttonName
    On Error GoTo 0

    Add1_StyleGeneratedButton shp
    Add1_SetGeneratedButtonCaption shp, captionText

    shp.OnAction = macroName
    shp.Placement = xlFreeFloating

    Set Add1_CreateNewButton = shp

End Function


'===============================================================================
' 24. НАЛАШТУВАТИ ВЖЕ ІСНУЮЧУ КНОПКУ 2 АБО 3
'===============================================================================
Private Sub Add1_ConfigureGeneratedButton( _
    ByVal shp As Shape, _
    ByVal captionText As String, _
    ByVal macroName As String)

    Add1_StyleGeneratedButton shp
    Add1_SetGeneratedButtonCaption shp, captionText

    shp.OnAction = macroName
    shp.Placement = xlFreeFloating

End Sub


'===============================================================================
' 25. СТИЛЬ НОВИХ КНОПОК
'===============================================================================
Private Sub Add1_StyleGeneratedButton(ByVal shp As Shape)

    On Error Resume Next

    With shp

        .Fill.Visible = msoTrue
        .Fill.Solid
        .Fill.ForeColor.RGB = RGB(0, 0, 0)

        .Line.Visible = msoFalse

        'Shadow не є критичною для роботи.
        .Shadow.Visible = msoTrue

        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.MarginLeft = 8
        .TextFrame2.MarginRight = 8
        .TextFrame2.MarginTop = 6
        .TextFrame2.MarginBottom = 6

    End With

    On Error GoTo 0

End Sub


'===============================================================================
' 26. ТЕКСТ НОВОЇ КНОПКИ
'===============================================================================
Private Sub Add1_SetGeneratedButtonCaption( _
    ByVal shp As Shape, _
    ByVal captionText As String)

    On Error Resume Next

    With shp.TextFrame2.TextRange

        .text = captionText
        .ParagraphFormat.Alignment = msoAlignCenter

        With .Font
            .Bold = msoTrue
            .Size = 16
            .Fill.Visible = msoTrue
            .Fill.ForeColor.RGB = RGB(255, 192, 0)
        End With

    End With

    'Резервний варіант, якщо TextFrame2 недоступний.
    If Err.Number <> 0 Then

        Err.Clear
        shp.TextFrame.Characters.text = captionText
        shp.TextFrame.HorizontalAlignment = xlHAlignCenter
        shp.TextFrame.VerticalAlignment = xlVAlignCenter

    End If

    On Error GoTo 0

End Sub




```

---

## 3. Add14.bas — Відомість № 14 (по всій частині)
### 3.1. Призначення
Зводить залишки майна у розрізі субрахунків бухгалтерського обліку по всіх підрозділах військової частини. Використовує захищені функції конвертації `SafeStr`, `ToNum`, `TryGetDate`.

### 3.2. Вихідний код Add14.bas
```vba
Attribute VB_Name = "Add14"
Option Explicit

' ============================================================
'  Додаток 13 (Книга обліку д14) - виправлена версія
'  Головна зміна: ПЕРЕД заповненням перевіряються всі дані бази
'  (дата, ціна, кількість, тип операції) і показується точний
'  рядок аркуша "База", де формат неправильний.
'  Потрібні в книзі: функція CleanFileName, форма ProgressBar.
'  Функція FindValue більше не потрібна.
' ============================================================

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

        Select Case typeOperation

            Case "надходження"
                quantity = ToNum(filteredData(n, 6))

                If Not nameExist Then
                    If Not DocExist Then
                        wsAdd13.Cells(lastRow13, LastColumn + 1).value = quantity
                        wsAdd13.Cells(lastRow13, LastColumn + 3).value = quantity
                    Else
                        wsAdd13.Cells(DocRow, LastColumn + 1).value = quantity
                        wsAdd13.Cells(DocRow, LastColumn + 3).value = quantity
                    End If
                Else
                    previousQuantyty = ToNum(wsAdd13.Cells(lastRow13 - 1, foundCell.Column + 2).value)
                    actualQuantyty = previousQuantyty + quantity

                    If Not DocExist Then
                        wsAdd13.Cells(lastRow13, foundCell.Column + 2).value = actualQuantyty
                        wsAdd13.Cells(lastRow13, foundCell.Column).value = quantity
                    Else
                        wsAdd13.Cells(DocRow, foundCell.Column + 2).value = actualQuantyty
                        ' додаємо попередню кількість (однаковий документ, але різні МВО)
                        wsAdd13.Cells(DocRow, foundCell.Column).value = quantity + ToNum(wsAdd13.Cells(DocRow, foundCell.Column).value)
                    End If
                End If

            Case "вибуття"
                quantity = ToNum(filteredData(n, 7))

                If Not nameExist Then
                    If Not DocExist Then
                        wsAdd13.Cells(lastRow13, LastColumn + 2).value = quantity
                        wsAdd13.Cells(lastRow13, LastColumn + 3).value = -quantity
                    Else
                        wsAdd13.Cells(DocRow, LastColumn + 2).value = quantity
                        wsAdd13.Cells(DocRow, LastColumn + 3).value = -quantity
                    End If
                Else
                    previousQuantyty = ToNum(wsAdd13.Cells(lastRow13 - 1, foundCell.Column + 2).value)
                    actualQuantyty = previousQuantyty - quantity

                    If Not DocExist Then
                        wsAdd13.Cells(lastRow13, foundCell.Column + 2).value = actualQuantyty
                        wsAdd13.Cells(lastRow13, foundCell.Column + 1).value = quantity
                    Else
                        wsAdd13.Cells(DocRow, foundCell.Column + 2).value = actualQuantyty
                        wsAdd13.Cells(DocRow, foundCell.Column + 1).value = quantity + ToNum(wsAdd13.Cells(DocRow, foundCell.Column + 1).value)
                    End If
                End If

        End Select

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

    MsgBox "Додаток сформовано. Оброблено записів: " & rowCount, vbInformation, "Готово"

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


```

---

## 4. Add14OneTroop.bas — Відомість № 14 для одного підрозділу
### 4.1. Призначення
Аналог Відомості № 14, але сформований індивідуально для обраного підрозділу (для звірки з фінансово-економічною службою).

### 4.2. Вихідний код Add14OneTroop.bas
```vba
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
    If Not dictUnits.exists(celll.value) And Not IsEmpty(celll.value) Then
        dictUnits.Add celll.value, Nothing
    End If
Next celll

For Each celll In colTo
    If Not dictUnits.exists(celll.value) And Not IsEmpty(celll.value) Then
        dictUnits.Add celll.value, Nothing
    End If
Next celll


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

' Зчитати назад у масив
Dim lastRowT As Long
lastRowT = wsTroops.Cells(wsTroops.Rows.count, 1).End(xlUp).Row


frm.Caption = "Виберіть підрозділ"
frm.ListData = wsTroops.Range("A1:A" & lastRowT)

 ' Показати форму
frm.show

' Отримати вибране
If frm.ListBox1.ListIndex <> -1 Then
    selectedUnit = frm.ListBox1.value
Else
    MsgBox "Підрозділ не вибрано!", vbExclamation
    Exit Sub
End If
   
 '''''''''''''''''''''''''''''''''''''''''''''''

    
    'Видаляємо попередні дані
    'видаляємо колонки
Dim lastColumnStart As Long
 lastColumnStart = wsAdd13.Cells(22, wsAdd13.Columns.count).End(xlToLeft).Column + 1
 
 'видаляємо рядки
Dim LastRowStart As Long
 LastRowStart = wsAdd13.Cells.Find(What:="*", After:=wsAdd13.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
                            SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
On Error Resume Next
    wsAdd13.Range(Cells(1, 9), Cells(LastRowStart, lastColumnStart)).EntireColumn.Delete ' Видалити колонки від U до останньої

    wsAdd13.Rows("24:" & wsAdd13.Rows.count).Delete

On Error GoTo 0

Dim RngToSearchData As Range
Set RngToSearchData = wsAdd13.Range("12:12") 'рядок з комбінацією даних в якому робити пошук на предмет співпадіння
    
Dim filteredData() As Variant 'масив з відфільтрованими даними

Dim criteriaRange14OneTroop As Range
Set criteriaRange14OneTroop = wsAdd13.Range("criteriaRange14OneTroop") 'критерії фільтрування, (надійшло, вибуло)будемо міняти на конкретний підрозділ

criteriaRange14OneTroop.Cells(2, 1) = selectedUnit
criteriaRange14OneTroop.Cells(3, 2) = selectedUnit


'знімаємо фільтри,якщо були поставлені
    If wsAllData.FilterMode Then
        wsAllData.ShowAllData
    End If
            
'фільтруємо по заданих критеріях
tbAllData.Range.AdvancedFilter Action:=xlFilterInPlace, criteriaRange:=criteriaRange14OneTroop, Unique:=False


Dim rowCount As Long, colCount As Long
Dim visibleRange As Range
Dim cell As Range
Dim area As Range

 ' Визначаємо кількість видимих рядків і стовпців
    rowCount = 0
    colCount = tbAllData.ListColumns.count
    On Error Resume Next
    Set visibleRange = tbAllData.DataBodyRange.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    
    If Not visibleRange Is Nothing Then
        For Each area In visibleRange.Areas
            rowCount = rowCount + area.Rows.count
        Next area
        
        ' Ініціалізуємо масив з урахуванням заголовків
        ReDim filteredData(1 To rowCount + 1, 1 To colCount)
        
        ' Записуємо заголовки таблиці у масив
        For k = 1 To colCount
            filteredData(1, k) = tbAllData.HeaderRowRange.Cells(1, k).value
        Next k
        
        ' Записуємо відфільтровані дані у масив
        i = 2
        For Each area In visibleRange.Areas
            For Each cell In area.Rows
                For j = 1 To colCount
                    filteredData(i, j) = cell.Cells(1, j).value
                Next j
                i = i + 1
            Next cell
        Next area
            
  End If
'''''''''''''''''''''''''''''''''''''''''''''''''''

Dim lastRow13 As Long
Dim CopyRange1 As Range
Dim foundCell As Range
' заголовок з назвою майна і категоріями
   Set CopyRange1 = wsAdd13.Range(wsAdd13.Cells(13, 6), wsAdd13.Cells(22, 8))
     'визначаємо, звідки копіювати шаблон для нового майна
    
'витягуємо дані з кожного рядка відфільтрованої бази даних
Dim n As Long

'On Error GoTo DontSelect
 For n = LBound(filteredData) + 1 To UBound(filteredData)
 
    Dim Document As String
    Dim DateDocument As Date
    Dim NumberDocument As String
    Dim FromTo As String
    Dim FromUnit As String
    Dim toUnit As String
    
    FromUnit = filteredData(n, 8) 'підрозділ звідки
    toUnit = filteredData(n, 9) 'підрозділ куди
    
    
    Document = filteredData(n, 25)
    DateDocument = filteredData(n, 13)
    NumberDocument = filteredData(n, 14)
    FromTo = FromUnit & Chr(10) & "(" & toUnit & ")"
    
    Dim typeOperation As String
    
    Dim name As String
    Dim Nomenclature As String
    Dim unit As String
    Dim unit2 As String
    Dim price As Double
    Dim DensityCount As Double
    
    Dim previousQuantyty As Double
    Dim previousQuantytyL As Double
    
    Dim actualQuantyty As Double
    Dim actualQuantytyL As Double
    
        typeOperation = filteredData(n, 26) 'тип операції
        
        name = filteredData(n, 2) 'найменування
'        Nomenclature = filteredData(n, 18) 'номенклатура
        unit = filteredData(n, 4) ' одиниця виміру кг
'        unit2 = filteredData(n, 23) ' одиниця виміру л
        Nomenclature = filteredData(n, 18)
        price = filteredData(n, 20) 'ціна
'        DensityCount = filteredData(n, 16) 'густина
        
        
      Dim generalInfo As String
      
      generalInfo = name & Nomenclature & unit & price 'сукупність інформації, за якою будуть формуватися унікальні колонки з найменуваннями
    
    Dim nameExist As Boolean
    nameExist = FindValue(generalInfo, RngToSearchData) 'шукаємо за сукупністю інформації майно
    
    Set foundCell = RngToSearchData.Find(generalInfo, LookIn:=xlValues, LookAt:=xlWhole)

    ''''''''''''''''''''''''''''''''''''''''''''''''''
    lastRow13 = Cells(wsAdd13.Rows.count, "A").End(xlUp).Row + 1 'зназходимо останній рядок і додаємо 1
    Dim LastColumn As Long
                LastColumn = wsAdd13.Cells.Find(What:="*", After:=wsAdd13.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    
    
    Dim lastNumber As Long
                                
        lastNumber = wsAdd13.Cells(22, LastColumn)
    
    'перевірка чи є вже такий документ
    Dim t As Long, DocExist As Boolean
    Dim DocRow As Long

    ' Порівняння кожного рядка масиву з відповідним рядком на листі
    For t = 24 To lastRow13 '24рядок, де починаються документи
        DocExist = True
            'перевірка, чи існує вже первинний документ
            If wsAdd13.Cells(t, 2).value <> Document Or wsAdd13.Cells(t, 3).value <> NumberDocument Or _
            wsAdd13.Cells(t, 4).value <> DateDocument Or wsAdd13.Cells(t, 5).value <> FromTo Then
                DocExist = False
            Else
            'якщо знайшли документ - забираємо рядок в якому він є і виходимо з циклу
                DocRow = wsAdd13.Cells(t, 1).Row
                Exit For
            End If
      
    Next t
    
   If Not DocExist Then
    
        '      Додаємо інформацію про документ
                With wsAdd13
                            .Cells(lastRow13, 1).value = filteredData(n, 1) 'дата внесення
                            .Cells(lastRow13, 2).value = Document 'найменування документа
                            .Cells(lastRow13, 3).NumberFormat = "@"
                            .Cells(lastRow13, 3).value = NumberDocument 'номер документа
                            On Error GoTo WrongDate
                            .Cells(lastRow13, 4).value = DateDocument  'дата документа
                            On Error GoTo 0
                            .Cells(lastRow13, 5).value = FromTo 'підрозділ звідки (підрозділ куди)
                End With
            
        'і ще дублювати всі колонки де є становить...
             Dim cellTotal As Range
             
        
            ' Перевірка кожної клітинки в третьому рядку
            For Each cellTotal In wsAdd13.Rows(21).Cells
                If InStr(cellTotal.value, "становить") > 0 Then
                    ' Копіювання інформації з попереднього рядка до поточного рядка у відповідну клітинку
                    wsAdd13.Cells(lastRow13, cellTotal.Column).value = wsAdd13.Cells(lastRow13 - 1, cellTotal.Column).value
'                    wsAdd13.Cells(lastRow13, cellTotal.Column + 1).value = wsAdd13.Cells(lastRow13 - 1, cellTotal.Column + 1).value
                    
                End If
            Next cellTotal
         
            Else
            
            
           If Not nameExist Then

                        wsAdd13.Activate
                    
                            CopyRange1.Copy
                            
                             wsAdd13.Cells(13, LastColumn + 1).PasteSpecial xlPasteAll
                             wsAdd13.Cells(12, LastColumn + 1) = generalInfo
                             wsAdd13.Cells(14, LastColumn + 1) = name
                             wsAdd13.Cells(16, LastColumn + 1) = Nomenclature
                             wsAdd13.Cells(18, LastColumn + 1) = unit
                             wsAdd13.Cells(20, LastColumn + 1) = price
                                    
                                    
                                'номерація колонок
                                wsAdd13.Cells(22, LastColumn + 1) = lastNumber + 1
                                wsAdd13.Cells(22, LastColumn + 2) = lastNumber + 2
                                wsAdd13.Cells(22, LastColumn + 3) = lastNumber + 3
                   
            End If
        
   
    End If
'    якщо є, то не додавати рядок, а знайти де знаходиться накладна, потім знайти де знаходиться майно і вставити туди кількість, якщо нема - додати колонки з майном


            
            
            Dim quantity As Double
'            Dim quantityL As Double
            If Not nameExist Then
'                        MsgBox "Працюємо"
                        wsAdd13.Activate
            
                                             
                            CopyRange1.Copy
                            
                             wsAdd13.Cells(13, LastColumn + 1).PasteSpecial xlPasteAll
                             wsAdd13.Cells(12, LastColumn + 1) = generalInfo
                             wsAdd13.Cells(14, LastColumn + 1) = name
                             wsAdd13.Cells(16, LastColumn + 1) = Nomenclature
                             wsAdd13.Cells(18, LastColumn + 1) = unit
                             wsAdd13.Cells(20, LastColumn + 1) = price
                                    
                                    
                                'номерація колонок
                                wsAdd13.Cells(22, LastColumn + 1) = lastNumber + 1
                                wsAdd13.Cells(22, LastColumn + 2) = lastNumber + 2
                                wsAdd13.Cells(22, LastColumn + 3) = lastNumber + 3
                   
            End If
            
            'визначаємо кількість, додаємо надходження і вибуття, розраховуючи, що одночасно в одному рядку одне з них 0.
            quantity = filteredData(n, 6) + filteredData(n, 7)
'            quantityL = filteredData(n, 3) + filteredData(n, 19)
            
           
            
            
         'якщо вибуває з підрозділу
         If FromUnit = selectedUnit Then
                               quantity = filteredData(n, 7)
'                        quantityL = filteredData(n, 19)
 
                    If Not nameExist Then
                        If Not DocExist Then
                            wsAdd13.Cells(lastRow13, LastColumn + 2).value = quantity
'                            wsAdd13.Cells(lastRow13, LastColumn + 4).value = quantityL
                            
                            wsAdd13.Cells(lastRow13, LastColumn + 3).value = -quantity
'                            wsAdd13.Cells(lastRow13, LastColumn + 6).value = -quantityL
                        Else
                            wsAdd13.Cells(DocRow, LastColumn + 2).value = quantity + wsAdd13.Cells(DocRow, LastColumn + 2).value
'                            wsAdd13.Cells(DocRow, LastColumn + 4).value = quantityL
                            wsAdd13.Cells(DocRow, LastColumn + 3).value = -quantity
'                            wsAdd13.Cells(DocRow, LastColumn + 6).value = -quantityL
                        End If
                    Else
                    
                        previousQuantyty = wsAdd13.Cells(lastRow13 - 1, foundCell.Column + 2).value
'                        previousQuantytyL = wsAdd13.Cells(lastRow13 - 1, foundCell.Column + 5).value
                        actualQuantyty = previousQuantyty - quantity 'віднімаємо від попередньої кількості
'                        actualQuantytyL = previousQuantytyL - quantityL
                        If Not DocExist Then
                            'вставляємо в поточний рядок кількість
                            wsAdd13.Cells(lastRow13, foundCell.Column + 2).value = actualQuantyty
'                            wsAdd13.Cells(lastRow13, foundCell.Column + 5).value = actualQuantytyL
                            wsAdd13.Cells(lastRow13, foundCell.Column + 1).value = quantity
'                            wsAdd13.Cells(lastRow13, foundCell.Column + 3).value = quantityL
                        Else
                            wsAdd13.Cells(DocRow, foundCell.Column + 2).value = actualQuantyty
'                            wsAdd13.Cells(DocRow, foundCell.Column + 5).value = actualQuantytyL
'                            ============ 3.9.15 ===========
                            'додаємо попередню кількість з клітинки (у випадку якщо однаковий документ, однакове найменування, однаковий підрозділ, але різні МВО)
                            wsAdd13.Cells(DocRow, foundCell.Column + 1).value = quantity + wsAdd13.Cells(DocRow, foundCell.Column + 1).value
'                            ============ 3.9.15 ===========
'                            wsAdd13.Cells(DocRow, foundCell.Column + 3).value = quantityL
                        End If
                       
                    End If
         End If
         
         
         'якщо надходить в підрозділ
         If toUnit = selectedUnit Then
                         If Not nameExist Then
                               If Not DocExist Then
                                     'вставляємо в поточний рядок
                                        wsAdd13.Cells(lastRow13, LastColumn + 1).value = quantity '+ wsAdd13.Cells(lastRow13, LastColumn + 1).value 'надійшло
'                                        wsAdd13.Cells(lastRow13, LastColumn + 2).value = quantityL
                                        wsAdd13.Cells(lastRow13, LastColumn + 3).value = quantity 'становить
'                                        wsAdd13.Cells(lastRow13, LastColumn + 6).value = quantityL
                               Else
                               'якщо нема назви і є документ...
                                        wsAdd13.Cells(DocRow, LastColumn + 1).value = quantity ' надійшло
'                                        wsAdd13.Cells(DocRow, LastColumn + 2).value = quantityL ' надійшло
                                        wsAdd13.Cells(DocRow, LastColumn + 3).value = quantity 'становить
'                                        wsAdd13.Cells(DocRow, LastColumn + 6).value = quantityL
                               End If
                    
                         
                     Else
                        previousQuantyty = wsAdd13.Cells(lastRow13 - 1, foundCell.Column + 2).value
'                        previousQuantytyL = wsAdd13.Cells(lastRow13 - 1, foundCell.Column + 5).value
                        actualQuantyty = previousQuantyty + quantity
'                        actualQuantytyL = previousQuantytyL + quantityL
                           If Not DocExist Then
                              
                                     wsAdd13.Cells(lastRow13, foundCell.Column + 2).value = actualQuantyty
'                                     wsAdd13.Cells(lastRow13, foundCell.Column + 5).value = actualQuantytyL
                                     
                                     wsAdd13.Cells(lastRow13, foundCell.Column).value = quantity
'                                     wsAdd13.Cells(lastRow13, foundCell.Column + 1).value = quantityL
                                 Else
                                      'якщо є назва і є документ...
                                      
                                    wsAdd13.Cells(DocRow, foundCell.Column + 2).value = actualQuantyty
'                                    wsAdd13.Cells(DocRow, foundCell.Column + 5).value = actualQuantytyL
'                                    ============ 3.9.15 ===========
                                    'додаємо попередню кількість з клітинки (у випадку якщо однаковий документ, однакове найменування, однаковий підрозділ, але різні МВО)
                                    wsAdd13.Cells(DocRow, foundCell.Column).value = quantity + wsAdd13.Cells(DocRow, foundCell.Column).value
'                                    ============ 3.9.15 ===========
'                                    wsAdd13.Cells(DocRow, foundCell.Column + 1).value = quantityL
                                                 
                            End If
                     End If
               
         End If
            
        
         'копіюємо стилі загальної іформації
        wsAdd13.Range(Cells(23, 1), Cells(23, 11)).Copy
        wsAdd13.Range(Cells(lastRow13, 1), Cells(lastRow13, 11)).PasteSpecial Paste:=xlPasteFormats
  
        Set foundCell = Nothing
        nameExist = False
        DocExist = False
    Call ProgressBar.Progress(n / UBound(filteredData) * 100)
    
    
    Next n
    
    'знаходимо останню колонку після всіх маніпуляцій
     Dim LastColumnFinal As Long
     LastColumnFinal = wsAdd13.Cells.Find(What:="*", After:=wsAdd13.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
     'копіюємо стилі основної інформації (для показників підрозділів)
       wsAdd13.Range(Cells(23, 6), Cells(23, 8)).Copy
       wsAdd13.Range(Cells(23, 9), Cells(lastRow13, LastColumnFinal)).PasteSpecial Paste:=xlPasteFormats


'скидаємо фільтри
On Error Resume Next
wsAllData.ShowAllData
Exit Sub
'*помилки*******************************************************
WrongDate:
MsgBox "Ймовірно, не правильний формат дати в " & Document & " № " & NumberDocument & " виправте дату і спробуйтие сформувати додаток ще раз!", vbCritical, "Не правильна дата"
wsAllData.ShowAllData 'скидаємо фільтри
Exit Sub

End Sub


```

---

## 5. Add47.bas та add47PNP.bas — Додаток 47
### 5.1. Призначення
Формують подетальну розгортку передачі окремо взятого виду майна між усіма підрозділами із зазначенням номерів документів та дат.

### 5.2. Вихідний код Add47.bas
```vba
Attribute VB_Name = "Add47"
Sub CreateAdd47()

Dim wbAllData As Workbook
Dim wsAllData As Worksheet
Dim tbAllData As ListObject
    Set wbAllData = ThisWorkbook
    Set wsAllData = wbAllData.Worksheets("База")
    Set tbAllData = wsAllData.ListObjects("AllData")
    
Dim wsAdd47 As Worksheet
    Set wsAdd47 = wbAllData.Worksheets("Додаток 47")
    
    
Dim RngToSearchData As Range
Set RngToSearchData = wsAdd47.Range("16:16")

Dim ptReduses As PivotTable     'таблиця з переліком МЦ
'оновити таблицю, перед вибором МЦ
Set ptReduses = ThisWorkbook.Worksheets("Довідник д47").PivotTables("ptReduses47")
ptReduses.RefreshTable

'Видаляємо попередні дані
Dim lastColumnStart As Long
 lastColumnStart = wsAdd47.Cells(19, wsAdd47.Columns.count).End(xlToLeft).Column + 1
Dim LastRowStart As Long
 LastRowStart = wsAdd47.Cells.Find(What:="*", After:=wsAdd47.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
                            SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
On Error Resume Next
    wsAdd47.Range(Cells(1, 20), Cells(LastRowStart, lastColumnStart)).EntireColumn.Delete ' Видалити колонки від U до останньої

    wsAdd47.Rows("21:" & wsAdd47.Rows.count).Delete

On Error GoTo 0
    
Dim filteredData() As Variant 'масив з відфільтрованими даними

'викликаємо форму для вибору майна
     Dim frm     As UserForm1
        Set frm = UserForms.Add(UserForm1.name)
        frm.ListData = ThisWorkbook.Worksheets("Довідник д47").Range("Goods47")
        
        frm.show
        
        NameOfThings = frm.TextBox1.value
        
        Dim rngWithName As Range
        
        Dim ListGoods As Range
        Set ListGoods = ThisWorkbook.Worksheets("Довідник д47").Range("Goods47")
        'Шукаємо задане значення в діапазоні
        Set rngWithName = ListGoods.Find(NameOfThings, LookIn:=xlValues, LookAt:=xlWhole)
        
        'Перевіряємо, чи знайдено задане значення
        If Not rngWithName Is Nothing Then
            Dim name As String
            Dim Nomenclature As String
            Dim unit As String
            Dim Pice As Double
            
            
            'Визначаємо значення за замовчуванням
            name = rngWithName.Offset(, -12).value
            
'            On Error GoTo WrongNomenclature
            Nomenclature = rngWithName.Offset(, -10).value
            If Nomenclature = "(пусто)" Then
            MsgBox "Ймовірно, ви не вказали код номенклатури вибраного майна, вкажіть код номенклатури (якщо не відомо, поставте б/н) і спробуйте створити додаток ще раз!", vbCritical, "Код номенклатури не вказаний"
            Exit Sub
            End If
'            On Error GoTo 0
            
'            On Error GoTo WrongUnit
            unit = rngWithName.Offset(, -11).value
            If unit = "(пусто)" Then
            MsgBox "Ймовірно, ви не вказали одиницю виміру вибраного майна, вкажіть одиницю виміру і спробуйте створити додаток ще раз!", vbCritical, "Одиниця виміру не вказана"
            Exit Sub
            End If
'            On Error GoTo 0
            
            On Error GoTo WrongPrice
            Pice = rngWithName.Offset(, -9).value
            On Error GoTo 0
            
'заповнюємо критерії, по яких будемо фільтрувати базу даних для витяжки потрібного майна
            Dim criteriaRange47 As Range
            
            Set criteriaRange47 = wsAdd47.Range("criteriaRange47")
            
            With criteriaRange47
                .Cells(2, 1) = "'=" & name 'найменуванян
                .Cells(2, 2) = "'=" & unit 'одиниця виміру
                .Cells(2, 3) = "'=" & Nomenclature    'код номенклатури
                .Cells(2, 4) = Pice  'ціна
            End With
            
            'знімаємо фільтри,якщо були поставлені
            If wsAllData.FilterMode Then
                wsAllData.ShowAllData
            End If
            
            'фільтруємо по заданих критеріях
            tbAllData.Range.AdvancedFilter Action:=xlFilterInPlace, criteriaRange:=criteriaRange47, Unique:=False
            
Dim rowCount As Long, colCount As Long
Dim visibleRange As Range
Dim cell As Range
Dim area As Range

 ' Визначаємо кількість видимих рядків і стовпців
    rowCount = 0
    colCount = tbAllData.ListColumns.count
    On Error Resume Next
    Set visibleRange = tbAllData.DataBodyRange.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    
    If Not visibleRange Is Nothing Then
        For Each area In visibleRange.Areas
            rowCount = rowCount + area.Rows.count
        Next area
        
        ' Ініціалізуємо масив з урахуванням заголовків
        ReDim filteredData(1 To rowCount + 1, 1 To colCount)
        
        ' Записуємо заголовки таблиці у масив
        For k = 1 To colCount
            filteredData(1, k) = tbAllData.HeaderRowRange.Cells(1, k).value
        Next k
        
        ' Записуємо відфільтровані дані у масив
        i = 2
        For Each area In visibleRange.Areas
            For Each cell In area.Rows
                For j = 1 To colCount
                    filteredData(i, j) = CStr(cell.Cells(1, j).text)
                Next j
                i = i + 1
            Next cell
        Next area
            
  End If
        
End If

Dim lastRow47 As Long
Dim CopyRange1 As Range
Dim foundCell As Range

'витягуємо дані з кожного рядка відфільтрованої бази даних
Dim n As Long
On Error GoTo DontSelect
 For n = LBound(filteredData) + 1 To UBound(filteredData)
 'якщо у вас помилка в цьому місці, скоріше всього вибране майно відсутнє в базі, або ви не підтвердили вибір майна:
 'вибране майно повинне "перестрибнути" у верхнє поле
 On Error GoTo 0
 
    lastRow47 = Cells(wsAdd47.Rows.count, "A").End(xlUp).Row + 1
    
    
    
    
    
    Dim LastColumn As Long
    LastColumn = wsAdd47.Cells.Find(What:="*", After:=wsAdd47.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
  
       
       'Якщо один і той же документ і ті ж підрозділи тоді рядок не додаємо
       If wsAdd47.Cells(lastRow47 - 1, 1).value = filteredData(n, 1) And _
       wsAdd47.Cells(lastRow47 - 1, 2).value = filteredData(n, 25) And _
       wsAdd47.Cells(lastRow47 - 1, 3).value = filteredData(n, 14) And _
       wsAdd47.Cells(lastRow47 - 1, 4).value = filteredData(n, 13) And _
       wsAdd47.Cells(lastRow47 - 1, 5).value = filteredData(n, 8) & Chr(10) & "(" & filteredData(n, 9) & ")" _
       Then
       
        lastRow47 = lastRow47 - 1
       
       Else
        wsAdd47.Range(wsAdd47.Cells(lastRow47 - 1, 14), wsAdd47.Cells(lastRow47 - 1, LastColumn)).Copy
        wsAdd47.Range(wsAdd47.Cells(lastRow47, 14), wsAdd47.Cells(lastRow47, LastColumn)).PasteSpecial xlPasteAll
       
       End If
            
  
  
  
  '================================================
  'скопіювати останій рядок і вставити
'    wsAdd47.Range(wsAdd47.Cells(lastRow47 - 1, 14), wsAdd47.Cells(lastRow47 - 1, LastColumn)).Copy
'    wsAdd47.Range(wsAdd47.Cells(lastRow47, 14), wsAdd47.Cells(lastRow47, LastColumn)).PasteSpecial xlPasteAll
'    wsAdd47.Cells(lastRow47, LastColumn)
    '==================================================
    
    Dim CategoryFrom As String
    Dim CategoryTo As String
    Dim categoryCol As Long
    'обнуляємо колонку з категоріями
    categoryCol = 0
    ' категорія відправника і отримувача з бази даних
    CategoryFrom = filteredData(n, 5)
    CategoryTo = filteredData(n, 27)
 
    
           Dim typeOperation As String
           typeOperation = filteredData(n, 26)
           
            Dim quantity As Double
            Dim FromUnit As String
            Dim toUnit As String
            Dim fromUnitExist As Boolean
            Dim toUnitExist As Boolean
            
            Dim previousQuantyty As Double
            Dim actualQuantyty As Double
            ' заголовок з назвою майна і категоріями
            Set CopyRange1 = wsAdd47.Range(wsAdd47.Cells(16, 14), wsAdd47.Cells(20, 19))
            
            Select Case typeOperation
    '*********************************************************************************************************
               Case "надходження"
                        quantity = filteredData(n, 6) 'кількість
                        toUnit = filteredData(n, 9) 'куди (підрозділ)
                        wsAdd47.Cells(lastRow47, 6).value = wsAdd47.Cells(lastRow47, 6).value + quantity
                        toUnitExist = FindValue(toUnit, RngToSearchData)
                        
                        
                    'якщо не знайшли підрозділ то копіюємо набір даних для підрозділу
                    If Not toUnitExist Then
'                        MsgBox "Працюємо"
                        wsAdd47.Activate

                        'копіюємо стовпчики
                       
                            CopyRange1.Copy
                            
                             wsAdd47.Cells(16, LastColumn + 1).PasteSpecial xlPasteAll
                            
                             wsAdd47.Cells(16, LastColumn + 1) = toUnit
                             

                    'колонка з потрібною категорією
                    
    'якщо тут помилка - в базі не правильно вказана категорія:
    'повинно бути 1 або 2 або 3 або 4 або 5,
    'в текстовому форматі(зліва зверху в клітинці з категорією зелений трикутник)
    'переконайтеся що у вас правильно проставлені категорії відправника або одержувача
                    categoryCol = wsAdd47.Cells.Find(What:=CategoryTo, After:=wsAdd47.Cells(18, LastColumn), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext).Column
                    
                    previousQuantyty = wsAdd47.Cells(lastRow47, categoryCol).value
'                    previousQuantyty = wsAdd47.Cells(lastRow47, categoryCol).End(xlUp).value 'визначаємо попередню кількість, за відповідною категорією.
                    
                    actualQuantyty = previousQuantyty + quantity 'додаємо до попередньої кількості
                    'вставляємо в поточний рядок
                    wsAdd47.Cells(lastRow47, categoryCol).value = actualQuantyty
                    
                         Else
                        
                        
                        
                        
                        
                        

                            'Шукаємо підрозділ в діапазоні
                            Set foundCell = RngToSearchData.Find(toUnit, LookIn:=xlValues, LookAt:=xlWhole)
                        
                            'Перевіряємо, чи знайдено підрозділ
                            If Not foundCell Is Nothing Then
                            
                            'тут вставити перевірку чи ж вже такий документ
                            
                            'якщо знайдений
                            
                            'якщо не знайдений
                            

                            'якщо тут помилка - в базі не правильно вказана категорія:
                            'повинно бути 1 або 2 або 3 або 4 або 5,
                            'в текстовому форматі(зліва зверху в клітинці з категорією зелений трикутник)
                            'переконайтеся що у вас правильно проставлені категорії відправника або одержувача
                                categoryCol = wsAdd47.Cells.Find(What:=CategoryTo, After:=wsAdd47.Cells(18, foundCell.Column), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext).Column
                                previousQuantyty = wsAdd47.Cells(lastRow47, categoryCol).value
'                                previousQuantyty = wsAdd47.Cells(lastRow47, categoryCol).End(xlUp).value 'визначаємо попередню кількість, за відповідною категорією.
                                actualQuantyty = previousQuantyty + quantity 'додаємо до попередньої кількості
                                'вставляємо в поточний рядок
                                wsAdd47.Cells(lastRow47, categoryCol).value = actualQuantyty
                               
                            Else
'                                FindValue = False
                            End If
 
                    End If
'******************************************************************************************************************************
                        
               Case "вибуття"
                        quantity = filteredData(n, 7)
                        wsAdd47.Cells(lastRow47, 7).value = wsAdd47.Cells(lastRow47, 7).value + quantity
                        FromUnit = filteredData(n, 8)
                        fromUnitExist = FindValue(FromUnit, RngToSearchData)
                        
                        
                    If Not fromUnitExist Then
'                        MsgBox "Працюємо"
                        wsAdd47.Activate
 
                                             
                            CopyRange1.Copy
                            wsAdd47.Cells(16, LastColumn + 1).PasteSpecial xlPasteAll
                            
                             wsAdd47.Cells(16, LastColumn + 1) = FromUnit
                    
                    'якщо тут помилка - в базі не правильно вказана категорія:
                    'повинно бути 1 або 2 або 3 або 4 або 5,
                    'в текстовому форматі(зліва зверху в клітинці з категорією зелений трикутник)
                    'переконайтеся що у вас правильно проставлені категорії відправника або одержувача
                    categoryCol = wsAdd47.Cells.Find(What:=CategoryFrom, After:=wsAdd47.Cells(18, LastColumn), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext).Column
                    previousQuantyty = wsAdd47.Cells(lastRow47, categoryCol).value
'                    previousQuantyty = wsAdd47.Cells(lastRow47, categoryCol).End(xlUp).value  'визначаємо попередню кількість, за відповідною категорією.
                    actualQuantyty = previousQuantyty - quantity 'віднімаємо до попередньої кількості
                    'вставляємо в поточний рядок
                    wsAdd47.Cells(lastRow47, categoryCol).value = actualQuantyty
'

                         Else

                            'Шукаємо задане значення в діапазоні, якщо вже є підрозділ
                            Set foundCell = RngToSearchData.Find(FromUnit, LookIn:=xlValues, LookAt:=xlWhole)
                        
                            'Перевіряємо, чи знайдено задане значення
                            If Not foundCell Is Nothing Then
                            
                                'якщо тут помилка - в базі не правильно вказана категорія:
                                'повинно бути 1 або 2 або 3 або 4 або 5,
                                'в текстовому форматі(зліва зверху в клітинці з категорією зелений трикутник)
                                'переконайтеся що у вас правильно проставлені категорії відправника або одержувача
                                categoryCol = wsAdd47.Cells.Find(What:=CategoryFrom, After:=wsAdd47.Cells(18, foundCell.Column), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext).Column
                                previousQuantyty = wsAdd47.Cells(lastRow47, categoryCol).value
'                                previousQuantyty = wsAdd47.Cells(lastRow47, categoryCol).End(xlUp).value  'визначаємо попередню кількість, за відповідною категорією.
                                actualQuantyty = previousQuantyty - quantity 'віднімаємо від попередньої кількості
                                
                                'вставляємо в поточний рядок кількість
                                wsAdd47.Cells(lastRow47, categoryCol).value = actualQuantyty
                               
                            Else
'                                FindValue = False
                            End If
                    End If
                    
                  
                        '***************************************************************************************************************
     
               Case "внутрішнє переміщення"
                        quantity = filteredData(n, 7)
                        toUnit = filteredData(n, 9)
                        FromUnit = filteredData(n, 8)
                        
                        fromUnitExist = FindValue(FromUnit, RngToSearchData)
                        toUnitExist = FindValue(toUnit, RngToSearchData)
                        
                    'для підрозділу, з якого вибуває
                        If Not fromUnitExist Then
'                        MsgBox "Працюємо"
                        wsAdd47.Activate
                        
                         
                            CopyRange1.Copy
                            wsAdd47.Cells(16, LastColumn + 1).PasteSpecial xlPasteAll
                            
                             wsAdd47.Cells(16, LastColumn + 1) = FromUnit
                            
                    'якщо тут помилка - в базі не правильно вказана категорія:
                    'повинно бути 1 або 2 або 3 або 4 або 5,
                    'в текстовому форматі(зліва зверху в клітинці з категорією зелений трикутник)
                    'переконайтеся що у вас правильно проставлені категорії відправника або одержувача
                    categoryCol = wsAdd47.Cells.Find(What:=CategoryFrom, After:=wsAdd47.Cells(18, LastColumn), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext).Column
                    previousQuantyty = wsAdd47.Cells(lastRow47, categoryCol).value
'                    previousQuantyty = wsAdd47.Cells(lastRow47, categoryCol).End(xlUp).value  'визначаємо попередню кількість, за відповідною категорією.
                    actualQuantyty = previousQuantyty - quantity 'віднімаємо до попередньої кількості
                    'вставляємо в поточний рядок
                    wsAdd47.Cells(lastRow47, categoryCol).value = actualQuantyty
                    

                         Else
                         
                            'Шукаємо задане значення в діапазоні
                            Set foundCell = RngToSearchData.Find(FromUnit, LookIn:=xlValues, LookAt:=xlWhole)
                        
                            'Перевіряємо, чи знайдено задане значення
                            If Not foundCell Is Nothing Then
                                'якщо тут помилка - в базі не правильно вказана категорія:
                                'повинно бути 1 або 2 або 3 або 4 або 5,
                                'в текстовому форматі(зліва зверху в клітинці з категорією зелений трикутник)
                                'переконайтеся що у вас правильно проставлені категорії відправника або одержувача
                                categoryCol = wsAdd47.Cells.Find(What:=CategoryFrom, After:=wsAdd47.Cells(18, foundCell.Column), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext).Column
                                previousQuantyty = wsAdd47.Cells(lastRow47, categoryCol).value
'                                previousQuantyty = wsAdd47.Cells(lastRow47, categoryCol).End(xlUp).value  'визначаємо попередню кількість, за відповідною категорією.
                                actualQuantyty = previousQuantyty - quantity 'додаємо до попередньої кількості
                                'вставляємо в поточний рядок
                                wsAdd47.Cells(lastRow47, categoryCol).value = actualQuantyty
                               
                            Else
'                                FindValue = False
                            End If

                    End If
                    
                    
                       If Not toUnitExist Then
'                        MsgBox "Працюємо"
                        wsAdd47.Activate
                        Dim addCol As Long
                        addCol = 0
                        
                        If Not fromUnitExist Then
                        
                        addCol = 6
                        End If
                        

                            CopyRange1.Copy
                            wsAdd47.Cells(16, LastColumn + addCol + 1).PasteSpecial xlPasteAll
                            
                             wsAdd47.Cells(16, LastColumn + addCol + 1) = toUnit
                             
                             
                     'якщо тут помилка - в базі не правильно вказана категорія:
                    'повинно бути 1 або 2 або 3 або 4 або 5,
                    'в текстовому форматі(зліва зверху в клітинці з категорією зелений трикутник)
                    'переконайтеся що у вас правильно проставлені категорії відправника або одержувача
                    categoryCol = wsAdd47.Cells.Find(What:=CategoryTo, After:=wsAdd47.Cells(18, LastColumn + addCol), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext).Column
                   
                    previousQuantyty = wsAdd47.Cells(lastRow47, categoryCol).value
'                    previousQuantyty = wsAdd47.Cells(lastRow47, categoryCol).End(xlUp).value  'визначаємо попередню кількість, за відповідною категорією.
                    actualQuantyty = previousQuantyty + quantity 'додаємо до попередньої кількості
                    
                    
                    'вставляємо в поточний рядок
                    wsAdd47.Cells(lastRow47, categoryCol).value = actualQuantyty
                 
                         Else

                            'Шукаємо задане значення в діапазоні
                            Set foundCell = RngToSearchData.Find(toUnit, LookIn:=xlValues, LookAt:=xlWhole)
                        
                            'Перевіряємо, чи знайдено задане значення
                            If Not foundCell Is Nothing Then
                            
                                'якщо тут помилка - в базі не правильно вказана категорія:
                                'повинно бути 1 або 2 або 3 або 4 або 5,
                                'в текстовому форматі(зліва зверху в клітинці з категорією зелений трикутник)
                                'переконайтеся що у вас правильно проставлені категорії відправника або одержувача
                                categoryCol = wsAdd47.Cells.Find(What:=CategoryTo, After:=wsAdd47.Cells(18, foundCell.Column), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext).Column
                                previousQuantyty = wsAdd47.Cells(lastRow47, categoryCol).value
'                                previousQuantyty = wsAdd47.Cells(lastRow47, categoryCol).End(xlUp).value  'визначаємо попередню кількість, за відповідною категорією.
                                actualQuantyty = previousQuantyty + quantity 'додаємо до попередньої кількості
                                'вставляємо в поточний рядок
                                wsAdd47.Cells(lastRow47, categoryCol).value = actualQuantyty

                            End If
    
                    End If
'
            End Select
'            Dim NDoc As String
'            NDoc = CStr()
            With wsAdd47
                .Cells(lastRow47, 1).value = filteredData(n, 1) 'дата внесення
                .Cells(lastRow47, 2).value = filteredData(n, 25) 'найменування документа
                .Cells(lastRow47, 3).NumberFormat = "@"
                .Cells(lastRow47, 3).value = filteredData(n, 14) 'номер документа
                .Cells(lastRow47, 4).value = filteredData(n, 13)  'дата документа
                .Cells(lastRow47, 5).value = filteredData(n, 8) & Chr(10) & "(" & filteredData(n, 9) & ")"  'підрозділ звідки (підрозділ куди)
        
            End With
             
    
         'копіюємо стилі загальної іформації
        wsAdd47.Range(Cells(20, 1), Cells(20, 19)).Copy
        wsAdd47.Range(Cells(lastRow47, 1), Cells(lastRow47, 19)).PasteSpecial Paste:=xlPasteFormats
  
        Set foundCell = Nothing
    
    Call ProgressBar.Progress(n / UBound(filteredData) * 100)
    
    
    Next n
    
    'знаходимо останню колонку після всіх маніпуляцій
     Dim LastColumnFinal As Long
     LastColumnFinal = wsAdd47.Cells.Find(What:="*", After:=wsAdd47.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
     'копіюємо стилі основної інформації (для показників підрозділів)
       wsAdd47.Range(Cells(20, 14), Cells(20, 19)).Copy
       wsAdd47.Range(Cells(21, 14), Cells(lastRow47, LastColumnFinal)).PasteSpecial Paste:=xlPasteFormats
  

       'проставляємо суму
       Dim cellAll As Range
       Dim cellWithSum As Range
       For Each cellAll In wsAdd47.Rows(17).Cells 'рядок з усього
        If cellAll.value = "Усього" Then
            Set cellWithSum = cellAll.Offset(2, 0)
            cellWithSum.AutoFill destination:=wsAdd47.Range(wsAdd47.Cells(cellWithSum.Row, cellWithSum.Column), wsAdd47.Cells(lastRow47, cellWithSum.Column))
        End If
    Next cellAll
    
   
    
    Dim sumRange As Range
    Dim criteriaRange As Range
    Dim criteriaValue As Variant
    
    
   'формула по категоріях
   Dim result As Double
    For j = 9 To 13
        For i = 21 To lastRow47
            
            Set sumRange = wsAdd47.Range(wsAdd47.Cells(i, 14), wsAdd47.Cells(i, LastColumnFinal))
            Set criteriaRange = wsAdd47.Range(wsAdd47.Cells(18, 14), wsAdd47.Cells(18, LastColumnFinal))
            criteriaValue = wsAdd47.Cells(18, j).value
        
'            wsAdd47.Range(Cells(i, j), Cells(lastRow47, j)).Formula = "=SUMIFS(" & wsAdd47.Range(Cells(i, 14), Cells(i, LastColumnFinal)).Address & "," & wsAdd47.Range(Cells(18, 14), Cells(18, LastColumnFinal)).Address & "," & wsAdd47.Cells(18, j).Address & ")"
        
        result = Application.WorksheetFunction.SumIfs(sumRange, criteriaRange, criteriaValue)
        wsAdd47.Cells(i, j).value = result
        Next i
    Next j
 'скидаємо область друку
wsAdd47.PageSetup.PrintArea = ""
'скидаємо фільтри
On Error Resume Next
wsAllData.ShowAllData
Exit Sub
'***помилки****************************************
DontSelect:
MsgBox "Ймовірно,майно відсутнє в базі, або Ви не підтвердили вибір майна: вибране майно повинне ""перестрибнути"" з нижнього у верхнє поле вікна вибору майна, спробуйтие сформувати додаток ще раз!", vbCritical, "Майно не знайдено"
'wsAllData.ShowAllData
Exit Sub

WrongPrice:
MsgBox "Ймовірно,ви не вказали ціну вибраного майна, вкажіть ціну (якщо не відома, поставте 1 грн.) і спробуйте створити додаток ще раз!", vbCritical, "Ціна не вказана"
Exit Sub


End Sub

Sub SaveAdd47()

        ActiveSheet.Select 'копіюємо
        ActiveSheet.Copy 'зберігаємо як нову книгу
     
     Dim WbNewWO As Workbook
     
        Set WbNewWO = ActiveWorkbook
    
    Dim nameAdd47 As String
    Dim nomenclature47 As String
    Dim unit47 As String
    Dim price47 As Double
    Dim formattedPrice As String
    
    nameAdd47 = WbNewWO.Worksheets(1).Range("nameAdd47")
    
    nomenclature47 = WbNewWO.Worksheets(1).Range("nomenclature47")
    unit47 = WbNewWO.Worksheets(1).Range("unit47")
    price47 = WbNewWO.Worksheets(1).Range("price47")
    
    formattedPrice = Format(price47, "#,##0.0000")

    Dim shp As Shape
    For Each shp In WbNewWO.Worksheets(1).Shapes
        shp.Delete
    Next shp
        
      Dim fileName As String, fileName2 As String, lineBrake As String
        'Шлях до файла зберігання
        SavePath = WbNewWO.Worksheets(1).Range("PathToSaveAdd47")
        Dim DateAdd47 As Date
        DateAdd47 = Format(Date, "dd.mm.yyyy")
        
        
              On Error Resume Next
         If Dir(SavePath, vbDirectory) = vbNullString Then
         On Error GoTo 0
        With Application.FileDialog(msoFileDialogFolderPicker)
            .Title = "Виберіть папку для збереження"
            .AllowMultiSelect = False
            If .show = -1 Then
                SavePath = .SelectedItems(1)
                'Записуємо новий шлях у клітинку
                ThisWorkbook.Sheets("Додаток 47").Range("PathToSaveAdd47").value = SavePath
            Else
                MsgBox "Збереження скасовано", vbExclamation
                WbNew.Close SaveChanges:=False
                Application.DisplayAlerts = True
                Exit Sub
            End If
        End With
    End If
        
        
        
        'складаємо назву файла
        fileName = SavePath & Application.PathSeparator & CleanFileName(DateAdd47 & " д47 " & nameAdd47 & " " & nomenclature47 & " " & unit47 & " " & formattedPrice & ".xlsx")
        'заміняємо перенесення рядка, косі лінії
'        fileName2 = CleanFileName(fileName)
        
        'зберігаємо файл
        On Error GoTo WrongPath
         WbNewWO.SaveAs fileName:=fileName, FileFormat:=xlOpenXMLWorkbook
         'якщо у вас тут помилка, ймовіно ви не змінили шлях збереження файлу, або допустити помилку в шляху збереження (наприклад лишній знак "/" вкінці)
        On Error GoTo 0
        Exit Sub
WrongPath:
     MsgBox "Ймовірно, не правильно вказана адреса збереження файла, виправте адресу збереження файлу( на зразок ""D:\База обліку майна ЗСУ 2024"" *без косої лінії вкінці) і спробуйтие сформувати додаток ще раз!", vbCritical, "Не правильний шлях збереження"
     Exit Sub
End Sub








```

### 5.3. Вихідний код add47PNP.bas
```vba
Attribute VB_Name = "add47PNP"
Sub CreateAdd47PNP()
'для категорій придатне/непридатне/потребує ремонту
Dim wbAllData As Workbook
Dim wsAllData As Worksheet
Dim tbAllData As ListObject
    Set wbAllData = ThisWorkbook
    Set wsAllData = wbAllData.Worksheets("База")
    Set tbAllData = wsAllData.ListObjects("AllData")
    
Dim wsAdd47 As Worksheet
    Set wsAdd47 = wbAllData.Worksheets("Додаток 47 пр-непр")
    
    
Dim RngToSearchData As Range
Set RngToSearchData = wsAdd47.Range("16:16")

Dim ptReduses As PivotTable     'таблиця з переліком МЦ
'оновити таблицю, перед вибором МЦ
Set ptReduses = ThisWorkbook.Worksheets("Довідник д47").PivotTables("ptReduses47")
ptReduses.RefreshTable

'Видаляємо попередні дані
Dim lastColumnStart As Long
 lastColumnStart = wsAdd47.Cells(19, wsAdd47.Columns.count).End(xlToLeft).Column + 1
Dim LastRowStart As Long
 LastRowStart = wsAdd47.Cells.Find(What:="*", After:=wsAdd47.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
                            SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
On Error Resume Next
    wsAdd47.Range(Cells(1, 16), Cells(LastRowStart, lastColumnStart)).EntireColumn.Delete ' Видалити колонки від до останньої

    wsAdd47.Rows("21:" & wsAdd47.Rows.count).Delete

On Error GoTo 0
    
Dim filteredData() As Variant 'масив з відфільтрованими даними

'викликаємо форму для вибору майна
     Dim frm     As UserForm1
        Set frm = UserForms.Add(UserForm1.name)
        frm.ListData = ThisWorkbook.Worksheets("Довідник д47").Range("Goods47")
        
        frm.show
        
        NameOfThings = frm.TextBox1.value
        
        Dim rngWithName As Range
        
        Dim ListGoods As Range
        Set ListGoods = ThisWorkbook.Worksheets("Довідник д47").Range("Goods47")
        'Шукаємо задане значення в діапазоні
        Set rngWithName = ListGoods.Find(NameOfThings, LookIn:=xlValues, LookAt:=xlWhole)
        
        'Перевіряємо, чи знайдено задане значення
        If Not rngWithName Is Nothing Then
            Dim name As String
            Dim Nomenclature As String
            Dim unit As String
            Dim Pice As Double
            
            
            'Визначаємо значення за замовчуванням
            name = rngWithName.Offset(, -12).value
            
'            Nomenclature = rngWithName.Offset(, -10).value
'
'            Unit = rngWithName.Offset(, -11).value
'
'            Pice = rngWithName.Offset(, -9).value
            
'             On Error GoTo WrongNomenclature
            Nomenclature = rngWithName.Offset(, -10).value
            If Nomenclature = "(пусто)" Then
            MsgBox "Ймовірно, ви не вказали код номенклатури вибраного майна, вкажіть код номенклатури (якщо не відомо, поставте б/н) і спробуйте створити додаток ще раз!", vbCritical, "Код номенклатури не вказаний"
            Exit Sub
            End If
'            On Error GoTo 0
            
'            On Error GoTo WrongUnit
            unit = rngWithName.Offset(, -11).value
            If unit = "(пусто)" Then
            MsgBox "Ймовірно, ви не вказали одиницю виміру вибраного майна, вкажіть одиницю виміру і спробуйте створити додаток ще раз!", vbCritical, "Одиниця виміру не вказана"
            Exit Sub
            End If
'            On Error GoTo 0
            
            On Error GoTo WrongPrice
            Pice = rngWithName.Offset(, -9).value
            On Error GoTo 0
            
            
           'заповнюємо критерії, по яких будемо фільтрувати базу даних для витяжки потрібного майна
            Dim criteriaRange47 As Range
            
            Set criteriaRange47 = wsAdd47.Range("criteriaRange47PNP")
            
            With criteriaRange47
                .Cells(2, 1) = "'=" & name 'найменуванян
                .Cells(2, 2) = "'=" & unit 'одиниця виміру
                .Cells(2, 3) = "'=" & Nomenclature    'код номенклатури
                .Cells(2, 4) = Pice  'ціна
            End With
            
            'знімаємо фільтри,якщо були поставлені
            If wsAllData.FilterMode Then
                wsAllData.ShowAllData
            End If
            
            'фільтруємо по заданих критеріях
            tbAllData.Range.AdvancedFilter Action:=xlFilterInPlace, criteriaRange:=criteriaRange47, Unique:=False
            
Dim rowCount As Long, colCount As Long
Dim visibleRange As Range
Dim cell As Range
Dim area As Range

 ' Визначаємо кількість видимих рядків і стовпців
    rowCount = 0
    colCount = tbAllData.ListColumns.count
    On Error Resume Next
    Set visibleRange = tbAllData.DataBodyRange.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    
    If Not visibleRange Is Nothing Then
        For Each area In visibleRange.Areas
            rowCount = rowCount + area.Rows.count
        Next area
        
        ' Ініціалізуємо масив з урахуванням заголовків
        ReDim filteredData(1 To rowCount + 1, 1 To colCount)
        
        ' Записуємо заголовки таблиці у масив
        For k = 1 To colCount
            filteredData(1, k) = tbAllData.HeaderRowRange.Cells(1, k).value
        Next k
        
        ' Записуємо відфільтровані дані у масив
        i = 2
        For Each area In visibleRange.Areas
            For Each cell In area.Rows
                For j = 1 To colCount
                    filteredData(i, j) = cell.Cells(1, j).value
                Next j
                i = i + 1
            Next cell
        Next area
            
  End If
        
End If
Dim lastRow47 As Long
Dim CopyRange1 As Range
Dim foundCell As Range

Dim n As Long
    On Error GoTo DontSelect
    For n = LBound(filteredData) + 1 To UBound(filteredData)
    On Error GoTo 0
    lastRow47 = Cells(wsAdd47.Rows.count, "A").End(xlUp).Row + 1
    Dim LastColumn As Long
    LastColumn = wsAdd47.Cells.Find(What:="*", After:=wsAdd47.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    
    'скопіювати останій рядок і вставити
    wsAdd47.Range(wsAdd47.Cells(lastRow47 - 1, 12), wsAdd47.Cells(lastRow47 - 1, LastColumn)).Copy
    wsAdd47.Range(wsAdd47.Cells(lastRow47, 12), wsAdd47.Cells(lastRow47, LastColumn)).PasteSpecial xlPasteAll
'    wsAdd47.Cells(lastRow47, LastColumn)
    
    Dim CategoryFrom As String
    Dim CategoryTo As String
    Dim categoryCol As Long
    categoryCol = 0
    
    CategoryFrom = filteredData(n, 5)
    CategoryTo = filteredData(n, 27)
 
    
           Dim typeOperation As String
           typeOperation = filteredData(n, 26)
           
            Dim quantity As Double
            Dim FromUnit As String
            Dim toUnit As String
            Dim fromUnitExist As Boolean
            Dim toUnitExist As Boolean
            
            Dim previousQuantyty As Double
            Dim actualQuantyty As Double
            
            Set CopyRange1 = wsAdd47.Range(wsAdd47.Cells(16, 12), wsAdd47.Cells(20, 15))
            
            Select Case typeOperation
    '*********************************************************************************************************
               Case "надходження"
                        quantity = filteredData(n, 6) 'кількість
                        toUnit = filteredData(n, 9) 'куди (підрозділ)
                        wsAdd47.Cells(lastRow47, 6).value = quantity
                        toUnitExist = FindValue(toUnit, RngToSearchData)
                        
                    If Not toUnitExist Then
'                        MsgBox "Працюємо"
                        wsAdd47.Activate

                        'копіюємо стовпчики
                       
                            CopyRange1.Copy
                            
                             wsAdd47.Cells(16, LastColumn + 1).PasteSpecial xlPasteAll
                            
                             wsAdd47.Cells(16, LastColumn + 1) = toUnit
                             
                    'якщо тут помилка - в базі не правильно вказана категорія:
                    'повинно бути пр. або не пр. або пот. рем.
                    'переконайтеся що у вас правильно проставлені категорії відправника або одержувача
                    categoryCol = wsAdd47.Cells.Find(What:=CategoryTo, After:=wsAdd47.Cells(18, LastColumn), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext).Column
                    
                    previousQuantyty = wsAdd47.Cells(lastRow47, categoryCol).value
'                    previousQuantyty = wsAdd47.Cells(lastRow47, categoryCol).End(xlUp).value 'визначаємо попередню кількість, за відповідною категорією.
                    actualQuantyty = previousQuantyty + quantity 'додаємо до попередньої кількості
                    'вставляємо в поточний рядок
                    wsAdd47.Cells(lastRow47, categoryCol).value = actualQuantyty
                   
                         Else
                        

                            'Шукаємо задане значення в діапазоні
                            Set foundCell = RngToSearchData.Find(toUnit, LookIn:=xlValues, LookAt:=xlWhole)
                        
                            'Перевіряємо, чи знайдено задане значення
                            If Not foundCell Is Nothing Then
                    
                                'якщо тут помилка - в базі не правильно вказана категорія:
                                'повинно бути пр. або не пр. або пот. рем.
                                'переконайтеся що у вас правильно проставлені категорії відправника або одержувача
                                categoryCol = wsAdd47.Cells.Find(What:=CategoryTo, After:=wsAdd47.Cells(18, foundCell.Column), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext).Column
                                previousQuantyty = wsAdd47.Cells(lastRow47, categoryCol).value
'                                previousQuantyty = wsAdd47.Cells(lastRow47, categoryCol).End(xlUp).value 'визначаємо попередню кількість, за відповідною категорією.
                                actualQuantyty = previousQuantyty + quantity 'додаємо до попередньої кількості
                                'вставляємо в поточний рядок
                                wsAdd47.Cells(lastRow47, categoryCol).value = actualQuantyty
                               
                            Else
'                                FindValue = False
                            End If
 
                    End If
'******************************************************************************************************************************
                        
               Case "вибуття"
                        quantity = filteredData(n, 7)
                        wsAdd47.Cells(lastRow47, 7).value = quantity
                        FromUnit = filteredData(n, 8)
                        fromUnitExist = FindValue(FromUnit, RngToSearchData)
                        
                        
                    If Not fromUnitExist Then
'                        MsgBox "Працюємо"
                        wsAdd47.Activate
 
                    
                            CopyRange1.Copy
                            wsAdd47.Cells(16, LastColumn + 1).PasteSpecial xlPasteAll
                            
                             wsAdd47.Cells(16, LastColumn + 1) = FromUnit
                    'якщо тут помилка - в базі не правильно вказана категорія:
                    'повинно бути пр. або не пр. або пот. рем.
                    'переконайтеся що у вас правильно проставлені категорії відправника або одержувача
                    categoryCol = wsAdd47.Cells.Find(What:=CategoryFrom, After:=wsAdd47.Cells(18, LastColumn), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext).Column
                    previousQuantyty = wsAdd47.Cells(lastRow47, categoryCol).value
'                    previousQuantyty = wsAdd47.Cells(lastRow47, categoryCol).End(xlUp).value  'визначаємо попередню кількість, за відповідною категорією.
                    actualQuantyty = previousQuantyty - quantity 'віднімаємо до попередньої кількості
                    'вставляємо в поточний рядок
                    wsAdd47.Cells(lastRow47, categoryCol).value = actualQuantyty


                         Else

                            'Шукаємо задане значення в діапазоні, якщо вже є підрозділ
                            Set foundCell = RngToSearchData.Find(FromUnit, LookIn:=xlValues, LookAt:=xlWhole)
                        
                            'Перевіряємо, чи знайдено задане значення
                            If Not foundCell Is Nothing Then
                                'якщо тут помилка - в базі не правильно вказана категорія:
                                'повинно бути пр. або не пр. або пот. рем.
                                'переконайтеся що у вас правильно проставлені категорії відправника або одержувача
                                categoryCol = wsAdd47.Cells.Find(What:=CategoryFrom, After:=wsAdd47.Cells(18, foundCell.Column), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext).Column
                                previousQuantyty = wsAdd47.Cells(lastRow47, categoryCol).value
'                                previousQuantyty = wsAdd47.Cells(lastRow47, categoryCol).End(xlUp).value  'визначаємо попередню кількість, за відповідною категорією.
                                actualQuantyty = previousQuantyty - quantity 'віднімаємо від попередньої кількості
                                
                                'вставляємо в поточний рядок кількість
                                wsAdd47.Cells(lastRow47, categoryCol).value = actualQuantyty
                               
                            Else
'                                FindValue = False
                            End If
                    End If
                    
                  
                        '***************************************************************************************************************
     
               Case "внутрішнє переміщення"
                        quantity = filteredData(n, 7)
                        toUnit = filteredData(n, 9)
                        FromUnit = filteredData(n, 8)
                        
                        fromUnitExist = FindValue(FromUnit, RngToSearchData)
                        toUnitExist = FindValue(toUnit, RngToSearchData)
                        
                    'для підрозділу, з якого вибуває
                        If Not fromUnitExist Then
'                        MsgBox "Працюємо"
                        wsAdd47.Activate
                        

  
                        
                            CopyRange1.Copy
                            wsAdd47.Cells(16, LastColumn + 1).PasteSpecial xlPasteAll
                            
                             wsAdd47.Cells(16, LastColumn + 1) = FromUnit
                    'якщо тут помилка - в базі не правильно вказана категорія:
                    'повинно бути пр. або не пр. або пот. рем.
                    'переконайтеся що у вас правильно проставлені категорії відправника або одержувача
                    categoryCol = wsAdd47.Cells.Find(What:=CategoryFrom, After:=wsAdd47.Cells(18, LastColumn), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext).Column
                    previousQuantyty = wsAdd47.Cells(lastRow47, categoryCol).value
'                    previousQuantyty = wsAdd47.Cells(lastRow47, categoryCol).End(xlUp).value  'визначаємо попередню кількість, за відповідною категорією.
                    actualQuantyty = previousQuantyty - quantity 'віднімаємо до попередньої кількості
                    'вставляємо в поточний рядок
                    wsAdd47.Cells(lastRow47, categoryCol).value = actualQuantyty
                    

                         Else
                         
'                        Dim foundCell As Range

                            'Шукаємо задане значення в діапазоні
                            Set foundCell = RngToSearchData.Find(FromUnit, LookIn:=xlValues, LookAt:=xlWhole)
                        
                            'Перевіряємо, чи знайдено задане значення
                            If Not foundCell Is Nothing Then
                                'якщо тут помилка - в базі не правильно вказана категорія:
                                'повинно бути пр. або не пр. або пот. рем.
                                'переконайтеся що у вас правильно проставлені категорії відправника або одержувача
                                categoryCol = wsAdd47.Cells.Find(What:=CategoryFrom, After:=wsAdd47.Cells(18, foundCell.Column), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext).Column
                                previousQuantyty = wsAdd47.Cells(lastRow47, categoryCol).value
'                                previousQuantyty = wsAdd47.Cells(lastRow47, categoryCol).End(xlUp).value  'визначаємо попередню кількість, за відповідною категорією.
                                actualQuantyty = previousQuantyty - quantity 'додаємо до попередньої кількості
                                'вставляємо в поточний рядок
                                wsAdd47.Cells(lastRow47, categoryCol).value = actualQuantyty
                               
                            Else
'                                FindValue = False
                            End If
                        
                    End If
                    
                    
                       If Not toUnitExist Then
'                        MsgBox "Працюємо"
                        wsAdd47.Activate
                        
                        
                        Dim addCol As Long
                        addCol = 0
                        
                        If Not fromUnitExist Then
                        
                        addCol = 4
                        End If
 
                            CopyRange1.Copy
                            wsAdd47.Cells(16, LastColumn + addCol + 1).PasteSpecial xlPasteAll
                            
                             wsAdd47.Cells(16, LastColumn + addCol + 1) = toUnit
                     'якщо тут помилка - в базі не правильно вказана категорія:
                    'повинно бути пр. або не пр. або пот. рем.
                    'переконайтеся що у вас правильно проставлені категорії відправника або одержувача
                    categoryCol = wsAdd47.Cells.Find(What:=CategoryTo, After:=wsAdd47.Cells(18, LastColumn + addCol - 1), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext).Column
                    '****** ще треба скопіювати весь попередній рядок...., щоб можна було віднімати (додавати) поточну кількість до попередньої
                    previousQuantyty = wsAdd47.Cells(lastRow47, categoryCol).value
'                    previousQuantyty = wsAdd47.Cells(lastRow47, categoryCol).End(xlUp).value  'визначаємо попередню кількість, за відповідною категорією.
                    actualQuantyty = previousQuantyty + quantity 'додаємо до попередньої кількості
                    'вставляємо в поточний рядок
                    wsAdd47.Cells(lastRow47, categoryCol).value = actualQuantyty
                   
                         Else
'                        Dim foundCell As Range

                            'Шукаємо задане значення в діапазоні
                            Set foundCell = RngToSearchData.Find(toUnit, LookIn:=xlValues, LookAt:=xlWhole)
                        
                            'Перевіряємо, чи знайдено задане значення
                            If Not foundCell Is Nothing Then
                                'якщо тут помилка - в базі не правильно вказана категорія:
                                'повинно бути пр. або не пр. або пот. рем.
                                'переконайтеся що у вас правильно проставлені категорії відправника або одержувача
                                categoryCol = wsAdd47.Cells.Find(What:=CategoryTo, After:=wsAdd47.Cells(18, foundCell.Column), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext).Column
                                previousQuantyty = wsAdd47.Cells(lastRow47, categoryCol).value
'                                previousQuantyty = wsAdd47.Cells(lastRow47, categoryCol).End(xlUp).value  'визначаємо попередню кількість, за відповідною категорією.
                                actualQuantyty = previousQuantyty + quantity 'додаємо до попередньої кількості
                                'вставляємо в поточний рядок
                                wsAdd47.Cells(lastRow47, categoryCol).value = actualQuantyty
                               
                            Else
'                                FindValue = False
                            End If
    
                    End If
'
'
            End Select
            
            With wsAdd47
                .Cells(lastRow47, 1).value = filteredData(n, 1) 'дата внесення
                .Cells(lastRow47, 2).value = filteredData(n, 25) 'найменування документа
                .Cells(lastRow47, 3).NumberFormat = "@"
                .Cells(lastRow47, 3).value = filteredData(n, 14) 'номер документа
                .Cells(lastRow47, 4).value = filteredData(n, 13)  'дата документа
                .Cells(lastRow47, 5).value = filteredData(n, 8) & Chr(10) & "(" & filteredData(n, 9) & ")"  'підрозділ звідки (підрозділ куди)
        
            End With
 
         'копіюємо стилі загальної іформації
        wsAdd47.Range(Cells(20, 1), Cells(20, 19)).Copy
        wsAdd47.Range(Cells(lastRow47, 1), Cells(lastRow47, 19)).PasteSpecial Paste:=xlPasteFormats
     
     
     Set foundCell = Nothing
        
        Call ProgressBar.Progress(n / UBound(filteredData) * 100)
    
    Next n
    
    'знаходимо останню колонку після всіх маніпуляцій
     Dim LastColumnFinal As Long
     LastColumnFinal = wsAdd47.Cells.Find(What:="*", After:=wsAdd47.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
     'копіюємо стилі основної інформації (для показників підрозділів)
       wsAdd47.Range(Cells(20, 12), Cells(20, 15)).Copy
       wsAdd47.Range(Cells(21, 12), Cells(lastRow47, LastColumnFinal)).PasteSpecial Paste:=xlPasteFormats
 
       'проставляємо суму
       Dim cellAll As Range
       Dim cellWithSum As Range
       For Each cellAll In wsAdd47.Rows(17).Cells 'рядок з усього
        If cellAll.value = "Усього" Then
            Set cellWithSum = cellAll.Offset(2, 0)
            cellWithSum.AutoFill destination:=wsAdd47.Range(wsAdd47.Cells(cellWithSum.Row, cellWithSum.Column), wsAdd47.Cells(lastRow47, cellWithSum.Column))
        End If
    Next cellAll
    
  
    
'   'формула по категоріях
'    For j = 9 To 11
'       For i = 21 To lastRow47
'           wsAdd47.Range(Cells(i, j), Cells(lastRow47, j)).Formula = "=SUMIFS(" & wsAdd47.Range(Cells(i, 12), Cells(i, LastColumnFinal)).Address & "," & wsAdd47.Range(Cells(18, 12), Cells(18, LastColumnFinal)).Address & "," & wsAdd47.Cells(18, j).Address & ")"
'
'       Next i
'    Next j
    
    Dim result As Double
    For j = 9 To 11
        For i = 21 To lastRow47
            
            Set sumRange = wsAdd47.Range(wsAdd47.Cells(i, 12), wsAdd47.Cells(i, LastColumnFinal))
            Set criteriaRange = wsAdd47.Range(wsAdd47.Cells(18, 12), wsAdd47.Cells(18, LastColumnFinal))
            criteriaValue = wsAdd47.Cells(18, j).value
        
'            wsAdd47.Range(Cells(i, j), Cells(lastRow47, j)).Formula = "=SUMIFS(" & wsAdd47.Range(Cells(i, 14), Cells(i, LastColumnFinal)).Address & "," & wsAdd47.Range(Cells(18, 14), Cells(18, LastColumnFinal)).Address & "," & wsAdd47.Cells(18, j).Address & ")"
        
        result = Application.WorksheetFunction.SumIfs(sumRange, criteriaRange, criteriaValue)
        wsAdd47.Cells(i, j).value = result
        Next i
    Next j
  
 'скидаємо область друку
wsAdd47.PageSetup.PrintArea = ""
'скидаємо фільтри
            If wsAllData.FilterMode Then
            On Error Resume Next
                wsAllData.ShowAllData
            End If
            
            Exit Sub
'***помилки****************************************
DontSelect:
MsgBox "Ймовірно,майно відсутнє в базі, або Ви не підтвердили вибір майна: вибране майно повинне ""перестрибнути"" з нижнього у верхнє поле вікна вибору майна, спробуйтие сформувати додаток ще раз!", vbCritical, "Майно не знайдено"
'wsAllData.ShowAllData
Exit Sub

WrongPrice:
MsgBox "Ймовірно,ви не вказали ціну вибраного майна, вкажіть ціну (якщо не відома, поставте 1 грн.) і спробуйте створити додаток ще раз!", vbCritical, "Ціна не вказана"
Exit Sub


End Sub


Sub SaveAdd47PNP()

        ActiveSheet.Select 'копіюємо
        ActiveSheet.Copy 'зберігаємо як нову книгу
     
     Dim WbNewWO As Workbook
     
        Set WbNewWO = ActiveWorkbook
    
    Dim nameAdd47 As String
    Dim nomenclature47 As String
    Dim unit47 As String
    Dim price47 As Double
    Dim formattedPrice As String
    
    nameAdd47 = WbNewWO.Worksheets(1).Range("nameAdd47PNP")
    
    nomenclature47 = WbNewWO.Worksheets(1).Range("nomenclature47PNP")
    unit47 = WbNewWO.Worksheets(1).Range("unit47PNP")
    price47 = WbNewWO.Worksheets(1).Range("price47PNP")
    
    formattedPrice = Format(price47, "#,##0.00")

    Dim shp As Shape
    For Each shp In WbNewWO.Worksheets(1).Shapes
        shp.Delete
    Next shp
        
      Dim fileName As String, fileName2 As String, lineBrake As String
        'Шлях до файла зберігання
        SavePath = WbNewWO.Worksheets(1).Range("PathToSaveAdd47PNP")
        Dim DateAdd47 As Date
        DateAdd47 = Format(Date, "dd.mm.yyyy")
        
        
        
              On Error Resume Next
         If Dir(SavePath, vbDirectory) = vbNullString Then
         On Error GoTo 0
        With Application.FileDialog(msoFileDialogFolderPicker)
            .Title = "Виберіть папку для збереження"
            .AllowMultiSelect = False
            If .show = -1 Then
                SavePath = .SelectedItems(1)
                'Записуємо новий шлях у клітинку
                ThisWorkbook.Sheets("Додаток 47 пр-непр").Range("PathToSaveAdd47PNP").value = SavePath
            Else
                MsgBox "Збереження скасовано", vbExclamation
                WbNew.Close SaveChanges:=False
                Application.DisplayAlerts = True
                Exit Sub
            End If
        End With
    End If
        
        
        
        
        'складаємо назву файла
        fileName = SavePath & Application.PathSeparator & CleanFileName(DateAdd47 & " д47 " & nameAdd47 & " " & nomenclature47 & " " & unit47 & " " & formattedPrice & ".xlsx")
        'заміняємо перенесення рядка, косі лінії
'        fileName2 = CleanFileName(fileName)
        
        'зберігаємо файл
        On Error GoTo WrongPath
         WbNewWO.SaveAs fileName:=fileName, FileFormat:=xlOpenXMLWorkbook
         On Error GoTo 0
     Exit Sub
WrongPath:
     MsgBox "Ймовірно, не правильно вказана адреса збереження файла, виправте адресу збереження файлу( на зразок ""D:\База обліку майна ЗСУ 2024"" *без косої лінії вкінці) і спробуйтие сформувати додаток ще раз!", vbCritical, "Не правильний шлях збереження"
     Exit Sub
     
End Sub





```

---

## 6. ReportingMove.bas — Звітна відомість ф. 18 (наказ МОУ 440)
### 6.1. Призначення
Офіційна звітна відомість за формою 18 наказу МОУ № 440. Агрегує рух за період у таблицю `tbReportingMove` із формуванням переліку всіх виправдних документів.

### 6.2. Вихідний код ReportingMove.bas
```vba
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



```
