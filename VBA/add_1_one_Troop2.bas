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



