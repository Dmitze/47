Attribute VB_Name = "Module1"
Option Explicit

' =====================================================================
'  CheckBase - перевірка бази даних (аркуш "База", таблиця AllData)
'
'  ЗАПУСК:  Alt+F8 -> CheckBaseErrors  (або призначте макрос на кнопку)
'    CheckBaseErrors  - перевіряє ВСЮ базу і створює аркуш "Помилки бази"
'                       (база НЕ змінюється: ні заливки, ні примітки)
'    MarkBaseErrors   - (за бажанням) зафарбовує проблемні клітинки в базі
'    UnmarkBaseErrors - знімає ТІЛЬКИ ці кольори
'
'  У звіті: клік по номеру рядка переходить до клітинки в базі;
'           колонки можна фільтрувати (Помилка / Увага).
'
'  Структура бази: 27 колонок A-AA, шапка в рядку 2, у рядку 1 - примітки "Обов'язкове!".
'  Помилка - те, що ламає формування документів (Додаток 47 тощо).
'  Увага   - підозріле, потребує перегляду людиною.
' =====================================================================

' ------------------------- НАЛАШТУВАННЯ ------------------------------
Private Const SHEET_DATA As String = "База"
Private Const TABLE_DATA As String = "AllData"
Private Const SHEET_REPORT As String = "Помилки бази"

Private Const CHECK_CATEGORY_AS_TEXT As Boolean = False  ' True = вимагати категорію саме текстом (у документі на аркуш База вона числова)
Private Const CHECK_ENTRY_DATE As Boolean = True         ' перша колонка = дата внесення
Private Const CHECK_DUPLICATES As Boolean = True         ' однакові рядки
Private Const CHECK_SPELLING As Boolean = True           ' різне написання одного й того ж
Private Const MIN_DOC_YEAR As Long = 2014                ' раніше цього року - підозріла дата
' ---------------------------------------------------------------------

Private Const LVL_ERR As String = "Помилка"
Private Const LVL_WARN As String = "Увага"
Private Const FIRST_REPORT_ROW As Long = 5

Private mRes() As Variant      ' (1 To 7, 1 To cap): рядок, рівень, колонка, значення, перевірка, опис
Private mCnt As Long
Private mCap As Long
Private mHdr() As String       ' заголовки колонок таблиці
Private mCats As Variant       ' допустимі категорії


' =====================================================================
'  ГОЛОВНИЙ МАКРОС
' =====================================================================
Public Sub CheckBaseErrors()

    Dim wsData As Worksheet
    Dim lo As ListObject
    Dim rng As Range
    Dim data As Variant
    Dim nRows As Long, nCols As Long
    Dim r As Long, c As Long, sr As Long, firstRow As Long
    Dim missing As String

    Dim cName As Long, cUnit As Long, cCode As Long, cCat As Long
    Dim cIn As Long, cOut As Long, cFrom As Long, cTo As Long
    Dim cDocDate As Long, cDocNum As Long, cPrice As Long
    Dim cDocType As Long, cOp As Long, cCatTo As Long
    Dim cF10 As Long, cF12 As Long
    Dim wrOff As String

    Dim kind As Long
    Dim hasIn As Boolean, hasOut As Boolean
    Dim numIn As Double, numOut As Double
    Dim fromN As String, toN As String
    Dim key As String
    Dim dupKeys As Object
    Dim dPlaces As Object, dNames As Object, dUnits As Object, dNameUnit As Object
    Dim nErr As Long, nWarn As Long, nBadRows As Long
    Dim i As Long
    Dim badRows As Object

    On Error GoTo EH
    Application.ScreenUpdating = False

    Set wsData = ThisWorkbook.Worksheets(SHEET_DATA)
    Set lo = wsData.ListObjects(TABLE_DATA)

    If lo.DataBodyRange Is Nothing Then
        MsgBox "У таблиці " & TABLE_DATA & " немає рядків з даними.", vbInformation, "Перевірка бази"
        GoTo Done
    End If

    ' ---- заголовки і пошук потрібних колонок ----
    nCols = lo.ListColumns.count
    ReDim mHdr(1 To nCols)
    For c = 1 To nCols
        mHdr(c) = CStr(lo.ListColumns(c).name)
    Next c

    cName = CB_FindCol("Найменування", missing)
    cUnit = CB_FindCol("Одиниця виміру", missing)
    cCode = CB_FindCol("Код номенклатури", missing)
    cCat = CB_FindCol("Категорія", missing)
    cIn = CB_FindCol("Надійшло", missing)
    cOut = CB_FindCol("Вибуло", missing)
    cFrom = CB_FindCol("Звідки", missing)
    cTo = CB_FindCol("Куди", missing)
    cDocDate = CB_FindCol("Дата документа", missing)
    cDocNum = CB_FindCol("№ документа", missing)
    cPrice = CB_FindCol("Ціна", missing)
    cDocType = CB_FindCol("Тип документа", missing)
    cOp = CB_FindCol("Запасне поле2", missing)      ' тип операції
    cCatTo = CB_FindCol("Запасне поле3", missing)   ' категорія "куди"
    cF10 = CB_FindColOpt("Поле 10")                 ' потрібне для донесень
    cF12 = CB_FindColOpt("Поле 12")                 ' субрахунок для інвентаризації

    If Len(missing) > 0 Then
        MsgBox "У таблиці " & TABLE_DATA & " не знайдено колонки:" & missing, vbExclamation, "Перевірка бази"
        GoTo Done
    End If

    ' ---- читаємо всю таблицю одним махом ----
    Set rng = lo.DataBodyRange
    nRows = rng.Rows.count
    firstRow = rng.Row
    If rng.Cells.count = 1 Then
        ReDim data(1 To 1, 1 To 1)
        data(1, 1) = rng.Value2
    Else
        data = rng.Value2
    End If

    ' ---- ініціалізація ----
    mCap = 2000
    mCnt = 0
    ReDim mRes(1 To 7, 1 To mCap)
    mCats = Array("1", "2", "3", "4", "5", "пр.", "не пр.", "пот. рем.")
    wrOff = CB_Norm("списан")   ' "Списане" / "Списання"

    Set dupKeys = CreateObject("Scripting.Dictionary")
    Set dPlaces = CreateObject("Scripting.Dictionary")
    Set dNames = CreateObject("Scripting.Dictionary")
    Set dUnits = CreateObject("Scripting.Dictionary")
    Set dNameUnit = CreateObject("Scripting.Dictionary")

    ' ---- перевірка рядок за рядком ----
    For r = 1 To nRows
        sr = firstRow + r - 1
        If r Mod 500 = 0 Then Application.StatusBar = "Перевірка бази: " & r & " з " & nRows

        ' текстові поля
        CB_CheckText sr, data(r, cName), cName, True, True
        CB_CheckText sr, data(r, cUnit), cUnit, True, True
        CB_CheckText sr, data(r, cCode), cCode, True, False
        CB_CheckText sr, data(r, cFrom), cFrom, True, True
        CB_CheckText sr, data(r, cTo), cTo, True, True
        CB_CheckText sr, data(r, cDocNum), cDocNum, False, False
        CB_CheckText sr, data(r, cDocType), cDocType, False, True

        ' категорії та тип операції
        CB_CheckCategory sr, data(r, cCat), cCat
        CB_CheckCategory sr, data(r, cCatTo), cCatTo
        kind = CB_CheckOp(sr, data(r, cOp), cOp)

        ' кількість
        CB_CheckQty sr, data(r, cIn), cIn, hasIn, numIn
        CB_CheckQty sr, data(r, cOut), cOut, hasOut, numOut

        If hasIn And hasOut Then
            CB_Add sr, LVL_ERR, cOut, CB_Show(data(r, cOut)), "Кількість", _
                   "Заповнені і ""Надійшло"", і ""Вибуло"". Кількість має бути лише в одній колонці: кожна операція - окремий рядок."
        ElseIf (Not hasIn) And (Not hasOut) Then
            CB_Add sr, LVL_ERR, cIn, "(порожньо)", "Кількість", _
                   "Не вказана кількість: порожні і ""Надійшло"", і ""Вибуло""."
        ElseIf kind = 1 And Not hasIn Then
            CB_Add sr, LVL_ERR, cOp, CB_Show(data(r, cOp)), "Тип і кількість", _
                   "Тип операції ""надходження"", але кількість стоїть у ""Вибуло"". Для надходження кількість береться з ""Надійшло""."
        ElseIf (kind = 2 Or kind = 3) And Not hasOut Then
            CB_Add sr, LVL_ERR, cOp, CB_Show(data(r, cOp)), "Тип і кількість", _
                   "Тип операції ""вибуття"" / ""внутрішнє переміщення"", але кількість стоїть у ""Надійшло"". Для цих типів кількість береться з ""Вибуло""."
        End If

        ' ціна, дати
        CB_CheckPrice sr, data(r, cPrice), cPrice
        CB_CheckDate sr, data(r, cDocDate), cDocDate, True, LVL_ERR
        If CHECK_ENTRY_DATE Then CB_CheckDate sr, data(r, 1), 1, False, LVL_WARN

        ' Звідки / Куди: переміщення в той самий підрозділ; вибуття = "Списане"
        fromN = CB_Norm(CB_Txt(data(r, cFrom)))
        toN = CB_Norm(CB_Txt(data(r, cTo)))
        If kind = 3 Then
            If Len(fromN) > 0 And fromN = toN Then
                CB_Add sr, LVL_ERR, cTo, CB_Show(data(r, cTo)), "Переміщення", _
                       "Внутрішнє переміщення в той самий підрозділ (""Звідки"" = ""Куди"")."
            End If
        End If
        If Len(toN) > 0 And kind > 0 Then
            If Left$(toN, Len(wrOff)) = wrOff Then
                If kind <> 2 Then
                    CB_Add sr, LVL_WARN, cTo, CB_Show(data(r, cTo)), "Списання", _
                           "У ""Куди"" вказано списання, але тип операції не ""вибуття""."
                End If
            ElseIf kind = 2 Then
                CB_Add sr, LVL_WARN, cTo, CB_Show(data(r, cTo)), "Вибуття", _
                       "Тип операції ""вибуття"", але в ""Куди"" не ""Списане"" (за схемою бази списання завжди йде в ""Списане"")."
            End If
        End If

        ' службові поля, обов'язкові за примітками в шапці бази
        If cF10 > 0 Then
            CB_CheckSoftFilled sr, data(r, cF10), cF10, _
                "Порожнє поле, а за приміткою в шапці воно обов'язкове для формування донесення."
        End If
        If cF12 > 0 Then
            CB_CheckSoftFilled sr, data(r, cF12), cF12, _
                "Порожній субрахунок, а за приміткою в шапці він обов'язковий для інвентаризаційного опису: рядок не потрапить в опис."
        End If

        ' дублі
        If CHECK_DUPLICATES Then
            key = CB_K(data(r, cDocDate)) & "|" & CB_K(data(r, cDocNum)) & "|" & CB_K(data(r, cDocType)) & "|" & _
                  CB_K(data(r, cName)) & "|" & CB_K(data(r, cCode)) & "|" & CB_K(data(r, cUnit)) & "|" & _
                  CB_K(data(r, cPrice)) & "|" & CB_K(data(r, cFrom)) & "|" & CB_K(data(r, cTo)) & "|" & _
                  CB_K(data(r, cIn)) & "|" & CB_K(data(r, cOut)) & "|" & CB_K(data(r, cCat)) & "|" & _
                  CB_K(data(r, cCatTo)) & "|" & CB_K(data(r, cOp))
            If dupKeys.exists(key) Then
                CB_Add sr, LVL_WARN, cName, CB_Show(data(r, cName)), "Дубль рядка", _
                       "Рядок повністю збігається з рядком " & dupKeys(key) & " (можливий подвійний імпорт)."
            Else
                dupKeys.Add key, sr
            End If
        End If

        ' збір даних для перевірки написання
        If CHECK_SPELLING Then
            CB_Collect dPlaces, data(r, cFrom), sr, cFrom
            CB_Collect dPlaces, data(r, cTo), sr, cTo
            CB_Collect dNames, data(r, cName), sr, cName
            CB_Collect dUnits, data(r, cUnit), sr, cUnit
            CB_CollectNameUnit dNameUnit, data(r, cName), data(r, cUnit), sr, cUnit
        End If
    Next r

    If CHECK_SPELLING Then
        CB_ReportSpelling dPlaces, "Різне написання підрозділу"
        CB_ReportSpelling dNames, "Різне написання найменування"
        CB_ReportSpelling dUnits, "Різне написання одиниці виміру"
        CB_ReportNameUnit dNameUnit
    End If

    ' ---- підрахунок і звіт ----
    Set badRows = CreateObject("Scripting.Dictionary")
    For i = 1 To mCnt
        If mRes(2, i) = LVL_ERR Then nErr = nErr + 1 Else nWarn = nWarn + 1
        If Not badRows.exists(mRes(1, i)) Then badRows.Add mRes(1, i), 1
    Next i
    nBadRows = badRows.count

    CB_WriteReport lo, nRows, nErr, nWarn, nBadRows

    Application.StatusBar = False
    Application.ScreenUpdating = True
    If mCnt = 0 Then
        MsgBox "Перевірено рядків: " & nRows & vbCrLf & "Помилок не знайдено.", vbInformation, "Перевірка бази"
    Else
        MsgBox "Перевірено рядків: " & nRows & vbCrLf & _
               "Рядків із проблемами: " & nBadRows & vbCrLf & _
               "Помилок: " & nErr & vbCrLf & _
               "Уваг: " & nWarn & vbCrLf & vbCrLf & _
               "Звіт на аркуші """ & SHEET_REPORT & """. Клік по номеру рядка відкриє клітинку в базі.", _
               vbInformation, "Перевірка бази"
    End If

Done:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Exit Sub

EH:
    MsgBox "Перевірку перервано помилкою:" & vbCrLf & Err.Number & " - " & Err.Description, vbCritical, "Перевірка бази"
    Resume Done
End Sub


' =====================================================================
'  ЗАФАРБУВАТИ / ЗНЯТИ ЗАФАРБОВУВАННЯ (за останнім звітом)
' =====================================================================
Public Sub MarkBaseErrors()
    Dim wr As Worksheet, wd As Worksheet
    Dim lastRow As Long, i As Long, pass As Long
    Dim arr As Variant

    On Error GoTo EH
    Set wr = ThisWorkbook.Worksheets(SHEET_REPORT)
    Set wd = ThisWorkbook.Worksheets(SHEET_DATA)

    lastRow = wr.Cells(wr.Rows.count, 8).End(xlUp).Row
    If lastRow < FIRST_REPORT_ROW Then
        MsgBox "У звіті немає записів. Спочатку запустіть CheckBaseErrors.", vbInformation, "Позначення помилок"
        Exit Sub
    End If

    If MsgBox("Зафарбувати проблемні клітинки в базі (червоний - помилка, жовтий - увага)?" & vbCrLf & vbCrLf & _
              "Існуюче заливання цих клітинок буде замінено. Робіть це одразу після перевірки: " & _
              "якщо рядки в базі змінювалися, спершу запустіть перевірку ще раз.", _
              vbYesNo + vbQuestion, "Позначення помилок") <> vbYes Then Exit Sub

    arr = wr.Range(wr.Cells(FIRST_REPORT_ROW, 1), wr.Cells(lastRow, 9)).value
    Application.ScreenUpdating = False

    For pass = 1 To 2          ' спершу "Увага", потім "Помилка" (помилка має пріоритет)
        For i = 1 To UBound(arr, 1)
            If pass = 1 And arr(i, 3) = LVL_WARN Then
                wd.Cells(arr(i, 8), arr(i, 9)).Interior.Color = CB_ColorWarn()
            ElseIf pass = 2 And arr(i, 3) = LVL_ERR Then
                wd.Cells(arr(i, 8), arr(i, 9)).Interior.Color = CB_ColorErr()
            End If
        Next i
    Next pass

    Application.ScreenUpdating = True
    wd.Activate
    Exit Sub
EH:
    Application.ScreenUpdating = True
    MsgBox "Помилка: " & Err.Description, vbCritical, "Позначення помилок"
End Sub

Public Sub UnmarkBaseErrors()
    Dim wr As Worksheet, wd As Worksheet
    Dim lastRow As Long, i As Long
    Dim arr As Variant
    Dim c As Range

    On Error GoTo EH
    Set wr = ThisWorkbook.Worksheets(SHEET_REPORT)
    Set wd = ThisWorkbook.Worksheets(SHEET_DATA)

    lastRow = wr.Cells(wr.Rows.count, 8).End(xlUp).Row
    If lastRow < FIRST_REPORT_ROW Then Exit Sub

    arr = wr.Range(wr.Cells(FIRST_REPORT_ROW, 1), wr.Cells(lastRow, 9)).value
    Application.ScreenUpdating = False

    For i = 1 To UBound(arr, 1)
        Set c = wd.Cells(arr(i, 8), arr(i, 9))
        If c.Interior.Color = CB_ColorErr() Or c.Interior.Color = CB_ColorWarn() Then
            c.Interior.ColorIndex = xlNone
        End If
    Next i

    Application.ScreenUpdating = True
    Exit Sub
EH:
    Application.ScreenUpdating = True
    MsgBox "Помилка: " & Err.Description, vbCritical, "Зняття позначок"
End Sub


' =====================================================================
'  ПЕРЕВІРКИ ОКРЕМИХ ПОЛІВ
' =====================================================================

Private Sub CB_CheckText(ByVal sr As Long, ByVal v As Variant, ByVal ci As Long, _
                         ByVal noWildcard As Boolean, ByVal checkScript As Boolean)
    Dim s As String
    Dim issues As String
    Dim w As String

    If IsError(v) Then
        CB_Add sr, LVL_ERR, ci, CB_Show(v), "Помилка Excel у клітинці", _
               "У клітинці значення помилки Excel (#Н/Д, #ЗНАЧ! тощо)."
        Exit Sub
    End If

    If CB_IsBlank(v) Then
        CB_Add sr, LVL_ERR, ci, "(порожньо)", "Порожнє поле", "Поле не повинно бути порожнім."
        Exit Sub
    End If

    s = CStr(v)

    If noWildcard Then
        If InStr(s, "*") > 0 Or InStr(s, "?") > 0 Or InStr(s, "~") > 0 Then
            CB_Add sr, LVL_ERR, ci, CB_Show(v), "Символи * ? ~", _
                   "Символи * ? ~ у пошуку і фільтрі працюють як підстановочні - таке значення не буде знайдено. Приберіть їх."
        End If
    End If

    If InStr(s, vbLf) > 0 Or InStr(s, vbCr) > 0 Then
        CB_Add sr, LVL_ERR, ci, CB_Show(v), "Перенос рядка", "У клітинці є перенос рядка - це не допускається."
    End If

    If InStr(s, vbTab) > 0 Then issues = issues & "символ табуляції; "
    If InStr(s, ChrW(160)) > 0 Then issues = issues & "нерозривний пробіл (невидимий); "
    If Left$(s, 1) = " " Or Right$(s, 1) = " " Then issues = issues & "пробіл на початку або в кінці; "
    If InStr(s, "  ") > 0 Then issues = issues & "подвійний пробіл; "
    If Len(issues) > 0 Then
        CB_Add sr, LVL_WARN, ci, CB_Show(v), "Пробіли", _
               "Невидимі символи, що ламають точний пошук: " & Left$(issues, Len(issues) - 2) & "."
    End If

    If checkScript Then
        w = CB_MixedScriptWord(s)
        If Len(w) > 0 Then
            CB_Add sr, LVL_WARN, ci, CB_Show(v), "Змішана абетка", _
                   "У слові """ & w & """ змішані латинські та кириличні літери (схожі літери з різних алфавітів)."
        End If
    End If
End Sub

Private Sub CB_CheckCategory(ByVal sr As Long, ByVal v As Variant, ByVal ci As Long)
    Dim raw As String

    If IsError(v) Then
        CB_Add sr, LVL_ERR, ci, CB_Show(v), "Помилка Excel у клітинці", _
               "У клітинці значення помилки Excel (#Н/Д, #ЗНАЧ! тощо)."
        Exit Sub
    End If

    If CB_IsBlank(v) Then
        CB_Add sr, LVL_ERR, ci, "(порожньо)", "Категорія", "Поле не повинно бути порожнім."
        Exit Sub
    End If

    raw = CStr(v)

    If CB_InList(raw, mCats) Then
        If CHECK_CATEGORY_AS_TEXT And VarType(v) <> vbString Then
            CB_Add sr, LVL_WARN, ci, CB_Show(v), "Категорія як число", _
                   "Категорія збережена як число, а не як текст (в базі має бути текстом - із зеленим трикутником у клітинці)."
        End If
    ElseIf CB_InList(CB_Strip(raw), mCats) Then
        CB_Add sr, LVL_ERR, ci, CB_Show(v), "Категорія", _
               "Зайві або невидимі пробіли в категорії - Додаток 47 її не знайде. Залиште лише саме значення."
    Else
        CB_Add sr, LVL_ERR, ci, CB_Show(v), "Категорія", _
               "Недопустиме значення. Допустимі: 1, 2, 3, 4, 5, пр., не пр., пот. рем."
    End If
End Sub

' повертає: 0 - невідомо, 1 - надходження, 2 - вибуття, 3 - внутрішнє переміщення
Private Function CB_CheckOp(ByVal sr As Long, ByVal v As Variant, ByVal ci As Long) As Long
    Dim s As String, t As String
    Dim k As Long

    If IsError(v) Then
        CB_Add sr, LVL_ERR, ci, CB_Show(v), "Помилка Excel у клітинці", _
               "У клітинці значення помилки Excel (#Н/Д, #ЗНАЧ! тощо)."
        CB_CheckOp = 0
        Exit Function
    End If

    If CB_IsBlank(v) Then
        CB_Add sr, LVL_ERR, ci, "(порожньо)", "Тип операції", "Поле не повинно бути порожнім."
        CB_CheckOp = 0
        Exit Function
    End If

    s = CStr(v)
    t = LCase$(CB_Strip(s))

    Select Case t
        Case "надходження": k = 1
        Case "вибуття": k = 2
        Case "внутрішнє переміщення": k = 3
        Case Else: k = 0
    End Select

    If k = 0 Then
        CB_Add sr, LVL_ERR, ci, CB_Show(v), "Тип операції", _
               "Допустимі значення: надходження, вибуття, внутрішнє переміщення."
    ElseIf StrComp(s, t, vbBinaryCompare) <> 0 Then
        CB_Add sr, LVL_ERR, ci, CB_Show(v), "Тип операції", _
               "Значення має бути точно """ & t & """ (малими літерами, без зайвих пробілів): Додаток 47 порівнює з урахуванням регістру."
    End If

    CB_CheckOp = k
End Function

Private Sub CB_CheckQty(ByVal sr As Long, ByVal v As Variant, ByVal ci As Long, _
                        ByRef isFilled As Boolean, ByRef num As Double)
    Dim isTxt As Boolean

    isFilled = False
    num = 0

    If IsError(v) Then
        isFilled = True
        CB_Add sr, LVL_ERR, ci, CB_Show(v), "Помилка Excel у клітинці", _
               "У клітинці значення помилки Excel (#Н/Д, #ЗНАЧ! тощо)."
        Exit Sub
    End If

    If CB_IsBlank(v) Then Exit Sub
    isFilled = True

    If Not CB_ParseNum(v, num, isTxt) Then
        CB_Add sr, LVL_ERR, ci, CB_Show(v), "Кількість", "Значення не є числом."
    ElseIf num <= 0 Then
        CB_Add sr, LVL_ERR, ci, CB_Show(v), "Кількість", "Кількість має бути більшою за нуль."
    ElseIf isTxt Then
        CB_Add sr, LVL_WARN, ci, CB_Show(v), "Кількість", "Число збережене як текст."
    End If
End Sub

Private Sub CB_CheckPrice(ByVal sr As Long, ByVal v As Variant, ByVal ci As Long)
    Dim num As Double
    Dim isTxt As Boolean

    If IsError(v) Then
        CB_Add sr, LVL_ERR, ci, CB_Show(v), "Помилка Excel у клітинці", _
               "У клітинці значення помилки Excel (#Н/Д, #ЗНАЧ! тощо)."
        Exit Sub
    End If

    If CB_IsBlank(v) Then
        CB_Add sr, LVL_ERR, ci, "(порожньо)", "Ціна", "Не вказана ціна (якщо ціна невідома - поставте 1 грн.)."
        Exit Sub
    End If

    If Not CB_ParseNum(v, num, isTxt) Then
        CB_Add sr, LVL_ERR, ci, CB_Show(v), "Ціна", "Значення не є числом."
    ElseIf num <= 0 Then
        CB_Add sr, LVL_ERR, ci, CB_Show(v), "Ціна", "Ціна має бути більшою за нуль (якщо невідома - поставте 1 грн.)."
    Else
        If isTxt Then
            CB_Add sr, LVL_WARN, ci, CB_Show(v), "Ціна", _
                   "Число збережене як текст - фільтр і підсумки можуть його не бачити."
        End If
        If Abs(num - Round(num, 4)) > 0.0000001 Then
            CB_Add sr, LVL_WARN, ci, CB_Show(v), "Ціна", _
                   "Більше 4 знаків після коми - у Додатку 47 ціна може не збігтися при пошуку."
        End If
    End If
End Sub

Private Sub CB_CheckDate(ByVal sr As Long, ByVal v As Variant, ByVal ci As Long, _
                         ByVal isRequired As Boolean, ByVal softLevel As String)
    Dim d As Date
    Dim s As String

    If IsError(v) Then
        CB_Add sr, LVL_ERR, ci, CB_Show(v), "Помилка Excel у клітинці", _
               "У клітинці значення помилки Excel (#Н/Д, #ЗНАЧ! тощо)."
        Exit Sub
    End If

    If CB_IsBlank(v) Then
        If isRequired Then CB_Add sr, LVL_ERR, ci, "(порожньо)", "Дата", "Поле не повинно бути порожнім."
        Exit Sub
    End If

    Select Case VarType(v)
        Case vbDouble
            If v < 1 Or v > 2958465 Then
                CB_Add sr, softLevel, ci, CB_Show(v), "Дата", "Значення не є датою."
                Exit Sub
            End If
            d = CDate(v)
        Case vbString
            s = CB_Strip(CStr(v))
            If IsDate(s) Then
                d = CDate(s)
                CB_Add sr, LVL_WARN, ci, CB_Show(v), "Дата", "Дата збережена як текст, а не як справжня дата Excel."
            Else
                CB_Add sr, softLevel, ci, CB_Show(v), "Дата", "Значення не є датою."
                Exit Sub
            End If
        Case Else
            CB_Add sr, softLevel, ci, CB_Show(v), "Дата", "Значення не є датою."
            Exit Sub
    End Select

    If Year(d) < MIN_DOC_YEAR Then
        CB_Add sr, LVL_WARN, ci, Format$(d, "dd.mm.yyyy"), "Дата", _
               "Дуже стара дата (раніше " & MIN_DOC_YEAR & " року) - можлива друкарська помилка."
    ElseIf d > Date Then
        CB_Add sr, LVL_WARN, ci, Format$(d, "dd.mm.yyyy"), "Дата", "Дата в майбутньому."
    End If
End Sub


' =====================================================================
'  ПЕРЕВІРКА НАПИСАННЯ (різні варіанти одного й того ж значення)
' =====================================================================

Private Sub CB_Collect(ByVal dic As Object, ByVal v As Variant, ByVal sr As Long, ByVal ci As Long)
    Dim raw As String, nk As String
    Dim inner As Object
    Dim info As Variant

    If IsError(v) Then Exit Sub
    If CB_IsBlank(v) Then Exit Sub
    raw = CStr(v)
    nk = CB_Norm(raw)
    If Len(nk) = 0 Then Exit Sub

    If dic.exists(nk) Then
        Set inner = dic(nk)
    Else
        Set inner = CreateObject("Scripting.Dictionary")
        dic.Add nk, inner
    End If

    If inner.exists(raw) Then
        info = inner(raw)
        info(0) = info(0) + 1
        inner(raw) = info
    Else
        inner.Add raw, Array(1, sr, ci)
    End If
End Sub

Private Sub CB_ReportSpelling(ByVal dic As Object, ByVal checkName As String)
    Dim k As Variant, rk As Variant
    Dim inner As Object
    Dim info As Variant
    Dim mainRaw As String
    Dim mainCnt As Long

    For Each k In dic.Keys
        Set inner = dic(k)
        If inner.count > 1 Then
            mainRaw = ""
            mainCnt = -1
            For Each rk In inner.Keys
                info = inner(rk)
                If info(0) > mainCnt Then
                    mainCnt = info(0)
                    mainRaw = CStr(rk)
                End If
            Next rk
            For Each rk In inner.Keys
                If CStr(rk) <> mainRaw Then
                    info = inner(rk)
                    CB_Add CLng(info(1)), LVL_WARN, CLng(info(2)), CB_Show(CStr(rk)), checkName, _
                           "Схоже написання: основне """ & mainRaw & """ (" & mainCnt & " рядків), це - """ & _
                           CStr(rk) & """ (" & info(0) & " рядків). Перевірте, чи це не помилка написання " & _
                           "(регістр, пробіли, крапки, схожі літери)."
                End If
            Next rk
        End If
    Next k
End Sub

Private Sub CB_CollectNameUnit(ByVal dic As Object, ByVal vName As Variant, ByVal vUnit As Variant, _
                               ByVal sr As Long, ByVal ci As Long)
    Dim nk As String, uk As String
    Dim inner As Object
    Dim info As Variant

    If IsError(vName) Or IsError(vUnit) Then Exit Sub
    If CB_IsBlank(vName) Or CB_IsBlank(vUnit) Then Exit Sub
    nk = CB_Norm(CStr(vName))
    uk = CB_Norm(CStr(vUnit))
    If Len(nk) = 0 Or Len(uk) = 0 Then Exit Sub

    If dic.exists(nk) Then
        Set inner = dic(nk)
    Else
        Set inner = CreateObject("Scripting.Dictionary")
        dic.Add nk, inner
    End If

    If inner.exists(uk) Then
        info = inner(uk)
        info(0) = info(0) + 1
        inner(uk) = info
    Else
        inner.Add uk, Array(1, sr, ci, CStr(vUnit), CStr(vName))
    End If
End Sub

Private Sub CB_ReportNameUnit(ByVal dic As Object)
    Dim k As Variant, uk As Variant
    Dim inner As Object
    Dim info As Variant
    Dim mainUnit As String, mainKey As String
    Dim mainCnt As Long

    For Each k In dic.Keys
        Set inner = dic(k)
        If inner.count > 1 Then
            mainCnt = -1
            For Each uk In inner.Keys
                info = inner(uk)
                If info(0) > mainCnt Then
                    mainCnt = info(0)
                    mainUnit = CStr(info(3))
                    mainKey = CStr(uk)
                End If
            Next uk
            For Each uk In inner.Keys
                If CStr(uk) <> mainKey Then
                    info = inner(uk)
                    CB_Add CLng(info(1)), LVL_WARN, CLng(info(2)), CB_Show(CStr(info(3))), "Різні одиниці виміру", _
                           "Для найменування """ & CStr(info(4)) & """ основна одиниця виміру """ & mainUnit & _
                           """ (" & mainCnt & " рядків), а тут """ & CStr(info(3)) & """ (" & info(0) & " рядків)."
                End If
            Next uk
        End If
    Next k
End Sub


' =====================================================================
'  ЗВІТ
' =====================================================================

Private Sub CB_WriteReport(ByVal lo As ListObject, ByVal nRows As Long, _
                           ByVal nErr As Long, ByVal nWarn As Long, ByVal nBadRows As Long)
    Dim wr As Worksheet
    Dim i As Long
    Dim baseCol As Long
    Dim outP() As Variant
    Dim outF() As Variant
    Dim sumDic As Object
    Dim k As Variant
    Dim sumArr() As Variant
    Dim j As Long
    Dim lastRow As Long
    Dim fc As Object

    baseCol = lo.DataBodyRange.Column

    On Error Resume Next
    Set wr = ThisWorkbook.Worksheets(SHEET_REPORT)
    On Error GoTo 0

    If wr Is Nothing Then
        Set wr = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        wr.name = SHEET_REPORT
    Else
        If wr.AutoFilterMode Then wr.AutoFilterMode = False
        wr.Cells.EntireColumn.Hidden = False
        wr.Cells.Clear
    End If

    wr.Range("A1").value = "Перевірка бази даних (аркуш """ & SHEET_DATA & """, таблиця " & TABLE_DATA & ")"
    wr.Range("A1").Font.Bold = True
    wr.Range("A1").Font.Size = 14
    wr.Range("A2").value = "Перевірено рядків: " & nRows & "   |   Рядків із проблемами: " & nBadRows & _
                           "   |   Помилок: " & nErr & "   |   Уваг: " & nWarn & "   |   " & Format$(Now, "dd.mm.yyyy hh:nn")

    wr.Range("A4:I4").value = Array("№", "Рядок", "Рівень", "Колонка", "Значення", "Перевірка", "Що не так", "рядок (тех.)", "колонка (тех.)")
    wr.Range("A4:I4").Font.Bold = True
    wr.Range("A4:I4").Interior.Color = RGB(68, 84, 106)
    wr.Range("A4:I4").Font.Color = RGB(255, 255, 255)

    If mCnt = 0 Then
        wr.Range("A5").value = "Помилок не знайдено."
        wr.Range("A5").Font.Bold = True
        wr.Columns("A").ColumnWidth = 24
        Exit Sub
    End If

    ReDim outP(1 To mCnt, 1 To 9)
    ReDim outF(1 To mCnt, 1 To 1)

    For i = 1 To mCnt
        outP(i, 1) = i
        outP(i, 3) = mRes(2, i)
        outP(i, 4) = mHdr(mRes(3, i))
        outP(i, 5) = mRes(4, i)
        outP(i, 6) = mRes(5, i)
        outP(i, 7) = mRes(6, i)
        outP(i, 8) = mRes(1, i)
        outP(i, 9) = baseCol + mRes(3, i) - 1
        outF(i, 1) = "=HYPERLINK(""#'" & SHEET_DATA & "'!" & CB_ColLetter(baseCol + mRes(3, i) - 1) & mRes(1, i) & """," & mRes(1, i) & ")"
    Next i

    lastRow = FIRST_REPORT_ROW + mCnt - 1

    wr.Range(wr.Cells(FIRST_REPORT_ROW, 5), wr.Cells(lastRow, 5)).NumberFormat = "@"
    wr.Cells(FIRST_REPORT_ROW, 1).Resize(mCnt, 9).value = outP
    wr.Cells(FIRST_REPORT_ROW, 2).Resize(mCnt, 1).Formula = outF

    ' зведення за видами перевірок
    Set sumDic = CreateObject("Scripting.Dictionary")
    For i = 1 To mCnt
        If sumDic.exists(mRes(5, i)) Then
            sumDic(mRes(5, i)) = sumDic(mRes(5, i)) + 1
        Else
            sumDic.Add mRes(5, i), 1
        End If
    Next i

    wr.Range("K4:L4").value = Array("Зведення за видами", "Кількість")
    wr.Range("K4:L4").Font.Bold = True
    wr.Range("K4:L4").Interior.Color = RGB(68, 84, 106)
    wr.Range("K4:L4").Font.Color = RGB(255, 255, 255)

    ReDim sumArr(1 To sumDic.count, 1 To 2)
    j = 0
    For Each k In sumDic.Keys
        j = j + 1
        sumArr(j, 1) = CStr(k)
        sumArr(j, 2) = sumDic(k)
    Next k
    wr.Cells(5, 11).Resize(sumDic.count, 2).value = sumArr

    ' оформлення
    wr.Columns("A").ColumnWidth = 6
    wr.Columns("B").ColumnWidth = 9
    wr.Columns("C").ColumnWidth = 11
    wr.Columns("D").ColumnWidth = 22
    wr.Columns("E").ColumnWidth = 38
    wr.Columns("F").ColumnWidth = 28
    wr.Columns("G").ColumnWidth = 90
    wr.Columns("K").ColumnWidth = 34
    wr.Columns("L").ColumnWidth = 10
    wr.Range(wr.Cells(FIRST_REPORT_ROW, 7), wr.Cells(lastRow, 7)).WrapText = True
    wr.Range(wr.Cells(FIRST_REPORT_ROW, 1), wr.Cells(lastRow, 9)).VerticalAlignment = xlTop
    wr.Columns("H:I").Hidden = True

    Set fc = wr.Range(wr.Cells(FIRST_REPORT_ROW, 3), wr.Cells(lastRow, 3))
    fc.FormatConditions.Delete
    With fc.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""" & LVL_ERR & """")
        .Interior.Color = CB_ColorErr()
    End With
    With fc.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""" & LVL_WARN & """")
        .Interior.Color = CB_ColorWarn()
    End With

    wr.Range(wr.Cells(4, 1), wr.Cells(lastRow, 7)).AutoFilter

    wr.Activate
    ActiveWindow.FreezePanes = False
    wr.Range("A5").Select
    ActiveWindow.FreezePanes = True
End Sub


' =====================================================================
'  ДОПОМІЖНІ ФУНКЦІЇ
' =====================================================================

Private Sub CB_Add(ByVal sr As Long, ByVal level As String, ByVal ci As Long, _
                   ByVal shown As String, ByVal check As String, ByVal msg As String)
    If mCnt >= mCap Then
        mCap = mCap * 2
        ReDim Preserve mRes(1 To 7, 1 To mCap)
    End If
    mCnt = mCnt + 1
    mRes(1, mCnt) = sr
    mRes(2, mCnt) = level
    mRes(3, mCnt) = ci
    mRes(4, mCnt) = shown
    mRes(5, mCnt) = check
    mRes(6, mCnt) = msg
End Sub

Private Function CB_FindCol(ByVal nm As String, ByRef missing As String) As Long
    Dim c As Long
    For c = 1 To UBound(mHdr)
        If StrComp(CB_Strip(mHdr(c)), nm, vbTextCompare) = 0 Then
            CB_FindCol = c
            Exit Function
        End If
    Next c
    missing = missing & vbCrLf & " - " & nm
End Function

' необов'язкова колонка: 0, якщо такої немає
Private Function CB_FindColOpt(ByVal nm As String) As Long
    Dim c As Long
    For c = 1 To UBound(mHdr)
        If StrComp(CB_Strip(mHdr(c)), nm, vbTextCompare) = 0 Then
            CB_FindColOpt = c
            Exit Function
        End If
    Next c
End Function

Private Sub CB_CheckSoftFilled(ByVal sr As Long, ByVal v As Variant, ByVal ci As Long, ByVal msg As String)
    If IsError(v) Then Exit Sub
    If CB_IsBlank(v) Then CB_Add sr, LVL_WARN, ci, "(порожньо)", "Службове поле", msg
End Sub

' текст клітинки або "" якщо там помилка
Private Function CB_Txt(ByVal v As Variant) As String
    If IsError(v) Then Exit Function
    CB_Txt = CStr(v)
End Function

' прибирає пробіли, нерозривні пробіли, табуляції та переноси по краях
Private Function CB_Strip(ByVal s As String) As String
    Dim t As String
    t = Replace(s, ChrW(160), " ")
    t = Replace(t, vbTab, " ")
    t = Replace(t, vbCr, " ")
    t = Replace(t, vbLf, " ")
    CB_Strip = Trim$(t)
End Function

Private Function CB_IsBlank(ByVal v As Variant) As Boolean
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then
        CB_IsBlank = True
        Exit Function
    End If
    CB_IsBlank = (Len(CB_Strip(CStr(v))) = 0)
End Function

' значення клітинки для звіту (текст - у квадратних дужках, щоб було видно пробіли)
Private Function CB_Show(ByVal v As Variant) As String
    Dim s As String
    If IsError(v) Then CB_Show = "(помилка Excel)": Exit Function
    If IsEmpty(v) Then CB_Show = "(порожньо)": Exit Function
    s = CStr(v)
    If VarType(v) = vbString Then
        s = Replace(s, vbCrLf, ChrW(182))
        s = Replace(s, vbLf, ChrW(182))
        s = Replace(s, vbCr, ChrW(182))
        s = Replace(s, vbTab, "[tab]")
        s = Replace(s, ChrW(160), "[nbsp]")
        If Len(s) > 120 Then s = Left$(s, 120) & "..."
        s = "[" & s & "]"
    End If
    CB_Show = s
End Function

' ключ для порівняння рядків (дублі)
Private Function CB_K(ByVal v As Variant) As String
    If IsError(v) Then CB_K = "#err": Exit Function
    CB_K = LCase$(CB_Strip(CStr(v)))
End Function

Private Function CB_InList(ByVal s As String, ByVal lst As Variant) As Boolean
    Dim i As Long
    For i = LBound(lst) To UBound(lst)
        If StrComp(s, CStr(lst(i)), vbTextCompare) = 0 Then
            CB_InList = True
            Exit Function
        End If
    Next i
End Function

' Розпізнає число. isText = True, якщо число лежить у клітинці як текст.
Private Function CB_ParseNum(ByVal v As Variant, ByRef num As Double, ByRef isText As Boolean) As Boolean
    Dim s As String
    Dim ch As String
    Dim i As Long
    Dim decCnt As Long

    isText = False
    num = 0
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function

    Select Case VarType(v)
        Case vbDouble, vbSingle, vbInteger, vbLong, vbCurrency, vbDecimal, vbByte
            num = CDbl(v)
            CB_ParseNum = True
            Exit Function
        Case vbString
            ' розбираємо нижче
        Case Else
            Exit Function
    End Select

    isText = True
    s = CB_Strip(CStr(v))
    s = Replace(s, " ", "")
    If Len(s) = 0 Then Exit Function

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        Select Case ch
            Case "0" To "9"
            Case ",", "."
                decCnt = decCnt + 1
                If decCnt > 1 Then Exit Function
            Case "-"
                If i > 1 Then Exit Function
            Case Else
                Exit Function
        End Select
    Next i

    num = val(Replace(s, ",", "."))
    CB_ParseNum = True
End Function

' Нормалізація для порівняння: великі літери, лише літери й цифри,
' латинські літери, схожі на кириличні, замінюються кириличними.
Private Function CB_Norm(ByVal s As String) As String
    Dim i As Long
    Dim ch As Long
    Dim res As String

    s = UCase$(s)
    For i = 1 To Len(s)
        ch = AscW(Mid$(s, i, 1))
        If ch < 0 Then ch = ch + 65536
        Select Case ch
            Case 48 To 57, 1024 To 1279
                res = res & ChrW(ch)
            Case 65 To 90
                res = res & ChrW(CB_LatinToCyr(ch))
        End Select
    Next i
    CB_Norm = res
End Function

Private Function CB_LatinToCyr(ByVal ch As Long) As Long
    Select Case ch
        Case 65: CB_LatinToCyr = 1040   ' A -> кирилична А
        Case 66: CB_LatinToCyr = 1042   ' B -> В
        Case 67: CB_LatinToCyr = 1057   ' C -> С
        Case 69: CB_LatinToCyr = 1045   ' E -> Е
        Case 72: CB_LatinToCyr = 1053   ' H -> Н
        Case 73: CB_LatinToCyr = 1030   ' I -> І
        Case 75: CB_LatinToCyr = 1050   ' K -> К
        Case 77: CB_LatinToCyr = 1052   ' M -> М
        Case 79: CB_LatinToCyr = 1054   ' O -> О
        Case 80: CB_LatinToCyr = 1056   ' P -> Р
        Case 84: CB_LatinToCyr = 1058   ' T -> Т
        Case 88: CB_LatinToCyr = 1061   ' X -> Х
        Case 89: CB_LatinToCyr = 1059   ' Y -> У
        Case Else: CB_LatinToCyr = ch
    End Select
End Function

' Повертає перше слово, де змішані латинські й кириличні літери (або "")
Private Function CB_MixedScriptWord(ByVal s As String) As String
    Dim i As Long
    Dim ch As Long
    Dim hasLat As Boolean, hasCyr As Boolean
    Dim word As String

    For i = 1 To Len(s) + 1
        If i <= Len(s) Then
            ch = AscW(Mid$(s, i, 1))
            If ch < 0 Then ch = ch + 65536
        Else
            ch = 32
        End If

        If (ch >= 65 And ch <= 90) Or (ch >= 97 And ch <= 122) Then
            hasLat = True
            word = word & ChrW(ch)
        ElseIf ch >= 1024 And ch <= 1279 Then
            hasCyr = True
            word = word & ChrW(ch)
        ElseIf ch >= 48 And ch <= 57 Then
            word = word & ChrW(ch)
        Else
            If hasLat And hasCyr Then
                CB_MixedScriptWord = word
                Exit Function
            End If
            hasLat = False
            hasCyr = False
            word = ""
        End If
    Next i
End Function

Private Function CB_ColLetter(ByVal c As Long) As String
    Dim s As String
    Do While c > 0
        s = Chr$(65 + ((c - 1) Mod 26)) & s
        c = (c - 1) \ 26
    Loop
    CB_ColLetter = s
End Function

Private Function CB_ColorErr() As Long
    CB_ColorErr = RGB(255, 199, 206)
End Function

Private Function CB_ColorWarn() As Long
    CB_ColorWarn = RGB(255, 235, 156)
End Function
