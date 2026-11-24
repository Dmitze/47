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
' Основна процедура: Залишки по підрозділах 2.0
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
    Dim cOpType As Long

    ' змінні для рядка
    Dim fromVal As String, toVal As String, cat As String
    Dim prefix As String, suffix As String, key As String
    Dim vIn As Double, vOut As Double, v As Double
    Dim opType As String, movQty As Double

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

    '--- Індекси колонок ---
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
        cOpType = .ListColumns("Запасне поле2").Index
    End With

    ' Читаємо всю базу в масив
    data = tbAllData.DataBodyRange.value

    Set dictUnits = CreateObject("Scripting.Dictionary")

    '--- Один прохід по базі з урахуванням типу операції ---
    For i = 1 To dataCount
        fromVal = SafeStr(data(i, cFrom))
        toVal = SafeStr(data(i, cTo))
        opType = LCase(Trim(SafeStr(data(i, cOpType))))
        vIn = SafeNum(data(i, cIn))
        vOut = SafeNum(data(i, cOut))

        prefix = SafeStr(data(i, cCode)) & "||" & _
                 SafeStr(data(i, cName)) & "||" & _
                 SafeStr(data(i, cPrice)) & "||" & _
                 SafeStr(data(i, cUnit)) & "||"
        suffix = "||" & SafeStr(data(i, cNote)) & "||" & _
                 SafeStr(data(i, c11)) & "||" & _
                 SafeStr(data(i, c12))

        Select Case opType
            Case "надходження"
                ' Впливає ТІЛЬКИ на отримувача (Куди)
                If toVal <> "" And vIn <> 0 Then
                    cat = SafeStr(data(i, cCat3))
                    If Len(cat) = 0 Then cat = SafeStr(data(i, cCat))
                    key = prefix & cat & suffix
                    If Not dictUnits.exists(toVal) Then dictUnits.Add toVal, CreateObject("Scripting.Dictionary")
                    Set d = dictUnits(toVal)
                    If Not d.exists(key) Then d.Add key, 0#
                    d(key) = d(key) + vIn
                End If

            Case "вибуття"
                ' Впливає ТІЛЬКИ на відправника (Звідки)
                If fromVal <> "" And vOut <> 0 Then
                    cat = SafeStr(data(i, cCat))
                    key = prefix & cat & suffix
                    If Not dictUnits.exists(fromVal) Then dictUnits.Add fromVal, CreateObject("Scripting.Dictionary")
                    Set d = dictUnits(fromVal)
                    If Not d.exists(key) Then d.Add key, 0#
                    d(key) = d(key) - vOut
                End If

            Case "внутрішнє переміщення"
                ' Кількість береться з Вибуло (або Надійшло як резерв)
                movQty = vOut
                If movQty = 0 Then movQty = vIn

                If movQty <> 0 Then
                    ' Додаємо отримувачу (Куди)
                    If toVal <> "" Then
                        cat = SafeStr(data(i, cCat3))
                        If Len(cat) = 0 Then cat = SafeStr(data(i, cCat))
                        key = prefix & cat & suffix
                        If Not dictUnits.exists(toVal) Then dictUnits.Add toVal, CreateObject("Scripting.Dictionary")
                        Set d = dictUnits(toVal)
                        If Not d.exists(key) Then d.Add key, 0#
                        d(key) = d(key) + movQty
                    End If

                    ' Віднімаємо у відправника (Звідки)
                    If fromVal <> "" Then
                        cat = SafeStr(data(i, cCat))
                        key = prefix & cat & suffix
                        If Not dictUnits.exists(fromVal) Then dictUnits.Add fromVal, CreateObject("Scripting.Dictionary")
                        Set d = dictUnits(fromVal)
                        If Not d.exists(key) Then d.Add key, 0#
                        d(key) = d(key) - movQty
                    End If
                End If
        End Select

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
