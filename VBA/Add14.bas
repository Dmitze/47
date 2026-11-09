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

