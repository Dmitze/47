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

