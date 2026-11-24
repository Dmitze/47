Attribute VB_Name = "add_1_one_Troop"
Option Explicit

' =====================================================================
'  Додаток 1 (по підрозділу): рух майна з аркуша "База"
'
'  Колонки таблиці tbAddOne:
'    B  Найменування, C  Одиниця виміру
'    D  Надійшло  = сума кількостей зі статусом "Видано"
'    E  Вибуло    = сума кількостей зі статусом "Повернено" + "Знищено"
'    F  Наявне    = D - E
'    G  Примітка  = документи (тип, номер, дата) по статусах:
'                   спочатку Видано, потім Повернено, потім Знищено
'
'  Статус визначається за схемою бази:
'    Видано    - майно надійшло на підрозділ (внутрішнє переміщення зі складу)
'    Повернено - майно залишило підрозділ і повернулось на склад
'    Знищено   - вибуття з підрозділу (списання / знищення)
'  Групування лише за найменуванням та одиницею виміру (ціна не враховується).
'  Враховується весь рух до вказаної дати включно.
' =====================================================================

' Документи цього типу вказуються в примітці БЕЗ дати (наприклад: накази № 1/БН/дск)
Private Const NO_DATE_DOC_WORD As String = "наказ"
' True = у "Знищено" рядок із серійним номером у "Примітці" виводиться як "втрачено с/н ... від ..."
Private Const LOST_SERIAL_IN_NOTE As Boolean = True
Private Const QTY_LABEL As String = "од."

Sub CreateAddOneForOneTroop()

    Dim wbAllData As Workbook
    Dim wsAllData As Worksheet
    Dim wsReportingMove As Worksheet
    Dim wsTroops As Worksheet
    Dim tbAllData As ListObject
    Dim tbReportingMove As ListObject
    Dim frm As UserForm1
    Dim dictUnits As Object
    Dim dict As Object
    Dim cell As Range
    Dim rngData As Range
    Dim unitKey As Variant
    Dim itemKey As Variant

    Dim k As Long, i As Long
    Dim lastRowT As Long
    Dim selectedUnit As String
    Dim endText As String
    Dim periodDefault As String
    Dim endDate As Date
    Dim data As Variant
    Dim nRows As Long

    Dim cDateDoc As Long, cFrom As Long, cTo As Long
    Dim cName As Long, cUnit As Long
    Dim cIn As Long, cOut As Long, cOpType As Long
    Dim cDocType As Long, cDocNum As Long, cNote As Long

    Dim dVal As Double, minDate As Double
    Dim unitFrom As String, unitTo As String, opType As String
    Dim st As Long
    Dim qty As Double
    Dim itemName As String, itemUnit As String
    Dim key As String
    Dim info As Variant
    Dim docs As Object
    Dim docType As String, docNum As String, serial As String
    Dim docKey As String

    Dim counter As Long, skippedNoDate As Long, negCount As Long
    Dim dIn As Double, dOut As Double, dBal As Double
    Dim msg As String

    On Error GoTo EH

    Set wbAllData = ThisWorkbook
    Set wsAllData = wbAllData.Worksheets("База")
    Set tbAllData = wsAllData.ListObjects("AllData")
    Set wsReportingMove = wbAllData.Worksheets("Додаток 1")
    Set tbReportingMove = wsReportingMove.ListObjects("tbAddOne")

    Set rngData = tbAllData.DataBodyRange
    If rngData Is Nothing Then
        MsgBox "У таблиці AllData немає даних.", vbInformation, "Додаток 1"
        Exit Sub
    End If

    ' ---------------- вибір підрозділу ----------------
    Set frm = New UserForm1
    Set dictUnits = CreateObject("Scripting.Dictionary")

    For Each cell In tbAllData.ListColumns("Звідки").DataBodyRange
        If Len(Trim$(A1Txt(cell.value))) > 0 Then
            If Not dictUnits.exists(cell.value) Then dictUnits.Add cell.value, Nothing
        End If
    Next cell
    For Each cell In tbAllData.ListColumns("Куди").DataBodyRange
        If Len(Trim$(A1Txt(cell.value))) > 0 Then
            If Not dictUnits.exists(cell.value) Then dictUnits.Add cell.value, Nothing
        End If
    Next cell

    Set wsTroops = wbAllData.Worksheets("Підрозділи")
    wsTroops.Range("A1:A" & wsTroops.Rows.count).ClearContents
    k = 1
    For Each unitKey In dictUnits.Keys
        wsTroops.Cells(k, 1).value = unitKey
        k = k + 1
    Next unitKey
    lastRowT = wsTroops.Cells(wsTroops.Rows.count, 1).End(xlUp).Row

    frm.Caption = "Виберіть підрозділ"
    frm.ListData = wsTroops.Range("A1:A" & lastRowT)
    frm.show

    If frm.ListBox1.ListIndex <> -1 Then
        selectedUnit = Trim$(CStr(frm.ListBox1.value))
    Else
        MsgBox "Підрозділ не вибрано!", vbExclamation, "Додаток 1"
        Exit Sub
    End If

    wsReportingMove.Range("AddOneNameTroop") = selectedUnit

    ' ---------------- дата, станом на яку формуємо ----------------
    periodDefault = ""
    If IsDate(wsReportingMove.Range("AddOnePeriodTo").value) Then
        periodDefault = Format$(CDate(wsReportingMove.Range("AddOnePeriodTo").value), "dd.mm.yyyy")
    End If
    endText = InputBox("Введіть дату у форматі дд.мм.рррр, станом на яку формувати Додаток 1." & vbCrLf & _
                       "Враховується весь рух майна підрозділу до цієї дати включно.", "Дата ПО", periodDefault)
    If Len(Trim$(endText)) = 0 Then Exit Sub
    dVal = A1ParseDate(endText)
    If dVal = 0 Then
        MsgBox "Некоректна дата: " & endText, vbExclamation, "Додаток 1"
        Exit Sub
    End If
    endDate = CDate(dVal)

    ' ---------------- колонки бази ----------------
    With tbAllData
        cDateDoc = .ListColumns("Дата документа").Index
        cFrom = .ListColumns("Звідки").Index
        cTo = .ListColumns("Куди").Index
        cName = .ListColumns("Найменування").Index
        cUnit = .ListColumns("Одиниця виміру").Index
        cIn = .ListColumns("Надійшло").Index
        cOut = .ListColumns("Вибуло").Index
        cOpType = .ListColumns("Запасне поле2").Index
        cDocType = .ListColumns("Тип документа").Index
        cDocNum = .ListColumns("№ документа").Index
        cNote = .ListColumns("Примітка").Index
    End With

    ' ---------------- очистити таблицю ----------------
    For i = tbReportingMove.ListRows.count To 1 Step -1
        tbReportingMove.ListRows(i).Delete
    Next i

    ' ---------------- читаємо всю базу одним махом ----------------
    nRows = rngData.Rows.count
    If rngData.Cells.count = 1 Then
        ReDim data(1 To 1, 1 To 1)
        data(1, 1) = rngData.value
    Else
        data = rngData.value
    End If

    Set dict = CreateObject("Scripting.Dictionary")
    minDate = 0

    For i = 1 To nRows
        If i = nRows Or i Mod 100 = 0 Then Call ProgressBar.Progress(i / nRows * 100)

        unitFrom = Trim$(A1Txt(data(i, cFrom)))
        unitTo = Trim$(A1Txt(data(i, cTo)))

        If A1Same(unitFrom, selectedUnit) Or A1Same(unitTo, selectedUnit) Then
            dVal = A1ParseDate(data(i, cDateDoc))
            If dVal = 0 Then
                skippedNoDate = skippedNoDate + 1
            ElseIf dVal <= CDbl(endDate) Then
                opType = LCase$(Trim$(A1Txt(data(i, cOpType))))
                st = A1Status(opType, unitFrom, unitTo, selectedUnit)
                qty = A1Qty(data(i, cIn), data(i, cOut), opType)
                itemName = Trim$(A1Txt(data(i, cName)))

                If st > 0 And qty <> 0 And Len(itemName) > 0 Then
                    itemUnit = Trim$(A1Txt(data(i, cUnit)))
                    key = A1NormKey(itemName) & "||" & A1NormKey(itemUnit)

                    If dict.exists(key) Then
                        info = dict(key)
                    Else
                        ReDim info(0 To 7)
                        info(0) = itemName
                        info(1) = itemUnit
                        info(2) = 0#
                        info(3) = 0#
                        info(4) = 0#
                        Set info(5) = CreateObject("Scripting.Dictionary")
                        Set info(6) = CreateObject("Scripting.Dictionary")
                        Set info(7) = CreateObject("Scripting.Dictionary")
                    End If

                    info(1 + st) = info(1 + st) + qty

                    docType = Trim$(A1Txt(data(i, cDocType)))
                    docNum = Trim$(A1Txt(data(i, cDocNum)))
                    serial = ""
                    If st = 3 And LOST_SERIAL_IN_NOTE Then serial = A1SerialText(A1Txt(data(i, cNote)))

                    Set docs = info(4 + st)
                    docKey = LCase$(docType) & "|" & LCase$(docNum) & "|" & CStr(dVal) & "|" & LCase$(serial)
                    If Not docs.exists(docKey) Then docs.Add docKey, Array(docType, docNum, dVal, serial)

                    dict(key) = info

                    If minDate = 0 Or dVal < minDate Then minDate = dVal
                End If
            End If
        End If
    Next i

    ' ---------------- період у шапці ----------------
    If minDate = 0 Then minDate = CDbl(endDate)
    wsReportingMove.Range("AddOnePeriodFrom") = CDate(minDate)
    wsReportingMove.Range("AddOnePeriodTo") = endDate
    wsReportingMove.Range("AddOnePeriod") = Format$(CDate(minDate), "dd.mm.yyyy") & "-" & Format$(endDate, "dd.mm.yyyy")

    ' ---------------- запис у таблицю ----------------
    counter = 0
    For Each itemKey In dict.Keys
        info = dict(itemKey)
        dIn = info(2)
        dOut = info(3) + info(4)
        dBal = dIn - dOut

        If Round(dIn, 6) <> 0 Or Round(dOut, 6) <> 0 Then
            counter = counter + 1
            If Round(dBal, 6) < 0 Then negCount = negCount + 1

            With tbReportingMove.ListRows.Add
                .Range(1, 1).value = counter
                .Range(1, 2).value = info(0)             ' Найменування
                .Range(1, 3).value = info(1)             ' Одиниця виміру
                .Range(1, 4).value = dIn                 ' Надійшло (Видано)
                .Range(1, 5).value = dOut                ' Вибуло (Повернено + Знищено)
                .Range(1, 6).value = dBal                ' Наявне
                .Range(1, 7).NumberFormat = "General"
                .Range(1, 7).WrapText = True
                .Range(1, 7).value = A1BuildNote(info)   ' Примітка: документи
                .Range(1, 8).value = 0                   ' Залишок на початок (враховується весь рух)
            End With
        End If
    Next itemKey

    On Error Resume Next
    tbReportingMove.DataBodyRange.Rows.AutoFit
    On Error GoTo EH

    msg = "Станом на " & Format$(endDate, "dd.mm.yyyy") & " сформовано рядків: " & counter
    If skippedNoDate > 0 Then msg = msg & vbCrLf & "Записів без дати документа пропущено: " & skippedNoDate
    If negCount > 0 Then msg = msg & vbCrLf & vbCrLf & "УВАГА: рядків з від'ємним наявним (повернено/знищено більше, ніж видано): " & negCount
    MsgBox msg, IIf(negCount > 0, vbExclamation, vbInformation), "Додаток 1: " & selectedUnit
    Exit Sub

EH:
    MsgBox "Помилка " & Err.Number & ": " & Err.Description, vbCritical, "Додаток 1"
End Sub


' =====================================================================
'  Статус і кількість
' =====================================================================

' 1 = Видано (надійшло на підрозділ), 2 = Повернено (на склад), 3 = Знищено (вибуття), 0 = не стосується
Private Function A1Status(ByVal op As String, ByVal unitFrom As String, ByVal unitTo As String, _
                          ByVal sel As String) As Long
    Dim fromSel As Boolean, toSel As Boolean
    fromSel = A1Same(unitFrom, sel)
    toSel = A1Same(unitTo, sel)

    Select Case op
        Case "внутрішнє переміщення", "надходження"
            If toSel And Not fromSel Then
                A1Status = 1
            ElseIf fromSel And Not toSel Then
                A1Status = 2
            End If
        Case "вибуття"
            If fromSel Then A1Status = 3
    End Select
End Function

' Кількість: для надходження спершу "Надійшло", для решти спершу "Вибуло"
Private Function A1Qty(ByVal vIn As Variant, ByVal vOut As Variant, ByVal op As String) As Double
    Dim a As Double, b As Double
    a = A1Num(vIn)
    b = A1Num(vOut)
    If op = "надходження" Then
        If a <> 0 Then
            A1Qty = a
        Else
            A1Qty = b
        End If
    Else
        If b <> 0 Then
            A1Qty = b
        Else
            A1Qty = a
        End If
    End If
End Function

Private Function A1Num(ByVal v As Variant) As Double
    If IsError(v) Then Exit Function
    If IsNull(v) Or IsEmpty(v) Then Exit Function
    If IsNumeric(v) Then A1Num = CDbl(v)
End Function

Private Function A1Txt(ByVal v As Variant) As String
    If IsError(v) Then Exit Function
    If IsNull(v) Then Exit Function
    A1Txt = CStr(v)
End Function

Private Function A1Same(ByVal a As String, ByVal b As String) As Boolean
    a = Trim$(a)
    b = Trim$(b)
    If Len(b) = 0 Then Exit Function
    A1Same = (StrComp(a, b, vbTextCompare) = 0)
End Function

' ключ найменування / одиниці: нижній регістр, без пробілів
Private Function A1NormKey(ByVal s As String) As String
    s = Replace(s, ChrW(160), "")
    s = Replace(s, " ", "")
    A1NormKey = LCase$(s)
End Function


' =====================================================================
'  Примітка: документи по статусах
' =====================================================================

Private Function A1BuildNote(ByVal info As Variant) As String
    Dim st As Long
    Dim q As Double
    Dim res As String

    For st = 1 To 3
        q = info(1 + st)
        If Round(q, 6) <> 0 Then
            If Len(res) > 0 Then res = res & vbLf
            res = res & A1StatusLine(st, q, info(4 + st))
        End If
    Next st
    A1BuildNote = res
End Function

Private Function A1StatusLine(ByVal st As Long, ByVal q As Double, ByVal docs As Object) As String
    Dim head As String
    Dim body As String

    Select Case st
        Case 1
            head = "Видано "
        Case 2
            head = "Повернено "
        Case Else
            head = "Знищено "
    End Select
    head = head & A1FmtQty(q) & " " & QTY_LABEL
    If st = 3 Then head = head & " під час виконання бойових завдань"

    body = A1DocsText(docs)
    If Len(body) > 0 Then head = head & " (" & body & ")"
    A1StatusLine = head
End Function

' "накази № 1/БН/дск, № 5/БН/дск; рапорти № 1981 від 03.02.26, ...; втрачено с/н 25009 від 17.02.26"
Private Function A1DocsText(ByVal docs As Object) As String
    Dim items As Variant
    Dim e As Variant
    Dim n As Long, i As Long, j As Long, cur As Long, cnt As Long
    Dim yr() As Long
    Dim num() As Double
    Dim dt() As Double
    Dim idx() As Long
    Dim grp As Object
    Dim typeKey As String
    Dim typeKeys() As String
    Dim tk As Variant
    Dim tmpS As String
    Dim piece As String
    Dim lost As String
    Dim res As String

    n = docs.count
    If n = 0 Then Exit Function
    items = docs.items

    ReDim yr(0 To n - 1)
    ReDim num(0 To n - 1)
    ReDim dt(0 To n - 1)
    ReDim idx(0 To n - 1)

    For i = 0 To n - 1
        e = items(i)
        dt(i) = CDbl(e(2))
        If dt(i) = 0 Then
            yr(i) = 9999
        Else
            yr(i) = Year(dt(i))
        End If
        num(i) = A1LeadNum(CStr(e(1)))
        idx(i) = i
    Next i

    ' сортування вставками: рік -> номер -> дата
    For i = 1 To n - 1
        cur = idx(i)
        j = i - 1
        Do While j >= 0
            If A1CmpKey(yr(idx(j)), num(idx(j)), dt(idx(j)), yr(cur), num(cur), dt(cur)) <= 0 Then Exit Do
            idx(j + 1) = idx(j)
            j = j - 1
        Loop
        idx(j + 1) = cur
    Next i

    Set grp = CreateObject("Scripting.Dictionary")
    For i = 0 To n - 1
        e = items(idx(i))
        If Len(CStr(e(3))) > 0 Then
            ' втрачений серійний номер
            If Len(lost) > 0 Then lost = lost & ", "
            lost = lost & CStr(e(3))
            If CDbl(e(2)) > 0 Then lost = lost & " від " & Format$(CDate(e(2)), "dd.mm.yy")
        Else
            typeKey = LCase$(Trim$(CStr(e(0))))
            If Len(CStr(e(1))) = 0 Then
                piece = "без номера"
            Else
                piece = ChrW(8470) & " " & CStr(e(1))
            End If
            If InStr(typeKey, NO_DATE_DOC_WORD) = 0 And CDbl(e(2)) > 0 Then
                piece = piece & " від " & Format$(CDate(e(2)), "dd.mm.yy")
            End If
            If grp.exists(typeKey) Then
                grp(typeKey) = grp(typeKey) & ", " & piece
            Else
                grp.Add typeKey, piece
            End If
        End If
    Next i

    ' типи документів - за алфавітом
    cnt = grp.count
    If cnt > 0 Then
        ReDim typeKeys(0 To cnt - 1)
        i = 0
        For Each tk In grp.Keys
            typeKeys(i) = CStr(tk)
            i = i + 1
        Next tk
        For i = 1 To cnt - 1
            tmpS = typeKeys(i)
            j = i - 1
            Do While j >= 0
                If StrComp(typeKeys(j), tmpS, vbTextCompare) <= 0 Then Exit Do
                typeKeys(j + 1) = typeKeys(j)
                j = j - 1
            Loop
            typeKeys(j + 1) = tmpS
        Next i
        For i = 0 To cnt - 1
            If Len(res) > 0 Then res = res & "; "
            res = res & A1Plural(typeKeys(i)) & " " & grp(typeKeys(i))
        Next i
    End If

    If Len(lost) > 0 Then
        If Len(res) > 0 Then res = res & "; "
        res = res & "втрачено с/н " & lost
    End If

    A1DocsText = res
End Function

Private Function A1CmpKey(ByVal ya As Long, ByVal na As Double, ByVal da As Double, _
                          ByVal yb As Long, ByVal nb As Double, ByVal db As Double) As Long
    If ya < yb Then
        A1CmpKey = -1
    ElseIf ya > yb Then
        A1CmpKey = 1
    ElseIf na < nb Then
        A1CmpKey = -1
    ElseIf na > nb Then
        A1CmpKey = 1
    ElseIf da < db Then
        A1CmpKey = -1
    ElseIf da > db Then
        A1CmpKey = 1
    End If
End Function

' "112/п/ІНЖ" -> 112 ; без числа на початку -> в кінець
Private Function A1LeadNum(ByVal s As String) As Double
    Dim i As Long
    s = Trim$(s)
    i = 1
    Do While i <= Len(s)
        If Not (Mid$(s, i, 1) Like "#") Then Exit Do
        i = i + 1
    Loop
    If i > 1 Then
        A1LeadNum = val(Left$(s, i - 1))
    Else
        A1LeadNum = 1E+300
    End If
End Function

' множина типу документа: накладна -> накладні, рапорт -> рапорти, наказ -> накази, акт -> акти
Private Function A1Plural(ByVal t As String) As String
    Dim words() As String
    Dim w As String
    Dim lastCh As String

    t = LCase$(Trim$(t))
    If Len(t) = 0 Then
        A1Plural = "документи"
        Exit Function
    End If

    words = Split(t, " ")
    w = words(0)
    If Len(w) >= 2 Then
        lastCh = Right$(w, 1)
        If lastCh = "а" Then
            If InStr("кгхчшщж", Mid$(w, Len(w) - 1, 1)) > 0 Then
                w = Left$(w, Len(w) - 1) & "и"
            Else
                w = Left$(w, Len(w) - 1) & "і"
            End If
        ElseIf InStr("бвгґджзйклмнпрстфхцчшщ", lastCh) > 0 Then
            w = w & "и"
        End If
    End If
    words(0) = w
    A1Plural = Join(words, " ")
End Function

Private Function A1FmtQty(ByVal x As Double) As String
    If Abs(x - Round(x, 0)) < 0.000001 Then
        A1FmtQty = CStr(CLng(Round(x, 0)))
    Else
        A1FmtQty = Replace(CStr(Round(x, 3)), ".", ",")
    End If
End Function


' =====================================================================
'  Серійний номер у "Примітці"
' =====================================================================

' Повертає серійний номер (як у клітинці), якщо примітка схожа на нього: є цифри, 4-30 знаків літер/цифр
Private Function A1SerialText(ByVal s As String) As String
    Dim i As Long, ch As Long
    Dim cnt As Long
    Dim hasDigit As Boolean

    s = Trim$(s)
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        ch = AscW(Mid$(s, i, 1))
        If ch < 0 Then ch = ch + 65536
        If ch >= 48 And ch <= 57 Then
            hasDigit = True
            cnt = cnt + 1
        ElseIf (ch >= 65 And ch <= 90) Or (ch >= 97 And ch <= 122) Or (ch >= 1024 And ch <= 1279) Then
            cnt = cnt + 1
        End If
    Next i
    If hasDigit And cnt >= 4 And cnt <= 30 And Len(s) <= 40 Then A1SerialText = s
End Function


' =====================================================================
'  Дата з клітинки або тексту (дд.мм.рррр / рррр-мм-дд). 0 = не розпізнано.
' =====================================================================

Private Function A1ParseDate(ByVal v As Variant) As Double
    Dim s As String
    Dim p() As String
    Dim res As Double

    If IsError(v) Then Exit Function
    If IsNull(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbDate Then
        A1ParseDate = CDbl(v)
        Exit Function
    End If
    If VarType(v) = vbDouble Then
        If v > 0 And v < 2958466 Then A1ParseDate = CDbl(v)
        Exit Function
    End If

    s = Trim$(CStr(v))
    If Len(s) = 0 Then Exit Function
    If InStr(s, " ") > 0 Then s = Left$(s, InStr(s, " ") - 1)

    If InStr(s, ".") > 0 Then
        p = Split(s, ".")
        If UBound(p) = 2 Then
            If IsNumeric(p(0)) And IsNumeric(p(1)) And IsNumeric(p(2)) Then
                res = A1MakeDate(CLng(p(2)), CLng(p(1)), CLng(p(0)))
                If res > 0 Then
                    A1ParseDate = res
                    Exit Function
                End If
            End If
        End If
    ElseIf InStr(s, "-") > 0 Then
        p = Split(s, "-")
        If UBound(p) = 2 Then
            If IsNumeric(p(0)) And IsNumeric(p(1)) And IsNumeric(p(2)) Then
                res = A1MakeDate(CLng(p(0)), CLng(p(1)), CLng(p(2)))
                If res > 0 Then
                    A1ParseDate = res
                    Exit Function
                End If
            End If
        End If
    End If

    If IsDate(s) Then A1ParseDate = CDbl(CDate(s))
End Function

Private Function A1MakeDate(ByVal y As Long, ByVal m As Long, ByVal d As Long) As Double
    If y < 100 Then y = y + 2000
    If y >= 1900 And y <= 2200 And m >= 1 And m <= 12 And d >= 1 And d <= 31 Then
        A1MakeDate = CDbl(DateSerial(y, m, d))
    End If
End Function
