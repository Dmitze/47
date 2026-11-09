Attribute VB_Name = "Inventory"
'Inventory Формування інвентаризиційного опису

Option Explicit

' === Головна процедура: формує Інвентаризаційний опис з проміжними підсумками по рахунках ===
Public Sub BuildInventoryReport()
'(ByVal unitFilter As String)
    Dim unitFilter As String
    Dim wsSrc As Worksheet, ws As Worksheet
    Dim lastRow As Long, r As Long

    Dim dictBal As Object, dictUom As Object, dictCode As Object, dictPrice As Object, dictAccount As Object
    Dim itemName As String, uom As String, itemCode As String
    Dim unitIn As String, unitOut As String, opType As String
    Dim qty As Double, price As Double, d As Date, key As String
    Dim account As String

    Dim outRow As Long
    Dim k As Variant, parts As Variant, bal As Double
    Dim currentName As String, prevName As String
    Dim currentAccount As String, prevAccount As String
    Dim sumF As Double, sumH As Double, sumI As Double, sumK As Double
    Dim accF As Double, accH As Double, acci As Double, accK As Double
    Dim grandF As Double, grandH As Double, grandI As Double, grandK As Double
    Dim posCounter As Long, groupStartRow As Long
    Dim keysSorted As Variant
    Dim kIdx As Long
    Dim totalRow As Long, lastPos As Long
    
    Dim wbAllData As Workbook
    Dim wsAllData As Worksheet
    Dim wsReportingMove As Worksheet
    Dim tbAllData As ListObject
    Dim tbReportingMove As ListObject

    Set wbAllData = ThisWorkbook
    Set wsAllData = wbAllData.Worksheets("База")
    Set tbAllData = wsAllData.ListObjects("AllData")
    
    ' --- Підготовка аркуша призначення ---
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Інвентаризаційний опис")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=wsSrc)
        ws.name = "Інвентаризаційний опис"
    End If
    On Error GoTo 0
    
    
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
            
            Dim t As Long
            t = 1
            wsTroops.Range("A1:A" & wsTroops.Rows.count).ClearContents ' Очистити перед вставкою
            
            Dim unitKey As Variant
            For Each unitKey In dictUnits.Keys
                wsTroops.Cells(t, 1).value = unitKey
                t = t + 1
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
                unitFilter = frm.ListBox1.value
            Else
                MsgBox "Підрозділ не вибрано!", vbExclamation
                Exit Sub
            End If
    
      
'список субрахунків відповідно до вибраного типу
Dim dictAllowedAccounts As Object
Dim tbAccaunts As ListObject
Dim accRow As ListRow
Dim accountNumber As String
Dim accountType As String
Dim selectedAccountType As String

Set dictAllowedAccounts = CreateObject("Scripting.Dictionary")
dictAllowedAccounts.CompareMode = vbTextCompare

Set tbAccaunts = ThisWorkbook.Worksheets("Довідники").ListObjects("tbAccaunts")

selectedAccountType = ws.Range("invAccauntType").value

' Формуємо перелік дозволених субрахунків
For Each accRow In tbAccaunts.ListRows

    accountNumber = Trim$(CStr(accRow.Range.Cells(1, 1).value))
    accountType = Trim$(CStr(accRow.Range.Cells(1, 3).value))

    If accountNumber <> "" Then
        If StrComp(accountType, selectedAccountType, vbTextCompare) = 0 Then
            If Not dictAllowedAccounts.exists(accountNumber) Then
                dictAllowedAccounts.Add accountNumber, True
            End If
        End If
    End If

Next accRow








    ' Посилання на джерело
    Set wsSrc = ThisWorkbook.Sheets("База")
    lastRow = wsSrc.Cells(wsSrc.Rows.count, "A").End(xlUp).Row

    ' Ініціалізація словників
    Set dictBal = CreateObject("Scripting.Dictionary")
    Set dictUom = CreateObject("Scripting.Dictionary")
    Set dictCode = CreateObject("Scripting.Dictionary")
    Set dictPrice = CreateObject("Scripting.Dictionary")
    Set dictAccount = CreateObject("Scripting.Dictionary")

    ' --- Обхід рядків БАЗИ: накопичення балансу по unitFilter ---
    For r = 2 To lastRow
        If IsDate(wsSrc.Cells(r, "A").value) Then
            d = CDate(wsSrc.Cells(r, "A").value)
        Else
            GoTo NextRow
        End If

        itemName = Trim$(CStr(wsSrc.Cells(r, "B").value))
        If itemName = "" Then GoTo NextRow

        itemCode = Trim$(CStr(wsSrc.Cells(r, "R").value))
        uom = Trim$(CStr(wsSrc.Cells(r, "D").value))
        unitOut = Trim$(CStr(wsSrc.Cells(r, "H").value))
        unitIn = Trim$(CStr(wsSrc.Cells(r, "I").value))
        opType = LCase$(Trim$(CStr(wsSrc.Cells(r, "Z").value)))
        account = Trim$(CStr(wsSrc.Cells(r, "X").value))
        
        If Not dictAllowedAccounts.exists(account) Then GoTo NextRow
        
        price = 0#
        If IsNumeric(wsSrc.Cells(r, "T").value) Then
            price = CDbl(wsSrc.Cells(r, "T").value)
        End If

' === 3.9.27 ===
        ' Ключ групування: itemName|price (можна змінити на інший варіант)
        key = itemName & "|" & CStr(price) & "|" & CStr(uom) & "|" & CStr(itemCode)
'key = itemName & "|" & CStr(price)
' === 3.9.27 ===

' --- Надходження ---
        If opType = "надходження" And IsNumeric(wsSrc.Cells(r, "F").value) Then
            qty = CDbl(wsSrc.Cells(r, "F").value)
            If unitIn = unitFilter Then
                AccumulateBalance dictBal, dictUom, dictCode, dictPrice, dictAccount, key, qty, uom, itemCode, price, account
            End If
        End If

        ' --- Внутрішнє переміщення ---
        If opType = "внутрішнє переміщення" And IsNumeric(wsSrc.Cells(r, "G").value) Then
            qty = CDbl(wsSrc.Cells(r, "G").value)
            If unitOut = unitFilter Then
                AccumulateBalance dictBal, dictUom, dictCode, dictPrice, dictAccount, key, -qty, uom, itemCode, price, account
            End If
            If unitIn = unitFilter Then
                AccumulateBalance dictBal, dictUom, dictCode, dictPrice, dictAccount, key, qty, uom, itemCode, price, account
            End If
        End If

        ' --- Вибуття ---
        If opType = "вибуття" And IsNumeric(wsSrc.Cells(r, "G").value) Then
            qty = CDbl(wsSrc.Cells(r, "G").value)
            If unitOut = unitFilter Then
                AccumulateBalance dictBal, dictUom, dictCode, dictPrice, dictAccount, key, -qty, uom, itemCode, price, account
            End If
        End If

NextRow:
    Next r

    
    
            
     ' Записати найменування підрозділу
     ws.Range("invUnit") = "у " & unitFilter
     
     ' Формуємо перелік задіяних субрахунків
     FillInventorySubaccounts dictAccount
    
    ' Очистити аркуш
    Dim ivnLastRowTable As Long 'клітинка, де фраза "Разом за описом"
    ivnLastRowTable = ws.Range("ivnLastRowTable").Row
    
    ws.Range("A24:L" & ivnLastRowTable - 2).Delete

   
    ' --- Вивантаження з груповими підсумками і проміжними підсумками по рахунках ---
    outRow = 24
    prevName = ""
    prevAccount = ""
    posCounter = 1
    groupStartRow = 0
    sumF = 0#: sumH = 0#: sumI = 0#: sumK = 0#
    accF = 0#: accH = 0#: acci = 0#: accK = 0#
    grandF = 0#: grandH = 0#: grandI = 0#: grandK = 0#

    ' Відсортовані ключі: спочатку за рахунком, потім за найменуванням
    keysSorted = SortedKeysByAccountName(dictAccount, dictBal)

    For kIdx = LBound(keysSorted) To UBound(keysSorted)
        k = keysSorted(kIdx)
        bal = NzNum(dictBal(k))
        If bal > 0 Then
            currentAccount = KeyAccount(k, dictAccount)
            currentName = KeyItemName(k)

            ' Якщо змінився рахунок — закриваємо попередній рахунок (включно з останньою групою)
            If prevAccount <> "" And currentAccount <> prevAccount Then
                ' Закрити останню групу попереднього рахунку, якщо вона відкрита
                If prevName <> "" Then
                
                ws.Rows(outRow).Insert Shift:=xlDown
                
                    ws.Cells(outRow, 3).value = "ВСЬОГО " & prevName
                    ws.Cells(outRow, 6).value = sumF
                    ws.Cells(outRow, 8).value = sumH
                    ws.Cells(outRow, 9).value = sumI
                    ws.Cells(outRow, 11).value = sumK
                    ws.Rows(outRow).Font.Bold = True

                    If groupStartRow > 0 Then
                        If (outRow - 1) > groupStartRow Then
                            With ws.Range("A" & groupStartRow & ":A" & outRow - 1)
                                .Merge
                                .value = posCounter
                                .HorizontalAlignment = xlCenter
                                .VerticalAlignment = xlCenter
                            End With
                        Else
                            ws.Cells(groupStartRow, "A").value = posCounter
                            ws.Cells(groupStartRow, "A").HorizontalAlignment = xlCenter
                            ws.Cells(groupStartRow, "A").VerticalAlignment = xlCenter
                        End If
                        posCounter = posCounter + 1
                    End If

                    accF = accF + sumF
                    accH = accH + sumH
                    acci = acci + sumI
                    accK = accK + sumK

                    outRow = outRow + 1
                    sumF = 0#: sumH = 0#: sumI = 0#: sumK = 0#
                    groupStartRow = 0
                    prevName = ""
                End If

                ' Виводимо проміжний підсумок по попередньому рахунку
                
                ws.Rows(outRow).Insert Shift:=xlDown
                
                ws.Range("A" & outRow & ":E" & outRow).Merge
                ws.Range("A" & outRow).value = "ВСЬОГО по рахунку " & prevAccount
                ws.Range("A" & outRow).HorizontalAlignment = xlCenter
                ws.Range("A" & outRow).VerticalAlignment = xlCenter
                ws.Range("A" & outRow).Font.Bold = True

                ws.Cells(outRow, 6).value = accF
                ws.Cells(outRow, 8).value = accH
                ws.Cells(outRow, 9).value = acci
                ws.Cells(outRow, 11).value = accK
                ws.Rows(outRow).Font.Bold = True

                ' Додаємо до загального підсумку
                grandF = grandF + accF
                grandH = grandH + accH
                grandI = grandI + acci
                grandK = grandK + accK

                outRow = outRow + 1

                ' Скидаємо проміжні підсумки по рахунку
                accF = 0#: accH = 0#: acci = 0#: accK = 0#
            End If

            ' Якщо змінилась група найменування — закриваємо попередню групу
            If prevName <> "" And currentName <> prevName Then
            
            ws.Rows(outRow).Insert Shift:=xlDown
            
                ws.Cells(outRow, 3).value = "ВСЬОГО " & prevName
                ws.Cells(outRow, 6).value = sumF
                ws.Cells(outRow, 8).value = sumH
                ws.Cells(outRow, 9).value = sumI
                ws.Cells(outRow, 11).value = sumK
                ws.Rows(outRow).Font.Bold = True

                If groupStartRow > 0 Then
                    If (outRow - 1) > groupStartRow Then
                        With ws.Range("A" & groupStartRow & ":A" & outRow - 1)
                            .Merge
                            .value = posCounter
                            .HorizontalAlignment = xlCenter
                            .VerticalAlignment = xlCenter
                        End With
                    Else
                        ws.Cells(groupStartRow, "A").value = posCounter
                        ws.Cells(groupStartRow, "A").HorizontalAlignment = xlCenter
                        ws.Cells(groupStartRow, "A").VerticalAlignment = xlCenter
                    End If
                    posCounter = posCounter + 1
                End If

                accF = accF + sumF
                accH = accH + sumH
                acci = acci + sumI
                accK = accK + sumK

                outRow = outRow + 1
                sumF = 0#: sumH = 0#: sumI = 0#: sumK = 0#
                groupStartRow = 0
            End If

            ' Основний рядок
            ws.Rows(outRow).Insert Shift:=xlDown
            
            If dictAccount.exists(k) Then
                ws.Cells(outRow, 2).value = dictAccount(k)
            Else
                ws.Cells(outRow, 2).value = ""
            End If
            ws.Cells(outRow, 3).value = currentName
            ws.Cells(outRow, 4).value = dictCode(k)
            ws.Cells(outRow, 5).value = dictUom(k)
            ws.Cells(outRow, 6).value = bal
            ws.Cells(outRow, 7).value = dictPrice(k)
            ws.Cells(outRow, 8).FormulaR1C1 = "=RC[-2]*RC[-1]"
            ws.Rows(outRow).Font.Bold = False

            ' Для сумарних колонок дублюємо (за потреби)
            ws.Cells(outRow, 9).value = ws.Cells(outRow, 6).value
            ws.Cells(outRow, 10).value = ws.Cells(outRow, 7).value
            ws.Cells(outRow, 11).value = ws.Cells(outRow, 8).value

            If groupStartRow = 0 Then groupStartRow = outRow

            sumF = sumF + NzNum(ws.Cells(outRow, 6).value)
            sumH = sumH + NzNum(ws.Cells(outRow, 8).value)
            sumI = sumI + NzNum(ws.Cells(outRow, 9).value)
            sumK = sumK + NzNum(ws.Cells(outRow, 11).value)

            prevName = currentName
            prevAccount = currentAccount
            outRow = outRow + 1
        End If
    Next kIdx

    ' Закриваємо останню відкриту групу (якщо є)
    If prevName <> "" Then
    
    ws.Rows(outRow).Insert Shift:=xlDown
    
        ws.Cells(outRow, 3).value = "ВСЬОГО " & prevName
        ws.Cells(outRow, 6).value = sumF
        ws.Cells(outRow, 8).value = sumH
        ws.Cells(outRow, 9).value = sumI
        ws.Cells(outRow, 11).value = sumK
        ws.Rows(outRow).Font.Bold = True

        If groupStartRow > 0 Then
            If (outRow - 1) > groupStartRow Then
                With ws.Range("A" & groupStartRow & ":A" & outRow - 1)
                    .Merge
                    .value = posCounter
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                End With
            Else
                ws.Cells(groupStartRow, "A").value = posCounter
                ws.Cells(groupStartRow, "A").HorizontalAlignment = xlCenter
                ws.Cells(groupStartRow, "A").VerticalAlignment = xlCenter
            End If
            posCounter = posCounter + 1
        End If

        accF = accF + sumF
        accH = accH + sumH
        acci = acci + sumI
        accK = accK + sumK

        outRow = outRow + 1
    End If

    ' Якщо залишився незакритий рахунок — виводимо його проміжний підсумок
    If prevAccount <> "" Then
    
    ws.Rows(outRow).Insert Shift:=xlDown
    
        ws.Range("A" & outRow & ":E" & outRow).Merge
        ws.Range("A" & outRow).value = "ВСЬОГО по рахунку " & prevAccount
        ws.Range("A" & outRow).HorizontalAlignment = xlCenter
        ws.Range("A" & outRow).VerticalAlignment = xlCenter
        ws.Range("A" & outRow).Font.Bold = True

        ws.Cells(outRow, 6).value = accF
        ws.Cells(outRow, 8).value = accH
        ws.Cells(outRow, 9).value = acci
        ws.Cells(outRow, 11).value = accK
        ws.Rows(outRow).Font.Bold = True

        grandF = grandF + accF
        grandH = grandH + accH
        grandI = grandI + acci
        grandK = grandK + accK

        outRow = outRow + 1
    End If

    ' Загальний підсумок по всіх рахунках
    ws.Rows(outRow).Insert Shift:=xlDown
    
    ws.Range("A" & outRow & ":E" & outRow).Merge
    ws.Range("A" & outRow).value = "РАЗОМ"
    ws.Range("A" & outRow).HorizontalAlignment = xlCenter
    ws.Range("A" & outRow).VerticalAlignment = xlCenter
    ws.Range("A" & outRow).Font.Bold = True

    ws.Cells(outRow, 6).value = grandF
    ws.Cells(outRow, 8).value = grandH
    ws.Cells(outRow, 9).value = grandI
    ws.Cells(outRow, 11).value = grandK
    ws.Rows(outRow).Font.Bold = True

    totalRow = outRow
    lastPos = posCounter - 1

      ' --- Форматування таблиці: вирівнювання колонок, без AutoFit ---
    With ws.Range("A24:L" & totalRow)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Font.name = "Times New Roman"
        .Font.Size = 12
        .VerticalAlignment = xlCenter
    End With
    
    ' Колонка B — по центру
    With ws.Range("B24:B" & totalRow)
        .NumberFormat = "@"          ' текстовий формат
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    
    ' Колонка C — по лівому краю
    With ws.Range("C24:C" & totalRow)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    
    ' Колонка D — по центру
    With ws.Range("D24:D" & totalRow)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Колонки E:L — по центру
    With ws.Range("E24:L" & totalRow)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
     
    ' Формати чисел залишаються як раніше
    ws.Range("G24:G" & totalRow).NumberFormat = "0.00#########################"
    ws.Range("J24:J" & totalRow).NumberFormat = "0.00#########################"
    ws.Range("H24:H" & totalRow).NumberFormat = "0.00"
    ws.Range("K24:K" & totalRow).NumberFormat = "0.00"

'    ' --- Блок "Разом за описом" ---

' а) кількість порядкових номерів:
    ws.Range("invTotalNumbers").value = ЧислоПрописом11(lastPos)
    
    ws.Range("invlastPos").value = lastPos

'  б) загальна кількість одиниць (фактично):
    ws.Range("invTotalQuantity").value = ЧислоПрописом11(grandF)

'  в) вартість фактична:
    ws.Range("invActualCost").value = LCase(СумаПрописом1(grandH))


' г) загальна кількість одиниць за даними бухгалтерського обліку:
    ws.Range("invTotalQuantityAccaunting").value = ЧислоПрописом11(grandI)


'ґ) вартість за даними бухгалтерського обліку:
    ws.Range("invActualCostAccaunting").value = LCase(СумаПрописом1(grandK))


With ws.PageSetup
    .PaperSize = xlPaperA4
    .Orientation = xlLandscape
        
    .Zoom = False
    .FitToPagesWide = 1
    .FitToPagesTall = False
    'тимчасово
    .PrintArea = "$A$1:$L$" & ws.Range("invLastRow").Row
    
End With

    MsgBox "Інвентаризаційний опис підрозділу: " & unitFilter, vbInformation, "Інвентаризаційний опис сформовано"
End Sub

' === Допоміжні процедури та функції ===

' Безпечне накопичення в словниках (тепер з dictAccount)
Private Sub AccumulateBalance(ByRef dictBal As Object, ByRef dictUom As Object, ByRef dictCode As Object, ByRef dictPrice As Object, ByRef dictAccount As Object, _
                              ByVal key As String, ByVal deltaQty As Double, ByVal uom As String, ByVal itemCode As String, ByVal price As Double, ByVal account As String)
    If Not dictBal.exists(key) Then dictBal.Add key, 0#
    dictBal(key) = dictBal(key) + deltaQty

    If Not dictUom.exists(key) Then dictUom.Add key, uom
    If Not dictCode.exists(key) Then dictCode.Add key, itemCode
    If Not dictPrice.exists(key) Then dictPrice.Add key, price
    If Not dictAccount.exists(key) Then dictAccount.Add key, account
End Sub

' Повертає найменування з ключа itemName|price
Private Function KeyItemName(ByVal k As String) As String
    Dim p As Long
    p = InStr(1, k, "|", vbBinaryCompare)
    If p > 0 Then
        KeyItemName = Left$(k, p - 1)
    Else
        KeyItemName = k
    End If
End Function

' Повертає рахунок для ключа (через dictAccount)
Private Function KeyAccount(ByVal k As String, ByRef dictAccount As Object) As String
    If dictAccount.exists(k) Then
        KeyAccount = CStr(dictAccount(k))
    Else
        KeyAccount = ""
    End If
End Function

' Сортує ключі спочатку за рахунком, потім за найменуванням
Private Function SortedKeysByAccountName(ByRef dictAccount As Object, ByRef dictBal As Object) As Variant
    Dim arr As Variant, i As Long, j As Long, tmp As Variant
    arr = dictBal.Keys
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            Dim ai As String, aj As String, acci As String, accj As String
            ai = KeyItemName(CStr(arr(i)))
            aj = KeyItemName(CStr(arr(j)))
            acci = KeyAccount(CStr(arr(i)), dictAccount)
            accj = KeyAccount(CStr(arr(j)), dictAccount)
            If StrComp(Trim$(acci) & "|" & Trim$(ai), Trim$(accj) & "|" & Trim$(aj), vbTextCompare) > 0 Then
                tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            End If
        Next j
    Next i
    SortedKeysByAccountName = arr
End Function

' Захисна функція: якщо Null або нечисло — повернути 0
Private Function NzNum(ByVal v As Variant) As Double
    If IsError(v) Then NzNum = 0: Exit Function
    If IsNull(v) Then NzNum = 0: Exit Function
    If v = "" Then NzNum = 0: Exit Function
    If Not IsNumeric(v) Then NzNum = 0: Exit Function
    NzNum = CDbl(v)
End Function



Private Sub FillInventorySubaccounts(ByVal dictAccount As Object)

    Dim wsDict As Worksheet
    Dim tbAccounts As ListObject
    Dim dictNames As Object
    Dim dictUsed As Object
    
    Dim cell As Range
    Dim account As Variant
    Dim accountName As String
    Dim result As String
    
    Set wsDict = ThisWorkbook.Worksheets("Довідники")
    Set tbAccounts = wsDict.ListObjects("tbAccaunts")
    
    Set dictNames = CreateObject("Scripting.Dictionary")
    Set dictUsed = CreateObject("Scripting.Dictionary")
    
    dictNames.CompareMode = vbTextCompare
    dictUsed.CompareMode = vbTextCompare
    
    ' Завантажуємо довідник субрахунків
    If Not tbAccounts.DataBodyRange Is Nothing Then
        
        For Each cell In tbAccounts.ListColumns(1).DataBodyRange
            
            account = Trim$(CStr(cell.value))
            
            If account <> "" Then
                accountName = Trim$(CStr(cell.Offset(0, 1).value))
                
                If Not dictNames.exists(account) Then
                    dictNames.Add account, accountName
                End If
            End If
            
        Next cell
        
    End If
    
    ' Визначаємо унікальні задіяні субрахунки
    For Each account In dictAccount.items
        
        account = Trim$(CStr(account))
        
        If account <> "" Then
            
            If Not dictUsed.exists(account) Then
                dictUsed.Add account, True
            End If
            
        End If
        
    Next account
    
    ' Формуємо текст
    For Each account In dictUsed.Keys
        
        If result <> "" Then result = result & ", "
        
        result = result & account
        
        If dictNames.exists(account) Then
            accountName = dictNames(account)
            
            If accountName <> "" Then
                result = result & " " & accountName
            End If
        Else
            result = result & " *****не знайдено в довіднику*****"
        End If
        
    Next account
    
    ' Запис у поле
    ThisWorkbook.Worksheets("Інвентаризаційний опис") _
        .Range("invSubs").value = result

End Sub










' === Примітка ===
' Функції NumberToWordsUA та NumberToWordsUA_Grn викликаються у блоці "Разом за описом".
' Якщо їх немає в проекті, додайте реалізації або замініть виклики на інші.
' При потребі я можу додати реалізації перетворення чисел у слова українською.

' === 3.9.27 ===
Sub SaveInventory()

        ThisWorkbook.Sheets("Інвентаризаційний опис").Select 'копіюємо перший лист з нарядом
        ThisWorkbook.Sheets("Інвентаризаційний опис").Copy 'зберігаємо як нову книгу
   
     Dim WbNew As Workbook
     Dim cell
        Set WbNew = ActiveWorkbook
   
    Dim shp As Shape
    For Each shp In WbNew.Worksheets(1).Shapes
        shp.Delete
    Next shp
    
    For Each cell In WbNew.Worksheets(1).Range(Cells(1, 1), Cells(WbNew.Worksheets(1).Range("invLastRow").Row, 17))
        If cell.HasFormula Then
            cell.value = cell.value
        End If
    Next cell
    
    
        'зберігаємо файл
'        On Error GoTo WrongPath
         WbNew.SaveAs fileName:=ThisWorkbook.Path & Application.PathSeparator & _
         CleanFileName("Інвентаризаційний опис запасів " & WbNew.Worksheets(1).Range("invAsOf").value & " " & _
          WbNew.Worksheets(1).Range("invUnit").value & ".xlsx"), FileFormat:=xlOpenXMLWorkbook
'     On Error GoTo 0
'      Exit Sub
'WrongPath:
     MsgBox "Інвентаризаційний опис збережений і тій же папці, де файл бази даних", vbInformation, ThisWorkbook.Path
'     Exit Sub
     
End Sub


' === 3.9.27 ===

