Attribute VB_Name = "add47PNP"
Option Explicit

Sub CreateAdd47PNP()
    ' Для категорій придатне / непридатне / потребує ремонту
    Dim wbAllData As Workbook
    Dim wsAllData As Worksheet
    Dim tbAllData As ListObject
    Dim wsAdd47 As Worksheet
    
    Set wbAllData = ThisWorkbook
    Set wsAllData = wbAllData.Worksheets("База")
    Set tbAllData = wsAllData.ListObjects("AllData")
    Set wsAdd47 = wbAllData.Worksheets("Додаток 47 пр-непр")
    
    Dim ptReduses As PivotTable
    On Error Resume Next
    Set ptReduses = ThisWorkbook.Worksheets("Довідник д47").PivotTables("ptReduses47")
    If Not ptReduses Is Nothing Then ptReduses.RefreshTable
    On Error GoTo 0

    ' Видаляємо попередні динамічні дані (колонки від P/16 і рядки від 21)
    Dim lastColumnStart As Long
    lastColumnStart = wsAdd47.Cells(19, wsAdd47.Columns.Count).End(xlToLeft).Column
    If lastColumnStart < 16 Then lastColumnStart = 16
    
    Dim LastRowStart As Long
    LastRowStart = wsAdd47.Cells.Find(What:="*", After:=wsAdd47.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
                            SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    If LastRowStart < 21 Then LastRowStart = 21

    On Error Resume Next
    If lastColumnStart >= 16 Then
        wsAdd47.Range(wsAdd47.Cells(1, 16), wsAdd47.Cells(LastRowStart, lastColumnStart)).EntireColumn.Delete
    End If
    If LastRowStart >= 21 Then
        wsAdd47.Rows("21:" & wsAdd47.Rows.Count).Delete
    End If
    On Error GoTo 0

    ' Викликаємо форму для вибору майна
    Dim frm As UserForm1
    Set frm = UserForms.Add(UserForm1.Name)
    frm.ListData = ThisWorkbook.Worksheets("Довідник д47").Range("Goods47")
    frm.Show
    
    Dim NameOfThings As String
    NameOfThings = frm.TextBox1.Value
    If Len(Trim$(NameOfThings)) = 0 Then Exit Sub
    
    Dim ListGoods As Range
    Dim rngWithName As Range
    Set ListGoods = ThisWorkbook.Worksheets("Довідник д47").Range("Goods47")
    Set rngWithName = ListGoods.Find(NameOfThings, LookIn:=xlValues, LookAt:=xlWhole)
    
    If rngWithName Is Nothing Then
        MsgBox "Майно не знайдено в довіднику. Спробуйте ще раз!", vbCritical, "Помилка вибору"
        Exit Sub
    End If
    
    Dim name As String
    Dim Nomenclature As String
    Dim unit As String
    Dim Pice As Double
    
    name = rngWithName.Offset(, -12).Value
    Nomenclature = rngWithName.Offset(, -10).Value
    If Nomenclature = "(пусто)" Or Len(Trim$(Nomenclature)) = 0 Then
        MsgBox "Ймовірно, ви не вказали код номенклатури вибраного майна (якщо не відомо, поставте б/н)!", vbCritical, "Код номенклатури не вказаний"
        Exit Sub
    End If
    
    unit = rngWithName.Offset(, -11).Value
    If unit = "(пусто)" Or Len(Trim$(unit)) = 0 Then
        MsgBox "Ймовірно, ви не вказали одиницю виміру вибраного майна!", vbCritical, "Одиниця виміру не вказана"
        Exit Sub
    End If
    
    On Error GoTo WrongPrice
    Pice = CDbl(rngWithName.Offset(, -9).Value)
    On Error GoTo 0
    
    ' Заповнюємо критерії фільтрації
    Dim criteriaRange47 As Range
    Set criteriaRange47 = wsAdd47.Range("criteriaRange47")
    With criteriaRange47
        .Cells(2, 1) = "'=" & name
        .Cells(2, 2) = "'=" & unit
        .Cells(2, 3) = "'=" & Nomenclature
        .Cells(2, 4) = Pice
    End With
    
    If wsAllData.FilterMode Then wsAllData.ShowAllData
    
    ' Фільтруємо базу
    tbAllData.Range.AdvancedFilter Action:=xlFilterInPlace, criteriaRange:=criteriaRange47, Unique:=False
    
    Dim visibleRange As Range
    On Error Resume Next
    Set visibleRange = tbAllData.DataBodyRange.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    
    If visibleRange Is Nothing Then
        If wsAllData.FilterMode Then wsAllData.ShowAllData
        MsgBox "Записів по вибраному майну в базі не знайдено!", vbInformation, "Немає даних"
        Exit Sub
    End If
    
    Dim rowCount As Long, colCount As Long, area As Range, cell As Range
    rowCount = 0
    colCount = tbAllData.ListColumns.Count
    For Each area In visibleRange.Areas
        rowCount = rowCount + area.Rows.Count
    Next area
    
    If rowCount = 0 Then
        If wsAllData.FilterMode Then wsAllData.ShowAllData
        MsgBox "Записів по вибраному майну в базі не знайдено!", vbInformation, "Немає даних"
        Exit Sub
    End If
    
    Dim filteredData() As Variant
    ReDim filteredData(1 To rowCount + 1, 1 To colCount)
    
    Dim k As Long, i As Long, j As Long
    For k = 1 To colCount
        filteredData(1, k) = tbAllData.HeaderRowRange.Cells(1, k).Value
    Next k
    
    i = 2
    For Each area In visibleRange.Areas
        For Each cell In area.Rows
            For j = 1 To colCount
                filteredData(i, j) = CStr(cell.Cells(1, j).Value)
            Next j
            i = i + 1
        Next cell
    Next area
    
    If wsAllData.FilterMode Then wsAllData.ShowAllData
    
    ' Сортуємо хронологічно: спочатку за датою (роком), потім за номером документа
    SortFilteredData filteredData
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Dim lastRow47 As Long
    Dim lastColFinal As Long
    lastRow47 = 20
    lastColFinal = 15 ' Колонки 12-15 - це 'Склад'
    
    For j = 8 To 15
        If j = 8 Or (j >= 12 And j <= 15) Then
            wsAdd47.Cells(20, j).Value = 0
        End If
    Next j
    
    Dim n As Long
    Dim docEntryDate As Variant, docType As String, docNumber As String, docDate As Variant, docParties As String
    Dim typeOperation As String, fromUnit As String, toUnit As String
    Dim catFrom As Long, catTo As Long, qty As Double
    Dim isSameDoc As Boolean
    Dim uColFrom As Long, uColTo As Long
    
    For n = 2 To UBound(filteredData, 1)
        docEntryDate = filteredData(n, 1)
        docType = Trim$(CStr(filteredData(n, 25)))
        docNumber = Trim$(CStr(filteredData(n, 14)))
        docDate = filteredData(n, 13)
        fromUnit = Trim$(CStr(filteredData(n, 8)))
        toUnit = Trim$(CStr(filteredData(n, 9)))
        If Len(fromUnit) = 0 Then fromUnit = "Склад"
        If Len(toUnit) = 0 Then toUnit = "Склад"
        
        docParties = fromUnit & vbLf & "(" & toUnit & ")"
        typeOperation = LCase$(Trim$(CStr(filteredData(n, 26))))
        
        If typeOperation = "надходження" Then
            qty = SafeVal(filteredData(n, 6))
            If qty = 0 Then qty = SafeVal(filteredData(n, 7))
        Else
            qty = SafeVal(filteredData(n, 7))
            If qty = 0 Then qty = SafeVal(filteredData(n, 6))
        End If
        
        If qty > 0 Then
            catFrom = MapCatPNP(filteredData(n, 5))
            catTo = MapCatPNP(filteredData(n, 27))
            If catTo = 0 Then catTo = catFrom
            If catFrom = 0 Then catFrom = catTo
            If catTo = 0 Then catTo = 1
            If catFrom = 0 Then catFrom = 1
            
            isSameDoc = False
            If lastRow47 > 20 Then
                If CStr(wsAdd47.Cells(lastRow47, 1).Value) = CStr(docEntryDate) And _
                   Trim$(CStr(wsAdd47.Cells(lastRow47, 2).Value)) = docType And _
                   Trim$(CStr(wsAdd47.Cells(lastRow47, 3).Value)) = docNumber And _
                   CStr(wsAdd47.Cells(lastRow47, 4).Value) = CStr(docDate) And _
                   Trim$(CStr(wsAdd47.Cells(lastRow47, 5).Value)) = docParties Then
                    isSameDoc = True
                End If
            End If
            
            If Not isSameDoc Then
                lastRow47 = lastRow47 + 1
                For j = 12 To lastColFinal
                    wsAdd47.Cells(lastRow47, j).Value = wsAdd47.Cells(lastRow47 - 1, j).Value
                Next j
                
                wsAdd47.Cells(lastRow47, 1).Value = docEntryDate
                wsAdd47.Cells(lastRow47, 2).Value = docType
                wsAdd47.Cells(lastRow47, 3).NumberFormat = "@"
                wsAdd47.Cells(lastRow47, 3).Value = docNumber
                wsAdd47.Cells(lastRow47, 4).Value = docDate
                wsAdd47.Cells(lastRow47, 5).Value = docParties
                wsAdd47.Cells(lastRow47, 6).Value = 0
                wsAdd47.Cells(lastRow47, 7).Value = 0
            End If
            
            Select Case typeOperation
                Case "надходження"
                    wsAdd47.Cells(lastRow47, 6).Value = wsAdd47.Cells(lastRow47, 6).Value + qty
                    If Not IsWrittenOffUnit(toUnit) Then
                        uColTo = GetOrCreateUnitColPNP(wsAdd47, toUnit, lastColFinal, lastRow47)
                        wsAdd47.Cells(lastRow47, uColTo + catTo).Value = wsAdd47.Cells(lastRow47, uColTo + catTo).Value + qty
                    End If
                    
                Case "вибуття"
                    wsAdd47.Cells(lastRow47, 7).Value = wsAdd47.Cells(lastRow47, 7).Value + qty
                    uColFrom = GetOrCreateUnitColPNP(wsAdd47, fromUnit, lastColFinal, lastRow47)
                    DeductFromUnitPNP wsAdd47, lastRow47, uColFrom, catFrom, qty
                    
                Case "внутрішнє переміщення"
                    If IsWrittenOffUnit(toUnit) Then
                        wsAdd47.Cells(lastRow47, 7).Value = wsAdd47.Cells(lastRow47, 7).Value + qty
                        uColFrom = GetOrCreateUnitColPNP(wsAdd47, fromUnit, lastColFinal, lastRow47)
                        DeductFromUnitPNP wsAdd47, lastRow47, uColFrom, catFrom, qty
                    Else
                        uColFrom = GetOrCreateUnitColPNP(wsAdd47, fromUnit, lastColFinal, lastRow47)
                        DeductFromUnitPNP wsAdd47, lastRow47, uColFrom, catFrom, qty
                        
                        uColTo = GetOrCreateUnitColPNP(wsAdd47, toUnit, lastColFinal, lastRow47)
                        wsAdd47.Cells(lastRow47, uColTo + catTo).Value = wsAdd47.Cells(lastRow47, uColTo + catTo).Value + qty
                    End If
                    
                Case Else
                    If IsWrittenOffUnit(toUnit) Then
                        wsAdd47.Cells(lastRow47, 7).Value = wsAdd47.Cells(lastRow47, 7).Value + qty
                        uColFrom = GetOrCreateUnitColPNP(wsAdd47, fromUnit, lastColFinal, lastRow47)
                        DeductFromUnitPNP wsAdd47, lastRow47, uColFrom, catFrom, qty
                    Else
                        wsAdd47.Cells(lastRow47, 6).Value = wsAdd47.Cells(lastRow47, 6).Value + qty
                        uColTo = GetOrCreateUnitColPNP(wsAdd47, toUnit, lastColFinal, lastRow47)
                        wsAdd47.Cells(lastRow47, uColTo + catTo).Value = wsAdd47.Cells(lastRow47, uColTo + catTo).Value + qty
                    End If
            End Select
        End If
        
        On Error Resume Next
        Call ProgressBar.Progress(n / UBound(filteredData, 1) * 100)
        On Error GoTo 0
    Next n
    
    Dim r As Long, c As Long, catK As Long, catSum As Double, uSum As Double
    For r = 21 To lastRow47
        For c = 12 To lastColFinal Step 4
            uSum = 0
            For catK = 1 To 3
                uSum = uSum + SafeVal(wsAdd47.Cells(r, c + catK).Value)
            Next catK
            wsAdd47.Cells(r, c).Value = uSum
        Next c
        
        For catK = 1 To 3
            catSum = 0
            For c = 12 To lastColFinal Step 4
                catSum = catSum + SafeVal(wsAdd47.Cells(r, c + catK).Value)
            Next c
            wsAdd47.Cells(r, 8 + catK).Value = catSum
        Next catK
        
        wsAdd47.Cells(r, 8).Value = wsAdd47.Cells(r, 9).Value + wsAdd47.Cells(r, 10).Value + wsAdd47.Cells(r, 11).Value
    Next r
    
    If lastRow47 >= 21 Then
        wsAdd47.Range(wsAdd47.Cells(20, 1), wsAdd47.Cells(20, 11)).Copy
        wsAdd47.Range(wsAdd47.Cells(21, 1), wsAdd47.Cells(lastRow47, 11)).PasteSpecial xlPasteFormats
        
        wsAdd47.Range(wsAdd47.Cells(20, 12), wsAdd47.Cells(20, 15)).Copy
        wsAdd47.Range(wsAdd47.Cells(21, 12), wsAdd47.Cells(lastRow47, lastColFinal)).PasteSpecial xlPasteFormats
    End If
    
    For c = 16 To lastColFinal
        wsAdd47.Cells(19, c).Value = c
    Next c
    
    wsAdd47.PageSetup.PrintArea = ""
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "Додаток 47 (пр-непр) успішно сформовано! Записів: " & (lastRow47 - 20), vbInformation, "Готово"
    Exit Sub

WrongPrice:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Ймовірно, ви не вказали ціну вибраного майна (введіть хоча б 1 грн)!", vbCritical, "Ціна не вказана"
End Sub

Private Function SafeVal(v As Variant) As Double
    Dim s As String
    s = Trim$(CStr(v))
    If Len(s) = 0 Then
        SafeVal = 0
        Exit Function
    End If
    s = Replace(s, " ", "")
    s = Replace(s, Chr(160), "")
    s = Replace(s, ",", ".")
    If IsNumeric(s) Then
        SafeVal = Val(s)
    Else
        SafeVal = 0
    End If
End Function

Private Function SafeDate(v As Variant) As Double
    Dim s As String
    s = Trim$(CStr(v))
    If Len(s) = 0 Then
        SafeDate = 0
        Exit Function
    End If
    If IsNumeric(s) Then
        SafeDate = Val(s)
    ElseIf IsDate(s) Then
        SafeDate = CDbl(CDate(s))
    Else
        SafeDate = 0
    End If
End Function

Private Function MapCatPNP(v As Variant) As Long
    ' 1 = придатне, 2 = потребує ремонту, 3 = непридатне
    Dim s As String
    s = LCase$(Trim$(CStr(v)))
    If Len(s) = 0 Then
        MapCatPNP = 1
        Exit Function
    End If
    If InStr(s, "не") > 0 Or s = "5" Then
        MapCatPNP = 3 ' непридатне
    ElseIf InStr(s, "рем") > 0 Or s = "3" Or s = "4" Then
        MapCatPNP = 2 ' потребує ремонту
    Else
        MapCatPNP = 1 ' придатне (1 або 2 категорія)
    End If
End Function

Private Function IsWrittenOffUnit(ByVal unitName As String) As Boolean
    Dim u As String
    u = LCase$(Trim$(unitName))
    If Len(u) = 0 Then
        IsWrittenOffUnit = False
        Exit Function
    End If
    If InStr(u, "спис") > 0 Or InStr(u, "знищ") > 0 Or InStr(u, "втрат") > 0 Or InStr(u, "утиліз") > 0 Then
        IsWrittenOffUnit = True
    Else
        IsWrittenOffUnit = False
    End If
End Function

Private Function GetOrCreateUnitColPNP(ws As Worksheet, ByVal unitName As String, ByRef lastCol As Long, ByVal currRow As Long) As Long
    Dim c As Long
    For c = 12 To lastCol Step 4
        If Trim$(CStr(ws.Cells(16, c).Value)) = Trim$(unitName) Then
            GetOrCreateUnitColPNP = c
            Exit Function
        End If
    Next c
    
    Dim newCol As Long
    newCol = lastCol + 1
    
    ws.Range(ws.Cells(16, 12), ws.Cells(20, 15)).Copy
    ws.Cells(16, newCol).PasteSpecial xlPasteAll
    ws.Cells(16, newCol).Value = unitName
    
    Dim rIdx As Long, k As Long
    For rIdx = 20 To currRow
        For k = 0 To 3
            ws.Cells(rIdx, newCol + k).Value = 0
        Next k
    Next rIdx
    
    If currRow >= 21 Then
        ws.Range(ws.Cells(20, 12), ws.Cells(20, 15)).Copy
        ws.Range(ws.Cells(21, newCol), ws.Cells(currRow, newCol + 3)).PasteSpecial xlPasteFormats
    End If
    
    lastCol = newCol + 3
    GetOrCreateUnitColPNP = newCol
End Function

Private Sub DeductFromUnitPNP(ws As Worksheet, ByVal r As Long, ByVal uCol As Long, ByVal preferredCat As Long, ByVal qty As Double)
    Dim currVal As Double
    If preferredCat >= 1 And preferredCat <= 3 Then
        currVal = SafeVal(ws.Cells(r, uCol + preferredCat).Value)
        If currVal >= qty Then
            ws.Cells(r, uCol + preferredCat).Value = currVal - qty
            Exit Sub
        ElseIf currVal > 0 Then
            ws.Cells(r, uCol + preferredCat).Value = 0
            qty = qty - currVal
        End If
    End If
    
    Dim order(1 To 3) As Long
    order(1) = 1: order(2) = 2: order(3) = 3
    Dim i As Long, cat As Long
    For i = 1 To 3
        cat = order(i)
        If cat <> preferredCat Then
            currVal = SafeVal(ws.Cells(r, uCol + cat).Value)
            If currVal >= qty Then
                ws.Cells(r, uCol + cat).Value = currVal - qty
                Exit Sub
            ElseIf currVal > 0 Then
                ws.Cells(r, uCol + cat).Value = 0
                qty = qty - currVal
            End If
        End If
    Next i
    
    If preferredCat < 1 Or preferredCat > 3 Then preferredCat = 1
    ws.Cells(r, uCol + preferredCat).Value = SafeVal(ws.Cells(r, uCol + preferredCat).Value) - qty
End Sub

Private Sub SortFilteredData(ByRef arr() As Variant)
    Dim numRows As Long
    numRows = UBound(arr, 1) - 1
    If numRows <= 1 Then Exit Sub
    
    Dim idx() As Long
    ReDim idx(1 To numRows)
    Dim i As Long
    For i = 1 To numRows
        idx(i) = i + 1
    Next i
    
    QuickSortRows arr, idx, 1, numRows
    
    Dim numCols As Long
    numCols = UBound(arr, 2)
    Dim temp() As Variant
    ReDim temp(1 To numRows + 1, 1 To numCols)
    
    Dim c As Long
    For c = 1 To numCols
        temp(1, c) = arr(1, c)
    Next c
    
    For i = 1 To numRows
        For c = 1 To numCols
            temp(i + 1, c) = arr(idx(i), c)
        Next c
    Next i
    
    arr = temp
End Sub

Private Sub QuickSortRows(ByRef arr() As Variant, ByRef idx() As Long, ByVal first As Long, ByVal last As Long)
    Dim i As Long, j As Long, pivotIdx As Long, temp As Long
    i = first
    j = last
    pivotIdx = idx((first + last) \ 2)
    
    Do While i <= j
        Do While CompareRows(arr, idx(i), pivotIdx) < 0
            i = i + 1
        Loop
        Do While CompareRows(arr, idx(j), pivotIdx) > 0
            j = j - 1
        Loop
        If i <= j Then
            temp = idx(i)
            idx(i) = idx(j)
            idx(j) = temp
            i = i + 1
            j = j - 1
        End If
    Loop
    
    If first < j Then QuickSortRows arr, idx, first, j
    If i < last Then QuickSortRows arr, idx, i, last
End Sub

Private Function CompareRows(ByRef arr() As Variant, ByVal a As Long, ByVal b As Long) As Long
    Dim dateA As Double, dateB As Double
    dateA = SafeDate(arr(a, 13))
    If dateA = 0 Then dateA = SafeDate(arr(a, 1))
    
    dateB = SafeDate(arr(b, 13))
    If dateB = 0 Then dateB = SafeDate(arr(b, 1))
    
    If dateA < dateB Then
        CompareRows = -1
        Exit Function
    ElseIf dateA > dateB Then
        CompareRows = 1
        Exit Function
    End If
    
    Dim docA As String, docB As String
    docA = Trim$(CStr(arr(a, 14)))
    docB = Trim$(CStr(arr(b, 14)))
    
    If IsNumeric(docA) And IsNumeric(docB) Then
        If Val(docA) < Val(docB) Then
            CompareRows = -1
            Exit Function
        ElseIf Val(docA) > Val(docB) Then
            CompareRows = 1
            Exit Function
        End If
    Else
        Dim comp As Long
        comp = StrComp(docA, docB, vbTextCompare)
        If comp <> 0 Then
            CompareRows = comp
            Exit Function
        End If
    End If
    
    If a < b Then
        CompareRows = -1
    ElseIf a > b Then
        CompareRows = 1
    Else
        CompareRows = 0
    End If
End Function

Sub SaveAdd47PNP()
    ActiveSheet.Select
    ActiveSheet.Copy
    
    Dim WbNewWO As Workbook
    Set WbNewWO = ActiveWorkbook
    
    Dim nameAdd47 As String
    Dim nomenclature47 As String
    Dim unit47 As String
    Dim price47 As Double
    Dim formattedPrice As String
    
    On Error Resume Next
    nameAdd47 = WbNewWO.Worksheets(1).Range("nameAdd47").Value
    nomenclature47 = WbNewWO.Worksheets(1).Range("nomenclature47").Value
    unit47 = WbNewWO.Worksheets(1).Range("unit47").Value
    price47 = WbNewWO.Worksheets(1).Range("price47").Value
    On Error GoTo 0
    
    formattedPrice = Format(price47, "#,##0.0000")

    Dim shp As Shape
    For Each shp In WbNewWO.Worksheets(1).Shapes
        shp.Delete
    Next shp
        
    Dim fileName As String, SavePath As String
    SavePath = ThisWorkbook.Sheets("Додаток 47 пр-непр").Range("PathToSaveAdd47").Value
    Dim DateAdd47 As String
    DateAdd47 = Format(Date, "dd.mm.yyyy")
    
    If Len(SavePath) = 0 Or Dir(SavePath, vbDirectory) = vbNullString Then
        With Application.FileDialog(msoFileDialogFolderPicker)
            .Title = "Виберіть папку для збереження Додатка 47 пр-непр"
            .AllowMultiSelect = False
            If .Show = -1 Then
                SavePath = .SelectedItems(1)
                ThisWorkbook.Sheets("Додаток 47 пр-непр").Range("PathToSaveAdd47").Value = SavePath
            Else
                MsgBox "Збереження скасовано", vbExclamation
                WbNewWO.Close SaveChanges:=False
                Exit Sub
            End If
        End With
    End If
        
    fileName = SavePath & Application.PathSeparator & CleanFileName(DateAdd47 & " д47_пр_непр " & nameAdd47 & " " & nomenclature47 & " " & unit47 & " " & formattedPrice & ".xlsx")
    
    Application.DisplayAlerts = False
    On Error GoTo WrongPath
    WbNewWO.SaveAs fileName:=fileName, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    On Error GoTo 0
    MsgBox "Файл успішно збережено: " & vbCrLf & fileName, vbInformation, "Збережено"
    Exit Sub

WrongPath:
    Application.DisplayAlerts = True
    MsgBox "Помилка адреси збереження файлу. Перевірте шлях збереження!", vbCritical, "Помилка збереження"
End Sub
