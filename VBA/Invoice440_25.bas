Attribute VB_Name = "Invoice440_25"
Sub addPosInv440_25()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Накладна440_25")
    
    Dim NameOfThings As String

    Dim tbInv440_25 As ListObject
    Set tbInv440_25 = ws.ListObjects("tbInv440_25")
    
    
    
'=============3.9.15
If IsError(ws.Range("FromUnitInv440_25").value) Then
    MsgBox "Схоже ви не заповнили довідник і/або не вибрали хто буде передавати(здавати, видавати) майно." _
    & vbNewLine & vbNewLine & "Заповінсть довідник і виберіть, хто буде ПЕРДАВАТИ майно.", _
    vbCritical, "Заповніть довідник, виберіть хто буде передавати"
    
    ws.Range("MRPInv440_25").Activate
    Exit Sub
End If

If IsError(ws.Range("toUnitInv440_25").value) Then
    MsgBox "Схоже ви не заповнили довідник і/або не вибрали хто буде отримувати (приймати) майно." _
    & vbNewLine & vbNewLine & "Заповінсть довідник і виберіть, хто буде ОТРИМУВАТИ майно.", _
    vbCritical, "Заповніть довідник, виберіть отримувачів"
    
    ws.Range("MRPToInv440_25").Activate
    Exit Sub
End If
'=============3.9.15
    
'=============3.9.11
'визначаємо підрозділ-відправник, беремо з документа
Dim FromUnit As String
FromUnit = ws.Range("FromUnitInv440_25")

Dim tempPoolOThings As Range
Set tempPoolOThings = GetResidusePool(wsResiduse2, FromUnit)

If tempPoolOThings Is Nothing Then Exit Sub
'=============3.9.11

''Визначаємо останній рядок з даними
'Dim lastRow As Long
'    lastRow = wsResiduse2.Cells(wsResiduse2.Rows.count, "N").End(xlUp).Row
'    If lastRow < 3 Then
'        MsgBox "У залишках немає даних.", vbExclamation
'        Exit Sub
'    End If
'
''пропонуємо, де шукати по цьому підрозділу, чи по всій базі
'AllOrOneUnit = MsgBox("Шукати лише в підрозділі: " & vbNewLine & FromUnit & " - так" & vbNewLine & _
'"шукати по всій базі - ні", vbYesNo, "Де шукати потрібне майно?")
'
''Приймаємо відповідь від користувача
''в залежності від цього формуємо діапазон для пошуку
'Dim tempPoolOThings As Range
'Dim Residuse As Range
'Set Residuse = wsResiduse2.Range("N3:N" & lastRow)
''Костиль, щоб передавався завжди масив, а не значення (коли тільки одне значення)
'Dim ResiduseCheet As Range
''тому робимо діапазон на одну клітинку більше
'Set ResiduseCheet = wsResiduse2.Range("N3:N" & lastRow + 1)
'
'
'If AllOrOneUnit = vbYes Then
'
''    накласти фільтр на залишки по підрозділах 2
'    Dim hdr As Range
'    Set hdr = wsResiduse2.Range("A2").CurrentRegion
'
'    On Error Resume Next
'    If wsResiduse2.FilterMode Then wsResiduse2.ShowAllData
'    On Error GoTo 0
'
'    hdr.AutoFilter Field:=12, Criteria1:=FromUnit
'
'
''забрати в масив тільки видимі клітинки
'    Set tempPoolOThings = ResiduseCheet.SpecialCells(xlCellTypeVisible)
'
'    If Application.WorksheetFunction.CountA(tempPoolOThings) = 0 Then
'    MsgBox "По підрозділу '" & FromUnit & "' немає позицій.", vbExclamation
'    Call ClearUnitFilter_Residues(wsResiduse2)
'    Exit Sub
'    End If
'
'    'зняти фільтр
'    Call ClearUnitFilter_Residues(wsResiduse2)
'Else
'     Set tempPoolOThings = ResiduseCheet
'End If
''=============3.9.11


 Dim Basis As String
 Basis = InputBox("Підстава (мета) операції: ", NameOfThings, ws.Range("basis440_25").value)
 ws.Range("basis440_25") = Basis
    Do
        
        Dim frm     As UserForm1
        Set frm = UserForms.Add(UserForm1.name)
 
         frm.ListData = tempPoolOThings
        
        frm.show
        
        NameOfThings = frm.TextBox1.value
        
        Dim rngWithName As Range
        
        Dim ListGoods As Range
        Set ListGoods = wsResiduse2.Range("Goods")
        'Шукаємо задане значення в діапазоні
        Set rngWithName = ListGoods.Find(NameOfThings, LookIn:=xlValues, LookAt:=xlWhole)
        
        'Перевіряємо, чи знайдено задане значення
        If Not rngWithName Is Nothing Then
            Dim name As String
            Dim Nomenclature As String
            Dim unit As String
            Dim Category As String
            Dim NoteKt As String
            Dim price As Double
            Dim VAT As Double
            Dim QuantityRest As Double
            Dim NoteDt As String
            
'            Dim Place As String
            Dim CategoryFrom As String
            Dim CategoryTo As String
            Dim Note1 As String
            Dim Note2 As String
            Dim Note3 As String
            'Визначаємо значення за замовчуванням
            name = rngWithName.Offset(, -12).value
            Nomenclature = rngWithName.Offset(, -6).value
            unit = rngWithName.Offset(, -7).value
'            category = rngWithName.Offset(, -10).value
            Note = rngWithName.Offset(, -9).value
            price = rngWithName.Offset(, -11).value
            QuantityRest = rngWithName.Offset(, -10).value
'            Place = rngWithName.Offset(, -3).value
            CategoryFrom = rngWithName.Offset(, -8).value
            
'==============
NextIteration:
'==============

            quantity = InputBox("Кількість: ", NameOfThings, QuantityRest)
            
'========
'If quantity = "" Then Exit Sub
            
            If Len(Trim$(CStr(quantity))) = 0 Then GoTo NextIteration
            If quantity <= 0 Then
                MsgBox "Кількість має бути більше 0.", vbExclamation
                GoTo NextIteration
            End If
            If Not IsNumeric(quantity) Then
                MsgBox "Кількість має бути числом.", vbExclamation
                GoTo NextIteration
            End If
            
'========

'            NoteDt = InputBox("Субрахунок для списання, залишити 1 (якщо не знаєте який - залишіть порожнім): ", NameOfThings, "6611 8013 8113")
                  
            
            
            Category = InputBox("Категорія відправника: ", NameOfThings, rngWithName.Offset(, -8).value)
            
            CategoryTo = InputBox("Категорія отримувача: ", NameOfThings, rngWithName.Offset(, -8).value)
'            VAT = InputBox("ПДВ (за наявності): ", NameOfThings, 0)
'            Note1 = InputBox("Додаткове поле 1", NameOfThings, rngWithName.Offset(, -5).value)
            Note2 = InputBox("Додаткове поле 2", NameOfThings, rngWithName.Offset(, -4).value)
            Note3 = InputBox("Додаткове поле 3", NameOfThings, rngWithName.Offset(, -3).value)
            If tbInv440_25.DataBodyRange.Rows(1).Cells(1, 2) = "" Then
                With tbInv440_25.DataBodyRange.Rows
                    .Cells(1, 2) = name
                    .Cells(1, 3) = Nomenclature
                    .Cells(1, 4) = unit
                    .Cells(1, 5) = Category
                    .Cells(1, 6) = price
                    .Cells(1, 7) = CDbl(quantity)
                    .Cells(1, 10) = Note
                    .Cells(1, 11) = CategoryTo
                    
'                    .Cells(1, 12) = Note1
                    .Cells(1, 13) = Note2
                    .Cells(1, 14) = Note3
                End With
            Else
                    Dim newRow As ListRow
                    Set newRow = tbInv440_25.ListRows.Add
                With newRow.Range
                    .Cells(1, 2) = name
                    .Cells(1, 3) = Nomenclature
                    .Cells(1, 4) = unit
                    .Cells(1, 5) = Category
                    .Cells(1, 6) = price
                    .Cells(1, 7) = CDbl(quantity)
                    .Cells(1, 10) = Note
                    .Cells(1, 11) = CategoryTo
                    
'                    .Cells(1, 12) = Note1
                    .Cells(1, 13) = Note2
                    .Cells(1, 14) = Note3
                    
                End With
            End If
            
            Dim AddNext
            
            AddNext = MsgBox("Майно додано! Додати наступну позицію?", vbYesNo, NameOfThings)
            
        End If
        
    Loop While AddNext = 6
    

End Sub

Sub clearTable440_25()
Dim tbActual As ListObject
Set tbActual = ThisWorkbook.Worksheets("Накладна440_25").ListObjects("tbInv440_25")
Dim i As Long
If tbActual.ListRows.count > 1 Then
                ' Видаляємо всі рядки, крім першого
                For i = tbActual.ListRows.count To 2 Step -1
                    tbActual.ListRows(i).Delete
                Next i
End If

Dim firstRow As ListRow
Set firstRow = tbActual.ListRows(1)
Dim j As Long
 For j = 1 To tbActual.ListColumns.count
        If j <> 1 And j <> 8 And j <> 9 Then
            firstRow.Range.Cells(1, j).ClearContents
        End If
 Next j
 
End Sub


Sub addRow440_25()
Dim tbActual As ListObject
Set tbActual = ThisWorkbook.Worksheets("Накладна440_25").ListObjects("tbInv440_25")

Dim newRow As ListRow
Set newRow = tbActual.ListRows.Add

End Sub

'Sub SaveInv440_25()
'Application.DisplayAlerts = False
' Dim typeOperation As String
'
'   typeOperation = InputBox("Тип операції" & vbNewLine & "надходження - 1" _
'                                           & vbNewLine & "внутрішнє переміщення - 2" _
'                                           & vbNewLine & "вибуття - 3", "Тип операції", 2)
'
'   Select Case typeOperation
'   Case 1
'   ThisWorkbook.Sheets("Накладна440_25").Range("ContentInv440_25") = "надходження"
'   Case 2
'    ThisWorkbook.Sheets("Накладна440_25").Range("ContentInv440_25") = "внутрішнє переміщення"
'   Case 3
'    ThisWorkbook.Sheets("Накладна440_25").Range("ContentInv440_25") = "вибуття"
'   End Select
'
''збільшуємо номер документа на 1
'Dim clearNumber As Long
'
'clearNumber = ThisWorkbook.Sheets("Накладна440_25").Range("clearNumberInv440_25").value
'
'ThisWorkbook.Sheets("Накладна440_25").Range("clearNumberInv440_25") = clearNumber + 1
'
''забираємо суму і кількість
'Dim allCount As String
'Dim allSum As String
'
'allCount = ThisWorkbook.Sheets("Накладна440_25").Range("allCountInv440_25").value
'allSum = ThisWorkbook.Sheets("Накладна440_25").Range("allSumInv440_25").value
'
'        ThisWorkbook.Sheets("Накладна440_25").Select 'копіюємо перший лист з нарядом
'        ThisWorkbook.Sheets("Накладна440_25").Copy 'зберігаємо як нову книгу
'
'     Dim WbNew As Workbook
'
'        Set WbNew = ActiveWorkbook
'
'     'замінюємо на значення клітинки, які підтягуються по формулі з іншого листа, щоб у збережених документах не злітали дані
'
'      For Each cell In WbNew.Worksheets(1).Range(Cells(1, 1), Cells(WbNew.Worksheets(1).Range("bookkeepingInv440_25").Row, 16))
'        If cell.HasFormula Then
'            cell.value = cell.value
'        End If
'     Next cell
'
'    WbNew.Worksheets(1).Range("allCountInv440_25") = allCount
'    WbNew.Worksheets(1).Range("allSumInv440_25") = allSum
'
'   Dim NumberInv440_25 As String
'     NumberInv440_25 = WbNew.Worksheets(1).Range("NumberInv440_25")
'
'   Dim fromUnit As String
'   Dim toUnit As String
'        fromUnit = WbNew.Worksheets(1).Range("FromUnitInv440_25")
'        toUnit = WbNew.Worksheets(1).Range("toUnitInv440_25")
'    Dim shp As Shape
'    For Each shp In WbNew.Worksheets(1).Shapes
'        shp.Delete
'    Next shp
'
'      Dim fileName As String, fileName2 As String, lineBrake As String
'        'Шлях до файла зберігання
'        SavePath = WbNew.Worksheets(1).Range("PathToSaveInv440_25")
'        Dim DateInv440_25 As Date
'        DateInv440_25 = WbNew.Worksheets(1).Range("DateInv440_25")
'        'складаємо назву файла
'        fileName = SavePath & Application.PathSeparator & CleanFileName(DateInv440_25 & _
'        " Накладна(вим.)№ " & NumberInv440_25 & " " & fromUnit & " - " & toUnit & ".xlsx")
'        'заміняємо перенесення рядка, косі лінії
'
'
'          'зберігаємо файл
'        On Error GoTo WrongPath
'         WbNew.SaveAs fileName:=fileName, FileFormat:=xlOpenXMLWorkbook
'     On Error GoTo 0
'     Application.DisplayAlerts = True
'      Exit Sub
'WrongPath:
'
'
'
'      MsgBox "Ймовірно, не правильно вказана адреса збереження файла, виправте адресу збереження файлу( на зразок ""D:\База обліку майна ЗСУ 2024"" *без косої лінії вкінці) і спробуйтие зберегти файл ще раз!", vbCritical, "Не правильний шлях збереження"
'     Application.DisplayAlerts = True
'
'        'якщо у вас тут помилка, ймовіно ви не змінили шлях збереження файлу, або допустити помилку в шляху збереження (наприклад лишній знак "/" вкінці)
'
'End Sub


Sub SaveInv440_25()
    Application.DisplayAlerts = False
    Dim typeOperation As String
    
    typeOperation = InputBox("Тип операції" & vbNewLine & "надходження - 1" _
                                            & vbNewLine & "внутрішнє переміщення - 2" _
                                            & vbNewLine & "вибуття - 3", "Тип операції", 2)
    
    Select Case typeOperation
        Case 1
            ThisWorkbook.Sheets("Накладна440_25").Range("ContentInv440_25") = "надходження"
        Case 2
            ThisWorkbook.Sheets("Накладна440_25").Range("ContentInv440_25") = "внутрішнє переміщення"
        Case 3
            ThisWorkbook.Sheets("Накладна440_25").Range("ContentInv440_25") = "вибуття"
    End Select
    
    
     'запитуємо додатковий тип операції
   typeOperationAdd = InputBox("Детальний тип операції" & vbNewLine & "надійшло з інших в/ч (підрозділів) - 1" _
                                           & vbNewLine & "надійшло зі складу в/ч - 2" _
                                           & vbNewLine & "інші надходження - 3" _
                                           & vbNewLine & "витрата - 4" _
                                           & vbNewLine & "передано до інших в/ч (підрозділів)- 5" _
                                           & vbNewLine & "інші вибуття - 6", "Тип операції", 5)
Select Case typeOperationAdd
   Case 1
   ThisWorkbook.ActiveSheet.Range("ContentInvAdd440_25") = "надійшло з інших в/ч (підрозділів)"
   Case 2
    ThisWorkbook.ActiveSheet.Range("ContentInvAdd440_25") = "надійшло зі складу в/ч"
   Case 3
    ThisWorkbook.ActiveSheet.Range("ContentInvAdd440_25") = "інші надходження"
    Case 4
    ThisWorkbook.ActiveSheet.Range("ContentInvAdd440_25") = "витрата"
    Case 5
    ThisWorkbook.ActiveSheet.Range("ContentInvAdd440_25") = "передано до інших в/ч (підрозділів)"
    Case 6
    ThisWorkbook.ActiveSheet.Range("ContentInvAdd440_25") = "інші вибуття"
   End Select
   
    
    
    
    
    
    
    
    'збільшуємо номер документа на 1
    Dim clearNumber As Long
    clearNumber = ThisWorkbook.Sheets("Накладна440_25").Range("clearNumberInv440_25").value
    ThisWorkbook.Sheets("Накладна440_25").Range("clearNumberInv440_25") = clearNumber + 1
    
    'забираємо суму і кількість
    Dim allCount As String, allSum As String
    allCount = ThisWorkbook.Sheets("Накладна440_25").Range("allCountInv440_25").value
    allSum = ThisWorkbook.Sheets("Накладна440_25").Range("allSumInv440_25").value
    
    ThisWorkbook.Sheets("Накладна440_25").Copy 'зберігаємо як нову книгу
    Dim WbNew As Workbook
    Set WbNew = ActiveWorkbook
    
    'замінюємо формули на значення
    Dim cell As Range
    For Each cell In WbNew.Worksheets(1).Range(Cells(1, 1), Cells(WbNew.Worksheets(1).Range("bookkeepingInv440_25").Row, 16))
        If cell.HasFormula Then
            cell.value = cell.value
        End If
    Next cell
    
    WbNew.Worksheets(1).Range("allCountInv440_25") = allCount
    WbNew.Worksheets(1).Range("allSumInv440_25") = allSum
    
    'реквізити для назви кінцевого файлу
    Dim NumberInv440_25 As String, FromUnit As String, toUnit As String, fromMRP As String, toMRP As String
    NumberInv440_25 = WbNew.Worksheets(1).Range("NumberInv440_25")
    FromUnit = WbNew.Worksheets(1).Range("FromUnitInv440_25")
    toUnit = WbNew.Worksheets(1).Range("toUnitInv440_25")
    '=================3.9.11
    fromMRP = WbNew.Worksheets(1).Range("MRPInv440_25")
    toMRP = WbNew.Worksheets(1).Range("MRPToInv440_25")
    '=================
    
    Dim shp As Shape
    For Each shp In WbNew.Worksheets(1).Shapes
        shp.Delete
    Next shp
    
    Dim fileName As String, SavePath As String
    Dim DateInv440_25 As Date
    DateInv440_25 = WbNew.Worksheets(1).Range("DateInv440_25")
    
    'Спочатку беремо шлях з комірки
    SavePath = WbNew.Worksheets(1).Range("PathToSaveInv440_25").value
    
    'Перевіряємо чи існує папка, якщо ні – відкриваємо діалог вибору
    On Error Resume Next
    If Dir(SavePath, vbDirectory) = vbNullString Then
    On Error GoTo 0
        With Application.FileDialog(msoFileDialogFolderPicker)
            .Title = "Виберіть папку для збереження"
            .AllowMultiSelect = False
            If .show = -1 Then
                SavePath = .SelectedItems(1)
                'Записуємо новий шлях у клітинку
                ThisWorkbook.Sheets("Накладна440_25").Range("PathToSaveInv440_25").value = SavePath
            Else
                MsgBox "Збереження скасовано", vbExclamation
                WbNew.Close SaveChanges:=False
                Application.DisplayAlerts = True
                Exit Sub
            End If
        End With
    End If
    
    'Формуємо ім'я файлу
    '3.9.11========
    fileName = SavePath & Application.PathSeparator & CleanFileName(DateInv440_25 & _
        " Накладна(вим.)№ " & NumberInv440_25 & " " & FromUnit & " (" & fromMRP & ") - " & toUnit & " (" & toMRP & ").xlsx")
    '========
    'Зберігаємо
    WbNew.SaveAs fileName:=fileName, FileFormat:=xlOpenXMLWorkbook
    
    Application.DisplayAlerts = True
End Sub


'============3.9.16=========
Sub addFromTableInv440_25()

'Dim lo As ListObject
'Set lo = Worksheets("База").ListObjects("AllData")

Dim ws As Worksheet
Set ws = Selection.Worksheet 'лист, де був клік

Dim wsDoc As Worksheet
Set wsDoc = ThisWorkbook.Worksheets("Накладна440_25")
Dim tbActual As ListObject

Set tbActual = wsDoc.ListObjects("tbInv440_25")

Dim NameOfThings As String
Dim Nomenclature As String
Dim unit As String
Dim Category As String

Dim price As Double

Dim QuantityRest As Double
Dim quantity As Double

Dim CategoryFrom As String
Dim CategoryTo As String
Dim Note As String
Dim Note1 As String
Dim Note2 As String
Dim Note3 As String

Dim MRP As String

Dim rowsArr

rowsArr = GetSelectedRowsArrayCell()

Dim colName As Long
Dim colQty1 As Long
Dim colQty2 As Long
Dim colNomenclature As Long
Dim colUnit As Long
Dim colCategory As Long
Dim colPrice As Long
Dim colNote As Long
Dim colMRP As Long
Dim colNote2 As Long
Dim colNote3 As Long

Select Case ws.name

    Case "База"
        colName = 2
        colQty1 = 6
        colQty2 = 7
        colNomenclature = 18
        colUnit = 4
        colCategory = 27
        colPrice = 20
        colNote = 16
        colMRP = 11
        colNote2 = 23
        colNote3 = 24

    Case "Залишки по підрозділах2"
        colName = 2
        colQty1 = 4
'        colQty2 = 4
        colNomenclature = 8
        colUnit = 7
        colCategory = 6
        colPrice = 3
        colNote = 5
'        colMRP = 7
        colNote2 = 10
        colNote3 = 11
        
        Case "Довідка по підрозділу"
        colName = 3
        colQty1 = 5
'        colQty2 = 4
        colNomenclature = 9
        colUnit = 8
        colCategory = 7
        colPrice = 4
        colNote = 6
'        colMRP = 7
'        colNote2 = 10
'        colNote3 = 11

    Case "Довідка по людині"
        colName = 2
        colQty1 = 4
'        colQty2 = 4
        colNomenclature = 8
        colUnit = 7
        colCategory = 6
        colPrice = 3
        colNote = 5
'        colMRP = 7
'        colNote2 = 10
'        colNote3 = 11
    Case Else
        MsgBox "Невідомий лист: " & ws.name, vbCritical
        Exit Sub

End Select

Dim qQuantity As VbMsgBoxResult
Dim qKategory As VbMsgBoxResult
qQuantity = MsgBox("Запитувати кількість кожного найменування?(Натисніть - так)" & vbNewLine & _
"Взяти дані кількість з поточного листа.(Натисніть - ні)", vbYesNo, "Кількість кожного найменування?")

qKategory = MsgBox("Запитувати категорію кожного найменування?(Натисніть - так)" & vbNewLine & _
"Категорія не буде змінюватись, тому можна взяти з поточного листа.(Натисніть - ні)", vbYesNo, "Категорія кожного найменування?")


Dim listOfThings As String
 Dim rowIndex
  Dim rw As Range

    For Each rowIndex In rowsArr
        Set rw = Rows(rowIndex)
            NameOfThings = rw.Cells(1, colName)
            
            If colQty2 = 0 Then
                QuantityRest = rw.Cells(1, colQty1)
            Else
                QuantityRest = rw.Cells(1, colQty1) + rw.Cells(1, colQty2)
            End If
            
            If qQuantity = vbYes Then
                quantity = InputBox("Кількість: ", NameOfThings, QuantityRest)
            Else
                quantity = QuantityRest
            End If
            
            Nomenclature = rw.Cells(1, colNomenclature)
            unit = rw.Cells(1, colUnit)
            
            If qKategory = vbYes Then
                Category = InputBox("Категорія відправника:", NameOfThings, rw.Cells(1, colCategory))
                CategoryTo = InputBox("Категорія отримувача:", NameOfThings, rw.Cells(1, colCategory))
            Else
                Category = rw.Cells(1, colCategory)
                CategoryTo = rw.Cells(1, colCategory)
            End If
            
            
            
            price = rw.Cells(1, colPrice)
            Note = rw.Cells(1, colNote)
            
            If colNote2 = 0 Then
            Else
                Note2 = rw.Cells(1, colNote2)
            End If
            
            If colNote3 = 0 Then
            Else
                Note3 = rw.Cells(1, colNote3)
            End If
            
            
            If colMRP = 0 Then
            Else
                MRP = rw.Cells(1, colMRP)
                wsDoc.Range("MRPInv440_25") = MRP
            End If
            
            
            
            
        wsDoc.Range("MRPInv440_25") = MRP
    
        'додаємо інформацію в документ
            If tbActual.DataBodyRange.Rows(1).Cells(1, 2) = "" Then
                With tbActual.DataBodyRange.Rows
                    .Cells(1, 2) = NameOfThings
                    .Cells(1, 3) = Nomenclature
                    .Cells(1, 4) = unit
                    .Cells(1, 5) = Category
                    .Cells(1, 6) = price
                    .Cells(1, 7) = CDbl(quantity)
                    .Cells(1, 10) = Note
                    .Cells(1, 11) = CategoryTo
                    
'                    .Cells(1, 12) = Note1
                    .Cells(1, 13) = Note2
                    .Cells(1, 14) = Note3
                End With
            Else
                    Dim newRow As ListRow
                    Set newRow = tbActual.ListRows.Add
                With newRow.Range
                    .Cells(1, 2) = NameOfThings
                    .Cells(1, 3) = Nomenclature
                    .Cells(1, 4) = unit
                    .Cells(1, 5) = Category
                    .Cells(1, 6) = price
                    .Cells(1, 7) = CDbl(quantity)
                    .Cells(1, 10) = Note
                    .Cells(1, 11) = CategoryTo
                    
'                    .Cells(1, 12) = Note1
                    .Cells(1, 13) = Note2
                    .Cells(1, 14) = Note3
                    
                End With
            End If

    listOfThings = listOfThings & NameOfThings & " " & CDbl(quantity) & " " & unit & vbNewLine
    
    
    
    Next rowIndex

MsgBox listOfThings, vbInformation, "Майно додане в Накладну д.25"

On Error Resume Next
    Application.CommandBars("Cell").Controls("Додати в документ").Delete
On Error GoTo 0


End Sub
'============3.9.16=========
