Attribute VB_Name = "ActChangeState81_14"

'ActChangeState81_14

Sub addPosActChangeState81_14()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Акт зм. якіс. ст. д14 81")
    
    Dim NameOfThings As String

    Dim tbActChangeState81_14 As ListObject
    Set tbActChangeState81_14 = ws.ListObjects("tbActChangeState81_14")
    
    
    
'=============3.9.15
If IsError(ws.Range("FromUnitActChangeState81_14").value) Then
    MsgBox "Схоже ви не заповнили довідник і/або не вибрали хто буде передавати(здавати, видавати) майно." _
    & vbNewLine & vbNewLine & "Заповінсть довідник і виберіть, хто буде ПЕРДАВАТИ майно.", _
    vbCritical, "Заповніть довідник, виберіть хто буде передавати"
    
    ws.Range("MRPActChangeState81_14").Activate
    Exit Sub
End If

'If IsError(ws.Range("toUnitActChangeState81_14").value) Then
'    MsgBox "Схоже ви не заповнили довідник і/або не вибрали хто буде отримувати (приймати) майно." _
'    & vbNewLine & vbNewLine & "Заповінсть довідник і виберіть, хто буде ОТРИМУВАТИ майно.", _
'    vbCritical, "Заповніть довідник, виберіть отримувачів"
'
'    ws.Range("MRPToActChangeState81_14").Activate
'    Exit Sub
'End If
'=============3.9.15

'=============3.9.19
Dim answer As VbMsgBoxResult

answer = MsgBox("Додати в частину списання - 'Так'" & vbNewLine & "Додати в частину оприбуткування - 'Ні'", vbYesNo + vbQuestion, "Підтвердження")

'=============3.9.19

    
'=============3.9.11
'визначаємо підрозділ-відправник, беремо з документа
Dim FromUnit As String
FromUnit = ws.Range("FromUnitActChangeState81_14")

Dim tempPoolOThings As Range
Set tempPoolOThings = GetResidusePool(wsResiduse2, FromUnit)

If tempPoolOThings Is Nothing Then Exit Sub


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
            
            
            Dim newRow As ListRow
            If answer = vbYes Then
            For i = 1 To tbActChangeState81_14.ListRows.count
                    If tbActChangeState81_14.DataBodyRange.Cells(i, 18).value = "" Then
                        Set newRow = tbActChangeState81_14.ListRows(i)
                        Exit For
                    End If
                Next i
                ' Якщо порожній рядок не знайдено – додаємо новий
                If newRow Is Nothing Then
                    Set newRow = tbActChangeState81_14.ListRows.Add
                End If
            
   
'            If tbActChangeState81_14.DataBodyRange.Rows(1).Cells(1, 18) = "" Then
'                With tbActChangeState81_14.DataBodyRange.Rows
'                    .Cells(1, 18) = name
'                    .Cells(1, 3) = Nomenclature
'                    .Cells(1, 4) = unit
'                    .Cells(1, 5) = Category
'                    .Cells(1, 7) = Price
'                    .Cells(1, 6) = CDbl(quantity)
'                    .Cells(1, 19) = Note
'                    .Cells(1, 21) = CategoryTo
'
'                    .Cells(1, 23) = Note2
'                    .Cells(1, 24) = Note3
'                End With
'            Else
'                    Dim NewRow As ListRow
'                    Set NewRow = tbActChangeState81_14.ListRows.Add
                With newRow.Range
                    .Cells(1, 18) = name
                    .Cells(1, 3) = Nomenclature
                    .Cells(1, 4) = unit
                    .Cells(1, 5) = Category
                    .Cells(1, 7) = price
                    .Cells(1, 6) = CDbl(quantity)
                    .Cells(1, 19) = Note
                    .Cells(1, 21) = CategoryTo
                    
                    .Cells(1, 23) = Note2
                    .Cells(1, 24) = Note3
                    
                End With
'            End If
            
            Else
                For i = 1 To tbActChangeState81_14.ListRows.count
                    If tbActChangeState81_14.DataBodyRange.Cells(i, 11).value = "" Then
                        Set newRow = tbActChangeState81_14.ListRows(i)
                        Exit For
                    End If
                Next i
                ' Якщо порожній рядок не знайдено – додаємо новий
                If newRow Is Nothing Then
                    Set newRow = tbActChangeState81_14.ListRows.Add
                End If

                   
                
                With newRow.Range
                    .Cells(1, 11) = name
                    .Cells(1, 12) = Nomenclature
                    .Cells(1, 13) = unit
                    .Cells(1, 14) = Category
                    .Cells(1, 16) = price
                    .Cells(1, 15) = CDbl(quantity)
                    .Cells(1, 20) = Note
                    .Cells(1, 22) = CategoryTo
                    
                    .Cells(1, 25) = Note2
                    .Cells(1, 26) = Note3
                    
                End With
            
    
    
    
    
            End If
            
            Dim AddNext
            
            AddNext = MsgBox("Майно додано! Додати наступну позицію?", vbYesNo, NameOfThings)
            
        End If
        
    Loop While AddNext = 6
    
    
    '***********Для майна, яке оприбутковується***********
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    

End Sub

Sub clearTableActChangeState81_14()
Dim tbActual As ListObject
Set tbActual = ThisWorkbook.Worksheets("Акт зм. якіс. ст. д14 81").ListObjects("tbActChangeState81_14")
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
        If j <> 1 And j <> 2 And j <> 8 And j <> 17 Then
            firstRow.Range.Cells(1, j).ClearContents
        End If
 Next j
 
End Sub


Sub addRowActChangeState81_14()
Dim tbActual As ListObject
Set tbActual = ThisWorkbook.Worksheets("Акт зм. якіс. ст. д14 81").ListObjects("tbActChangeState81_14")

Dim newRow As ListRow
Set newRow = tbActual.ListRows.Add

End Sub




Sub SaveActChangeState81_14()
    Application.DisplayAlerts = False
    
 Dim Basis As String
 Basis = InputBox("Вкажіть на підставі якого документу або події оформлюється даний документ ", "Підстава (мета) операції:", ThisWorkbook.ActiveSheet.Range("basisActChangeState81_14").value)
 ThisWorkbook.ActiveSheet.Range("basisActChangeState81_14") = Basis
    
    
    
    Dim typeOperation As String
    
    typeOperation = InputBox("Тип операції" & vbNewLine & "надходження - 1" _
                                            & vbNewLine & "внутрішнє переміщення - 2" _
                                            & vbNewLine & "вибуття - 3", "Тип операції", 3)
    
    Select Case typeOperation
        Case 1
            ThisWorkbook.Sheets("Акт зм. якіс. ст. д14 81").Range("ContentActChangeState81_14") = "надходження"
        Case 2
            ThisWorkbook.Sheets("Акт зм. якіс. ст. д14 81").Range("ContentActChangeState81_14") = "внутрішнє переміщення"
        Case 3
            ThisWorkbook.Sheets("Акт зм. якіс. ст. д14 81").Range("ContentActChangeState81_14") = "вибуття"
    End Select
    
    
     'запитуємо додатковий тип операції
   typeOperationAdd = InputBox("Детальний тип операції" & vbNewLine & "надійшло з інших в/ч (підрозділів) - 1" _
                                           & vbNewLine & "надійшло зі складу в/ч - 2" _
                                           & vbNewLine & "інші надходження - 3" _
                                           & vbNewLine & "витрата - 4" _
                                           & vbNewLine & "передано до інших в/ч (підрозділів)- 5" _
                                           & vbNewLine & "інші вибуття - 6", "Тип операції", 4)
Select Case typeOperationAdd
   Case 1
   ThisWorkbook.ActiveSheet.Range("ContentAddActChangeState81_14") = "надійшло з інших в/ч (підрозділів)"
   Case 2
    ThisWorkbook.ActiveSheet.Range("ContentAddActChangeState81_14") = "надійшло зі складу в/ч"
   Case 3
    ThisWorkbook.ActiveSheet.Range("ContentAddActChangeState81_14") = "інші надходження"
    Case 4
    ThisWorkbook.ActiveSheet.Range("ContentAddActChangeState81_14") = "витрата"
    Case 5
    ThisWorkbook.ActiveSheet.Range("ContentAddActChangeState81_14") = "передано до інших в/ч (підрозділів)"
    Case 6
    ThisWorkbook.ActiveSheet.Range("ContentAddActChangeState81_14") = "інші вибуття"
   End Select
   
  
    
    'збільшуємо номер документа на 1
Dim clearNumber As Long
'=======3.9.19=======
Dim increase As String
increase = ThisWorkbook.Sheets("Акт зм. якіс. ст. д14 81").Range("increaseActChangeState81_14").value
If increase = "так" Then
    clearNumber = ThisWorkbook.Sheets("Акт зм. якіс. ст. д14 81").Range("clearNumberActChangeState81_14").value

    ThisWorkbook.Sheets("Акт зм. якіс. ст. д14 81").Range("clearNumberActChangeState81_14") = clearNumber + 1
End If
'=======3.9.19=======
    
    
    
    
    
'    'збільшуємо номер документа на 1
'    Dim clearNumber As Long
'    clearNumber = ThisWorkbook.Sheets("Акт зм. якіс. ст. д14 81").Range("clearNumberActChangeState81_14").value
'    ThisWorkbook.Sheets("Акт зм. якіс. ст. д14 81").Range("clearNumberActChangeState81_14") = clearNumber + 1
    
'    'забираємо суму і кількість
'    Dim allCount As String, allSum As String
'    allCount = ThisWorkbook.Sheets("Акт зм. якіс. ст. д14 81").Range("allCountActChangeState81_14").value
'    allSum = ThisWorkbook.Sheets("Акт зм. якіс. ст. д14 81").Range("allSumActChangeState81_14").value
    
    ThisWorkbook.Sheets("Акт зм. якіс. ст. д14 81").Copy 'зберігаємо як нову книгу
    Dim WbNew As Workbook
    Set WbNew = ActiveWorkbook
    
    'замінюємо формули на значення
    Dim cell As Range
    For Each cell In WbNew.Worksheets(1).Range(Cells(1, 1), Cells(WbNew.Worksheets(1).Range("bookkeepingActChangeState81_14").Row, 16))
        If cell.HasFormula Then
            cell.value = cell.value
        End If
    Next cell
    
'    WbNew.Worksheets(1).Range("allCountActChangeState81_14") = allCount
'    WbNew.Worksheets(1).Range("allSumActChangeState81_14") = allSum
    
    'реквізити для назви кінцевого файлу
    Dim NumberActChangeState81_14 As String, FromUnit As String, toUnit As String, fromMRP As String, toMRP As String
    NumberActChangeState81_14 = WbNew.Worksheets(1).Range("NumberActChangeState81_14")
    'оскільки зміна якості відбувається в одному підрозділі, то тільки 1 підрозділ
    
    FromUnit = WbNew.Worksheets(1).Range("FromUnitActChangeState81_14")
'    toUnit = WbNew.Worksheets(1).Range("toUnitActChangeState81_14")
    '=================3.9.18
    'оскільки зміна якості відбувається в одному підрозділі, то тільки 1 МВО
    fromMRP = WbNew.Worksheets(1).Range("MRPActChangeState81_14")
'    toMRP = WbNew.Worksheets(1).Range("MRPToActChangeState81_14")
    '=================
    
    Dim shp As Shape
    For Each shp In WbNew.Worksheets(1).Shapes
        shp.Delete
    Next shp
    
    Dim fileName As String, SavePath As String
    Dim DateActChangeState81_14 As Date
    DateActChangeState81_14 = WbNew.Worksheets(1).Range("DateActChangeState81_14")
    
    
    Dim typeDocActChangeState81_14 As String
        typeDocActChangeState81_14 = WbNew.Worksheets(1).Range("typeDocActChangeState81_14")
    'Спочатку беремо шлях з комірки
    SavePath = WbNew.Worksheets(1).Range("PathToSaveActChangeState81_14").value
    
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
                ThisWorkbook.Sheets("Акт зм. якіс. ст. д14 81").Range("PathToSaveActChangeState81_14").value = SavePath
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
    fileName = SavePath & Application.PathSeparator & CleanFileName(DateActChangeState81_14 & _
        " " & typeDocActChangeState81_14 & " " & NumberActChangeState81_14 & " " & FromUnit & " (" & fromMRP & ").xlsx")
    '========
    'Зберігаємо
    WbNew.SaveAs fileName:=fileName, FileFormat:=xlOpenXMLWorkbook
    
    Application.DisplayAlerts = True
End Sub


'============3.9.16=========
Sub addFromTableActChangeState81_14()

'Dim lo As ListObject
'Set lo = Worksheets("База").ListObjects("AllData")

Dim ws As Worksheet
Set ws = Selection.Worksheet 'лист, де був клік

Dim wsDoc As Worksheet
Set wsDoc = ThisWorkbook.Worksheets("Акт зм. якіс. ст. д14 81")
Dim tbActual As ListObject

Set tbActual = wsDoc.ListObjects("tbActChangeState81_14")

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
                wsDoc.Range("MRPActChangeState81_14") = MRP
            End If
            
            
            
            
        wsDoc.Range("MRPActChangeState81_14") = MRP
    
        'додаємо інформацію в документ
            If tbActual.DataBodyRange.Rows(1).Cells(1, 18) = "" Then
                With tbActual.DataBodyRange.Rows
                    .Cells(1, 18) = NameOfThings
                    .Cells(1, 3) = Nomenclature
                    .Cells(1, 4) = unit
                    .Cells(1, 5) = Category
                    .Cells(1, 7) = price
                    .Cells(1, 6) = CDbl(quantity)
                    .Cells(1, 19) = Note
                    .Cells(1, 21) = CategoryTo
                    
                    .Cells(1, 23) = Note2
                    .Cells(1, 24) = Note3
                End With
            Else
                    Dim newRow As ListRow
                    Set newRow = tbActual.ListRows.Add
                With newRow.Range
                    .Cells(1, 18) = NameOfThings
                    .Cells(1, 3) = Nomenclature
                    .Cells(1, 4) = unit
                    .Cells(1, 5) = Category
                    .Cells(1, 7) = price
                    .Cells(1, 6) = CDbl(quantity)
                    .Cells(1, 19) = Note
                    .Cells(1, 21) = CategoryTo
                    
                    .Cells(1, 23) = Note2
                    .Cells(1, 24) = Note3
                    
                End With
            End If

    listOfThings = listOfThings & NameOfThings & " " & CDbl(quantity) & " " & unit & vbNewLine
    
    
    
    Next rowIndex

MsgBox listOfThings, vbInformation, "Майно додане в Акт зміни якісного стану"

On Error Resume Next
    Application.CommandBars("Cell").Controls("Додати в документ").Delete
On Error GoTo 0


End Sub
'============3.9.16=========


