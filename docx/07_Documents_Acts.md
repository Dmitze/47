# 07 — Модулі первинних документів: Акти та Заявки

У цьому документі наведено повну документацію та вихідний код спеціалізованих актів та модуля заявок:
1. **ActRTFA440_23.bas** — Акт приймання-передачі основних засобів (наказ МОУ № 440, форма 23)
2. **ActRTFA440_24.bas** — Акт передачі основних засобів (наказ МОУ № 440, форма 24)
3. **ActRTFA818.bas** — Акт приймання-передачі основних засобів (наказ Мінфіну № 818)
4. **ActChangeState81_14.bas** — Акт зміни якісного стану майна (наказ МОУ № 81, форма 14)
5. **ActQState81_1.bas** — Акт технічного (якісного) стану майна (наказ МОУ № 81, форма 1)
6. **ActPlacement440_45.bas** — Акт розміщення (прийняття) майна (наказ МОУ № 440, форма 45)
7. **RequestM.bas** — Модуль формування заявок/запитів на отримання майна

---

## 1. ActRTFA440_23.bas — Акт ОЗ ф. 23 (наказ МОУ 440)
### 1.1. Призначення
Застосовується при надходженні, перерозподілі або переміщенні об'єктів основних засобів (техніка, агрегати, озброєння, прилади) між підрозділами. Таблиця `tbFA440_23` на аркуші «Акт_пр_пд440_23».

### 1.2. Вихідний код ActRTFA440_23.bas
```vba
Attribute VB_Name = "ActRTFA440_23"
Sub addPosFA440_23()
  
    Dim NameOfThings As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Акт_пп_ОЗ440_23")
    
    Dim tbFA440_23 As ListObject
    Set tbFA440_23 = ws.ListObjects("tbFA440_23")
    
Dim ptReduses As PivotTable     'таблиця з переліком МЦ
'оновити таблицю, перед вибором МЦ
'Set ptReduses = wsResiduse2.PivotTables("ptReduses")
'ptReduses.RefreshTable
'=============3.9.15
'If IsError(ws.Range("FromUnitInv440_25").value) Then
'    MsgBox "Схоже ви не заповнили довідник і/або не вибрали хто буде передавати(здавати, видавати) майно." _
'    & vbNewLine & vbNewLine & "Заповінсть довідник і виберіть, хто буде ПЕРДАВАТИ майно.", _
'    vbCritical, "Заповніть довідник, виберіть хто буде передавати"
'
'    ws.Range("MRPInv440_25").Activate
'    Exit Sub
'End If
'
'If IsError(ws.Range("toUnitInv440_25").value) Then
'    MsgBox "Схоже ви не заповнили довідник і/або не вибрали хто буде отримувати (приймати) майно." _
'    & vbNewLine & vbNewLine & "Заповінсть довідник і виберіть, хто буде ОТРИМУВАТИ майно.", _
'    vbCritical, "Заповніть довідник, виберіть отримувачів"
'
'    ws.Range("MRPToInv440_25").Activate
'    Exit Sub
'End If
'=============3.9.15


'=============3.9.11
'визначаємо підрозділ-відправник, беремо з документа
Dim FromUnit As String
FromUnit = ws.Range("FromUnitFA440_23")

Dim tempPoolOThings As Range
Set tempPoolOThings = GetResidusePool(wsResiduse2, FromUnit)

If tempPoolOThings Is Nothing Then Exit Sub
'=============3.9.11



 Dim Basis As String
 Basis = InputBox("Підстава (мета) операції: ", NameOfThings, ThisWorkbook.Worksheets("Акт_пп_ОЗ440_23").Range("basisFA440_23").value)
 ThisWorkbook.Worksheets("Акт_пп_ОЗ440_23").Range("basisFA440_23") = Basis
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
            
            quantity = InputBox("Кількість: ", NameOfThings, QuantityRest)

'            NoteDt = InputBox("Субрахунок для списання, залишити 1 (якщо не знаєте який - залишіть порожнім): ", NameOfThings, "6611 8013 8113")
            
           
            
            Category = InputBox("Категорія відправника: ", NameOfThings, rngWithName.Offset(, -8).value)
            CategoryTo = InputBox("Категорія отримувача: ", NameOfThings, rngWithName.Offset(, -8).value)
'            VAT = InputBox("ПДВ (за наявності): ", NameOfThings, 0)
'            Note1 = InputBox("Додаткове поле 1", NameOfThings, rngWithName.Offset(, -5).value)
            Note2 = InputBox("Додаткове поле 2", NameOfThings, rngWithName.Offset(, -4).value)
            Note3 = InputBox("Додаткове поле 3", NameOfThings, rngWithName.Offset(, -3).value)
            If tbFA440_23.DataBodyRange.Rows(1).Cells(1, 2) = "" Then
                With tbFA440_23.DataBodyRange.Rows
                    .Cells(1, 2) = name
                    .Cells(1, 3) = Nomenclature
                    .Cells(1, 16) = unit
                    .Cells(1, 7) = Category
                    .Cells(1, 5) = price
                    .Cells(1, 6) = CDbl(quantity)
                    .Cells(1, 8) = CDbl(quantity)
                    .Cells(1, 14) = Note
                    .Cells(1, 9) = CategoryTo
                    
'                    .Cells(1, 17) = Note1
                    .Cells(1, 18) = Note2
                    .Cells(1, 19) = Note3
                End With
            Else
                    Dim newRow As ListRow
                    Set newRow = tbFA440_23.ListRows.Add
                With newRow.Range
                    .Cells(1, 2) = name
                    .Cells(1, 3) = Nomenclature
                    .Cells(1, 16) = unit
                    .Cells(1, 7) = Category
                    .Cells(1, 5) = price
                    .Cells(1, 6) = CDbl(quantity)
                    .Cells(1, 8) = CDbl(quantity)
                    .Cells(1, 14) = Note
                    .Cells(1, 9) = CategoryTo
                    
'                    .Cells(1, 17) = Note1
                    .Cells(1, 18) = Note2
                    .Cells(1, 19) = Note3
                    
                End With
            End If
           
            
            Dim AddNext
            
            AddNext = MsgBox("Майно додано! Додати наступну позицію?", vbYesNo, NameOfThings)
            
        End If
        
    Loop While AddNext = 6
    

End Sub

Sub clearTableFA440_23()
Dim tbActual As ListObject
Set tbActual = ThisWorkbook.Worksheets("Акт_пп_ОЗ440_23").ListObjects("tbFA440_23")
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
        If j <> 1 And j <> 10 And j <> 12 Then
            firstRow.Range.Cells(1, j).ClearContents
        End If
 Next j
 
End Sub


Sub addRowFA440_23()
Dim tbActual As ListObject
Set tbActual = ThisWorkbook.Worksheets("Акт_пп_ОЗ440_23").ListObjects("tbFA440_23")

Dim newRow As ListRow
Set newRow = tbActual.ListRows.Add

End Sub

Sub SaveFA440_23()

 Dim typeOperation As String
   
   typeOperation = InputBox("Тип операції" & vbNewLine & "надходження - 1" _
                                           & vbNewLine & "внутрішнє переміщення - 2" _
                                           & vbNewLine & "вибуття - 3", "Тип операції", 3)
   
   Select Case typeOperation
   Case 1
   ThisWorkbook.Sheets("Акт_пп_ОЗ440_23").Range("ContentFA440_23") = "надходження"
   Case 2
    ThisWorkbook.Sheets("Акт_пп_ОЗ440_23").Range("ContentFA440_23") = "внутрішнє переміщення"
   Case 3
    ThisWorkbook.Sheets("Акт_пп_ОЗ440_23").Range("ContentFA440_23") = "вибуття"
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
   ThisWorkbook.ActiveSheet.Range("ContentInvAdd440_23") = "надійшло з інших в/ч (підрозділів)"
   Case 2
    ThisWorkbook.ActiveSheet.Range("ContentInvAdd440_23") = "надійшло зі складу в/ч"
   Case 3
    ThisWorkbook.ActiveSheet.Range("ContentInvAdd440_23") = "інші надходження"
    Case 4
    ThisWorkbook.ActiveSheet.Range("ContentInvAdd440_23") = "витрата"
    Case 5
    ThisWorkbook.ActiveSheet.Range("ContentInvAdd440_23") = "передано до інших в/ч (підрозділів)"
    Case 6
    ThisWorkbook.ActiveSheet.Range("ContentInvAdd440_23") = "інші вибуття"
   End Select
   
   
   
   
   
'збільшуємо номер документа на 1
Dim clearNumber As Long

clearNumber = ThisWorkbook.Sheets("Акт_пп_ОЗ440_23").Range("clearNumberFA440_23").value

ThisWorkbook.Sheets("Акт_пп_ОЗ440_23").Range("clearNumberFA440_23") = clearNumber + 1

'забираємо суму і кількість
'Dim allCount As String
'Dim allSum As String
'
'allCount = ThisWorkbook.Sheets("Акт_пп_ОЗ440_23").Range("allCountFA440_23").value
'allSum = ThisWorkbook.Sheets("Акт_пп_ОЗ440_23").Range("allSumFA440_23").value
'
        ThisWorkbook.Sheets("Акт_пп_ОЗ440_23").Select 'копіюємо перший лист з нарядом
        ThisWorkbook.Sheets("Акт_пп_ОЗ440_23").Copy 'зберігаємо як нову книгу
   
     Dim WbNew As Workbook
     
        Set WbNew = ActiveWorkbook
   
     'замінюємо на значення клітинки, які підтягуються по формулі з іншого листа, щоб у збережених документах не злітали дані
     
      For Each cell In WbNew.Worksheets(1).Range(Cells(1, 1), Cells(WbNew.Worksheets(1).Range("bookkeepingFA440_23").Row, 16))
        If cell.HasFormula Then
            cell.value = cell.value
        End If
     Next cell
     
'    WbNew.Worksheets(1).Range("allCountFA440_23") = allCount
'    WbNew.Worksheets(1).Range("allSumFA440_23") = allSum
    
   Dim NumberFA440_23 As String
     NumberFA440_23 = WbNew.Worksheets(1).Range("NumberFA440_23")
     
   Dim FromUnit As String
   Dim toUnit As String
        FromUnit = WbNew.Worksheets(1).Range("FromUnitFA440_23")
        toUnit = WbNew.Worksheets(1).Range("toUnitFA440_23")
   
   '=================3.9.11
   Dim fromMRP  As String
   Dim toMRP As String
    
    fromMRP = WbNew.Worksheets(1).Range("MRPFromFA440_23")
    toMRP = WbNew.Worksheets(1).Range("MRPToFA440_23")
    '=================
        
        
    Dim shp As Shape
    For Each shp In WbNew.Worksheets(1).Shapes
        shp.Delete
    Next shp
        
      Dim fileName As String, fileName2 As String, lineBrake As String
        'Шлях до файла зберігання
        SavePath = WbNew.Worksheets(1).Range("PathToSaveFA440_23")
        Dim DateFA440_23 As Date
        DateFA440_23 = WbNew.Worksheets(1).Range("DateFA440_23")
        
        Dim typeDocFA440_23 As String
        typeDocFA440_23 = WbNew.Worksheets(1).Range("typeDocFA440_23")
        
        On Error Resume Next
         If Dir(SavePath, vbDirectory) = vbNullString Then
         On Error GoTo 0
        With Application.FileDialog(msoFileDialogFolderPicker)
            .Title = "Виберіть папку для збереження"
            .AllowMultiSelect = False
            If .show = -1 Then
                SavePath = .SelectedItems(1)
                'Записуємо новий шлях у клітинку
                ThisWorkbook.Sheets("Акт_пп_ОЗ440_23").Range("PathToSaveFA440_23").value = SavePath
            Else
                MsgBox "Збереження скасовано", vbExclamation
                WbNew.Close SaveChanges:=False
                Application.DisplayAlerts = True
                Exit Sub
            End If
        End With
    End If
   
        
        'складаємо назву файла
        '3.9.11========
        fileName = SavePath & Application.PathSeparator & CleanFileName(DateFA440_23 & _
        " " & typeDocFA440_23 & " " & NumberFA440_23 & " " & FromUnit & " (" & fromMRP & ") - " & toUnit & " (" & toMRP & ").xlsx")
        '========
        
        
        'заміняємо перенесення рядка, косі лінії

        
        'зберігаємо файл
        On Error GoTo WrongPath
         WbNew.SaveAs fileName:=fileName, FileFormat:=xlOpenXMLWorkbook
     On Error GoTo 0
      Exit Sub
WrongPath:
     MsgBox "Ймовірно, не правильно вказана адреса збереження файла, виправте адресу збереження файлу( на зразок ""D:\База обліку майна ЗСУ 2024"" *без косої лінії вкінці) і спробуйтие зберегти файл ще раз!", vbCritical, "Не правильний шлях збереження"
     Exit Sub
     
End Sub




```

---

## 2. ActRTFA440_24.bas — Акт передачі ОЗ ф. 24 (наказ МОУ 440)
### 2.1. Призначення
Спеціалізована форма передачі техніки та основних засобів усередині частини чи стороннім організаціям. Таблиця `tbFA440_24` на аркуші «Акт_пр_прд440_24».

### 2.2. Вихідний код ActRTFA440_24.bas
```vba
Attribute VB_Name = "ActRTFA440_24"
Sub addPosFA440_24()
  
    Dim NameOfThings As String
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Акт_пп_Зап440_24")

    Dim tbFA440_24 As ListObject
    Set tbFA440_24 = ws.ListObjects("tbFA440_24")
    

'=============3.9.11
'визначаємо підрозділ-відправник, беремо з документа
Dim FromUnit As String
FromUnit = ws.Range("FromUnitFA440_24")

Dim tempPoolOThings As Range
Set tempPoolOThings = GetResidusePool(wsResiduse2, FromUnit)

If tempPoolOThings Is Nothing Then Exit Sub
'=============3.9.11



 Dim Basis As String
 Basis = InputBox("Підстава (мета) операції: ", NameOfThings, ThisWorkbook.Worksheets("Акт_пп_Зап440_24").Range("basisFA440_24").value)
 ThisWorkbook.Worksheets("Акт_пп_Зап440_24").Range("basisFA440_24") = Basis
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
            
            quantity = InputBox("Кількість: ", NameOfThings, QuantityRest)

'            NoteDt = InputBox("Субрахунок для списання, залишити 1 (якщо не знаєте який - залишіть порожнім): ", NameOfThings, "6611 8013 8113")
            
           
            
            Category = InputBox("Категорія відправника: ", NameOfThings, rngWithName.Offset(, -8).value)
            CategoryTo = InputBox("Категорія отримувача: ", NameOfThings, rngWithName.Offset(, -8).value)
'            VAT = InputBox("ПДВ (за наявності): ", NameOfThings, 0)
'            Note1 = InputBox("Додаткове поле 1", NameOfThings, rngWithName.Offset(, -5).value)
            Note2 = InputBox("Додаткове поле 2", NameOfThings, rngWithName.Offset(, -4).value)
            Note3 = InputBox("Додаткове поле 3", NameOfThings, rngWithName.Offset(, -3).value)
            If tbFA440_24.DataBodyRange.Rows(1).Cells(1, 2) = "" Then
                With tbFA440_24.DataBodyRange.Rows
                    .Cells(1, 2) = name
                    .Cells(1, 3) = Nomenclature
                    .Cells(1, 5) = unit
                    .Cells(1, 8) = Category
                    .Cells(1, 6) = price
                    .Cells(1, 7) = CDbl(quantity)
                    .Cells(1, 9) = CDbl(quantity)
                    .Cells(1, 12) = Note
                    .Cells(1, 10) = CategoryTo
                    
'                    .Cells(1, 13) = Note1
                    .Cells(1, 14) = Note2
                    .Cells(1, 15) = Note3
                End With
            Else
                    Dim newRow As ListRow
                    Set newRow = tbFA440_24.ListRows.Add
                With newRow.Range
                    .Cells(1, 2) = name
                    .Cells(1, 3) = Nomenclature
                    .Cells(1, 5) = unit
                    .Cells(1, 8) = Category
                    .Cells(1, 6) = price
                    .Cells(1, 7) = CDbl(quantity)
                    .Cells(1, 9) = CDbl(quantity)
                    .Cells(1, 12) = Note
                    .Cells(1, 10) = CategoryTo
                    
'                    .Cells(1, 13) = Note1
                    .Cells(1, 14) = Note2
                    .Cells(1, 15) = Note3
                    
                End With
            End If
           
            
            Dim AddNext
            
            AddNext = MsgBox("Майно додано! Додати наступну позицію?", vbYesNo, NameOfThings)
            
        End If
        
    Loop While AddNext = 6
    

End Sub

Sub clearTableFA440_24()
Dim tbActual As ListObject
Set tbActual = ThisWorkbook.Worksheets("Акт_пп_Зап440_24").ListObjects("tbFA440_24")
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
 'додав обмеження і на 4 колонку
 '==========3.9.25  ========
        If j <> 1 And j <> 4 And j <> 11 Then
 '==========3.9.25  ========
            firstRow.Range.Cells(1, j).ClearContents
        End If
 Next j
 
End Sub


Sub addRowFA440_24()
Dim tbActual As ListObject
Set tbActual = ThisWorkbook.Worksheets("Акт_пп_Зап440_24").ListObjects("tbFA440_24")

Dim newRow As ListRow
Set newRow = tbActual.ListRows.Add

End Sub

Sub SaveFA440_24()

 Dim typeOperation As String
   
   typeOperation = InputBox("Тип операції" & vbNewLine & "надходження - 1" _
                                           & vbNewLine & "внутрішнє переміщення - 2" _
                                           & vbNewLine & "вибуття - 3", "Тип операції", 3)
   
   Select Case typeOperation
   Case 1
   ThisWorkbook.Sheets("Акт_пп_Зап440_24").Range("ContentFA440_24") = "надходження"
   Case 2
    ThisWorkbook.Sheets("Акт_пп_Зап440_24").Range("ContentFA440_24") = "внутрішнє переміщення"
   Case 3
    ThisWorkbook.Sheets("Акт_пп_Зап440_24").Range("ContentFA440_24") = "вибуття"
   End Select
   
    typeOperationAdd = InputBox("Детальний тип операції" & vbNewLine & "надійшло з інших в/ч (підрозділів) - 1" _
                                           & vbNewLine & "надійшло зі складу в/ч - 2" _
                                           & vbNewLine & "інші надходження - 3" _
                                           & vbNewLine & "витрата - 4" _
                                           & vbNewLine & "передано до інших в/ч (підрозділів)- 5" _
                                           & vbNewLine & "інші вибуття - 6", "Тип операції", 5)
Select Case typeOperationAdd
   Case 1
   ThisWorkbook.ActiveSheet.Range("ContentInvAdd440_24") = "надійшло з інших в/ч (підрозділів)"
   Case 2
    ThisWorkbook.ActiveSheet.Range("ContentInvAdd440_24") = "надійшло зі складу в/ч"
   Case 3
    ThisWorkbook.ActiveSheet.Range("ContentInvAdd440_24") = "інші надходження"
    Case 4
    ThisWorkbook.ActiveSheet.Range("ContentInvAdd440_24") = "витрата"
    Case 5
    ThisWorkbook.ActiveSheet.Range("ContentInvAdd440_24") = "передано до інших в/ч (підрозділів)"
    Case 6
    ThisWorkbook.ActiveSheet.Range("ContentInvAdd440_24") = "інші вибуття"
   End Select
   
   
'збільшуємо номер документа на 1
Dim clearNumber As Long

clearNumber = ThisWorkbook.Sheets("Акт_пп_Зап440_24").Range("clearNumberFA440_24").value

ThisWorkbook.Sheets("Акт_пп_Зап440_24").Range("clearNumberFA440_24") = clearNumber + 1

'забираємо суму і кількість
Dim allCount As String
'Dim allSum As String
'
allCount = ThisWorkbook.Sheets("Акт_пп_Зап440_24").Range("allSumAndQuantity440_24").value
'allSum = ThisWorkbook.Sheets("Акт_пп_Зап440_24").Range("allSumFA440_24").value
'
        ThisWorkbook.Sheets("Акт_пп_Зап440_24").Select 'копіюємо перший лист з нарядом
        ThisWorkbook.Sheets("Акт_пп_Зап440_24").Copy 'зберігаємо як нову книгу
   
     Dim WbNew As Workbook
     
        Set WbNew = ActiveWorkbook
   
     'замінюємо на значення клітинки, які підтягуються по формулі з іншого листа, щоб у збережених документах не злітали дані
     
      For Each cell In WbNew.Worksheets(1).Range(Cells(1, 1), Cells(WbNew.Worksheets(1).Range("bookkeepingFA440_24").Row, 16))
        If cell.HasFormula Then
            cell.value = cell.value
        End If
     Next cell
     
    WbNew.Worksheets(1).Range("allSumAndQuantity440_24") = allCount
'    WbNew.Worksheets(1).Range("allSumFA440_24") = allSum
    
   Dim NumberFA440_24 As String
     NumberFA440_24 = WbNew.Worksheets(1).Range("NumberFA440_24")
     
   Dim FromUnit As String
   Dim toUnit As String
        FromUnit = WbNew.Worksheets(1).Range("FromUnitFA440_24")
        toUnit = WbNew.Worksheets(1).Range("toUnitFA440_24")
        
'=================3.9.11
   Dim fromMRP  As String
   Dim toMRP As String
    
    fromMRP = WbNew.Worksheets(1).Range("MRPFromFA440_24")
    toMRP = WbNew.Worksheets(1).Range("MRPToFA440_24")
'=================
        
        
        
    Dim shp As Shape
    For Each shp In WbNew.Worksheets(1).Shapes
        shp.Delete
    Next shp
        
      Dim fileName As String, fileName2 As String, lineBrake As String
        'Шлях до файла зберігання
        SavePath = WbNew.Worksheets(1).Range("PathToSaveFA440_24")
        Dim DateFA440_24 As Date
        DateFA440_24 = WbNew.Worksheets(1).Range("DateFA440_24")
        
        Dim typeDocFA440_24 As String
        typeDocFA440_24 = WbNew.Worksheets(1).Range("typeDocFA440_24")
        
        
        On Error Resume Next
         If Dir(SavePath, vbDirectory) = vbNullString Then
         On Error GoTo 0
        With Application.FileDialog(msoFileDialogFolderPicker)
            .Title = "Виберіть папку для збереження"
            .AllowMultiSelect = False
            If .show = -1 Then
                SavePath = .SelectedItems(1)
                'Записуємо новий шлях у клітинку
                ThisWorkbook.Sheets("Акт_пп_Зап440_24").Range("PathToSaveFA440_24").value = SavePath
            Else
                MsgBox "Збереження скасовано", vbExclamation
                WbNew.Close SaveChanges:=False
                Application.DisplayAlerts = True
                Exit Sub
            End If
        End With
    End If
        
        
       
        
        'складаємо назву файла
        '3.9.11========
        fileName = SavePath & Application.PathSeparator & CleanFileName(DateFA440_24 & _
        " " & typeDocFA440_24 & " " & NumberFA440_24 & " " & FromUnit & " (" & fromMRP & ") - " & toUnit & " (" & toMRP & ").xlsx")
        '========
     

        
        
           'зберігаємо файл
        On Error GoTo WrongPath
         WbNew.SaveAs fileName:=fileName, FileFormat:=xlOpenXMLWorkbook
     On Error GoTo 0
      Exit Sub
WrongPath:
     MsgBox "Ймовірно, не правильно вказана адреса збереження файла, виправте адресу збереження файлу( на зразок ""D:\База обліку майна ЗСУ 2024"" *без косої лінії вкінці) і спробуйтие зберегти файл ще раз!", vbCritical, "Не правильний шлях збереження"
     Exit Sub
     
        'якщо у вас тут помилка, ймовіно ви не змінили шлях збереження файлу, або допустити помилку в шляху збереження (наприклад лишній знак "/" вкінці)
    
     
End Sub






```

---

## 3. ActRTFA818.bas — Акт ОЗ за типовою формою 818
### 3.1. Призначення
Форма цивільного/бюджетного стандарту бухгалтерського обліку згідно з наказом Мінфіну № 818. Таблиця `tbInvFA818` на аркуші «Акт пр. прд. ОЗ 818».

### 3.2. Вихідний код ActRTFA818.bas
```vba
Attribute VB_Name = "ActRTFA818"
Sub addPosInvFA818()
  
    Dim NameOfThings As String


    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Акт вн. пер. ОЗ 818")
    
    Dim tbInvFA818 As ListObject
    Set tbInvFA818 = ws.ListObjects("tbInvFA818")
    
    '=============3.9.15
If IsError(ws.Range("FromUnitInvFA818").value) Then
    MsgBox "Схоже ви не заповнили довідник і/або не вибрали хто буде передавати(здавати, видавати) майно." _
    & vbNewLine & vbNewLine & "Заповінсть довідник і виберіть, хто буде ПЕРДАВАТИ майно.", _
    vbCritical, "Заповніть довідник, виберіть хто буде передавати"
    
    ws.Range("MRPInvFromFA818").Activate
    Exit Sub
End If

If IsError(ws.Range("toUnitInvFA818").value) Then
    MsgBox "Схоже ви не заповнили довідник і/або не вибрали хто буде отримувати (приймати) майно." _
    & vbNewLine & vbNewLine & "Заповінсть довідник і виберіть, хто буде ОТРИМУВАТИ майно.", _
    vbCritical, "Заповніть довідник, виберіть отримувачів"
    
    ws.Range("MRPToInvFA818").Activate
    Exit Sub
End If
'=============3.9.15
    
    
    
'=============3.9.11
'визначаємо підрозділ-відправник, беремо з документа
Dim FromUnit As String
FromUnit = ws.Range("FromUnitInvFA818")

Dim tempPoolOThings As Range
Set tempPoolOThings = GetResidusePool(wsResiduse2, FromUnit)

If tempPoolOThings Is Nothing Then Exit Sub
'=============3.9.11

 Dim Basis As String
 Basis = InputBox("Підстава (мета) операції: ", NameOfThings, ThisWorkbook.Worksheets("Акт вн. пер. ОЗ 818").Range("basisFA818").value)
 ThisWorkbook.Worksheets("Акт вн. пер. ОЗ 818").Range("basisFA818") = Basis
    Do
        
        Dim frm     As UserForm1
        Set frm = UserForms.Add(UserForm1.name)
        frm.ListData = tempPoolOThings
'        frm.ListData = wsResiduse2.Range("Goods")
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
            
            quantity = InputBox("Кількість: ", NameOfThings, QuantityRest)

'            NoteDt = InputBox("Субрахунок для списання, залишити 1 (якщо не знаєте який - залишіть порожнім): ", NameOfThings, "6611 8013 8113")
            
           
            
            Category = InputBox("Категорія відправника: ", NameOfThings, rngWithName.Offset(, -8).value)
            CategoryTo = InputBox("Категорія отримувача: ", NameOfThings, rngWithName.Offset(, -8).value)
'            VAT = InputBox("ПДВ (за наявності): ", NameOfThings, 0)
'            Note1 = InputBox("Додаткове поле 1", NameOfThings, rngWithName.Offset(, -5).value)
            Note2 = InputBox("Додаткове поле 2", NameOfThings, rngWithName.Offset(, -4).value)
            Note3 = InputBox("Додаткове поле 3", NameOfThings, rngWithName.Offset(, -3).value)
            If tbInvFA818.DataBodyRange.Rows(1).Cells(1, 2) = "" Then
                With tbInvFA818.DataBodyRange.Rows
                    .Cells(1, 1) = name
                    .Cells(1, 2) = Nomenclature
                    .Cells(1, 3) = unit
                    .Cells(1, 8) = Category
                    .Cells(1, 5) = price
                    .Cells(1, 4) = CDbl(quantity)
                    .Cells(1, 7) = Note
                    .Cells(1, 9) = CategoryTo
                    
'                    .Cells(1, 10) = Note1
                    .Cells(1, 11) = Note2
                    .Cells(1, 12) = Note3
                End With
            Else
                    Dim newRow As ListRow
                    Set newRow = tbInvFA818.ListRows.Add
                With newRow.Range
                    .Cells(1, 1) = name
                    .Cells(1, 2) = Nomenclature
                    .Cells(1, 3) = unit
                    .Cells(1, 8) = Category
                    .Cells(1, 5) = price
                    .Cells(1, 4) = CDbl(quantity)
                    .Cells(1, 7) = Note
                    .Cells(1, 9) = CategoryTo
                    
'                    .Cells(1, 10) = Note1
                    .Cells(1, 11) = Note2
                    .Cells(1, 12) = Note3
                    
                End With
            End If
           
            
            Dim AddNext
            
            AddNext = MsgBox("Майно додано! Додати наступну позицію?", vbYesNo, NameOfThings)
            
        End If
        
    Loop While AddNext = 6
    

End Sub

Sub clearTableFA818()
Dim tbActual As ListObject
Set tbActual = ThisWorkbook.Worksheets("Акт вн. пер. ОЗ 818").ListObjects("tbInvFA818")
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
        If j <> 6 Then
            firstRow.Range.Cells(1, j).ClearContents
        End If
 Next j
 
End Sub


Sub addRowFA818()
Dim tbActual As ListObject
Set tbActual = ThisWorkbook.Worksheets("Акт вн. пер. ОЗ 818").ListObjects("tbInvFA818")

Dim newRow As ListRow
Set newRow = tbActual.ListRows.Add

End Sub

Sub SaveInvFA818()

 Dim typeOperation As String
   
   typeOperation = InputBox("Тип операції" & vbNewLine & "надходження - 1" _
                                           & vbNewLine & "внутрішнє переміщення - 2" _
                                           & vbNewLine & "вибуття - 3", "Тип операції", 2)
   
   Select Case typeOperation
   Case 1
   ThisWorkbook.Sheets("Акт вн. пер. ОЗ 818").Range("ContentInvFA818") = "надходження"
   Case 2
    ThisWorkbook.Sheets("Акт вн. пер. ОЗ 818").Range("ContentInvFA818") = "внутрішнє переміщення"
   Case 3
    ThisWorkbook.Sheets("Акт вн. пер. ОЗ 818").Range("ContentInvFA818") = "вибуття"
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
   ThisWorkbook.ActiveSheet.Range("ContentInvAdd818") = "надійшло з інших в/ч (підрозділів)"
   Case 2
    ThisWorkbook.ActiveSheet.Range("ContentInvAdd818") = "надійшло зі складу в/ч"
   Case 3
    ThisWorkbook.ActiveSheet.Range("ContentInvAdd818") = "інші надходження"
    Case 4
    ThisWorkbook.ActiveSheet.Range("ContentInvAdd818") = "витрата"
    Case 5
    ThisWorkbook.ActiveSheet.Range("ContentInvAdd818") = "передано до інших в/ч (підрозділів)"
    Case 6
    ThisWorkbook.ActiveSheet.Range("ContentInvAdd818") = "інші вибуття"
   End Select
   
   
   
'збільшуємо номер документа на 1
Dim clearNumber As Long

clearNumber = ThisWorkbook.Sheets("Акт вн. пер. ОЗ 818").Range("clearNumberInvFA818").value

ThisWorkbook.Sheets("Акт вн. пер. ОЗ 818").Range("clearNumberInvFA818") = clearNumber + 1

'забираємо суму і кількість
Dim allCount As String
Dim allSum As String

allCount = ThisWorkbook.Sheets("Акт вн. пер. ОЗ 818").Range("allCountInvFA818").value
allSum = ThisWorkbook.Sheets("Акт вн. пер. ОЗ 818").Range("allSumInvFA818").value

        ThisWorkbook.Sheets("Акт вн. пер. ОЗ 818").Select 'копіюємо перший лист з нарядом
        ThisWorkbook.Sheets("Акт вн. пер. ОЗ 818").Copy 'зберігаємо як нову книгу
   
     Dim WbNew As Workbook
     
        Set WbNew = ActiveWorkbook
   
     'замінюємо на значення клітинки, які підтягуються по формулі з іншого листа, щоб у збережених документах не злітали дані
     
      For Each cell In WbNew.Worksheets(1).Range(Cells(1, 1), Cells(WbNew.Worksheets(1).Range("bookkeepingInvFA818").Row, 16))
        If cell.HasFormula Then
            cell.value = cell.value
        End If
     Next cell
     
    WbNew.Worksheets(1).Range("allCountInvFA818") = allCount
    WbNew.Worksheets(1).Range("allSumInvFA818") = allSum
    
   Dim NumberInvFA818 As String
     NumberInvFA818 = WbNew.Worksheets(1).Range("NumberInvFA818")
     
   Dim FromUnit As String
   Dim toUnit As String
        FromUnit = WbNew.Worksheets(1).Range("FromUnitInvFA818")
        toUnit = WbNew.Worksheets(1).Range("toUnitInvFA818")
        
    '=================3.9.11
   Dim fromMRP  As String
   Dim toMRP As String
    
    fromMRP = WbNew.Worksheets(1).Range("MRPInvFromFA818")
    toMRP = WbNew.Worksheets(1).Range("MRPToInvFA818")
    '=================
        
        
        
    Dim shp As Shape
    For Each shp In WbNew.Worksheets(1).Shapes
        shp.Delete
    Next shp
        
      Dim fileName As String, fileName2 As String, lineBrake As String
        'Шлях до файла зберігання
        SavePath = WbNew.Worksheets(1).Range("PathToSaveInvFA818")
        Dim DateInvFA818 As Date
        DateInvFA818 = WbNew.Worksheets(1).Range("DateInvFA818")
        Dim typeDocInvFA818 As String
        typeDocInvFA818 = WbNew.Worksheets(1).Range("typeDocInvFA818")
        
        
        
        
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
                ThisWorkbook.Sheets("Акт вн. пер. ОЗ 818").Range("PathToSaveInvFA818").value = SavePath
            Else
                MsgBox "Збереження скасовано", vbExclamation
                WbNew.Close SaveChanges:=False
                Application.DisplayAlerts = True
                Exit Sub
            End If
        End With
    End If
        
        
        
        'складаємо назву файла
        '3.9.11========
        fileName = SavePath & Application.PathSeparator & CleanFileName(DateInvFA818 & _
        " " & typeDocInvFA818 & " " & NumberInvFA818 & " " & FromUnit & " (" & fromMRP & ") - " & toUnit & " (" & toMRP & ").xlsx")
        '========
        
        
          'зберігаємо файл
        On Error GoTo WrongPath
         WbNew.SaveAs fileName:=fileName, FileFormat:=xlOpenXMLWorkbook
     On Error GoTo 0
      Exit Sub
WrongPath:
     MsgBox "Ймовірно, не правильно вказана адреса збереження файла, виправте адресу збереження файлу( на зразок ""D:\База обліку майна ЗСУ 2024"" *без косої лінії вкінці) і спробуйтие зберегти файл ще раз!", vbCritical, "Не правильний шлях збереження"
     Exit Sub
     
    
End Sub




```

---

## 4. ActChangeState81_14.bas — Акт зміни стану майна ф. 14 (наказ МОУ 81)
### 4.1. Особливість обліку категорій
Цей акт фіксує переведення майна з однієї категорії якісного стану в іншу (наприклад, з 1-ї в 2-гу категорію після видачі, або з 2-ї в 4-ту/5-ту через знос чи пошкодження). Макрос спеціально запитує: категорію «З якої» та категорію «В яку». Таблиця `tbActChangeState81_14` на аркуші «Акт зм. стан. мн. №14 81».

### 4.2. Вихідний код ActChangeState81_14.bas
```vba
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



```

---

## 5. ActQState81_1.bas — Акт якісного/технічного стану ф. 1 (наказ МОУ 81)
### 5.1. Призначення
Складається комісією для встановлення фактичного технічного стану зброї, засобів зв'язку, приладів тощо. Таблиця `tbActQState81_1` на аркуші «Акт якіс. стан. майна №1 81».

### 5.2. Вихідний код ActQState81_1.bas
```vba
Attribute VB_Name = "ActQState81_1"
'ActQState81_1

Sub addPosActQState81_1()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Акт якісн. тех. стану д1 81")
    
    Dim NameOfThings As String

    Dim tbActQState81_1 As ListObject
    Set tbActQState81_1 = ws.ListObjects("tbActQState81_1")
    
    
    
'=============3.9.15
If IsError(ws.Range("FromUnitActQState81_1").value) Then
    MsgBox "Схоже ви не заповнили довідник і/або не вибрали хто буде передавати(здавати, видавати) майно." _
    & vbNewLine & vbNewLine & "Заповінсть довідник і виберіть, хто буде ПЕРДАВАТИ майно.", _
    vbCritical, "Заповніть довідник, виберіть хто буде передавати"
    
    ws.Range("MRPActQState81_1").Activate
    Exit Sub
End If

'If IsError(ws.Range("toUnitActQState81_1").value) Then
'    MsgBox "Схоже ви не заповнили довідник і/або не вибрали хто буде отримувати (приймати) майно." _
'    & vbNewLine & vbNewLine & "Заповінсть довідник і виберіть, хто буде ОТРИМУВАТИ майно.", _
'    vbCritical, "Заповніть довідник, виберіть отримувачів"
'
'    ws.Range("MRPToActQState81_1").Activate
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
FromUnit = ws.Range("FromUnitActQState81_1")

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
            
            '========3.9.20=====
            Dim newRow As ListRow
            If answer = vbYes Then
            
            For i = 1 To tbActQState81_1.ListRows.count
                    If tbActQState81_1.DataBodyRange.Cells(i, 18).value = "" Then
                        Set newRow = tbActQState81_1.ListRows(i)
                                                    
                        Exit For
                    End If
             Next i
             
            ' Якщо порожній рядок не знайдено – додаємо новий
            If newRow Is Nothing Then
                Set newRow = tbActQState81_1.ListRows.Add
            End If
        
            
            
            
   
'            If tbActQState81_1.DataBodyRange.Rows(1).Cells(1, 18) = "" Then
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
'            Else
'                    Dim NewRow As ListRow
'                    Set NewRow = tbActQState81_1.ListRows.Add
'                With NewRow.Range
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
'
'                End With
'            End If
            
            Else
                For i = 1 To tbActQState81_1.ListRows.count
                    If tbActQState81_1.DataBodyRange.Cells(i, 11).value = "" Then
                        Set newRow = tbActQState81_1.ListRows(i)
                            
                        
                        Exit For
                    End If
             Next i
                ' Якщо порожній рядок не знайдено – додаємо новий
                If newRow Is Nothing Then
                    Set newRow = tbActQState81_1.ListRows.Add
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
    '========3.9.20=====
            
            Dim AddNext
            
            AddNext = MsgBox("Майно додано! Додати наступну позицію?", vbYesNo, NameOfThings)
            
        End If
        
    Loop While AddNext = 6
    
    
    '***********Для майна, яке оприбутковується***********
    
    
      
    

End Sub

Sub clearTableActQState81_1()
Dim tbActual As ListObject
Set tbActual = ThisWorkbook.Worksheets("Акт якісн. тех. стану д1 81").ListObjects("tbActQState81_1")
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


Sub addRowActQState81_1()
Dim tbActual As ListObject
Set tbActual = ThisWorkbook.Worksheets("Акт якісн. тех. стану д1 81").ListObjects("tbActQState81_1")

Dim newRow As ListRow
Set newRow = tbActual.ListRows.Add

End Sub




Sub SaveActQState81_1()
    Application.DisplayAlerts = False
    
 Dim Basis As String
 Basis = InputBox("Вкажіть на підставі якого документу або події оформлюється даний документ ", "Підстава (мета) операції:", ThisWorkbook.ActiveSheet.Range("basisActQState81_1").value)
 ThisWorkbook.ActiveSheet.Range("basisActQState81_1") = Basis
    
    
    
    Dim typeOperation As String
    
    typeOperation = InputBox("Тип операції" & vbNewLine & "надходження - 1" _
                                            & vbNewLine & "внутрішнє переміщення - 2" _
                                            & vbNewLine & "вибуття - 3", "Тип операції", 3)
    
    Select Case typeOperation
        Case 1
            ThisWorkbook.Sheets("Акт якісн. тех. стану д1 81").Range("ContentActQState81_1") = "надходження"
        Case 2
            ThisWorkbook.Sheets("Акт якісн. тех. стану д1 81").Range("ContentActQState81_1") = "внутрішнє переміщення"
        Case 3
            ThisWorkbook.Sheets("Акт якісн. тех. стану д1 81").Range("ContentActQState81_1") = "вибуття"
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
   ThisWorkbook.ActiveSheet.Range("ContentAddActQState81_1") = "надійшло з інших в/ч (підрозділів)"
   Case 2
    ThisWorkbook.ActiveSheet.Range("ContentAddActQState81_1") = "надійшло зі складу в/ч"
   Case 3
    ThisWorkbook.ActiveSheet.Range("ContentAddActQState81_1") = "інші надходження"
    Case 4
    ThisWorkbook.ActiveSheet.Range("ContentAddActQState81_1") = "витрата"
    Case 5
    ThisWorkbook.ActiveSheet.Range("ContentAddActQState81_1") = "передано до інших в/ч (підрозділів)"
    Case 6
    ThisWorkbook.ActiveSheet.Range("ContentAddActQState81_1") = "інші вибуття"
   End Select
   
  
       
    'збільшуємо номер документа на 1
Dim clearNumber As Long
Dim increase As String
increase = ThisWorkbook.Sheets("Акт якісн. тех. стану д1 81").Range("increaseActQState81_1").value
If increase = "так" Then
    clearNumber = ThisWorkbook.Sheets("Акт якісн. тех. стану д1 81").Range("clearNumberActQState81_1").value

    ThisWorkbook.Sheets("Акт якісн. тех. стану д1 81").Range("clearNumberActQState81_1") = clearNumber + 1
End If


'    'збільшуємо номер документа на 1
'    Dim clearNumber As Long
'    clearNumber = ThisWorkbook.Sheets("Акт якісн. тех. стану д1 81").Range("clearNumberActQState81_1").value
'    ThisWorkbook.Sheets("Акт якісн. тех. стану д1 81").Range("clearNumberActQState81_1") = clearNumber + 1
    
'    'забираємо суму і кількість
'    Dim allCount As String, allSum As String
'    allCount = ThisWorkbook.Sheets("Акт якісн. тех. стану д1 81").Range("allCountActQState81_1").value
'    allSum = ThisWorkbook.Sheets("Акт якісн. тех. стану д1 81").Range("allSumActQState81_1").value
    
    ThisWorkbook.Sheets("Акт якісн. тех. стану д1 81").Copy 'зберігаємо як нову книгу
    Dim WbNew As Workbook
    Set WbNew = ActiveWorkbook
    
    'замінюємо формули на значення
    Dim cell As Range
    For Each cell In WbNew.Worksheets(1).Range(Cells(1, 1), Cells(WbNew.Worksheets(1).Range("bookkeepingActQState81_1").Row, 16))
        If cell.HasFormula Then
            cell.value = cell.value
        End If
    Next cell
    
'    WbNew.Worksheets(1).Range("allCountActQState81_1") = allCount
'    WbNew.Worksheets(1).Range("allSumActQState81_1") = allSum
    
    'реквізити для назви кінцевого файлу
    Dim NumberActQState81_1 As String, FromUnit As String, toUnit As String, fromMRP As String, toMRP As String
    NumberActQState81_1 = WbNew.Worksheets(1).Range("NumberActQState81_1")
    'оскільки зміна якості відбувається в одному підрозділі, то тільки 1 підрозділ
    
    FromUnit = WbNew.Worksheets(1).Range("FromUnitActQState81_1")
'    toUnit = WbNew.Worksheets(1).Range("toUnitActQState81_1")
    '=================3.9.18
    'оскільки зміна якості відбувається в одному підрозділі, то тільки 1 МВО
    fromMRP = WbNew.Worksheets(1).Range("MRPActQState81_1")
'    toMRP = WbNew.Worksheets(1).Range("MRPToActQState81_1")
    '=================
    
    Dim shp As Shape
    For Each shp In WbNew.Worksheets(1).Shapes
        shp.Delete
    Next shp
    
    Dim fileName As String, SavePath As String
    Dim DateActQState81_1 As Date
    DateActQState81_1 = WbNew.Worksheets(1).Range("DateActQState81_1")
    
    
    Dim typeDocActQState81_1 As String
        typeDocActQState81_1 = WbNew.Worksheets(1).Range("typeDocActQState81_1")
    'Спочатку беремо шлях з комірки
    SavePath = WbNew.Worksheets(1).Range("PathToSaveActQState81_1").value
    
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
                ThisWorkbook.Sheets("Акт якісн. тех. стану д1 81").Range("PathToSaveActQState81_1").value = SavePath
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
    fileName = SavePath & Application.PathSeparator & CleanFileName(DateActQState81_1 & _
        " " & typeDocActQState81_1 & " " & NumberActQState81_1 & " " & FromUnit & " (" & fromMRP & ").xlsx")
    '========
    'Зберігаємо
    WbNew.SaveAs fileName:=fileName, FileFormat:=xlOpenXMLWorkbook
    
    Application.DisplayAlerts = True
End Sub


'============3.9.16=========
Sub addFromTableActQState81_1()

'Dim lo As ListObject
'Set lo = Worksheets("База").ListObjects("AllData")

Dim ws As Worksheet
Set ws = Selection.Worksheet 'лист, де був клік

Dim wsDoc As Worksheet
Set wsDoc = ThisWorkbook.Worksheets("Акт якісн. тех. стану д1 81")
Dim tbActual As ListObject

Set tbActual = wsDoc.ListObjects("tbActQState81_1")

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
                wsDoc.Range("MRPActQState81_1") = MRP
            End If
            
            
            
            
        wsDoc.Range("MRPActQState81_1") = MRP
    
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

MsgBox listOfThings, vbInformation, "Майно додане в Акт якісного (технічного) стану"

On Error Resume Next
    Application.CommandBars("Cell").Controls("Додати в документ").Delete
On Error GoTo 0


End Sub
'============3.9.16=========





```

---

## 6. ActPlacement440_45.bas — Акт розміщення (прийняття) ф. 45 (наказ МОУ 440)
### 6.1. Призначення
Застосовується для обліку розміщення майна на складах зберігання або прийняття від постачальників. Таблиця `tbActPlacement440_45` на аркуші «Акт розм.(прийн.)440_45».

### 6.2. Вихідний код ActPlacement440_45.bas
```vba
Attribute VB_Name = "ActPlacement440_45"

'ActPlacement440_45

Sub addPosActPlacement440_45()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Акт закл.(освіж.)440_45")
    
    Dim NameOfThings As String

    Dim tbActPlacement440_45 As ListObject
    Set tbActPlacement440_45 = ws.ListObjects("tbActPlacement440_45")
    
    
    
'=============3.9.15
If IsError(ws.Range("FromUnitActPlacement440_45").value) Then
    MsgBox "Схоже ви не заповнили довідник і/або не вибрали хто буде передавати(здавати, видавати) майно." _
    & vbNewLine & vbNewLine & "Заповінсть довідник і виберіть, хто буде ПЕРДАВАТИ майно.", _
    vbCritical, "Заповніть довідник, виберіть хто буде передавати"
    
    ws.Range("MRPActPlacement440_45").Activate
    Exit Sub
End If

'If IsError(ws.Range("toUnitActPlacement440_45").value) Then
'    MsgBox "Схоже ви не заповнили довідник і/або не вибрали хто буде отримувати (приймати) майно." _
'    & vbNewLine & vbNewLine & "Заповінсть довідник і виберіть, хто буде ОТРИМУВАТИ майно.", _
'    vbCritical, "Заповніть довідник, виберіть отримувачів"
'
'    ws.Range("MRPToActPlacement440_45").Activate
'    Exit Sub
'End If
'=============3.9.15

'=============3.9.19
Dim answer As VbMsgBoxResult

answer = MsgBox("Додати в частину 'Закладено' - 'Так'" & vbNewLine & "Додати в частину 'Вилучено' - 'Ні'", vbYesNo + vbQuestion, "Підтвердження")

'=============3.9.19

    
'=============3.9.11
'визначаємо підрозділ-відправник, беремо з документа
Dim FromUnit As String
FromUnit = ws.Range("FromUnitActPlacement440_45")

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
            
            For i = 1 To tbActPlacement440_45.ListRows.count
                    If tbActPlacement440_45.DataBodyRange.Cells(i, 2).value = "" Then
                        Set newRow = tbActPlacement440_45.ListRows(i)
                        Exit For
                    End If
                Next i
                ' Якщо порожній рядок не знайдено – додаємо новий
                If newRow Is Nothing Then
                    Set newRow = tbActPlacement440_45.ListRows.Add
                End If
            
           If answer = vbYes Then
                With newRow.Range
                    .Cells(1, 2) = name
                    .Cells(1, 3) = Nomenclature
                    .Cells(1, 4) = unit
                    .Cells(1, 5) = Category
                    .Cells(1, 7) = price
                    .Cells(1, 6) = CDbl(quantity)
                    .Cells(1, 10) = Note
                    .Cells(1, 18) = CategoryTo
                    
                    .Cells(1, 19) = Note2
                    .Cells(1, 20) = Note3
                    
                End With
'            End If
            
            Else
                
                With newRow.Range
                    .Cells(1, 2) = name
                    .Cells(1, 3) = Nomenclature
                    .Cells(1, 4) = unit
                    .Cells(1, 11) = Category
                    .Cells(1, 13) = price
                    .Cells(1, 12) = CDbl(quantity)
                    .Cells(1, 16) = Note
                    .Cells(1, 21) = CategoryTo
                    
                    .Cells(1, 22) = Note2
                    .Cells(1, 23) = Note3
                    
                End With
            
    
            End If
            Set newRow = Nothing
            
            Dim AddNext
            
            AddNext = MsgBox("Майно додано! Додати наступну позицію?", vbYesNo, NameOfThings)
            
        End If
        
    Loop While AddNext = 6
    
    
    '***********Для майна, яке оприбутковується***********
 
End Sub

Sub clearTableActPlacement440_45()
Dim tbActual As ListObject
Set tbActual = ThisWorkbook.Worksheets("Акт закл.(освіж.)440_45").ListObjects("tbActPlacement440_45")
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
        If j <> 1 And j <> 8 And j <> 14 Then
            firstRow.Range.Cells(1, j).ClearContents
        End If
 Next j
 
End Sub


Sub addRowActPlacement440_45()
Dim tbActual As ListObject
Set tbActual = ThisWorkbook.Worksheets("Акт закл.(освіж.)440_45").ListObjects("tbActPlacement440_45")

Dim newRow As ListRow
Set newRow = tbActual.ListRows.Add

End Sub




Sub SaveActPlacement440_45()
    Application.DisplayAlerts = False
    
 Dim Basis As String
 Basis = InputBox("Вкажіть на підставі якого документу або події оформлюється даний документ ", "Підстава (мета) операції:", ThisWorkbook.ActiveSheet.Range("basisActPlacement440_45").value)
 ThisWorkbook.ActiveSheet.Range("basisActPlacement440_45") = Basis
    
    
    
    Dim typeOperation As String
    
    typeOperation = InputBox("Тип операції" & vbNewLine & "надходження - 1" _
                                            & vbNewLine & "внутрішнє переміщення - 2" _
                                            & vbNewLine & "вибуття - 3", "Тип операції", 3)
    
    Select Case typeOperation
        Case 1
            ThisWorkbook.Sheets("Акт закл.(освіж.)440_45").Range("ContentActPlacement440_45") = "надходження"
        Case 2
            ThisWorkbook.Sheets("Акт закл.(освіж.)440_45").Range("ContentActPlacement440_45") = "внутрішнє переміщення"
        Case 3
            ThisWorkbook.Sheets("Акт закл.(освіж.)440_45").Range("ContentActPlacement440_45") = "вибуття"
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
   ThisWorkbook.ActiveSheet.Range("ContentAddActPlacement440_45") = "надійшло з інших в/ч (підрозділів)"
   Case 2
    ThisWorkbook.ActiveSheet.Range("ContentAddActPlacement440_45") = "надійшло зі складу в/ч"
   Case 3
    ThisWorkbook.ActiveSheet.Range("ContentAddActPlacement440_45") = "інші надходження"
    Case 4
    ThisWorkbook.ActiveSheet.Range("ContentAddActPlacement440_45") = "витрата"
    Case 5
    ThisWorkbook.ActiveSheet.Range("ContentAddActPlacement440_45") = "передано до інших в/ч (підрозділів)"
    Case 6
    ThisWorkbook.ActiveSheet.Range("ContentAddActPlacement440_45") = "інші вибуття"
   End Select
   
  
    
    'збільшуємо номер документа на 1
Dim clearNumber As Long
'=======3.9.19=======
Dim increase As String
increase = ThisWorkbook.Sheets("Акт закл.(освіж.)440_45").Range("increaseActPlacement440_45").value
If increase = "так" Then
    clearNumber = ThisWorkbook.Sheets("Акт закл.(освіж.)440_45").Range("clearNumberActPlacement440_45").value

    ThisWorkbook.Sheets("Акт закл.(освіж.)440_45").Range("clearNumberActPlacement440_45") = clearNumber + 1
End If
'=======3.9.19=======
    
    
    
    
    
'    'збільшуємо номер документа на 1
'    Dim clearNumber As Long
'    clearNumber = ThisWorkbook.Sheets("Акт закл.(освіж.)440_45").Range("clearNumberActPlacement440_45").value
'    ThisWorkbook.Sheets("Акт закл.(освіж.)440_45").Range("clearNumberActPlacement440_45") = clearNumber + 1
    
'    'забираємо суму і кількість
'    Dim allCount As String, allSum As String
'    allCount = ThisWorkbook.Sheets("Акт закл.(освіж.)440_45").Range("allCountActPlacement440_45").value
'    allSum = ThisWorkbook.Sheets("Акт закл.(освіж.)440_45").Range("allSumActPlacement440_45").value
    
    ThisWorkbook.Sheets("Акт закл.(освіж.)440_45").Copy 'зберігаємо як нову книгу
    Dim WbNew As Workbook
    Set WbNew = ActiveWorkbook
    
    'замінюємо формули на значення
    Dim cell As Range
    For Each cell In WbNew.Worksheets(1).Range(Cells(1, 1), Cells(WbNew.Worksheets(1).Range("bookkeepingActPlacement440_45").Row, 16))
        If cell.HasFormula Then
            cell.value = cell.value
        End If
    Next cell
    
'    WbNew.Worksheets(1).Range("allCountActPlacement440_45") = allCount
'    WbNew.Worksheets(1).Range("allSumActPlacement440_45") = allSum
    
    'реквізити для назви кінцевого файлу
    Dim NumberActPlacement440_45 As String, FromUnit As String, toUnit As String, fromMRP As String, toMRP As String
    NumberActPlacement440_45 = WbNew.Worksheets(1).Range("NumberActPlacement440_45")
    'оскільки зміна якості відбувається в одному підрозділі, то тільки 1 підрозділ
    
    FromUnit = WbNew.Worksheets(1).Range("FromUnitActPlacement440_45")
'    toUnit = WbNew.Worksheets(1).Range("toUnitActPlacement440_45")
    '=================3.9.18
    'оскільки зміна якості відбувається в одному підрозділі, то тільки 1 МВО
    fromMRP = WbNew.Worksheets(1).Range("MRPActPlacement440_45")
'    toMRP = WbNew.Worksheets(1).Range("MRPToActPlacement440_45")
    '=================
    
    Dim shp As Shape
    For Each shp In WbNew.Worksheets(1).Shapes
        shp.Delete
    Next shp
    
    Dim fileName As String, SavePath As String
    Dim DateActPlacement440_45 As Date
    DateActPlacement440_45 = WbNew.Worksheets(1).Range("DateActPlacement440_45")
    
    
    Dim typeDocActPlacement440_45 As String
        typeDocActPlacement440_45 = WbNew.Worksheets(1).Range("typeDocActPlacement440_45")
    'Спочатку беремо шлях з комірки
    SavePath = WbNew.Worksheets(1).Range("PathToSaveActPlacement440_45").value
    
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
                ThisWorkbook.Sheets("Акт закл.(освіж.)440_45").Range("PathToSaveActPlacement440_45").value = SavePath
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
    fileName = SavePath & Application.PathSeparator & CleanFileName(DateActPlacement440_45 & _
        " " & typeDocActPlacement440_45 & " " & NumberActPlacement440_45 & " " & FromUnit & " (" & fromMRP & ").xlsx")
    '========
    'Зберігаємо
    WbNew.SaveAs fileName:=fileName, FileFormat:=xlOpenXMLWorkbook
    
    Application.DisplayAlerts = True
End Sub


'============3.9.16=========
Sub addFromTableActPlacement440_45()

'Dim lo As ListObject
'Set lo = Worksheets("База").ListObjects("AllData")

Dim ws As Worksheet
Set ws = Selection.Worksheet 'лист, де був клік

Dim wsDoc As Worksheet
Set wsDoc = ThisWorkbook.Worksheets("Акт закл.(освіж.)440_45")
Dim tbActual As ListObject

Set tbActual = wsDoc.ListObjects("tbActPlacement440_45")

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
                wsDoc.Range("MRPActPlacement440_45") = MRP
            End If
            
            
            
            
        wsDoc.Range("MRPActPlacement440_45") = MRP
    
        'додаємо інформацію в документ
            If tbActual.DataBodyRange.Rows(1).Cells(1, 2) = "" Then
                With tbActual.DataBodyRange.Rows
                    .Cells(1, 2) = NameOfThings
                    .Cells(1, 3) = Nomenclature
                    .Cells(1, 4) = unit
                    .Cells(1, 5) = Category
                    .Cells(1, 7) = price
                    .Cells(1, 6) = CDbl(quantity)
                    .Cells(1, 10) = Note
                    .Cells(1, 18) = CategoryTo
                    
                    .Cells(1, 19) = Note2
                    .Cells(1, 20) = Note3
                End With
            Else
                    Dim newRow As ListRow
                    Set newRow = tbActual.ListRows.Add
                With newRow.Range
                    .Cells(1, 2) = NameOfThings
                    .Cells(1, 3) = Nomenclature
                    .Cells(1, 4) = unit
                    .Cells(1, 5) = Category
                    .Cells(1, 7) = price
                    .Cells(1, 6) = CDbl(quantity)
                    .Cells(1, 10) = Note
                    .Cells(1, 18) = CategoryTo
                    
                    .Cells(1, 19) = Note2
                    .Cells(1, 20) = Note3
                    
                End With
            End If

    listOfThings = listOfThings & NameOfThings & " " & CDbl(quantity) & " " & unit & vbNewLine
    
    
    
    Next rowIndex

MsgBox listOfThings, vbInformation, "Майно додане в Акт закладання(освіження)"

On Error Resume Next
    Application.CommandBars("Cell").Controls("Додати в документ").Delete
On Error GoTo 0


End Sub
'============3.9.16=========





```

---

## 7. RequestM.bas — Модуль формування внутрішніх заявок
### 7.1. Призначення
Служить для формування заявок/запитів підрозділів на поповнення або отримання майна зі складу. Таблиця `tbRequest` на аркуші «Заявка». Позиції підбираються зі складу залишків за допомогою форми вибору.

### 7.2. Вихідний код RequestM.bas
```vba
Attribute VB_Name = "RequestM"

'Request

Sub addPosRequest()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Заявка")
    
    Dim NameOfThings As String

    Dim tbRequest As ListObject
    Set tbRequest = ws.ListObjects("tbRequest")
    
    
    
'=============3.9.15
If IsError(ws.Range("FromUnitRequest").value) Then
    MsgBox "Схоже ви не заповнили довідник і/або не вибрали хто буде передавати(здавати, видавати) майно." _
    & vbNewLine & vbNewLine & "Заповінсть довідник і виберіть, хто буде ПЕРДАВАТИ майно.", _
    vbCritical, "Заповніть довідник, виберіть хто буде передавати"
    
    ws.Range("MRPRequest").Activate
    Exit Sub
End If

'If IsError(ws.Range("toUnitRequest").value) Then
'    MsgBox "Схоже ви не заповнили довідник і/або не вибрали хто буде отримувати (приймати) майно." _
'    & vbNewLine & vbNewLine & "Заповінсть довідник і виберіть, хто буде ОТРИМУВАТИ майно.", _
'    vbCritical, "Заповніть довідник, виберіть отримувачів"
'
'    ws.Range("MRPToRequest").Activate
'    Exit Sub
'End If
'=============3.9.15

''=============3.9.19
'Dim answer As VbMsgBoxResult
'
'answer = MsgBox("Додати в частину 'Закладено' - 'Так'" & vbNewLine & "Додати в частину 'Вилучено' - 'Ні'", vbYesNo + vbQuestion, "Підтвердження")
'
''=============3.9.19

    
'=============3.9.11
'визначаємо підрозділ-відправник, беремо з документа
Dim FromUnit As String
FromUnit = ws.Range("FromUnitRequest")

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
            Dim Service As String
            Dim QuantityRest As Double
            Dim Note As String
            
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
                  
            
            
'            Category = InputBox("Категорія відправника: ", NameOfThings, rngWithName.Offset(, -8).value)
'
'            CategoryTo = InputBox("Категорія отримувача: ", NameOfThings, rngWithName.Offset(, -8).value)
'            VAT = InputBox("ПДВ (за наявності): ", NameOfThings, 0)
            Note = InputBox("Примітка", NameOfThings, tbRequest.ListColumns(5).DataBodyRange.Cells(tbRequest.ListRows.count, 1).value)
            Service = InputBox("Служба забезпечення", NameOfThings, tbRequest.ListColumns(6).DataBodyRange.Cells(tbRequest.ListRows.count, 1).value)
'            Note3 = InputBox("Додаткове поле 3", NameOfThings, rngWithName.Offset(, -3).value)
            
            
            Dim newRow As ListRow
            
            For i = 1 To tbRequest.ListRows.count
                    If tbRequest.DataBodyRange.Cells(i, 2).value = "" Then
                        Set newRow = tbRequest.ListRows(i)
                        Exit For
                    End If
                Next i
                ' Якщо порожній рядок не знайдено – додаємо новий
                If newRow Is Nothing Then
                    Set newRow = tbRequest.ListRows.Add
                End If
            
'           If answer = vbYes Then
                With newRow.Range
                    .Cells(1, 2) = name
'                    .Cells(1, 3) = Nomenclature
                    .Cells(1, 3) = unit
'                    .Cells(1, 5) = Category
'                    .Cells(1, 7) = Price
                    .Cells(1, 4) = CDbl(quantity)
                    .Cells(1, 5) = Note
'                    .Cells(1, 18) = CategoryTo
                    
                    .Cells(1, 6) = Service
'                    .Cells(1, 20) = Note3
                    
                End With
'            End If
            
'            Else
'
'                With NewRow.Range
'                    .Cells(1, 2) = name
'                    .Cells(1, 3) = Nomenclature
'                    .Cells(1, 4) = unit
'                    .Cells(1, 11) = Category
'                    .Cells(1, 13) = Price
'                    .Cells(1, 12) = CDbl(quantity)
'                    .Cells(1, 16) = Note
'                    .Cells(1, 21) = CategoryTo
'
'                    .Cells(1, 22) = Note2
'                    .Cells(1, 23) = Note3
'
'                End With
'
'
'            End If
            Set newRow = Nothing
            
            Dim AddNext
            
            AddNext = MsgBox("Майно додано! Додати наступну позицію?", vbYesNo, NameOfThings)
            
        End If
        
    Loop While AddNext = 6
    
    
    '***********Для майна, яке оприбутковується***********
 
End Sub

Sub clearTableRequest()
Dim tbActual As ListObject
Set tbActual = ThisWorkbook.Worksheets("Заявка").ListObjects("tbRequest")
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
        If j <> 1 Then
            firstRow.Range.Cells(1, j).ClearContents
        End If
 Next j
 
End Sub


Sub addRowRequest()
Dim tbActual As ListObject
Set tbActual = ThisWorkbook.Worksheets("Заявка").ListObjects("tbRequest")

Dim newRow As ListRow
Set newRow = tbActual.ListRows.Add

End Sub




Sub SaveRequest()
    Application.DisplayAlerts = False
    
' Dim Basis As String
' Basis = InputBox("Вкажіть на підставі якого документу або події оформлюється даний документ ", "Підстава (мета) операції:", ThisWorkbook.ActiveSheet.Range("basisRequest").value)
' ThisWorkbook.ActiveSheet.Range("basisRequest") = Basis
'
'
'
'    Dim typeOperation As String
'
'    typeOperation = InputBox("Тип операції" & vbNewLine & "надходження - 1" _
'                                            & vbNewLine & "внутрішнє переміщення - 2" _
'                                            & vbNewLine & "вибуття - 3", "Тип операції", 3)
'
'    Select Case typeOperation
'        Case 1
'            ThisWorkbook.Sheets("Заявка").Range("ContentRequest") = "надходження"
'        Case 2
'            ThisWorkbook.Sheets("Заявка").Range("ContentRequest") = "внутрішнє переміщення"
'        Case 3
'            ThisWorkbook.Sheets("Заявка").Range("ContentRequest") = "вибуття"
'    End Select
'
'
'     'запитуємо додатковий тип операції
'   typeOperationAdd = InputBox("Детальний тип операції" & vbNewLine & "надійшло з інших в/ч (підрозділів) - 1" _
'                                           & vbNewLine & "надійшло зі складу в/ч - 2" _
'                                           & vbNewLine & "інші надходження - 3" _
'                                           & vbNewLine & "витрата - 4" _
'                                           & vbNewLine & "передано до інших в/ч (підрозділів)- 5" _
'                                           & vbNewLine & "інші вибуття - 6", "Тип операції", 4)
'Select Case typeOperationAdd
'   Case 1
'   ThisWorkbook.ActiveSheet.Range("ContentAddRequest") = "надійшло з інших в/ч (підрозділів)"
'   Case 2
'    ThisWorkbook.ActiveSheet.Range("ContentAddRequest") = "надійшло зі складу в/ч"
'   Case 3
'    ThisWorkbook.ActiveSheet.Range("ContentAddRequest") = "інші надходження"
'    Case 4
'    ThisWorkbook.ActiveSheet.Range("ContentAddRequest") = "витрата"
'    Case 5
'    ThisWorkbook.ActiveSheet.Range("ContentAddRequest") = "передано до інших в/ч (підрозділів)"
'    Case 6
'    ThisWorkbook.ActiveSheet.Range("ContentAddRequest") = "інші вибуття"
'   End Select
'
'
    
    'збільшуємо номер документа на 1
Dim clearNumber As Long
'=======3.9.19=======
Dim increase As String
increase = ThisWorkbook.Sheets("Заявка").Range("increaseRequest").value
If increase = "так" Then
    clearNumber = ThisWorkbook.Sheets("Заявка").Range("clearNumberRequest").value

    ThisWorkbook.Sheets("Заявка").Range("clearNumberRequest") = clearNumber + 1
End If
'=======3.9.19=======
    
    
    
    ThisWorkbook.Sheets("Заявка").Copy 'зберігаємо як нову книгу
    Dim WbNew As Workbook
    Set WbNew = ActiveWorkbook
    
    'замінюємо формули на значення
    Dim cell As Range
    For Each cell In WbNew.Worksheets(1).Range(Cells(1, 1), Cells(WbNew.Worksheets(1).Range("bookkeepingRequest").Row, 16))
        If cell.HasFormula Then
            cell.value = cell.value
        End If
    Next cell
    
    
    'реквізити для назви кінцевого файлу
    Dim NumberRequest As String, FromUnit As String, toUnit As String, fromMRP As String, toMRP As String
    NumberRequest = WbNew.Worksheets(1).Range("NumberRequest")
    'оскільки зміна якості відбувається в одному підрозділі, то тільки 1 підрозділ
    
    FromUnit = WbNew.Worksheets(1).Range("FromUnitRequest")
'    toUnit = WbNew.Worksheets(1).Range("toUnitRequest")
    '=================3.9.18
    'оскільки зміна якості відбувається в одному підрозділі, то тільки 1 МВО
    fromMRP = WbNew.Worksheets(1).Range("MRPRequest")
'    toMRP = WbNew.Worksheets(1).Range("MRPToRequest")
    '=================
    
    Dim shp As Shape
    For Each shp In WbNew.Worksheets(1).Shapes
        shp.Delete
    Next shp
    
    Dim fileName As String, SavePath As String
    Dim DateRequest As Date
    DateRequest = WbNew.Worksheets(1).Range("DateRequest")
    
    
    Dim typeDocRequest As String
        typeDocRequest = WbNew.Worksheets(1).Range("typeDocRequest")
    'Спочатку беремо шлях з комірки
    SavePath = WbNew.Worksheets(1).Range("PathToSaveRequest").value
    
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
                ThisWorkbook.Sheets("Заявка").Range("PathToSaveRequest").value = SavePath
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
    fileName = SavePath & Application.PathSeparator & CleanFileName(DateRequest & _
        " " & typeDocRequest & " " & NumberRequest & " " & FromUnit & " (" & fromMRP & ").xlsx")
    '========
    'Зберігаємо
    WbNew.SaveAs fileName:=fileName, FileFormat:=xlOpenXMLWorkbook
    
    Application.DisplayAlerts = True
End Sub


'============3.9.16=========
Sub addFromTableRequest()

'Dim lo As ListObject
'Set lo = Worksheets("База").ListObjects("AllData")

Dim ws As Worksheet
Set ws = Selection.Worksheet 'лист, де був клік

Dim wsDoc As Worksheet
Set wsDoc = ThisWorkbook.Worksheets("Заявка")
Dim tbActual As ListObject

Set tbActual = wsDoc.ListObjects("tbRequest")

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


Case "База заявок"
        colName = 2
        colQty1 = 4
'        colQty2 = 4
'        colNomenclature = 8
        colUnit = 3
'        colCategory = 6
'        colPrice = 3
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

'qKategory = MsgBox("Запитувати категорію кожного найменування?(Натисніть - так)" & vbNewLine & _
'"Категорія не буде змінюватись, тому можна взяти з поточного листа.(Натисніть - ні)", vbYesNo, "Категорія кожного найменування?")


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
            
            
'            Nomenclature = rw.Cells(1, colNomenclature)
             If colNomenclature = 0 Then
            Else
                Nomenclature = rw.Cells(1, colNomenclature)
            End If
            
            unit = rw.Cells(1, colUnit)
            
'            If qKategory = vbYes Then
'                Category = InputBox("Категорія відправника:", NameOfThings, rw.Cells(1, colCategory))
'                CategoryTo = InputBox("Категорія отримувача:", NameOfThings, rw.Cells(1, colCategory))
'            Else
'                Category = rw.Cells(1, colCategory)
'                CategoryTo = rw.Cells(1, colCategory)
'            End If
            
            
            
'            Price = rw.Cells(1, colPrice)
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
                wsDoc.Range("MRPRequest") = MRP
            End If
            
            
            
            
        wsDoc.Range("MRPRequest") = MRP
    
        
        
'        Dim NewRow As ListRow
'
'            For i = 1 To tbActual.ListRows.count
'                    If tbActual.DataBodyRange.Cells(i, 2).value = "" Then
'                        Set NewRow = tbActual.ListRows(i)
'                        Exit For
'                    End If
'                Next i
'                ' Якщо порожній рядок не знайдено – додаємо новий
'                If NewRow Is Nothing Then
'                    Set NewRow = tbActual.ListRows.Add
'                End If
        
        
        'додаємо інформацію в документ
            If tbActual.DataBodyRange.Rows(1).Cells(1, 2) = "" Then
                With tbActual.DataBodyRange.Rows
                    .Cells(1, 2) = NameOfThings
'                    .Cells(1, 3) = Nomenclature
                    .Cells(1, 3) = unit
'                    .Cells(1, 5) = Category
'                    .Cells(1, 7) = Price
                    .Cells(1, 4) = CDbl(quantity)
'                    .Cells(1, 10) = Note
'                    .Cells(1, 18) = CategoryTo
'
'                    .Cells(1, 19) = Note2
'                    .Cells(1, 20) = Note3
                End With
            Else
                    Dim newRow As ListRow
                    Set newRow = tbActual.ListRows.Add
                With newRow.Range
                    .Cells(1, 2) = NameOfThings
'                    .Cells(1, 3) = Nomenclature
                    .Cells(1, 3) = unit
'                    .Cells(1, 5) = Category
'                    .Cells(1, 7) = Price
                    .Cells(1, 4) = CDbl(quantity)
'                    .Cells(1, 10) = Note
'                    .Cells(1, 18) = CategoryTo
                    
'                    .Cells(1, 19) = Note2
'                    .Cells(1, 20) = Note3
                    
                End With
            End If

    listOfThings = listOfThings & NameOfThings & " " & CDbl(quantity) & " " & unit & vbNewLine
    
    
    
    Next rowIndex

MsgBox listOfThings, vbInformation, "Майно додане в документ"

On Error Resume Next
    Application.CommandBars("Cell").Controls("Додати в документ").Delete
On Error GoTo 0


End Sub
'============3.9.16=========







```
