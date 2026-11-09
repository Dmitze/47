# 06 — Модулі первинних документів: Акти списання та Накладні

У цьому документі детально описано модулі генерації та обліку основних первинних документів:
1. **WOff431.bas** — Акт списання майна (наказ МОУ № 431)
2. **WOff81.bas** — Акт списання майна (наказ МОУ № 81, форма 3)
3. **WOffLaunch81_19.bas** — Акт введення в експлуатацію / виведення з експлуатації ОЗ (наказ МОУ № 81, форма 19)
4. **invoice431.bas** — Накладна на видачу/передачу майна (наказ МОУ № 431)
5. **Invoice440_25.bas** — Накладна на внутрішнє переміщення (наказ МОУ № 440, форма 25)
6. **payroll440_57.bas** — Роздавальна/здавальна відомість (наказ МОУ № 440, форма 57)

## Спільний архітектурний шаблон первинного документа

Кожен модуль реалізує три взаємопов'язані сценарії:
1. **`addPos...()` (Поштучне додавання):** Користувач обирає підрозділ, система отримує доступний пул залишків через `GetResidusePool`, відкриває пошуковий діалог `UserForm1`, запитує кількість, категорію, підставу і додає новий рядок у відповідний `ListObject`.
2. **`addFromTable...()` (Масове перенесення):** Користувач виділяє рядки в таблиці `AllData` або на аркуші залишків і через праве контекстне меню переносить вибрані позиції безпосередньо в активний документ.
3. **`Save...()` (Збереження та архівація):** Автоматично інкрементує лічильник номера, копіює аркуш у чисту книгу Excel, видаляє формули (лишаючи тільки значення), видаляє кнопки та зберігає файл у робочий каталог із стандартизованою назвою.

---

## 1. WOff431.bas — Акт списання (наказ 431)
### 1.1. Реквізити та таблиці
- Аркуш: «Акт списання»
- Таблиця: `tbWriteOff`
- Іменовані діапазони: `invNumWOff`, `dateWOff`, `FromUnitWOff`, `fromMRPWOff`, `toUnitInv431`, `allSumWOff`, `allCountWOff`

### 1.2. Вихідний код WOff431.bas
```vba
Attribute VB_Name = "WOff431"
Sub copyFromCurentBookWrightOff()
 
    Dim NameOfThings As String


    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Акт списання")
    Dim tbWriteOff As ListObject
    Set tbWriteOff = ws.ListObjects("tbWriteOff")
    
    '=============3.9.15
If IsError(ws.Range("fromUnit").value) Then
    MsgBox "Схоже ви не заповнили довідник і/або не вибрали хто буде передавати(здавати, видавати) майно." _
    & vbNewLine & vbNewLine & "Заповінсть довідник і виберіть, хто буде ПЕРДАВАТИ майно.", _
    vbCritical, "Заповніть довідник, виберіть хто буде передавати"

    ws.Range("MRP").Activate
    Exit Sub
End If
'
'If IsError(ws.Range("toUnitInv431").value) Then
'    MsgBox "Схоже ви не заповнили довідник і/або не вибрали хто буде отримувати (приймати) майно." _
'    & vbNewLine & vbNewLine & "Заповінсть довідник і виберіть, хто буде ОТРИМУВАТИ майно.", _
'    vbCritical, "Заповніть довідник, виберіть отримувачів"
'
'    ws.Range("MRPToInv431").Activate
'    Exit Sub
'End If
'=============3.9.15
     
'=============3.9.11
'визначаємо підрозділ-відправник, беремо з документа
Dim FromUnit As String
FromUnit = ws.Range("fromUnit")

Dim tempPoolOThings As Range
Set tempPoolOThings = GetResidusePool(wsResiduse2, FromUnit)

If tempPoolOThings Is Nothing Then Exit Sub
'=============3.9.11

    
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
            Dim SubAccountKt As String
            Dim price As Double
            Dim QuantityRest As Double
            Dim SubAccountDt As String
            Dim Basis As String
            Dim Place As String
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
            
            SubAccount = rngWithName.Offset(, -7).value
            SubAccountDt = InputBox("Субрахунок для списання, залишити 1 (якщо не знаєте який - залишіть порожнім): ", NameOfThings, "6611 8013 8113")
            
            quantity = InputBox("Кількість: ", NameOfThings, QuantityRest)
            Basis = InputBox("Підстава для списання: ", NameOfThings)

'            NoteDt = InputBox("Субрахунок для списання, залишити 1 (якщо не знаєте який - залишіть порожнім): ", NameOfThings, "6611 8013 8113")
            
            Category = InputBox("Категорія відправника: ", NameOfThings, rngWithName.Offset(, -8).value)
'            CategoryTo = InputBox("Категорія отримувача: ", NameOfThings, rngWithName.Offset(, -8).value)
            CategoryTo = InputBox("Категорія для списання (категорія отримувача)не обов'язково: ", NameOfThings)
'            VAT = InputBox("ПДВ (за наявності): ", NameOfThings, 0)
'            Note1 = InputBox("Додаткове поле 1", NameOfThings, rngWithName.Offset(, -5).value)
            Note2 = InputBox("Додаткове поле 2", NameOfThings, rngWithName.Offset(, -4).value)
            Note3 = InputBox("Додаткове поле 3", NameOfThings, rngWithName.Offset(, -3).value)
            
          
            If tbWriteOff.DataBodyRange.Rows(1).Cells(1, 2) = "" Then
                With tbWriteOff.DataBodyRange.Rows
                    .Cells(1, 2) = name
                    .Cells(1, 3) = Nomenclature
                    .Cells(1, 4) = unit
                    .Cells(1, 5) = CDbl(quantity)
                    .Cells(1, 6) = CDbl(price)
                    .Cells(1, 10) = SubAccount
                    .Cells(1, 9) = SubAccountDt
                    .Cells(1, 8) = Basis
                    .Cells(1, 11) = Category
'                    .Cells(1, 12) = CategoryFrom
                    .Cells(1, 12) = CategoryTo
'                    .Cells(1, 13) = Note1
                    .Cells(1, 14) = Note2
                    .Cells(1, 15) = Note3
                End With
            Else
                    Dim newRow As ListRow
                    Set newRow = tbWriteOff.ListRows.Add
                With newRow.Range
                    .Cells(1, 2) = name
                    .Cells(1, 3) = Nomenclature
                    .Cells(1, 4) = unit
                    .Cells(1, 5) = CDbl(quantity)
                    .Cells(1, 6) = price
                    .Cells(1, 10) = SubAccount
                    .Cells(1, 9) = SubAccountDt
                    .Cells(1, 8) = Basis
                    .Cells(1, 11) = Category
                    .Cells(1, 12) = CategoryTo
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






Sub addRow()
Dim tbWriteOff As ListObject
Set tbWriteOff = ThisWorkbook.Worksheets("Акт списання").ListObjects("tbWriteOff")

Dim newRow As ListRow
Set newRow = tbWriteOff.ListRows.Add

End Sub
Sub clearTable()
Dim tbWriteOff As ListObject
Set tbWriteOff = ThisWorkbook.Worksheets("Акт списання").ListObjects("tbWriteOff")
Dim i As Long
If tbWriteOff.ListRows.count > 1 Then
                ' Видаляємо всі рядки, крім першого
                For i = tbWriteOff.ListRows.count To 2 Step -1
                    tbWriteOff.ListRows(i).Delete
                Next i
End If

Dim firstRow As ListRow
Set firstRow = tbWriteOff.ListRows(1)
Dim j As Long
 For j = 1 To tbWriteOff.ListColumns.count
        If j <> 1 And j <> 7 Then
            firstRow.Range.Cells(1, j).ClearContents
        End If
 Next j
 
End Sub

Sub SaveWO()

Dim typeOperation As String

  
'збільшуємо номер акта на 1
Dim clearNumber As Long

clearNumber = ThisWorkbook.Sheets("Акт списання").Range("clearNumberWOff").value

ThisWorkbook.Sheets("Акт списання").Range("clearNumberWOff") = clearNumber + 1

'забираємо суму і кількість
Dim allCountWOff As String
Dim allSumWOff As String

allCountWOff = ThisWorkbook.Sheets("Акт списання").Range("allCountWOff").value
allSumWOff = ThisWorkbook.Sheets("Акт списання").Range("allSumWOff").value

        ThisWorkbook.Sheets("Акт списання").Select 'копіюємо перший лист з нарядом
        ThisWorkbook.Sheets("Акт списання").Copy 'зберігаємо як нову книгу
        

    
     Dim WbNewWO As Workbook
     
        Set WbNewWO = ActiveWorkbook
   
     'замінюємо на значення клітинки, які підтягуються по формулі з іншого листа, щоб у збережених документах не злітали дані
     
      For Each cell In WbNewWO.Worksheets(1).Range(Cells(1, 1), Cells(WbNewWO.Worksheets(1).Range("bookkeeping").Row, 15))
        If cell.HasFormula Then
            cell.value = cell.value
        End If
     Next cell
     
    WbNewWO.Worksheets(1).Range("allCountWOff") = allCountWOff
    WbNewWO.Worksheets(1).Range("allSumWOff") = allSumWOff
    
   Dim NumberWOff As String
     NumberWOff = WbNewWO.Worksheets(1).Range("numberWOff")
   Dim FromUnit As String
    FromUnit = WbNewWO.Worksheets(1).Range("fromUnit")
    
    '=================3.9.11
   Dim fromMRP  As String
    
    fromMRP = WbNewWO.Worksheets(1).Range("MRP")
   '=================
   
    Dim shp As Shape
    For Each shp In WbNewWO.Worksheets(1).Shapes
        shp.Delete
    Next shp
        
      Dim fileName As String, fileName2 As String, lineBrake As String
        'Шлях до файла зберігання
        SavePath = WbNewWO.Worksheets(1).Range("PathToSaveWiteOff")
        Dim DateWOff As Date
        DateWOff = WbNewWO.Worksheets(1).Range("DateWOff")
        
        
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
                ThisWorkbook.Sheets("Акт списання").Range("PathToSaveWiteOff").value = SavePath
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
        fileName = SavePath & Application.PathSeparator & CleanFileName(DateWOff & _
        " Акт спис. запасів № " & NumberWOff & " з " & FromUnit & " (" & fromMRP & ").xlsx")
         '========

        'зберігаємо файл
        On Error GoTo WrongPath
         WbNewWO.SaveAs fileName:=fileName, FileFormat:=xlOpenXMLWorkbook
    On Error GoTo 0
      Exit Sub
WrongPath:
      MsgBox "Ймовірно, не правильно вказана адреса збереження файла, виправте адресу збереження файлу( на зразок ""D:\База обліку майна ЗСУ 2024"" *без косої лінії вкінці) і спробуйтие зберегти файл ще раз!", vbCritical, "Не правильний шлях збереження"
     Exit Sub
End Sub






```

---

## 2. WOff81.bas — Акт списання (наказ 81, ф. 3)
### 2.1. Реквізити та таблиці
- Аркуш: «Акт списання 81»
- Таблиця: `tbWriteOff81`
- Іменовані діапазони: `invNumWOff81`, `dateWOff81`, `FromUnitWOff81`, `fromMRPWOff81`, `toUnitWOff81`, `basisWOff81`, `ServiceWOff81`, `allCountWOff81`, `allSumWOff81`

### 2.2. Вихідний код WOff81.bas
```vba
Attribute VB_Name = "WOff81"
Sub addPosWOff81()
 
    Dim NameOfThings As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Акт списання 81")
    Dim tbWOff81 As ListObject
    Set tbWOff81 = ws.ListObjects("tbWriteOff81")
    
'=============3.9.11
'визначаємо підрозділ-відправник, беремо з документа
Dim FromUnit As String
FromUnit = ws.Range("FromUnitWOff81")

Dim tempPoolOThings As Range
Set tempPoolOThings = GetResidusePool(wsResiduse2, FromUnit)

If tempPoolOThings Is Nothing Then Exit Sub
'=============3.9.11

 Dim Basis As String
 Basis = InputBox("Підстава (мета) операції: ", NameOfThings, ThisWorkbook.Worksheets("Акт списання 81").Range("basisWOff81").value)
 ThisWorkbook.Worksheets("Акт списання 81").Range("basisWOff81") = Basis
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
            If tbWOff81.DataBodyRange.Rows(1).Cells(1, 2) = "" Then
                With tbWOff81.DataBodyRange.Rows
                    .Cells(1, 2) = name
                    .Cells(1, 3) = Nomenclature
                    .Cells(1, 4) = unit
                    
                    .Cells(1, 5) = Category
                    .Cells(1, 6) = price
                    .Cells(1, 7) = CDbl(quantity)
'                    .Cells(1, 9) = VAT
                    
                    .Cells(1, 14) = Note
'                    .Cells(1, 13) = Note
'                    .Cells(1, 8) = Basis
'                    .Cells(1, 10) = Quantity
                    .Cells(1, 18) = CategoryTo
'                    .Cells(1, 15) = Note1
                    .Cells(1, 16) = Note2
                    .Cells(1, 17) = Note3
                End With
            Else
                    Dim newRow As ListRow
                    Set newRow = tbWOff81.ListRows.Add
                With newRow.Range
                    .Cells(1, 2) = name
                    .Cells(1, 3) = Nomenclature
                    .Cells(1, 4) = unit
                    
                    .Cells(1, 5) = Category
                    .Cells(1, 6) = price
                    .Cells(1, 7) = CDbl(quantity)
'                    .Cells(1, 9) = VAT
                    
                    .Cells(1, 14) = Note
'                    .Cells(1, 13) = Note
'                    .Cells(1, 8) = Basis
'                    .Cells(1, 10) = Quantity
                    .Cells(1, 18) = CategoryTo
'                    .Cells(1, 15) = Note1
                    .Cells(1, 16) = Note2
                    .Cells(1, 17) = Note3
                    
                End With
            End If
           
            Dim AddNext
            
            AddNext = MsgBox("Майно додано! Додати наступну позицію?", vbYesNo, NameOfThings)
            
        End If
        
    Loop While AddNext = 6
    

End Sub

Sub addRow81()
Dim tbWriteOff As ListObject
Set tbWriteOff = ThisWorkbook.Worksheets("Акт списання 81").ListObjects("tbWriteOff81")

Dim newRow As ListRow
Set newRow = tbWriteOff.ListRows.Add

End Sub
Sub clearTable81()
Dim tbWriteOff As ListObject
Set tbWriteOff = ThisWorkbook.Worksheets("Акт списання 81").ListObjects("tbWriteOff81")
Dim i As Long
If tbWriteOff.ListRows.count > 1 Then
                ' Видаляємо всі рядки, крім першого
                For i = tbWriteOff.ListRows.count To 2 Step -1
                    tbWriteOff.ListRows(i).Delete
                Next i
End If

Dim firstRow As ListRow
Set firstRow = tbWriteOff.ListRows(1)
Dim j As Long
 For j = 1 To tbWriteOff.ListColumns.count
        If j <> 1 And j <> 8 And j <> 9 Then
            firstRow.Range.Cells(1, j).ClearContents
        End If
 Next j
 
End Sub


Sub SaveWOff81()
 
'збільшуємо номер акта на 1
Dim clearNumber As Long

clearNumber = ThisWorkbook.Sheets("Акт списання 81").Range("clearNumberWOff81").value

ThisWorkbook.Sheets("Акт списання 81").Range("clearNumberWOff81") = clearNumber + 1

'забираємо суму і кількість
'Dim allCount As String
'Dim allSum As String
'
'allCount = ThisWorkbook.Sheets("Акт списання 81").Range("allCountWOff81").value
'allSum = ThisWorkbook.Sheets("Акт списання 81").Range("allSumWOff81").value

        ThisWorkbook.Sheets("Акт списання 81").Select 'копіюємо перший лист з нарядом
        ThisWorkbook.Sheets("Акт списання 81").Copy 'зберігаємо як нову книгу
   
     Dim WbNew As Workbook
     
        Set WbNew = ActiveWorkbook
   
     'замінюємо на значення клітинки, які підтягуються по формулі з іншого листа, щоб у збережених документах не злітали дані
     
      For Each cell In WbNew.Worksheets(1).Range(Cells(1, 1), Cells(WbNew.Worksheets(1).Range("bookkeepingWOff81").Row, 20))
        If cell.HasFormula Then
            cell.value = cell.value
        End If
     Next cell
     
'    WbNew.Worksheets(1).Range("allCountWOff81") = allCount
'    WbNew.Worksheets(1).Range("allSumWOff81") = allSum
    
   Dim NumberWOff81 As String
     NumberWOff81 = WbNew.Worksheets(1).Range("NumberWOff81")
     
   Dim FromUnit As String
   Dim toUnit As String
        FromUnit = WbNew.Worksheets(1).Range("FromUnitWOff81")
        toUnit = WbNew.Worksheets(1).Range("toUnitWOff81")


'=================3.9.11
   Dim fromMRP  As String
    
    fromMRP = WbNew.Worksheets(1).Range("MRPWOff81")
'=================
        
        
    Dim shp As Shape
    For Each shp In WbNew.Worksheets(1).Shapes
        shp.Delete
    Next shp
        
      Dim fileName As String, fileName2 As String, lineBrake As String
        'Шлях до файла зберігання
        SavePath = WbNew.Worksheets(1).Range("PathToSaveWOff81")
        Dim DateWOff81 As Date
        DateWOff81 = WbNew.Worksheets(1).Range("DateWOff81")
        
        
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
                ThisWorkbook.Sheets("Акт списання 81").Range("PathToSaveWOff81").value = SavePath
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
        fileName = SavePath & Application.PathSeparator & CleanFileName(DateWOff81 & _
        " Акт списання № " & NumberWOff81 & " з " & FromUnit & " (" & fromMRP & ").xlsx")
        '========
        'зберігаємо файл
         WbNew.SaveAs fileName:=fileName, FileFormat:=xlOpenXMLWorkbook
    
     
End Sub

Sub addFromTableWOff81()

'Dim lo As ListObject
'Set lo = Worksheets("База").ListObjects("AllData")

Dim ws As Worksheet
Set ws = Selection.Worksheet 'лист, де був клік

Dim wsDoc As Worksheet
Set wsDoc = ThisWorkbook.Worksheets("Акт списання 81")
Dim tbActual As ListObject

Set tbActual = wsDoc.ListObjects("tbWriteOff81")

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
                wsDoc.Range("MRPWOff81") = MRP
            End If
            
            
            
            
        wsDoc.Range("MRPWOff81") = MRP
    
        'додаємо інформацію в документ
            If tbActual.DataBodyRange.Rows(1).Cells(1, 2) = "" Then
                With tbActual.DataBodyRange.Rows
                    .Cells(1, 2) = NameOfThings
                    .Cells(1, 3) = Nomenclature
                    .Cells(1, 4) = unit
                    .Cells(1, 5) = Category
                    .Cells(1, 6) = price
                    .Cells(1, 7) = CDbl(quantity)
                    .Cells(1, 14) = Note
                    .Cells(1, 18) = CategoryTo
                    
'                    .Cells(1, 12) = Note1
                    .Cells(1, 16) = Note2
                    .Cells(1, 17) = Note3
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
                    .Cells(1, 14) = Note
                    .Cells(1, 18) = CategoryTo
                    
'                    .Cells(1, 12) = Note1
                    .Cells(1, 16) = Note2
                    .Cells(1, 17) = Note3
                    
                End With
            End If

    listOfThings = listOfThings & NameOfThings & " " & CDbl(quantity) & " " & unit & vbNewLine
    
    
    
    Next rowIndex

MsgBox listOfThings, vbInformation, "Майно додане в Акт списання д.3"

On Error Resume Next
    Application.CommandBars("Cell").Controls("Додати в документ").Delete
On Error GoTo 0


End Sub


```

---

## 3. WOffLaunch81_19.bas — Введення / виведення з експлуатації ОЗ (ф. 19)
### 3.1. Реквізити та таблиці
- Аркуш: «Акт списання №19 (вОЗ)»
- Таблиця: `tbWOffLaunch81_19`
- Підтримує 6 видів операцій введення/передачі/списання ОЗ.

### 3.2. Вихідний код WOffLaunch81_19.bas
```vba
Attribute VB_Name = "WOffLaunch81_19"
Sub addPosWOffLaunch81_19()
  
    Dim NameOfThings As String
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Акт списання д19 (пуск)")

    Dim tbWOffLaunch81_19 As ListObject
    Set tbWOffLaunch81_19 = ws.ListObjects("tbWOffLaunch81_19")
    

'=============3.9.11
'визначаємо підрозділ-відправник, беремо з документа
Dim FromUnit As String
FromUnit = ws.Range("FromUnitWOffLaunch81_19")

Dim tempPoolOThings As Range
Set tempPoolOThings = GetResidusePool(wsResiduse2, FromUnit)

If tempPoolOThings Is Nothing Then Exit Sub
'=============3.9.11



' Dim Basis As String
' Basis = InputBox("Підстава (мета) операції: ", NameOfThings, ThisWorkbook.Worksheets("Акт списання д19 (пуск)").Range("basisWOffLaunch81_19").value)
' ThisWorkbook.Worksheets("Акт списання д19 (пуск)").Range("basisWOffLaunch81_19") = Basis
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
            If tbWOffLaunch81_19.DataBodyRange.Rows(1).Cells(1, 2) = "" Then
                With tbWOffLaunch81_19.DataBodyRange.Rows
                    .Cells(1, 2) = name
                    .Cells(1, 4) = Nomenclature
                    .Cells(1, 5) = unit
                    .Cells(1, 6) = Category
                    .Cells(1, 8) = price
                    .Cells(1, 7) = CDbl(quantity)
'                    .Cells(1, 9) = CDbl(quantity)
                    .Cells(1, 3) = Note
                    .Cells(1, 10) = CategoryTo
                    
'                    .Cells(1, 13) = Note1
                    .Cells(1, 11) = Note2
                    .Cells(1, 12) = Note3
                End With
            Else
                    Dim newRow As ListRow
                    Set newRow = tbWOffLaunch81_19.ListRows.Add
                With newRow.Range
                    .Cells(1, 2) = name
                    .Cells(1, 4) = Nomenclature
                    .Cells(1, 5) = unit
                    .Cells(1, 6) = Category
                    .Cells(1, 8) = price
                    .Cells(1, 7) = CDbl(quantity)
'                    .Cells(1, 9) = CDbl(quantity)
                    .Cells(1, 3) = Note
                    .Cells(1, 10) = CategoryTo
                    
'                    .Cells(1, 13) = Note1
                    .Cells(1, 11) = Note2
                    .Cells(1, 12) = Note3
                End With
            End If
           
            
            Dim AddNext
            
            AddNext = MsgBox("Майно додано! Додати наступну позицію?", vbYesNo, NameOfThings)
            
        End If
        
    Loop While AddNext = 6
    

End Sub

Sub clearTableWOffLaunch81_19()
Dim tbActual As ListObject
Set tbActual = ThisWorkbook.Worksheets("Акт списання д19 (пуск)").ListObjects("tbWOffLaunch81_19")
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
        If j <> 1 And j <> 9 Then
            firstRow.Range.Cells(1, j).ClearContents
        End If
 Next j
 
End Sub


Sub addRowWOffLaunch81_19()
Dim tbActual As ListObject
Set tbActual = ThisWorkbook.Worksheets("Акт списання д19 (пуск)").ListObjects("tbWOffLaunch81_19")

Dim newRow As ListRow
Set newRow = tbActual.ListRows.Add

End Sub

Sub SaveWOffLaunch81_19()

 Dim typeOperation As String
   
   typeOperation = InputBox("Тип операції" & vbNewLine & "надходження - 1" _
                                           & vbNewLine & "внутрішнє переміщення - 2" _
                                           & vbNewLine & "вибуття - 3", "Тип операції", 3)
   
   Select Case typeOperation
   Case 1
   ThisWorkbook.Sheets("Акт списання д19 (пуск)").Range("ContentWOffLaunch81_19") = "надходження"
   Case 2
    ThisWorkbook.Sheets("Акт списання д19 (пуск)").Range("ContentWOffLaunch81_19") = "внутрішнє переміщення"
   Case 3
    ThisWorkbook.Sheets("Акт списання д19 (пуск)").Range("ContentWOffLaunch81_19") = "вибуття"
   End Select
   
    typeOperationAdd = InputBox("Детальний тип операції" & vbNewLine & "надійшло з інших в/ч (підрозділів) - 1" _
                                           & vbNewLine & "надійшло зі складу в/ч - 2" _
                                           & vbNewLine & "інші надходження - 3" _
                                           & vbNewLine & "витрата - 4" _
                                           & vbNewLine & "передано до інших в/ч (підрозділів)- 5" _
                                           & vbNewLine & "інші вибуття - 6", "Тип операції", 4)
Select Case typeOperationAdd
   Case 1
   ThisWorkbook.ActiveSheet.Range("ContentAddWOffLaunch81_19") = "надійшло з інших в/ч (підрозділів)"
   Case 2
    ThisWorkbook.ActiveSheet.Range("ContentAddWOffLaunch81_19") = "надійшло зі складу в/ч"
   Case 3
    ThisWorkbook.ActiveSheet.Range("ContentAddWOffLaunch81_19") = "інші надходження"
    Case 4
    ThisWorkbook.ActiveSheet.Range("ContentAddWOffLaunch81_19") = "витрата"
    Case 5
    ThisWorkbook.ActiveSheet.Range("ContentAddWOffLaunch81_19") = "передано до інших в/ч (підрозділів)"
    Case 6
    ThisWorkbook.ActiveSheet.Range("ContentAddWOffLaunch81_19") = "інші вибуття"
   End Select
   
 '=======3.9.16=======
 Dim Basis As String
 Basis = InputBox("Дата пусків ", "Коли були здійснені пуски:", ThisWorkbook.Worksheets("Акт списання д19 (пуск)").Range("basisWOffLaunch81_19").value)
 ThisWorkbook.Worksheets("Акт списання д19 (пуск)").Range("basisWOffLaunch81_19") = Basis
 '=======3.9.16=======
   
   
'збільшуємо номер документа на 1
Dim clearNumber As Long
'=======3.9.16=======
Dim increase As String
increase = ThisWorkbook.Sheets("Акт списання д19 (пуск)").Range("increaseWOffLaunch81_19").value
If increase = "так" Then
    clearNumber = ThisWorkbook.Sheets("Акт списання д19 (пуск)").Range("clearNumberWOffLaunch81_19").value

    ThisWorkbook.Sheets("Акт списання д19 (пуск)").Range("clearNumberWOffLaunch81_19") = clearNumber + 1
End If
'=======3.9.16=======

'забираємо суму і кількість
Dim allCount As String
'Dim allSum As String
'
'allCount = ThisWorkbook.Sheets("Акт списання д19 (пуск)").Range("allSumAndQuantity440_24").value
'allSum = ThisWorkbook.Sheets("Акт списання д19 (пуск)").Range("allSumWOffLaunch81_19").value
'
        ThisWorkbook.Sheets("Акт списання д19 (пуск)").Select 'копіюємо лист з документом
        ThisWorkbook.Sheets("Акт списання д19 (пуск)").Copy 'зберігаємо як нову книгу
   
     Dim WbNew As Workbook
     
        Set WbNew = ActiveWorkbook
   
     'замінюємо на значення клітинки, які підтягуються по формулі з іншого листа, щоб у збережених документах не злітали дані
     
      For Each cell In WbNew.Worksheets(1).Range(Cells(1, 1), Cells(WbNew.Worksheets(1).Range("bookkeepingWOffLaunch81_19").Row, 16))
        If cell.HasFormula Then
            cell.value = cell.value
        End If
     Next cell
     
'    WbNew.Worksheets(1).Range("allSumAndQuantity440_24") = allCount
'    WbNew.Worksheets(1).Range("allSumWOffLaunch81_19") = allSum
    
   Dim NumberWOffLaunch81_19 As String
     NumberWOffLaunch81_19 = WbNew.Worksheets(1).Range("NumberWOffLaunch81_19")
     
   Dim FromUnit As String
   Dim toUnit As String
        FromUnit = WbNew.Worksheets(1).Range("FromUnitWOffLaunch81_19")
        toUnit = WbNew.Worksheets(1).Range("toUnitWOffLaunch81_19")
        
'=================3.9.11
   Dim fromMRP  As String
   Dim toMRP As String
    
    fromMRP = WbNew.Worksheets(1).Range("MRPFromWOffLaunch81_19")
'    toMRP = WbNew.Worksheets(1).Range("MRPToWOffLaunch81_19")
'==================
        
        
        
    Dim shp As Shape
    For Each shp In WbNew.Worksheets(1).Shapes
        shp.Delete
    Next shp
        
      Dim fileName As String, fileName2 As String, lineBrake As String
        'Шлях до файла зберігання
        SavePath = WbNew.Worksheets(1).Range("PathToSaveWOffLaunch81_19")
        Dim DateWOffLaunch81_19 As Date
        DateWOffLaunch81_19 = WbNew.Worksheets(1).Range("DateWOffLaunch81_19")
        
        Dim typeDocWOffLaunch81_19 As String
        typeDocWOffLaunch81_19 = WbNew.Worksheets(1).Range("typeDocWOffLaunch81_19")
        
        
        On Error Resume Next
         If Dir(SavePath, vbDirectory) = vbNullString Then
         On Error GoTo 0
        With Application.FileDialog(msoFileDialogFolderPicker)
            .Title = "Виберіть папку для збереження"
            .AllowMultiSelect = False
            If .show = -1 Then
                SavePath = .SelectedItems(1)
                'Записуємо новий шлях у клітинку
                ThisWorkbook.Sheets("Акт списання д19 (пуск)").Range("PathToSaveWOffLaunch81_19").value = SavePath
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
        fileName = SavePath & Application.PathSeparator & CleanFileName(DateWOffLaunch81_19 & _
        " " & typeDocWOffLaunch81_19 & " " & NumberWOffLaunch81_19 & " " & FromUnit & " (" & fromMRP & ") - " & toUnit & ".xlsx")
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

## 4. invoice431.bas — Накладна за наказом 431
### 4.1. Реквізити та таблиці
- Аркуш: «Накладна431»
- Таблиця: `tbInv431`
- Іменовані діапазони: `invNum431`, `dateInv431`, `FromUnitInv431`, `toUnitInv431`, `basis431`, `allSumInv431`, `allCountInv431`

### 4.2. Вихідний код invoice431.bas
```vba
Attribute VB_Name = "invoice431"
Sub addPosInv431()
  
    
    Dim NameOfThings As String
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Накладна431")

    Dim tbInv431 As ListObject
    Set tbInv431 = ws.ListObjects("tbInv431")
       '=============3.9.15
If IsError(ws.Range("FromUnitInv431").value) Then
    MsgBox "Схоже ви не заповнили довідник і/або не вибрали хто буде передавати(здавати, видавати) майно." _
    & vbNewLine & vbNewLine & "Заповінсть довідник і виберіть, хто буде ПЕРДАВАТИ майно.", _
    vbCritical, "Заповніть довідник, виберіть хто буде передавати"

    ws.Range("MRPInv431").Activate
    Exit Sub
End If

If IsError(ws.Range("toUnitInv431").value) Then
    MsgBox "Схоже ви не заповнили довідник і/або не вибрали хто буде отримувати (приймати) майно." _
    & vbNewLine & vbNewLine & "Заповінсть довідник і виберіть, хто буде ОТРИМУВАТИ майно.", _
    vbCritical, "Заповніть довідник, виберіть отримувачів"

    ws.Range("MRPToInv431").Activate
    Exit Sub
End If
'=============3.9.15
    
    
'=============3.9.11
'визначаємо підрозділ-відправник, беремо з документа
Dim FromUnit As String
FromUnit = ws.Range("FromUnitInv431")

Dim tempPoolOThings As Range
Set tempPoolOThings = GetResidusePool(wsResiduse2, FromUnit)

If tempPoolOThings Is Nothing Then Exit Sub
'=============3.9.11

 Dim Basis As String
 Basis = InputBox("Підстава (мета) операції: ", NameOfThings, ThisWorkbook.Worksheets("Накладна431").Range("basis431").value)
 ThisWorkbook.Worksheets("Накладна431").Range("basis431") = Basis
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
            VAT = InputBox("ПДВ (за наявності): ", NameOfThings, 0)
'            Note1 = InputBox("Додаткове поле 1", NameOfThings, rngWithName.Offset(, -5).value)
            Note2 = InputBox("Додаткове поле 2", NameOfThings, rngWithName.Offset(, -4).value)
            Note3 = InputBox("Додаткове поле 3", NameOfThings, rngWithName.Offset(, -3).value)
            If tbInv431.DataBodyRange.Rows(1).Cells(1, 2) = "" Then
                With tbInv431.DataBodyRange.Rows
                    .Cells(1, 2) = name
                    .Cells(1, 3) = Nomenclature
                    .Cells(1, 4) = unit
                    .Cells(1, 5) = CDbl(quantity)
                    .Cells(1, 7) = price
                    .Cells(1, 9) = VAT
                    
                    .Cells(1, 12) = Note
                    .Cells(1, 13) = Note
'                    .Cells(1, 8) = Basis
                    .Cells(1, 10) = Category
                    .Cells(1, 11) = CategoryTo
'                    .Cells(1, 14) = Note1
                    .Cells(1, 15) = Note2
                    .Cells(1, 16) = Note3
                End With
            Else
                    Dim newRow As ListRow
                    Set newRow = tbInv431.ListRows.Add
                With newRow.Range
                    .Cells(1, 2) = name
                    .Cells(1, 3) = Nomenclature
                    .Cells(1, 4) = unit
                    .Cells(1, 5) = CDbl(quantity)
                    .Cells(1, 7) = price
                    .Cells(1, 9) = VAT
                    
                    .Cells(1, 12) = Note
                    .Cells(1, 13) = Note
'                    .Cells(1, 8) = Basis
                    .Cells(1, 10) = Category
                    .Cells(1, 11) = CategoryTo
'                    .Cells(1, 14) = Note1
                    .Cells(1, 15) = Note2
                    .Cells(1, 16) = Note3
                    
                End With
            End If
           
            
            Dim AddNext
            
            AddNext = MsgBox("Майно додано! Додати наступну позицію?", vbYesNo, NameOfThings)
            
        End If
        
    Loop While AddNext = 6
    

End Sub

Sub clearTable431()
Dim tbActual As ListObject
Set tbActual = ThisWorkbook.Worksheets("Накладна431").ListObjects("tbInv431")
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
        If j <> 1 And j <> 6 And j <> 8 Then
            firstRow.Range.Cells(1, j).ClearContents
        End If
 Next j
 
End Sub


Sub addRow431()
Dim tbActual As ListObject
Set tbActual = ThisWorkbook.Worksheets("Накладна431").ListObjects("tbInv431")

Dim newRow As ListRow
Set newRow = tbActual.ListRows.Add

End Sub

Sub SaveInv431()

 Dim typeOperation As String
   
   typeOperation = InputBox("Тип операції" & vbNewLine & "надходження - 1" _
                                           & vbNewLine & "внутрішнє переміщення - 2" _
                                           & vbNewLine & "вибуття - 3", "Тип операції", 2)
   
   Select Case typeOperation
   Case 1
   ThisWorkbook.Sheets("Накладна431").Range("ContentInv431") = "надходження"
   Case 2
    ThisWorkbook.Sheets("Накладна431").Range("ContentInv431") = "внутрішнє переміщення"
   Case 3
    ThisWorkbook.Sheets("Накладна431").Range("ContentInv431") = "вибуття"
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
   ThisWorkbook.ActiveSheet.Range("ContentInvAdd431") = "надійшло з інших в/ч (підрозділів)"
   Case 2
    ThisWorkbook.ActiveSheet.Range("ContentInvAdd431") = "надійшло зі складу в/ч"
   Case 3
    ThisWorkbook.ActiveSheet.Range("ContentInvAdd431") = "інші надходження"
    Case 4
    ThisWorkbook.ActiveSheet.Range("ContentInvAdd431") = "витрата"
    Case 5
    ThisWorkbook.ActiveSheet.Range("ContentInvAdd431") = "передано до інших в/ч (підрозділів)"
    Case 6
    ThisWorkbook.ActiveSheet.Range("ContentInvAdd431") = "інші вибуття"
   End Select
'збільшуємо номер акта на 1
Dim clearNumber As Long

clearNumber = ThisWorkbook.Sheets("Накладна431").Range("clearNumberInv431").value

ThisWorkbook.Sheets("Накладна431").Range("clearNumberInv431") = clearNumber + 1

'забираємо суму і кількість
Dim allCount As String
Dim allSum As String

allCount = ThisWorkbook.Sheets("Накладна431").Range("allCountInv431").value
allSum = ThisWorkbook.Sheets("Накладна431").Range("allSumInv431").value

        ThisWorkbook.Sheets("Накладна431").Select 'копіюємо перший лист з нарядом
        ThisWorkbook.Sheets("Накладна431").Copy 'зберігаємо як нову книгу
        

    
     Dim WbNew As Workbook
     
        Set WbNew = ActiveWorkbook
   
     'замінюємо на значення клітинки, які підтягуються по формулі з іншого листа, щоб у збережених документах не злітали дані
     
      For Each cell In WbNew.Worksheets(1).Range(Cells(1, 1), Cells(WbNew.Worksheets(1).Range("bookkeepingInv431").Row, 16))
        If cell.HasFormula Then
            cell.value = cell.value
        End If
     Next cell
     
    WbNew.Worksheets(1).Range("allCountInv431") = allCount
    WbNew.Worksheets(1).Range("allSumInv431") = allSum
    
   Dim NumberInv431 As String
     NumberInv431 = WbNew.Worksheets(1).Range("NumberInv431")
     
   Dim FromUnit As String
   Dim toUnit As String
        FromUnit = WbNew.Worksheets(1).Range("FromUnitInv431")
        toUnit = WbNew.Worksheets(1).Range("toUnitInv431")
        
    '=================3.9.11
   Dim fromMRP  As String
   Dim toMRP As String
    
    fromMRP = WbNew.Worksheets(1).Range("MRPInv431")
    toMRP = WbNew.Worksheets(1).Range("MRPToInv431")
    '=================
        
        
        
        
    Dim shp As Shape
    For Each shp In WbNew.Worksheets(1).Shapes
        shp.Delete
    Next shp
        
      Dim fileName As String, fileName2 As String, lineBrake As String
        'Шлях до файла зберігання
        SavePath = WbNew.Worksheets(1).Range("PathToSaveInv431")
        Dim DateInv431 As Date
        DateInv431 = WbNew.Worksheets(1).Range("DateInv431")
        'складаємо назву файла
        
        
               On Error Resume Next
         If Dir(SavePath, vbDirectory) = vbNullString Then
         On Error GoTo 0
        With Application.FileDialog(msoFileDialogFolderPicker)
            .Title = "Виберіть папку для збереження"
            .AllowMultiSelect = False
            If .show = -1 Then
                SavePath = .SelectedItems(1)
                'Записуємо новий шлях у клітинку
                ThisWorkbook.Sheets("Накладна431").Range("PathToSaveInv431").value = SavePath
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
        fileName = SavePath & Application.PathSeparator & CleanFileName(DateInv431 & _
        " Накладна(вим.)№ " & NumberInv431 & " " & FromUnit & " (" & fromMRP & ") - " & toUnit & " (" & toMRP & ").xlsx")
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

## 5. Invoice440_25.bas — Накладна ф. 25 (наказ МОУ 440)
### 5.1. Реквізити та таблиці
- Аркуш: «Накладна440_25»
- Таблиця: `tbInv440_25`
- Основна робоча форма частини для оформлення внутрішнього переміщення між підрозділами та складами.
- Іменовані діапазони: `invNum440_25`, `dateInv440_25`, `FromUnitInv440_25`, `fromMRPInv440_25`, `toUnitInv440_25`, `toMRPInv440_25`, `basis440_25`, `ServiceInv440_25`

### 5.2. Вихідний код Invoice440_25.bas
```vba
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

```

---

## 6. payroll440_57.bas — Відомість ф. 57 (наказ МОУ 440)
### 6.1. Особливості архітектури
На відміну від звичайних вертикальних накладних, форма 57 має матричну структуру: по вертикалі розташовані отримувачі (військовослужбовці), а по горизонталі — до 19 колонок товарів. Процедура `addPosInv440_57` динамічно знаходить наступну вільну колонку в заголовку таблиці та вписує реквізити майна.

### 6.2. Вихідний код payroll440_57.bas
```vba
Attribute VB_Name = "payroll440_57"
Sub addPosInv440_57()
  
    Dim NameOfThings As String
    Dim wsInv440_57 As Worksheet
    Set wsInv440_57 = ThisWorkbook.Worksheets("Розд(З)Відом д57")
    
    Dim tbInv440_57 As ListObject
    Set tbInv440_57 = wsInv440_57.ListObjects("tbInv440_57")
    
    
'=============3.9.11
'визначаємо підрозділ-відправник, беремо з документа
Dim FromUnit As String
FromUnit = wsInv440_57.Range("FromUnitInv440_57")

Dim tempPoolOThings As Range
Set tempPoolOThings = GetResidusePool(wsResiduse2, FromUnit)

If tempPoolOThings Is Nothing Then Exit Sub
'=============3.9.11


 Dim Basis As String
 Basis = InputBox("Підстава (мета) операції: ", NameOfThings, ThisWorkbook.Worksheets("Розд(З)Відом д57").Range("basis440_57").value)
 ThisWorkbook.Worksheets("Розд(З)Відом д57").Range("basis440_57") = Basis
    
 Dim assetRow As Integer
 assetRow = 12 'номер рядка з найменуванням майна
 
 ' Знаходимо першу вільну колонку в діапазоні з 3 до 21
    For assetCol = 3 To 21
        If wsInv440_57.Cells(assetRow, assetCol).value = "" Then
            Exit For ' Знайдено першу порожню клітинку
        End If
    Next assetCol
    
    ' Перевірка, чи є хоча б одна вільна клітинка
    If assetCol > 21 Then
        MsgBox "Усі поля для майна вже заповнені."
        Exit Sub
    End If
    
    
    Dim exists As Boolean
    Dim i As Integer
    Do
        Dim frm     As UserForm1
        Set frm = UserForms.Add(UserForm1.name)
        frm.ListData = tempPoolOThings
        
        frm.show
        NameOfThings = frm.TextBox1.value
        ' Якщо користувач натисне "Скасувати" або не введе нічого
        If NameOfThings = "" Then Exit Sub
         
         
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
        
        exists = False
        For i = 3 To 21
            If wsInv440_57.Cells(assetRow, i).value = name Then
                exists = True
                Exit For
            End If
        Next i
        
        If exists Then
         Dim answer As VbMsgBoxResult
        
            answer = MsgBox(name & " вже є у відомості. Все одно додати?", vbYesNo + vbExclamation, "Майно вже існує")
            If answer = vbYes Then
            GoTo answerYes
            End If
            
        Else
answerYes:

'Визначаємо значення за замовчуванням
            
            Nomenclature = rngWithName.Offset(, -6).value
            unit = rngWithName.Offset(, -7).value
'            category = rngWithName.Offset(, -10).value
            Note = rngWithName.Offset(, -9).value
            price = rngWithName.Offset(, -11).value
            QuantityRest = rngWithName.Offset(, -10).value
'            Place = rngWithName.Offset(, -3).value
            CategoryFrom = rngWithName.Offset(, -8).value
            
'            quantity = InputBox("Кількість: ", NameOfThings, QuantityRest)

'            NoteDt = InputBox("Субрахунок для списання, залишити 1 (якщо не знаєте який - залишіть порожнім): ", NameOfThings, "6611 8013 8113")
            
           
            
            Category = InputBox("Категорія відправника: ", NameOfThings, rngWithName.Offset(, -8).value)
            CategoryTo = InputBox("Категорія отримувача: ", NameOfThings, rngWithName.Offset(, -8).value)
'            VAT = InputBox("ПДВ (за наявності): ", NameOfThings, 0)
'            Note1 = InputBox("Додаткове поле 1", NameOfThings, rngWithName.Offset(, -5).value)
            Note2 = InputBox("Додаткове поле 2", NameOfThings, rngWithName.Offset(, -4).value)
            Note3 = InputBox("Додаткове поле 3", NameOfThings, rngWithName.Offset(, -3).value)

' Додаємо майно в першу вільну клітинку
            With wsInv440_57
                .Cells(assetRow, assetCol).value = name
                .Cells(14, assetCol).value = Nomenclature
                .Cells(16, assetCol).value = unit
                .Cells(18, assetCol).value = Category
                .Cells(20, assetCol).value = price
                .Cells(22, assetCol).value = CategoryTo
                .Cells(24, assetCol).value = Note
'                .Cells(26, assetCol).value = Note1
                .Cells(28, assetCol).value = Note2
                .Cells(30, assetCol).value = Note3
            End With
            ' Переходимо до наступного поля для майна
            assetCol = assetCol + 1
            
            ' Знаходимо наступну вільну клітинку
            For assetCol = 3 To 21
                If wsInv440_57.Cells(assetRow, assetCol).value = "" Then
                    Exit For ' Знайдено першу порожню клітинку
                End If
            Next assetCol
            
            ' Перевірка, чи закінчилися поля для майна
            If assetCol > 21 Then
                MsgBox "Усі поля для майна заповнені."
                Exit Sub
            End If
        End If
'
            
            Dim AddNext
            
            AddNext = MsgBox("Майно додано! Додати наступну позицію?", vbYesNo, NameOfThings)
            
        End If
        
    Loop While AddNext = 6
    

End Sub




Sub AddPerson()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim personName As String
    Dim answer As VbMsgBoxResult
    Dim lastRow As Range
    Dim isFound As Boolean
    Dim cell As Range
    
    ' Встановлюємо лист "РЗВ д48"
    Set ws = ThisWorkbook.Worksheets("Розд(З)Відом д57")
    
    ' Встановлюємо таблицю "Таблиця1"
    Set tbl = ws.ListObjects("tbInv440_57")
    
    ' Додаємо людей
    Do
        ' Отримуємо введення від користувача
'        Dim frm As UserForm1
'        Set frm = UserForms.Add(UserForm1.Name)
'        frm.ListData = ThisWorkbook.Worksheets("Довідники").Range("peopleList")
'        frm.Caption = "Виберіть людину"
'        frm.show
        
        personName = InputBox("Введіть одержувача:", "Одержувач")
        
        ' Якщо користувач натисне "Скасувати" або не введе нічого
        If personName = "" Then Exit Sub
        
'         Dim ListGoods As Range
'            Dim rngPersonName As Range
'
'            Set ListGoods = ThisWorkbook.Worksheets("Довідники").Range("peopleList")
'            ' Шукаємо задане значення в діапазоні
'            Set rngPersonName = ListGoods.Find(personName, LookIn:=xlValues, LookAt:=xlWhole)
'
'            If Not rngPersonName Is Nothing Then
'                Dim Unit As String
'                Dim rankPersName As String
'                ' Визначаємо значення за замовчуванням
'                rankPersName = rngPersonName.Offset(, 1).value
'                Unit = rngPersonName.Offset(, 2).value
'
'
'            End If
            
        
        ' Перевірка, чи таке ім'я вже існує у таблиці
        isFound = False
        For Each cell In tbl.ListColumns(2).DataBodyRange
            If cell.value = personName Then
                isFound = True
                Exit For
            End If
        Next cell
        
        If isFound Then
            MsgBox "Одержувач вже існує у відомості. Введіть іншого.", vbExclamation, "Одержувач вже у відомості"
        Else
            ' Перевіряємо, чи є вільні рядки
            Set lastRow = Nothing
            For Each cell In tbl.ListColumns(2).DataBodyRange
                If cell.value = "" Then
                    Set lastRow = cell
                    Exit For
                End If
            Next cell
            
            ' Якщо не знайдено вільних клітинок, запитуємо, чи потрібно додати новий рядок
            If lastRow Is Nothing Then
                 ' Додаємо новий рядок
                Set lastRow = tbl.ListRows.Add.Range.Cells(1, 2)
            End If
            
            ' Додаємо ім'я людини у першу вільну клітинку
'            lastRow.value = personName
            
'           lastRow.Offset(, 1).value = Unit
           lastRow = personName
            
            ' Перевіряємо, чи знайдено задане значення
           
        End If
        
        ' Питаємо, чи додавати наступне ім'я
        answer = MsgBox("Додати наступного одержувача?", vbYesNo + vbQuestion, "Продовжити?")
        
    Loop While answer = vbYes

    MsgBox "Додавання завершено.", vbInformation, "Готово"
End Sub

Sub DistributeAssets()
    Dim ws As Worksheet
    Dim assetTbl As ListObject
'    Dim personTbl As ListObject
    Dim assetRow As Integer
    Dim personRow As Integer
    Dim assetName As String
    Dim personName As String
    Dim quantity As Variant
    
    ' Встановлюємо лист "РЗВ д48"
    Set ws = ThisWorkbook.Worksheets("Розд(З)Відом д57")
    
    ' Встановлюємо таблиці
    Set assetTbl = ws.ListObjects("tbInv440_57") ' Таблиця з іменами людей
'    Set personTbl = ws.ListObjects("tbInv440_57") ' Таблиця з майном (можна змінити, якщо у вас інша таблиця)
    
    ' Проходимо по всім людям у таблиці
    For personRow = 1 To assetTbl.ListRows.count
        personName = assetTbl.ListColumns(2).DataBodyRange.Cells(personRow, 1).value
        
        ' Проходимо по всім активам
        For assetRow = 3 To 21 ' Змінити, якщо ваша таблиця має інші стовпці
            assetName = ws.Cells(12, assetRow).value
            
            If assetName <> "" Then
                ' Запитуємо, скільки майна видати
                quantity = InputBox(assetName & vbNewLine & " видати:", personName, "0")
                
                ' Перевіряємо, чи введено числове значення
                If IsNumeric(quantity) Then
                    ' Записуємо кількість на перехресті
                    ws.Cells(personRow + 31, assetRow).value = CDbl(quantity) ' Додаємо 31, щоб потрапити в правильний рядок (рядки над таблицею)
                Else
                    MsgBox "Введіть коректне числове значення.", vbExclamation, "Помилка введення"
                End If
            End If
        Next assetRow
    Next personRow
    
    MsgBox "Видача майна завершена.", vbInformation, "Готово"
End Sub

Sub clearTable440_57()
Dim tbActual As ListObject
Set tbActual = ThisWorkbook.Worksheets("Розд(З)Відом д57").ListObjects("tbInv440_57")
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
        If j <> 1 And j <> 22 Then
            firstRow.Range.Cells(1, j).ClearContents
        End If
 Next j
 
Worksheets("Розд(З)Відом д57").Range("rngNames").ClearContents
Worksheets("Розд(З)Відом д57").Range("rngNomenclatures").ClearContents
Worksheets("Розд(З)Відом д57").Range("rngQuantityes").ClearContents
Worksheets("Розд(З)Відом д57").Range("rngCategoriesFrom").ClearContents
Worksheets("Розд(З)Відом д57").Range("rngPrices").ClearContents
Worksheets("Розд(З)Відом д57").Range("rngCategoriesTo").ClearContents
Worksheets("Розд(З)Відом д57").Range("rngNotes").ClearContents
Worksheets("Розд(З)Відом д57").Range("rngNotes1").ClearContents
Worksheets("Розд(З)Відом д57").Range("rngNotes2").ClearContents
Worksheets("Розд(З)Відом д57").Range("rngNotes3").ClearContents
   
 
End Sub


Sub addRow440_57()
Dim tbActual As ListObject
Set tbActual = ThisWorkbook.Worksheets("Розд(З)Відом д57").ListObjects("tbInv440_57")

Dim newRow As ListRow
Set newRow = tbActual.ListRows.Add

End Sub

Sub SaveInv440_57()

 Dim typeOperation As String
   
   typeOperation = InputBox("Тип операції" & vbNewLine & "надходження - 1" _
                                           & vbNewLine & "внутрішнє переміщення - 2" _
                                           & vbNewLine & "вибуття - 3", "Тип операції", 2)
   
   Select Case typeOperation
   Case 1
   ThisWorkbook.Sheets("Розд(З)Відом д57").Range("ContentInv440_57") = "надходження"
   Case 2
    ThisWorkbook.Sheets("Розд(З)Відом д57").Range("ContentInv440_57") = "внутрішнє переміщення"
   Case 3
    ThisWorkbook.Sheets("Розд(З)Відом д57").Range("ContentInv440_57") = "вибуття"
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
   ThisWorkbook.ActiveSheet.Range("ContentInvAdd440_57") = "надійшло з інших в/ч (підрозділів)"
   Case 2
    ThisWorkbook.ActiveSheet.Range("ContentInvAdd440_57") = "надійшло зі складу в/ч"
   Case 3
    ThisWorkbook.ActiveSheet.Range("ContentInvAdd440_57") = "інші надходження"
    Case 4
    ThisWorkbook.ActiveSheet.Range("ContentInvAdd440_57") = "витрата"
    Case 5
    ThisWorkbook.ActiveSheet.Range("ContentInvAdd440_57") = "передано до інших в/ч (підрозділів)"
    Case 6
    ThisWorkbook.ActiveSheet.Range("ContentInvAdd440_57") = "інші вибуття"
   End Select
   
   
   
'збільшуємо номер документа на 1
Dim clearNumber As Long

clearNumber = ThisWorkbook.Sheets("Розд(З)Відом д57").Range("clearNumberInv440_57").value

ThisWorkbook.Sheets("Розд(З)Відом д57").Range("clearNumberInv440_57") = clearNumber + 1

'забираємо суму і кількість
'Dim allCount As String
'Dim allSum As String
'
'allCount = ThisWorkbook.Sheets("Розд(З)Відом д57").Range("allCountInv440_57").value
'allSum = ThisWorkbook.Sheets("Розд(З)Відом д57").Range("allSumInv440_57").value
    
    Dim tbInv440_57 As ListObject
    Set tbInv440_57 = ThisWorkbook.Worksheets("Розд(З)Відом д57").ListObjects("tbInv440_57")

Dim srcRangeCount As Range, dstRangeCount As Range

Set srcRangeCount = tbInv440_57.ListColumns(22).DataBodyRange

        ThisWorkbook.Sheets("Розд(З)Відом д57").Select 'копіюємо перший лист з нарядом
        ThisWorkbook.Sheets("Розд(З)Відом д57").Copy 'зберігаємо як нову книгу
   
     Dim WbNew As Workbook
     
        Set WbNew = ActiveWorkbook
   
     'замінюємо на значення клітинки, які підтягуються по формулі з іншого листа, щоб у збережених документах не злітали дані
     
      For Each cell In WbNew.Worksheets(1).Range(Cells(1, 1), Cells(WbNew.Worksheets(1).Range("bookkeepingInv440_57").Row, 16))
        If cell.HasFormula Then
            cell.value = cell.value
        End If
     Next cell

Dim tbNewInv As ListObject
Set tbNewInv = WbNew.Worksheets(1).ListObjects(1)
Set dstRangeCount = tbNewInv.ListColumns(22).DataBodyRange

 Dim i As Integer
    For i = 1 To srcRangeCount.Rows.count
        dstRangeCount.Cells(i, 1).value = srcRangeCount.Cells(i, 1).value
    Next i



'    WbNew.Worksheets(1).Range("allCountInv440_57") = allCount
'    WbNew.Worksheets(1).Range("allSumInv440_57") = allSum
    
   Dim NumberInv440_57 As String
     NumberInv440_57 = WbNew.Worksheets(1).Range("NumberInv440_57")
     
   Dim FromUnit As String
   Dim toUnit As String
        FromUnit = WbNew.Worksheets(1).Range("FromUnitInv440_57")
        toUnit = WbNew.Worksheets(1).Range("toUnitInv440_57")
    Dim shp As Shape
    For Each shp In WbNew.Worksheets(1).Shapes
        shp.Delete
    Next shp
        
      Dim fileName As String, fileName2 As String, lineBrake As String
        'Шлях до файла зберігання
        SavePath = WbNew.Worksheets(1).Range("PathToSaveInv440_57")
        Dim DateInv440_57 As Date
        DateInv440_57 = WbNew.Worksheets(1).Range("DateInv440_57")
        Dim typeDocInv440_57 As String
        typeDocInv440_57 = WbNew.Worksheets(1).Range("typeDocInv440_57")
        
        
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
                ThisWorkbook.Sheets("Розд(З)Відом д57").Range("PathToSaveInv440_57").value = SavePath
            Else
                MsgBox "Збереження скасовано", vbExclamation
                WbNew.Close SaveChanges:=False
                Application.DisplayAlerts = True
                Exit Sub
            End If
        End With
    End If
        
        
        
        
        
        'складаємо назву файла
        fileName = SavePath & Application.PathSeparator & CleanFileName(DateInv440_57 & " " & _
        typeDocInv440_57 & "№ " & NumberInv440_57 & " " & FromUnit & " - " & toUnit & ".xlsx")
        'заміняємо перенесення рядка, косі лінії

        
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


'============ 3.9.26 ПКМ -> Розд(З)Відом д57 ============
Public Sub addFromTableInv440_57()

    Const FIRST_ASSET_COL As Long = 3   'C
    Const LAST_ASSET_COL As Long = 21   'U

    Const ROW_NAME As Long = 12
    Const ROW_NOMENCLATURE As Long = 14
    Const ROW_UNIT As Long = 16
    Const ROW_CATEGORY_FROM As Long = 18
    Const ROW_PRICE As Long = 20
    Const ROW_CATEGORY_TO As Long = 22
    Const ROW_NOTE As Long = 24
    Const ROW_NOTE2 As Long = 28
    Const ROW_NOTE3 As Long = 30

    Dim ws As Worksheet
    Dim wsDoc As Worksheet
    Dim lo As ListObject

    Dim rowsArr As Variant
    Dim rowIndex As Variant
    Dim rw As Range

    Dim colName As Long
    Dim colNomenclature As Long
    Dim colUnit As Long
    Dim colCategory As Long
    Dim colPrice As Long
    Dim colNote As Long
    Dim colNote2 As Long
    Dim colNote3 As Long
    Dim minDataRow As Long

    Dim NameOfThings As Variant
    Dim Nomenclature As Variant
    Dim UnitValue As Variant
    Dim CategoryValue As Variant
    Dim PriceValue As Variant
    Dim NoteValue As Variant
    Dim Note2Value As Variant
    Dim Note3Value As Variant

    Dim freeCols() As Long
    Dim freeCount As Long
    Dim selectedCount As Long
    Dim targetCol As Long
    Dim i As Long

    Dim listOfThings As String

    On Error GoTo ErrHandler

    Set ws = Selection.Worksheet
    Set wsDoc = ThisWorkbook.Worksheets("Розд(З)Відом д57")

    rowsArr = GetSelectedRowsArrayCell()

    'Визначаємо структуру аркуша-джерела.
    'Мапінг відповідає вже існуючим addFromTable... у версії 3.9.25.
    Select Case ws.name

        Case "База"
            colName = 2
            colNomenclature = 18
            colUnit = 4
            colCategory = 27
            colPrice = 20
            colNote = 16
            colNote2 = 23
            colNote3 = 24

        Case "Залишки по підрозділах2"
            minDataRow = 3
            colName = 2
            colNomenclature = 8
            colUnit = 7
            colCategory = 6
            colPrice = 3
            colNote = 5
            colNote2 = 10
            colNote3 = 11

        Case "Довідка по підрозділу"
            minDataRow = 8
            colName = 3
            colNomenclature = 9
            colUnit = 8
            colCategory = 7
            colPrice = 4
            colNote = 6
            colNote2 = 0
            colNote3 = 0

        Case "Довідка по людині"
            minDataRow = 8
            colName = 2
            colNomenclature = 8
            colUnit = 7
            colCategory = 6
            colPrice = 3
            colNote = 5
            colNote2 = 0
            colNote3 = 0

        Case Else
            MsgBox "Невідомий лист: " & ws.name, vbCritical, "Розд(З)Відом д57"
            GoTo SafeExit

    End Select

    'Перевіряємо, що у виділенні тільки допустимі рядки.
    If ws.name = "База" Then
        Set lo = ws.ListObjects("AllData")
    End If

    For Each rowIndex In rowsArr

        If ws.name = "База" Then
            If Intersect(ws.Rows(CLng(rowIndex)), lo.DataBodyRange) Is Nothing Then
                MsgBox "У виділенні є рядок поза таблицею AllData." & vbCrLf & _
                       "Нічого не додано.", vbExclamation, "Розд(З)Відом д57"
                GoTo SafeExit
            End If
        Else
            If CLng(rowIndex) < minDataRow Then
                MsgBox "У виділенні є службовий рядок." & vbCrLf & _
                       "Нічого не додано.", vbExclamation, "Розд(З)Відом д57"
                GoTo SafeExit
            End If
        End If

        If Len(Trim$(CStr(ws.Cells(CLng(rowIndex), colName).value))) = 0 Then
            MsgBox "У виділенні є порожній рядок." & vbCrLf & _
                   "Нічого не додано.", vbExclamation, "Розд(З)Відом д57"
            GoTo SafeExit
        End If

        selectedCount = selectedCount + 1

    Next rowIndex

    'Збираємо всі вільні колонки майна C:U.
    ReDim freeCols(1 To LAST_ASSET_COL - FIRST_ASSET_COL + 1)

    For i = FIRST_ASSET_COL To LAST_ASSET_COL
        If Len(Trim$(CStr(wsDoc.Cells(ROW_NAME, i).value))) = 0 Then
            freeCount = freeCount + 1
            freeCols(freeCount) = i
        End If
    Next i

    If freeCount = 0 Then
        MsgBox "Усі 19 полів для майна у Розд(З)Відом д57 вже заповнені.", _
               vbExclamation, "Розд(З)Відом д57"
        GoTo SafeExit
    End If

    'Не допускаємо часткового перенесення.
    If selectedCount > freeCount Then
        MsgBox "Вибрано позицій: " & selectedCount & vbCrLf & _
               "Вільних колонок у відомості: " & freeCount & vbCrLf & vbCrLf & _
               "Нічого не додано. Звільніть потрібну кількість колонок або виберіть менше позицій.", _
               vbExclamation, "Розд(З)Відом д57"
        GoTo SafeExit
    End If

    i = 0

    For Each rowIndex In rowsArr

        Set rw = ws.Rows(CLng(rowIndex))

        NameOfThings = rw.Cells(1, colName).value
        Nomenclature = rw.Cells(1, colNomenclature).value
        UnitValue = rw.Cells(1, colUnit).value
        CategoryValue = rw.Cells(1, colCategory).value
        PriceValue = rw.Cells(1, colPrice).value
        NoteValue = rw.Cells(1, colNote).value

        Note2Value = vbNullString
        Note3Value = vbNullString

        If colNote2 > 0 Then Note2Value = rw.Cells(1, colNote2).value
        If colNote3 > 0 Then Note3Value = rw.Cells(1, colNote3).value

        i = i + 1
        targetCol = freeCols(i)

        With wsDoc
            .Cells(ROW_NAME, targetCol).value = NameOfThings
            .Cells(ROW_NOMENCLATURE, targetCol).value = Nomenclature
            .Cells(ROW_UNIT, targetCol).value = UnitValue
            .Cells(ROW_CATEGORY_FROM, targetCol).value = CategoryValue
            .Cells(ROW_PRICE, targetCol).value = PriceValue

            'За замовчуванням категорія отримувача така сама.
            'За потреби її можна змінити безпосередньо у відомості.
            .Cells(ROW_CATEGORY_TO, targetCol).value = CategoryValue

            .Cells(ROW_NOTE, targetCol).value = NoteValue
            .Cells(ROW_NOTE2, targetCol).value = Note2Value
            .Cells(ROW_NOTE3, targetCol).value = Note3Value
        End With

        listOfThings = listOfThings & "- " & CStr(NameOfThings) & vbCrLf

    Next rowIndex

    MsgBox "До Розд(З)Відом д57 додано позицій: " & selectedCount & vbCrLf & vbCrLf & _
           listOfThings & vbCrLf & _
           "Кількість не переносилась.", _
           vbInformation, "Розд(З)Відом д57"

SafeExit:
    On Error Resume Next
    Application.CommandBars("Cell").Controls("Додати в документ").Delete
    On Error GoTo 0
    Exit Sub

ErrHandler:
    MsgBox "Помилка під час додавання майна у Розд(З)Відом д57:" & vbCrLf & _
           Err.Number & " - " & Err.Description, _
           vbCritical, "Розд(З)Відом д57"
    Resume SafeExit

End Sub




```
