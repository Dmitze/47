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





