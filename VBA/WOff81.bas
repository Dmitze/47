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

