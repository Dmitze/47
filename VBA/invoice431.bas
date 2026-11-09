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

