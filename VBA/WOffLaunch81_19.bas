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







