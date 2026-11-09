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



