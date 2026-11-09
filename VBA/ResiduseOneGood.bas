Attribute VB_Name = "ResiduseOneGood"
Sub ResiduseForGood()
Dim wbAllData As Workbook
Dim wsAllData As Worksheet
Dim wsResiduse As Worksheet
Dim tbAllData As ListObject

Set wbAllData = ThisWorkbook
Set wsAllData = wbAllData.Worksheets("База")
Set tbAllData = wsAllData.ListObjects("AllData")
Set wsResiduse = wbAllData.Worksheets("Залишки по майну")

Dim tbGoodCriteria As ListObject
Set tbGoodCriteria = wsResiduse.ListObjects("tbGoodCriteria")

'очищаємо дані з попереднього запиту
tbGoodCriteria.DataBodyRange.ClearContents

Dim frm As UserForm1
Set frm = New UserForm1

'Dim colFrom As Range, colTo As Range, cell As Range
'Dim dictUnits As Object
'Set dictUnits = CreateObject("Scripting.Dictionary")
'
'Set colFrom = tbAllData.ListColumns("Звідки").DataBodyRange
'Set colTo = tbAllData.ListColumns("Куди").DataBodyRange
'
'' Формуємо словник унікальних підрозділів
'For Each cell In colFrom
'    If Not dictUnits.exists(cell.value) And Not IsEmpty(cell.value) Then
'        dictUnits.Add cell.value, Nothing
'    End If
'Next cell
'
'For Each cell In colTo
'    If Not dictUnits.exists(cell.value) And Not IsEmpty(cell.value) Then
'        dictUnits.Add cell.value, Nothing
'    End If
'Next cell

' Виводимо список підрозділів у колонку A, починаючи з A7
'Dim n As Long
'n = 7
'wsResiduse.Range("A7:A" & wsResiduse.Rows.count).ClearContents ' Очистити перед вставкою
'
'Dim unitKey As Variant
'For Each unitKey In dictUnits.Keys
'    wsResiduse.Cells(n, 1).value = unitKey
'    n = n + 1
'Next unitKey
'
'' Зчитати назад у масив
'Dim lastRow As Long
'lastRow = wsResiduse.Cells(wsResiduse.Rows.count, 1).End(xlUp).Row

'Dim unitArray As Variant
'unitArray = wsResiduse.Range("A7:A" & lastRow).value

' Передати в ListBox
'If IsArray(unitArray) Then
    frm.Caption = "Виберіть майно"
    frm.ListData = wsResiduse2.Range("Goods")
'Else
'    frm.ListBox1.AddItem unitArray ' якщо один елемент
'End If

' Показати форму
frm.show

Dim NameOfThings As String
Dim ListGoods As Range
        Set ListGoods = wsResiduse2.Range("Goods")

' Отримати вибране
If frm.ListBox1.ListIndex <> -1 Then
    NameOfThings = frm.TextBox1.value
Else
    MsgBox "Майно не вибрано!", vbExclamation
    Exit Sub
End If

Set rngWithName = ListGoods.Find(NameOfThings, LookIn:=xlValues, LookAt:=xlWhole)
'================================закінчив тут, треба відібрати в верхню таблицю-фільтр критерії, якщо вони в полях проставлено "так"
'відфільтрувати базу даних по цих катергоріях
'відібрати унікальні підрозділи
'якщо в них нема залишків не виводити
'якщо є залишки - вивести в таблицю
'може зробити mvp тільки по найменуванню?
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
       'рядок, де прописуємо "так", якщо хочемо щоб цей параметр включився в пошук
       Dim rowYes As Long
       rowYes = 8
       'найменування
       If wsResiduse.Cells(rowYes, 1) = "так" Then
            name = rngWithName.Offset(, -12).value
       End If
       'ціна
       If wsResiduse.Cells(rowYes, 2) = "так" Then
            price = rngWithName.Offset(, -11).value
       End If
       'одиниця виміру
       If wsResiduse.Cells(rowYes, 3) = "так" Then
            unit = rngWithName.Offset(, -7).value
       End If
       'категорія
       If wsResiduse.Cells(rowYes, 4) = "так" Then
            Category = rngWithName.Offset(, -8).value
       End If
       'примітка
       If wsResiduse.Cells(rowYes, 5) = "так" Then
            Note = rngWithName.Offset(, -9).value
       End If
       'код номенклатури
       If wsResiduse.Cells(rowYes, 6) = "так" Then
            Nomenclature = rngWithName.Offset(, -6).value
       End If
       'додаткове поле 2
       If wsResiduse.Cells(rowYes, 7) = "так" Then
           Note2 = rngWithName.Offset(, -4).value
       End If
       'додаткове поле 3
       If wsResiduse.Cells(rowYes, 8) = "так" Then
            Note3 = rngWithName.Offset(, -3).value
       End If
            
        
            
                With tbGoodCriteria.DataBodyRange.Rows
                    .Cells(1, 1) = name
                    .Cells(1, 2) = price
                    .Cells(1, 3) = unit
                    .Cells(1, 4) = Category
                    .Cells(1, 5) = Note
                    .Cells(1, 6) = Nomenclature
                    .Cells(1, 7) = Note2
                    .Cells(1, 8) = Note3
                    
'                     .Cells(1, 7) = CDbl(quantity)
'
'                    .Cells(1, 11) = CategoryTo
                    
'                    .Cells(1, 12) = Note1
                    
                End With
           End If



'
' wsResiduse.Range("selectedUnit").value = selectedUnit ' Вибраний підрозділ
' wsResiduse.Range("A7:A" & lastRow).ClearContents

Dim dict As Object
Set dict = CreateObject("Scripting.Dictionary")

Dim i As Long
Dim key As Variant
Dim result As Double

' Очистити старі дані
wsResiduse.Range("b8:i" & wsResiduse.Rows.count).ClearContents

Dim dataCount As Long
dataCount = tbAllData.ListRows.count

' Формуємо унікальні ключі на основі декількох полів
For i = 1 To dataCount
    If tbAllData.DataBodyRange(i, tbAllData.ListColumns("Куди").Index).value = selectedUnit _
    Or tbAllData.DataBodyRange(i, tbAllData.ListColumns("Звідки").Index).value = selectedUnit Then
    
      Dim dynamicCategory As String

' Вибираємо значення для категорії в залежності від напрямку
If tbAllData.DataBodyRange(i, tbAllData.ListColumns("Куди").Index).value = selectedUnit Then
    dynamicCategory = tbAllData.DataBodyRange(i, tbAllData.ListColumns("Запасне поле3").Index).value
Else
    dynamicCategory = tbAllData.DataBodyRange(i, tbAllData.ListColumns("Категорія").Index).value
End If

key = tbAllData.DataBodyRange(i, tbAllData.ListColumns("Код номенклатури").Index).value & "||" & _
      tbAllData.DataBodyRange(i, tbAllData.ListColumns("Найменування").Index).value & "||" & _
      tbAllData.DataBodyRange(i, tbAllData.ListColumns("Ціна").Index).value & "||" & _
      tbAllData.DataBodyRange(i, tbAllData.ListColumns("Одиниця виміру").Index).value & "||" & _
      dynamicCategory & "||" & _
      tbAllData.DataBodyRange(i, tbAllData.ListColumns("Примітка").Index).value


        If Not dict.exists(key) Then
            dict.Add key, 0
        End If
        
        ' Додаємо/віднімаємо залишок
        If tbAllData.DataBodyRange(i, tbAllData.ListColumns("Куди").Index).value = selectedUnit Then
            dict(key) = dict(key) + tbAllData.DataBodyRange(i, tbAllData.ListColumns("Надійшло").Index).value
            dict(key) = dict(key) + tbAllData.DataBodyRange(i, tbAllData.ListColumns("Вибуло").Index).value
        End If
        
        If tbAllData.DataBodyRange(i, tbAllData.ListColumns("Звідки").Index).value = selectedUnit Then
            dict(key) = dict(key) - tbAllData.DataBodyRange(i, tbAllData.ListColumns("Вибуло").Index).value
        End If
    End If
Next i

' Виводимо результати
Dim rowIndex As Long
rowIndex = 8
Dim parts() As String

Dim k As Long
Dim totalKeys As Long
Dim allKeys As Variant

allKeys = dict.Keys
totalKeys = dict.count
d = 0
For k = 1 To totalKeys
    If dict(allKeys(k - 1)) <> 0 Then
        parts = Split(allKeys(k - 1), "||")
        d = d + 1
        wsResiduse.Cells(rowIndex, 2).value = d
        wsResiduse.Cells(rowIndex, 3).value = parts(1) ' Найменування
        wsResiduse.Cells(rowIndex, 4).value = CDbl(parts(2)) ' Ціна
        wsResiduse.Cells(rowIndex, 5).value = CDbl(dict(allKeys(k - 1))) ' Залишок
        wsResiduse.Cells(rowIndex, 6).value = parts(5) ' Примітка
        wsResiduse.Cells(rowIndex, 7).value = parts(4) ' Категорія
        wsResiduse.Cells(rowIndex, 8).value = parts(3) ' Одиниця виміру
        wsResiduse.Cells(rowIndex, 9).value = parts(0) ' Код номенклатури
        rowIndex = rowIndex + 1
    End If
    
    Call ProgressBar.Progress(k / totalKeys * 100)
Next k
End Sub

