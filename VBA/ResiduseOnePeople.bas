Attribute VB_Name = "ResiduseOnePeople"
'=========3.9.14=========



Sub UpdateResiduseOnePeople()

Dim wbAllData As Workbook
Dim wsAllData As Worksheet
Dim wsResiduse As Worksheet
Dim tbAllData As ListObject

Set wbAllData = ThisWorkbook
Set wsAllData = wbAllData.Worksheets("База")
Set tbAllData = wsAllData.ListObjects("AllData")
Set wsResiduse = wbAllData.Worksheets("Довідка по людині") ' ОНОВЛЕНО


Dim frm As UserForm1
Set frm = New UserForm1

Dim colFrom As Range, colTo As Range, cell As Range
Dim dictUnits As Object
Set dictUnits = CreateObject("Scripting.Dictionary")

Set colFrom = tbAllData.ListColumns("МВО звідки").DataBodyRange
Set colTo = tbAllData.ListColumns("МВО куди").DataBodyRange

' Формуємо словник унікальних людей
For Each cell In colFrom
    If Not dictUnits.exists(cell.value) And Not IsEmpty(cell.value) Then
        dictUnits.Add cell.value, Nothing
    End If
Next cell

For Each cell In colTo
    If Not dictUnits.exists(cell.value) And Not IsEmpty(cell.value) Then
        dictUnits.Add cell.value, Nothing
    End If
Next cell

' Виводимо список людей
Dim wsTroops As Worksheet
Set wsTroops = wbAllData.Worksheets("Підрозділи")

Dim k As Long
k = 1
wsTroops.Range("B1:B" & wsTroops.Rows.count).ClearContents ' Очистити перед вставкою

Dim unitKey As Variant
For Each unitKey In dictUnits.Keys
    wsTroops.Cells(k, 2).value = unitKey
    k = k + 1
Next unitKey


Dim lastRowT As Long
lastRowT = wsTroops.Cells(wsTroops.Rows.count, 2).End(xlUp).Row


' Показуємо форму для вибору
frm.Caption = "Виберіть людину"
frm.ListData = wsTroops.Range("B1:B" & lastRowT)
    frm.show

    Dim selectedUnit As String
    If frm.ListBox1.ListIndex <> -1 Then
        selectedUnit = frm.ListBox1.value
    Else
        MsgBox "Людину не вибрано! Спробуйте ще раз!", vbExclamation
        Exit Sub
    End If

 wsResiduse.Range("selectedPeople").value = selectedUnit ' Вибраний підрозділ


Dim dict As Object
Set dict = CreateObject("Scripting.Dictionary")

Dim i As Long
Dim key As Variant
Dim result As Double

' Очистити старі дані
wsResiduse.Range("A8:I" & wsResiduse.Rows.count).ClearContents

Dim dataCount As Long
dataCount = tbAllData.ListRows.count

' Формуємо унікальні ключі на основі декількох полів
For i = 1 To dataCount
    If tbAllData.DataBodyRange(i, tbAllData.ListColumns("МВО куди").Index).value = selectedUnit _
    Or tbAllData.DataBodyRange(i, tbAllData.ListColumns("МВО звідки").Index).value = selectedUnit Then
    
      Dim dynamicCategory As String

' Вибираємо значення для категорії в залежності від напрямку
If tbAllData.DataBodyRange(i, tbAllData.ListColumns("МВО куди").Index).value = selectedUnit Then
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
'      & "||" & _
'      tbAllData.DataBodyRange(i, tbAllData.ListColumns("Куди").Index).value

        If Not dict.exists(key) Then
            dict.Add key, 0
        End If
        
        ' Додаємо/віднімаємо залишок
        If tbAllData.DataBodyRange(i, tbAllData.ListColumns("МВО куди").Index).value = selectedUnit Then
            dict(key) = dict(key) + tbAllData.DataBodyRange(i, tbAllData.ListColumns("Надійшло").Index).value
            dict(key) = dict(key) + tbAllData.DataBodyRange(i, tbAllData.ListColumns("Вибуло").Index).value
        End If
        
        If tbAllData.DataBodyRange(i, tbAllData.ListColumns("МВО звідки").Index).value = selectedUnit Then
            dict(key) = dict(key) - tbAllData.DataBodyRange(i, tbAllData.ListColumns("Вибуло").Index).value
        End If
    End If
Next i

' Виводимо результати
Dim rowIndex As Long
rowIndex = 8
Dim parts() As String

Dim m As Long
Dim totalKeys As Long
Dim allKeys As Variant

allKeys = dict.Keys
totalKeys = dict.count
d = 0
For m = 1 To totalKeys
    If dict(allKeys(m - 1)) <> 0 Then
        parts = Split(allKeys(m - 1), "||")
        d = d + 1
        wsResiduse.Cells(rowIndex, 1).value = d
        wsResiduse.Cells(rowIndex, 2).value = parts(1) ' Найменування
        wsResiduse.Cells(rowIndex, 3).value = CDbl(parts(2)) ' Ціна
        wsResiduse.Cells(rowIndex, 4).value = CDbl(dict(allKeys(m - 1))) ' Залишок
        wsResiduse.Cells(rowIndex, 5).value = parts(5) ' Примітка
        wsResiduse.Cells(rowIndex, 6).value = parts(4) ' Категорія
        wsResiduse.Cells(rowIndex, 7).value = parts(3) ' Одиниця виміру
        wsResiduse.Cells(rowIndex, 8).value = parts(0) ' Код номенклатури
'        wsResiduse.Cells(rowIndex, 9).value = parts(6) ' підрозділ
        rowIndex = rowIndex + 1
    End If
    
    Call ProgressBar.Progress(m / totalKeys * 100)
Next m
'ставимо фільтри для зручності
    If wsResiduse.AutoFilterMode Then wsResiduse.AutoFilterMode = False
    wsResiduse.Range("A7").CurrentRegion.AutoFilter
    

End Sub




Sub SaveResiduseOnePeople()

        ThisWorkbook.Sheets("Довідка по людині").Select 'копіюємо перший лист з нарядом
        ThisWorkbook.Sheets("Довідка по людині").Copy 'зберігаємо як нову книгу
   
     Dim WbNew As Workbook
     
        Set WbNew = ActiveWorkbook
   
    Dim shp As Shape
    For Each shp In WbNew.Worksheets(1).Shapes
        shp.Delete
    Next shp
    
'    WbNew.Worksheets(1).Columns("N").ClearContents
        
      Dim fileName As String
     fileName = ", залишки МЦ станом на " & Date
        'складаємо назву файла
        
        WbNew.Worksheets(1).Range("A3") = fileName
        
        'зберігаємо файл
'        On Error GoTo WrongPath
         WbNew.SaveAs fileName:=ThisWorkbook.Path & Application.PathSeparator & _
         CleanFileName(WbNew.Worksheets(1).Range("selectedPeople").value & _
         fileName & ".xlsx"), FileFormat:=xlOpenXMLWorkbook
'     On Error GoTo 0
'      Exit Sub
'WrongPath:
     MsgBox "Файл з залишками матеріальних цінностей збережений і тій же папці, де файл бази даних", vbInformation, ThisWorkbook.Path
'     Exit Sub
     
End Sub




