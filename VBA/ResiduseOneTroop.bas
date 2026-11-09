Attribute VB_Name = "ResiduseOneTroop"
Sub UpdateResiduseDetailed()

Dim wbAllData As Workbook
Dim wsAllData As Worksheet
Dim wsResiduse As Worksheet
Dim tbAllData As ListObject

Set wbAllData = ThisWorkbook
Set wsAllData = wbAllData.Worksheets("База")
Set tbAllData = wsAllData.ListObjects("AllData")
Set wsResiduse = wbAllData.Worksheets("Довідка по підрозділу") ' ОНОВЛЕНО


Dim frm As UserForm1
Set frm = New UserForm1

Dim colFrom As Range, colTo As Range, cell As Range
Dim dictUnits As Object
Set dictUnits = CreateObject("Scripting.Dictionary")

Set colFrom = tbAllData.ListColumns("Звідки").DataBodyRange
Set colTo = tbAllData.ListColumns("Куди").DataBodyRange

' Формуємо словник унікальних підрозділів
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

' Виводимо список підрозділів у колонку A, починаючи з A7
Dim n As Long
n = 7
wsResiduse.Range("A7:A" & wsResiduse.Rows.count).ClearContents ' Очистити перед вставкою

Dim unitKey As Variant
For Each unitKey In dictUnits.Keys
    wsResiduse.Cells(n, 1).value = unitKey
    n = n + 1
Next unitKey

' Зчитати назад у масив
Dim lastRow As Long
lastRow = wsResiduse.Cells(wsResiduse.Rows.count, 1).End(xlUp).Row

'Dim unitArray As Variant
'unitArray = wsResiduse.Range("A7:A" & lastRow).value

' Передати в ListBox
'If IsArray(unitArray) Then
    frm.Caption = "Виберіть підрозділ"
    frm.ListData = wsResiduse.Range("A7:A" & lastRow)
'Else
'    frm.ListBox1.AddItem unitArray ' якщо один елемент
'End If

' Показати форму
frm.show

' Отримати вибране
If frm.ListBox1.ListIndex <> -1 Then
    selectedUnit = frm.ListBox1.value
Else
    MsgBox "Підрозділ не вибрано!", vbExclamation
    Exit Sub
End If



 wsResiduse.Range("selectedUnit").value = selectedUnit ' Вибраний підрозділ
 wsResiduse.Range("A7:A" & lastRow).ClearContents

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

If wsResiduse.AutoFilterMode Then wsResiduse.AutoFilterMode = False
    wsResiduse.Range("B7").CurrentRegion.AutoFilter


End Sub




Sub SaveResiduseOneUnit()

        ThisWorkbook.Sheets("Довідка по підрозділу").Select 'копіюємо перший лист з нарядом
        ThisWorkbook.Sheets("Довідка по підрозділу").Copy 'зберігаємо як нову книгу
   
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
         CleanFileName(WbNew.Worksheets(1).Range("selectedUnit").value & _
         fileName & ".xlsx"), FileFormat:=xlOpenXMLWorkbook
'     On Error GoTo 0
'      Exit Sub
'WrongPath:
     MsgBox "Файл з залишками матеріальних цінностей збережений і тій же папці, де файл бази даних", vbInformation, ThisWorkbook.Path
'     Exit Sub
     
End Sub


