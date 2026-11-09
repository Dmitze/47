Attribute VB_Name = "MoveForUnit"
Option Explicit
Sub MoveUnitGoods() 'рух по підрозділу

    Dim wbAllData As Workbook
    Dim wsAllData As Worksheet
    Dim wsMove As Worksheet
    Dim tbAllData As ListObject

    Set wbAllData = ThisWorkbook
    Set wsAllData = wbAllData.Worksheets("База")
    Set tbAllData = wsAllData.ListObjects("AllData")
    Set wsMove = wbAllData.Worksheets("Рух по підрозділу")

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
    wsMove.Range("A7:A" & wsMove.Rows.count).ClearContents ' Очистити перед вставкою

    Dim unitKey As Variant
    For Each unitKey In dictUnits.Keys
        wsMove.Cells(n, 1).value = unitKey
        n = n + 1
    Next unitKey

    ' Зчитати назад у масив
    Dim lastRow As Long
    lastRow = wsMove.Cells(wsMove.Rows.count, 1).End(xlUp).Row

    frm.Caption = "Виберіть підрозділ"
    frm.ListData = wsMove.Range("A7:A" & lastRow)
    frm.show

    Dim selectedUnit As String
    If frm.ListBox1.ListIndex <> -1 Then
        selectedUnit = frm.ListBox1.value
    Else
        MsgBox "Підрозділ не вибрано!", vbExclamation
        Exit Sub
    End If

wsMove.Range("moveUnit") = selectedUnit
wsMove.Range("CreateDate") = Date


    ' Очистити тимчасові значення
    wsMove.Range("A7:A" & lastRow).ClearContents

    ' Очистити попередні фільтри
    If Not tbAllData.AutoFilter Is Nothing Then
    If tbAllData.AutoFilter.FilterMode Then
        tbAllData.AutoFilter.ShowAllData
    End If
End If

    ' Очистити старі дані на листі "Рух по підрозділу"
    wsMove.Rows("5:" & wsMove.Rows.count).ClearContents

    ' Вставити заголовки у рядок 4
    'tbAllData.HeaderRowRange.Copy destination:=wsMove.Range("b4")

' Зчитуємо діапазон дат
    Dim startDate As Date, endDate As Date
    On Error Resume Next
    startDate = CDate(InputBox("Введіть дату початку періоду у форматі дд.мм.рррр", "Дата З (включно)", Date))
    endDate = CDate(InputBox("Введіть дату кінця періоду у форматі дд.мм.рррр", "Дата ПО (включно)", Date))
    On Error GoTo 0


'==========3.9.25==============

Dim periodMove As String
    ' Створити тимчасові критерії у комірках
    With wsMove
        .Range("c1:d3").ClearContents
        .Range("c1").value = "Звідки"
        .Range("d1").value = "Куди"
        .Range("c2").value = selectedUnit
        .Range("d3").value = selectedUnit
        
        .Range("e1").value = "Дата документа"
        .Range("f1").value = "Дата документа"
        
    If startDate = 0 Then
    .Range("e2").value = ""
    .Range("e3").value = ""
     periodMove = "з (не застосовувався)"
    Else
    .Range("e2").value = ">=" & CLng(startDate)
    .Range("e3").value = ">=" & CLng(startDate)
    periodMove = "з " & startDate
    End If
          
    If endDate = 0 Then
    .Range("f2").value = ""
    .Range("f3").value = ""
    periodMove = periodMove & " по(не застосовувався)"
    Else
    .Range("f2").value = "<=" & CLng(endDate)
    .Range("f3").value = "<=" & CLng(endDate)
    periodMove = periodMove & " по " & endDate
    End If
    
    End With
wsMove.Range("periodMove") = periodMove


'==========3.9.25==============


    ' Застосувати розширений фільтр
    tbAllData.Range.AdvancedFilter _
        Action:=xlFilterCopy, _
        criteriaRange:=wsMove.Range("c1:f3"), _
        CopyToRange:=wsMove.Range("a5"), _
        Unique:=False
        
'CopyToRange:=wsMove.Range("b5"), _
'     Очистити тимчасові критерії
    wsMove.Range("c1:f3").ClearContents

    ' Скинути фільтр у джерельній таблиці
    If Not tbAllData.AutoFilter Is Nothing Then
    If tbAllData.AutoFilter.FilterMode Then
        tbAllData.AutoFilter.ShowAllData
    End If
End If

End Sub

Sub SaveMoveUnitGoods()

        ThisWorkbook.Sheets("Рух по підрозділу").Select 'копіюємо перший лист з нарядом
        ThisWorkbook.Sheets("Рух по підрозділу").Copy 'зберігаємо як нову книгу
   
     Dim WbNew As Workbook
     
        Set WbNew = ActiveWorkbook
   
    Dim shp As Shape
    For Each shp In WbNew.Worksheets(1).Shapes
        shp.Delete
    Next shp
    
'    WbNew.Worksheets(1).Columns("N").ClearContents
        
      Dim fileName As String
     fileName = ", рух майна станом на " & WbNew.Worksheets(1).Range("CreateDate")
        'складаємо назву файла
        
'        WbNew.Worksheets(1).Range("A3") = fileName
        
        'зберігаємо файл
'        On Error GoTo WrongPath
         WbNew.SaveAs fileName:=ThisWorkbook.Path & Application.PathSeparator & _
         CleanFileName(WbNew.Worksheets(1).Range("moveUnit").value & _
         fileName & ".xlsx"), FileFormat:=xlOpenXMLWorkbook
'     On Error GoTo 0
'      Exit Sub
'WrongPath:
     MsgBox "Файл з рухом матеріальних цінностей збережений і тій же папці, де файл бази даних", vbInformation, ThisWorkbook.Path
'     Exit Sub
     
End Sub
