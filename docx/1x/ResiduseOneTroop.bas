Attribute VB_Name = "ResiduseOneTroop"
Option Explicit

Sub UpdateResiduseDetailed()

    Dim wbAllData As Workbook
    Dim wsAllData As Worksheet
    Dim wsResiduse As Worksheet
    Dim tbAllData As ListObject

    Set wbAllData = ThisWorkbook
    Set wsAllData = wbAllData.Worksheets("База")
    Set tbAllData = wsAllData.ListObjects("AllData")
    Set wsResiduse = wbAllData.Worksheets("Довідка по підрозділу")

    Dim frm As UserForm1
    Set frm = New UserForm1

    Dim colFrom As Range, colTo As Range, cell As Range
    Dim dictUnits As Object
    Set dictUnits = CreateObject("Scripting.Dictionary")

    Set colFrom = tbAllData.ListColumns("Звідки").DataBodyRange
    Set colTo = tbAllData.ListColumns("Куди").DataBodyRange

    ' Формуємо словник унікальних підрозділів
    For Each cell In colFrom
        If Not dictUnits.exists(cell.Value) And Len(Trim(CStr(cell.Value))) > 0 Then
            dictUnits.Add cell.Value, Nothing
        End If
    Next cell

    For Each cell In colTo
        If Not dictUnits.exists(cell.Value) And Len(Trim(CStr(cell.Value))) > 0 Then
            dictUnits.Add cell.Value, Nothing
        End If
    Next cell

    ' Виводимо список підрозділів у колонку A
    Dim n As Long
    n = 7
    wsResiduse.Range("A7:A" & wsResiduse.Rows.Count).ClearContents

    Dim unitKey As Variant
    For Each unitKey In dictUnits.Keys
        wsResiduse.Cells(n, 1).Value = unitKey
        n = n + 1
    Next unitKey

    Dim lastRow As Long
    lastRow = wsResiduse.Cells(wsResiduse.Rows.Count, 1).End(xlUp).Row

    frm.Caption = "Виберіть підрозділ"
    frm.ListData = wsResiduse.Range("A7:A" & lastRow)
    frm.show

    Dim selectedUnit As String
    If frm.ListBox1.ListIndex <> -1 Then
        selectedUnit = frm.ListBox1.Value
    Else
        MsgBox "Підрозділ не вибрано!", vbExclamation
        Exit Sub
    End If

    wsResiduse.Range("selectedUnit").Value = selectedUnit
    wsResiduse.Range("A7:A" & lastRow).ClearContents

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim i As Long
    Dim key As Variant
    Dim dataCount As Long
    dataCount = tbAllData.ListRows.Count

    ' Очистити старі дані
    wsResiduse.Range("b8:i" & wsResiduse.Rows.Count).ClearContents

    Dim cFrom As Long, cTo As Long, cIn As Long, cOut As Long, cOpType As Long
    Dim cName As Long, cCode As Long, cPrice As Long, cUnit As Long
    Dim cCat As Long, cCat3 As Long, cNote As Long

    With tbAllData
        cFrom = .ListColumns("Звідки").Index
        cTo = .ListColumns("Куди").Index
        cIn = .ListColumns("Надійшло").Index
        cOut = .ListColumns("Вибуло").Index
        cOpType = .ListColumns("Запасне поле2").Index
        cName = .ListColumns("Найменування").Index
        cCode = .ListColumns("Код номенклатури").Index
        cPrice = .ListColumns("Ціна").Index
        cUnit = .ListColumns("Одиниця виміру").Index
        cCat = .ListColumns("Категорія").Index
        cCat3 = .ListColumns("Запасне поле3").Index
        cNote = .ListColumns("Примітка").Index
    End With

    Dim uFrom As String, uTo As String, opType As String
    Dim valIn As Double, valOut As Double, movQty As Double
    Dim dynamicCat As String

    For i = 1 To dataCount
        uFrom = Trim(CStr(tbAllData.DataBodyRange(i, cFrom).Value))
        uTo = Trim(CStr(tbAllData.DataBodyRange(i, cTo).Value))

        If uFrom = selectedUnit Or uTo = selectedUnit Then
            opType = LCase(Trim(CStr(tbAllData.DataBodyRange(i, cOpType).Value)))
            valIn = 0#: valOut = 0#
            If IsNumeric(tbAllData.DataBodyRange(i, cIn).Value) Then valIn = CDbl(tbAllData.DataBodyRange(i, cIn).Value)
            If IsNumeric(tbAllData.DataBodyRange(i, cOut).Value) Then valOut = CDbl(tbAllData.DataBodyRange(i, cOut).Value)

            If uTo = selectedUnit Then
                dynamicCat = CStr(tbAllData.DataBodyRange(i, cCat3).Value)
                If Len(Trim(dynamicCat)) = 0 Then dynamicCat = CStr(tbAllData.DataBodyRange(i, cCat).Value)
            Else
                dynamicCat = CStr(tbAllData.DataBodyRange(i, cCat).Value)
            End If

            key = CStr(tbAllData.DataBodyRange(i, cCode).Value) & "||" & _
                  CStr(tbAllData.DataBodyRange(i, cName).Value) & "||" & _
                  CStr(tbAllData.DataBodyRange(i, cPrice).Value) & "||" & _
                  CStr(tbAllData.DataBodyRange(i, cUnit).Value) & "||" & _
                  dynamicCat & "||" & _
                  CStr(tbAllData.DataBodyRange(i, cNote).Value)

            If Not dict.exists(key) Then dict.Add key, 0#

            Select Case opType
                Case "надходження"
                    If uTo = selectedUnit Then
                        dict(key) = dict(key) + valIn
                    End If

                Case "вибуття"
                    If uFrom = selectedUnit Then
                        dict(key) = dict(key) - valOut
                    End If

                Case "внутрішнє переміщення"
                    movQty = valOut
                    If movQty = 0 Then movQty = valIn

                    If uTo = selectedUnit Then
                        dict(key) = dict(key) + movQty
                    End If
                    If uFrom = selectedUnit Then
                        dict(key) = dict(key) - movQty
                    End If
            End Select
        End If
    Next i

    ' Виведення результатів
    Dim rowIndex As Long
    rowIndex = 8
    Dim parts() As String
    Dim k As Long, totalKeys As Long, allKeys As Variant
    Dim d As Long

    allKeys = dict.Keys
    totalKeys = dict.Count
    d = 0

    For k = 1 To totalKeys
        If Round(dict(allKeys(k - 1)), 6) <> 0 Then
            parts = Split(allKeys(k - 1), "||")
            d = d + 1
            wsResiduse.Cells(rowIndex, 2).Value = d
            wsResiduse.Cells(rowIndex, 3).Value = parts(1) ' Найменування
            wsResiduse.Cells(rowIndex, 4).Value = CDbl(parts(2)) ' Ціна
            wsResiduse.Cells(rowIndex, 5).Value = CDbl(dict(allKeys(k - 1))) ' Залишок
            wsResiduse.Cells(rowIndex, 6).Value = parts(5) ' Примітка
            wsResiduse.Cells(rowIndex, 7).Value = parts(4) ' Категорія
            wsResiduse.Cells(rowIndex, 8).Value = parts(3) ' Одиниця виміру
            wsResiduse.Cells(rowIndex, 9).Value = parts(0) ' Код номенклатури
            rowIndex = rowIndex + 1
        End If
        If totalKeys > 0 Then Call ProgressBar.Progress(k / totalKeys * 100)
    Next k

    If wsResiduse.AutoFilterMode Then wsResiduse.AutoFilterMode = False
    wsResiduse.Range("B7").CurrentRegion.AutoFilter

End Sub

Sub SaveResiduseOneUnit()

    ThisWorkbook.Sheets("Довідка по підрозділу").Copy
    Dim WbNew As Workbook
    Set WbNew = ActiveWorkbook

    Dim shp As Shape
    For Each shp In WbNew.Worksheets(1).Shapes
        shp.Delete
    Next shp

    Dim fileName As String
    fileName = ", залишки МЦ станом на " & Date
    WbNew.Worksheets(1).Range("A3") = fileName

    Application.DisplayAlerts = False
    WbNew.SaveAs fileName:=ThisWorkbook.Path & Application.PathSeparator & _
        CleanFileName(WbNew.Worksheets(1).Range("selectedUnit").Value & _
        fileName & ".xlsx"), FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True

    MsgBox "Файл з залишками матеріальних цінностей збережений у тій же папці, де файл бази даних", vbInformation, ThisWorkbook.Path

End Sub
