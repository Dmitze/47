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
        If Not dictUnits.exists(cell.value) And Len(Trim(CStr(cell.value))) > 0 Then
            dictUnits.Add cell.value, Nothing
        End If
    Next cell

    For Each cell In colTo
        If Not dictUnits.exists(cell.value) And Len(Trim(CStr(cell.value))) > 0 Then
            dictUnits.Add cell.value, Nothing
        End If
    Next cell

    ' Виводимо список підрозділів у колонку A
    Dim n As Long
    n = 7
    wsResiduse.Range("A7:A" & wsResiduse.Rows.count).ClearContents

    Dim unitKey As Variant
    For Each unitKey In dictUnits.Keys
        wsResiduse.Cells(n, 1).value = unitKey
        n = n + 1
    Next unitKey

    Dim lastRow As Long
    lastRow = wsResiduse.Cells(wsResiduse.Rows.count, 1).End(xlUp).Row

    frm.Caption = "Виберіть підрозділ"
    frm.ListData = wsResiduse.Range("A7:A" & lastRow)
    frm.show

    Dim selectedUnit As String
    If frm.ListBox1.ListIndex <> -1 Then
        selectedUnit = frm.ListBox1.value
    Else
        MsgBox "Підрозділ не вибрано!", vbExclamation
        Exit Sub
    End If

    wsResiduse.Range("selectedUnit").value = selectedUnit
    wsResiduse.Range("A7:A" & lastRow).ClearContents

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim i As Long
    Dim key As Variant
    Dim dataCount As Long
    dataCount = tbAllData.ListRows.count

    ' Очистити старі дані
    wsResiduse.Range("b8:i" & wsResiduse.Rows.count).ClearContents

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
        uFrom = Trim(CStr(tbAllData.DataBodyRange(i, cFrom).value))
        uTo = Trim(CStr(tbAllData.DataBodyRange(i, cTo).value))

        If uFrom = selectedUnit Or uTo = selectedUnit Then
            opType = LCase(Trim(CStr(tbAllData.DataBodyRange(i, cOpType).value)))
            valIn = 0#: valOut = 0#
            If IsNumeric(tbAllData.DataBodyRange(i, cIn).value) Then valIn = CDbl(tbAllData.DataBodyRange(i, cIn).value)
            If IsNumeric(tbAllData.DataBodyRange(i, cOut).value) Then valOut = CDbl(tbAllData.DataBodyRange(i, cOut).value)

            If uTo = selectedUnit Then
                dynamicCat = CStr(tbAllData.DataBodyRange(i, cCat3).value)
                If Len(Trim(dynamicCat)) = 0 Then dynamicCat = CStr(tbAllData.DataBodyRange(i, cCat).value)
            Else
                dynamicCat = CStr(tbAllData.DataBodyRange(i, cCat).value)
            End If

            key = CStr(tbAllData.DataBodyRange(i, cCode).value) & "||" & _
                  CStr(tbAllData.DataBodyRange(i, cName).value) & "||" & _
                  CStr(tbAllData.DataBodyRange(i, cPrice).value) & "||" & _
                  CStr(tbAllData.DataBodyRange(i, cUnit).value) & "||" & _
                  dynamicCat & "||" & _
                  CStr(tbAllData.DataBodyRange(i, cNote).value)

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
    totalKeys = dict.count
    d = 0

    For k = 1 To totalKeys
        If Round(dict(allKeys(k - 1)), 6) <> 0 Then
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
        CleanFileName(WbNew.Worksheets(1).Range("selectedUnit").value & _
        fileName & ".xlsx"), FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True

    MsgBox "Файл з залишками матеріальних цінностей збережений у тій же папці, де файл бази даних", vbInformation, ThisWorkbook.Path

End Sub


