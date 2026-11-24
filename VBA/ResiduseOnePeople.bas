Attribute VB_Name = "ResiduseOnePeople"
Option Explicit

Sub UpdateResiduseOnePeople()

    Dim wbAllData As Workbook
    Dim wsAllData As Worksheet
    Dim wsResiduse As Worksheet
    Dim tbAllData As ListObject

    Set wbAllData = ThisWorkbook
    Set wsAllData = wbAllData.Worksheets("База")
    Set tbAllData = wsAllData.ListObjects("AllData")
    Set wsResiduse = wbAllData.Worksheets("Довідка по людині")

    Dim frm As UserForm1
    Set frm = New UserForm1

    Dim colFrom As Range, colTo As Range, cell As Range
    Dim dictPeople As Object
    Set dictPeople = CreateObject("Scripting.Dictionary")

    Set colFrom = tbAllData.ListColumns("МВО звідки").DataBodyRange
    Set colTo = tbAllData.ListColumns("МВО куди").DataBodyRange

    ' Формуємо словник унікальних людей
    For Each cell In colFrom
        If Not dictPeople.exists(cell.value) And Len(Trim(CStr(cell.value))) > 0 Then
            dictPeople.Add cell.value, Nothing
        End If
    Next cell

    For Each cell In colTo
        If Not dictPeople.exists(cell.value) And Len(Trim(CStr(cell.value))) > 0 Then
            dictPeople.Add cell.value, Nothing
        End If
    Next cell

    ' Виводимо список людей
    Dim wsTroops As Worksheet
    Set wsTroops = wbAllData.Worksheets("Підрозділи")

    Dim k As Long
    k = 1
    wsTroops.Range("B1:B" & wsTroops.Rows.count).ClearContents

    Dim personKey As Variant
    For Each personKey In dictPeople.Keys
        wsTroops.Cells(k, 2).value = personKey
        k = k + 1
    Next personKey

    Dim lastRowT As Long
    lastRowT = wsTroops.Cells(wsTroops.Rows.count, 2).End(xlUp).Row

    frm.Caption = "Виберіть людину"
    frm.ListData = wsTroops.Range("B1:B" & lastRowT)
    frm.show

    Dim selectedPerson As String
    If frm.ListBox1.ListIndex <> -1 Then
        selectedPerson = frm.ListBox1.value
    Else
        MsgBox "Людину не вибрано! Спробуйте ще раз!", vbExclamation
        Exit Sub
    End If

    wsResiduse.Range("selectedPeople").value = selectedPerson

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim i As Long
    Dim key As Variant
    Dim dataCount As Long
    dataCount = tbAllData.ListRows.count

    ' Очистити старі дані
    wsResiduse.Range("A8:I" & wsResiduse.Rows.count).ClearContents

    Dim cMvoFrom As Long, cMvoTo As Long, cIn As Long, cOut As Long, cOpType As Long
    Dim cName As Long, cCode As Long, cPrice As Long, cUnit As Long
    Dim cCat As Long, cCat3 As Long, cNote As Long

    With tbAllData
        cMvoFrom = .ListColumns("МВО звідки").Index
        cMvoTo = .ListColumns("МВО куди").Index
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

    Dim mFrom As String, mTo As String, opType As String
    Dim valIn As Double, valOut As Double, movQty As Double
    Dim dynamicCat As String

    For i = 1 To dataCount
        mFrom = Trim(CStr(tbAllData.DataBodyRange(i, cMvoFrom).value))
        mTo = Trim(CStr(tbAllData.DataBodyRange(i, cMvoTo).value))

        If mFrom = selectedPerson Or mTo = selectedPerson Then
            opType = LCase(Trim(CStr(tbAllData.DataBodyRange(i, cOpType).value)))
            valIn = 0#: valOut = 0#
            If IsNumeric(tbAllData.DataBodyRange(i, cIn).value) Then valIn = CDbl(tbAllData.DataBodyRange(i, cIn).value)
            If IsNumeric(tbAllData.DataBodyRange(i, cOut).value) Then valOut = CDbl(tbAllData.DataBodyRange(i, cOut).value)

            If mTo = selectedPerson Then
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
                    If mTo = selectedPerson Then
                        dict(key) = dict(key) + valIn
                    End If

                Case "вибуття"
                    If mFrom = selectedPerson Then
                        dict(key) = dict(key) - valOut
                    End If

                Case "внутрішнє переміщення"
                    movQty = valOut
                    If movQty = 0 Then movQty = valIn

                    If mTo = selectedPerson Then
                        dict(key) = dict(key) + movQty
                    End If
                    If mFrom = selectedPerson Then
                        dict(key) = dict(key) - movQty
                    End If
            End Select
        End If
    Next i

    ' Виведення результатів
    Dim rowIndex As Long
    rowIndex = 8
    Dim parts() As String
    Dim m As Long, totalKeys As Long, allKeys As Variant
    Dim d As Long

    allKeys = dict.Keys
    totalKeys = dict.count
    d = 0

    For m = 1 To totalKeys
        If Round(dict(allKeys(m - 1)), 6) <> 0 Then
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
            rowIndex = rowIndex + 1
        End If
        If totalKeys > 0 Then Call ProgressBar.Progress(m / totalKeys * 100)
    Next m

    If wsResiduse.AutoFilterMode Then wsResiduse.AutoFilterMode = False
    wsResiduse.Range("A7").CurrentRegion.AutoFilter

End Sub

Sub SaveResiduseOnePeople()

    ThisWorkbook.Sheets("Довідка по людині").Copy
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
        CleanFileName(WbNew.Worksheets(1).Range("selectedPeople").value & _
        fileName & ".xlsx"), FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True

    MsgBox "Файл з залишками матеріальних цінностей збережений у тій же папці, де файл бази даних", vbInformation, ThisWorkbook.Path

End Sub
