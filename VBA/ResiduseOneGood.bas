Attribute VB_Name = "ResiduseOneGood"
Option Explicit

Sub ResiduseForGood()
    Dim wbAllData As Workbook
    Dim wsAllData As Worksheet
    Dim wsResiduse As Worksheet
    Dim wsResiduse2 As Worksheet
    Dim tbAllData As ListObject
    Dim tbGoodCriteria As ListObject

    Set wbAllData = ThisWorkbook
    Set wsAllData = wbAllData.Worksheets("База")
    Set tbAllData = wsAllData.ListObjects("AllData")
    Set wsResiduse = wbAllData.Worksheets("Залишки по майну")
    Set wsResiduse2 = wbAllData.Worksheets("Залишки по підрозділах2")
    Set tbGoodCriteria = wsResiduse.ListObjects("tbGoodCriteria")

    ' Очищаємо критерії попереднього запиту
    tbGoodCriteria.DataBodyRange.ClearContents

    Dim frm As UserForm1
    Set frm = New UserForm1

    frm.Caption = "Виберіть майно"
    frm.ListData = wsResiduse2.Range("Goods")
    frm.show

    Dim NameOfThings As String
    Dim ListGoods As Range
    Set ListGoods = wsResiduse2.Range("Goods")

    If frm.ListBox1.ListIndex <> -1 Then
        NameOfThings = frm.TextBox1.value
    Else
        MsgBox "Майно не вибрано!", vbExclamation
        Exit Sub
    End If

    Dim rngWithName As Range
    Set rngWithName = ListGoods.Find(NameOfThings, LookIn:=xlValues, LookAt:=xlWhole)

    Dim critName As String, critPrice As Double, critUnit As String
    Dim critCat As String, critNote As String, critNom As String
    Dim crit11 As String, crit12 As String
    Dim useName As Boolean, usePrice As Boolean, useUnit As Boolean
    Dim useCat As Boolean, useNote As Boolean, useNom As Boolean
    Dim use11 As Boolean, use12 As Boolean

    Dim rowYes As Long
    rowYes = 8

    If Not rngWithName Is Nothing Then
        If wsResiduse.Cells(rowYes, 1).value = "так" Then
            critName = CStr(rngWithName.Offset(, -12).value)
            useName = True
        End If
        If wsResiduse.Cells(rowYes, 2).value = "так" Then
            critPrice = val(Replace(CStr(rngWithName.Offset(, -11).value), ",", "."))
            usePrice = True
        End If
        If wsResiduse.Cells(rowYes, 3).value = "так" Then
            critUnit = CStr(rngWithName.Offset(, -7).value)
            useUnit = True
        End If
        If wsResiduse.Cells(rowYes, 4).value = "так" Then
            critCat = CStr(rngWithName.Offset(, -8).value)
            useCat = True
        End If
        If wsResiduse.Cells(rowYes, 5).value = "так" Then
            critNote = CStr(rngWithName.Offset(, -9).value)
            useNote = True
        End If
        If wsResiduse.Cells(rowYes, 6).value = "так" Then
            critNom = CStr(rngWithName.Offset(, -6).value)
            useNom = True
        End If
        If wsResiduse.Cells(rowYes, 7).value = "так" Then
            crit11 = CStr(rngWithName.Offset(, -4).value)
            use11 = True
        End If
        If wsResiduse.Cells(rowYes, 8).value = "так" Then
            crit12 = CStr(rngWithName.Offset(, -3).value)
            use12 = True
        End If

        With tbGoodCriteria.DataBodyRange.Rows(1)
            .Cells(1, 1).value = critName
            .Cells(1, 2).value = critPrice
            .Cells(1, 3).value = critUnit
            .Cells(1, 4).value = critCat
            .Cells(1, 5).value = critNote
            .Cells(1, 6).value = critNom
            .Cells(1, 7).value = crit11
            .Cells(1, 8).value = crit12
        End With
    End If

    ' Очищаємо старі дані таблиці починаючи з рядка 10
    wsResiduse.Range("A10:J" & wsResiduse.Rows.count).ClearContents

    Dim cFrom As Long, cTo As Long, cIn As Long, cOut As Long, cOpType As Long
    Dim cName As Long, cCode As Long, cPrice As Long, cUnit As Long
    Dim cCat As Long, cCat3 As Long, cNote As Long, c11 As Long, c12 As Long

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
        c11 = .ListColumns("Поле 11").Index
        c12 = .ListColumns("Поле 12").Index
    End With

    ' Словник: Unit -> Key (де Key містить параметри майна) -> Кількість
    Dim dictUnits As Object
    Set dictUnits = CreateObject("Scripting.Dictionary")

    Dim i As Long, dataCount As Long
    dataCount = tbAllData.ListRows.count

    Dim rName As String, rUnit As String, rCode As String, rNote As String
    Dim r11 As String, r12 As String, rFrom As String, rTo As String, opType As String
    Dim rPrice As Double, vIn As Double, vOut As Double, movQty As Double
    Dim matchItem As Boolean

    For i = 1 To dataCount
        rName = CStr(tbAllData.DataBodyRange(i, cName).value)
        
        ' Перевіряємо базову відповідність обраному найменуванню
        If InStr(1, rName, NameOfThings, vbTextCompare) > 0 Or (useName And rName = critName) Then
            matchItem = True
            
            If usePrice Then
                If val(Replace(CStr(tbAllData.DataBodyRange(i, cPrice).value), ",", ".")) <> critPrice Then matchItem = False
            End If
            If useUnit And matchItem Then
                If CStr(tbAllData.DataBodyRange(i, cUnit).value) <> critUnit Then matchItem = False
            End If
            If useNom And matchItem Then
                If CStr(tbAllData.DataBodyRange(i, cCode).value) <> critNom Then matchItem = False
            End If
            If useNote And matchItem Then
                If CStr(tbAllData.DataBodyRange(i, cNote).value) <> critNote Then matchItem = False
            End If
            If use11 And matchItem Then
                If CStr(tbAllData.DataBodyRange(i, c11).value) <> crit11 Then matchItem = False
            End If
            If use12 And matchItem Then
                If CStr(tbAllData.DataBodyRange(i, c12).value) <> crit12 Then matchItem = False
            End If

            If matchItem Then
                rFrom = Trim(CStr(tbAllData.DataBodyRange(i, cFrom).value))
                rTo = Trim(CStr(tbAllData.DataBodyRange(i, cTo).value))
                opType = LCase(Trim(CStr(tbAllData.DataBodyRange(i, cOpType).value)))

                vIn = 0#: vOut = 0#
                If IsNumeric(tbAllData.DataBodyRange(i, cIn).value) Then vIn = CDbl(tbAllData.DataBodyRange(i, cIn).value)
                If IsNumeric(tbAllData.DataBodyRange(i, cOut).value) Then vOut = CDbl(tbAllData.DataBodyRange(i, cOut).value)

                Dim itemInfo As String, catTo As String, catFrom As String
                catTo = CStr(tbAllData.DataBodyRange(i, cCat3).value)
                If Len(Trim(catTo)) = 0 Then catTo = CStr(tbAllData.DataBodyRange(i, cCat).value)
                catFrom = CStr(tbAllData.DataBodyRange(i, cCat).value)

                Select Case opType
                    Case "надходження"
                        If Len(rTo) > 0 And vIn <> 0 Then
                            itemInfo = rName & "||" & CStr(tbAllData.DataBodyRange(i, cPrice).value) & "||" & _
                                       CStr(tbAllData.DataBodyRange(i, cUnit).value) & "||" & catTo & "||" & _
                                       CStr(tbAllData.DataBodyRange(i, cNote).value) & "||" & _
                                       CStr(tbAllData.DataBodyRange(i, cCode).value) & "||" & _
                                       CStr(tbAllData.DataBodyRange(i, c11).value) & "||" & _
                                       CStr(tbAllData.DataBodyRange(i, c12).value) & "||" & rTo
                            If Not dictUnits.exists(itemInfo) Then dictUnits.Add itemInfo, 0#
                            dictUnits(itemInfo) = dictUnits(itemInfo) + vIn
                        End If

                    Case "вибуття"
                        If Len(rFrom) > 0 And vOut <> 0 Then
                            itemInfo = rName & "||" & CStr(tbAllData.DataBodyRange(i, cPrice).value) & "||" & _
                                       CStr(tbAllData.DataBodyRange(i, cUnit).value) & "||" & catFrom & "||" & _
                                       CStr(tbAllData.DataBodyRange(i, cNote).value) & "||" & _
                                       CStr(tbAllData.DataBodyRange(i, cCode).value) & "||" & _
                                       CStr(tbAllData.DataBodyRange(i, c11).value) & "||" & _
                                       CStr(tbAllData.DataBodyRange(i, c12).value) & "||" & rFrom
                            If Not dictUnits.exists(itemInfo) Then dictUnits.Add itemInfo, 0#
                            dictUnits(itemInfo) = dictUnits(itemInfo) - vOut
                        End If

                    Case "внутрішнє переміщення"
                        movQty = vOut
                        If movQty = 0 Then movQty = vIn

                        If movQty <> 0 Then
                            If Len(rTo) > 0 Then
                                itemInfo = rName & "||" & CStr(tbAllData.DataBodyRange(i, cPrice).value) & "||" & _
                                           CStr(tbAllData.DataBodyRange(i, cUnit).value) & "||" & catTo & "||" & _
                                           CStr(tbAllData.DataBodyRange(i, cNote).value) & "||" & _
                                           CStr(tbAllData.DataBodyRange(i, cCode).value) & "||" & _
                                           CStr(tbAllData.DataBodyRange(i, c11).value) & "||" & _
                                           CStr(tbAllData.DataBodyRange(i, c12).value) & "||" & rTo
                                If Not dictUnits.exists(itemInfo) Then dictUnits.Add itemInfo, 0#
                                dictUnits(itemInfo) = dictUnits(itemInfo) + movQty
                            End If

                            If Len(rFrom) > 0 Then
                                itemInfo = rName & "||" & CStr(tbAllData.DataBodyRange(i, cPrice).value) & "||" & _
                                           CStr(tbAllData.DataBodyRange(i, cUnit).value) & "||" & catFrom & "||" & _
                                           CStr(tbAllData.DataBodyRange(i, cNote).value) & "||" & _
                                           CStr(tbAllData.DataBodyRange(i, cCode).value) & "||" & _
                                           CStr(tbAllData.DataBodyRange(i, c11).value) & "||" & _
                                           CStr(tbAllData.DataBodyRange(i, c12).value) & "||" & rFrom
                                If Not dictUnits.exists(itemInfo) Then dictUnits.Add itemInfo, 0#
                                dictUnits(itemInfo) = dictUnits(itemInfo) - movQty
                            End If
                        End If
                End Select
            End If
        End If
    Next i

    ' Виведення результатів з рядка 10
    Dim rowIndex As Long
    rowIndex = 10
    Dim parts() As String
    Dim k As Long, totalKeys As Long, allKeys As Variant
    Dim currentQty As Double

    allKeys = dictUnits.Keys
    totalKeys = dictUnits.count

    For k = 1 To totalKeys
        currentQty = dictUnits(allKeys(k - 1))
        If Round(currentQty, 6) <> 0 Then
            parts = Split(allKeys(k - 1), "||")
            wsResiduse.Cells(rowIndex, 1).value = parts(0) ' Найменування
            wsResiduse.Cells(rowIndex, 2).value = CDbl(parts(1)) ' Ціна
            wsResiduse.Cells(rowIndex, 3).value = parts(2) ' Одиниця виміру
            wsResiduse.Cells(rowIndex, 4).value = parts(3) ' Категорія
            wsResiduse.Cells(rowIndex, 5).value = parts(4) ' Примітка
            wsResiduse.Cells(rowIndex, 6).value = parts(5) ' Код номенклатури
            wsResiduse.Cells(rowIndex, 7).value = parts(6) ' Поле 11
            wsResiduse.Cells(rowIndex, 8).value = parts(7) ' Поле 12
            wsResiduse.Cells(rowIndex, 9).value = currentQty ' Кількість
            wsResiduse.Cells(rowIndex, 10).value = parts(8) ' Підрозділ
            rowIndex = rowIndex + 1
        End If
        If totalKeys > 0 Then Call ProgressBar.Progress(k / totalKeys * 100)
    Next k

    If wsResiduse.AutoFilterMode Then wsResiduse.AutoFilterMode = False
    wsResiduse.Range("A9").CurrentRegion.AutoFilter

    MsgBox "Пошук завершено. Знайдено підрозділів з залишками: " & (rowIndex - 10), vbInformation, "Залишки по майну"
End Sub
