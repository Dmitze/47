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
        NameOfThings = frm.TextBox1.Value
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
        If wsResiduse.Cells(rowYes, 1).Value = "так" Then
            critName = CStr(rngWithName.Offset(, -12).Value)
            useName = True
        End If
        If wsResiduse.Cells(rowYes, 2).Value = "так" Then
            critPrice = Val(Replace(CStr(rngWithName.Offset(, -11).Value), ",", "."))
            usePrice = True
        End If
        If wsResiduse.Cells(rowYes, 3).Value = "так" Then
            critUnit = CStr(rngWithName.Offset(, -7).Value)
            useUnit = True
        End If
        If wsResiduse.Cells(rowYes, 4).Value = "так" Then
            critCat = CStr(rngWithName.Offset(, -8).Value)
            useCat = True
        End If
        If wsResiduse.Cells(rowYes, 5).Value = "так" Then
            critNote = CStr(rngWithName.Offset(, -9).Value)
            useNote = True
        End If
        If wsResiduse.Cells(rowYes, 6).Value = "так" Then
            critNom = CStr(rngWithName.Offset(, -6).Value)
            useNom = True
        End If
        If wsResiduse.Cells(rowYes, 7).Value = "так" Then
            crit11 = CStr(rngWithName.Offset(, -4).Value)
            use11 = True
        End If
        If wsResiduse.Cells(rowYes, 8).Value = "так" Then
            crit12 = CStr(rngWithName.Offset(, -3).Value)
            use12 = True
        End If

        With tbGoodCriteria.DataBodyRange.Rows(1)
            .Cells(1, 1).Value = critName
            .Cells(1, 2).Value = critPrice
            .Cells(1, 3).Value = critUnit
            .Cells(1, 4).Value = critCat
            .Cells(1, 5).Value = critNote
            .Cells(1, 6).Value = critNom
            .Cells(1, 7).Value = crit11
            .Cells(1, 8).Value = crit12
        End With
    End If

    ' Очищаємо старі дані таблиці починаючи з рядка 10
    wsResiduse.Range("A10:J" & wsResiduse.Rows.Count).ClearContents

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
    dataCount = tbAllData.ListRows.Count

    Dim rName As String, rUnit As String, rCode As String, rNote As String
    Dim r11 As String, r12 As String, rFrom As String, rTo As String, opType As String
    Dim rPrice As Double, vIn As Double, vOut As Double, movQty As Double
    Dim matchItem As Boolean

    For i = 1 To dataCount
        rName = CStr(tbAllData.DataBodyRange(i, cName).Value)
        
        ' Перевіряємо базову відповідність обраному найменуванню
        If InStr(1, rName, NameOfThings, vbTextCompare) > 0 Or (useName And rName = critName) Then
            matchItem = True
            
            If usePrice Then
                If Val(Replace(CStr(tbAllData.DataBodyRange(i, cPrice).Value), ",", ".")) <> critPrice Then matchItem = False
            End If
            If useUnit And matchItem Then
                If CStr(tbAllData.DataBodyRange(i, cUnit).Value) <> critUnit Then matchItem = False
            End If
            If useNom And matchItem Then
                If CStr(tbAllData.DataBodyRange(i, cCode).Value) <> critNom Then matchItem = False
            End If
            If useNote And matchItem Then
                If CStr(tbAllData.DataBodyRange(i, cNote).Value) <> critNote Then matchItem = False
            End If
            If use11 And matchItem Then
                If CStr(tbAllData.DataBodyRange(i, c11).Value) <> crit11 Then matchItem = False
            End If
            If use12 And matchItem Then
                If CStr(tbAllData.DataBodyRange(i, c12).Value) <> crit12 Then matchItem = False
            End If

            If matchItem Then
                rFrom = Trim(CStr(tbAllData.DataBodyRange(i, cFrom).Value))
                rTo = Trim(CStr(tbAllData.DataBodyRange(i, cTo).Value))
                opType = LCase(Trim(CStr(tbAllData.DataBodyRange(i, cOpType).Value)))

                vIn = 0#: vOut = 0#
                If IsNumeric(tbAllData.DataBodyRange(i, cIn).Value) Then vIn = CDbl(tbAllData.DataBodyRange(i, cIn).Value)
                If IsNumeric(tbAllData.DataBodyRange(i, cOut).Value) Then vOut = CDbl(tbAllData.DataBodyRange(i, cOut).Value)

                Dim itemInfo As String, catTo As String, catFrom As String
                catTo = CStr(tbAllData.DataBodyRange(i, cCat3).Value)
                If Len(Trim(catTo)) = 0 Then catTo = CStr(tbAllData.DataBodyRange(i, cCat).Value)
                catFrom = CStr(tbAllData.DataBodyRange(i, cCat).Value)

                Select Case opType
                    Case "надходження"
                        If Len(rTo) > 0 And vIn <> 0 Then
                            itemInfo = rName & "||" & CStr(tbAllData.DataBodyRange(i, cPrice).Value) & "||" & _
                                       CStr(tbAllData.DataBodyRange(i, cUnit).Value) & "||" & catTo & "||" & _
                                       CStr(tbAllData.DataBodyRange(i, cNote).Value) & "||" & _
                                       CStr(tbAllData.DataBodyRange(i, cCode).Value) & "||" & _
                                       CStr(tbAllData.DataBodyRange(i, c11).Value) & "||" & _
                                       CStr(tbAllData.DataBodyRange(i, c12).Value) & "||" & rTo
                            If Not dictUnits.exists(itemInfo) Then dictUnits.Add itemInfo, 0#
                            dictUnits(itemInfo) = dictUnits(itemInfo) + vIn
                        End If

                    Case "вибуття"
                        If Len(rFrom) > 0 And vOut <> 0 Then
                            itemInfo = rName & "||" & CStr(tbAllData.DataBodyRange(i, cPrice).Value) & "||" & _
                                       CStr(tbAllData.DataBodyRange(i, cUnit).Value) & "||" & catFrom & "||" & _
                                       CStr(tbAllData.DataBodyRange(i, cNote).Value) & "||" & _
                                       CStr(tbAllData.DataBodyRange(i, cCode).Value) & "||" & _
                                       CStr(tbAllData.DataBodyRange(i, c11).Value) & "||" & _
                                       CStr(tbAllData.DataBodyRange(i, c12).Value) & "||" & rFrom
                            If Not dictUnits.exists(itemInfo) Then dictUnits.Add itemInfo, 0#
                            dictUnits(itemInfo) = dictUnits(itemInfo) - vOut
                        End If

                    Case "внутрішнє переміщення"
                        movQty = vOut
                        If movQty = 0 Then movQty = vIn

                        If movQty <> 0 Then
                            If Len(rTo) > 0 Then
                                itemInfo = rName & "||" & CStr(tbAllData.DataBodyRange(i, cPrice).Value) & "||" & _
                                           CStr(tbAllData.DataBodyRange(i, cUnit).Value) & "||" & catTo & "||" & _
                                           CStr(tbAllData.DataBodyRange(i, cNote).Value) & "||" & _
                                           CStr(tbAllData.DataBodyRange(i, cCode).Value) & "||" & _
                                           CStr(tbAllData.DataBodyRange(i, c11).Value) & "||" & _
                                           CStr(tbAllData.DataBodyRange(i, c12).Value) & "||" & rTo
                                If Not dictUnits.exists(itemInfo) Then dictUnits.Add itemInfo, 0#
                                dictUnits(itemInfo) = dictUnits(itemInfo) + movQty
                            End If

                            If Len(rFrom) > 0 Then
                                itemInfo = rName & "||" & CStr(tbAllData.DataBodyRange(i, cPrice).Value) & "||" & _
                                           CStr(tbAllData.DataBodyRange(i, cUnit).Value) & "||" & catFrom & "||" & _
                                           CStr(tbAllData.DataBodyRange(i, cNote).Value) & "||" & _
                                           CStr(tbAllData.DataBodyRange(i, cCode).Value) & "||" & _
                                           CStr(tbAllData.DataBodyRange(i, c11).Value) & "||" & _
                                           CStr(tbAllData.DataBodyRange(i, c12).Value) & "||" & rFrom
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
    totalKeys = dictUnits.Count

    For k = 1 To totalKeys
        currentQty = dictUnits(allKeys(k - 1))
        If Round(currentQty, 6) <> 0 Then
            parts = Split(allKeys(k - 1), "||")
            wsResiduse.Cells(rowIndex, 1).Value = parts(0) ' Найменування
            wsResiduse.Cells(rowIndex, 2).Value = CDbl(parts(1)) ' Ціна
            wsResiduse.Cells(rowIndex, 3).Value = parts(2) ' Одиниця виміру
            wsResiduse.Cells(rowIndex, 4).Value = parts(3) ' Категорія
            wsResiduse.Cells(rowIndex, 5).Value = parts(4) ' Примітка
            wsResiduse.Cells(rowIndex, 6).Value = parts(5) ' Код номенклатури
            wsResiduse.Cells(rowIndex, 7).Value = parts(6) ' Поле 11
            wsResiduse.Cells(rowIndex, 8).Value = parts(7) ' Поле 12
            wsResiduse.Cells(rowIndex, 9).Value = currentQty ' Кількість
            wsResiduse.Cells(rowIndex, 10).Value = parts(8) ' Підрозділ
            rowIndex = rowIndex + 1
        End If
        If totalKeys > 0 Then Call ProgressBar.Progress(k / totalKeys * 100)
    Next k

    If wsResiduse.AutoFilterMode Then wsResiduse.AutoFilterMode = False
    wsResiduse.Range("A9").CurrentRegion.AutoFilter

    MsgBox "Пошук завершено. Знайдено підрозділів з залишками: " & (rowIndex - 10), vbInformation, "Залишки по майну"
End Sub
