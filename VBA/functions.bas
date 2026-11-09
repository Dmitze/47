Attribute VB_Name = "functions"
Function CleanFileName(str As String) As String
    Dim invalidChars As String
    Dim i As Integer
    Dim char As String
    
    ' Задаємо всі недопустимі символи, включаючи перенесення рядка
    invalidChars = "/\:*?""<>|" & Chr(10) & Chr(13)
    
    ' Проходимо через кожен символ в рядку
    For i = 1 To Len(str)
        char = Mid(str, i, 1)
        ' Якщо символ є недопустимим, замінюємо його на дефіс
        If InStr(invalidChars, char) > 0 Then
            CleanFileName = CleanFileName & "-"
        Else
            CleanFileName = CleanFileName & char
        End If
    Next i
End Function
Function FindValue(searchValue As String, searchRange As Range) As Boolean
    Dim foundCell As Range

    'Шукаємо задане значення в діапазоні
    Set foundCell = searchRange.Find(searchValue, LookIn:=xlValues, LookAt:=xlWhole)

    'Перевіряємо, чи знайдено задане значення
    If Not foundCell Is Nothing Then
        FindValue = True
    Else
        FindValue = False
    End If
End Function
'Пошук документа в рядку (повторюваного значення)
Function IsDocumentInList(docList As String, docName As String) As Boolean
    Dim docs() As String
    Dim i As Long
    
    docs = Split(docList, ";")
    For i = LBound(docs) To UBound(docs)
        If Trim(docs(i)) = Trim(docName) Then
            IsDocumentInList = True
            Exit Function
        End If
    Next i
    IsDocumentInList = False
End Function

'очистка фільтрів на сторінці
Sub ClearUnitFilter_Residues(ByVal ws As Worksheet)
    On Error Resume Next
    If ws.FilterMode Then ws.ShowAllData
    On Error GoTo 0
End Sub

'функція для пошуку матеріалів тільки в конкретному підрозділі
Function GetResidusePool(ByVal wsResiduse2 As Worksheet, ByVal FromUnit As String) As Range

    Dim lastRow As Long
    lastRow = wsResiduse2.Cells(wsResiduse2.Rows.count, "N").End(xlUp).Row
    If lastRow < 3 Then
        MsgBox "У залишках немає даних.", vbExclamation
        Exit Function
    End If

    Dim AllOrOneUnit As VbMsgBoxResult
    AllOrOneUnit = MsgBox( _
        "Шукати лише в підрозділі:" & vbNewLine & _
        FromUnit & " - так" & vbNewLine & _
        "шукати по всій базі - ні", _
        vbYesNo, "Де шукати потрібне майно?" _
    )

    Dim ResiduseCheet As Range
    Set ResiduseCheet = wsResiduse2.Range("N3:N" & lastRow) 'костиль для масиву

    If AllOrOneUnit = vbYes Then

        Dim hdr As Range
        Set hdr = wsResiduse2.Range("A2").CurrentRegion

        On Error Resume Next
        If wsResiduse2.FilterMode Then wsResiduse2.ShowAllData
        On Error GoTo 0

        hdr.AutoFilter Field:=12, Criteria1:=FromUnit

        On Error Resume Next
        Set GetResidusePool = ResiduseCheet.SpecialCells(xlCellTypeVisible)
        On Error GoTo 0

        If GetResidusePool Is Nothing Then
'
' Application.WorksheetFunction.CountA(GetResidusePool) = 0
            MsgBox "По підрозділу '" & FromUnit & "' немає позицій.", vbExclamation
            Call ClearUnitFilter_Residues(wsResiduse2)
            Exit Function
        End If

        Call ClearUnitFilter_Residues(wsResiduse2)

    Else
        Set GetResidusePool = ResiduseCheet
    End If

End Function

'===========3.9.16=============
'Function GetSelectedRowsArray()
'
'    Dim dict As Object
'    Dim cell As Range
'    Dim lo As ListObject
'    Dim rowIndex As Long
'
'    Set dict = CreateObject("Scripting.Dictionary")
'    Set lo = Worksheets("База").ListObjects("AllData")
'
'    For Each cell In Selection
'
'        If Not Intersect(cell, lo.DataBodyRange) Is Nothing Then
'
'            If cell.EntireRow.Hidden = False Then
'
'                rowIndex = cell.Row - lo.DataBodyRange.Row + 1
'
'                If Not dict.exists(rowIndex) Then
'                    dict.Add rowIndex, rowIndex
'                End If
'
'            End If
'
'        End If
'
'    Next cell
'
'    GetSelectedRowsArray = dict.Keys
'
'End Function

'===========3.9.16=============
'===========3.9.17=============

Function GetSelectedRowsArrayCell()

    Dim dict As Object
    Dim cell As Range
    
    Set dict = CreateObject("Scripting.Dictionary")

    For Each cell In Selection
        
        If cell.EntireRow.Hidden = False Then
            
            If Not dict.exists(cell.Row) Then
                dict.Add cell.Row, cell.Row
            End If
            
        End If
        
    Next cell
    
    GetSelectedRowsArrayCell = dict.Keys

End Function


Sub RemoveCustomContextMenu()

    Dim cb As CommandBar
    
    On Error Resume Next
    Set cb = Application.CommandBars("Cell")
    cb.Controls("Додати в документ").Delete
    On Error GoTo 0

End Sub
'===========3.9.17=============
