Attribute VB_Name = "checkErrors"
Sub HighlightRowsWithAllChecks()
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim rw As ListRow
    Dim fields As Variant
    Dim colIndex As Object
    Dim colName As Variant
    Dim cell As Range
    Dim allowedValues As Variant
    Dim allowedZap2 As Variant
    Dim val As String
    Dim commentText As String
    Dim hasIssue As Boolean
    
    Set ws = ThisWorkbook.Worksheets("База")
    Set lo = ws.ListObjects("AllData")
    
    ' Колонки для перевірки
    fields = Array("Категорія", "Запасне поле2", "Запасне поле3", _
                   "Тип документа", "Ціна", "Код номенклатури", _
                   "№ документа", "Дата документа", "Куди", _
                   "Звідки", "Одиниця виміру", "Найменування", _
                   "Надійшло", "Вибуло")
    
    ' Індекс колонок
    Set colIndex = CreateObject("Scripting.Dictionary")
    For Each colName In fields
        On Error Resume Next
        colIndex(colName) = lo.ListColumns(colName).Index
        On Error GoTo 0
        If colIndex(colName) = 0 Then
            MsgBox "Не знайдено колонку: " & colName, vbExclamation
            Exit Sub
        End If
    Next colName
    
    ' Очистка попередніх підсвіток і приміток
    lo.DataBodyRange.Interior.ColorIndex = xlNone
    For Each colName In fields
        For Each cell In lo.ListColumns(colIndex(colName)).DataBodyRange
            If Not cell.Comment Is Nothing Then cell.Comment.Delete
        Next cell
    Next colName
    
    ' Дозволені значення
    allowedValues = Array("1", "2", "3", "4", "5", "пр.", "не пр.", "пот. рем.")
    allowedZap2 = Array("надходження", "вибуття", "внутрішнє переміщення")
    
    ' Перебір рядків
    For Each rw In lo.ListRows
        hasIssue = False
        
        ' Перевірка окремих колонок
        For Each colName In fields
            Set cell = rw.Range.Cells(1, colIndex(colName))
            commentText = ""
            val = Trim(cell.value)
            
            Select Case colName
                Case "Категорія", "Запасне поле3"
                    If Len(val) = 0 Or IsError(Application.Match(val, allowedValues, 0)) Then
                        commentText = "Поле не повинно бути пустим" & vbNewLine & _
                                      "Допустимі значення для поля: " & Join(allowedValues, ", ")
                        hasIssue = True
                    End If
                Case "Запасне поле2"
                    If Len(val) = 0 Or IsError(Application.Match(val, allowedZap2, 0)) Then
                        commentText = "Поле не повинно бути пустим" & vbNewLine & _
                                      "Допустимі значення для поля: " & Join(allowedZap2, ", ")
                        hasIssue = True
                    End If
                Case "Дата документа"
                    If Len(val) = 0 Or Not IsDate(val) Then
                        commentText = "Поле не повинно бути пустим" & vbNewLine & "Значення має бути датою"
                        hasIssue = True
                    End If
                Case "Ціна"
                    If Len(val) = 0 Or Not IsValidNumber(val) Then
                        commentText = "Поле не повинно бути пустим" & vbNewLine & "Значення має бути числом"
                        hasIssue = True
                    End If
                Case "Найменування"
                    If Len(val) = 0 Then
                        commentText = "Поле не повинно бути пустим"
                        hasIssue = True
                    ElseIf InStr(val, "~") > 0 Or InStr(val, Chr(10)) > 0 Or InStr(val, Chr(13)) > 0 Then
                        commentText = "Не допускаються перенос рядків та символ ~"
                        hasIssue = True
                    End If
                Case "Надійшло", "Вибуло"
                    If Len(val) > 0 And Not IsValidNumber(val) Then
                        commentText = "Значення має бути числом"
                        hasIssue = True
                    End If
                Case Else
                    If Len(val) = 0 Then
                        commentText = "Поле не повинно бути пустим"
                        hasIssue = True
                    End If
            End Select
            
            ' Додаємо примітку
            If Len(commentText) > 0 Then cell.AddComment commentText
        Next colName
        
        ' Перевірка одночасного заповнення Надійшло і Вибуло
        Dim valNad As String, valVyb As String
        valNad = Trim(rw.Range.Cells(1, colIndex("Надійшло")).value)
        valVyb = Trim(rw.Range.Cells(1, colIndex("Вибуло")).value)
        If Len(valNad) > 0 And Len(valVyb) > 0 Then
            rw.Range.Cells(1, colIndex("Надійшло")).AddComment _
                "Кількість має бути або в Надійшло або у Вибуло. Кожна операція окремий рядок"
            rw.Range.Cells(1, colIndex("Вибуло")).AddComment _
                "Кількість має бути або в Надійшло або у Вибуло. Кожна операція окремий рядок"
            hasIssue = True
        End If
        
        ' Перевірка узгодженості Надійшло - Запасне поле2
        Dim valZap2 As String
        valZap2 = Trim(rw.Range.Cells(1, colIndex("Запасне поле2")).value)
        If Len(valNad) > 0 And LCase(valZap2) <> "надходження" Then
            rw.Range.Cells(1, colIndex("Запасне поле2")).AddComment _
                "Змініть тип операції або перемістіть кількість з Надійшло у Вибуло"
            hasIssue = True
        End If
        If Len(valNad) = 0 And LCase(valZap2) = "надходження" Then
            rw.Range.Cells(1, colIndex("Запасне поле2")).AddComment _
                "Змініть тип операції або перемістіть кількість з Надійшло у Вибуло"
            hasIssue = True
        End If
        
        ' Підсвітка рядка
        If hasIssue Then rw.Range.Interior.Color = vbYellow
    Next rw
    MsgBox "Жовтим кольором підсвічені рядки з помилками. Читайте примітки в клітинках з помилками.", vbInformation, "Перевірку завершено"
End Sub

'------------------------------------------
Function IsValidNumber(s As String) As Boolean
    Dim i As Long, dotCount As Long, ch As String
    s = Trim(s)
    If Len(s) = 0 Then IsValidNumber = False: Exit Function
    
    dotCount = 0
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        Select Case ch
            Case "0" To "9"
            Case ","
                dotCount = dotCount + 1
                If dotCount > 1 Then IsValidNumber = False: Exit Function
            Case Else
                IsValidNumber = False
                Exit Function
        End Select
    Next i
    IsValidNumber = True
End Function


