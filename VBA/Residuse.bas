Attribute VB_Name = "Residuse"

Option Explicit

Sub UpdateResiduse()

    Dim wbAllData As Workbook
    Dim wsAllData As Worksheet
    Dim wsResiduse As Worksheet
    Dim tbAllData As ListObject

    Set wbAllData = ThisWorkbook
    Set wsAllData = wbAllData.Worksheets("База")
    Set tbAllData = wsAllData.ListObjects("AllData")
    Set wsResiduse = wbAllData.Worksheets("Залишки по підрозділах")
 
    Dim dict As Object
    Dim cell As Range
    Dim colFrom As Range, colTo As Range
    Dim i As Long
    
    wsResiduse.Range("D6", wsResiduse.Cells(wsResiduse.Rows.count, wsResiduse.Columns.count)).ClearContents
    
    ' Посилання на стовпці таблиці
    Set colFrom = tbAllData.ListColumns(8).DataBodyRange ' Підрозділ звідки
    Set colTo = tbAllData.ListColumns(9).DataBodyRange   ' Підрозділ куди
    
    ' Словник для зберігання унікальних значень
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Унікальні значення з першого стовпця
    For Each cell In colFrom
        If Not dict.exists(cell.value) And Len(Trim(CStr(cell.value))) > 0 Then
            dict.Add cell.value, Nothing
        End If
    Next cell
    
    ' Унікальні значення з другого стовпця
    For Each cell In colTo
        If Not dict.exists(cell.value) And Len(Trim(CStr(cell.value))) > 0 Then
            dict.Add cell.value, Nothing
        End If
    Next cell
    
    Dim arr() As Variant
    Dim j As Long, temp As Variant
    
    arr = dict.Keys
    
    ' Сортування масиву підрозділів
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) > arr(j) Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
    
    Dim k As Long
    ' Запис унікальних значень у рядок на аркуші
    For i = LBound(arr) To UBound(arr)
        k = i + 3
        wsResiduse.Cells(6, k + 1).value = arr(i)
    Next i
    
    Dim ptUnitResiduse As PivotTable
    On Error Resume Next
    Set ptUnitResiduse = wsResiduse.PivotTables("ptUnitResiduse")
    If Not ptUnitResiduse Is Nothing Then ptUnitResiduse.RefreshTable
    On Error GoTo 0
    
    Dim totalSum As Double
    Dim unitName As String
    Dim destination As String
    
    Dim s As Long, l As Long
    Dim lastRow As Long
    Dim lastCol As Long
    
    lastRow = wsResiduse.Cells(wsResiduse.Rows.count, "A").End(xlUp).Row
    lastCol = wsResiduse.Cells(6, wsResiduse.Columns.count).End(xlToLeft).Column
    
    On Error GoTo CleanExit
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    For s = 8 To lastRow
        unitName = wsResiduse.Cells(s, 1).value ' Значення з колонки A
        
        For l = 4 To lastCol ' Починаємо зі стовпця 4 (D)
            destination = wsResiduse.Cells(6, l).value ' Значення з рядка 6
            
            ' Універсальний розрахунок залишків:
            ' Додаємо отримане підрозділом destination (через Надійшло або Вибуло)
            ' Віднімаємо передане підрозділом destination (через Надійшло або Вибуло)
            totalSum = WorksheetFunction.SumIfs(tbAllData.ListColumns("Надійшло").DataBodyRange, _
                                                tbAllData.ListColumns("Куди").DataBodyRange, destination, _
                                                tbAllData.ListColumns("Найменування").DataBodyRange, unitName) _
                      + WorksheetFunction.SumIfs(tbAllData.ListColumns("Вибуло").DataBodyRange, _
                                                tbAllData.ListColumns("Куди").DataBodyRange, destination, _
                                                tbAllData.ListColumns("Найменування").DataBodyRange, unitName) _
                      - WorksheetFunction.SumIfs(tbAllData.ListColumns("Надійшло").DataBodyRange, _
                                                tbAllData.ListColumns("Звідки").DataBodyRange, destination, _
                                                tbAllData.ListColumns("Найменування").DataBodyRange, unitName) _
                      - WorksheetFunction.SumIfs(tbAllData.ListColumns("Вибуло").DataBodyRange, _
                                                tbAllData.ListColumns("Звідки").DataBodyRange, destination, _
                                                tbAllData.ListColumns("Найменування").DataBodyRange, unitName)
            
            wsResiduse.Cells(s, l).value = totalSum
        Next l
        If lastRow > 0 Then Call ProgressBar.Progress(s / lastRow * 100)
    Next s

CleanExit:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub


