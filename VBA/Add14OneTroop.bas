Attribute VB_Name = "Add14OneTroop"
Sub CreateAdd14OneTroop()
    Dim wbAllData As Workbook
    Dim wsAllData As Worksheet
    Dim tbAllData As ListObject
        Set wbAllData = ThisWorkbook
        Set wsAllData = wbAllData.Worksheets("База")
        Set tbAllData = wsAllData.ListObjects("AllData") 'таблиця з базою даних
    
    Dim wsAdd13 As Worksheet
    Set wsAdd13 = wbAllData.Worksheets("Книга обліку д14")
    
    
   Dim frm As UserForm1
Set frm = New UserForm1

Dim colFrom As Range, colTo As Range, celll As Range
Dim dictUnits As Object
Set dictUnits = CreateObject("Scripting.Dictionary")

Set colFrom = tbAllData.ListColumns("Звідки").DataBodyRange
Set colTo = tbAllData.ListColumns("Куди").DataBodyRange

' Формуємо словник унікальних підрозділів
For Each celll In colFrom
    If Not dictUnits.exists(celll.value) And Not IsEmpty(celll.value) Then
        dictUnits.Add celll.value, Nothing
    End If
Next celll

For Each celll In colTo
    If Not dictUnits.exists(celll.value) And Not IsEmpty(celll.value) Then
        dictUnits.Add celll.value, Nothing
    End If
Next celll


Dim wsTroops As Worksheet
Set wsTroops = wbAllData.Worksheets("Підрозділи")

Dim k As Long
k = 1
wsTroops.Range("A1:A" & wsTroops.Rows.count).ClearContents ' Очистити перед вставкою

Dim unitKey As Variant
For Each unitKey In dictUnits.Keys
    wsTroops.Cells(k, 1).value = unitKey
    k = k + 1
Next unitKey

' Зчитати назад у масив
Dim lastRowT As Long
lastRowT = wsTroops.Cells(wsTroops.Rows.count, 1).End(xlUp).Row


frm.Caption = "Виберіть підрозділ"
frm.ListData = wsTroops.Range("A1:A" & lastRowT)

 ' Показати форму
frm.show

' Отримати вибране
If frm.ListBox1.ListIndex <> -1 Then
    selectedUnit = frm.ListBox1.value
Else
    MsgBox "Підрозділ не вибрано!", vbExclamation
    Exit Sub
End If
   
 '''''''''''''''''''''''''''''''''''''''''''''''

    
    'Видаляємо попередні дані
    'видаляємо колонки
Dim lastColumnStart As Long
 lastColumnStart = wsAdd13.Cells(22, wsAdd13.Columns.count).End(xlToLeft).Column + 1
 
 'видаляємо рядки
Dim LastRowStart As Long
 LastRowStart = wsAdd13.Cells.Find(What:="*", After:=wsAdd13.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
                            SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
On Error Resume Next
    wsAdd13.Range(Cells(1, 9), Cells(LastRowStart, lastColumnStart)).EntireColumn.Delete ' Видалити колонки від U до останньої

    wsAdd13.Rows("24:" & wsAdd13.Rows.count).Delete

On Error GoTo 0

Dim RngToSearchData As Range
Set RngToSearchData = wsAdd13.Range("12:12") 'рядок з комбінацією даних в якому робити пошук на предмет співпадіння
    
Dim filteredData() As Variant 'масив з відфільтрованими даними

Dim criteriaRange14OneTroop As Range
Set criteriaRange14OneTroop = wsAdd13.Range("criteriaRange14OneTroop") 'критерії фільтрування, (надійшло, вибуло)будемо міняти на конкретний підрозділ

criteriaRange14OneTroop.Cells(2, 1) = selectedUnit
criteriaRange14OneTroop.Cells(3, 2) = selectedUnit


'знімаємо фільтри,якщо були поставлені
    If wsAllData.FilterMode Then
        wsAllData.ShowAllData
    End If
            
'фільтруємо по заданих критеріях
tbAllData.Range.AdvancedFilter Action:=xlFilterInPlace, criteriaRange:=criteriaRange14OneTroop, Unique:=False


Dim rowCount As Long, colCount As Long
Dim visibleRange As Range
Dim cell As Range
Dim area As Range

 ' Визначаємо кількість видимих рядків і стовпців
    rowCount = 0
    colCount = tbAllData.ListColumns.count
    On Error Resume Next
    Set visibleRange = tbAllData.DataBodyRange.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    
    If Not visibleRange Is Nothing Then
        For Each area In visibleRange.Areas
            rowCount = rowCount + area.Rows.count
        Next area
        
        ' Ініціалізуємо масив з урахуванням заголовків
        ReDim filteredData(1 To rowCount + 1, 1 To colCount)
        
        ' Записуємо заголовки таблиці у масив
        For k = 1 To colCount
            filteredData(1, k) = tbAllData.HeaderRowRange.Cells(1, k).value
        Next k
        
        ' Записуємо відфільтровані дані у масив
        i = 2
        For Each area In visibleRange.Areas
            For Each cell In area.Rows
                For j = 1 To colCount
                    filteredData(i, j) = cell.Cells(1, j).value
                Next j
                i = i + 1
            Next cell
        Next area
            
  End If
'''''''''''''''''''''''''''''''''''''''''''''''''''

Dim lastRow13 As Long
Dim CopyRange1 As Range
Dim foundCell As Range
' заголовок з назвою майна і категоріями
   Set CopyRange1 = wsAdd13.Range(wsAdd13.Cells(13, 6), wsAdd13.Cells(22, 8))
     'визначаємо, звідки копіювати шаблон для нового майна
    
'витягуємо дані з кожного рядка відфільтрованої бази даних
Dim n As Long

'On Error GoTo DontSelect
 For n = LBound(filteredData) + 1 To UBound(filteredData)
 
    Dim Document As String
    Dim DateDocument As Date
    Dim NumberDocument As String
    Dim FromTo As String
    Dim FromUnit As String
    Dim toUnit As String
    
    FromUnit = filteredData(n, 8) 'підрозділ звідки
    toUnit = filteredData(n, 9) 'підрозділ куди
    
    
    Document = filteredData(n, 25)
    DateDocument = filteredData(n, 13)
    NumberDocument = filteredData(n, 14)
    FromTo = FromUnit & Chr(10) & "(" & toUnit & ")"
    
    Dim typeOperation As String
    
    Dim name As String
    Dim Nomenclature As String
    Dim unit As String
    Dim unit2 As String
    Dim price As Double
    Dim DensityCount As Double
    
    Dim previousQuantyty As Double
    Dim previousQuantytyL As Double
    
    Dim actualQuantyty As Double
    Dim actualQuantytyL As Double
    
        typeOperation = filteredData(n, 26) 'тип операції
        
        name = filteredData(n, 2) 'найменування
'        Nomenclature = filteredData(n, 18) 'номенклатура
        unit = filteredData(n, 4) ' одиниця виміру кг
'        unit2 = filteredData(n, 23) ' одиниця виміру л
        Nomenclature = filteredData(n, 18)
        price = filteredData(n, 20) 'ціна
'        DensityCount = filteredData(n, 16) 'густина
        
        
      Dim generalInfo As String
      
      generalInfo = name & Nomenclature & unit & price 'сукупність інформації, за якою будуть формуватися унікальні колонки з найменуваннями
    
    Dim nameExist As Boolean
    nameExist = FindValue(generalInfo, RngToSearchData) 'шукаємо за сукупністю інформації майно
    
    Set foundCell = RngToSearchData.Find(generalInfo, LookIn:=xlValues, LookAt:=xlWhole)

    ''''''''''''''''''''''''''''''''''''''''''''''''''
    lastRow13 = Cells(wsAdd13.Rows.count, "A").End(xlUp).Row + 1 'зназходимо останній рядок і додаємо 1
    Dim LastColumn As Long
                LastColumn = wsAdd13.Cells.Find(What:="*", After:=wsAdd13.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    
    
    Dim lastNumber As Long
                                
        lastNumber = wsAdd13.Cells(22, LastColumn)
    
    'перевірка чи є вже такий документ
    Dim t As Long, DocExist As Boolean
    Dim DocRow As Long

    ' Порівняння кожного рядка масиву з відповідним рядком на листі
    For t = 24 To lastRow13 '24рядок, де починаються документи
        DocExist = True
            'перевірка, чи існує вже первинний документ
            If wsAdd13.Cells(t, 2).value <> Document Or wsAdd13.Cells(t, 3).value <> NumberDocument Or _
            wsAdd13.Cells(t, 4).value <> DateDocument Or wsAdd13.Cells(t, 5).value <> FromTo Then
                DocExist = False
            Else
            'якщо знайшли документ - забираємо рядок в якому він є і виходимо з циклу
                DocRow = wsAdd13.Cells(t, 1).Row
                Exit For
            End If
      
    Next t
    
   If Not DocExist Then
    
        '      Додаємо інформацію про документ
                With wsAdd13
                            .Cells(lastRow13, 1).value = filteredData(n, 1) 'дата внесення
                            .Cells(lastRow13, 2).value = Document 'найменування документа
                            .Cells(lastRow13, 3).NumberFormat = "@"
                            .Cells(lastRow13, 3).value = NumberDocument 'номер документа
                            On Error GoTo WrongDate
                            .Cells(lastRow13, 4).value = DateDocument  'дата документа
                            On Error GoTo 0
                            .Cells(lastRow13, 5).value = FromTo 'підрозділ звідки (підрозділ куди)
                End With
            
        'і ще дублювати всі колонки де є становить...
             Dim cellTotal As Range
             
        
            ' Перевірка кожної клітинки в третьому рядку
            For Each cellTotal In wsAdd13.Rows(21).Cells
                If InStr(cellTotal.value, "становить") > 0 Then
                    ' Копіювання інформації з попереднього рядка до поточного рядка у відповідну клітинку
                    wsAdd13.Cells(lastRow13, cellTotal.Column).value = wsAdd13.Cells(lastRow13 - 1, cellTotal.Column).value
'                    wsAdd13.Cells(lastRow13, cellTotal.Column + 1).value = wsAdd13.Cells(lastRow13 - 1, cellTotal.Column + 1).value
                    
                End If
            Next cellTotal
         
            Else
            
            
           If Not nameExist Then

                        wsAdd13.Activate
                    
                            CopyRange1.Copy
                            
                             wsAdd13.Cells(13, LastColumn + 1).PasteSpecial xlPasteAll
                             wsAdd13.Cells(12, LastColumn + 1) = generalInfo
                             wsAdd13.Cells(14, LastColumn + 1) = name
                             wsAdd13.Cells(16, LastColumn + 1) = Nomenclature
                             wsAdd13.Cells(18, LastColumn + 1) = unit
                             wsAdd13.Cells(20, LastColumn + 1) = price
                                    
                                    
                                'номерація колонок
                                wsAdd13.Cells(22, LastColumn + 1) = lastNumber + 1
                                wsAdd13.Cells(22, LastColumn + 2) = lastNumber + 2
                                wsAdd13.Cells(22, LastColumn + 3) = lastNumber + 3
                   
            End If
        
   
    End If
'    якщо є, то не додавати рядок, а знайти де знаходиться накладна, потім знайти де знаходиться майно і вставити туди кількість, якщо нема - додати колонки з майном


            
            
            Dim quantity As Double
'            Dim quantityL As Double
            If Not nameExist Then
'                        MsgBox "Працюємо"
                        wsAdd13.Activate
            
                                             
                            CopyRange1.Copy
                            
                             wsAdd13.Cells(13, LastColumn + 1).PasteSpecial xlPasteAll
                             wsAdd13.Cells(12, LastColumn + 1) = generalInfo
                             wsAdd13.Cells(14, LastColumn + 1) = name
                             wsAdd13.Cells(16, LastColumn + 1) = Nomenclature
                             wsAdd13.Cells(18, LastColumn + 1) = unit
                             wsAdd13.Cells(20, LastColumn + 1) = price
                                    
                                    
                                'номерація колонок
                                wsAdd13.Cells(22, LastColumn + 1) = lastNumber + 1
                                wsAdd13.Cells(22, LastColumn + 2) = lastNumber + 2
                                wsAdd13.Cells(22, LastColumn + 3) = lastNumber + 3
                   
            End If
            
            'визначаємо кількість, додаємо надходження і вибуття, розраховуючи, що одночасно в одному рядку одне з них 0.
            quantity = filteredData(n, 6) + filteredData(n, 7)
'            quantityL = filteredData(n, 3) + filteredData(n, 19)
            
           
            
            
         'якщо вибуває з підрозділу
         If FromUnit = selectedUnit Then
                               quantity = filteredData(n, 7)
'                        quantityL = filteredData(n, 19)
 
                    If Not nameExist Then
                        If Not DocExist Then
                            wsAdd13.Cells(lastRow13, LastColumn + 2).value = quantity
'                            wsAdd13.Cells(lastRow13, LastColumn + 4).value = quantityL
                            
                            wsAdd13.Cells(lastRow13, LastColumn + 3).value = -quantity
'                            wsAdd13.Cells(lastRow13, LastColumn + 6).value = -quantityL
                        Else
                            wsAdd13.Cells(DocRow, LastColumn + 2).value = quantity + wsAdd13.Cells(DocRow, LastColumn + 2).value
'                            wsAdd13.Cells(DocRow, LastColumn + 4).value = quantityL
                            wsAdd13.Cells(DocRow, LastColumn + 3).value = -quantity
'                            wsAdd13.Cells(DocRow, LastColumn + 6).value = -quantityL
                        End If
                    Else
                    
                        previousQuantyty = wsAdd13.Cells(lastRow13 - 1, foundCell.Column + 2).value
'                        previousQuantytyL = wsAdd13.Cells(lastRow13 - 1, foundCell.Column + 5).value
                        actualQuantyty = previousQuantyty - quantity 'віднімаємо від попередньої кількості
'                        actualQuantytyL = previousQuantytyL - quantityL
                        If Not DocExist Then
                            'вставляємо в поточний рядок кількість
                            wsAdd13.Cells(lastRow13, foundCell.Column + 2).value = actualQuantyty
'                            wsAdd13.Cells(lastRow13, foundCell.Column + 5).value = actualQuantytyL
                            wsAdd13.Cells(lastRow13, foundCell.Column + 1).value = quantity
'                            wsAdd13.Cells(lastRow13, foundCell.Column + 3).value = quantityL
                        Else
                            wsAdd13.Cells(DocRow, foundCell.Column + 2).value = actualQuantyty
'                            wsAdd13.Cells(DocRow, foundCell.Column + 5).value = actualQuantytyL
'                            ============ 3.9.15 ===========
                            'додаємо попередню кількість з клітинки (у випадку якщо однаковий документ, однакове найменування, однаковий підрозділ, але різні МВО)
                            wsAdd13.Cells(DocRow, foundCell.Column + 1).value = quantity + wsAdd13.Cells(DocRow, foundCell.Column + 1).value
'                            ============ 3.9.15 ===========
'                            wsAdd13.Cells(DocRow, foundCell.Column + 3).value = quantityL
                        End If
                       
                    End If
         End If
         
         
         'якщо надходить в підрозділ
         If toUnit = selectedUnit Then
                         If Not nameExist Then
                               If Not DocExist Then
                                     'вставляємо в поточний рядок
                                        wsAdd13.Cells(lastRow13, LastColumn + 1).value = quantity '+ wsAdd13.Cells(lastRow13, LastColumn + 1).value 'надійшло
'                                        wsAdd13.Cells(lastRow13, LastColumn + 2).value = quantityL
                                        wsAdd13.Cells(lastRow13, LastColumn + 3).value = quantity 'становить
'                                        wsAdd13.Cells(lastRow13, LastColumn + 6).value = quantityL
                               Else
                               'якщо нема назви і є документ...
                                        wsAdd13.Cells(DocRow, LastColumn + 1).value = quantity ' надійшло
'                                        wsAdd13.Cells(DocRow, LastColumn + 2).value = quantityL ' надійшло
                                        wsAdd13.Cells(DocRow, LastColumn + 3).value = quantity 'становить
'                                        wsAdd13.Cells(DocRow, LastColumn + 6).value = quantityL
                               End If
                    
                         
                     Else
                        previousQuantyty = wsAdd13.Cells(lastRow13 - 1, foundCell.Column + 2).value
'                        previousQuantytyL = wsAdd13.Cells(lastRow13 - 1, foundCell.Column + 5).value
                        actualQuantyty = previousQuantyty + quantity
'                        actualQuantytyL = previousQuantytyL + quantityL
                           If Not DocExist Then
                              
                                     wsAdd13.Cells(lastRow13, foundCell.Column + 2).value = actualQuantyty
'                                     wsAdd13.Cells(lastRow13, foundCell.Column + 5).value = actualQuantytyL
                                     
                                     wsAdd13.Cells(lastRow13, foundCell.Column).value = quantity
'                                     wsAdd13.Cells(lastRow13, foundCell.Column + 1).value = quantityL
                                 Else
                                      'якщо є назва і є документ...
                                      
                                    wsAdd13.Cells(DocRow, foundCell.Column + 2).value = actualQuantyty
'                                    wsAdd13.Cells(DocRow, foundCell.Column + 5).value = actualQuantytyL
'                                    ============ 3.9.15 ===========
                                    'додаємо попередню кількість з клітинки (у випадку якщо однаковий документ, однакове найменування, однаковий підрозділ, але різні МВО)
                                    wsAdd13.Cells(DocRow, foundCell.Column).value = quantity + wsAdd13.Cells(DocRow, foundCell.Column).value
'                                    ============ 3.9.15 ===========
'                                    wsAdd13.Cells(DocRow, foundCell.Column + 1).value = quantityL
                                                 
                            End If
                     End If
               
         End If
            
        
         'копіюємо стилі загальної іформації
        wsAdd13.Range(Cells(23, 1), Cells(23, 11)).Copy
        wsAdd13.Range(Cells(lastRow13, 1), Cells(lastRow13, 11)).PasteSpecial Paste:=xlPasteFormats
  
        Set foundCell = Nothing
        nameExist = False
        DocExist = False
    Call ProgressBar.Progress(n / UBound(filteredData) * 100)
    
    
    Next n
    
    'знаходимо останню колонку після всіх маніпуляцій
     Dim LastColumnFinal As Long
     LastColumnFinal = wsAdd13.Cells.Find(What:="*", After:=wsAdd13.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
     'копіюємо стилі основної інформації (для показників підрозділів)
       wsAdd13.Range(Cells(23, 6), Cells(23, 8)).Copy
       wsAdd13.Range(Cells(23, 9), Cells(lastRow13, LastColumnFinal)).PasteSpecial Paste:=xlPasteFormats


'скидаємо фільтри
On Error Resume Next
wsAllData.ShowAllData
Exit Sub
'*помилки*******************************************************
WrongDate:
MsgBox "Ймовірно, не правильний формат дати в " & Document & " № " & NumberDocument & " виправте дату і спробуйтие сформувати додаток ще раз!", vbCritical, "Не правильна дата"
wsAllData.ShowAllData 'скидаємо фільтри
Exit Sub

End Sub

