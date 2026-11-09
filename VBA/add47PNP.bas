Attribute VB_Name = "add47PNP"
Sub CreateAdd47PNP()
'для категорій придатне/непридатне/потребує ремонту
Dim wbAllData As Workbook
Dim wsAllData As Worksheet
Dim tbAllData As ListObject
    Set wbAllData = ThisWorkbook
    Set wsAllData = wbAllData.Worksheets("База")
    Set tbAllData = wsAllData.ListObjects("AllData")
    
Dim wsAdd47 As Worksheet
    Set wsAdd47 = wbAllData.Worksheets("Додаток 47 пр-непр")
    
    
Dim RngToSearchData As Range
Set RngToSearchData = wsAdd47.Range("16:16")

Dim ptReduses As PivotTable     'таблиця з переліком МЦ
'оновити таблицю, перед вибором МЦ
Set ptReduses = ThisWorkbook.Worksheets("Довідник д47").PivotTables("ptReduses47")
ptReduses.RefreshTable

'Видаляємо попередні дані
Dim lastColumnStart As Long
 lastColumnStart = wsAdd47.Cells(19, wsAdd47.Columns.count).End(xlToLeft).Column + 1
Dim LastRowStart As Long
 LastRowStart = wsAdd47.Cells.Find(What:="*", After:=wsAdd47.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
                            SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
On Error Resume Next
    wsAdd47.Range(Cells(1, 16), Cells(LastRowStart, lastColumnStart)).EntireColumn.Delete ' Видалити колонки від до останньої

    wsAdd47.Rows("21:" & wsAdd47.Rows.count).Delete

On Error GoTo 0
    
Dim filteredData() As Variant 'масив з відфільтрованими даними

'викликаємо форму для вибору майна
     Dim frm     As UserForm1
        Set frm = UserForms.Add(UserForm1.name)
        frm.ListData = ThisWorkbook.Worksheets("Довідник д47").Range("Goods47")
        
        frm.show
        
        NameOfThings = frm.TextBox1.value
        
        Dim rngWithName As Range
        
        Dim ListGoods As Range
        Set ListGoods = ThisWorkbook.Worksheets("Довідник д47").Range("Goods47")
        'Шукаємо задане значення в діапазоні
        Set rngWithName = ListGoods.Find(NameOfThings, LookIn:=xlValues, LookAt:=xlWhole)
        
        'Перевіряємо, чи знайдено задане значення
        If Not rngWithName Is Nothing Then
            Dim name As String
            Dim Nomenclature As String
            Dim unit As String
            Dim Pice As Double
            
            
            'Визначаємо значення за замовчуванням
            name = rngWithName.Offset(, -12).value
            
'            Nomenclature = rngWithName.Offset(, -10).value
'
'            Unit = rngWithName.Offset(, -11).value
'
'            Pice = rngWithName.Offset(, -9).value
            
'             On Error GoTo WrongNomenclature
            Nomenclature = rngWithName.Offset(, -10).value
            If Nomenclature = "(пусто)" Then
            MsgBox "Ймовірно, ви не вказали код номенклатури вибраного майна, вкажіть код номенклатури (якщо не відомо, поставте б/н) і спробуйте створити додаток ще раз!", vbCritical, "Код номенклатури не вказаний"
            Exit Sub
            End If
'            On Error GoTo 0
            
'            On Error GoTo WrongUnit
            unit = rngWithName.Offset(, -11).value
            If unit = "(пусто)" Then
            MsgBox "Ймовірно, ви не вказали одиницю виміру вибраного майна, вкажіть одиницю виміру і спробуйте створити додаток ще раз!", vbCritical, "Одиниця виміру не вказана"
            Exit Sub
            End If
'            On Error GoTo 0
            
            On Error GoTo WrongPrice
            Pice = rngWithName.Offset(, -9).value
            On Error GoTo 0
            
            
           'заповнюємо критерії, по яких будемо фільтрувати базу даних для витяжки потрібного майна
            Dim criteriaRange47 As Range
            
            Set criteriaRange47 = wsAdd47.Range("criteriaRange47PNP")
            
            With criteriaRange47
                .Cells(2, 1) = "'=" & name 'найменуванян
                .Cells(2, 2) = "'=" & unit 'одиниця виміру
                .Cells(2, 3) = "'=" & Nomenclature    'код номенклатури
                .Cells(2, 4) = Pice  'ціна
            End With
            
            'знімаємо фільтри,якщо були поставлені
            If wsAllData.FilterMode Then
                wsAllData.ShowAllData
            End If
            
            'фільтруємо по заданих критеріях
            tbAllData.Range.AdvancedFilter Action:=xlFilterInPlace, criteriaRange:=criteriaRange47, Unique:=False
            
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
        
End If
Dim lastRow47 As Long
Dim CopyRange1 As Range
Dim foundCell As Range

Dim n As Long
    On Error GoTo DontSelect
    For n = LBound(filteredData) + 1 To UBound(filteredData)
    On Error GoTo 0
    lastRow47 = Cells(wsAdd47.Rows.count, "A").End(xlUp).Row + 1
    Dim LastColumn As Long
    LastColumn = wsAdd47.Cells.Find(What:="*", After:=wsAdd47.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    
    'скопіювати останій рядок і вставити
    wsAdd47.Range(wsAdd47.Cells(lastRow47 - 1, 12), wsAdd47.Cells(lastRow47 - 1, LastColumn)).Copy
    wsAdd47.Range(wsAdd47.Cells(lastRow47, 12), wsAdd47.Cells(lastRow47, LastColumn)).PasteSpecial xlPasteAll
'    wsAdd47.Cells(lastRow47, LastColumn)
    
    Dim CategoryFrom As String
    Dim CategoryTo As String
    Dim categoryCol As Long
    categoryCol = 0
    
    CategoryFrom = filteredData(n, 5)
    CategoryTo = filteredData(n, 27)
 
    
           Dim typeOperation As String
           typeOperation = filteredData(n, 26)
           
            Dim quantity As Double
            Dim FromUnit As String
            Dim toUnit As String
            Dim fromUnitExist As Boolean
            Dim toUnitExist As Boolean
            
            Dim previousQuantyty As Double
            Dim actualQuantyty As Double
            
            Set CopyRange1 = wsAdd47.Range(wsAdd47.Cells(16, 12), wsAdd47.Cells(20, 15))
            
            Select Case typeOperation
    '*********************************************************************************************************
               Case "надходження"
                        quantity = filteredData(n, 6) 'кількість
                        toUnit = filteredData(n, 9) 'куди (підрозділ)
                        wsAdd47.Cells(lastRow47, 6).value = quantity
                        toUnitExist = FindValue(toUnit, RngToSearchData)
                        
                    If Not toUnitExist Then
'                        MsgBox "Працюємо"
                        wsAdd47.Activate

                        'копіюємо стовпчики
                       
                            CopyRange1.Copy
                            
                             wsAdd47.Cells(16, LastColumn + 1).PasteSpecial xlPasteAll
                            
                             wsAdd47.Cells(16, LastColumn + 1) = toUnit
                             
                    'якщо тут помилка - в базі не правильно вказана категорія:
                    'повинно бути пр. або не пр. або пот. рем.
                    'переконайтеся що у вас правильно проставлені категорії відправника або одержувача
                    categoryCol = wsAdd47.Cells.Find(What:=CategoryTo, After:=wsAdd47.Cells(18, LastColumn), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext).Column
                    
                    previousQuantyty = wsAdd47.Cells(lastRow47, categoryCol).value
'                    previousQuantyty = wsAdd47.Cells(lastRow47, categoryCol).End(xlUp).value 'визначаємо попередню кількість, за відповідною категорією.
                    actualQuantyty = previousQuantyty + quantity 'додаємо до попередньої кількості
                    'вставляємо в поточний рядок
                    wsAdd47.Cells(lastRow47, categoryCol).value = actualQuantyty
                   
                         Else
                        

                            'Шукаємо задане значення в діапазоні
                            Set foundCell = RngToSearchData.Find(toUnit, LookIn:=xlValues, LookAt:=xlWhole)
                        
                            'Перевіряємо, чи знайдено задане значення
                            If Not foundCell Is Nothing Then
                    
                                'якщо тут помилка - в базі не правильно вказана категорія:
                                'повинно бути пр. або не пр. або пот. рем.
                                'переконайтеся що у вас правильно проставлені категорії відправника або одержувача
                                categoryCol = wsAdd47.Cells.Find(What:=CategoryTo, After:=wsAdd47.Cells(18, foundCell.Column), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext).Column
                                previousQuantyty = wsAdd47.Cells(lastRow47, categoryCol).value
'                                previousQuantyty = wsAdd47.Cells(lastRow47, categoryCol).End(xlUp).value 'визначаємо попередню кількість, за відповідною категорією.
                                actualQuantyty = previousQuantyty + quantity 'додаємо до попередньої кількості
                                'вставляємо в поточний рядок
                                wsAdd47.Cells(lastRow47, categoryCol).value = actualQuantyty
                               
                            Else
'                                FindValue = False
                            End If
 
                    End If
'******************************************************************************************************************************
                        
               Case "вибуття"
                        quantity = filteredData(n, 7)
                        wsAdd47.Cells(lastRow47, 7).value = quantity
                        FromUnit = filteredData(n, 8)
                        fromUnitExist = FindValue(FromUnit, RngToSearchData)
                        
                        
                    If Not fromUnitExist Then
'                        MsgBox "Працюємо"
                        wsAdd47.Activate
 
                    
                            CopyRange1.Copy
                            wsAdd47.Cells(16, LastColumn + 1).PasteSpecial xlPasteAll
                            
                             wsAdd47.Cells(16, LastColumn + 1) = FromUnit
                    'якщо тут помилка - в базі не правильно вказана категорія:
                    'повинно бути пр. або не пр. або пот. рем.
                    'переконайтеся що у вас правильно проставлені категорії відправника або одержувача
                    categoryCol = wsAdd47.Cells.Find(What:=CategoryFrom, After:=wsAdd47.Cells(18, LastColumn), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext).Column
                    previousQuantyty = wsAdd47.Cells(lastRow47, categoryCol).value
'                    previousQuantyty = wsAdd47.Cells(lastRow47, categoryCol).End(xlUp).value  'визначаємо попередню кількість, за відповідною категорією.
                    actualQuantyty = previousQuantyty - quantity 'віднімаємо до попередньої кількості
                    'вставляємо в поточний рядок
                    wsAdd47.Cells(lastRow47, categoryCol).value = actualQuantyty


                         Else

                            'Шукаємо задане значення в діапазоні, якщо вже є підрозділ
                            Set foundCell = RngToSearchData.Find(FromUnit, LookIn:=xlValues, LookAt:=xlWhole)
                        
                            'Перевіряємо, чи знайдено задане значення
                            If Not foundCell Is Nothing Then
                                'якщо тут помилка - в базі не правильно вказана категорія:
                                'повинно бути пр. або не пр. або пот. рем.
                                'переконайтеся що у вас правильно проставлені категорії відправника або одержувача
                                categoryCol = wsAdd47.Cells.Find(What:=CategoryFrom, After:=wsAdd47.Cells(18, foundCell.Column), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext).Column
                                previousQuantyty = wsAdd47.Cells(lastRow47, categoryCol).value
'                                previousQuantyty = wsAdd47.Cells(lastRow47, categoryCol).End(xlUp).value  'визначаємо попередню кількість, за відповідною категорією.
                                actualQuantyty = previousQuantyty - quantity 'віднімаємо від попередньої кількості
                                
                                'вставляємо в поточний рядок кількість
                                wsAdd47.Cells(lastRow47, categoryCol).value = actualQuantyty
                               
                            Else
'                                FindValue = False
                            End If
                    End If
                    
                  
                        '***************************************************************************************************************
     
               Case "внутрішнє переміщення"
                        quantity = filteredData(n, 7)
                        toUnit = filteredData(n, 9)
                        FromUnit = filteredData(n, 8)
                        
                        fromUnitExist = FindValue(FromUnit, RngToSearchData)
                        toUnitExist = FindValue(toUnit, RngToSearchData)
                        
                    'для підрозділу, з якого вибуває
                        If Not fromUnitExist Then
'                        MsgBox "Працюємо"
                        wsAdd47.Activate
                        

  
                        
                            CopyRange1.Copy
                            wsAdd47.Cells(16, LastColumn + 1).PasteSpecial xlPasteAll
                            
                             wsAdd47.Cells(16, LastColumn + 1) = FromUnit
                    'якщо тут помилка - в базі не правильно вказана категорія:
                    'повинно бути пр. або не пр. або пот. рем.
                    'переконайтеся що у вас правильно проставлені категорії відправника або одержувача
                    categoryCol = wsAdd47.Cells.Find(What:=CategoryFrom, After:=wsAdd47.Cells(18, LastColumn), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext).Column
                    previousQuantyty = wsAdd47.Cells(lastRow47, categoryCol).value
'                    previousQuantyty = wsAdd47.Cells(lastRow47, categoryCol).End(xlUp).value  'визначаємо попередню кількість, за відповідною категорією.
                    actualQuantyty = previousQuantyty - quantity 'віднімаємо до попередньої кількості
                    'вставляємо в поточний рядок
                    wsAdd47.Cells(lastRow47, categoryCol).value = actualQuantyty
                    

                         Else
                         
'                        Dim foundCell As Range

                            'Шукаємо задане значення в діапазоні
                            Set foundCell = RngToSearchData.Find(FromUnit, LookIn:=xlValues, LookAt:=xlWhole)
                        
                            'Перевіряємо, чи знайдено задане значення
                            If Not foundCell Is Nothing Then
                                'якщо тут помилка - в базі не правильно вказана категорія:
                                'повинно бути пр. або не пр. або пот. рем.
                                'переконайтеся що у вас правильно проставлені категорії відправника або одержувача
                                categoryCol = wsAdd47.Cells.Find(What:=CategoryFrom, After:=wsAdd47.Cells(18, foundCell.Column), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext).Column
                                previousQuantyty = wsAdd47.Cells(lastRow47, categoryCol).value
'                                previousQuantyty = wsAdd47.Cells(lastRow47, categoryCol).End(xlUp).value  'визначаємо попередню кількість, за відповідною категорією.
                                actualQuantyty = previousQuantyty - quantity 'додаємо до попередньої кількості
                                'вставляємо в поточний рядок
                                wsAdd47.Cells(lastRow47, categoryCol).value = actualQuantyty
                               
                            Else
'                                FindValue = False
                            End If
                        
                    End If
                    
                    
                       If Not toUnitExist Then
'                        MsgBox "Працюємо"
                        wsAdd47.Activate
                        
                        
                        Dim addCol As Long
                        addCol = 0
                        
                        If Not fromUnitExist Then
                        
                        addCol = 4
                        End If
 
                            CopyRange1.Copy
                            wsAdd47.Cells(16, LastColumn + addCol + 1).PasteSpecial xlPasteAll
                            
                             wsAdd47.Cells(16, LastColumn + addCol + 1) = toUnit
                     'якщо тут помилка - в базі не правильно вказана категорія:
                    'повинно бути пр. або не пр. або пот. рем.
                    'переконайтеся що у вас правильно проставлені категорії відправника або одержувача
                    categoryCol = wsAdd47.Cells.Find(What:=CategoryTo, After:=wsAdd47.Cells(18, LastColumn + addCol - 1), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext).Column
                    '****** ще треба скопіювати весь попередній рядок...., щоб можна було віднімати (додавати) поточну кількість до попередньої
                    previousQuantyty = wsAdd47.Cells(lastRow47, categoryCol).value
'                    previousQuantyty = wsAdd47.Cells(lastRow47, categoryCol).End(xlUp).value  'визначаємо попередню кількість, за відповідною категорією.
                    actualQuantyty = previousQuantyty + quantity 'додаємо до попередньої кількості
                    'вставляємо в поточний рядок
                    wsAdd47.Cells(lastRow47, categoryCol).value = actualQuantyty
                   
                         Else
'                        Dim foundCell As Range

                            'Шукаємо задане значення в діапазоні
                            Set foundCell = RngToSearchData.Find(toUnit, LookIn:=xlValues, LookAt:=xlWhole)
                        
                            'Перевіряємо, чи знайдено задане значення
                            If Not foundCell Is Nothing Then
                                'якщо тут помилка - в базі не правильно вказана категорія:
                                'повинно бути пр. або не пр. або пот. рем.
                                'переконайтеся що у вас правильно проставлені категорії відправника або одержувача
                                categoryCol = wsAdd47.Cells.Find(What:=CategoryTo, After:=wsAdd47.Cells(18, foundCell.Column), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext).Column
                                previousQuantyty = wsAdd47.Cells(lastRow47, categoryCol).value
'                                previousQuantyty = wsAdd47.Cells(lastRow47, categoryCol).End(xlUp).value  'визначаємо попередню кількість, за відповідною категорією.
                                actualQuantyty = previousQuantyty + quantity 'додаємо до попередньої кількості
                                'вставляємо в поточний рядок
                                wsAdd47.Cells(lastRow47, categoryCol).value = actualQuantyty
                               
                            Else
'                                FindValue = False
                            End If
    
                    End If
'
'
            End Select
            
            With wsAdd47
                .Cells(lastRow47, 1).value = filteredData(n, 1) 'дата внесення
                .Cells(lastRow47, 2).value = filteredData(n, 25) 'найменування документа
                .Cells(lastRow47, 3).NumberFormat = "@"
                .Cells(lastRow47, 3).value = filteredData(n, 14) 'номер документа
                .Cells(lastRow47, 4).value = filteredData(n, 13)  'дата документа
                .Cells(lastRow47, 5).value = filteredData(n, 8) & Chr(10) & "(" & filteredData(n, 9) & ")"  'підрозділ звідки (підрозділ куди)
        
            End With
 
         'копіюємо стилі загальної іформації
        wsAdd47.Range(Cells(20, 1), Cells(20, 19)).Copy
        wsAdd47.Range(Cells(lastRow47, 1), Cells(lastRow47, 19)).PasteSpecial Paste:=xlPasteFormats
     
     
     Set foundCell = Nothing
        
        Call ProgressBar.Progress(n / UBound(filteredData) * 100)
    
    Next n
    
    'знаходимо останню колонку після всіх маніпуляцій
     Dim LastColumnFinal As Long
     LastColumnFinal = wsAdd47.Cells.Find(What:="*", After:=wsAdd47.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
     'копіюємо стилі основної інформації (для показників підрозділів)
       wsAdd47.Range(Cells(20, 12), Cells(20, 15)).Copy
       wsAdd47.Range(Cells(21, 12), Cells(lastRow47, LastColumnFinal)).PasteSpecial Paste:=xlPasteFormats
 
       'проставляємо суму
       Dim cellAll As Range
       Dim cellWithSum As Range
       For Each cellAll In wsAdd47.Rows(17).Cells 'рядок з усього
        If cellAll.value = "Усього" Then
            Set cellWithSum = cellAll.Offset(2, 0)
            cellWithSum.AutoFill destination:=wsAdd47.Range(wsAdd47.Cells(cellWithSum.Row, cellWithSum.Column), wsAdd47.Cells(lastRow47, cellWithSum.Column))
        End If
    Next cellAll
    
  
    
'   'формула по категоріях
'    For j = 9 To 11
'       For i = 21 To lastRow47
'           wsAdd47.Range(Cells(i, j), Cells(lastRow47, j)).Formula = "=SUMIFS(" & wsAdd47.Range(Cells(i, 12), Cells(i, LastColumnFinal)).Address & "," & wsAdd47.Range(Cells(18, 12), Cells(18, LastColumnFinal)).Address & "," & wsAdd47.Cells(18, j).Address & ")"
'
'       Next i
'    Next j
    
    Dim result As Double
    For j = 9 To 11
        For i = 21 To lastRow47
            
            Set sumRange = wsAdd47.Range(wsAdd47.Cells(i, 12), wsAdd47.Cells(i, LastColumnFinal))
            Set criteriaRange = wsAdd47.Range(wsAdd47.Cells(18, 12), wsAdd47.Cells(18, LastColumnFinal))
            criteriaValue = wsAdd47.Cells(18, j).value
        
'            wsAdd47.Range(Cells(i, j), Cells(lastRow47, j)).Formula = "=SUMIFS(" & wsAdd47.Range(Cells(i, 14), Cells(i, LastColumnFinal)).Address & "," & wsAdd47.Range(Cells(18, 14), Cells(18, LastColumnFinal)).Address & "," & wsAdd47.Cells(18, j).Address & ")"
        
        result = Application.WorksheetFunction.SumIfs(sumRange, criteriaRange, criteriaValue)
        wsAdd47.Cells(i, j).value = result
        Next i
    Next j
  
 'скидаємо область друку
wsAdd47.PageSetup.PrintArea = ""
'скидаємо фільтри
            If wsAllData.FilterMode Then
            On Error Resume Next
                wsAllData.ShowAllData
            End If
            
            Exit Sub
'***помилки****************************************
DontSelect:
MsgBox "Ймовірно,майно відсутнє в базі, або Ви не підтвердили вибір майна: вибране майно повинне ""перестрибнути"" з нижнього у верхнє поле вікна вибору майна, спробуйтие сформувати додаток ще раз!", vbCritical, "Майно не знайдено"
'wsAllData.ShowAllData
Exit Sub

WrongPrice:
MsgBox "Ймовірно,ви не вказали ціну вибраного майна, вкажіть ціну (якщо не відома, поставте 1 грн.) і спробуйте створити додаток ще раз!", vbCritical, "Ціна не вказана"
Exit Sub


End Sub


Sub SaveAdd47PNP()

        ActiveSheet.Select 'копіюємо
        ActiveSheet.Copy 'зберігаємо як нову книгу
     
     Dim WbNewWO As Workbook
     
        Set WbNewWO = ActiveWorkbook
    
    Dim nameAdd47 As String
    Dim nomenclature47 As String
    Dim unit47 As String
    Dim price47 As Double
    Dim formattedPrice As String
    
    nameAdd47 = WbNewWO.Worksheets(1).Range("nameAdd47PNP")
    
    nomenclature47 = WbNewWO.Worksheets(1).Range("nomenclature47PNP")
    unit47 = WbNewWO.Worksheets(1).Range("unit47PNP")
    price47 = WbNewWO.Worksheets(1).Range("price47PNP")
    
    formattedPrice = Format(price47, "#,##0.00")

    Dim shp As Shape
    For Each shp In WbNewWO.Worksheets(1).Shapes
        shp.Delete
    Next shp
        
      Dim fileName As String, fileName2 As String, lineBrake As String
        'Шлях до файла зберігання
        SavePath = WbNewWO.Worksheets(1).Range("PathToSaveAdd47PNP")
        Dim DateAdd47 As Date
        DateAdd47 = Format(Date, "dd.mm.yyyy")
        
        
        
              On Error Resume Next
         If Dir(SavePath, vbDirectory) = vbNullString Then
         On Error GoTo 0
        With Application.FileDialog(msoFileDialogFolderPicker)
            .Title = "Виберіть папку для збереження"
            .AllowMultiSelect = False
            If .show = -1 Then
                SavePath = .SelectedItems(1)
                'Записуємо новий шлях у клітинку
                ThisWorkbook.Sheets("Додаток 47 пр-непр").Range("PathToSaveAdd47PNP").value = SavePath
            Else
                MsgBox "Збереження скасовано", vbExclamation
                WbNew.Close SaveChanges:=False
                Application.DisplayAlerts = True
                Exit Sub
            End If
        End With
    End If
        
        
        
        
        'складаємо назву файла
        fileName = SavePath & Application.PathSeparator & CleanFileName(DateAdd47 & " д47 " & nameAdd47 & " " & nomenclature47 & " " & unit47 & " " & formattedPrice & ".xlsx")
        'заміняємо перенесення рядка, косі лінії
'        fileName2 = CleanFileName(fileName)
        
        'зберігаємо файл
        On Error GoTo WrongPath
         WbNewWO.SaveAs fileName:=fileName, FileFormat:=xlOpenXMLWorkbook
         On Error GoTo 0
     Exit Sub
WrongPath:
     MsgBox "Ймовірно, не правильно вказана адреса збереження файла, виправте адресу збереження файлу( на зразок ""D:\База обліку майна ЗСУ 2024"" *без косої лінії вкінці) і спробуйтие сформувати додаток ще раз!", vbCritical, "Не правильний шлях збереження"
     Exit Sub
     
End Sub




