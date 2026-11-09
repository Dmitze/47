Attribute VB_Name = "ManualAdd"
Sub ManualAdding()

Dim AddManual

AddManual = MsgBox("Внести дані вручну?", vbQuestion + vbOKCancel, "Внести вручну?")
 
If AddManual = 2 Then
 Exit Sub
End If

Dim wbAllData As Workbook
Dim wsAllData As Worksheet
Dim tbAllData As ListObject
    Set wbAllData = ThisWorkbook
    Set wsAllData = wbAllData.Worksheets("База")
    Set tbAllData = wsAllData.ListObjects("AllData")
    
'Dim wsAccountingBook As Worksheet
'Dim tbAccountingBook As ListObject
'
'    Set wsAccountingBook = wbAllData.Worksheets("Книга обліку д13")
'    Set tbAccountingBook = wsAccountingBook.ListObjects("AccountingBook")
'
''для перевірки чи є в книзі обліку мц, яке розносимо в БД
'Dim RngToSearchData As Range
'Dim DataThing As String

            Dim Executor As String  '+
            Dim FromUnit As String '+
            Dim fromMRP As String '+
            Dim toMRP As String '+
            Dim toUnit As String '+
            Dim NumberInv As String '+
'            Dim NumberInvClear As String
            Dim DateInv As Date '+
            Dim NumberDemand As String '+
             
            Dim NameThings As String '+
            Dim MensureThings As String '+
            Dim QuantityThings As Double '+
            Dim СategoryThings As String '+
            Dim СategoryThingsTo As String
            Dim ItemNumber As String '+
            Dim Note As String '+
            Dim price As Double '+
            
            Dim Field10 '+
            Dim Field11 '+
            Dim Field12 '+
'            Dim TimeInv
            
            Dim Service As String '+
            Dim Content As String '+
            
            Dim InputDate As Date
            Dim typeDoc As String '+


InputDate = Date




'перевірка, чи є такий документ в базі
NumberInv = InputBox("Введіть номер документа:", "№ документа") '+

Dim foundCell As Range
Dim firstColumn As Range
        
            Set firstColumn = tbAllData.ListColumns(14).DataBodyRange
            
            Set foundCell = firstColumn.Find(What:=NumberInv, LookIn:=xlValues, LookAt:=xlWhole)
    
            If Not foundCell Is Nothing Then
            Dim EnterAnyway
                EnterAnyway = MsgBox("Документ з номером " & NumberInv & " вже є в базі." & vbNewLine & _
                "Все одно додати в базу?", vbYesNo, "Документ вже є в базі")
                If EnterAnyway = 7 Then
                    Exit Sub
                End If
            
            End If

typeDoc = InputBox("Введіть назву(акт, накладна і т.д.) документа:", "Тип документа", "Акт") '+

DateInv = InputBox("Введіть дату документа " & typeDoc & " №" & NumberInv & " в форматі 27.04.2023 :", "Дата документа") '+


FromUnit = InputBox("Підрозділ (організація, база, в/ч), з якого надійшло (який передає) майно", "Постачальник")
toUnit = InputBox("Підрозділ (організація, база, в/ч), отримувач майна", "Отримувач")

NumberDemand = InputBox("Назва, номер, і дата документа підстави операції (наряд, рапорт, заявка і т.д.)", "Підстава операцї")

Service = InputBox("До номенклатури якої служби забезпечення відноситься майно (зв'язок, медичне, прод, КЕС і т.д.)?", "Служба забезпечення")
Content = InputBox("Виберіть тип операції:" & vbNewLine & _
                    "надходження - 1" & vbNewLine & _
                    "вибуття - 2" & vbNewLine & _
                    "внутрішнє переміщення - 3", "Тип операції", 1)
'
        If Content = 2 Then
            Content = "вибуття"
        ElseIf Content = 1 Then
            Content = "надходження"
        Else
            Content = "внутрішнє переміщення"
        End If

Executor = InputBox("Прізвище, хто здійснює занесення документа в базу", "Виконавець")
toMRP = InputBox("Прізвище мателіально відповідальної особи, яка отримує", "МВО отримувача")
fromMRP = InputBox("Прізвище мателіально відповідальної особи, яка передає", "МВО відправника")

NextPosition:

NameThings = InputBox("Введіть найменування майна", "Найменування")

ItemNumber = InputBox("Код номенклатури (за наявності) " & NameThings, "Код", "б/н")

MensureThings = InputBox("Одиниця виміру: " & NameThings, "Одиниця виміру", "к-т")

QuantityThings = InputBox("Кількість: " & NameThings, "Кількість", 1)

СategoryThings = InputBox("Категорія відправника: " & NameThings, "Категорія", "прид.")

СategoryThingsTo = InputBox("Категорія одержувача: " & NameThings, "Категорія", "прид.")

Dim txtPrice As String

txtPrice = InputBox("Ціна за одиницю: " & NameThings, "Ціна")


If txtPrice <> "" Then
    txtPrice = Replace(txtPrice, ".", Application.DecimalSeparator)
    txtPrice = Replace(txtPrice, ",", Application.DecimalSeparator)

    If IsNumeric(txtPrice) Then
        price = CDbl(txtPrice)
    Else
        MsgBox "Некоректне число"
    End If
End If


Note = InputBox("Примітка: " & NameThings, "Примітка")

Field10 = InputBox("Поле 10 (додаткова інформація, яку ви бажаєте додати про майно) для " & NameThings, "Поле 10")
Field11 = InputBox("Поле 11 (додаткова інформація, яку ви бажаєте додати про майно) для " & NameThings, "Поле 11")
Field12 = InputBox("Поле 12 (додаткова інформація, яку ви бажаєте додати про майно) для " & NameThings, "Поле 12")


wsAllData.Activate 'активуємо лист з базою

            Dim ArrInv()
            Dim ArrSizeX
            Dim ArrSizeY
                ArrSizeX = tbAllData.ListColumns.count
                ArrSizeY = 1
                
                            
                            ReDim ArrInv(1 To ArrSizeY, 1 To ArrSizeX)
                            arrRow = 1
                           
                           ArrInv(arrRow, 1) = InputDate
                           ArrInv(arrRow, 2) = NameThings
                           ArrInv(arrRow, 4) = MensureThings
                           ArrInv(arrRow, 5) = СategoryThings
                           
                           If Content = "вибуття" Or Content = "внутрішнє переміщення" Then
                                ArrInv(arrRow, 7) = QuantityThings
                           Else
                                ArrInv(arrRow, 6) = QuantityThings
                           End If
                                                                                 
                           ArrInv(arrRow, 8) = FromUnit
                           ArrInv(arrRow, 9) = toUnit
                           ArrInv(arrRow, 10) = fromMRP
                           ArrInv(arrRow, 11) = toMRP
                           ArrInv(arrRow, 12) = NumberDemand
                           ArrInv(arrRow, 13) = DateInv
                           ArrInv(arrRow, 14) = NumberInv
                           ArrInv(arrRow, 15) = NumberInvClear
                           ArrInv(arrRow, 16) = Note
                           
                           ArrInv(arrRow, 17) = Executor
                           ArrInv(arrRow, 18) = ItemNumber
                           
'                           ArrInv(arrRow, 19) = TimeInv
                           ArrInv(arrRow, 20) = price
                           ArrInv(arrRow, 21) = Service
                                                      
                           ArrInv(arrRow, 22) = Field10
                           ArrInv(arrRow, 23) = Field11
                           ArrInv(arrRow, 24) = Field12
                           ArrInv(arrRow, 25) = typeDoc
                           ArrInv(arrRow, 26) = Content
                           ArrInv(arrRow, 27) = СategoryThingsTo


Dim lastRow As Long
    lastRow = wsAllData.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row + 1
Dim arrTarget
wsAllData.Activate
    Set arrTarget = wsAllData.Range(Cells(lastRow, 1), Cells(lastRow + UBound(ArrInv()) - 1, ArrSizeX))
    arrTarget.value = ArrInv
    

'  'Якщо нема мц в книзі обліку - додаємо колонки з такими назвами
'
'            DataThing = NameThings & ItemNumber & MensureThings & Price
'            Set RngToSearchData = wsAccountingBook.Range("16:16")
'            ThingExist = FindValue(DataThing, RngToSearchData)
'
'                    If Not ThingExist Then
''                        MsgBox "Працюємо"
'                        wsAccountingBook.Activate
'                        Dim LastColumn As Long
'                        LastColumn = tbAccountingBook.Range.Columns.count
'
'                        Dim CopyRange1 As Range
'                        Dim CopyRange2 As Range
'
'                        'копіюємо стовпчики
'                        Set CopyRange1 = wsAccountingBook.Range(Cells(15, LastColumn - 2), Cells(16, LastColumn))
'
'                            CopyRange1.Copy
'                            wsAccountingBook.Cells(15, LastColumn + 1).PasteSpecial xlPasteAll
'
'                        Set CopyRange2 = wsAccountingBook.Range(Cells(7, LastColumn - 2), Cells(14, LastColumn))
'                            CopyRange2.Copy
'                            wsAccountingBook.Cells(7, LastColumn + 1).PasteSpecial xlPasteFormats
'
'                        tbAccountingBook.ListColumns(LastColumn - 2).DataBodyRange.Copy
'                        wsAccountingBook.Cells(18, LastColumn + 1).PasteSpecial xlPasteAll
'
'                        tbAccountingBook.ListColumns(LastColumn - 1).DataBodyRange.Copy
'                        wsAccountingBook.Cells(18, LastColumn + 2).PasteSpecial xlPasteAll
'
'                        tbAccountingBook.ListColumns(LastColumn).DataBodyRange.Copy
'                        wsAccountingBook.Cells(18, LastColumn + 3).PasteSpecial xlPasteAll
'
'                        'додаємо позицію в книгу
'                        wsAccountingBook.Cells(8, LastColumn + 1) = NameThings
'                        wsAccountingBook.Cells(10, LastColumn + 1) = ItemNumber
'                        wsAccountingBook.Cells(12, LastColumn + 1) = MensureThings
'                        wsAccountingBook.Cells(14, LastColumn + 1) = Price
'
'
'                    End If
'


Dim WhatIsNext

WhatIsNext = MsgBox("Позиція згідно " & vbNewLine & typeDoc & " №" & NumberInv & " від " & DateInv & _
vbNewLine & NameThings & vbNewLine & _
"номенклатурний номер: " & ItemNumber & vbNewLine & "Одиниця виміру: " & MensureThings & vbNewLine & "Кількість:" & _
QuantityThings & vbNewLine & "Ціна: " & price & vbNewLine & "Примітка" & Note _
& vbNewLine & "рознесена. Перейти до наступної позиції?", vbYesNo, "Рознести наступну позицію?")

If WhatIsNext = 6 Then
    GoTo NextPosition
End If
        
 
    
'
'            'додаємо накладну в книгу обліку
'        Set rowAccBook = tbAccountingBook.ListRows.Add
'
'        rowAccBook.Range(1) = InputDate
'        rowAccBook.Range(2) = typeDoc
'        rowAccBook.Range(3) = NumberInv
'        rowAccBook.Range(4) = DateInv
'
'        If Content = "вибуття" Or Content = "внутрішнє переміщення" Then
'             rowAccBook.Range(5) = toUnit
'        Else
'             rowAccBook.Range(5) = fromUnit
'        End If



MsgBox "Документ рознесений", , NumberInv

            
End Sub




