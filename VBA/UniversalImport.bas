Attribute VB_Name = "UniversalImport"
Sub UnivImport()


Dim wbAllData As Workbook
Dim wsAllData As Worksheet
Dim tbAllData As ListObject
    Set wbAllData = ThisWorkbook
    Set wsAllData = wbAllData.Worksheets("База")
    Set tbAllData = wsAllData.ListObjects("AllData")
    
Dim lastRow As Long
Dim arrTarget
    
    
'Dim wsAccountingBook As Worksheet
'Dim tbAccountingBook As ListObject
'
'    Set wsAccountingBook = wbAllData.Worksheets("Книга обліку д13")
'    Set tbAccountingBook = wsAccountingBook.ListObjects("AccountingBook")
    
'для перевірки чи є в книзі обліку мц, яке розносимо в БД
Dim RngToSearchData As Range
Dim DataThing As String

wsAllData.Activate 'активуємо лист з базою
 '====3.9.18====
On Error Resume Next
wsAllData.ShowAllData
On Error GoTo 0
'====3.9.18====


Application.AskToUpdateLinks = False
Application.DisplayAlerts = False

Dim sFolder As String, sFiles As String

With Application.FileDialog(msoFileDialogFolderPicker)
           If .show = False Then Exit Sub
           sFolder = .SelectedItems(1)

End With

    sFolder = sFolder & IIf(Right(sFolder, 1) = Application.PathSeparator, "", Application.PathSeparator)
    sFiles = Dir(sFolder & "*.xls*")
    
    Do While sFiles <> "" 'цикл по файлах
    'Відкриваємо книгу
    Workbooks.Open sFolder & sFiles
    
    'Відкриваємо книгу
Dim wbActualDoc As Workbook
        Set wbActualDoc = ActiveWorkbook
Dim wsActualDoc As Worksheet
        Set wsActualDoc = wbActualDoc.Worksheets(1)

Dim Executor    As String
Dim FromUnit    As String
Dim fromMRP As String
Dim toMRP   As String
Dim toUnit  As String
Dim Number  As String
Dim NumberClear As String
Dim DateDoc As Date
Dim NumberBasis As String
Dim NameThings  As String
Dim MensureThings   As String
Dim QuantityThings  As Double
Dim СategoryThings  As String
Dim СategoryThingsTo    As String
Dim ItemNumber  As String
Dim Note    As String
Dim Note1   As String
Dim Note2   As String
Dim Note3   As String
Dim TimeInv As Long
Dim price   As Double
Dim Service As String
Dim Content As String
Dim InputDate   As Date
Dim typeDoc As String
   
Dim RangeToSearch As Range
    Set RangeToSearch = tbAllData.ListColumns(14).DataBodyRange
Dim DocExist
Dim msgExist

Dim ArrDoc()
Dim ArrSizeX As Long
Dim ArrSizeY As Long

Dim LastColumn As Long

Dim CopyRange1 As Range
Dim CopyRange2 As Range

Dim structureDoc As String

Dim tbActual As ListObject
        
    structureDoc = wbActualDoc.Worksheets(1).Range("Z1")
                
   Select Case structureDoc
 '***********************************************************************************************
       Case "WriteOff431"
       
       Number = Trim(wsActualDoc.Range("numberWOff").value)
        'перевірка, чи існує вже такий акт
                    On Error Resume Next
                                        
                        DocExist = FindValue(Number, RangeToSearch)
                        
                        If DocExist Then
                            
                            msgExist = MsgBox("Документ: " & Number & " вже рознесений в таблицю. Все одно рознести?", vbYesNo, "Накладна існує")
                                
                                If msgExist = 7 Then
                                    GoTo nextDoc
                                End If
                                
                        End If
                     On Error GoTo 0
                     
                NumberClear = Trim(wsActualDoc.Range("clearNumberWOff"))
                    
                Service = Trim(wsActualDoc.Range("Service"))

                typeDoc = Trim(wsActualDoc.Range("typeDoc"))
                
                
                 
                InputDate = Date 'дата введення
                Executor = wsActualDoc.Range("executor") ' виконавець (якщо тут помилка, скоріше всього клітинці
                'не присвоєне ім'я Exsecutor. Скопіювати Exsecutor, вставити в назву клітинки з виконавцем, натиснути Enter
'                FromUnit = Trim(wsActualDoc.Range("FromUnit")) 'відправник
                
                toUnit = Trim(wsActualDoc.Range("WriteOff")) 'одержувач
                
                Content = wsActualDoc.Range("rngTypeOperation")
                '******************************************************************
                ContentAdd = Trim(wsActualDoc.Range("ContentInvAdd431_WO"))
                
                
                On Error Resume Next
                fromMRP = Trim(wsActualDoc.Range("MRP"))
                On Error GoTo 0
                toMRP = Trim(wsActualDoc.Range("WriteOff"))
                

                FromUnit = wsActualDoc.Range("fromUnit")
                DateDoc = wsActualDoc.Range("DateWOff") 'дата акта списання
'                TimeInv = wsActualDoc.Range("TimeInv")

'
            Dim tbWriteOff As ListObject
            
            Set tbWriteOff = wsActualDoc.ListObjects("tbWriteOff")
         
                ArrSizeX = tbAllData.ListColumns.count
                ArrSizeY = 0
                
               
                        ArrSizeY = tbWriteOff.DataBodyRange.Rows.count
                            
                            
                            ReDim ArrDoc(1 To ArrSizeY, 1 To ArrSizeX)
                            
                            
                arrRow = 1 ' лічильник рядків масиву
                 
                For Each WOffRow In tbWriteOff.DataBodyRange.Rows
                wsActualDoc.Activate
'                    If InvRow.value <> "" Then
                    
                          NameThings = Cells(WOffRow.Row, 2).value     'найменування
                          MensureThings = Cells(WOffRow.Row, 4).value 'одиниця вимцру
                          
                          QuantityThings = Cells(WOffRow.Row, 5).value 'кількість
                          ItemNumber = Cells(WOffRow.Row, 3).value 'номенклатурний номер
                          NumberBasis = Cells(WOffRow.Row, 8).value 'підстава для списання
                          price = Cells(WOffRow.Row, 6).value 'ціна
                          
                          
'                          NumberDemand
                          Note = Cells(WOffRow.Row, 10).value 'субрахунок
                          СategoryThings = Cells(WOffRow.Row, 11).value 'додаткове поле
                          СategoryThingsTo = Cells(WOffRow.Row, 12).value 'додаткове поле
                          
                          
                          Note1 = ContentAdd
                          Note2 = Cells(WOffRow.Row, 14).value
                          Note3 = Cells(WOffRow.Row, 15).value
                           
                           ArrDoc(arrRow, 1) = InputDate
                           ArrDoc(arrRow, 2) = NameThings
                           ArrDoc(arrRow, 4) = MensureThings
                           ArrDoc(arrRow, 5) = СategoryThings
                           
'                           If Content = "вибуває" Then
                                ArrDoc(arrRow, 7) = QuantityThings
'                           Else
'                                ArrDoc(arrRow, 6) = QuantityThings
'                           End If
'
                           ArrDoc(arrRow, 8) = FromUnit
                           ArrDoc(arrRow, 9) = toUnit
                           ArrDoc(arrRow, 10) = fromMRP
                           ArrDoc(arrRow, 11) = toMRP
                           ArrDoc(arrRow, 12) = NumberBasis
                           ArrDoc(arrRow, 13) = DateDoc
                           ArrDoc(arrRow, 14) = Number
                           ArrDoc(arrRow, 15) = NumberClear
                           ArrDoc(arrRow, 16) = Note
                           
                           ArrDoc(arrRow, 17) = Executor
                           ArrDoc(arrRow, 18) = ItemNumber
                           
'                           ArrDoc(arrRow, 19) = TimeInv
                           ArrDoc(arrRow, 20) = price
                           ArrDoc(arrRow, 21) = Service
                                                      
                           ArrDoc(arrRow, 22) = Note1
                           ArrDoc(arrRow, 23) = Note2
                           ArrDoc(arrRow, 24) = Note3
                           ArrDoc(arrRow, 25) = typeDoc
                           ArrDoc(arrRow, 26) = Content
                           ArrDoc(arrRow, 27) = СategoryThingsTo
                           
                        arrRow = arrRow + 1
        
                  
                
                Next WOffRow
    
    '****************************************************************************************
       Case "WriteOff81"
       
        Number = Trim(wsActualDoc.Range("NumberWOff81").value) 'номер документа
'                NumberClear = Trim(wsActual.Range("clearNumber"))
                    
               
                    'перевірка, чи існує вже такий ордер
                    On Error Resume Next
                    
                    
                        Set RangeToSearch = tbAllData.ListColumns(14).DataBodyRange
                        DocExist = FindValue(Number, RangeToSearch)
                        
                        If DocExist Then
                            
                            msgExist = MsgBox("Документ: " & Number & " вже рознесений в таблицю. Все одно рознести?", vbYesNo, "Документ з таким номером є в базі")
                                
                                If msgExist = 7 Then
                                    GoTo nextDoc
                                End If
                                
                        End If
                     On Error GoTo 0
                
                Service = Trim(wsActualDoc.Range("ServiceWOff81"))
                
                Content = Trim(wsActualDoc.Range("ContentWOff81"))
                'додатковий тип операції, для донесення
                ContentAdd = Trim(wsActualDoc.Range("ContentInvAdd81_WO"))
                
                typeDoc = Trim(wsActualDoc.Range("typeDoc81"))
                InputDate = Date 'дата введення
                Executor = wsActualDoc.Range("ExecutorWOff81") ' виконавець (якщо тут помилка, скоріше всього клітинці
                'не присвоєне ім'я Exsecutor. Скопіювати Exsecutor, вставити в назву клітинки з виконавцем, натиснути Enter
                FromUnit = Trim(wsActualDoc.Range("FromUnitWOff81")) 'відправник
                
                toUnit = Trim(wsActualDoc.Range("toUnitWOff81")) 'одержувач

                On Error Resume Next
                fromMRP = Trim(wsActualDoc.Range("MRPWOff81"))
                On Error GoTo 0
                toMRP = Trim(wsActualDoc.Range("toUnitWOff81"))
                

                DateDoc = wsActualDoc.Range("DateWOff81") 'дата документа
                NumberBasis = wsActualDoc.Range("basisWOff81")

            
            Set tbActual = wsActualDoc.ListObjects("tbWriteOff81")
            
            
            
           
                ArrSizeX = tbAllData.ListColumns.count
                ArrSizeY = 0
                
                            
                        ArrSizeY = tbActual.DataBodyRange.Rows.count
                            
                            
                            ReDim ArrDoc(1 To ArrSizeY, 1 To ArrSizeX)
                            
                            
                arrRow = 1 ' лічильник рядків масиву
                 
                For Each DocRow In tbActual.DataBodyRange.Rows
                wsActualDoc.Activate
'                    If InvRow.value <> "" Then
                    
                          NameThings = Cells(DocRow.Row, 2).value     'найменування
                          MensureThings = Cells(DocRow.Row, 4).value 'одиниця вимцру
                          
                          QuantityThings = Cells(DocRow.Row, 7).value 'кількість
                          ItemNumber = Cells(DocRow.Row, 3).value 'номенклатурний номер
'                          NumberBasis = Cells(DocRow.Row, 8).value 'підстава для списання
                          price = Cells(DocRow.Row, 6).value 'ціна
                          Note = Cells(DocRow.Row, 14).value
                          Note1 = ContentAdd
                          Note2 = Cells(DocRow.Row, 16).value
                          Note3 = Cells(DocRow.Row, 17).value
                         
                          
                          СategoryThings = Cells(DocRow.Row, 5).value 'додаткове поле
                          СategoryThingsTo = Cells(DocRow.Row, 18).value 'додаткове поле
                           
                           ArrDoc(arrRow, 1) = InputDate
                           ArrDoc(arrRow, 2) = NameThings
                           ArrDoc(arrRow, 4) = MensureThings
                           ArrDoc(arrRow, 5) = СategoryThings
                           
                           If Content = "вибуття" Or Content = "внутрішнє переміщення" Then
                                ArrDoc(arrRow, 7) = QuantityThings
                                
                           Else
                                ArrDoc(arrRow, 6) = QuantityThings
                           End If
'
                           ArrDoc(arrRow, 8) = FromUnit
                           ArrDoc(arrRow, 9) = toUnit
                           ArrDoc(arrRow, 10) = fromMRP
                           ArrDoc(arrRow, 11) = toMRP
                           ArrDoc(arrRow, 12) = NumberBasis
                           ArrDoc(arrRow, 13) = DateDoc
                           ArrDoc(arrRow, 14) = Number
                           ArrDoc(arrRow, 15) = NumberClear
                           ArrDoc(arrRow, 16) = Note
                           
                           ArrDoc(arrRow, 17) = Executor
                           ArrDoc(arrRow, 18) = ItemNumber
                           
'                           ArrDoc(arrRow, 19) = TimeInv
                           ArrDoc(arrRow, 20) = price
                           ArrDoc(arrRow, 21) = Service
                                                      
                           ArrDoc(arrRow, 22) = Note1
                           ArrDoc(arrRow, 23) = Note2
                           ArrDoc(arrRow, 24) = Note3
                           ArrDoc(arrRow, 25) = typeDoc
                           ArrDoc(arrRow, 26) = Content
                           ArrDoc(arrRow, 27) = СategoryThingsTo
                           
                        arrRow = arrRow + 1

             
                Next DocRow
 '*****************************************************************************************
    
       Case "Inv431"
       Number = Trim(wsActualDoc.Range("NumberInv431").value)
       
       On Error Resume Next
            
                DocExist = FindValue(Number, RangeToSearch)
                
                If DocExist Then
                    
                    msgExist = MsgBox("Документ: " & Number & " вже рознесений в таблицю. Все одно рознести?", vbYesNo, "Документ з таким номером є в базі")
                        
                        If msgExist = 7 Then
                            GoTo nextDoc
                        End If
                        
                End If
        On Error GoTo 0
       
       Service = Trim(wsActualDoc.Range("ServiceInv431"))
                
       Content = Trim(wsActualDoc.Range("ContentInv431"))
        'додатковий тип операції, для донесення
        ContentAdd = Trim(wsActualDoc.Range("ContentInvAdd431"))
                
       typeDoc = Trim(wsActualDoc.Range("typeDocInv431"))
       
       InputDate = Date 'дата введення
                Executor = wsActualDoc.Range("ExecutorInv431") ' виконавець (якщо тут помилка, скоріше всього клітинці
                'не присвоєне ім'я Exsecutor. Скопіювати Exsecutor, вставити в назву клітинки з виконавцем, натиснути Enter
                FromUnit = Trim(wsActualDoc.Range("FromUnitInv431")) 'відправник
                
                toUnit = Trim(wsActualDoc.Range("toUnitInv431")) 'одержувач

                On Error Resume Next
                fromMRP = Trim(wsActualDoc.Range("MRPInv431"))
                On Error GoTo 0
                toMRP = Trim(wsActualDoc.Range("MRPToInv431"))
                

                DateDoc = wsActualDoc.Range("DateInv431") 'дата документа
                NumberBasis = wsActualDoc.Range("basis431")
            
            
            
            Set tbActual = wsActualDoc.ListObjects("tbInv431")
            
          
                ArrSizeX = tbAllData.ListColumns.count
                ArrSizeY = 0
                
                            
                        ArrSizeY = tbActual.DataBodyRange.Rows.count
                            
                            
                            ReDim ArrDoc(1 To ArrSizeY, 1 To ArrSizeX)
                            
                            
                arrRow = 1 ' лічильник рядків масиву
                 
                For Each DocRow In tbActual.DataBodyRange.Rows
                wsActualDoc.Activate
'                    If InvRow.value <> "" Then
                    
                          NameThings = Cells(DocRow.Row, 2).value     'найменування
                          MensureThings = Cells(DocRow.Row, 4).value 'одиниця вимцру
                          
                          QuantityThings = Cells(DocRow.Row, 5).value 'кількість
                          ItemNumber = Cells(DocRow.Row, 3).value 'номенклатурний номер
'                          NumberBasis = Cells(DocRow.Row, 8).value 'підстава для списання
                          price = Cells(DocRow.Row, 7).value 'ціна
                          Note = Cells(DocRow.Row, 12).value
                          Note1 = ContentAdd
                          Note2 = Cells(DocRow.Row, 15).value
                          Note3 = Cells(DocRow.Row, 16).value
                         
                          
                          СategoryThings = Cells(DocRow.Row, 10).value 'додаткове поле
                          СategoryThingsTo = Cells(DocRow.Row, 11).value 'додаткове поле
                           
                           ArrDoc(arrRow, 1) = InputDate
                           ArrDoc(arrRow, 2) = NameThings
                           ArrDoc(arrRow, 4) = MensureThings
                           ArrDoc(arrRow, 5) = СategoryThings
                           
                           If Content = "вибуття" Or Content = "внутрішнє переміщення" Then
                                ArrDoc(arrRow, 7) = QuantityThings
                                
                           Else
                                ArrDoc(arrRow, 6) = QuantityThings
                           End If
'
                           ArrDoc(arrRow, 8) = FromUnit
                           ArrDoc(arrRow, 9) = toUnit
                           ArrDoc(arrRow, 10) = fromMRP
                           ArrDoc(arrRow, 11) = toMRP
                           ArrDoc(arrRow, 12) = NumberBasis
                           ArrDoc(arrRow, 13) = DateDoc
                           ArrDoc(arrRow, 14) = Number
                           ArrDoc(arrRow, 15) = NumberClear
                           ArrDoc(arrRow, 16) = Note
                           
                           ArrDoc(arrRow, 17) = Executor
                           ArrDoc(arrRow, 18) = ItemNumber
                           
'                           ArrDoc(arrRow, 19) = TimeInv
                           ArrDoc(arrRow, 20) = price
                           ArrDoc(arrRow, 21) = Service
                                                      
                           ArrDoc(arrRow, 22) = Note1
                           ArrDoc(arrRow, 23) = Note2
                           ArrDoc(arrRow, 24) = Note3
                           ArrDoc(arrRow, 25) = typeDoc
                           ArrDoc(arrRow, 26) = Content
                           ArrDoc(arrRow, 27) = СategoryThingsTo
                           
                        arrRow = arrRow + 1
'                    End If
  
                    
                Next DocRow
     '*********************************************************************************************************
       Case "Inv440"
          Number = Trim(wsActualDoc.Range("NumberInv").value) 'номер накладної
                
                
                    'перевірка, чи існує вже такий ордер
                    On Error Resume Next
                    
                   
                        DocExist = FindValue(Number, RangeToSearch)
                        
                        If DocExist Then
                            
                            msgExist = MsgBox("Накладна: " & Number & " вже рознесена в таблицю. Все одно рознести?", vbYesNo, "Накладна існує")
                                
                                If msgExist = 7 Then
                                    GoTo nextDoc
                                End If
                                
                        End If
                     On Error GoTo 0
                NumberClear = Trim(wsActualDoc.Range("NumberInvClear"))
                    
                Service = Trim(wsActualDoc.Range("Service"))
                
                Content = Trim(wsActualDoc.Range("Content"))
                
                'додатковий тип операції, для донесення
                ContentAdd = Trim(wsActualDoc.Range("ContentInvAdd431"))
                
                typeDoc = Trim(wsActualDoc.Range("TypeDoc"))
                InputDate = Date 'дата введення
                Executor = wsActualDoc.Range("Executor") ' виконавець (якщо тут помилка, скоріше всього клітинці
                'не присвоєне ім'я Exsecutor. Скопіювати Exsecutor, вставити в назву клітинки з виконавцем, натиснути Enter
                FromUnit = Trim(wsActualDoc.Range("FromUnit")) 'відправник
                toUnit = Trim(wsActualDoc.Range("ToUnit")) 'одержувач
                
                On Error Resume Next
                fromMRP = Trim(wsActualDoc.Range("FromMRP"))
                On Error GoTo 0
                toMRP = Trim(wsActualDoc.Range("ToMRP"))
                

                
                DateDoc = wsActualDoc.Range("DateInv") 'дата накладної
                TimeInv = wsActualDoc.Range("TimeInv")
                NumberBasis = wsActualDoc.Range("NumberDemand")

            
            Dim InvFields As Range
            Set InvFields = wsActualDoc.Range("InvFields")
           
                ArrSizeX = tbAllData.ListColumns.count
                ArrSizeY = 0
                
                            For Each InvRow In InvFields.Columns(1).Cells ' шукаємо кількість позицій в наряді
                                If InvRow.value <> "" Then
                                   ArrSizeY = ArrSizeY + 1
                                End If
                            Next
                            
                            ReDim ArrDoc(1 To ArrSizeY, 1 To ArrSizeX)
                            
                            
                arrRow = 1 ' лічильник рядків масиву
                 
                For Each InvRow In InvFields.Columns(1).Cells
                wsActualDoc.Activate
                    If InvRow.value <> "" Then
                    
                          NameThings = Cells(InvRow.Row, 2).value     'найменування
                          MensureThings = Cells(InvRow.Row, 4).value 'одиниця вимцру
                          СategoryThings = Cells(InvRow.Row, 5).value 'категорія
                          СategoryThingsTo = Cells(InvRow.Row, 13).value
                          QuantityThings = Cells(InvRow.Row, 6).value 'кількість
                          ItemNumber = Cells(InvRow.Row, 3).value 'номенклатурний номер
                          Note = Cells(InvRow.Row, 8).value 'примітка
                          price = Cells(InvRow.Row, 9).value 'ціна
                          
                          Note1 = Cells(InvRow.Row, 10).value 'додаткове поле
                          Note2 = Cells(InvRow.Row, 11).value 'додаткове поле
                          Note3 = Cells(InvRow.Row, 12).value 'додаткове поле
                           
                           ArrDoc(arrRow, 1) = InputDate
                           ArrDoc(arrRow, 2) = NameThings
                           ArrDoc(arrRow, 4) = MensureThings
                           ArrDoc(arrRow, 5) = СategoryThings
                           
                           If Content = "вибуття" Or Content = "внутрішнє переміщення" Then
                                ArrDoc(arrRow, 7) = QuantityThings
                           Else
                                ArrDoc(arrRow, 6) = QuantityThings
                           End If
                                                                                 
                           ArrDoc(arrRow, 8) = FromUnit
                           ArrDoc(arrRow, 9) = toUnit
                           ArrDoc(arrRow, 10) = fromMRP
                           ArrDoc(arrRow, 11) = toMRP
                           ArrDoc(arrRow, 12) = NumberBasis
                           ArrDoc(arrRow, 13) = DateDoc
                           ArrDoc(arrRow, 14) = Number
                           ArrDoc(arrRow, 15) = NumberClear
                           ArrDoc(arrRow, 16) = Note
                           
                           ArrDoc(arrRow, 17) = Executor
                           ArrDoc(arrRow, 18) = ItemNumber
                           
                           ArrDoc(arrRow, 19) = TimeInv
                           ArrDoc(arrRow, 20) = price
                           ArrDoc(arrRow, 21) = Service
                                                      
                           ArrDoc(arrRow, 22) = Note1
                           ArrDoc(arrRow, 23) = Note2
                           ArrDoc(arrRow, 24) = Note3
                           
                           ArrDoc(arrRow, 25) = typeDoc
                           ArrDoc(arrRow, 26) = Content
                           ArrDoc(arrRow, 27) = СategoryThingsTo
                        arrRow = arrRow + 1
                    End If
            

                    
                Next InvRow
       Case "Inv440_25"
       
       Number = Trim(wsActualDoc.Range("NumberInv440_25").value)
       
       On Error Resume Next
            
                DocExist = FindValue(Number, RangeToSearch)
                
                If DocExist Then
                    
                    msgExist = MsgBox("Документ: " & Number & " вже рознесений в таблицю. Все одно рознести?", vbYesNo, "Документ з таким номером є в базі")
                        
                        If msgExist = 7 Then
                            GoTo nextDoc
                        End If
                        
                End If
        On Error GoTo 0
       
       Service = Trim(wsActualDoc.Range("ServiceInv440_25"))
                
       Content = Trim(wsActualDoc.Range("ContentInv440_25"))
       'додатковий тип операції, для донесення
        ContentAdd = Trim(wsActualDoc.Range("ContentInvAdd440_25"))
                
       typeDoc = Trim(wsActualDoc.Range("typeDocInv440_25"))
       
       InputDate = Date 'дата введення
                Executor = wsActualDoc.Range("ExecutorInv440_25") ' виконавець (якщо тут помилка, скоріше всього клітинці
                'не присвоєне ім'я Exsecutor. Скопіювати Exsecutor, вставити в назву клітинки з виконавцем, натиснути Enter
                FromUnit = Trim(wsActualDoc.Range("FromUnitInv440_25")) 'відправник
                
                toUnit = Trim(wsActualDoc.Range("toUnitInv440_25")) 'одержувач

                On Error Resume Next
                fromMRP = Trim(wsActualDoc.Range("MRPInv440_25"))
                On Error GoTo 0
                toMRP = Trim(wsActualDoc.Range("MRPToInv440_25"))
                

                DateDoc = wsActualDoc.Range("DateInv440_25") 'дата документа
                NumberBasis = wsActualDoc.Range("basis440_25")
            
            
            
            Set tbActual = wsActualDoc.ListObjects("tbInv440_25")
            
          
                ArrSizeX = tbAllData.ListColumns.count
                ArrSizeY = 0
                
                            
                        ArrSizeY = tbActual.DataBodyRange.Rows.count
                            
                            
                            ReDim ArrDoc(1 To ArrSizeY, 1 To ArrSizeX)
                            
                            
                arrRow = 1 ' лічильник рядків масиву
                 
                For Each DocRow In tbActual.DataBodyRange.Rows
                wsActualDoc.Activate
'                    If InvRow.value <> "" Then
                    
                          NameThings = Cells(DocRow.Row, 2).value     'найменування
                          MensureThings = Cells(DocRow.Row, 4).value 'одиниця вимцру
                          
                          QuantityThings = Cells(DocRow.Row, 7).value 'кількість
                          ItemNumber = Cells(DocRow.Row, 3).value 'номенклатурний номер
'                          NumberBasis = Cells(DocRow.Row, 8).value 'підстава для списання
                          price = Cells(DocRow.Row, 6).value 'ціна
                          Note = Cells(DocRow.Row, 10).value
                          Note1 = ContentAdd
                          Note2 = Cells(DocRow.Row, 13).value
                          Note3 = Cells(DocRow.Row, 14).value
                         
                          
                          СategoryThings = Cells(DocRow.Row, 5).value 'додаткове поле
                          СategoryThingsTo = Cells(DocRow.Row, 11).value 'додаткове поле
                           
                           ArrDoc(arrRow, 1) = InputDate
                           ArrDoc(arrRow, 2) = NameThings
                           ArrDoc(arrRow, 4) = MensureThings
                           ArrDoc(arrRow, 5) = СategoryThings
                           
                           If Content = "вибуття" Or Content = "внутрішнє переміщення" Then
                                ArrDoc(arrRow, 7) = QuantityThings
                                
                           Else
                                ArrDoc(arrRow, 6) = QuantityThings
                           End If
'
                           ArrDoc(arrRow, 8) = FromUnit
                           ArrDoc(arrRow, 9) = toUnit
                           ArrDoc(arrRow, 10) = fromMRP
                           ArrDoc(arrRow, 11) = toMRP
                           ArrDoc(arrRow, 12) = NumberBasis
                           ArrDoc(arrRow, 13) = DateDoc
                           ArrDoc(arrRow, 14) = Number
                           ArrDoc(arrRow, 15) = NumberClear
                           ArrDoc(arrRow, 16) = Note
                           
                           ArrDoc(arrRow, 17) = Executor
                           ArrDoc(arrRow, 18) = ItemNumber
                           
'                           ArrDoc(arrRow, 19) = TimeInv
                           ArrDoc(arrRow, 20) = price
                           ArrDoc(arrRow, 21) = Service
                                                      
                           ArrDoc(arrRow, 22) = Note1
                           ArrDoc(arrRow, 23) = Note2
                           ArrDoc(arrRow, 24) = Note3
                           ArrDoc(arrRow, 25) = typeDoc
                           ArrDoc(arrRow, 26) = Content
                           ArrDoc(arrRow, 27) = СategoryThingsTo
                           
                        arrRow = arrRow + 1
'                    End If
  
                    
                Next DocRow
                
  '========================================================================
  Case "FA440_23"

       
       Number = Trim(wsActualDoc.Range("NumberFA440_23").value)
       
       On Error Resume Next
            
                DocExist = FindValue(Number, RangeToSearch)
                
                If DocExist Then
                    
                    msgExist = MsgBox("Документ: " & Number & " вже рознесений в таблицю. Все одно рознести?", vbYesNo, "Документ з таким номером є в базі")
                        
                        If msgExist = 7 Then
                            GoTo nextDoc
                        End If
                        
                End If
        On Error GoTo 0
       
       Service = Trim(wsActualDoc.Range("ServiceFA440_23"))
                
       Content = Trim(wsActualDoc.Range("ContentFA440_23"))
       'додатковий тип операції, для донесення
       ContentAdd = Trim(wsActualDoc.Range("ContentInvAdd440_23"))
                
       typeDoc = Trim(wsActualDoc.Range("typeDocFA440_23"))
       
       InputDate = Date 'дата введення
                Executor = wsActualDoc.Range("ExecutorFA440_23") ' виконавець (якщо тут помилка, скоріше всього клітинці
                'не присвоєне ім'я Exsecutor. Скопіювати Exsecutor, вставити в назву клітинки з виконавцем, натиснути Enter
                FromUnit = Trim(wsActualDoc.Range("FromUnitFA440_23")) 'відправник
                
                toUnit = Trim(wsActualDoc.Range("toUnitFA440_23")) 'одержувач

                On Error Resume Next
                fromMRP = Trim(wsActualDoc.Range("MRPFromFA440_23"))
                On Error GoTo 0
                toMRP = Trim(wsActualDoc.Range("MRPToFA440_23"))
                

                DateDoc = wsActualDoc.Range("DateFA440_23") 'дата документа
                NumberBasis = wsActualDoc.Range("basisFA440_23")
            
            
            
            Set tbActual = wsActualDoc.ListObjects("tbFA440_23")
            
          
                ArrSizeX = tbAllData.ListColumns.count
                ArrSizeY = 0
                
                            
                        ArrSizeY = tbActual.DataBodyRange.Rows.count
                            
                            
                            ReDim ArrDoc(1 To ArrSizeY, 1 To ArrSizeX)
                            
                            
                arrRow = 1 ' лічильник рядків масиву
                 
                For Each DocRow In tbActual.DataBodyRange.Rows
                wsActualDoc.Activate
'                    If InvRow.value <> "" Then
                    
                          NameThings = Cells(DocRow.Row, 2).value     'найменування
                          MensureThings = Cells(DocRow.Row, 16).value 'одиниця вимцру
                          
                          QuantityThings = Cells(DocRow.Row, 8).value 'кількість
                          ItemNumber = Cells(DocRow.Row, 3).value 'номенклатурний номер
'                          NumberBasis = Cells(DocRow.Row, 8).value 'підстава для списання
                          price = Cells(DocRow.Row, 5).value 'ціна
                          Note = Cells(DocRow.Row, 14).value ' в цьому документі заводський номер
                          Note1 = ContentAdd '
                          Note2 = Cells(DocRow.Row, 18).value
                          Note3 = Cells(DocRow.Row, 19).value
                         
                          
                          СategoryThings = Cells(DocRow.Row, 7).value 'додаткове поле
                          СategoryThingsTo = Cells(DocRow.Row, 9).value 'додаткове поле
                           
                           ArrDoc(arrRow, 1) = InputDate
                           ArrDoc(arrRow, 2) = NameThings
                           ArrDoc(arrRow, 4) = MensureThings
                           ArrDoc(arrRow, 5) = СategoryThings
                           
                           If Content = "вибуття" Or Content = "внутрішнє переміщення" Then
                                ArrDoc(arrRow, 7) = QuantityThings
                                
                           Else
                                ArrDoc(arrRow, 6) = QuantityThings
                           End If
'
                           ArrDoc(arrRow, 8) = FromUnit
                           ArrDoc(arrRow, 9) = toUnit
                           ArrDoc(arrRow, 10) = fromMRP
                           ArrDoc(arrRow, 11) = toMRP
                           ArrDoc(arrRow, 12) = NumberBasis
                           ArrDoc(arrRow, 13) = DateDoc
                           ArrDoc(arrRow, 14) = Number
                           ArrDoc(arrRow, 15) = NumberClear
                           ArrDoc(arrRow, 16) = Note
                           
                           ArrDoc(arrRow, 17) = Executor
                           ArrDoc(arrRow, 18) = ItemNumber
                           
'                           ArrDoc(arrRow, 19) = TimeInv
                           ArrDoc(arrRow, 20) = price
                           ArrDoc(arrRow, 21) = Service
                                                      
                           ArrDoc(arrRow, 22) = Note1
                           ArrDoc(arrRow, 23) = Note2
                           ArrDoc(arrRow, 24) = Note3
                           ArrDoc(arrRow, 25) = typeDoc
                           ArrDoc(arrRow, 26) = Content
                           ArrDoc(arrRow, 27) = СategoryThingsTo
                           
                        arrRow = arrRow + 1
'                    End If
  
                    
                Next DocRow
  '========================================================================

  Case "FA440_24"

       
       Number = Trim(wsActualDoc.Range("NumberFA440_24").value)
       
       On Error Resume Next
            
                DocExist = FindValue(Number, RangeToSearch)
                
                If DocExist Then
                    
                    msgExist = MsgBox("Документ: " & Number & " вже рознесений в таблицю. Все одно рознести?", vbYesNo, "Документ з таким номером є в базі")
                        
                        If msgExist = 7 Then
                            GoTo nextDoc
                        End If
                        
                End If
        On Error GoTo 0
       
       Service = Trim(wsActualDoc.Range("ServiceFA440_24"))
                
       Content = Trim(wsActualDoc.Range("ContentFA440_24"))
       'додатковий тип операції, для донесення
       ContentAdd = Trim(wsActualDoc.Range("ContentInvAdd440_24"))
                
       typeDoc = Trim(wsActualDoc.Range("typeDocFA440_24"))
       
       InputDate = Date 'дата введення
                Executor = wsActualDoc.Range("ExecutorFA440_24") ' виконавець (якщо тут помилка, скоріше всього клітинці
                'не присвоєне ім'я Exsecutor. Скопіювати Exsecutor, вставити в назву клітинки з виконавцем, натиснути Enter
                FromUnit = Trim(wsActualDoc.Range("FromUnitFA440_24")) 'відправник
                
                toUnit = Trim(wsActualDoc.Range("toUnitFA440_24")) 'одержувач

                On Error Resume Next
                fromMRP = Trim(wsActualDoc.Range("MRPFromFA440_24"))
                On Error GoTo 0
                toMRP = Trim(wsActualDoc.Range("MRPToFA440_24"))
                

                DateDoc = wsActualDoc.Range("DateFA440_24") 'дата документа
                NumberBasis = wsActualDoc.Range("basisFA440_24")
            
            
            
            Set tbActual = wsActualDoc.ListObjects("tbFA440_24")
            
          
                ArrSizeX = tbAllData.ListColumns.count
                ArrSizeY = 0
                
                            
                        ArrSizeY = tbActual.DataBodyRange.Rows.count
                            
                            
                            ReDim ArrDoc(1 To ArrSizeY, 1 To ArrSizeX)
                            
                            
                arrRow = 1 ' лічильник рядків масиву
                 
                For Each DocRow In tbActual.DataBodyRange.Rows
                wsActualDoc.Activate
'                    If InvRow.value <> "" Then
                    
                          NameThings = Cells(DocRow.Row, 2).value     'найменування
                          MensureThings = Cells(DocRow.Row, 5).value 'одиниця вимцру
                          
                          QuantityThings = Cells(DocRow.Row, 9).value 'кількість
                          ItemNumber = Cells(DocRow.Row, 3).value 'номенклатурний номер
'                          NumberBasis = Cells(DocRow.Row, 8).value 'підстава для списання
                          price = Cells(DocRow.Row, 6).value 'ціна
                          Note = Cells(DocRow.Row, 12).value ' в цьому документі заводський номер
                          Note1 = ContentAdd
                          Note2 = Cells(DocRow.Row, 14).value
                          Note3 = Cells(DocRow.Row, 15).value
                         
                          
                          СategoryThings = Cells(DocRow.Row, 8).value 'додаткове поле
                          СategoryThingsTo = Cells(DocRow.Row, 10).value 'додаткове поле
                           
                           ArrDoc(arrRow, 1) = InputDate
                           ArrDoc(arrRow, 2) = NameThings
                           ArrDoc(arrRow, 4) = MensureThings
                           ArrDoc(arrRow, 5) = СategoryThings
                           
                           If Content = "вибуття" Or Content = "внутрішнє переміщення" Then
                                ArrDoc(arrRow, 7) = QuantityThings
                                
                           Else
                                ArrDoc(arrRow, 6) = QuantityThings
                           End If
'
                           ArrDoc(arrRow, 8) = FromUnit
                           ArrDoc(arrRow, 9) = toUnit
                           ArrDoc(arrRow, 10) = fromMRP
                           ArrDoc(arrRow, 11) = toMRP
                           ArrDoc(arrRow, 12) = NumberBasis
                           ArrDoc(arrRow, 13) = DateDoc
                           ArrDoc(arrRow, 14) = Number
                           ArrDoc(arrRow, 15) = NumberClear
                           ArrDoc(arrRow, 16) = Note
                           
                           ArrDoc(arrRow, 17) = Executor
                           ArrDoc(arrRow, 18) = ItemNumber
                           
'                           ArrDoc(arrRow, 19) = TimeInv
                           ArrDoc(arrRow, 20) = price
                           ArrDoc(arrRow, 21) = Service
                                                      
                           ArrDoc(arrRow, 22) = Note1
                           ArrDoc(arrRow, 23) = Note2
                           ArrDoc(arrRow, 24) = Note3
                           ArrDoc(arrRow, 25) = typeDoc
                           ArrDoc(arrRow, 26) = Content
                           ArrDoc(arrRow, 27) = СategoryThingsTo
                           
                        arrRow = arrRow + 1
'                    End If
  
                    
                Next DocRow
  '========================================================================
 Case "FA818"
       
       Number = Trim(wsActualDoc.Range("NumberInvFA818").value)
       
       On Error Resume Next
            
                DocExist = FindValue(Number, RangeToSearch)
                
                If DocExist Then
                    
                    msgExist = MsgBox("Документ: " & Number & " вже рознесений в таблицю. Все одно рознести?", vbYesNo, "Документ з таким номером є в базі")
                        
                        If msgExist = 7 Then
                            GoTo nextDoc
                        End If
                        
                End If
        On Error GoTo 0
       
       Service = Trim(wsActualDoc.Range("ServiceInvFA818"))
                
       Content = Trim(wsActualDoc.Range("ContentInvFA818"))
        'додатковий тип операції, для донесення
       ContentAdd = Trim(wsActualDoc.Range("ContentInvAdd818"))
                
       typeDoc = Trim(wsActualDoc.Range("typeDocInvFA818"))
       
       InputDate = Date 'дата введення
                Executor = wsActualDoc.Range("ExecutorInvFA818") ' виконавець (якщо тут помилка, скоріше всього клітинці
                'не присвоєне ім'я Exsecutor. Скопіювати Exsecutor, вставити в назву клітинки з виконавцем, натиснути Enter
                FromUnit = Trim(wsActualDoc.Range("FromUnitInvFA818")) 'відправник
                
                toUnit = Trim(wsActualDoc.Range("toUnitInvFA818")) 'одержувач

                On Error Resume Next
                fromMRP = Trim(wsActualDoc.Range("MRPInvFromFA818"))
                On Error GoTo 0
                toMRP = Trim(wsActualDoc.Range("MRPToInvFA818"))
                

                DateDoc = wsActualDoc.Range("DateInvFA818") 'дата документа
                NumberBasis = wsActualDoc.Range("basisFA818")
            
            
            
            Set tbActual = wsActualDoc.ListObjects("tbInvFA818")
            
          
                ArrSizeX = tbAllData.ListColumns.count
                ArrSizeY = 0
                
                            
                        ArrSizeY = tbActual.DataBodyRange.Rows.count
                            
                            
                            ReDim ArrDoc(1 To ArrSizeY, 1 To ArrSizeX)
                            
                            
                arrRow = 1 ' лічильник рядків масиву
                 
                For Each DocRow In tbActual.DataBodyRange.Rows
                wsActualDoc.Activate
'                    If InvRow.value <> "" Then
                    
                          NameThings = Cells(DocRow.Row, 1).value     'найменування
                          MensureThings = Cells(DocRow.Row, 3).value 'одиниця вимцру
                          
                          QuantityThings = Cells(DocRow.Row, 4).value 'кількість
                          ItemNumber = Cells(DocRow.Row, 2).value 'номенклатурний номер
'                          NumberBasis = Cells(DocRow.Row, 8).value 'підстава для списання
                          price = Cells(DocRow.Row, 5).value 'ціна
                          Note = Cells(DocRow.Row, 7).value
                          Note1 = ContentAdd
                          Note2 = Cells(DocRow.Row, 11).value
                          Note3 = Cells(DocRow.Row, 12).value
                         
                          
                          СategoryThings = Cells(DocRow.Row, 8).value 'додаткове поле
                          СategoryThingsTo = Cells(DocRow.Row, 9).value 'додаткове поле
                           
                           ArrDoc(arrRow, 1) = InputDate
                           ArrDoc(arrRow, 2) = NameThings
                           ArrDoc(arrRow, 4) = MensureThings
                           ArrDoc(arrRow, 5) = СategoryThings
                           
                           If Content = "вибуття" Or Content = "внутрішнє переміщення" Then
                                ArrDoc(arrRow, 7) = QuantityThings
                                
                           Else
                                ArrDoc(arrRow, 6) = QuantityThings
                           End If
'
                           ArrDoc(arrRow, 8) = FromUnit
                           ArrDoc(arrRow, 9) = toUnit
                           ArrDoc(arrRow, 10) = fromMRP
                           ArrDoc(arrRow, 11) = toMRP
                           ArrDoc(arrRow, 12) = NumberBasis
                           ArrDoc(arrRow, 13) = DateDoc
                           ArrDoc(arrRow, 14) = Number
                           ArrDoc(arrRow, 15) = NumberClear
                           ArrDoc(arrRow, 16) = Note
                           
                           ArrDoc(arrRow, 17) = Executor
                           ArrDoc(arrRow, 18) = ItemNumber
                           
'                           ArrDoc(arrRow, 19) = TimeInv
                           ArrDoc(arrRow, 20) = price
                           ArrDoc(arrRow, 21) = Service
                                                      
                           ArrDoc(arrRow, 22) = Note1
                           ArrDoc(arrRow, 23) = Note2
                           ArrDoc(arrRow, 24) = Note3
                           ArrDoc(arrRow, 25) = typeDoc
                           ArrDoc(arrRow, 26) = Content
                           ArrDoc(arrRow, 27) = СategoryThingsTo
                           
                        arrRow = arrRow + 1
'                    End If
  
                    
                Next DocRow
       '========================================================================
Case "payroll440_57"

       
       Number = Trim(wsActualDoc.Range("NumberInv440_57").value)
       
       On Error Resume Next
            
                DocExist = FindValue(Number, RangeToSearch)
                
                If DocExist Then
                    
                    msgExist = MsgBox("Документ: " & Number & " вже рознесений в таблицю. Все одно рознести?", vbYesNo, "Документ з таким номером є в базі")
                        
                        If msgExist = 7 Then
                            GoTo nextDoc
                        End If
                        
                End If
        On Error GoTo 0
       
       Service = Trim(wsActualDoc.Range("ServiceInv440_57"))
                
       Content = Trim(wsActualDoc.Range("ContentInv440_57"))
       'додатковий тип операції, для донесення
       ContentAdd = Trim(wsActualDoc.Range("ContentInvAdd440_57"))
                
       typeDoc = Trim(wsActualDoc.Range("typeDocInv440_57"))
       
       InputDate = Date 'дата введення
                Executor = wsActualDoc.Range("ExecutorInv440_57") ' виконавець (якщо тут помилка, скоріше всього клітинці
                'не присвоєне ім'я Exsecutor. Скопіювати Exsecutor, вставити в назву клітинки з виконавцем, натиснути Enter
                FromUnit = Trim(wsActualDoc.Range("FromUnitInv440_57")) 'відправник
                
                toUnit = Trim(wsActualDoc.Range("toUnitInv440_57")) 'одержувач

'                On Error Resume Next
'                FromMRP = Trim(wsActualDoc.Range("MRPInvFrom440_57"))
'                On Error GoTo 0
'                ToMRP = Trim(wsActualDoc.Range("MRPToInv440_57"))
                

                DateDoc = wsActualDoc.Range("DateInv440_57") 'дата документа
                NumberBasis = wsActualDoc.Range("basis440_57")
            
            
            
            Set tbActual = wsActualDoc.ListObjects("tbInv440_57")
            
          
                ArrSizeX = tbAllData.ListColumns.count
                ArrSizeY = 0
                
'                 tbActual.Range.Columns("3").Resize(, 19).Select
                 Dim countNonEmpty As Range
                 Set countNonEmpty = tbActual.DataBodyRange.Columns(3).Resize(, 19)
'                 countNonEmpty = Application.WorksheetFunction.CountA(wsActualDoc.Range(Cells(12, 3), Cells(12, 22)))
                  ArrSizeY = Application.WorksheetFunction.CountIf(countNonEmpty, ">" & 0)
                            
'                        ArrSizeY = tbActual.DataBodyRange.Rows.count * countNonEmpty
                            
                            
                            ReDim ArrDoc(1 To ArrSizeY, 1 To ArrSizeX)
                            
                'продумати нову номерацію для роздавальної відомості
                arrRow = 1 ' лічильник рядків масиву
                 
                For Each DocRow In tbActual.DataBodyRange.Rows
                wsActualDoc.Activate
'                    If InvRow.value <> "" Then
                    'перебрати кожного отримувача, і кожну колонку з цифрами, якщо там 0, тоді йдемо далі, якщо ні - присвоюємо дані
                            ' Перебір кожної клітинки у рядку
                            Dim cell As Range
'                            Переробити на цикл аочинаючи з 3 колонки до кінця
                            For Each cell In DocRow.Cells
                                If Not IsEmpty(cell.value) And cell.Column > 2 And cell.Column < 22 And cell.value <> 0 Then '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                                  NameThings = Cells(12, cell.Column).value
                                  ItemNumber = Cells(14, cell.Column).value
                                  MensureThings = Cells(16, cell.Column).value
                                  СategoryThings = Cells(18, cell.Column).value
                                  price = Cells(20, cell.Column).value
                                  СategoryThingsTo = Cells(22, cell.Column).value
                                  Note = Cells(24, cell.Column).value
                                  Note1 = ContentAdd
                                  Note2 = Cells(28, cell.Column).value
                                  Note3 = Cells(30, cell.Column).value
                                  QuantityThings = Cells(cell.Row, cell.Column).value
                                  
                           If Content = "вибуття" Or Content = "внутрішнє переміщення" Then
                                fromMRP = wsActualDoc.Range("MRP_57")
                                toMRP = Cells(cell.Row, 2).value
                           Else
                                fromMRP = Cells(cell.Row, 2).value
                                toMRP = wsActualDoc.Range("MRP_57")
                           End If
                                 
                                  
                        
                            
                                              
'                          NameThings = Cells(DocRow.Row, 1).value     'найменування
'                          MensureThings = Cells(DocRow.Row, 2).value 'одиниця вимцру
                          
'                          QuantityThings = Cells(DocRow.Row, 4).value 'кількість
'                          ItemNumber = Cells(DocRow.Row, 2).value 'номенклатурний номер
'                          NumberBasis = Cells(DocRow.Row, 8).value 'підстава для списанн
'                          Price = Cells(DocRow.Row, 5).value 'ціна
'                          Note = Cells(DocRow.Row, 7).value
'                          Note1 = Cells(DocRow.Row, 10).value '
'                          Note2 = Cells(DocRow.Row, 11).value
'                          Note3 = Cells(DocRow.Row, 12).value
                        
                          
'                          СategoryThings = Cells(DocRow.Row, 8).value 'додаткове поле
'                          СategoryThingsTo = Cells(DocRow.Row, 9).value 'додаткове поле
                          
                          
                          
                          
                          
                           
                           ArrDoc(arrRow, 1) = InputDate
                           ArrDoc(arrRow, 2) = NameThings
                           ArrDoc(arrRow, 4) = MensureThings
                           ArrDoc(arrRow, 5) = СategoryThings
                           
                           If Content = "вибуття" Or Content = "внутрішнє переміщення" Then
                                ArrDoc(arrRow, 7) = QuantityThings
                                
                           Else
                                ArrDoc(arrRow, 6) = QuantityThings
                           End If
'
                           ArrDoc(arrRow, 8) = FromUnit
                           ArrDoc(arrRow, 9) = toUnit
                           ArrDoc(arrRow, 10) = fromMRP
                           ArrDoc(arrRow, 11) = toMRP
                           ArrDoc(arrRow, 12) = NumberBasis
                           ArrDoc(arrRow, 13) = DateDoc
                           ArrDoc(arrRow, 14) = Number
                           ArrDoc(arrRow, 15) = NumberClear
                           ArrDoc(arrRow, 16) = Note
                           
                           ArrDoc(arrRow, 17) = Executor
                           ArrDoc(arrRow, 18) = ItemNumber
                           
'                           ArrDoc(arrRow, 19) = TimeInv
                           ArrDoc(arrRow, 20) = price
                           ArrDoc(arrRow, 21) = Service
                                                      
                           ArrDoc(arrRow, 22) = Note1
                           ArrDoc(arrRow, 23) = Note2
                           ArrDoc(arrRow, 24) = Note3
                           ArrDoc(arrRow, 25) = typeDoc
                           ArrDoc(arrRow, 26) = Content
                           ArrDoc(arrRow, 27) = СategoryThingsTo
                           
                        arrRow = arrRow + 1
'                    End If
                End If
                    Next cell
                    
                Next DocRow
       
       
'*********************************************************************************************************
Case "WOffLaunch81_19"
       
       Number = Trim(wsActualDoc.Range("NumberWOffLaunch81_19").value)
       
       On Error Resume Next
            
                DocExist = FindValue(Number, RangeToSearch)
                
                If DocExist Then
                    
                    msgExist = MsgBox("Документ: " & Number & " вже рознесений в таблицю. Все одно рознести?", vbYesNo, "Документ з таким номером є в базі")
                        
                        If msgExist = 7 Then
                            GoTo nextDoc
                        End If
                        
                End If
        On Error GoTo 0
       
       Service = Trim(wsActualDoc.Range("ServiceWOffLaunch81_19"))
                
       Content = Trim(wsActualDoc.Range("ContentWOffLaunch81_19"))
        'додатковий тип операції, для донесення
       ContentAdd = Trim(wsActualDoc.Range("ContentAddWOffLaunch81_19"))
                
       typeDoc = Trim(wsActualDoc.Range("typeDocWOffLaunch81_19"))
       
       InputDate = Date 'дата введення
                Executor = wsActualDoc.Range("ExecutorWOffLaunch81_19") ' виконавець (якщо тут помилка, скоріше всього клітинці
                'не присвоєне ім'я Exsecutor. Скопіювати Exsecutor, вставити в назву клітинки з виконавцем, натиснути Enter
                FromUnit = Trim(wsActualDoc.Range("FromUnitWOffLaunch81_19")) 'відправник
                
                toUnit = Trim(wsActualDoc.Range("toUnitWOffLaunch81_19")) 'одержувач

                On Error Resume Next
                fromMRP = Trim(wsActualDoc.Range("MRPFromWOffLaunch81_19"))
                On Error GoTo 0
                toMRP = Trim(wsActualDoc.Range("toUnitWOffLaunch81_19")) 'одержувач

                

                DateDoc = wsActualDoc.Range("DateWOffLaunch81_19") 'дата документа
                NumberBasis = wsActualDoc.Range("basisWOffLaunch81_19")
            
            
            
            Set tbActual = wsActualDoc.ListObjects("tbWOffLaunch81_19")
            
          
                ArrSizeX = tbAllData.ListColumns.count
                ArrSizeY = 0
                
                            
                        ArrSizeY = tbActual.DataBodyRange.Rows.count
                            
                            
                            ReDim ArrDoc(1 To ArrSizeY, 1 To ArrSizeX)
                            
                            
                arrRow = 1 ' лічильник рядків масиву
                 
                For Each DocRow In tbActual.DataBodyRange.Rows
                wsActualDoc.Activate
'                    If InvRow.value <> "" Then
                    
                          NameThings = Cells(DocRow.Row, 2).value     'найменування
                          MensureThings = Cells(DocRow.Row, 5).value 'одиниця вимцру
                          
                          QuantityThings = Cells(DocRow.Row, 7).value 'кількість
                          ItemNumber = Cells(DocRow.Row, 4).value 'номенклатурний номер
'                          NumberBasis = Cells(DocRow.Row, 8).value 'підстава для списання
                          price = Cells(DocRow.Row, 8).value 'ціна
                          Note = Cells(DocRow.Row, 3).value
                          Note1 = ContentAdd
                          Note2 = Cells(DocRow.Row, 11).value
                          Note3 = Cells(DocRow.Row, 12).value
                         
                          
                          СategoryThings = Cells(DocRow.Row, 6).value 'додаткове поле
                          СategoryThingsTo = Cells(DocRow.Row, 10).value 'додаткове поле
                           
                           ArrDoc(arrRow, 1) = InputDate
                           ArrDoc(arrRow, 2) = NameThings
                           ArrDoc(arrRow, 4) = MensureThings
                           ArrDoc(arrRow, 5) = СategoryThings
                           
                           If Content = "вибуття" Or Content = "внутрішнє переміщення" Then
                                ArrDoc(arrRow, 7) = QuantityThings
                                
                           Else
                                ArrDoc(arrRow, 6) = QuantityThings
                           End If
'
                           ArrDoc(arrRow, 8) = FromUnit
                           ArrDoc(arrRow, 9) = toUnit
                           ArrDoc(arrRow, 10) = fromMRP
                           ArrDoc(arrRow, 11) = toMRP
                           ArrDoc(arrRow, 12) = NumberBasis
                           ArrDoc(arrRow, 13) = DateDoc
                           ArrDoc(arrRow, 14) = Number
                           ArrDoc(arrRow, 15) = NumberClear
                           ArrDoc(arrRow, 16) = Note
                           
                           ArrDoc(arrRow, 17) = Executor
                           ArrDoc(arrRow, 18) = ItemNumber
                           
'                           ArrDoc(arrRow, 19) = TimeInv
                           ArrDoc(arrRow, 20) = price
                           ArrDoc(arrRow, 21) = Service
                                                      
                           ArrDoc(arrRow, 22) = Note1
                           ArrDoc(arrRow, 23) = Note2
                           ArrDoc(arrRow, 24) = Note3
                           ArrDoc(arrRow, 25) = typeDoc
                           ArrDoc(arrRow, 26) = Content
                           ArrDoc(arrRow, 27) = СategoryThingsTo
                           
                        arrRow = arrRow + 1
'                    End If
  
                    
                Next DocRow
   
   
   
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   
   
   Case "ActChangeState81_14"
       

       Number = Trim(wsActualDoc.Range("NumberActChangeState81_14").value)
       
       On Error Resume Next
            
                DocExist = FindValue(Number, RangeToSearch)
                
                If DocExist Then
                    
                    msgExist = MsgBox("Документ: " & Number & " вже рознесений в таблицю. Все одно рознести?", vbYesNo, "Документ з таким номером є в базі")
                        
                        If msgExist = 7 Then
                            GoTo nextDoc
                        End If
                        
                End If
        On Error GoTo 0
       
       Service = Trim(wsActualDoc.Range("ServiceActChangeState81_14"))
                
       Content = Trim(wsActualDoc.Range("ContentActChangeState81_14"))
        'додатковий тип операції, для донесення
       ContentAdd = Trim(wsActualDoc.Range("ContentAddActChangeState81_14"))
                
       typeDoc = Trim(wsActualDoc.Range("typeDocActChangeState81_14"))
       
       InputDate = Date 'дата введення
                Executor = wsActualDoc.Range("ExecutorActChangeState81_14") ' виконавець (якщо тут помилка, скоріше всього клітинці
                'не присвоєне ім'я Exsecutor. Скопіювати Exsecutor, вставити в назву клітинки з виконавцем, натиснути Enter
                
                FromUnit = Trim(wsActualDoc.Range("FromUnitActChangeState81_14")) 'відправник
                
                toUnit = Trim(wsActualDoc.Range("toUnitActChangeState81_14")) 'одержувач

                On Error Resume Next
                fromMRP = Trim(wsActualDoc.Range("MRPActChangeState81_14"))
'                МВО одержувач той же, бо зміна якісного стану відбувається в межах одного підрозділу
                On Error GoTo 0
                toMRP = Trim(wsActualDoc.Range("MRPActChangeState81_14")) 'одержувач

                

                DateDoc = wsActualDoc.Range("DateActChangeState81_14") 'дата документа
                NumberBasis = wsActualDoc.Range("basisActChangeState81_14")
            
            
            
            Set tbActual = wsActualDoc.ListObjects("tbActChangeState81_14")
            
          
                ArrSizeX = tbAllData.ListColumns.count
                ArrSizeY = 0
                
                            
'                        ArrSizeY = tbActual.DataBodyRange.Rows.count
'кількість рядків оприбуткування і списання може бути різною, тому перевіряємо чи є позиції в частині списання і оприбуткування
                    ArrSizeY = Application.WorksheetFunction.CountA(tbActual.ListColumns(18).DataBodyRange) + Application.WorksheetFunction.CountA(tbActual.ListColumns(11).DataBodyRange)
                            
                            ReDim ArrDoc(1 To ArrSizeY, 1 To ArrSizeX)
                            
                            
                arrRow = 1 ' лічильник рядків масиву
                 
                For Each DocRow In tbActual.DataBodyRange.Rows
                wsActualDoc.Activate
                    If Cells(DocRow.Row, 18).value <> "" Then
                    
                          NameThings = Cells(DocRow.Row, 18).value     'найменування
                          MensureThings = Cells(DocRow.Row, 4).value 'одиниця вимцру
                          
                          QuantityThings = Cells(DocRow.Row, 6).value 'кількість
                          ItemNumber = Cells(DocRow.Row, 3).value 'номенклатурний номер
                          price = Cells(DocRow.Row, 7).value 'ціна
                          Note = Cells(DocRow.Row, 19).value
                          Note1 = ContentAdd
                          Note2 = Cells(DocRow.Row, 23).value
                          Note3 = Cells(DocRow.Row, 24).value
                         
                          
                          СategoryThings = Cells(DocRow.Row, 5).value 'додаткове поле
                          СategoryThingsTo = Cells(DocRow.Row, 21).value 'додаткове поле
                           
                           ArrDoc(arrRow, 1) = InputDate
                           ArrDoc(arrRow, 2) = NameThings
                           ArrDoc(arrRow, 4) = MensureThings
                           ArrDoc(arrRow, 5) = СategoryThings
                           
                           If Content = "вибуття" Or Content = "внутрішнє переміщення" Then
                                ArrDoc(arrRow, 7) = QuantityThings
                                
                           Else
                                ArrDoc(arrRow, 6) = QuantityThings
                           End If
'
                           ArrDoc(arrRow, 8) = FromUnit
                           ArrDoc(arrRow, 9) = toUnit
                           ArrDoc(arrRow, 10) = fromMRP
                           ArrDoc(arrRow, 11) = toMRP
                           ArrDoc(arrRow, 12) = NumberBasis
                           ArrDoc(arrRow, 13) = DateDoc
                           ArrDoc(arrRow, 14) = Number
                           ArrDoc(arrRow, 15) = NumberClear
                           ArrDoc(arrRow, 16) = Note
                           
                           ArrDoc(arrRow, 17) = Executor
                           ArrDoc(arrRow, 18) = ItemNumber
                           
'                           ArrDoc(arrRow, 19) = TimeInv
                           ArrDoc(arrRow, 20) = price
                           ArrDoc(arrRow, 21) = Service
                                                      
                           ArrDoc(arrRow, 22) = Note1
                           ArrDoc(arrRow, 23) = Note2
                           ArrDoc(arrRow, 24) = Note3
                           ArrDoc(arrRow, 25) = typeDoc
                           ArrDoc(arrRow, 26) = Content
                           ArrDoc(arrRow, 27) = СategoryThingsTo
                           
                           arrRow = arrRow + 1
                    End If
  
                    
                Next DocRow
                               
                
'Оприбуткування:
                
       Service = Trim(wsActualDoc.Range("ServiceInActChangeState81_14"))
                
       Content = Trim(wsActualDoc.Range("ContentInActChangeState81_14"))
        'додатковий тип операції, для донесення
       ContentAdd = Trim(wsActualDoc.Range("ContentAddInActChangeState81_14"))
                
       FromUnit = Trim(wsActualDoc.Range("FromUnitInActChangeState81_14")) 'відправник
                
       toUnit = Trim(wsActualDoc.Range("FromUnitActChangeState81_14")) 'одержувач
                
'            'кількість рядків оприбуткування і списання може бути різною, тому перевіряємо чи є позиції в частині оприбуткування
'           ArrSizeY = Application.WorksheetFunction.CountA(tbActual.ListColumns(11).DataBodyRange)
'
'                            ReDim ArrDoc(1 To ArrSizeY, 1 To ArrSizeX)
'
'
'                arrRow = 1 ' лічильник рядків масиву
                 
                For Each DocRow In tbActual.DataBodyRange.Rows
                wsActualDoc.Activate
                    If Cells(DocRow.Row, 11).value <> "" Then
                    
                          NameThings = Cells(DocRow.Row, 11).value     'найменування
                          MensureThings = Cells(DocRow.Row, 13).value 'одиниця вимцру
                          
                          QuantityThings = Cells(DocRow.Row, 15).value 'кількість
                          ItemNumber = Cells(DocRow.Row, 12).value 'номенклатурний номер
'
                          price = Cells(DocRow.Row, 16).value 'ціна
                          Note = Cells(DocRow.Row, 20).value
                          Note1 = ContentAdd
                          Note2 = Cells(DocRow.Row, 25).value
                          Note3 = Cells(DocRow.Row, 26).value
                         
                          
                          СategoryThings = Cells(DocRow.Row, 14).value
                          СategoryThingsTo = Cells(DocRow.Row, 22).value
                           
                           ArrDoc(arrRow, 1) = InputDate
                           ArrDoc(arrRow, 2) = NameThings
                           ArrDoc(arrRow, 4) = MensureThings
                           ArrDoc(arrRow, 5) = СategoryThings
                           
                           If Content = "вибуття" Or Content = "внутрішнє переміщення" Then
                                ArrDoc(arrRow, 7) = QuantityThings
                                
                           Else
                                ArrDoc(arrRow, 6) = QuantityThings
                           End If
'
                           ArrDoc(arrRow, 8) = FromUnit
                           ArrDoc(arrRow, 9) = toUnit
                           ArrDoc(arrRow, 10) = fromMRP
                           ArrDoc(arrRow, 11) = toMRP
                           ArrDoc(arrRow, 12) = NumberBasis
                           ArrDoc(arrRow, 13) = DateDoc
                           ArrDoc(arrRow, 14) = Number
                           ArrDoc(arrRow, 15) = NumberClear
                           ArrDoc(arrRow, 16) = Note
                           
                           ArrDoc(arrRow, 17) = Executor
                           ArrDoc(arrRow, 18) = ItemNumber
                           
'                           ArrDoc(arrRow, 19) = TimeInv
                           ArrDoc(arrRow, 20) = price
                           ArrDoc(arrRow, 21) = Service
                                                      
                           ArrDoc(arrRow, 22) = Note1
                           ArrDoc(arrRow, 23) = Note2
                           ArrDoc(arrRow, 24) = Note3
                           ArrDoc(arrRow, 25) = typeDoc
                           ArrDoc(arrRow, 26) = Content
                           ArrDoc(arrRow, 27) = СategoryThingsTo
                           
                           arrRow = arrRow + 1
                    End If
  
                    
                Next DocRow
        
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
       Case "ActQState81_1"
       

       Number = Trim(wsActualDoc.Range("NumberActQState81_1").value)
       
       On Error Resume Next
            
                DocExist = FindValue(Number, RangeToSearch)
                
                If DocExist Then
                    
                    msgExist = MsgBox("Документ: " & Number & " вже рознесений в таблицю. Все одно рознести?", vbYesNo, "Документ з таким номером є в базі")
                        
                        If msgExist = 7 Then
                            GoTo nextDoc
                        End If
                        
                End If
        On Error GoTo 0
       
       Service = Trim(wsActualDoc.Range("ServiceActQState81_1"))
                
       Content = Trim(wsActualDoc.Range("ContentActQState81_1"))
        'додатковий тип операції, для донесення
       ContentAdd = Trim(wsActualDoc.Range("ContentAddActQState81_1"))
                
       typeDoc = Trim(wsActualDoc.Range("typeDocActQState81_1"))
       
       InputDate = Date 'дата введення
                Executor = wsActualDoc.Range("ExecutorActQState81_1") ' виконавець (якщо тут помилка, скоріше всього клітинці
                'не присвоєне ім'я Exsecutor. Скопіювати Exsecutor, вставити в назву клітинки з виконавцем, натиснути Enter
                
                FromUnit = Trim(wsActualDoc.Range("FromUnitActQState81_1")) 'відправник
                
                toUnit = Trim(wsActualDoc.Range("toUnitActQState81_1")) 'одержувач

                On Error Resume Next
                fromMRP = Trim(wsActualDoc.Range("MRPActQState81_1"))
'                МВО одержувач той же, бо зміна якісного стану відбувається в межах одного підрозділу
                On Error GoTo 0
                toMRP = Trim(wsActualDoc.Range("MRPActQState81_1")) 'одержувач

                

                DateDoc = wsActualDoc.Range("DateActQState81_1") 'дата документа
                NumberBasis = wsActualDoc.Range("basisActQState81_1")
            
            
            
            Set tbActual = wsActualDoc.ListObjects("tbActQState81_1")
            
          
                ArrSizeX = tbAllData.ListColumns.count
                ArrSizeY = 0
                
                            
'                        ArrSizeY = tbActual.DataBodyRange.Rows.count
'кількість рядків оприбуткування і списання може бути різною, тому перевіряємо чи є позиції в частині списання і оприбуткування
                    ArrSizeY = Application.WorksheetFunction.CountA(tbActual.ListColumns(18).DataBodyRange) + Application.WorksheetFunction.CountA(tbActual.ListColumns(11).DataBodyRange)
                            
                            ReDim ArrDoc(1 To ArrSizeY, 1 To ArrSizeX)
                            
                            
                arrRow = 1 ' лічильник рядків масиву
                 
                For Each DocRow In tbActual.DataBodyRange.Rows
                wsActualDoc.Activate
                    If Cells(DocRow.Row, 18).value <> "" Then
                    
                          NameThings = Cells(DocRow.Row, 18).value     'найменування
                          MensureThings = Cells(DocRow.Row, 4).value 'одиниця вимцру
                          
                          QuantityThings = Cells(DocRow.Row, 6).value 'кількість
                          ItemNumber = Cells(DocRow.Row, 3).value 'номенклатурний номер
                          price = Cells(DocRow.Row, 7).value 'ціна
                          Note = Cells(DocRow.Row, 19).value
                          Note1 = ContentAdd
                          Note2 = Cells(DocRow.Row, 23).value
                          Note3 = Cells(DocRow.Row, 24).value
                         
                          
                          СategoryThings = Cells(DocRow.Row, 5).value 'додаткове поле
                          СategoryThingsTo = Cells(DocRow.Row, 21).value 'додаткове поле
                           
                           ArrDoc(arrRow, 1) = InputDate
                           ArrDoc(arrRow, 2) = NameThings
                           ArrDoc(arrRow, 4) = MensureThings
                           ArrDoc(arrRow, 5) = СategoryThings
                           
                           If Content = "вибуття" Or Content = "внутрішнє переміщення" Then
                                ArrDoc(arrRow, 7) = QuantityThings
                                
                           Else
                                ArrDoc(arrRow, 6) = QuantityThings
                           End If
'
                           ArrDoc(arrRow, 8) = FromUnit
                           ArrDoc(arrRow, 9) = toUnit
                           ArrDoc(arrRow, 10) = fromMRP
                           ArrDoc(arrRow, 11) = toMRP
                           ArrDoc(arrRow, 12) = NumberBasis
                           ArrDoc(arrRow, 13) = DateDoc
                           ArrDoc(arrRow, 14) = Number
                           ArrDoc(arrRow, 15) = NumberClear
                           ArrDoc(arrRow, 16) = Note
                           
                           ArrDoc(arrRow, 17) = Executor
                           ArrDoc(arrRow, 18) = ItemNumber
                           
'                           ArrDoc(arrRow, 19) = TimeInv
                           ArrDoc(arrRow, 20) = price
                           ArrDoc(arrRow, 21) = Service
                                                      
                           ArrDoc(arrRow, 22) = Note1
                           ArrDoc(arrRow, 23) = Note2
                           ArrDoc(arrRow, 24) = Note3
                           ArrDoc(arrRow, 25) = typeDoc
                           ArrDoc(arrRow, 26) = Content
                           ArrDoc(arrRow, 27) = СategoryThingsTo
                           
                           arrRow = arrRow + 1
                    End If
  
                    
                Next DocRow
                               
                
'Оприбуткування:
                
       Service = Trim(wsActualDoc.Range("ServiceInActQState81_1"))
                
       Content = Trim(wsActualDoc.Range("ContentInActQState81_1"))
        'додатковий тип операції, для донесення
       ContentAdd = Trim(wsActualDoc.Range("ContentAddInActQState81_1"))
                
       FromUnit = Trim(wsActualDoc.Range("FromUnitInActQState81_1")) 'відправник
                
       toUnit = Trim(wsActualDoc.Range("FromUnitActQState81_1")) 'одержувач
                
'            'кількість рядків оприбуткування і списання може бути різною, тому перевіряємо чи є позиції в частині оприбуткування
'           ArrSizeY = Application.WorksheetFunction.CountA(tbActual.ListColumns(11).DataBodyRange)
'
'                            ReDim ArrDoc(1 To ArrSizeY, 1 To ArrSizeX)
'
'
'                arrRow = 1 ' лічильник рядків масиву
                 
                For Each DocRow In tbActual.DataBodyRange.Rows
                wsActualDoc.Activate
                    If Cells(DocRow.Row, 11).value <> "" Then
                    
                          NameThings = Cells(DocRow.Row, 11).value     'найменування
                          MensureThings = Cells(DocRow.Row, 13).value 'одиниця вимцру
                          
                          QuantityThings = Cells(DocRow.Row, 15).value 'кількість
                          ItemNumber = Cells(DocRow.Row, 12).value 'номенклатурний номер
'
                          price = Cells(DocRow.Row, 16).value 'ціна
                          Note = Cells(DocRow.Row, 20).value
                          Note1 = ContentAdd
                          Note2 = Cells(DocRow.Row, 25).value
                          Note3 = Cells(DocRow.Row, 26).value
                         
                          
                          СategoryThings = Cells(DocRow.Row, 14).value
                          СategoryThingsTo = Cells(DocRow.Row, 22).value
                           
                           ArrDoc(arrRow, 1) = InputDate
                           ArrDoc(arrRow, 2) = NameThings
                           ArrDoc(arrRow, 4) = MensureThings
                           ArrDoc(arrRow, 5) = СategoryThings
                           
                           If Content = "вибуття" Or Content = "внутрішнє переміщення" Then
                                ArrDoc(arrRow, 7) = QuantityThings
                                
                           Else
                                ArrDoc(arrRow, 6) = QuantityThings
                           End If
'
                           ArrDoc(arrRow, 8) = FromUnit
                           ArrDoc(arrRow, 9) = toUnit
                           ArrDoc(arrRow, 10) = fromMRP
                           ArrDoc(arrRow, 11) = toMRP
                           ArrDoc(arrRow, 12) = NumberBasis
                           ArrDoc(arrRow, 13) = DateDoc
                           ArrDoc(arrRow, 14) = Number
                           ArrDoc(arrRow, 15) = NumberClear
                           ArrDoc(arrRow, 16) = Note
                           
                           ArrDoc(arrRow, 17) = Executor
                           ArrDoc(arrRow, 18) = ItemNumber
                           
'                           ArrDoc(arrRow, 19) = TimeInv
                           ArrDoc(arrRow, 20) = price
                           ArrDoc(arrRow, 21) = Service
                                                      
                           ArrDoc(arrRow, 22) = Note1
                           ArrDoc(arrRow, 23) = Note2
                           ArrDoc(arrRow, 24) = Note3
                           ArrDoc(arrRow, 25) = typeDoc
                           ArrDoc(arrRow, 26) = Content
                           ArrDoc(arrRow, 27) = СategoryThingsTo
                           
                           arrRow = arrRow + 1
                    End If
  
                    
                Next DocRow
        
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
     '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
       Case "ActPlacement440_45"
       

       Number = Trim(wsActualDoc.Range("NumberActPlacement440_45").value)
       
       On Error Resume Next
            
                DocExist = FindValue(Number, RangeToSearch)
                
                If DocExist Then
                    
                    msgExist = MsgBox("Документ: " & Number & " вже рознесений в таблицю. Все одно рознести?", vbYesNo, "Документ з таким номером є в базі")
                        
                        If msgExist = 7 Then
                            GoTo nextDoc
                        End If
                        
                End If
        On Error GoTo 0
       
       Service = Trim(wsActualDoc.Range("ServiceActPlacement440_45"))
                
       
                
       typeDoc = Trim(wsActualDoc.Range("typeDocActPlacement440_45"))
       
       InputDate = Date 'дата введення
                Executor = wsActualDoc.Range("ExecutorActPlacement440_45") ' виконавець (якщо тут помилка, скоріше всього клітинці
                'не присвоєне ім'я Exsecutor. Скопіювати Exsecutor, вставити в назву клітинки з виконавцем, натиснути Enter
                
                

                On Error Resume Next
                fromMRP = Trim(wsActualDoc.Range("MRPActPlacement440_45"))
'                МВО одержувач той же, бо зміна якісного стану відбувається в межах одного підрозділу
                On Error GoTo 0
                toMRP = Trim(wsActualDoc.Range("MRPActPlacement440_45")) 'одержувач

                

                DateDoc = wsActualDoc.Range("DateActPlacement440_45") 'дата документа
                NumberBasis = wsActualDoc.Range("basisActPlacement440_45")
            
            
            
            Set tbActual = wsActualDoc.ListObjects("tbActPlacement440_45")
            
          
                ArrSizeX = tbAllData.ListColumns.count
                ArrSizeY = 0
                
                            
'                        ArrSizeY = tbActual.DataBodyRange.Rows.count
''кількість рядків оприбуткування і списання може бути різною, тому перевіряємо чи є позиції в частині списання і оприбуткування
                ArrSizeY = Application.WorksheetFunction.CountA(tbActual.ListColumns(6).DataBodyRange) + Application.WorksheetFunction.CountA(tbActual.ListColumns(12).DataBodyRange)
'
                            ReDim ArrDoc(1 To ArrSizeY, 1 To ArrSizeX)
                            
                            
                arrRow = 1 ' лічильник рядків масиву
                 
                For Each DocRow In tbActual.DataBodyRange.Rows
                wsActualDoc.Activate
'
                    
                 'перевірка закладаємо чи вилучаємо
                 'якщо закладаємо:
                    If Cells(DocRow.Row, 6).value <> "" Then
                    
                          NameThings = Cells(DocRow.Row, 2).value     'найменування
                          MensureThings = Cells(DocRow.Row, 4).value 'одиниця вимцру
                          ItemNumber = Cells(DocRow.Row, 3).value 'номенклатурний номер
                          
                    
                          Content = Trim(wsActualDoc.Range("ContentActPlacement440_45"))
                            'додатковий тип операції, для донесення
                          ContentAdd = Trim(wsActualDoc.Range("ContentAddActPlacement440_45"))
                          
                          FromUnit = Trim(wsActualDoc.Range("FromUnitActPlacement440_45")) 'відправник
                
                          toUnit = Trim(wsActualDoc.Range("toUnitActPlacement440_45")) 'одержувач
                        
                          
                          QuantityThings = Cells(DocRow.Row, 6).value 'кількість
                          
                          price = Cells(DocRow.Row, 7).value 'ціна
                          Note = Cells(DocRow.Row, 10).value
                          Note1 = ContentAdd
                          Note2 = Cells(DocRow.Row, 19).value
                          Note3 = Cells(DocRow.Row, 20).value
                                                   
                          СategoryThings = Cells(DocRow.Row, 5).value 'додаткове поле
                          СategoryThingsTo = Cells(DocRow.Row, 18).value 'додаткове поле
                          
                           ArrDoc(arrRow, 1) = InputDate
                           ArrDoc(arrRow, 2) = NameThings
                           ArrDoc(arrRow, 4) = MensureThings
                           ArrDoc(arrRow, 5) = СategoryThings
                           
                           If Content = "вибуття" Or Content = "внутрішнє переміщення" Then
                                ArrDoc(arrRow, 7) = QuantityThings
                                
                           Else
                                ArrDoc(arrRow, 6) = QuantityThings
                           End If
'
                           ArrDoc(arrRow, 8) = FromUnit
                           ArrDoc(arrRow, 9) = toUnit
                           ArrDoc(arrRow, 10) = fromMRP
                           ArrDoc(arrRow, 11) = toMRP
                           ArrDoc(arrRow, 12) = NumberBasis
                           ArrDoc(arrRow, 13) = DateDoc
                           ArrDoc(arrRow, 14) = Number
                           ArrDoc(arrRow, 15) = NumberClear
                           ArrDoc(arrRow, 16) = Note
                           
                           ArrDoc(arrRow, 17) = Executor
                           ArrDoc(arrRow, 18) = ItemNumber
                           
'                           ArrDoc(arrRow, 19) = TimeInv
                           ArrDoc(arrRow, 20) = price
                           ArrDoc(arrRow, 21) = Service
                                                      
                           ArrDoc(arrRow, 22) = Note1
                           ArrDoc(arrRow, 23) = Note2
                           ArrDoc(arrRow, 24) = Note3
                           ArrDoc(arrRow, 25) = typeDoc
                           ArrDoc(arrRow, 26) = Content
                           ArrDoc(arrRow, 27) = СategoryThingsTo
                           
                           arrRow = arrRow + 1
                    
                    End If
                Next DocRow
                
                'якщо вилучаємо:
                For Each DocRow In tbActual.DataBodyRange.Rows
                wsActualDoc.Activate
'
                   If Cells(DocRow.Row, 12).value <> "" Then
                    
                          
                          NameThings = Cells(DocRow.Row, 2).value     'найменування
                          MensureThings = Cells(DocRow.Row, 4).value 'одиниця вимцру
                          ItemNumber = Cells(DocRow.Row, 3).value 'номенклатурний номер
                          
                    
                    
                    
                          Content = Trim(wsActualDoc.Range("ContentInActPlacement440_45"))
                            'додатковий тип операції, для донесення
                          ContentAdd = Trim(wsActualDoc.Range("ContentAddInActPlacement440_45"))
                          
                          FromUnit = Trim(wsActualDoc.Range("FromUnitInActPlacement440_45")) 'відправник
                
                          toUnit = Trim(wsActualDoc.Range("FromUnitActPlacement440_45")) 'одержувач
                        
                    
                    
                    
                    
                          QuantityThings = Cells(DocRow.Row, 12).value 'кількість
                          
                          price = Cells(DocRow.Row, 13).value 'ціна
                          Note = Cells(DocRow.Row, 16).value
                          Note1 = ContentAdd
                          Note2 = Cells(DocRow.Row, 22).value
                          Note3 = Cells(DocRow.Row, 23).value
                                                   
                          СategoryThings = Cells(DocRow.Row, 11).value 'додаткове поле
                          СategoryThingsTo = Cells(DocRow.Row, 21).value 'додаткове поле
                                     
                           ArrDoc(arrRow, 1) = InputDate
                           ArrDoc(arrRow, 2) = NameThings
                           ArrDoc(arrRow, 4) = MensureThings
                           ArrDoc(arrRow, 5) = СategoryThings
                           
                           If Content = "вибуття" Or Content = "внутрішнє переміщення" Then
                                ArrDoc(arrRow, 7) = QuantityThings
                                
                           Else
                                ArrDoc(arrRow, 6) = QuantityThings
                           End If
'
                           ArrDoc(arrRow, 8) = FromUnit
                           ArrDoc(arrRow, 9) = toUnit
                           ArrDoc(arrRow, 10) = fromMRP
                           ArrDoc(arrRow, 11) = toMRP
                           ArrDoc(arrRow, 12) = NumberBasis
                           ArrDoc(arrRow, 13) = DateDoc
                           ArrDoc(arrRow, 14) = Number
                           ArrDoc(arrRow, 15) = NumberClear
                           ArrDoc(arrRow, 16) = Note
                           
                           ArrDoc(arrRow, 17) = Executor
                           ArrDoc(arrRow, 18) = ItemNumber
                           
'                           ArrDoc(arrRow, 19) = TimeInv
                           ArrDoc(arrRow, 20) = price
                           ArrDoc(arrRow, 21) = Service
                                                      
                           ArrDoc(arrRow, 22) = Note1
                           ArrDoc(arrRow, 23) = Note2
                           ArrDoc(arrRow, 24) = Note3
                           ArrDoc(arrRow, 25) = typeDoc
                           ArrDoc(arrRow, 26) = Content
                           ArrDoc(arrRow, 27) = СategoryThingsTo
                           
                           arrRow = arrRow + 1
                    End If
  
                    
                Next DocRow
                
                
                       
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    Case "Request"

Dim wsAllRequest As Worksheet
Dim tbAllRequests As ListObject
   
    Set wsAllRequest = wbAllData.Worksheets("База заявок")
    Set tbAllRequests = wsAllRequest.ListObjects("tbAllRequests")
    
'Dim wsAccountingBook As Worksheet
'Dim tbAccountingBook As ListObject
'
'    Set wsAccountingBook = wbAllData.Worksheets("Книга обліку д13")
'    Set tbAccountingBook = wsAccountingBook.ListObjects("AccountingBook")
    
'для перевірки чи є в книзі обліку мц, яке розносимо в БД

Set RangeToSearch = tbAllRequests.ListColumns(8).DataBodyRange

wsAllRequest.Activate 'активуємо лист з базою
 '====3.9.18====
On Error Resume Next
wsAllRequest.ShowAllData
On Error GoTo 0
'====3.9.18====
       
       
       
       Number = Trim(wsActualDoc.Range("NumberRequest").value)
       
       
       
       
       On Error Resume Next
            
                DocExist = FindValue(Number, RangeToSearch)
                
                If DocExist Then
                    
                    msgExist = MsgBox("Документ: " & Number & " вже рознесений в таблицю. Все одно рознести?", vbYesNo, "Документ з таким номером є в базі")
                        
                        If msgExist = 7 Then
                            GoTo nextDoc
                        End If
                        
                End If
        On Error GoTo 0
       
'       Service = Trim(wsActualDoc.Range("ServiceRequest"))
                
'       Content = Trim(wsActualDoc.Range("ContentRequest"))
'        'додатковий тип операції, для донесення
'       ContentAdd = Trim(wsActualDoc.Range("ContentAddRequest"))
                
       typeDoc = Trim(wsActualDoc.Range("typeDocRequest"))
       
       InputDate = Date 'дата введення
                Executor = wsActualDoc.Range("ExecutorRequest") ' виконавець (якщо тут помилка, скоріше всього клітинці
                'не присвоєне ім'я Exsecutor. Скопіювати Exsecutor, вставити в назву клітинки з виконавцем, натиснути Enter
                FromUnit = Trim(wsActualDoc.Range("FromUnitRequest")) 'відправник
                
                toUnit = Trim(wsActualDoc.Range("toUnitRequest")) 'одержувач

                On Error Resume Next
'                fromMRP = Trim(wsActualDoc.Range("MRPFromRequest"))
                On Error GoTo 0
                toMRP = Trim(wsActualDoc.Range("MRPRequest")) 'одержувач

                

                DateDoc = wsActualDoc.Range("DateRequest") 'дата документа
    Dim signerRequest As String
                signerRequest = wsActualDoc.Range("signerRequest") 'підписант документа
            
            
            
            Set tbActual = wsActualDoc.ListObjects("tbRequest")
            
          
                ArrSizeX = tbAllRequests.ListColumns.count
                ArrSizeY = 0
                
                            
                        ArrSizeY = tbActual.DataBodyRange.Rows.count
                            
                            
                            ReDim ArrDoc(1 To ArrSizeY, 1 To ArrSizeX)
                            
                            
                arrRow = 1 ' лічильник рядків масиву
                 
                For Each DocRow In tbActual.DataBodyRange.Rows
                wsActualDoc.Activate
'                    If InvRow.value <> "" Then
                          Service = Cells(DocRow.Row, 6).value 'служба забезпечення
                    
                          NameThings = Cells(DocRow.Row, 2).value     'найменування
                          MensureThings = Cells(DocRow.Row, 3).value 'одиниця вимцру
                          
                          QuantityThings = Cells(DocRow.Row, 4).value 'кількість
'                          ItemNumber = Cells(DocRow.Row, 4).value 'номенклатурний номер
'                          NumberBasis = Cells(DocRow.Row, 8).value 'підстава для списання
'                          Price = Cells(DocRow.Row, 8).value 'ціна
                          Note = Cells(DocRow.Row, 5).value
'                          Note1 = ContentAdd
'                          Note2 = Cells(DocRow.Row, 11).value
'                          Note3 = Cells(DocRow.Row, 12).value
                         
                          
'                          СategoryThings = Cells(DocRow.Row, 6).value 'додаткове поле
'                          СategoryThingsTo = Cells(DocRow.Row, 10).value 'додаткове поле
                           
                           ArrDoc(arrRow, 1) = InputDate
                           ArrDoc(arrRow, 2) = NameThings
                           ArrDoc(arrRow, 3) = MensureThings
'                           ArrDoc(arrRow, 5) = СategoryThings
                           
'                           If Content = "вибуття" Or Content = "внутрішнє переміщення" Then
                                ArrDoc(arrRow, 4) = QuantityThings
                                
'                           Else
'                                ArrDoc(arrRow, 6) = QuantityThings
'                           End If
'
                           ArrDoc(arrRow, 6) = FromUnit
                           ArrDoc(arrRow, 7) = toUnit
'                           ArrDoc(arrRow, 10) = fromMRP
                           ArrDoc(arrRow, 11) = toMRP
'                           ArrDoc(arrRow, 12) = NumberBasis
                           ArrDoc(arrRow, 9) = DateDoc
                           ArrDoc(arrRow, 8) = Number
                           ArrDoc(arrRow, 16) = signerRequest
                           ArrDoc(arrRow, 17) = Note
                           
                           ArrDoc(arrRow, 12) = Executor
'                           ArrDoc(arrRow, 18) = ItemNumber
                           
'                           ArrDoc(arrRow, 19) = TimeInv
'                           ArrDoc(arrRow, 20) = Price
                           ArrDoc(arrRow, 18) = Service
                                                      
'                           ArrDoc(arrRow, 22) = Note1
'                           ArrDoc(arrRow, 23) = Note2
'                           ArrDoc(arrRow, 24) = Note3
                           ArrDoc(arrRow, 19) = typeDoc
'                           ArrDoc(arrRow, 26) = Content
'                           ArrDoc(arrRow, 27) = СategoryThingsTo
                           
                        arrRow = arrRow + 1
'                    End If
  
                    
                Next DocRow
                
                
      wsAllRequest.Activate
            
'Dim lastRow As Long
    lastRow = wsAllRequest.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row + 1
'Dim arrTarget
wsAllRequest.Activate
    Set arrTarget = wsAllRequest.Range(Cells(lastRow, 1), Cells(lastRow + UBound(ArrDoc()) - 1, ArrSizeX))
    
    arrTarget.value = ArrDoc
                
GoTo nextDoc:
                
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
   End Select
   

   
   
    wbAllData.Activate
            
'Dim lastRow As Long
    lastRow = wsAllData.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row + 1
'Dim arrTarget
wsAllData.Activate
    Set arrTarget = wsAllData.Range(Cells(lastRow, 1), Cells(lastRow + UBound(ArrDoc()) - 1, ArrSizeX))
    
    arrTarget.value = ArrDoc
    
    'після рознесння доставляємо + накладної спереду
 Dim OldName As String
 Dim OldPath As String
 OldPath = wbActualDoc.Path & "\"
        OldName = wbActualDoc.name
        wbActualDoc.SaveAs OldPath & "+ " & OldName
  Kill OldPath & OldName
    
nextDoc:

wbActualDoc.Close False
sFiles = Dir

Loop 'цикл по файлах

Application.AskToUpdateLinks = True
Application.DisplayAlerts = True

MsgBox "Документи внесені в базу! Рекомендовано оновити залишки", , "Імпортовано"

Call UpdateResiduseDetailedAll

End Sub

