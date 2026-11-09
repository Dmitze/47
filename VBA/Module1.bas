Attribute VB_Name = "Module1"
Sub ExportAllVBAElements()
    Dim vbaProject As Object
    Dim vbaComponent As Object
    Dim exportPath As String
    Dim extension As String
    
    ' Путь к папке, куда будут сохранены файлы (в данном случае - папка с текущим файлом Excel)
    exportPath = ActiveWorkbook.Path & "\VBA_Export\"
    
    ' Создаем папку, если она не существует
    If Dir(exportPath, vbDirectory) = "" Then
        MkDir exportPath
    End If
    
    Set vbaProject = ActiveWorkbook.VBProject
    
    ' Перебираем все компоненты в проекте
    For Each vbaComponent In vbaProject.VBComponents
        ' Определяем расширение файла в зависимости от типа компонента
        Select Case vbaComponent.Type
            Case 1 ' Стандартный модуль
                extension = ".bas"
            Case 2 ' Модуль класса
                extension = ".cls"
            Case 3 ' Пользовательская форма (UserForm)
                extension = ".frm"
            Case 100 ' Документы (Листы, ЭтаКнига)
                extension = ".cls"
            Case Else
                extension = ".txt"
        End Select
        
        ' Экспортируем компонент, если в нем есть код
        If vbaComponent.CodeModule.CountOfLines > 0 Then
            vbaComponent.Export exportPath & vbaComponent.name & extension
        End If
    Next vbaComponent
    
    MsgBox "Все VBA-скрипты успешно выгружены в папку:" & vbCrLf & exportPath, vbInformation, "Готово"
End Sub

