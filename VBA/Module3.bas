Attribute VB_Name = "Module3"
Sub ExportAllVBA()
    Dim vbComp As Object
    Dim exportPath As String
    Dim ext As String
    
    ' Створюємо папку VBA_Export поруч із поточним файлом
    exportPath = ThisWorkbook.Path & "\VBA_Export\"
    If Dir(exportPath, vbDirectory) = "" Then MkDir exportPath
    
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Select Case vbComp.Type
            Case 1: ext = ".bas" ' Стандартний модуль
            Case 2: ext = ".cls" ' Модуль класу
            Case 3: ext = ".frm" ' UserForm (також створиться .frx)
            Case 100: ext = ".cls" ' Код аркушів та ThisWorkbook
            Case Else: ext = ""
        End Select
        
        ' Експортуємо лише якщо в модулі є код
        If ext <> "" And vbComp.CodeModule.CountOfLines > 0 Then
            vbComp.Export exportPath & vbComp.name & ext
        End If
    Next vbComp
    
    MsgBox "Готово! Усі скрипти збережено в: " & exportPath, vbInformation
End Sub
