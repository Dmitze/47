VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Аркуш11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)

    Dim lo As ListObject
    Dim cb As CommandBar
    Dim mainBtn As CommandBarControl
    Dim subBtn As CommandBarControl
    Dim i As Integer

    
For Each cell In Selection.Cells
    
    If cell.Row <= 7 Then Exit Sub
    
    If Trim(Me.Cells(cell.Row, "B").value) = "" Then Exit Sub

Next cell

    Cancel = True

    Set cb = Application.CommandBars("Cell")

    'видалити якщо вже є
    On Error Resume Next
    cb.Controls("Додати в документ").Delete
    On Error GoTo 0

    'головне меню
    Set mainBtn = cb.Controls.Add(Type:=msoControlPopup, Temporary:=True)

    mainBtn.Caption = "Додати в документ"

    Set subBtn = mainBtn.Controls.Add(Type:=msoControlButton)
    subBtn.Caption = "Накладна д.25 (440 нак.)"
    subBtn.OnAction = "addFromTableInv440_25"
    
    '=========3.9.26============
    Set subBtn = mainBtn.Controls.Add(Type:=msoControlButton)
    subBtn.Caption = "Розд(З)Відом д57"
    subBtn.OnAction = "addFromTableInv440_57"
    '=========3.9.26============
    
    Set subBtn = mainBtn.Controls.Add(Type:=msoControlButton)
    subBtn.Caption = "Акт списання д.3 (81 нак.)"
    subBtn.OnAction = "addFromTableWOff81"
    
    Set subBtn = mainBtn.Controls.Add(Type:=msoControlButton)
    subBtn.Caption = "Акт зміни якісного стану д14 81 наказ"
    subBtn.OnAction = "addFromTableActChangeState81_14"
    
    Set subBtn = mainBtn.Controls.Add(Type:=msoControlButton)
    subBtn.Caption = "Акт якісного (технічного) стану д1 81 наказ"
    subBtn.OnAction = "addFromTableActQState81_1"
    
    Set subBtn = mainBtn.Controls.Add(Type:=msoControlButton)
    subBtn.Caption = "Акт заклаладання (освіження) д45 440 наказ"
    subBtn.OnAction = "addFromTableActPlacement440_45"
    
    Set subBtn = mainBtn.Controls.Add(Type:=msoControlButton)
    subBtn.Caption = "Заявка/Рапорт на отримання майна"
    subBtn.OnAction = "addFromTableRequest"

    cb.ShowPopup

End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    RemoveCustomContextMenu
End Sub

