Attribute VB_Name = "RibbonCallbacks"
Option Explicit

Private Sub CallImport(control As IRibbonControl)
   Call UnivImport
End Sub

Private Sub CallManualAdd(control As IRibbonControl)
    Call ManualAdding
End Sub
