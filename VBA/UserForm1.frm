VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Майно"
   ClientHeight    =   3024
   ClientLeft      =   -345
   ClientTop       =   -1665
   ClientWidth     =   5475
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

 Private WithEvents oEventHandler As clsSearchableDropdown
Attribute oEventHandler.VB_VarHelpID = -1

'  Public Property Let ListData(ByVal rg As Range)
'      oEventHandler.List = rg.value
'  End Property
  
  Public Property Let ListData(ByVal v As Variant)
    If TypeName(v) = "Range" Then
        oEventHandler.List = v.value
    ElseIf IsArray(v) Then
        oEventHandler.List = v
    Else
        Err.Raise vbObjectError + 1, , "ListData приймає лише Range або масив"
    End If
End Property
  
  
  
  
  
  Private Sub UserForm_Initialize()

     Set oEventHandler = New clsSearchableDropdown
     With oEventHandler
         Set .SearchListBox = Me.ListBox1
         Set .SearchTextBox = Me.TextBox1

'       Settings
         .MaxRows = 20 ' set the number of items to be displayed
         .ShowAllMatches = True ' to show all the matches
         .CompareMethod = vbTextCompare ' use vbBinaryCompare for case sensitivity
     End With
    
 With Me
    .Zoom = 100
    .Width = 850
    .Height = 390
End With
  End Sub

  Private Sub UserForm_Terminate()
     Set oEventHandler = Nothing
  End Sub
Private Sub CommandButton1_Click()
Me.Hide
End Sub
