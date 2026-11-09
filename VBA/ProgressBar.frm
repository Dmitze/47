VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressBar 
   Caption         =   "Progress"
   ClientHeight    =   384
   ClientLeft      =   90
   ClientTop       =   360
   ClientWidth     =   8010
   OleObjectBlob   =   "ProgressBar.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Function Progress(Progress_percentage As Single, Optional Unload_After_Completion As Boolean = True)
 Me.show False
 
 VBA.DoEvents
 
 Dim Total_With, Current_W As Single
 
 Total_With = 400
 
 Current_W = (Total_With / 100) * Progress_percentage
 
  Me.lbl_Progress_bar.Width = Current_W
  Me.lbl_Value.Caption = Format(Progress_percentage, "0") & "%"
  
  If Unload_After_Completion = True And Progress_percentage = 100 Then Unload Me
End Function


