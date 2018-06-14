VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} OverwriteQuest 
   Caption         =   "UserForm1"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7575
   OleObjectBlob   =   "OverwriteQuest.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "OverwriteQuest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public OverwriteVal As Boolean

Private Sub DoNothing_Click()
    Me.Tag = 0
    Unload Me
End Sub

Private Sub EditExistingProject_Click()
   ' Unload Me
    'Unload AddProjectForm
    Me.Tag = 1
    Me.Hide
   ' EditProjectForm.Show
  End Sub


Private Sub Overwrite_Click()
    Me.Tag = 2
    Me.Hide
End Sub

