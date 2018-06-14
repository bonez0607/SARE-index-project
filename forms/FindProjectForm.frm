VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FindProjectForm 
   Caption         =   "Find Project"
   ClientHeight    =   1590
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   4800
   OleObjectBlob   =   "FindProjectForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FindProjectForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub GetProject_Click()
    Dim projectNumber As String
    
    projectNumber = ProjectNumberA & "-" & ProjectNumberB
     
    defineAddress (projectNumber)
    
     projectFound = Cells.Find(What:="*", After:=[A1], SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row + 1
     If address Is Nothing Then
          alert = MsgBox("Project not found", vbOKOnly)
        Else
        Unload FindProjectForm
        EditProjectForm.Show
    End If
End Sub

