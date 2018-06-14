VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddProjectForm 
   Caption         =   "Add New Project"
   ClientHeight    =   10800
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   9525
   OleObjectBlob   =   "AddProjectForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddProjectForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AddProjectForm_Init()
    Call blankForm
    Call loadProjectType(ProjectType)
    Call loadRegionBox(RegionBox)
    Call loadStates(StateBox)
    Call loadYearRange(EndYear)
    ProjectNumber1.SetFocus
    ProjectNumber1.MaxLength = 10
    ProjectNumber2.MaxLength = 10
    ProjectName.MaxLength = 400
    EndYear.MaxLength = 4
End Sub

Private Sub AddProject_Click()
    Dim emptyRow As Integer
    Dim projectNumber As String
    Dim placeItems As Object

    Sheet1.Activate
    projectNumber = ProjectNumber1.Value & "-" & ProjectNumber2.Value
    defineAddress (projectNumber)
   
   If address Is Nothing Then
        emptyRow = Cells.Find(What:="*", After:=[A1], SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row + 1
        Call addItemsToSheet(projectNumber, emptyRow)

        Else
        OverwriteQuest.Show
        
        Select Case OverwriteQuest.Tag
            Case 1
                Unload AddProjectForm
                EditProjectForm.Show
        
            Case 2
                Dim addressArr() As String
                Dim rowNum As Integer
                addressArr = Split(address.address, "$")
                rowNum = addressArr(UBound(addressArr))
             
                Call addItemsToSheet(projectNumber, rowNum)
    
                Unload Me
       
            Case Else
                'Do Nothing
        End Select
    End If
     Set address = Nothing
    'AddProjectForm_Init
End Sub

Private Sub cmdClear_Click()
    AddProjectForm_Init
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub blankForm()
    ProjectNumber1 = ""
    ProjectNumber2 = ""
    ProjectName = ""
    ProjectType = ""
    RegionBox = ""
    StateBox = ""
    EndYear = ""
    GrantRecipient = ""
    PrincipalInvestigator = ""
    AfPracticeBox = ""
    Link = ""
    OtherResources = ""
    SearchTerms = ""
    
    For i = ProjectType.ListCount - 1 To 0 Step -1
        ProjectType.RemoveItem i
    Next i
    
    For i = RegionBox.ListCount - 1 To 0 Step -1
        RegionBox.RemoveItem i
    Next i
    
     For i = StateBox.ListCount - 1 To 0 Step -1
        StateBox.RemoveItem i
    Next i
    
    'Practices CheckBox Default
    AlleyCroppingChk.Value = False
    ForestFarmingChk.Value = False
    GeneralChk.Value = False
    RiparianChk.Value = False
    SilvopastureChk.Value = False
    WindbreakChk.Value = False
    DontKnowChk.Value = False
    
    'Resources CheckBox Default
    BulletinChk.Value = False
    CurriculumChk.Value = False
    FactSheetChk.Value = False
    GuideChk.Value = False
    MultimediaChk.Value = False
End Sub
Public Function getCheckedPractices(arrLen As Integer)
    Dim practicesArr() As String
    arrLen = 0
    
    If AlleyCroppingChk.Value = True Then
            ReDim Preserve practicesArr(arrLen)
            practicesArr(0) = "Alley Cropping"
            arrLen = arrLen + 1
    End If
    
    If ForestFarmingChk.Value = True Then
        ReDim Preserve practicesArr(arrLen)
        practicesArr(arrLen) = "Forest Farming"
        arrLen = arrLen + 1
        
     End If
    
    If GeneralChk.Value = True Then
        ReDim Preserve practicesArr(arrLen)
        practicesArr(arrLen) = "General"
        arrLen = arrLen + 1
      
    End If
    
    If RiparianChk.Value = True Then
        ReDim Preserve practicesArr(arrLen)
        practicesArr(arrLen) = "Riparian Forest Buffer"
        arrLen = arrLen + 1
        
    End If
    
    If SilvopastureChk.Value = True Then
        ReDim Preserve practicesArr(arrLen)
        practicesArr(arrLen) = "Silvopasture"
        arrLen = arrLen + 1
        
    End If
    
    If WindbreakChk.Value = True Then
        ReDim Preserve practicesArr(arrLen)
        practicesArr(arrLen) = "Windbreak"
        arrLen = arrLen + 1
        
    End If
    
    If DontKnowChk.Value = True Then
        ReDim Preserve practicesArr(arrLen)
        practicesArr(arrLen) = "I don't know"
        arrLen = arrLen + 1
    End If
    
    getCheckedPractices = Join(practicesArr, ", ")
End Function

Public Function getCheckedResources(arrLen As Integer, moreResources As String)
    Dim resourcesArr() As String
    arrLen = 0
    
    If BulletinChk.Value = True Then
            ReDim Preserve resourcesArr(arrLen)
            resourcesArr(0) = "Bulletin"
            arrLen = arrLen + 1
    End If
    
    If CurriculumChk.Value = True Then
        ReDim Preserve resourcesArr(arrLen)
        resourcesArr(arrLen) = "Curriculum"
        arrLen = arrLen + 1
    End If
    
    If FactSheetChk.Value = True Then
        ReDim Preserve resourcesArr(arrLen)
        resourcesArr(arrLen) = "Fact Sheet"
        arrLen = arrLen + 1
      
    End If
    
    If GuideChk.Value = True Then
        ReDim Preserve resourcesArr(arrLen)
        resourcesArr(arrLen) = "Manual/Guide"
        arrLen = arrLen + 1
        
    End If
    
    If MultimediaChk.Value = True Then
        ReDim Preserve resourcesArr(arrLen)
        resourcesArr(arrLen) = "Multimedia"
        arrLen = arrLen + 1
    End If
    
    If Not moreResources = "" Then
        Dim moreResourcesArr() As String
        
        moreResourcesArr = Split(moreResources, ",")
        
        For Each resource In moreResourcesArr
             ReDim Preserve resourcesArr(arrLen)
            resourcesArr(arrLen) = Trim(resource)
            arrLen = arrLen + 1
        Next resource
    End If
    
    getCheckedResources = Join(resourcesArr, ", ")
End Function

'Alerts user when entering Project Name field that the project already exists
Private Sub ProjectName_Enter()
    
    projectNumber = ProjectNumber1.Value & "-" & ProjectNumber2.Value
    defineAddress (projectNumber)
    
     If Not address Is Nothing Then
        Dim answer As String
        answer = MsgBox("This project already exists in spreadsheet. Would you like you like to pull up the existing project?", vbQuestion + vbYesNo, "Project exists error")
        
        If answer = vbNo Then
            'Do nothing
        Else
            Unload Me
            timesShown = 1
            EditProjectForm.Show
        End If
    End If

End Sub

Private Sub addItemsToSheet(projectNumber As String, row As Integer)
    Cells(row, 1).Value = UCase(projectNumber)
    Cells(row, 2).Value = ProjectName.Value
    Cells(row, 3).Value = ProjectType.Value
    Cells(row, 4).Value = RegionBox.Value
    Cells(row, 5).Value = StateBox.Value
    Cells(row, 6).Value = EndYear.Value
    Cells(row, 7).Value = GrantRecipient.Value
    Cells(row, 8).Value = PrincipalInvestigator.Value
    Cells(row, 9).Value = getCheckedPractices(0)
    Cells(row, 10).Value = getCheckedResources(0, OtherResources)
    Cells(row, 11).Value = Link.Value
    Cells(row, 12).Value = SearchTerms.Value
End Sub

Private Sub UserForm_Click()

End Sub
