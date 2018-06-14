VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EditProjectForm 
   Caption         =   "Edit Project"
   ClientHeight    =   10260
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   9990
   OleObjectBlob   =   "EditProjectForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "EditProjectForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    Sheet1.Activate
    Call loadProjectType(ProjectType)
    Call loadRegionBox(RegionBox)
    Call loadStates(StateBox)
    Call loadYearRange(EndYear)
    ProjectNumber1.SetFocus
    ProjectNumber1.MaxLength = 10
    ProjectNumber2.MaxLength = 10
    ProjectName.MaxLength = 400
    EndYear.MaxLength = 4
    
     Call loadProject(rowNum)
End Sub

Private Sub loadProject(row)
    Dim projectNum() As String
    projectNum = Split(Cells(row, 1).Value, "-")
    
    ProjectNumber1 = projectNum(0)
    ProjectNumber2 = projectNum(1)
    ProjectName.Text = Cells(row, 2).Value
    ProjectType = Cells(row, 3).Value
   
    RegionBox = Cells(row, 4).Value
    StateBox = Cells(row, 5).Value
    EndYear = Cells(row, 6).Value
    
    GrantRecipient = Cells(row, 7).Value
    PrincipalInvestigator = Cells(row, 8).Value
  
    getCheckedPracticesFromSheet (Cells(row, 9).Value)
    getCheckedResourcesFromSheet (Cells(row, 10).Value)

    Link = Cells(rowNum, 11).Value
    SearchTerms = Cells(rowNum, 12).Value
End Sub

Function rowNum() As Integer
     rowNum = address.row
End Function

Private Sub getCheckedPracticesFromSheet(rowValue)
       practices = Split(rowValue, ", ")
       For Each practice In practices
         Select Case practice
            Case "Alley Cropping"
                AlleyCroppingChk.Value = True
            Case "Forest Farming"
                ForestFarmingChk.Value = True
            Case "General"
                GeneralChk.Value = True
            Case "Riparian Forest Buffer"
                RiparianChk = True
            Case "Silvopasture"
                SilvopastureChk.Value = True
            Case "Windbreak"
                WindbreakChk.Value = True
            Case "I don't know"
                DontKnowChk.Value = True
            End Select
    Next practice
End Sub

Private Sub getCheckedResourcesFromSheet(rowValue)
     resources = Split(rowValue, ", ")
        For Each resource In resources
         Select Case resource
            Case "Bulletin"
                BulletinChk.Value = True
            Case "Curriculum"
                CurriculumChk.Value = True
            Case "Fact Sheet"
                FactSheetChk = True
            Case "Manual/Guide"
                GuideChk.Value = True
            Case "Multimedia"
                MultimediaChk.Value = True
            Case Else
                OtherResources.Value = OtherResources & ", " & resource
            End Select
    Next resource
    
     If Left(OtherResources.Value, 2) = ", " Then
        OtherResources.Value = Right(OtherResources, Len(OtherResources) - 2)
    End If
End Sub


Public Sub UpdateProject_Click()
    Cells(rowNum, 1).Value = UCase(ProjectNumber1.Value & "-" & ProjectNumber2.Value)
    Cells(rowNum, 2).Value = ProjectName.Value
    Cells(rowNum, 3).Value = ProjectType.Value
    Cells(rowNum, 4).Value = RegionBox.Value
    Cells(rowNum, 5).Value = StateBox.Value
    Cells(rowNum, 6).Value = EndYear.Value
    Cells(rowNum, 7).Value = GrantRecipient.Value
    Cells(rowNum, 8).Value = PrincipalInvestigator.Value
    Cells(rowNum, 9).Value = getCheckedPractices(0)
    Cells(rowNum, 10).Value = getCheckedResources(0, OtherResources)
    Cells(rowNum, 11).Value = Link.Value
    Cells(rowNum, 12).Value = SearchTerms.Value

    UserForm_Initialize
    Unload Me
End Sub

Private Sub cmdClear_Click()
    UserForm_Initialize
End Sub

Private Sub cmdClose_Click()
    Unload Me
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




