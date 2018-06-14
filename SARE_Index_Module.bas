Attribute VB_Name = "SARE_Index_Module"
Public address As Range

Public Sub defineAddress(projectNumber)
    Set address = Cells.Find(What:=projectNumber, After:=[A1], SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
End Sub


Sub AddProject()
        AddProjectForm.Show
End Sub

Sub EditProject()
    FindProjectForm.Show
End Sub

Public Sub loadProjectType(boxName)
     With boxName
        .AddItem "Farmer/Rancher"
        .AddItem "Graduate Student"
        .AddItem "Matching Grants Program"
        .AddItem "Partnership"
        .AddItem "PDP State Program"
        .AddItem "Professional Development Program"
        .AddItem "Research and Education"
        .AddItem "Sustainable Community Innovation"
    End With
End Sub

Public Sub loadRegionBox(boxName)
        With boxName
        .AddItem "North Central"
        .AddItem "Northeast"
        .AddItem "Southern"
        .AddItem "Western"
    End With
End Sub

Public Sub loadYearRange(boxName)
    
    Dim intStartYear  As Integer
    Dim intEndYear   As Integer
 
    intEndYear = 2030
    boxName.Clear
    intStartYear = 1989
    
    Do While intStartYear <= intEndYear
        boxName.AddItem intStartYear
        intStartYear = intStartYear + 1
    Loop
End Sub
Public Sub loadStates(boxName)
    With boxName
        .AddItem "Alabama"
        .AddItem "Alaska"
        .AddItem "Arizona"
        .AddItem "Arkansas"
        .AddItem "California"
        .AddItem "Colorado"
        .AddItem "Connecticut"
        .AddItem "District of Columbia"
        .AddItem "Delaware"
        .AddItem "Florida"
        .AddItem "Georgia"
        .AddItem "Hawaii"
        .AddItem "Idaho"
        .AddItem "Illinois"
        .AddItem "Indiana"
        .AddItem "Iowa"
        .AddItem "Kansas"
        .AddItem "Kentucky"
        .AddItem "Louisiana"
        .AddItem "Maine"
        .AddItem "Maryland"
        .AddItem "Massachusetts"
        .AddItem "Michigan"
        .AddItem "Minnesota"
        .AddItem "Mississippi"
        .AddItem "Missouri"
        .AddItem "Montana"
        .AddItem "Nebraska"
        .AddItem "Nevada"
        .AddItem "New Hampshire"
        .AddItem "New Jersey"
        .AddItem "New Mexico"
        .AddItem "New York"
        .AddItem "North Carolina"
        .AddItem "North Dakota"
        .AddItem "Ohio"
        .AddItem "Oklahoma"
        .AddItem "Oregon"
        .AddItem "Pennsylvania"
        .AddItem "Rhode Island"
        .AddItem "South Carolina"
        .AddItem "South Dakota"
        .AddItem "Tennessee"
        .AddItem "Texas"
        .AddItem "Utah"
        .AddItem "Vermont"
        .AddItem "Virginia"
        .AddItem "Washington"
        .AddItem "West Virginia"
        .AddItem "Wisconsin"
        .AddItem "Wyoming"
        
        'Territories
        .AddItem "Guam"
        .AddItem "Puerto Rico"
        .AddItem "Northern Mariana Islands"
        .AddItem "U.S. Virgin Islands"
        .AddItem "Republic of Palau"
        .AddItem "Marshall Islands"
    End With
End Sub

