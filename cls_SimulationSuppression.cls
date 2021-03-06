VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_SimulationSuppression"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==========================================================================
'                         FIXTURE SUPPRESSION
'==========================================================================


'This class allows the user to hide or show specified simulation fixtures in specified
'   locations within the operations chain.

'Fixture solid names defined within the simulation parameters will be added as custom properties
'   within the 'Simulation' folder of each park cycle. Altering this setting will determine whether
'   the specified fixture solid will be shown, hidden or no change is made during simulation.

Dim WithEvents EspApp       As Application
Attribute EspApp.VB_VarHelpID = -1
Dim WithEvents MySimul      As Simulation
Attribute MySimul.VB_VarHelpID = -1
Dim IsChangingSimParams     As Boolean

Private Sub Class_Initialize()
    Set EspApp = Application
    Set MySimul = Document.Simulation
End Sub

Private Sub EspApp_BeforePromptDisplayed(ByVal nCmdId As Long, ByVal nSubCmdId As Long, PromptStr As String)
    If nCmdId = 3714 Then
        'Raise flag, user is in simulation params
        IsChangingSimParams = True
    End If
    
    If IsChangingSimParams Then
        If nCmdId = 1001 Then
            'User is exiting simulation params
            'Alter park cycle custom settings
            IsChangingSimParams = False
            UpdateParkSolids
        End If
    End If
End Sub


Private Sub UpdateParkSolids()
    Dim Op          As Esprit.Operation
    Dim Tech        As Technology
    Dim SimuSolid   As Esprit.SimulationSolid
    Dim OpCProps    As CustomProperties
    Dim CProps      As CustomProperties
    Dim CProp       As CustomProperty
    Dim Fixtures    As ArrayList
    Dim i           As Integer
            
    Set Fixtures = New ArrayList
    
    'Loop through each solid within the simulation params
    'Add to fixtures list if item is a fixture
    For i = 1 To Document.Simulation.Count
        Set SimuSolid = Document.Simulation.Item(i)
        If SimuSolid.Type = espSimSolidTypeFixture Then
            Call Fixtures.Add(SimuSolid.Name)
        End If
    Next i
    
    'Find the park cycles
    For Each Op In Document.Operations
        Set Tech = Op.Technology
        If Tech.TechnologyType = espTechLathePark Or Tech.TechnologyType = espTechMillPark Then
            'Create the "Simulation" folder
            Set OpCProps = Op.CustomProperties
            Set CProps = GetCustomPropertyByName(OpCProps, "Simulation", espPropertyTypeCustomProperties, True)
            'Add each  item within the fixtures list as a custom property
            For i = 0 To Fixtures.Count - 1
                'Call GetCustomPropertyByName(CProps, Fixtures.Item(i), espPropertyTypeBoolean, True, True)
                Call CreatePropertyList(CProps, Fixtures.Item(i))
            Next i
            'Clear any fixture properties that may've been removed from the simulation params
            For Each CProp In CProps
                If Not Fixtures.Contains(CProp.Name) Then
                    Call CProps.Remove(CProp.Name)
                End If
            Next CProp
        End If
    Next Op
    End Sub
    
Private Function GetCustomPropertyByName(Properties As CustomProperties, PropertyName As String, PropertyType As espPropertyType, Create As Boolean, Optional DefaultValue As Variant = Nothing) As CustomProperty
    'Attempt to set the specified customproperty
    On Error Resume Next
    Set GetCustomPropertyByName = Properties(PropertyName)
    On Error GoTo 0
    'Get out if property was found
    If Not GetCustomPropertyByName Is Nothing Then Exit Function
    
    If Not Create Then Exit Function
    'Create new customproperty
    Set GetCustomPropertyByName = Properties.Add(PropertyName, PropertyName, PropertyType, DefaultValue)
End Function
    
    
Private Sub MySimul_OnOperationEnd(ByVal Operation As Object)
    Dim Op          As Operation
    Dim Tech        As Technology
    Dim Sp          As Spindle
    Dim MySimu      As FullSimu
    Dim SimuAsm     As SimuAssembly
    Dim custprops   As CustomProperties
    Dim Fixtures    As ISimuFixtures
    Dim i           As Integer
    Dim CustProp    As CustomProperty
    
    Set Op = Operation
    Set Tech = Op.Technology
    
    'determine whether operation is a park cycle
    If Tech.TechnologyType = espTechLathePark Or Tech.TechnologyType = espTechMillPark Then
        Set MySimu = MySimul.CustomSimulationObject
        'Set attributes depending on machine type
        Select Case Document.MachiningMode
            Case Is = espLatheMachiningMode
                Set Sp = Document.LatheMachineSetup.Spindles.Item(CInt(Tech.Item("793")))
                Set MySimu = MySimul.CustomSimulationObject
                Set SimuAsm = MySimu.Assemblies(Sp.Name)
                Set Fixtures = SimuAsm.Axes(SimuAsm.Axes.Count).Child.Fixtures
            Case Is = espMillMachiningMode
                Set SimuAsm = MySimu.Assemblies("Table")
                Set Fixtures = SimuAsm.Fixtures
        End Select
        
        'Get simulation custom properties from the operation
        Set custprops = GetCustomPropertyByName(Op.CustomProperties, "Simulation", espPropertyTypeCustomProperties, False)
        
        'Loop through each custom property to set the visibility
        For i = 1 To Fixtures.Count
            If Not GetCustomPropertyByName(custprops, Fixtures.Item(i).Key, espPropertyTypeBoolean, False) Is Nothing Then
                If Not custprops.Item(Fixtures.Item(i).Key).Value = "No_Change" Then
                    Fixtures.Item(i).Enabled = custprops.Item(Fixtures.Item(i).Key).Value = "Show"
                    Fixtures.Item(i).Visible = custprops.Item(Fixtures.Item(i).Key).Value = "Show"
                End If
            End If
        Next i
        
        Document.Refresh
    End If
    
End Sub

Private Function CreatePropertyList(Properties As CustomProperties, Name As String) As CustomPropertyList
    Dim List            As CustomPropertyList
    Dim ListItem        As CustomPropertyListItem
    Dim CustProp        As CustomProperty
    Dim i               As Integer
    
    Set List = New CustomPropertyList
    
    For i = 1 To 3
        Set ListItem = New CustomPropertyListItem
        Select Case i
        Case 1
            ListItem.Value = "Show"
        Case 2
            ListItem.Value = "Hide"
        Case 3
            ListItem.Value = "No_Change"
        End Select
        ListItem.Caption = ListItem.Value
        Call List.Add(ListItem)
    Next i
    
    Set CustProp = Properties.AddPropertyList(Name, Name, List)
    CustProp.Value = "No_Change"
End Function




