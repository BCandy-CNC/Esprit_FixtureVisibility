VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_FixtureVisibilityManager 
   Caption         =   "Fixture Visibility Manager"
   ClientHeight    =   3135
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8505
   OleObjectBlob   =   "frm_FixtureVisibilityManager.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_FixtureVisibilityManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim OpDict      As New Dictionary

'==========================================|
' !DO NOT ASSIGN THESE VARIABLES DIRECTLY!'|
'     !USE GETTERS AND SETTERS ONLY!      '|
'-----------------------------------------'|
Dim pMyOp       As Operation              '|
Dim pMyProps    As CustomProperties       '|
Dim pMyProp     As CustomProperty         '|
Dim UserGroup   As New ArrayList          '|
'========================================='|


'================================================================================='
'=============================      PROPERTIES      =============================='
'================================================================================='


'|------------------------------------------------------------------------------'|
'|                                    MyProps                                   '|
'|This property handles the full simulation customproperties collection         '|
'|CustomProperties are at the Esprit Operation level                            '|
Private Property Get MyProps() As CustomProperties                              '|
    Set MyProps = pMyProps                                                      '|
End Property                                                                    '|
                                                                                '|
Private Property Let MyProps(Props As CustomProperties)                         '|
    Dim i As Integer                                                            '|
    lbo_PropertyName.Clear                                                      '|
                                                                                '|
    If Props Is Nothing Then                                                    '|
        Set pMyProps = Nothing                                                  '|
        MyProp = Nothing                                                        '|
    Else                                                                        '|
        Set pMyProps = Props                                                    '|
        MyProp = Nothing                                                        '|
        For i = 1 To pMyProps.Count                                             '|
            Call lbo_PropertyName.AddItem(pMyProps(i).Caption)                  '|
        Next i                                                                  '|
    End If                                                                      '|
End Property                                                                    '|
'--------------------------------------------------------------------------------|



'|------------------------------------------------------------------------------'|
'|                                    MyProp                                    '|
'|This property handles the user selected property                              '|
Private Property Get MyProp() As CustomProperty                                 '|
    Set MyProp = pMyProp                                                        '|
End Property                                                                    '|
                                                                                '|
Private Property Let MyProp(Prop As CustomProperty)                             '|
    Dim ListItem    As CustomPropertyListItem                                   '|
    Dim i           As Long                                                     '|
    Dim SimuSolid   As Esprit.SimulationSolid                                   '|
                                                                                '|
    cbo_PropertyValue.Clear                                                     '|
                                                                                '|
    Group.Clear                                                                 '|
    Call Group.Add(MyOp)                                                        '|
                                                                                '|
    If Prop Is Nothing Then                                                     '|
        Set pMyProp = Nothing                                                   '|
    Else                                                                        '|
        Set pMyProp = Prop                                                      '|
        For Each ListItem In MyProp.List                                        '|
            Call cbo_PropertyValue.AddItem(ListItem.Caption)                    '|
        Next                                                                    '|
        cbo_PropertyValue.Value = MyProp.Value                                  '|
                                                                                '|
        For i = 1 To Sim.Count                                                  '|
            Set SimuSolid = Sim.Item(i)                                         '|
            If SimuSolid.Name = lbo_PropertyName.Value Then                     '|
                Call Group.Add(SimuSolid.Profile)                               '|
            End If                                                              '|
        Next i                                                                  '|
    End If                                                                      '|
End Property                                                                    '|
'--------------------------------------------------------------------------------|



'|------------------------------------------------------------------------------'|
'|                                      MyOp                                    '|
'|This property handles the user selected operation                             '|
Private Property Get MyOp() As Operation                                        '|
    Set MyOp = pMyOp                                                            '|
End Property                                                                    '|
                                                                                '|
Private Property Let MyOp(Op As Operation)                                      '|
    Group.Clear                                                                 '|
    If Op Is Nothing Then                                                       '|
        Set pMyOp = Nothing                                                     '|
        MyProps = Nothing                                                       '|
        cbo_PropertyValue.Clear                                                 '|
        lbo_PropertyName.Clear                                                  '|
    Else                                                                        '|
        Set pMyOp = Op                                                          '|
        MyProps = Op.CustomProperties("Simulation")                             '|
        Call Group.Add(MyOp)                                                    '|
    End If                                                                      '|
End Property                                                                    '|
 '-------------------------------------------------------------------------------|
 '===============================================================================|



'Initializes the userform
Private Sub UserForm_Initialize()
    CaptureUserGroup
    PopulateOpDict
    PopulateOpList
    Group.Clear
End Sub



'Handles operation change
Private Sub lbo_Operation_Change()
    MyOp = Ops(OpDict.Keys(lbo_Operation.ListIndex))
End Sub



'Handles a property name change
Private Sub lbo_PropertyName_Change()
    If IsControlIndexed(lbo_PropertyName.Name) Then
        MyProp = MyProps(lbo_PropertyName.Text)
    End If
End Sub



'Handles a property value change
Private Sub cbo_PropertyValue_Change()
    If IsControlIndexed(cbo_PropertyValue.Name) Then
        MyProp.Value = cbo_PropertyValue.Text
    End If
End Sub



'Identifies all operations that contain the Simulation custom properties
Private Sub PopulateOpDict()
    Dim Op As Operation
    For Each Op In Ops
        If PropertyExists(Op.CustomProperties, "Simulation") Then
            Call OpDict.Add(Op.Key, Op.Name)
        End If
    Next Op
End Sub



'Populates the operations list from the OpDict dictionary
Private Sub PopulateOpList()
    Dim Key As Variant
    For Each Key In OpDict.Keys
        Call lbo_Operation.AddItem(OpDict(Key))
    Next Key
End Sub



'Captures all items currently grouped
Private Sub CaptureUserGroup()
    Dim i       As Long
    For i = 1 To Group.Count
        Call UserGroup.Add(Group.Item(i))
    Next i
End Sub



'Regroups all items within the UserGroups List
Private Sub ReinstateUserGroup()
    Dim i       As Long
    Group.Clear
    For i = 0 To UserGroup.Count - 1
        Call Group.Add(UserGroup.Item(i))
    Next i
End Sub



'Groups all simulation solids by name
Private Sub GroupSimSolids()
    Dim i           As Long
    Dim SimuSolid   As Esprit.SimulationSolid
    For i = 1 To Sim.Count
        Set SimuSolid = Sim.Item(i)
        If SimuSolid.Name = lbo_PropertyName.Value Then
            Call Group.Add(SimuSolid.Profile)
        End If
    Next i
End Sub



'Returns whether a control with a list has been positioned by the user.
'An index other than -1 must have been changed by the user.
Private Function IsControlIndexed(CtrlName As String) As Boolean
    IsControlIndexed = Not Controls(CtrlName).ListIndex = -1
End Function



'Returns whether a specified CustomProperty exists within a CustomProperties collection
Private Function PropertyExists(Properties As CustomProperties, Property As String) As Boolean
    On Error Resume Next
    PropertyExists = Not Properties(Property) Is Nothing
End Function



'Returns the Esprit Operations collection
Private Function Ops() As Operations
    Set Ops = Doc.Operations
End Function



'Returns the Esprit Document object
Private Function Doc() As Document
    Set Doc = Document
End Function



'Returns the Esprit Group object
Private Function Group() As Group
    Set Group = Doc.Group
End Function



'Returns the Esprit Simulation object
Private Function Sim() As Simulation
    Set Sim = Doc.Simulation
End Function




Private Sub UserForm_Terminate()
    ReinstateUserGroup
End Sub
