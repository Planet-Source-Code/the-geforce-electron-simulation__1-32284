VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form EngineConfigurationFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Engine Configuration"
   ClientHeight    =   3110
   ClientLeft      =   1730
   ClientTop       =   2100
   ClientWidth     =   5150
   Icon            =   "EngineConfigurationFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3110
   ScaleWidth      =   5150
   Begin VB.Frame RepulsionFrame 
      Caption         =   "Repulsion Settings"
      Height          =   1575
      Left            =   288
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   4575
      Begin VB.TextBox FrictionTxt 
         Height          =   375
         Left            =   3120
         TabIndex        =   10
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox SpeedTxt 
         Height          =   375
         Left            =   840
         TabIndex        =   9
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox NumParticlesTxt 
         Height          =   375
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "1"
         Top             =   600
         Width           =   380
      End
      Begin MSComCtl2.UpDown NumParticlesUpDown 
         Height          =   370
         Left            =   2300
         TabIndex        =   11
         Top             =   600
         Width           =   260
         _ExtentX        =   459
         _ExtentY        =   653
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "NumParticlesTxt"
         BuddyDispid     =   196623
         OrigLeft        =   2280
         OrigTop         =   600
         OrigRight       =   2505
         OrigBottom      =   975
         Max             =   1000
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Friction"
         Height          =   195
         Left            =   3090
         TabIndex        =   14
         Top             =   480
         Width           =   525
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Speed"
         Height          =   195
         Left            =   840
         TabIndex        =   13
         Top             =   480
         Width           =   465
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Number of particles"
         Height          =   195
         Left            =   1560
         TabIndex        =   12
         Top             =   240
         Width           =   1365
      End
   End
   Begin VB.CommandButton Done 
      Caption         =   "Done"
      Height          =   372
      Left            =   2022
      TabIndex        =   1
      Top             =   2520
      Width           =   1092
   End
   Begin VB.Frame VideoFrame 
      Caption         =   "Video"
      Height          =   1680
      Left            =   468
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   4215
      Begin VB.CheckBox WindowedBox 
         Caption         =   "Windowed"
         Height          =   252
         Left            =   1560
         TabIndex        =   6
         Top             =   1200
         Width           =   1092
      End
      Begin VB.ComboBox ResolutionBox 
         Height          =   280
         ItemData        =   "EngineConfigurationFrm.frx":000C
         Left            =   1200
         List            =   "EngineConfigurationFrm.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   720
         Width           =   1812
      End
      Begin VB.ComboBox AdapterBox 
         Height          =   280
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   1812
      End
      Begin VB.ComboBox DeviceTypeBox 
         Height          =   280
         ItemData        =   "EngineConfigurationFrm.frx":0010
         Left            =   2280
         List            =   "EngineConfigurationFrm.frx":0012
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   1812
      End
   End
   Begin MSComctlLib.TabStrip ConfigTabs 
      Height          =   3012
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5312
      _ExtentX        =   9366
      _ExtentY        =   5309
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Video"
            Object.ToolTipText     =   "Wildfire 3D Engine Graphics Settings"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Repulsion 3D"
            Object.ToolTipText     =   "Repulsion 3D Engine Settings"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "EngineConfigurationFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ScreenMdeChange As Boolean
Private CurrentState As InitDeviceData

Private Sub AdapterBox_Validate(Cancel As Boolean)
    If CurrentState.AdapterIndex <> AdapterBox.ListIndex Then
        CurrentState.AdapterIndex = AdapterBox.ListIndex
        UpdateDeviceTypeInfo CurrentState.AdapterIndex
        UpdateResolutionInfo CurrentState.AdapterIndex, CurrentState.DeviceType
    End If
End Sub

Private Sub ConfigTabs_Click()
    If ConfigTabs.SelectedItem.Index = 1 Then
        VideoFrame.Visible = True
    Else
        VideoFrame.Visible = False
    End If
    If ConfigTabs.SelectedItem.Index = 2 Then
        RepulsionFrame.Visible = True
    Else
        RepulsionFrame.Visible = False
    End If
End Sub

Private Sub DeviceTypeBox_Validate(Cancel As Boolean)
    If CurrentState.DeviceType <> DeviceTypeBox.ItemData(DeviceTypeBox.ListIndex) Then
        CurrentState.DeviceType = DeviceTypeBox.ItemData(DeviceTypeBox.ListIndex)
        UpdateResolutionInfo CurrentState.AdapterIndex, CurrentState.DeviceType
    End If
End Sub

Private Sub Done_Click()
    Dim InitData As InitDeviceData
    SaveData
    Me.Visible = False
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorH
    Dim Answer As String
    Open "Repulsion.ini" For Input As #1
        Input #1, Answer
        NumParticlesTxt.Text = Answer
        Input #1, Answer
        SpeedTxt.Text = Answer
        Input #1, Answer
        FrictionTxt.Text = Answer
    Close
    Dim DataCls As New DataClass
    Dim InTxt As String
    UpdateSize
    UpdateAdapterInfo
    UpdateDeviceTypeInfo 0
    UpdateResolutionInfo 0, DeviceTypeBox.ItemData(DeviceTypeBox.ListIndex)
    ConfigTabs_Click
    CurrentState = LoadInitDeviceData
    AdapterBox.ListIndex = CurrentState.AdapterIndex
    If DeviceTypeBox.ListCount > 1 Then
        DeviceTypeBox.ListIndex = CurrentState.DeviceType - 1
    End If
    ResolutionBox.ListIndex = CurrentState.Resolution
    If CurrentState.Windowed = True Then
        WindowedBox.Value = 1
    Else
        WindowedBox.Value = 0
    End If
    ScreenMdeChange = False
    
ErrorH:
    If Err = 53 Then
        SaveData
    End If
End Sub

Public Sub UpdateAdapterInfo()
    AdapterBox.Clear
    Dim i As Integer
    For i = 0 To Event3DEngine.Devices(1).NumAdapters - 1
        With AdaptersInfo(i)
            AdapterBox.AddItem .Description
            AdapterBox.ListIndex = 0
        End With
    Next i
End Sub

Public Sub UpdateDeviceTypeInfo(AdapterIndex As Long)
    DeviceTypeBox.Clear
    With AdaptersInfo(AdapterIndex)
        If .DevTypeInfo(1).NumModes > 0 Then
            DeviceTypeBox.AddItem "HAL"
            DeviceTypeBox.ItemData(DeviceTypeBox.ListCount - 1) = 1
        End If
        DeviceTypeBox.AddItem "REF"
        DeviceTypeBox.ItemData(DeviceTypeBox.ListCount - 1) = 2
        DeviceTypeBox.ListIndex = 0
    End With
End Sub

Public Sub UpdateResolutionInfo(AdapterIndex As Long, DeviceType As Long)
    Dim i As Integer
    ResolutionBox.Clear
    With AdaptersInfo(AdapterIndex).DevTypeInfo(DeviceType)
        For i = 0 To .NumModes - 1
            With .Modes(i)
                ResolutionBox.AddItem .Width & " X " & .Height & " X " & .BPP
            End With
        Next i
    End With
    ResolutionBox.ListIndex = 0
End Sub

Public Sub SaveData()
    Open "InitDevice.ini" For Output As #1
        If WindowedBox.Value < 1 Then
            Print #1, ".InitDevice: " & AdapterBox.ListIndex & "," & DeviceTypeBox.ItemData(DeviceTypeBox.ListIndex) & "," & ResolutionBox.ListIndex & "," & "False"
        Else
            Print #1, ".InitDevice: " & AdapterBox.ListIndex & "," & DeviceTypeBox.ItemData(DeviceTypeBox.ListIndex) & "," & ResolutionBox.ListIndex & "," & "True"
        End If
    Close
    Open "Repulsion.ini" For Output As #1
        Print #1, NumParticlesTxt.Text
        Print #1, SpeedTxt.Text
        Print #1, FrictionTxt.Text
    Close
End Sub

Public Sub UpdateSize()
    With ConfigTabs
        .Left = 0
        .Top = 0
        .Width = ScaleWidth
        .Height = ScaleHeight
    End With
End Sub

Private Sub Form_Resize()
    UpdateSize
End Sub

Private Sub ResolutionBox_Validate(Cancel As Boolean)
    If CurrentState.Resolution <> ResolutionBox.ListIndex Then
        CurrentState.Resolution = ResolutionBox.ListIndex
        ScreenMdeChange = True
    End If
End Sub

Private Sub WindowedBox_Click()
    ScreenMdeChange = True
End Sub
