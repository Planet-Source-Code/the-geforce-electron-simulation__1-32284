VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EventEngineConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Wildfire3DEngine"
Attribute VB_Ext_KEY = "Member1" ,"WildfireEngine"
Attribute VB_Ext_KEY = "Member2" ,"WildfireMusicEngine"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Event Engine by Daniel Bloemendal
'If you take any code from this engine you must give me credit!

'This engine is most likely going to be discontinued.
'This is because I am working on a C++ engine.
'If you would like to continue the production of this engine,
'You must get permission from me at teh_leet_programming_man@hotmail.com
'Thank you for downloading my code please give me feedback.
'Vote for me!!!

Public TargethFocusWindow As Long
Public EventEngine As New EventEngine
Public Event3DEngine As Event3DEngine
Public Event3DEngineCreationPassword As String
Public EventMusicEngine As EventMusicEngine
Public EventMusicEngineCreationPassword As String
Private AdaptersInfo() As AdapterInfo
Public Type InitDeviceData
    hFocusWindow As Long
    AdapterIndex As Long
    DeviceType As Long
    Resolution As Long
    Windowed As Boolean
End Type
Public StatusText As String
Public EchoOff As Boolean
Private CurrentState As InitDeviceData
'To fire this event, use RaiseEvent with the following syntax:
'RaiseEvent EngineEvent[(arg1, arg2, ... , argn)]
Public Event EngineEvent(Message As String)

Private Sub Init3DEngine()
    Set Event3DEngine = Nothing
    AddStatusText "Initializing 3D Engine..." & vbCrLf
    AddStatusText "Adding Instance... "
    Set Event3DEngine = EventEngine.Event3DEngineCreate(Event3DEngineCreationPassword)
    AddStatusText "Done!" & vbCrLf & "3D Engine Initialization Complete!" & vbCrLf
End Sub

Private Sub InitDevice(AdapterIndex As Long, DeviceType As Long, ScrnMde As Long, Windowed As Boolean)
    Dim i As Integer
    CurrentState.hFocusWindow = TargethFocusWindow
    CurrentState.AdapterIndex = AdapterIndex
    CurrentState.DeviceType = DeviceType
    CurrentState.Resolution = ScrnMde
    CurrentState.Windowed = Windowed
    AddStatusText "Adding Device... "
    Event3DEngine.Devices.Add
    AddStatusText "Done!" & vbCrLf
    With Event3DEngine.Devices(Event3DEngine.Devices.Count)
        .hFocusWindow = TargethFocusWindow
        AddStatusText "Detecting Hardware... "
        .DetectHardware
        ReDim AdaptersInfo(.NumAdapters - 1)
        For i = 0 To .NumAdapters - 1
            AdaptersInfo(i) = .GetAdaptersInfo(i)
        Next i
        AddStatusText "Done!" & vbCrLf
        If Windowed = False Then
            With AdaptersInfo(AdapterIndex).DevTypeInfo(DeviceType).Modes(ScrnMde)
                AddStatusText "Initializing Device (" & .Width & " X " & .Height & " X " & .BPP & ")... "
            End With
        Else
            AddStatusText "Initializing Device (Windowed)... "
        End If
        .InitDevice AdapterIndex, DeviceType, Windowed, ScrnMde
        AddStatusText "Done!" & vbCrLf & "Device Initialization Complete!" & vbCrLf
    End With
End Sub

Private Sub ResizeFullScreen(DeviceIndex As Long, ScrnMode As Long)
    CurrentState.Resolution = ScrnMode
    With AdaptersInfo(CurrentState.AdapterIndex).DevTypeInfo(CurrentState.DeviceType).Modes(CurrentState.Resolution)
        AddStatusText "Resizing Full Screen (" & .Width & " X " & .Height & " X " & .BPP & ")... "
    End With
    Event3DEngine.Devices(DeviceIndex).ResizeFullScreen CurrentState.Resolution
    AddStatusText "Done!" & vbCrLf
End Sub

Private Sub ResizeWindowed(DeviceIndex As Long)
    CurrentState.Windowed = True
    AddStatusText "Resizing Windowed... "
    Event3DEngine.Devices(DeviceIndex).ResizeWindowed
    AddStatusText "Done!" & vbCrLf
End Sub

Private Sub AddRenderState(DeviceIndex As Long, State As Long, Value As Long)
    AddStatusText "Adding Render State... "
    Event3DEngine.Devices(DeviceIndex).AddRenderState State, Value
    AddStatusText "Done!" & vbCrLf
End Sub

Private Sub SetRenderState(DeviceIndex As Long, Index As Long, State As Long, Value As Long)
    AddStatusText "Setting Render State... "
    Event3DEngine.Devices(DeviceIndex).SetRenderState Index, State, Value
    AddStatusText "Done!" & vbCrLf
End Sub

Private Sub RemoveRenderState(DeviceIndex As Long, Index As Long)
    AddStatusText "Removing Render State... "
    Event3DEngine.Devices(DeviceIndex).RemoveRenderState Index
    AddStatusText "Done!" & vbCrLf
End Sub

Private Sub InitMeshFromFile(DeviceIndex As Long, File As String)
    AddStatusText "Initializing Mesh From File (" & File & ")... "
    Event3DEngine.Devices(DeviceIndex).Meshes.AddMeshFromFile File
    AddStatusText "Done!" & vbCrLf
End Sub

Private Sub RemoveDevice(Index As Long)
    AddStatusText "Removing Device " & Index & "... "
    If Index > Event3DEngine.Devices.Count Then Err.Raise 2
    Event3DEngine.Devices.Remove Index
    AddStatusText "Done!" & vbCrLf
End Sub

Private Sub AddStatusText(Text As String)
    On Error GoTo ErrorH:
    If EchoOff = True Then Exit Sub
    StatusText = StatusText & Text
    RaiseEvent EngineEvent(Text)

ErrorH:
End Sub

Private Sub InitAudio()
    AddStatusText "Initializing Music Engine..." & vbCrLf
    AddStatusText "Adding Instance... "
    Set EventMusicEngine = EventEngine.EventMusicEngineCreate(EventMusicEngineCreationPassword)
    AddStatusText "Done!" & vbCrLf
    With EventMusicEngine
        AddStatusText "Initializing Audio... "
        .InitAudio CurrentState.hFocusWindow
        AddStatusText "Done!" & vbCrLf & "Music Engine Initialization Complete!" & vbCrLf
    End With
End Sub

Private Sub SetAudioSearchDirectory(SearchDirectory As String)
    AddStatusText "Setting Search Directory... "
        EventMusicEngine.SetSearchDirectory SearchDirectory
    AddStatusText "Done!" & vbCrLf
End Sub

Private Sub InitAudioSegment(File As String)
    AddStatusText "Initializing Audio File (" & File & ")... "
    With EventMusicEngine
        .CreateAudioSegment File
    End With
    AddStatusText "Done!" & vbCrLf
End Sub

Private Sub PlayEngineSegment(Index As Long, Repeats As Long)
    AddStatusText "Playing Segment " & Index & "... "
    If Index > EventMusicEngine.NumberOfSegments - 1 Then Err.Raise 2
    EventMusicEngine.PlayAudioSegment Index, Repeats
    AddStatusText "Done!" & vbCrLf
End Sub

Private Sub StopEngineSegment(Index As Long)
    AddStatusText "Stoping Segment " & Index & "... "
    If Index > EventMusicEngine.NumberOfSegments - 1 Then Err.Raise 2
    EventMusicEngine.StopAudioSegment Index
    AddStatusText "Done!" & vbCrLf
End Sub

Private Sub RemoveEngineSegment(Index As Long)
    AddStatusText "Removing Segment " & Index & "... "
    If Index > EventMusicEngine.NumberOfSegments - 1 Then Err.Raise 2
    EventMusicEngine.RemoveAudioSegment Index
    AddStatusText "Done!" & vbCrLf
End Sub

Public Sub SendCommand(Command As String)
    On Error GoTo ErrorH:
    Dim CommandSubj As String
    Static WithSubj As String, WithCaption As String
    Dim EngineSubj As String
    Dim DataCls As New DataClass
    Dim Windowed As Boolean
    Dim lValue1 As Long
    Dim lValue2 As Long
    Dim lValue3 As Long
    Dim lValue4 As Long
    Dim sValue1 As String
    AddStatusText WithCaption & "> " & Command & vbCrLf
    If InStr(1, Command, ".", vbTextCompare) > 0 Then
        EngineSubj = UCase(DataCls.DataRead(Command, " ", 0, "."))
        CommandSubj = UCase(DataCls.DataRead(Command, ".", 1, ":"))
    End If
    Select Case UCase(DataCls.DataRead(Command, " ", 0, " "))
            
        Case "RUNCONSOLESCRIPT:"
            sValue1 = DataCls.DataRead(Command, " ", 1, ",")
            RunConsoleScript sValue1
            
        Case "WITH:"
            If WithSubj = "" Then
                WithCaption = DataCls.DataRead(Command, " ", 1, ",")
                WithSubj = UCase(WithCaption)
            Else
                Err.Raise 2
            End If
            
        Case "ENDWITH:"
            WithCaption = ""
            WithSubj = ""
        
        Case "ENGINEINFO:"
            DetailedEngineInfo
            
        Case Else
            GoTo Continue1:
            
    End Select
    Exit Sub
Continue1:
    
    If EngineSubj = "3DENGINE" Or WithSubj = "3DENGINE" Then
        Select Case CommandSubj
            
            Case "INIT3DENGINE:"
                Init3DEngine
            
            Case "INITDEVICE:"
                lValue1 = Val(DataCls.DataRead(Command, " ", 1, ","))
                lValue2 = Val(DataCls.DataRead(Command, ",", 1))
                lValue3 = Val(DataCls.DataRead(Command, ",", 2))
                Windowed = UCase(DataCls.DataRead(Command, ",", 3))
                If Windowed = "TRUE" Then
                    InitDevice lValue1, lValue2, lValue3, True
                ElseIf Windowed = "FALSE" Then
                    InitDevice lValue1, lValue2, lValue3, False
                End If
                
            Case "RESIZEFULLSCREEN:"
                lValue1 = Val(DataCls.DataRead(Command, " ", 1, ","))
                lValue2 = Val(DataCls.DataRead(Command, ",", 1))
                ResizeFullScreen lValue1, lValue2
                
            Case "RESIZEWINDOWED:"
                lValue1 = Val(DataCls.DataRead(Command, " ", 1))
                ResizeWindowed lValue1
                
            Case "ADDRENDERSTATE:"
                lValue1 = Val(DataCls.DataRead(Command, " ", 1, ","))
                lValue2 = Val(DataCls.DataRead(Command, ",", 1, ","))
                lValue3 = Val(DataCls.DataRead(Command, ",", 2))
                AddRenderState lValue1, lValue2, lValue3
                                
            Case "SETRENDERSTATE:"
                lValue1 = Val(DataCls.DataRead(Command, " ", 1, ","))
                lValue2 = Val(DataCls.DataRead(Command, ",", 1, ","))
                lValue3 = Val(DataCls.DataRead(Command, ",", 2, ","))
                lValue4 = Val(DataCls.DataRead(Command, ",", 3))
                SetRenderState lValue1, lValue2, lValue3, lValue4
                                
            Case "REMOVERENDERSTATE:"
                lValue1 = Val(DataCls.DataRead(Command, " ", 1, ","))
                lValue2 = Val(DataCls.DataRead(Command, ",", 1))
                RemoveRenderState lValue1, lValue2
                
            Case "SETUPRENDERSTATES:"
                lValue1 = Val(DataCls.DataRead(Command, " ", 1))
                AddStatusText "Setting Up Render States... "
                Event3DEngine.Devices(lValue1).SetupRenderStates
                AddStatusText "Done!" & vbCrLf
                
            Case "INITMESHFROMFILE:"
                lValue1 = Val(DataCls.DataRead(Command, " ", 1, ","))
                sValue1 = DataCls.DataRead(Command, ",", 1)
                InitMeshFromFile lValue1, sValue1
                
            Case "REMOVEDEVICE:"
                lValue1 = DataCls.DataRead(Command, " ", 1, ",")
                RemoveDevice lValue1
                
            Case Else
                GoTo Continue2
                
        End Select
        Exit Sub
Continue2:
    End If
    
    If EngineSubj = "MUSICENGINE" Or WithSubj = "MUSICENGINE" Then
        Select Case CommandSubj
            
            Case "INITAUDIO:"
                InitAudio
                
            Case "SETSEARCHDIRECTORY:"
                sValue1 = DataCls.DataRead(Command, " ", 1)
                SetAudioSearchDirectory sValue1
                
            Case "INITAUDIOFILE:"
                sValue1 = DataCls.DataRead(Command, " ", 1)
                InitAudioSegment sValue1
                
            Case "PLAYENGINESEGMENT:"
                lValue1 = DataCls.DataRead(Command, " ", 1, ",")
                lValue2 = DataCls.DataRead(Command, ",", 1)
                PlayEngineSegment lValue1, lValue2
                
            Case "STOPENGINESEGMENT:"
                lValue1 = DataCls.DataRead(Command, " ", 1)
                StopEngineSegment lValue1
                        
            Case "REMOVEENGINESEGMENT:"
                lValue1 = DataCls.DataRead(Command, " ", 1)
                RemoveEngineSegment lValue1
            
            Case Else
                GoTo Continue3
                
        End Select
        Exit Sub
Continue3:
    End If
    If InStr(1, Command, ":", vbTextCompare) > 0 Then
        AddStatusText "Invalid console command!" & vbCrLf
    End If

ErrorH:
    If Err Then
        If InStr(1, Command, ":", vbTextCompare) > 0 Then
            AddStatusText "Invalid console command!" & vbCrLf
        End If
    End If
End Sub

Public Sub RunConsoleScript(File As String)
    AddStatusText "Running Console Script (" & File & ")..." & vbCrLf
    Dim FLine As String
    Dim FreeFileNumber As Long
    FreeFileNumber = FreeFile
    Open File For Input As FreeFileNumber
        Do Until EOF(FreeFileNumber)
            Line Input #FreeFileNumber, FLine
            If UCase(FLine) = "ECHO OFF" Then
                EchoOff = True
            ElseIf UCase(FLine) = "ECHO ON" Then
                EchoOff = False
            Else
                SendCommand FLine
            End If
        Loop
    Close FreeFileNumber
    AddStatusText "Done Console Script!" & vbCrLf
End Sub

Private Sub DetailedEngineInfo()
    EngineInfo
    AddStatusText "Version Info" & vbCrLf
    AddStatusText "Event 3D Engine: 0.5" & vbCrLf
    AddStatusText "Event Physics Engine: 0.01" & vbCrLf
    AddStatusText "Event Engine Console: 0.5" & vbCrLf & vbCrLf
End Sub

Private Sub EngineInfo()
    AddStatusText "Event Engine" & vbCrLf
    AddStatusText "Build " & App.Major & "." & App.Minor & App.Revision & vbCrLf
    AddStatusText "Event Entertainment© 2001" & vbCrLf & vbCrLf
End Sub

Private Sub Class_Initialize()
    EngineInfo
End Sub

Private Sub Class_Terminate()
    Set Event3DEngine = Nothing
    Set EventEngine = Nothing
End Sub
