Attribute VB_Name = "EventInitMod"
Public Sub Init()
    ChDir App.Path
    ResetThreadPriority GetCurrentThread, THREAD_PRIORITY_TIME_CRITICAL
    MainWin.Visible = True
    DegL1 = 0
    DegL2 = 180
    ConsoleFrm.InitEvents
    With EventEngineConsole
        .Event3DEngineCreationPassword = "! 4/\/\ 7|-|3 0/\/3"
        .SendCommand "3DEngine.Init3DEngine:"
        Set Event3DEngine = .Event3DEngine
        .EchoOff = True
    End With
    Event3DEngine.Devices.Add
    Event3DEngine.Devices(1).DetectHardware
    GetAdaptersInfo
    EnableMenus
End Sub

Public Sub GetAdaptersInfo()
    Dim i As Integer
    With Event3DEngine.Devices(1)
        .DetectHardware
        ReDim AdaptersInfo(.NumAdapters - 1)
        For i = 0 To .NumAdapters - 1
            AdaptersInfo(i) = .GetAdaptersInfo(i)
        Next i
    End With
End Sub

Public Sub PrepareGUI()
    ConsoleFrm.EndEvents
    EngineConfigurationFrm.Visible = False
    DisableMenus
    ViewWin.Visible = True
    ViewWin.UpdateSize
    DoEvents
End Sub

Public Sub Init3D()
    PrepareGUI
    InitRepulsion
    InitConsole
    InitView
    InitLights
    InitObjects
    EnableRendering
End Sub

Public Sub InitRepulsion()
    Dim Answer As String
    With Repulsion
        Open "Repulsion.ini" For Input As #1
            Input #1, Answer
            NumberOfParticles = Val(Answer)
            Input #1, Answer
            .Speed = Val(Answer)
            Input #1, Answer
            .Friction = Val(Answer)
        Close
        .SphereRadius = 1
    End With
End Sub

Public Sub InitConsole()
    On Error Resume Next
    Dim i As Integer
    With EventEngineConsole
        .TargethFocusWindow = ViewWin.hWnd
        .RunConsoleScript "Init.ini"
    End With
End Sub

Public Function LoadInitDeviceData() As InitDeviceData
    On Error GoTo ErrorH:
    Dim InTxt As String
    Dim DataCls As New DataClass
    LoadInitDeviceData.Windowed = False
    LoadInitDeviceData.hFocusWindow = ViewWin.hWnd
    Open "InitDevice.ini" For Input As #1
        Line Input #1, InTxt
        LoadInitDeviceData.AdapterIndex = Val(DataCls.DataRead(InTxt, " ", 1, ","))
        LoadInitDeviceData.DeviceType = Val(DataCls.DataRead(InTxt, ",", 1, ","))
        LoadInitDeviceData.Resolution = Val(DataCls.DataRead(InTxt, ",", 2, ","))
        If UCase$(DataCls.DataRead(InTxt, ",", 3)) = "TRUE" Then
            LoadInitDeviceData.Windowed = True
        End If
    Close

ErrorH:
Close
End Function

Public Sub InitView()
    Event3DEngine.NumberOfMatrices = 2
    Event3DEngine.SetNewMatrices False
    With MatCalc
        .MatrixIdentity 0
        .InMatrixIndex = 0
        .OutMatrixIndex = 1
        .MatrixLookAtLH Vector(0#, 0#, -4#), Vector(0#, 0#, 0#), Vector(0#, 1#, 0), After
        .Transform After
        Event3DEngine.Devices(2).SetViewMatrix 1
        .ResetTransform
    End With
End Sub

Public Sub InitLights()
    Dim col As ColorValue
    With col
        .A = 1
        .b = 1
        .g = 1
        .r = 1
    End With
    With RLight
       .Type = PointLight
        .diffuse = col
        .specular = col
        .Range = 8
        .position.z = -1.25
        .Attenuation1 = 1
    End With
    Event3DEngine.Devices(2).AddLight RLight
    Event3DEngine.Devices(2).AddLight RLight
End Sub

Public Sub InitObjects()
    Randomize Timer
    Dim i As Integer
    Dim r As Single
    Dim xp As Double, yp As Double, zp As Double
    With Event3DEngine.Devices(2)
        .Meshes.AddMeshFromFile "Background.x"
        .Meshes.AddMeshFromFile "Glass Sphere.x"
        .Meshes.AddMeshFromFile "Sphere.x"
        .Frames.Add .Meshes(1)
        .Frames.Add .Meshes(2)
        .Frames(1).SetFrameMatrix 0
        .Frames(2).SetFrameMatrix 0
        .Frames(2).AlphaBlend = True
    End With
    With Repulsion
        .NumberOfParticles = NumberOfParticles
        .SetNewParticleValues
        For i = 1 To NumberOfParticles
            Do
                xp = .SphereRadius * Rnd - .SphereRadius
                yp = .SphereRadius * Rnd - .SphereRadius
                zp = .SphereRadius * Rnd - .SphereRadius
            Loop Until xp ^ 2 + yp ^ 2 + zp ^ 2 <= .SphereRadius ^ 2
            .SetParticle i, xp, yp, zp, 1, 1
            With Event3DEngine.Devices(2)
                .Frames(2).ChildFrames.Add .Meshes(3)
            End With
        Next i
    End With
End Sub

Public Sub EnableMenus()
    MainWin.Game_Mnu.Visible = True
    MainWin.Options_Mnu.Visible = True
End Sub

Public Sub DisableMenus()
    MainWin.Game_Mnu.Visible = False
    MainWin.Options_Mnu.Visible = False
End Sub

Public Sub EnableRendering()
    Dim InitDat As InitDeviceData
    InitDat = LoadInitDeviceData
    If InitDat.Windowed = False Then ShowCursor False
    Rendering = True
    RenderLoop
End Sub

Public Sub DisableRendering()
    Dim InitDat As InitDeviceData
    InitDat = LoadInitDeviceData
    If InitDat.Windowed = False Then ShowCursor True
    Rendering = False
End Sub

Public Sub Cleanup()
    If Rendering = True Then
        DisableRendering
    End If
    Set Event3DEngine = Nothing
    Set EventEngineConsole = Nothing
    Set Repulsion = Nothing
    Unload ViewWin
    Unload EngineConfigurationFrm
    Unload ConsoleFrm
    Unload MainWin
    End
End Sub
