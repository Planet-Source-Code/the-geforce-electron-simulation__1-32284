Attribute VB_Name = "EventMain"
Public Sub SetupObjects()
    Dim i As Integer
    With MatCalc
        Repulsion.Calculate
        For i = 1 To Repulsion.NumberOfParticles
            Repulsion.OutParticles i
            .MatrixTranslation Repulsion.OutX, Repulsion.OutY, Repulsion.OutZ, After
            .Transform After
            Event3DEngine.Devices(2).Frames(2).ChildFrames(i).SetFrameMatrix 1
            .ResetTransform
        Next i
    End With
End Sub

Public Sub SetupLights()
    With RLight
        .position.X = 1.5 * Cos(DegL1 * (PI / 180))
        .position.Y = 1.5 * Sin(DegL1 * (PI / 180))
    End With
    Event3DEngine.Devices(2).SetLight 0, RLight
    With RLight
        .position.X = 1.5 * Cos(DegL2 * (PI / 180))
        .position.Y = 1.5 * Sin(DegL2 * (PI / 180))
    End With
    Event3DEngine.Devices(2).SetLight 1, RLight
    DegL1 = DegL1 + 1
    DegL2 = DegL2 - 1
    If DegL1 > 180 Then DegL1 = -180
    If DegL2 < -180 Then DegL2 = 180
End Sub

Public Sub RenderLoop()
    Dim i As Integer
    Dim Start As Single
    Do
        Start = Timer
        Do
            SetupObjects
            SetupLights
            Event3DEngine.Devices(2).Render &H0
        Loop Until Timer - Start > 0.1
        DoEvents
    Loop Until Rendering = False
    EventEngineConsole.SendCommand "3DEngine.RemoveDevice: 2"
    ViewWin.Visible = False
    If Command <> "" Then Cleanup
End Sub
