VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Device"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"Frames"
Attribute VB_Ext_KEY = "Member1" ,"Meshes"
Attribute VB_Ext_KEY = "Member2" ,"TextHandlers"
Attribute VB_Ext_KEY = "Member3" ,"TextStrings"
'Event Engine by Daniel Bloemendal
'If you take any code from this engine you must give me credit!

'This engine is most likely going to be discontinued.
'This is because I am working on a C++ engine.
'If you would like to continue the production of this engine,
'You must get permission from me at teh_leet_programming_man@hotmail.com
'Thank you for downloading my code please give me feedback.
'Vote for me!!!

'***** Public Start Type Declarations *****'
    '***** Start Adapters Info Type Declarations *****'
        Public Type ModeInfo
            Height As Long
            Width As Long
            BPP As Long
        End Type
        
        Public Type FormatInfo
            Usage As Long
            CanDoWindowed As Boolean
            CanDoFullscreen As Boolean
        End Type
        
        Public Type DevTypeInfo
            NumModes As Long
            Modes() As ModeInfo
            NumFormats As Long
            FormatInfo() As FormatInfo
            CurrentMode As Long
        End Type
        
        Public Type AdapterInfo
            Desc As String
            CanDoWindowed As Long
            DevTypeInfo(1 To 2) As DevTypeInfo
            Windowed As Boolean
            Reference As Boolean
            Description As String
        End Type
    '***** End Adapters Info Type Declarations *****'
    '***** Start Shading Type Declarations *****'
        Public Type ColorValue
            A As Single
            b As Single
            g As Single
            r As Single
        End Type
        
        Public Enum LightType
            PointLight = 1
            SpotLight = 2
            DirectionalLight = 3
        End Enum
        
        Public Type Light
            ambient As ColorValue
            diffuse As ColorValue
            specular As ColorValue
            type As LightType
            Attenuation0 As Single
            Attenuation1 As Single
            Attenuation2 As Single
            Direction As Vector
            Falloff As Single
            Phi As Single
            position As Vector
            Range As Single
            Theta As Single
        End Type
        
        Private Type State_D3DLight8
            Light As D3DLIGHT8
            Changed As Boolean
        End Type
        
        Public Type Material
            ambient As ColorValue
            diffuse As ColorValue
            emissive As ColorValue
            power As Single
            specular As ColorValue
        End Type
    '***** End Shading Type Declarations *****'
'***** End Type Declarations *****'
'***** Start Private Declarations *****'
    Private AdaptersInfo() As D3D_ADAPTERINFO
    Private D3DPP As D3DPRESENT_PARAMETERS
    Private CurrentAdapter As Long
    Private FormatNumber As Long
    Private VertexFlags As Long
    Private HardwareDetected As Boolean
    Private Initialized As Boolean
    Private D3DDevice As Direct3DDevice8
    Private Const D3DFVF_CUSTOMVERTEX = (D3DFVF_XYZ Or D3DFVF_NORMAL Or D3DFVF_TEX1)
    Private mvarFrames As Frames
    Private mvarMeshes As Meshes
    Private RenderStates() As RenderState
    Private NumberOfRenderStates As Long
    Private Lights() As State_D3DLight8
    Private NumberOfLights As Long
    Private ViewMatrix As D3DMATRIX
'***** End Private Declarations *****'
'***** Start Public Declarations *****'
    Public NumAdapters As Long
    Public hFocusWindow As Long
    Public PauseRender As Boolean
'***** End Public Declarations *****'

Public Sub DetectHardware()
    D3DEnum_BuildAdapterList D3DDevice
    HardwareDetected = True
End Sub

Public Function GetAdaptersInfo(AdapterIndex As Integer) As AdapterInfo
    On Error Resume Next
    Dim DevNum As Integer
    Dim ModeNum As Integer
    Dim FrmtNum As Integer
    Dim Description As String
    Dim ChrNum As Long
    With AdaptersInfo(AdapterIndex)
        GetAdaptersInfo.CanDoWindowed = .CanDoWindowed
        GetAdaptersInfo.Desc = .Desc
        For DevNum = 1 To 2
            With AdaptersInfo(AdapterIndex).DevTypeInfo(DevNum)
                GetAdaptersInfo.DevTypeInfo(DevNum).NumModes = .NumModes
                GetAdaptersInfo.DevTypeInfo(DevNum).NumFormats = .NumFormats
                GetAdaptersInfo.DevTypeInfo(DevNum).CurrentMode = .CurrentMode
                ReDim GetAdaptersInfo.DevTypeInfo(DevNum).Modes(.NumModes - 1)
                ReDim GetAdaptersInfo.DevTypeInfo(DevNum).FormatInfo(.NumFormats - 1)
                For ModeNum = 0 To .NumModes - 1
                    With AdaptersInfo(AdapterIndex).DevTypeInfo(DevNum).Modes(ModeNum)
                        GetAdaptersInfo.DevTypeInfo(DevNum).Modes(ModeNum).Width = .Width
                        GetAdaptersInfo.DevTypeInfo(DevNum).Modes(ModeNum).Height = .Height
                        If .Format = D3DFMT_X8R8G8B8 Or .Format = D3DFMT_A8R8G8B8 Or .Format = D3DFMT_R8G8B8 Then
                            GetAdaptersInfo.DevTypeInfo(DevNum).Modes(ModeNum).BPP = 32
                        Else
                            GetAdaptersInfo.DevTypeInfo(DevNum).Modes(ModeNum).BPP = 16
                        End If
                    End With
                Next ModeNum
                For FrmtNum = 0 To .NumFormats - 1
                    With AdaptersInfo(AdapterIndex).DevTypeInfo(dev).FormatInfo(FrmtNum)
                        GetAdaptersInfo.DevTypeInfo(DevNum).FormatInfo(frmt).CanDoFullscreen = .CanDoFullscreen
                        GetAdaptersInfo.DevTypeInfo(DevNum).FormatInfo(frmt).CanDoWindowed = .CanDoWindowed
                        GetAdaptersInfo.DevTypeInfo(DevNum).FormatInfo(frmt).Usage = .Usage
                    End With
                Next FrmtNum
            End With
        Next DevNum
        Description = vbNullString
        For ChrNum = 0 To 511
            Description = Description & Chr$(.D3DAI.Description(ChrNum))
        Next ChrNum
        Description = Replace$(Description, Chr$(0), " ")
        GetAdaptersInfo.Description = Description
    End With
End Function

Public Sub InitDevice(Adapter As Long, Optional DeviceType As CONST_D3DDEVTYPE = D3DDEVTYPE_HAL, Optional Windowed As Boolean = False, Optional Mode As Long)
    With AdaptersInfo(Adapter)
        .Windowed = Windowed
        .DeviceType = DeviceType
        If .DeviceType = D3DDEVTYPE_HAL And .DevTypeInfo(1).NumModes <= 0 Then
            .DeviceType = D3DDEVTYPE_REF
            .Reference = True
        ElseIf DeviceType = D3DDEVTYPE_SW Then
            .DeviceType = D3DDEVTYPE_REF
            .Reference = True
        End If
        If D3DEnum_FindInFormatList(Adapter, .DeviceType, .DesktopMode.Format) <> -1 Then
            If .DevTypeInfo(.DeviceType).FormatInfo(FormatNumber).CanDoWindowed = False Then
                .Windowed = False
                .CanDoWindowed = 0
            Else
                Call D3DEnum_FindInFormatList(Adapter, .DeviceType, .DevTypeInfo(.DeviceType).Modes(Mode).Format)
                If .DevTypeInfo(.DeviceType).FormatInfo(FormatNumber).CanDoFullscreen = False Then
                    .Windowed = True
                    .CanDoWindowed = 1
                End If
            End If
        Else
            .Windowed = False
            .CanDoWindowed = 0
        End If
        If .Windowed = True Then GoSub WindowInit
        If .Windowed = False Then GoSub FullScreenInit
        Exit Sub

WindowInit:
        D3DPP.Windowed = .Windowed
        D3DPP.BackBufferFormat = .DesktopMode.Format
        D3DPP.BackBufferWidth = 0
        D3DPP.BackBufferHeight = 0
        D3DPP.hDeviceWindow = hFocusWindow
        D3DPP.EnableAutoDepthStencil = 1
        D3DPP.AutoDepthStencilFormat = D3DFMT_D16
        D3DPP.SwapEffect = D3DSWAPEFFECT_DISCARD
        VertexFlags = D3DEnum_FindInFormatList(Adapter, .DeviceType, .DesktopMode.Format)
        If VertexFlags < 0 Then VertexFlags = D3DCREATE_SOFTWARE_VERTEXPROCESSING
        Set D3DDevice = D3D.CreateDevice(Adapter, .DeviceType, hFocusWindow, VertexFlags, D3DPP)
        CurrentAdapter = Adapter
        Set Meshes.D3DDevice = D3DDevice
        Set TextHandlers.D3DDevice = D3DDevice
        SetupDeviceObjects
        Return

FullScreenInit:
        D3DPP.Windowed = .Windowed
        D3DPP.BackBufferFormat = .DevTypeInfo(.DeviceType).Modes(Mode).Format
        D3DPP.BackBufferWidth = .DevTypeInfo(.DeviceType).Modes(Mode).Width
        D3DPP.BackBufferHeight = .DevTypeInfo(.DeviceType).Modes(Mode).Height
        D3DPP.hDeviceWindow = hFocusWindow
        D3DPP.EnableAutoDepthStencil = 1
        D3DPP.AutoDepthStencilFormat = D3DFMT_D16
        D3DPP.SwapEffect = D3DSWAPEFFECT_DISCARD
        VertexFlags = D3DEnum_FindInFormatList(Adapter, .DeviceType, .DevTypeInfo(.DeviceType).Modes(Mode).Format)
        If VertexFlags < 0 Then VertexFlags = D3DCREATE_SOFTWARE_VERTEXPROCESSING
        Set D3DDevice = D3D.CreateDevice(Adapter, .DeviceType, hFocusWindow, VertexFlags, D3DPP)
        CurrentAdapter = Adapter
        Set Meshes.D3DDevice = D3DDevice
        SetupDeviceObjects
        Return

    End With
End Sub

Public Sub ResizeWindowed()
    With AdaptersInfo(CurrentAdapter)
        .Windowed = True
        D3DPP.Windowed = 1
        D3DPP.BackBufferWidth = 0
        D3DPP.BackBufferHeight = 0
        D3DPP.BackBufferFormat = .DesktopMode.Format
        D3DPP.SwapEffect = D3DSWAPEFFECT_DISCARD
        Reset D3DPP
    End With
End Sub

Public Sub ResizeFullScreen(Mode As Long)
    With AdaptersInfo(CurrentAdapter)
        .Windowed = False
        D3DPP.Windowed = 0
        D3DPP.BackBufferWidth = .DevTypeInfo(.DeviceType).Modes(Mode).Width
        D3DPP.BackBufferHeight = .DevTypeInfo(.DeviceType).Modes(Mode).Height
        D3DPP.BackBufferFormat = .DevTypeInfo(.DeviceType).Modes(Mode).Format
        D3DPP.SwapEffect = D3DSWAPEFFECT_DISCARD
        Reset D3DPP
    End With
End Sub

Private Sub Reset(D3DPresentParameters As D3DPRESENT_PARAMETERS)
    On Error GoTo ErrorH
    PauseRender = True
    D3DDevice.Reset D3DPresentParameters
    SetupDeviceObjects
    PauseRender = False

ErrorH:
    If Err Then
        
        If D3DDevice Is Nothing Then
            Exit Sub
        End If
        
        If Err = D3DERR_INVALIDCALL Then
            Err.Clear
            Exit Sub
        End If
            
        If D3DDevice.TestCooperativeLevel = D3DERR_DEVICELOST Then
            Err.Clear
            Exit Sub
        Else
            Err.Raise 2
        End If

    End If
End Sub

Public Sub Render(Optional ClearColor As Long = &HFF)
    On Error GoTo ErrorH
    Dim i As Long
    If PauseRender = True Then Exit Sub
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, ClearColor, 1#, 0
        D3DDevice.BeginScene
            SetupWorldMatrix
            SetupLights
            With Frames
                For i = 1 To .Count
                    .Item(i).TransformAndDrawFrame D3DDevice, DrawOpaque
                Next i
                For i = 1 To .Count
                    .Item(i).TransformAndDrawFrame D3DDevice, DrawAlpha
                Next i
            End With
            D3DDevice.SetVertexShader D3DFVF_CUSTOMVERTEX
        D3DDevice.EndScene
    D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
    
ErrorH:
    If Err Then
        If D3DDevice Is Nothing Then
            Exit Sub
        End If
        If Err = D3DERR_DEVICELOST Then
            If D3DDevice.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then
                Reset D3DPP
            ElseIf D3DDevice.TestCooperativeLevel = D3DERR_DEVICELOST Then
                Err.Clear
                Exit Sub
            End If
        ElseIf Err = D3DERR_INVALIDCALL Then
            Err.Raise 2
        End If
    End If
End Sub

Public Sub SetViewMatrix(MatrixIndex As Integer)
    ViewMatrix = Matrices(MatrixIndex)
End Sub

Public Sub AddRenderState(State As Long, Value As Long)
    ReDim Preserve RenderStates(NumberOfRenderStates)
    RenderStates(NumberOfRenderStates).State = State
    RenderStates(NumberOfRenderStates).Value = Value
    NumberOfRenderStates = NumberOfRenderStates + 1
End Sub

Public Sub SetRenderState(Index As Long, State As Long, Value As Long)
    RenderStates(Index).State = State
    RenderStates(Index).Value = Value
End Sub

Public Sub RemoveRenderState(Index As Long)
    Dim i As Long
    For i = Index To NumberOfRenderStates - 2
        RenderStates(i) = RenderStates(i + 1)
    Next i
    NumberOfRenderStates = NumberOfRenderStates - 1
    ReDim Preserve RenderStates(NumberOfRenderStates - 1)
End Sub

Public Function GetNumberOfRenderStates() As Long
    GetNumberOfRenderStates = NumberOfRenderStates
End Function

Public Sub AddLight(Light As Light)
    ReDim Preserve Lights(NumberOfLights)
    With Lights(NumberOfLights)
        .Changed = True
        .Light = LightToD3DLight8(Light)
    End With
    NumberOfLights = NumberOfLights + 1
End Sub

Public Sub SetLight(Index As Long, Light As Light)
    With Lights(Index)
        .Changed = True
        .Light = LightToD3DLight8(Light)
    End With
End Sub

Public Sub RemoveLight(Index As Long)
    Dim i As Long
    For i = Index To NumberOfLights - 2
        Lights(i) = Lights(i + 1)
    Next i
    NumberOfLights = NumberOfLights - 1
    D3DDevice.LightEnable NumberOfLights, 0
    ReDim Preserve Lights(NumberOfLights - 1)
End Sub

Public Sub SetupDeviceObjects()
    SetupRenderStates
    SetupLights True
End Sub

Public Sub SetupRenderStates()
    Dim i As Long
    For i = 0 To NumberOfRenderStates - 1
        D3DDevice.SetRenderState RenderStates(i).State, RenderStates(i).Value
    Next i
End Sub

Public Sub SetupLights(Optional All As Boolean = False)
    Dim i As Integer
    For i = 0 To NumberOfLights - 1
        With Lights(i)
            If .Changed = True Or All = True Then
                D3DDevice.SetLight i, .Light
                D3DDevice.LightEnable i, 1
            End If
        End With
    Next i
End Sub

Private Sub SetupWorldMatrix()
    Const pi = 3.14159265359
    Dim matProj As D3DMATRIX
    Dim Aspect As Single
    Dim ClientRect As RECT
    Call GetWindowRect(hFocusWindow, ClientRect)
    With ClientRect
        If .bottom - .Top <> 0 And .Right - .Left Then Aspect = (.bottom - .Top) / (.Right - .Left)
    End With
    D3DDevice.SetTransform D3DTS_VIEW, ViewMatrix
    D3DXMatrixPerspectiveFovLH matProj, pi / 4, Aspect, 1, 1000
    D3DDevice.SetTransform D3DTS_PROJECTION, matProj
End Sub

'****************** Start Shaders ******************'
    '!!!!! Warning, not all devices support these shaders !!!!!'
    Public Function CreateVertexShader(VertexShaderCode As String) As Long
        Dim Code As D3DXBuffer
        Dim VertexShader() As Long
        Dim Size As Long
        Dim Decl(4) As Long
        Decl(0) = D3DVSD_STREAM(0)
        Decl(1) = D3DVSD_REG(D3DVSDE_POSITION, D3DVSDT_FLOAT3)
        Decl(2) = D3DVSD_REG(D3DVSDE_NORMAL, D3DVSDT_FLOAT3)
        Decl(3) = D3DVSD_REG(D3DVSDE_TEXCOORD1, D3DVSDT_FLOAT2)
        Decl(4) = D3DVSD_END()
        Set Code = D3DX.AssembleShader(VertexShaderCode, 0, Nothing)
        ReDim VertexShader(Code.GetBufferSize() / 4)
        D3DX.BufferGetData Code, 0, 1, Code.GetBufferSize(), VertexShader(0)
        D3DDevice.CreateVertexShader Decl(0), VertexShader(0), CreateVertexShader, 0
    End Function
    
    Public Sub SetVertexShaderConstants(Register As Long, Constants() As Single, NumberOfConstants As Long)
        D3DDevice.SetVertexShaderConstant Register, Constants(0), NumberOfConstants
    End Sub
    
    Public Sub DestroyVertexShader(VertexShaderHandle As Long)
        D3DDevice.DeleteVertexShader VertexShaderHandle
        D3DDevice.SetVertexShader 0
    End Sub
    
    Public Function CreatePixelShader(PixelShaderCode As String) As Long
        Dim Code As D3DXBuffer
        Dim PixelShader() As Long
        Dim Size As Long
        
        Set Code = D3DX.AssembleShader(PixelShaderCode, 0, Nothing)
        Size = Code.GetBufferSize() / 4
        ReDim PixelShader(Size - 1)
        D3DX.BufferGetData Code, 0, 1, Code.GetBufferSize, PixelShader(0)
        CreatePixelShader = D3DDevice.CreatePixelShader(PixelShader(0))
    End Function
    
    Public Sub SetPixelShaderConstants(Register As Long, Constants() As Single, NumberOfConstants As Long)
        D3DDevice.SetPixelShaderConstant Register, Constants(0), NumberOfConstants
    End Sub
    
    Public Sub DestroyPixelShader(PixelShaderHandle As Long)
        D3DDevice.DeletePixelShader PixelShaderHandle
        D3DDevice.SetPixelShader 0
    End Sub
    '!!!!! Warning, not all devices support these shaders !!!!!'
'****************** End Shaders ******************'

Private Function D3DEnum_BuildAdapterList(EnumerationCallback As Object) As Boolean
    
    On Local Error GoTo errOut
    
    Dim Adapter As Long
        
    Call D3DEnum_Cleanup
            
    If DX8 Is Nothing Then Set DX8 = New DirectX8
    If D3D Is Nothing Then Set D3D = DX8.Direct3DCreate
    If D3DX Is Nothing Then Set D3DX = New D3DX8
    
    Set EnumCallback = EnumerationCallback
    
    NumAdapters = D3D.GetAdapterCount
    ReDim AdaptersInfo(NumAdapters)
    
    For Adapter = 0 To NumAdapters - 1
    
        D3DEnum_BuildValidFormatList Adapter, D3DDEVTYPE_HAL
        D3DEnum_BuildValidFormatList Adapter, D3DDEVTYPE_REF
        
                
        D3DEnum_BuildDisplayModeList Adapter, D3DDEVTYPE_HAL
        D3DEnum_BuildDisplayModeList Adapter, D3DDEVTYPE_REF
        
        D3D.GetAdapterIdentifier Adapter, 0, AdaptersInfo(Adapter).D3DAI
        
    Next
    
    D3DEnum_BuildAdapterList = True
    Exit Function
    
errOut:
    Debug.Print "Failed D3DEnum_BuildAdapterList"
End Function

Private Sub D3DEnum_BuildDisplayModeList(Adapter As Long, DevType As CONST_D3DDEVTYPE)

        Dim Mode As Long
        Dim NumModes As Long
        Dim DisplayMode As D3DDISPLAYMODE

        With AdaptersInfo(Adapter).DevTypeInfo(DevType)

            .NumModes = 0

            NumModes = D3D.GetAdapterModeCount(Adapter)

            ReDim .Modes(NumModes)

            Call D3D.GetAdapterDisplayMode(Adapter, AdaptersInfo(Adapter).DesktopMode)

            For Mode = 0 To NumModes - 1

                Call D3D.EnumAdapterModes(Adapter, Mode, DisplayMode)

                'If -1 <> D3DEnum_FindInDisplayModeList(Adapter, DevType, DisplayMode) Then GoTo Continue


                If -1 = D3DEnum_FindInFormatList(Adapter, DevType, DisplayMode.Format) Then GoTo Continue

                .Modes(.NumModes).Height = DisplayMode.Height
                .Modes(.NumModes).Width = DisplayMode.Width
                .Modes(.NumModes).Format = DisplayMode.Format
                .NumModes = .NumModes + 1

Continue:
            Next

        End With

End Sub

Private Sub D3DEnum_BuildValidFormatList(Adapter As Long, DevType As CONST_D3DDEVTYPE)

        Dim Mode As Long
        Dim Usage As Long
        Dim NumModes As Long
        Dim DisplayMode As D3DDISPLAYMODE
        Dim CanDoWindowed As Boolean
        Dim CanDoFullscreen As Boolean


        With AdaptersInfo(Adapter).DevTypeInfo(DevType)

            .NumFormats = 0

            NumModes = D3D.GetAdapterModeCount(Adapter)
            ReDim .FormatInfo(NumModes)

            For Mode = 0 To NumModes - 1

                Call D3D.EnumAdapterModes(Adapter, Mode, DisplayMode)

                If -1 <> D3DEnum_FindInFormatList(Adapter, DevType, DisplayMode.Format) Then GoTo Continue

                Usage = D3DEnum_CheckFormatCompatibility(Adapter, DevType, DisplayMode.Format, CanDoWindowed, CanDoFullscreen)

                If -1 = Usage Then GoTo Continue

                .FormatInfo(.NumFormats).Format = DisplayMode.Format
                .FormatInfo(.NumFormats).Usage = Usage
                .FormatInfo(.NumFormats).CanDoWindowed = CanDoWindowed
                .FormatInfo(.NumFormats).CanDoFullscreen = CanDoFullscreen
                .NumFormats = .NumFormats + 1


Continue:
            Next

        End With

End Sub

Private Function D3DEnum_CheckFormatCompatibility(Adapter As Long, DeviceType As CONST_D3DDEVTYPE, Format As CONST_D3DFORMAT, ByRef OutCanDoWindowed As Boolean, ByRef OutCanDoFullscreen As Boolean) As Long
        On Local Error GoTo errOut

        D3DEnum_CheckFormatCompatibility = -1

        Dim D3DCaps As D3DCAPS8
        Dim Flags As Long

        OutCanDoWindowed = True: OutCanDoFullscreen = True
        If 0 <> D3D.CheckDeviceType(Adapter, DeviceType, Format, Format, 0) Then OutCanDoWindowed = False
        If 0 <> D3D.CheckDeviceType(Adapter, DeviceType, Format, Format, 1) Then OutCanDoFullscreen = False
        If (OutCanDoWindowed = False) And (OutCanDoFullscreen = False) Then Exit Function


        D3D.GetDeviceCaps Adapter, DeviceType, D3DCaps
        AdaptersInfo(Adapter).D3DCaps = D3DCaps

        D3DEnum_CheckFormatCompatibility = D3DCREATE_SOFTWARE_VERTEXPROCESSING
        If EnumCallback Is Nothing Then Exit Function

        Flags = D3DCREATE_HARDWARE_VERTEXPROCESSING
        D3DEnum_CheckFormatCompatibility = Flags
        If D3DCaps.DevCaps And D3DDEVCAPS_HWTRANSFORMANDLIGHT Then
           If EnumCallback.VerifyDevice(Flags, Format) Then Exit Function
        End If

        Flags = D3DCREATE_SOFTWARE_VERTEXPROCESSING
        D3DEnum_CheckFormatCompatibility = Flags
        If EnumCallback.VerifyDevice(Flags, Format) Then Exit Function

        D3DEnum_CheckFormatCompatibility = -1

        Exit Function
errOut:

End Function

Private Sub D3DEnum_Cleanup()
    Set EnumCallback = Nothing
    ReDim adapters(0)
End Sub

Private Function D3DEnum_FindInDisplayModeList(Adapter As Long, DevType As CONST_D3DDEVTYPE, DisplayMode As D3DDISPLAYMODE) As Long

    Dim Mode As Long
    Dim NumModes As Long

    NumModes = AdaptersInfo(Adapter).DevTypeInfo(DevType).NumModes
    D3DEnum_FindInDisplayModeList = -1

    For Mode = 0 To NumModes - 1
      With AdaptersInfo(Adapter).DevTypeInfo(DevType).Modes(Mode)
          If .Width = DisplayMode.Width And .Height = DisplayMode.Height And .Format = DisplayMode.Format Then
              D3DEnum_FindInDisplayModeList = Mode
              Exit Function
          End If
      End With
    Next

End Function

Private Function D3DEnum_FindInFormatList(Adapter As Long, DevType As CONST_D3DDEVTYPE, Format As CONST_D3DFORMAT) As Long
    
    Dim FormatNum As Long
    Dim NumFormats As Long

    NumFormats = AdaptersInfo(Adapter).DevTypeInfo(DevType).NumFormats
    D3DEnum_FindInFormatList = -1

    For FormatNum = 0 To NumFormats - 1
      With AdaptersInfo(Adapter).DevTypeInfo(DevType).FormatInfo(FormatNum)
          If .Format = Format Then
             D3DEnum_FindInFormatList = .Usage
             FormatNumber = FormatNum
             Exit Function
          End If
      End With
    Next

    D3DEnum_FindInFormatList = -1

End Function

Public Property Get Meshes() As Meshes
    If mvarMeshes Is Nothing Then
        Set mvarMeshes = New Meshes
    End If


    Set Meshes = mvarMeshes
End Property

Public Property Set Meshes(vData As Meshes)
    Set mvarMeshes = vData
End Property

Public Property Get Frames() As Frames
    If mvarFrames Is Nothing Then
        Set mvarFrames = New Frames
    End If


    Set Frames = mvarFrames
End Property

Public Property Set Frames(vData As Frames)
    Set mvarFrames = vData
End Property

Private Sub Class_Initialize()
    ResetThreadPriority GetCurrentThread, THREAD_PRIORITY_TIME_CRITICAL
End Sub

Private Sub Class_Terminate()
  Set mvarMeshes = Nothing
    Set mvarFrames = Nothing
End Sub

Public Sub ClearDeviceInfo()
    Set EnumCallback = Nothing
    ReDim adapters(0)
    HardwareDetected = False
End Sub

Private Sub Cleanup()
    Set DX8 = Nothing
    Set D3DX = Nothing
    Set D3D = Nothing
    Set D3DDevice = Nothing
End Sub
