Attribute VB_Name = "Direct3DInfoMod"
'Event Engine by Daniel Bloemendal
'If you take any code from this engine you must give me credit!

'This engine is most likely going to be discontinued.
'This is because I am working on a C++ engine.
'If you would like to continue the production of this engine,
'You must get permission from me at teh_leet_programming_man@hotmail.com
'Thank you for downloading my code please give me feedback.
'Vote for me!!!

'***** Start Init Declarations *****'
    Public Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, ByRef RECT As RECT) As Long
    Public Declare Function GetClientRect Lib "user32.dll" (ByVal hWnd As Long, ByRef RECT As RECT) As Long
    Public EnumCallback As Object
    Public Type D3D_MODEINFO
        Format As CONST_D3DFORMAT
        Height As Long
        Width As Long
        VertexBehavior As CONST_D3DCREATEFLAGS
    End Type
    
    Public Type D3D_FORMATINFO
        Format As CONST_D3DFORMAT
        Usage As Long
        CanDoWindowed As Boolean
        CanDoFullscreen As Boolean
    End Type
    
    Public Type D3D_DEVTYPEINFO
        NumModes As Long
        Modes() As D3D_MODEINFO
        NumFormats As Long
        FormatInfo() As D3D_FORMATINFO
        CurrentMode As Long
    End Type
    
    Public Type D3D_ADAPTERINFO
        DeviceType As CONST_D3DDEVTYPE
        D3DCaps As D3DCAPS8
        Desc As String
        CanDoWindowed As Long
        DevTypeInfo(1 To 2) As D3D_DEVTYPEINFO
        D3DAI As D3DADAPTER_IDENTIFIER8
        DesktopMode As D3DDISPLAYMODE
        Windowed As Boolean
        Reference As Boolean
    End Type

    Public Type RenderState
        State As Long
        Value As Long
    End Type
'***** End Init Declarations *****'
'***** Start In Render Declarations *****'
    Public Matrices() As D3DMATRIX
    Public NumMatrices As Integer
'***** End In Render Declarations *****'
