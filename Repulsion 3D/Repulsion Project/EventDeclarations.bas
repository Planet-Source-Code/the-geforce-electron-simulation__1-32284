Attribute VB_Name = "EventDeclarations"
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Event3DEngine As Event3DEngine
Public AdaptersInfo() As AdapterInfo
Public MatCalc As New MatrixCalculator
Public Const PI = 3.14159265358979
Public Rendering As Boolean
Public RLight As Light
Public EventEngineConsole As New EventEngineConsole
Public Repulsion As New Repulsion3DCore
Public NumberOfParticles As Long
Public DegL1 As Long
Public DegL2 As Long

Public Function Vector(X As Single, Y As Single, z As Single) As InVector
    With Vector
        .X = X
        .Y = Y
        .z = z
    End With
End Function

Public Function Rad(Deg As Single) As Single
    Rad = PI / 180 * Deg
End Function
