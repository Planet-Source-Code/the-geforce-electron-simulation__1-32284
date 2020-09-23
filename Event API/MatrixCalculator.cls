VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MatrixCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
'Event Engine by Daniel Bloemendal
'If you take any code from this engine you must give me credit!

'This engine is most likely going to be discontinued.
'This is because I am working on a C++ engine.
'If you would like to continue the production of this engine,
'You must get permission from me at teh_leet_programming_man@hotmail.com
'Thank you for downloading my code please give me feedback.
'Vote for me!!!

Public Type InVector
    x As Single
    y As Single
    z As Single
End Type
Public Enum TransformationOrder
    Before = 0
    After = 1
End Enum
Public InMatrixIndex As Integer
Public OutMatrixIndex As Integer
Private MidCalcMatrix As D3DMATRIX
Private MatrixTransformation As D3DMATRIX

Private Sub Class_Initialize()
    D3DXMatrixIdentity MatrixTransformation
End Sub

Public Sub MatrixIdentity(MatrixIndex As Integer)
    D3DXMatrixIdentity Matrices(Index)
End Sub

Public Sub MatrixMultiply(OutMatrixIndex As Integer, Matrix1Index As Integer, Matrix2Index As Integer)
    D3DXMatrixMultiply Matrices(OutMatrixIndex), Matrices(Matrix1Index), Matrices(Matrix2Index)
End Sub

Public Sub MatrixScale(ScaleX As Single, ScaleY As Single, ScaleZ As Single, Order As TransformationOrder)
    D3DXMatrixScaling MidCalcMatrix, ScaleX, ScaleY, ScaleZ
    OrderTransform Order
End Sub

Public Sub MatrixTranslation(x As Single, y As Single, z As Single, Order As TransformationOrder)
    D3DXMatrixTranslation MidCalcMatrix, x, y, z
    OrderTransform Order
End Sub

Public Sub MatrixYawPitchRoll(yaw As Single, pitch As Single, roll As Single, Order As TransformationOrder)
    D3DXMatrixRotationYawPitchRoll MidCalcMatrix, yaw, pitch, roll
    OrderTransform Order
End Sub

Public Sub MatrixLookAtLH(VEye As InVector, VAt As InVector, VUp As InVector, Order As TransformationOrder)
    D3DXMatrixLookAtLH MidCalcMatrix, Vector(VEye.x, VEye.y, VEye.z), Vector(VAt.x, VAt.y, VAt.z), Vector(VUp.x, VUp.y, VUp.z)
    OrderTransform Order
End Sub

Private Sub OrderTransform(Order As TransformationOrder)
    If Order = Before Then
        D3DXMatrixMultiply MatrixTransformation, MidCalcMatrix, MatrixTransformation
    ElseIf Order = After Then
        D3DXMatrixMultiply MatrixTransformation, MatrixTransformation, MidCalcMatrix
    End If
End Sub

Public Sub Transform(Order As TransformationOrder)
    If Order = Before Then
        D3DXMatrixMultiply Matrices(OutMatrixIndex), MatrixTransformation, Matrices(InMatrixIndex)
    ElseIf Order = After Then
        D3DXMatrixMultiply Matrices(OutMatrixIndex), Matrices(InMatrixIndex), MatrixTransformation
    End If
End Sub

Public Sub ResetTransform()
    D3DXMatrixIdentity MatrixTransformation
End Sub

Private Function Vector(x As Single, y As Single, z As Single) As D3DVECTOR
    Vector.x = x
    Vector.y = y
    Vector.z = z
End Function
