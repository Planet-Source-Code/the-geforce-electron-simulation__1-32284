VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Frame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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

Public Enum DrawSorting
    DrawOpaque = 1
    DrawAlpha = 2
    DrawBoth = 3
End Enum
Public ChildFrames As New Frames
Public ChildMesh As Mesh
Public MeshSubset As Long
Public HideSubset As Boolean
Public AlphaBlend As Boolean
Public VertexShaderHandle As Long
Public VertexShadeEnable As Boolean
Public PixelShaderHandle As Long
Public PixelShadeEnable As Boolean
Private WorldMatrix As D3DMATRIX
Private FrameMatrix As D3DMATRIX

Private Sub Class_Initialize()
    D3DXMatrixIdentity WorldMatrix
    D3DXMatrixIdentity FrameMatrix
End Sub

Public Sub SetWorldMatrix(Matrix As D3DMATRIX)
    WorldMatrix = Matrix
End Sub

Public Function GetFrameMatrix() As D3DMATRIX
    GetFrameMatrix = FrameMatrix
End Function

Public Sub SetFrameMatrix(MatrixIndex As Integer)
    FrameMatrix = Matrices(MatrixIndex)
End Sub

Public Sub TransformAndDrawFrame(D3DDeviceIn As Direct3DDevice8, Optional DrawSort As DrawSorting = 3)
    Dim i As Long
    D3DXMatrixMultiply WorldMatrix, FrameMatrix, WorldMatrix
    D3DDeviceIn.SetTransform D3DTS_WORLD, WorldMatrix
    If ((AlphaBlend = True And DrawSort = DrawAlpha) Or (AlphaBlend = False And DrawSort = DrawOpaque) Or (DrawSort = DrawBoth)) And HideSubset = False Then
        If AlphaBlend = True And DrawSort <> DrawOpaque Then
            D3DDeviceIn.SetRenderState D3DRS_ALPHABLENDENABLE, 1
            D3DDeviceIn.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
            D3DDeviceIn.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        Else
            D3DDeviceIn.SetRenderState D3DRS_ALPHABLENDENABLE, 0
        End If
        With ChildMesh
            If MeshSubset < .NumberOfMaterials Then
                D3DDeviceIn.SetMaterial .GetMaterial(MeshSubset)
                If .GetMaterialTextureIndex(MeshSubset) > -1 Then
                    D3DDeviceIn.SetTexture 0, .GetTexture(.GetMaterialTextureIndex(MeshSubset))
                    D3DDeviceIn.SetTextureStageState 0, D3DTSS_ADDRESSU, D3DTADDRESS_WRAP
                    D3DDeviceIn.SetTextureStageState 0, D3DTSS_ADDRESSV, D3DTADDRESS_WRAP
                End If
            End If
                If VertexShadeEnable = True Then
                    D3DDeviceIn.SetVertexShader VertexShaderHandle
                End If
                If PixelShadeEnable = True Then
                    D3DDeviceIn.SetPixelShader PixelShaderHandle
                End If
                'D3DDeviceIn.SetStreamSource
                .D3DMesh.DrawSubset MeshSubset
        End With
    End If
    For i = 1 To ChildFrames.Count
        With ChildFrames.Item(i)
            .SetWorldMatrix WorldMatrix
            .TransformAndDrawFrame D3DDeviceIn, DrawSort
        End With
    Next i
    D3DXMatrixIdentity WorldMatrix
    D3DDeviceIn.SetRenderState D3DRS_ALPHABLENDENABLE, 0
End Sub
