VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Mesh"
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

Private Type TextureIndexedMaterials
    Material As D3DMATERIAL8
    TextureIndex As Long
End Type
Public D3DMesh As D3DXMesh
Public NumberOfMaterials As Long
Public NumberOfTextures As Long
Private Materials() As TextureIndexedMaterials
Private Textures() As Direct3DTexture8
Private D3DDevice As Direct3DDevice8

Public Sub InitFromFile(File As String, D3DDeviceIn As Direct3DDevice8)
    Dim i As Long
    Dim MaterialBuffer As D3DXBuffer
    Dim TextureFile As String
    Set D3DDevice = D3DDeviceIn
    Set D3DMesh = D3DX.LoadMeshFromX(File, D3DXMESH_MANAGED, D3DDeviceIn, Nothing, MaterialBuffer, NumberOfMaterials)
    ReDim Materials(NumberOfMaterials - 1)
    ReDim Textures(NumberOfMaterials - 1)
    For i = 0 To NumberOfMaterials - 1
        D3DX.BufferGetMaterial MaterialBuffer, i, Materials(i).Material
        Materials(i).Material.ambient = Materials(i).Material.diffuse
        TextureFile = D3DX.BufferGetTextureName(MaterialBuffer, i)
        If TextureFile <> "" Then
            Set Textures(NumberOfTextures) = D3DX.CreateTextureFromFile(D3DDevice, TextureFile)
            NumberOfTextures = NumberOfTextures + 1
            Materials(i).TextureIndex = i
        Else
            Materials(i).TextureIndex = -1
        End If
    Next i
    If NumberOfMaterials = 0 Then
        NumberOfMaterials = 1
        Materials(NumberOfMaterials).Material = DefaultMat
        Materials(NumberOfMaterials).TextureIndex = -1
    End If
    If NumberOfTextures = 0 Then NumberOfTextures = 1
    If NumberOfTextures > 0 Then ReDim Preserve Textures(NumberOfTextures - 1)
End Sub

Public Sub InitFromText(Text As String, hDC As Long, Deviation As Single, Extrusion As Single, Material As Material, D3DDeviceIn As Direct3DDevice8, Optional TextureFile As String)
    Set D3DDevice = D3DDeviceIn
    D3DX.CreateText D3DDevice, hDC, Text, Deviation, Extrusion, D3DMesh, Nothing, ByVal 0
    ReDim Materials(0)
    ReDim Textures(0)
    NumberOfMaterials = 1
    NumberOfTextures = 1
    Materials(NumberOfMaterials - 1).Material = MaterialToD3DMaterial(Material)
    Materials(NumberOfMaterials - 1).TextureIndex = 0
    If TextureFile <> "" Then
        Set Textures(0) = D3DX.CreateTextureFromFile(D3DDevice, TextureFile)
    End If
End Sub

Public Function GetMaterial(Index As Long) As D3DMATERIAL8
    GetMaterial = Materials(Index).Material
End Function

Public Function SetMaterial(Index As Long, Material As Material)
    Materials(Index).Material = MaterialToD3DMaterial(Material)
End Function

Public Function GetTexture(Index As Long) As Direct3DTexture8
    Set GetTexture = Textures(Index)
End Function

Public Function GetMaterialTextureIndex(Index As Long) As Long
    GetMaterialTextureIndex = Materials(Index).TextureIndex
End Function

Public Function DefaultMat() As D3DMATERIAL8
    Dim col As D3DCOLORVALUE
    With col: .A = 1: .b = 1: .g = 1: .r = 1: End With
    With DefaultMat
        .diffuse = col
        .ambient = col
    End With
End Function
