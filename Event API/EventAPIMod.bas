Attribute VB_Name = "EventAPIMod"
'Event Engine by Daniel Bloemendal
'If you take any code from this engine you must give me credit!

'This engine is most likely going to be discontinued.
'This is because I am working on a C++ engine.
'If you would like to continue the production of this engine,
'You must get permission from me at teh_leet_programming_man@hotmail.com
'Thank you for downloading my code please give me feedback.
'Vote for me!!!

Public DX8 As New DirectX8
Public D3DX As New D3DX8
Public D3D As Direct3D8

Public Sub Main()
    Set D3D = DX8.Direct3DCreate
End Sub

Public Function LightToD3DLight8(Light As Light) As D3DLIGHT8
    With LightToD3DLight8
        .ambient = ColorValueToD3DColorValue(Light.ambient)
        .diffuse = ColorValueToD3DColorValue(Light.diffuse)
        .specular = ColorValueToD3DColorValue(Light.specular)
        .type = LightTypeToD3DLightType(Light.type)
        .Attenuation0 = Light.Attenuation0
        .Attenuation1 = Light.Attenuation1
        .Attenuation2 = Light.Attenuation2
        .Direction = VectorToD3DVector(Light.Direction)
        .Falloff = Light.Falloff
        .Phi = Light.Phi
        .position = VectorToD3DVector(Light.position)
        .Range = Light.Range
        .Theta = Light.Theta
    End With
End Function

Public Function LightTypeToD3DLightType(LightType As LightType) As CONST_D3DLIGHTTYPE
    Select Case LightType
            
        Case DirectionalLight
            LightTypeToD3DLightType = D3DLIGHT_DIRECTIONAL
            
        Case PointLight
            LightTypeToD3DLightType = D3DLIGHT_POINT
            
        Case SpotLight
            LightTypeToD3DLightType = D3DLIGHT_SPOT
            
    End Select
End Function

Public Function MaterialToD3DMaterial(Material As Material) As D3DMATERIAL8
    With MaterialToD3DMaterial
        .ambient = ColorValueToD3DColorValue(Material.ambient)
        .diffuse = ColorValueToD3DColorValue(Material.diffuse)
        .emissive = ColorValueToD3DColorValue(Material.emissive)
        .power = Material.power
        .specular = ColorValueToD3DColorValue(Material.specular)
    End With
End Function

Public Function ColorValueToD3DColorValue(ColorValue As ColorValue) As D3DCOLORVALUE
    With ColorValueToD3DColorValue
        .A = ColorValue.A
        .b = ColorValue.b
        .g = ColorValue.g
        .r = ColorValue.r
    End With
End Function

Public Function VectorToD3DVector(Vector As Vector) As D3DVECTOR
    With VectorToD3DVector
        .x = Vector.x
        .y = Vector.y
        .z = Vector.z
    End With
End Function
