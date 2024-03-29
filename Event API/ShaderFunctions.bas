Attribute VB_Name = "ShaderFunctions"
'Event Engine by Daniel Bloemendal
'If you take any code from this engine you must give me credit!

'This engine is most likely going to be discontinued.
'This is because I am working on a C++ engine.
'If you would like to continue the production of this engine,
'You must get permission from me at teh_leet_programming_man@hotmail.com
'Thank you for downloading my code please give me feedback.
'Vote for me!!!

Global Const D3DSI_COMMENTSIZE_SHIFT = 16
Global Const D3DSI_COMMENTSIZE_MASK = &H7FFF0000
Global Const D3DVS_INPUTREG_MAX_V1_1 = 16
Global Const D3DVS_TEMPREG_MAX_V1_1 = 12
Global Const D3DVS_CONSTREG_MAX_V1_1 = 96
Global Const D3DVS_TCRDOUTREG_MAX_V1_1 = 8
Global Const D3DVS_ADDRREG_MAX_V1_1 = 1
Global Const D3DVS_ATTROUTREG_MAX_V1_1 = 2
Global Const D3DVS_MAXINSTRUCTIONCOUNT_V1_1 = 128
Global Const D3DPS_INPUTREG_MAX_DX8 = 8
Global Const D3DPS_TEMPREG_MAX_DX8 = 8
Global Const D3DPS_CONSTREG_MAX_DX8 = 8
Global Const D3DPS_TEXTUREREG_MAX_DX8 = 8
Enum D3DVSD_TOKENTYPE
    D3DVSD_TOKEN_NOP = 0
    D3DVSD_TOKEN_STREAM = 1
    D3DVSD_TOKEN_STREAMDATA = 2
    D3DVSD_TOKEN_TESSELLATOR = 3
    D3DVSD_TOKEN_constMEM = 4
    D3DVSD_TOKEN_EXT = 5
    D3DVSD_TOKEN_END = 7
End Enum
Global Const D3DVSD_TOKENTYPESHIFT = 29
Global Const D3DVSD_TOKENTYPEMASK = &HE0000000
Global Const D3DVSD_STREAMNUMBERSHIFT = 0
Global Const D3DVSD_STREAMNUMBERMASK = &HF&
Global Const D3DVSD_DATALOADTYPESHIFT = 28
Global Const D3DVSD_DATALOADTYPEMASK = &H10000000
Global Const D3DVSD_DATATYPESHIFT = 16
Global Const D3DVSD_DATATYPEMASK = &HF& * 2 ^ D3DVSD_DATATYPESHIFT
Global Const D3DVSD_SKIPCOUNTSHIFT = 16
Global Const D3DVSD_SKIPCOUNTMASK = &HF& * 2 ^ D3DVSD_SKIPCOUNTSHIFT
Global Const D3DVSD_VERTEXREGSHIFT = 0
Global Const D3DVSD_VERTEXREGMASK = &HF& * 2 ^ D3DVSD_VERTEXREGSHIFT
Global Const D3DVSD_VERTEXREGINSHIFT = 20
Global Const D3DVSD_VERTEXREGINMASK = &HF& * 2 ^ D3DVSD_VERTEXREGINSHIFT
Global Const D3DVSD_CONSTCOUNTSHIFT = 25
Global Const D3DVSD_CONSTCOUNTMASK = &HF& * 2 ^ D3DVSD_CONSTCOUNTSHIFT
Global Const D3DVSD_CONSTADDRESSSHIFT = 0
Global Const D3DVSD_CONSTADDRESSMASK = &H7F&
Global Const D3DVSD_CONSTRSSHIFT = 16
Global Const D3DVSD_CONSTRSMASK = &H1FFF0000
Global Const D3DVSD_EXTCOUNTSHIFT = 24
Global Const D3DVSD_EXTCOUNTMASK = &H1F& * 2 ^ D3DVSD_EXTCOUNTSHIFT
Global Const D3DVSD_EXTINFOSHIFT = 0
Global Const D3DVSD_EXTINFOMASK = &HFFFFFF
Global Const D3DVSDT_FLOAT1 = 0&
Global Const D3DVSDT_FLOAT2 = 1&
Global Const D3DVSDT_FLOAT3 = 2&
Global Const D3DVSDT_FLOAT4 = 3&
Global Const D3DVSDT_D3DCOLOR = 4&
Global Const D3DVSDT_UBYTE4 = 5&
Global Const D3DVSDT_SHORT2 = 6&
Global Const D3DVSDT_SHORT4 = 7&
Global Const D3DVSDE_POSITION = 0&
Global Const D3DVSDE_BLENDWEIGHT = 1&
Global Const D3DVSDE_BLENDINDICES = 2&
Global Const D3DVSDE_NORMAL = 3&
Global Const D3DVSDE_PSIZE = 4&
Global Const D3DVSDE_DIFFUSE = 5&
Global Const D3DVSDE_SPECULAR = 6&
Global Const D3DVSDE_TEXCOORD0 = 7&
Global Const D3DVSDE_TEXCOORD1 = 8&
Global Const D3DVSDE_TEXCOORD2 = 9&
Global Const D3DVSDE_TEXCOORD3 = 10&
Global Const D3DVSDE_TEXCOORD4 = 11&
Global Const D3DVSDE_TEXCOORD5 = 12&
Global Const D3DVSDE_TEXCOORD6 = 13&
Global Const D3DVSDE_TEXCOORD7 = 14&
Global Const D3DVSDE_POSITION2 = 15&
Global Const D3DVSDE_NORMAL2 = 16&
Global Const D3DDP_MAXTEXCOORD = 8
Global Const D3DSI_OPCODE_MASK = &HFFFF&
Global Const D3DSI_COISSUE = &H40000000
Global Const D3DSP_REGNUM_MASK = &HFFF&
Global Const D3DSP_WRITEMASK_0 = &H10000
Global Const D3DSP_WRITEMASK_1 = &H20000
Global Const D3DSP_WRITEMASK_2 = &H40000
Global Const D3DSP_WRITEMASK_3 = &H80000
Global Const D3DSP_WRITEMASK_ALL = &HF0000
Global Const D3DSP_DSTMOD_SHIFT = 20
Global Const D3DSP_DSTMOD_MASK = &HF00000
Enum D3DSHADER_PARAM_DSTMOD_TYPE
    D3DSPDM_NONE = 0 * 2 ^ D3DSP_DSTMOD_SHIFT
    D3DSPDM_SATURATE = 1 * 2 ^ D3DSP_DSTMOD_SHIFT
End Enum
Global Const D3DSP_DSTSHIFT_SHIFT = 24
Global Const D3DSP_DSTSHIFT_MASK = &HF000000
Global Const D3DSP_REGTYPE_SHIFT = 28
Global Const D3DSP_REGTYPE_MASK = &H70000000
Global Const D3DVSD_STREAMTESSSHIFT = 28
Global Const D3DVSD_STREAMTESSMASK = 2 ^ D3DVSD_STREAMTESSSHIFT
Enum D3DSHADER_PARAM_REGISTER_TYPE
    D3DSPR_TEMP = &H0&
    D3DSPR_INPUT = &H20000000
    D3DSPR_CONST = &H40000000
    D3DSPR_ADDR = &H60000000
    D3DSPR_TEXTURE = &H60000000
    D3DSPR_RASTOUT = &H80000000
    D3DSPR_ATTROUT = &HA0000000
    D3DSPR_TEXCRDOUT = &HC0000000
End Enum
Enum D3DVS_RASTOUT_OFFSETS
    D3DSRO_POSITION = 0
    D3DSRO_FOG = 1
    D3DSRO_POINT_SIZE = 2
End Enum
Global Const D3DVS_ADDRESSMODE_SHIFT = 13
Global Const D3DVS_ADDRESSMODE_MASK = (2 ^ D3DVS_ADDRESSMODE_SHIFT)
Enum D3DVS_ADRRESSMODE_TYPE
    D3DVS_ADDRMODE_ABSOLUTE = 0
    D3DVS_ADDRMODE_RELATIVE = 2 ^ D3DVS_ADDRESSMODE_SHIFT
End Enum
Global Const D3DVS_SWIZZLE_SHIFT = 16
Global Const D3DVS_SWIZZLE_MASK = &HFF0000
Global Const D3DVS_X_X = (0 * 2 ^ D3DVS_SWIZZLE_SHIFT)
Global Const D3DVS_X_Y = (1 * 2 ^ D3DVS_SWIZZLE_SHIFT)
Global Const D3DVS_X_Z = (2 * 2 ^ D3DVS_SWIZZLE_SHIFT)
Global Const D3DVS_X_W = (3 * 2 ^ D3DVS_SWIZZLE_SHIFT)
Global Const D3DVS_Y_X = (0 * 2 ^ (D3DVS_SWIZZLE_SHIFT + 2))
Global Const D3DVS_Y_Y = (1 * 2 ^ (D3DVS_SWIZZLE_SHIFT + 2))
Global Const D3DVS_Y_Z = (2 * 2 ^ (D3DVS_SWIZZLE_SHIFT + 2))
Global Const D3DVS_Y_W = (3 * 2 ^ (D3DVS_SWIZZLE_SHIFT + 2))
Global Const D3DVS_Z_X = (0 * 2 ^ (D3DVS_SWIZZLE_SHIFT + 4))
Global Const D3DVS_Z_Y = (1 * 2 ^ (D3DVS_SWIZZLE_SHIFT + 4))
Global Const D3DVS_Z_Z = (2 * 2 ^ (D3DVS_SWIZZLE_SHIFT + 4))
Global Const D3DVS_Z_W = (3 * 2 ^ (D3DVS_SWIZZLE_SHIFT + 4))
Global Const D3DVS_W_X = (0 * 2 ^ (D3DVS_SWIZZLE_SHIFT + 6))
Global Const D3DVS_W_Y = (1 * 2 ^ (D3DVS_SWIZZLE_SHIFT + 6))
Global Const D3DVS_W_Z = (2 * 2 ^ (D3DVS_SWIZZLE_SHIFT + 6))
Global Const D3DVS_W_W = (3 * 2 ^ (D3DVS_SWIZZLE_SHIFT + 6))
Global Const D3DVS_NOSWIZZLE = (D3DVS_X_X Or D3DVS_Y_Y Or D3DVS_Z_Z Or D3DVS_W_W)
Global Const D3DSP_SWIZZLE_SHIFT = 16
Global Const D3DSP_SWIZZLE_MASK = &HFF0000
Global Const D3DSP_NOSWIZZLE = ((0 * 2 ^ (D3DSP_SWIZZLE_SHIFT + 0)) Or (1 * 2 ^ (D3DSP_SWIZZLE_SHIFT + 2)) Or (2 * 2 ^ (D3DSP_SWIZZLE_SHIFT + 4)) Or (3 * 2 ^ (D3DSP_SWIZZLE_SHIFT + 6)))
Global Const D3DSP_REPLICATEALPHA = ((3 * 2 ^ (D3DSP_SWIZZLE_SHIFT + 0)) Or (3 * 2 ^ (D3DSP_SWIZZLE_SHIFT + 2)) Or (3 * 2 ^ (D3DSP_SWIZZLE_SHIFT + 4)) Or (3 * 2 ^ (D3DSP_SWIZZLE_SHIFT + 6)))
Global Const D3DSP_SRCMOD_SHIFT = 24
Global Const D3DSP_SRCMOD_MASK = &HF000000
Enum D3DSHADER_PARAM_SRCMOD_TYPE
    D3DSPSM_NONE = 0 * 2 ^ D3DSP_SRCMOD_SHIFT
    D3DSPSM_NEG = 1 * 2 ^ D3DSP_SRCMOD_SHIFT
    D3DSPSM_BIAS = 2 * 2 ^ D3DSP_SRCMOD_SHIFT
    D3DSPSM_BIASNEG = 3 * 2 ^ D3DSP_SRCMOD_SHIFT
    D3DSPSM_SIGN = 4 * 2 ^ D3DSP_SRCMOD_SHIFT
    D3DSPSM_SIGNNEG = 5 * 2 ^ D3DSP_SRCMOD_SHIFT
    D3DSPSM_COMP = 6 * 2 ^ D3DSP_SRCMOD_SHIFT
End Enum

Function D3DPS_VERSION(Major As Long, Minor As Long) As Long
    D3DPS_VERSION = (&HFFFF0000 Or ((Major) * 2 ^ 8) Or (Minor))
End Function

Function D3DVS_VERSION(Major As Long, Minor As Long) As Long
    D3DVS_VERSION = (&HFFFE0000 Or ((Major) * 2 ^ 8) Or (Minor))
End Function

Function D3DSHADER_VERSION_MAJOR(Version As Long) As Long
    D3DSHADER_VERSION_MAJOR = (((Version) \ 8) And &HFF&)
End Function
    
Function D3DSHADER_VERSION_MINOR(Version As Long) As Long
    D3DSHADER_VERSION_MINOR = (((Version)) And &HFF&)
End Function

Function D3DSHADER_COMMENT(DWordSize As Long) As Long
    D3DSHADER_COMMENT = ((((DWordSize) * 2 ^ D3DSI_COMMENTSIZE_SHIFT) And D3DSI_COMMENTSIZE_MASK) Or D3DSIO_COMMENT)
End Function

Function D3DPS_END() As Long
    D3DPS_END = &HFFFF&
End Function

Function D3DVS_END() As Long
   D3DVS_END = &HFFFF&
End Function

Function D3DVSD_MAKETOKENTYPE(tokenType As Long) As Long
    Dim out As Long
    Select Case tokenType
        Case D3DVSD_TOKEN_NOP
            out = 0
        Case D3DVSD_TOKEN_STREAM
            out = &H20000000
        Case D3DVSD_TOKEN_STREAMDATA
            out = &H40000000
        Case D3DVSD_TOKEN_TESSELLATOR
            out = &H60000000
        Case D3DVSD_TOKEN_constMEM
            out = &H80000000
        Case D3DVSD_TOKEN_EXT
            out = &HA0000000
        Case D3DVSD_TOKEN_END
            out = &HFFFFFFFF
    End Select
    D3DVSD_MAKETOKENTYPE = out And D3DVSD_TOKENTYPEMASK
End Function

Function D3DVSD_STREAM(StreamNumber As Long) As Long
    D3DVSD_STREAM = (D3DVSD_MAKETOKENTYPE(D3DVSD_TOKEN_STREAM) Or (StreamNumber))
End Function

Function D3DVSD_STREAM_TESS() As Long
    D3DVSD_STREAM_TESS = (D3DVSD_MAKETOKENTYPE(D3DVSD_TOKEN_STREAM) Or (D3DVSD_STREAMTESSMASK))
End Function
    
Function D3DVSD_REG(VertexRegister As Long, dataType As Long) As Long
    D3DVSD_REG = (D3DVSD_MAKETOKENTYPE(D3DVSD_TOKEN_STREAMDATA) Or _
     ((dataType) * 2 ^ D3DVSD_DATATYPESHIFT) Or (VertexRegister))
End Function

Function D3DVSD_SKIP(DWORDCount As Long) As Long
    D3DVSD_SKIP = (D3DVSD_MAKETOKENTYPE(D3DVSD_TOKEN_STREAMDATA) Or &H10000000 Or _
     ((DWORDCount) * 2 ^ D3DVSD_SKIPCOUNTSHIFT))
End Function

Function D3DVSD_CONST(constantAddress As Long, Count As Long) As Long
    D3DVSD_CONST = (D3DVSD_MAKETOKENTYPE(D3DVSD_TOKEN_constMEM) Or _
     ((Count) * 2 ^ D3DVSD_CONSTCOUNTSHIFT) Or (constantAddress))
End Function

Function D3DVSD_TESSNORMAL(VertexRegisterIn As Long, VertexRegisterOut As Long) As Long
    D3DVSD_TESSNORMAL = (D3DVSD_MAKETOKENTYPE(D3DVSD_TOKEN_TESSELLATOR) Or _
     ((VertexRegisterIn) * 2 ^ D3DVSD_VERTEXREGINSHIFT) Or _
     ((&H2&) * 2 ^ D3DVSD_DATATYPESHIFT) Or (VertexRegisterOut))
End Function

Function D3DVSD_TESSUV(VertexRegister As Long) As Long
    D3DVSD_TESSUV = (D3DVSD_MAKETOKENTYPE(D3DVSD_TOKEN_TESSELLATOR) Or &H10000000 Or _
     ((&H1&) * 2 ^ D3DVSD_DATATYPESHIFT) Or (VertexRegister))
End Function

Function D3DVSD_END() As Long
        D3DVSD_END = &HFFFFFFFF
End Function

Function D3DVSD_NOP() As Long
    D3DVSD_NOP = 0
End Function
