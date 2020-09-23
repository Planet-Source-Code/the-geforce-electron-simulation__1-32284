VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Event3DEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"Devices"
'Event Engine by Daniel Bloemendal
'If you take any code from this engine you must give me credit!

'This engine is most likely going to be discontinued.
'This is because I am working on a C++ engine.
'If you would like to continue the production of this engine,
'You must get permission from me at teh_leet_programming_man@hotmail.com
'Thank you for downloading my code please give me feedback.
'Vote for me!!!

Private mvarDevices As Devices
Public NumberOfMatrices As Integer

Public Sub SetNewMatrices(ReDimPreserve As Boolean)
    If ReDimPreserve = True Then
        ReDim Preserve Matrices(NumberOfMatrices - 1)
    Else
        ReDim Matrices(NumberOfMatrices - 1)
    End If
End Sub

Public Sub SetMatrix(InMatrixIndex As Integer, OutMatrixIndex As Integer)
    Matrices(OutMatrixIndex) = Matrices(InMatrixIndex)
End Sub

Public Property Get Devices() As Devices
    If mvarDevices Is Nothing Then
        Set mvarDevices = New Devices
    End If


    Set Devices = mvarDevices
End Property

Public Property Set Devices(vData As Devices)
    Set mvarDevices = vData
End Property

Private Sub Class_Terminate()
    Set mvarDevices = Nothing
End Sub

Public Sub Cleanup()
    Set DX8 = Nothing
    Set D3DX = Nothing
    Set D3D = Nothing
    Set D3DDevice = Nothing
End Sub
