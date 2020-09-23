VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EventPhysicsEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"CollisionDetection"
'Event Engine by Daniel Bloemendal
'If you take any code from this engine you must give me credit!

'This engine is most likely going to be discontinued.
'This is because I am working on a C++ engine.
'If you would like to continue the production of this engine,
'You must get permission from me at teh_leet_programming_man@hotmail.com
'Thank you for downloading my code please give me feedback.
'Vote for me!!!

'***** Start Type Declarations *****'
    Public Type Vector
            x As Single
            y As Single
            z As Single
    End Type
    Public Type VectorFaucet
        Active As Boolean
        CreationTime As Double
        position As Vector
        m As Single
    End Type
    Public Type PhysicsObject
        VectorFaucets() As VectorFaucet
        NumberOfVectorFaucets As Long
    End Type
'***** End Type Declarations *****'
'***** Start Public Declarations *****'
    Public GlobalTime As Double
'***** End Public Declarations *****'
'***** Start Private Declarations *****'
    Private PhysicsObjects() As PhysicsObject
    Private NumberOfPhysicsObjects As Long
'***** End Private Declarations *****'

Public Sub AddPhyisicsObjectToStream(PhysicsObjectIn As PhysicsObject)
    ReDim Preserve PhysicsObjects(NumberOfPhysicsObjects)
    PhysicsObjects(NumberOfPhysicsObjects) = PhysicsObjectIn
    NumberOfPhysicsObjects = NumberOfPhysicsObjects + 1
End Sub

Private Function f(t As Double, A As D3DVECTOR, V As D3DVECTOR, i As D3DVECTOR) As D3DVECTOR
    f.x = 0.5 * A.x * t ^ 2 + V.x * t + i.x
    f.y = 0.5 * A.y * t ^ 2 + V.y * t + i.y
    f.z = 0.5 * A.z * t ^ 2 + V.z * t + i.z
End Function
