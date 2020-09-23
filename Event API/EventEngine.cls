VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EventEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Wildfire3DEngine"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member1" ,"WildfireMusicEngine"
Attribute VB_Ext_KEY = "Member2" ,"WildfirePhysicsEngine"
'Event Engine by Daniel Bloemendal
'If you take any code from this engine you must give me credit!

'This engine is most likely going to be discontinued.
'This is because I am working on a C++ engine.
'If you would like to continue the production of this engine,
'You must get permission from me at teh_leet_programming_man@hotmail.com
'Thank you for downloading my code please give me feedback.
'Vote for me!!!

Private Const Event3DEngineCreationPassword As String = "! 4/\/\ 7|-|3 0/\/3"
Private Const EventMusicEngineCreationPassword As String = "/\/\|_|5!( !5 /\/\`/ 1!|=3"
Private Const EventPhysicsEngineCreationPassword As String = "|=002314|2!5"

Public Function Event3DEngineCreate(Password As String) As Event3DEngine
    If Password = Event3DEngineCreationPassword Then
        Set Event3DEngineCreate = New Event3DEngine
    Else
        MsgBox "Incorrect creation perimiters!", vbCritical, "Event 3D Engine"
    End If
End Function

Public Function EventMusicEngineCreate(Password As String) As EventMusicEngine
    If Password = EventMusicEngineCreationPassword Then
        Set EventMusicEngineCreate = New EventMusicEngine
    Else
        MsgBox "Incorrect creation perimiters!", vbCritical, "Event Music Engine"
    End If
End Function

Public Function EventPhysicsEngineCreate(Password As String) As EventPhysicsEngine
    If Password = EventPhysicsEngineCreationPassword Then
        Set EventPhysicsEngineCreate = New EventPhysicsEngine
    Else
        MsgBox "Incorrect creation perimiters!", vbCritical, "Event Physics Engine"
    End If
End Function
