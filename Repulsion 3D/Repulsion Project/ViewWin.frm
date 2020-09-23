VERSION 5.00
Begin VB.Form ViewWin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Window"
   ClientHeight    =   3480
   ClientLeft      =   30
   ClientTop       =   320
   ClientWidth     =   5520
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   18
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   5520
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "ViewWin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And Rendering = True Then
        DisableRendering
        EnableMenus
        Me.Visible = False
    End If
End Sub

Public Sub UpdateSize()
    Dim InitDat As InitDeviceData
    InitDat = LoadInitDeviceData
    With AdaptersInfo(InitDat.AdapterIndex).DevTypeInfo(InitDat.DeviceType).Modes(InitDat.Resolution)
        Me.Width = ScaleY(.Width, vbPixels, vbTwips)
        Me.Height = ScaleX(.Width, vbPixels, vbTwips)
    End With
End Sub
