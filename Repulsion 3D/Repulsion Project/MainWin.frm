VERSION 5.00
Begin VB.MDIForm MainWin 
   BackColor       =   &H8000000C&
   Caption         =   "Event Engine"
   ClientHeight    =   6330
   ClientLeft      =   50
   ClientTop       =   320
   ClientWidth     =   8230
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox BackgroundImage 
      Align           =   2  'Align Bottom
      Height          =   576
      Left            =   0
      Picture         =   "MainWin.frx":0000
      ScaleHeight     =   540
      ScaleWidth      =   8190
      TabIndex        =   0
      Top             =   5750
      Visible         =   0   'False
      Width           =   8232
   End
   Begin VB.Menu Game_Mnu 
      Caption         =   "Game"
      Visible         =   0   'False
      Begin VB.Menu Strt_Itm 
         Caption         =   "Start Demo"
      End
      Begin VB.Menu Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu Exit_Itm 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu Options_Mnu 
      Caption         =   "Options"
      Visible         =   0   'False
      Begin VB.Menu EngineConfiguration_Itm 
         Caption         =   "Engine Configuration"
      End
   End
End
Attribute VB_Name = "MainWin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub EngineConfiguration_Itm_Click()
    EngineConfigurationFrm.Visible = True
End Sub

Private Sub Exit_Itm_Click()
    Cleanup
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Cleanup
End Sub

Private Sub Strt_Itm_Click()
    Init3D
End Sub
