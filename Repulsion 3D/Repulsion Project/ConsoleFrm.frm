VERSION 5.00
Begin VB.Form ConsoleFrm 
   BackColor       =   &H00000000&
   Caption         =   "Wildfire Console"
   ClientHeight    =   3580
   ClientLeft      =   50
   ClientTop       =   310
   ClientWidth     =   5220
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3580
   ScaleWidth      =   5220
   Begin VB.TextBox ConsoleSendTxt 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   3240
      Width           =   5292
   End
   Begin VB.TextBox Status 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   720
      Left            =   906
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   960
      Width           =   1692
   End
End
Attribute VB_Name = "ConsoleFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents ConsoleEvents As EventEngineConsole
Attribute ConsoleEvents.VB_VarHelpID = -1

Public Sub UpdateSize()
    On Error GoTo ErrorH
    With Status
        .Left = 0
        .Top = 0
        .Width = ScaleWidth
        .Height = ScaleHeight - ConsoleSendTxt.Height
    End With
    With ConsoleSendTxt
        .Left = 0
        .Top = Status.Height
        .Width = ScaleWidth
    End With
  
ErrorH:
End Sub

Private Sub ConsoleEvents_EngineEvent(Message As String)
    Me.Visible = True
    Status.Text = EventEngineConsole.StatusText
    DoEvents
End Sub

Public Sub InitEvents()
    Set ConsoleEvents = EventEngineConsole
End Sub

Public Sub EndEvents()
    Set ConsoleEvents = Nothing
    Me.Visible = False
End Sub

Private Sub ConsoleSendTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        EventEngineConsole.SendCommand ConsoleSendTxt.Text
        ConsoleSendTxt.Text = ""
    End If
End Sub

Private Sub Form_Load()
    UpdateSize
End Sub

Private Sub Form_Resize()
    UpdateSize
End Sub
