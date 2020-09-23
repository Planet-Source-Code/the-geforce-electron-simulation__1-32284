VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CalcThreadCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim Thread As New Thread

Public Sub BeginThread(ObjectIn As Object)
    Thread.CreateWin32Thread ObjectIn, "Calculate"
    Thread.ThreadPriority = THREAD_PRIORITY_HIGHEST
End Sub

Public Function IsRunning() As Boolean
    IsRunning = Thread.IsThreadRunning
End Function

Public Sub StopThread()
    Thread.TerminateWin32Thread
End Sub

