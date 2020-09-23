VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EventMusicEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"Segments"
'Event Engine by Daniel Bloemendal
'If you take any code from this engine you must give me credit!

'This engine is most likely going to be discontinued.
'This is because I am working on a C++ engine.
'If you would like to continue the production of this engine,
'You must get permission from me at teh_leet_programming_man@hotmail.com
'Thank you for downloading my code please give me feedback.
'Vote for me!!!

Implements DirectXEvent8
Public Enum MusicEvents
    SegmentEnded = 0
End Enum
Private DMEvent As Long
Private DirectMusicPerformance As DirectMusicPerformance8
Private DirectMusicLoader As DirectMusicLoader8
Private AudioParams As DMUS_AUDIOPARAMS
Private hFocusWindow As Long
Private InitTrue As Boolean
Private NumSegments As Long
Private Segments() As DirectMusicSegment8
Private Playing As Boolean
Private SegPlaying As Long
'To fire this event, use RaiseEvent with the following syntax:
'RaiseEvent MusicEvent[(arg1, arg2, ... , argn)]
Public Event MusicEvent(InEvent As MusicEvents)

Private Sub Class_Initialize()
    Set DirectMusicPerformance = DX8.DirectMusicPerformanceCreate
    Set DirectMusicLoader = DX8.DirectMusicLoaderCreate
End Sub

Public Sub InitAudio(hWnd As Long)
    hFocusWindow = hWnd
    DirectMusicPerformance.InitAudio hFocusWindow, DMUS_AUDIOF_ALL, AudioParams, Nothing, DMUS_APATH_SHARED_STEREOPLUSREVERB, 64
    DirectMusicPerformance.SetMasterAutoDownload True
    DirectMusicPerformance.AddNotificationType DMUS_NOTIFY_ON_SEGMENT
    DMEvent = DX8.CreateEvent(Me)
    DirectMusicPerformance.SetNotificationHandle DMEvent
    InitTrue = True
End Sub

Public Sub SetSearchDirectory(SearchDirectory As String)
    If InitTrue = False Then Err.Raise 4
    DirectMusicLoader.SetSearchDirectory SearchDirectory
End Sub

Public Sub CreateAudioSegment(File As String)
    If InitTrue = False Then Err.Raise 4
    ReDim Preserve Segments(NumSegments)
    Set Segments(NumSegments) = DirectMusicLoader.LoadSegment(File)
    Segments(NumSegments).Download DirectMusicPerformance.GetDefaultAudioPath
    If UCase(Right(File, 3)) = "MID" Then
        Segments(NumSegments).SetStandardMidiFile
    End If
    NumSegments = NumSegments + 1
End Sub

Public Sub PlayAudioSegment(SegmentIndex As Long, Optional Repeats As Long = 0)
    If InitTrue = False Then Err.Raise 4
    If Playing = True Then StopAudioSegment SegPlaying
    Segments(SegmentIndex).SetRepeats Repeats
    DirectMusicPerformance.PlaySegmentEx Segments(SegmentIndex), DMUS_SEGF_DEFAULT, 0
    SegPlaying = SegmentIndex
    Playing = True
End Sub

Public Sub StopAudioSegment(SegmentIndex As Long)
    If InitTrue = False Then Err.Raise 4
    DirectMusicPerformance.StopEx Segments(SegmentIndex), 0, 0
    DirectMusicPerformance.Reset 0
    Playing = False
End Sub

Public Sub RemoveAudioSegment(SegmentIndex As Long)
    If InitTrue = False Then Err.Raise 4
    If Playing = True Then
        If SegPlaying = SegmentIndex Then StopAudioSegment SegPlaying
    End If
    Dim SegI As Long
    If NumSegments > 0 Then
        For SegI = SegmentIndex To NumSegments - 2
            Set Segments(SegI) = Segments(SegI + 1)
        Next SegI
        NumSegments = NumSegments - 1
        ReDim Preserve Segments(NumSegments)
    Else
        Err.Raise 4
    End If
End Sub

Public Function NumberOfSegments()
    If InitTrue = False Then Err.Raise 4
    NumberOfSegments = NumSegments
End Function

Public Function GetSegmentLength(SegmentIndex As Long)
    If InitTrue = False Then Err.Raise 4
    GetSegmentLength = Segments(SegmentIndex).GetLength
End Function

Public Function GetElapsedTime()
    If InitTrue = False Then Err.Raise 4
    Dim SegmentState As DirectMusicSegmentState8
    With DirectMusicPerformance
        Set SegmentState = .GetSegmentState(.GetMusicTime)
        GetElapsedTime = SegmentState.GetSeek
    End With
End Function

Public Sub SetStartTime(SegmentIndex, MusicTime As Long)
    If InitTrue = False Then Err.Raise 4
    Segments(SegmentIndex).SetStartPoint MusicTime
End Sub

Public Function ConvertToClockTime(MusicTime As Long)
    If InitTrue = False Then Err.Raise 4
    ConvertToClockTime = DirectMusicPerformance.MusicToClockTime(MusicTime)
End Function

Private Sub DirectXEvent8_DXCallback(ByVal eventid As Long)
    Dim DMNotification As DMUS_NOTIFICATION_PMSG
    If Not DirectMusicPerformance.GetNotificationPMSG(DMNotification) Then
        Exit Sub
    Else
        If DMNotification.lNotificationOption = DMUS_NOTIFICATION_SEGEND Then
            RaiseEvent MusicEvent(SegmentEnded)
        End If
    End If
End Sub

Private Sub Class_Terminate()
    Cleanup
End Sub

Public Sub Cleanup()
    If InitTrue = True Then
        Dim i As Long
        If Playing = True Then
            StopAudioSegment SegPlaying
        End If
        Erase Segments
        ReDim Segments(0)
        DirectMusicPerformance.RemoveNotificationType DMUS_NOTIFY_ON_SEGMENT
        DirectMusicPerformance.CloseDown
    End If
    Set DirectMusicPerformance = Nothing
    Set DirectMusicLoader = Nothing
End Sub
