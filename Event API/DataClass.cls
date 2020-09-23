VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Event Engine by Daniel Bloemendal
'If you take any code from this engine you must give me credit!

'This engine is most likely going to be discontinued.
'This is because I am working on a C++ engine.
'If you would like to continue the production of this engine,
'You must get permission from me at teh_leet_programming_man@hotmail.com
'Thank you for downloading my code please give me feedback.
'Vote for me!!!

Dim SepS As Integer
Dim SepE As Integer
Dim SepSS As Integer

Public Function DataRead(Text As String, Seperator As String, Section As Variant, Optional EndSeperator As String) As String
    SepS = 0
    For DataN = 1 To Section
        SepS = InStr(SepS + 1, Text, Seperator, vbTextCompare)
    Next DataN
    If EndSeperator = "" Then
        SepE = Len(Text) + 1
    Else
        SepE = InStr(SepS + 1, Text, EndSeperator, vbTextCompare)
    End If
    If SepS >= SepE Then
        If SepS + 1 > Len(Text) Then Err.Raise 2
        DataRead = Mid(Text, SepS + 1)
        Exit Function
    End If
    If SepS = 0 Then SepSS = 1 Else SepSS = 0
    DataRead = Mid(Text, SepS + 1, SepE - SepSS - SepS)
End Function
