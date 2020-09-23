VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Repulsion3DCore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"DirectDrawGraphics"
'Repulsion Engine by Daniel Bloemendal
'If you take any code from this engine you must give me credit!

'This engine is most likely going to be discontinued.
'This is because I am working on a C++ engine.
'If you would like to continue the production of this engine,
'You must get permission from me at teh_leet_programming_man@hotmail.com
'Thank you for downloading my code please give me feedback.
'Vote for me!!!

'*** Start Public Calculation Declarations ***'
    Public SphereRadius As Double
    Public NumberOfParticles As Double
    Public Speed As Double
    Public Friction As Double
    Public OutX As Double
    Public OutY As Double
    Public OutZ As Double
    Public OutQ As Double
    Public OutM As Double
    Public CalculateInfo As Boolean
    Public KE As Double
    Public U As Double
    Public Time As Double
    Public StopCalculations
'*** End Public Calculation Declarations ***'

'*** Start Mid Calculation Declarations ***'
    Const pi = 3.141593
    Dim ax As Double
    Dim ay As Double
    Dim az As Double
    Dim r As Double
    Dim r0 As Double
    Dim v() As PointXYZ
    Dim d As Double
    Dim i As Integer
    Dim j As Integer
    Dim Particles() As ParticleXYZ
    Dim ParticlesOut() As PointXYZ
'*** End Mid Calculation Declarations ***'

Public Sub SetNewParticleValues()
On Error GoTo ErrorH
    ReDim Particles(1 To NumberOfParticles)
    ReDim ParticlesOut(1 To NumberOfParticles)
    ReDim v(1 To NumberOfParticles)

ErrorH:
    If Err Then
        MsgBox Err.Description, vbCritical
        Err.Clear
        Exit Sub
    End If
End Sub

Public Sub Calculate()
On Error GoTo ErrorH
    U = 0
    
    For i = 1 To NumberOfParticles
        ax = 0
        ay = 0
        az = 0
    
        For j = 1 To NumberOfParticles
            If j <> i Then
                r = Sqr((Particles(i).x - Particles(j).x) ^ 2 + (Particles(i).y - Particles(j).y) ^ 2 + (Particles(i).z - Particles(j).z) ^ 2)
                ax = ax + Particles(i).q * Particles(j).q * (Particles(i).x - Particles(j).x) / r ^ 3
                ay = ay + Particles(i).q * Particles(j).q * (Particles(i).y - Particles(j).y) / r ^ 3
                az = az + Particles(i).q * Particles(j).q * (Particles(i).z - Particles(j).z) / r ^ 3
                U = U + Particles(i).q * Particles(j).q / r
            End If
        Next j
        
        ax = (ax - Friction * v(i).x) / Particles(i).m
        ay = (ay - Friction * v(i).y) / Particles(i).m
        az = (az - Friction * v(i).z) / Particles(i).m
        v(i).x = v(i).x + ax * Speed
        v(i).y = v(i).y + ay * Speed
        v(i).z = v(i).z + az * Speed
        r0 = (Particles(i).x ^ 2 + Particles(i).y ^ 2 + Particles(i).z ^ 2)
            
        If r0 >= SphereRadius Then
            r0 = Sqr(r0)
            d = (Particles(i).x * v(i).x + Particles(i).y * v(i).y + Particles(i).z * v(i).z) / r0
            v(i).x = v(i).x - 2 * Particles(i).x * d / r0
            v(i).y = v(i).y - 2 * Particles(i).y * d / r0
            v(i).z = v(i).z - 2 * Particles(i).z * d / r0
        End If
        
        ParticlesOut(i).x = Particles(i).x + v(i).x * Speed
        ParticlesOut(i).y = Particles(i).y + v(i).y * Speed
        ParticlesOut(i).z = Particles(i).z + v(i).z * Speed
    Next i
    U = U / 2
    Time = Time + Speed
    For i = 1 To NumberOfParticles
        Particles(i).x = ParticlesOut(i).x
        Particles(i).y = ParticlesOut(i).y
        Particles(i).z = ParticlesOut(i).z
        If CalculateInfo = True Then
            KE = KE + (v(i).x ^ 2 + v(i).y ^ 2 + v(i).z ^ 2) * Particles(i).m
            KE = KE / 2
        End If
    Next i

ErrorH:
    If Err Then
        MsgBox Err.Description, vbCritical
        Err.Clear
        Exit Sub
    End If
End Sub

Public Sub OutParticles(Particle As Integer)
    OutX = ParticlesOut(Particle).x
    OutY = ParticlesOut(Particle).y
    OutZ = ParticlesOut(Particle).z
    OutQ = Particles(Particle).q
    OutM = Particles(Particle).m
End Sub

Public Sub SetParticle(Particle As Integer, x As Double, y As Double, z As Double, q As Double, m As Double)
    If Particle > NumberOfParticles Then Exit Sub
    Particles(Particle).x = x
    Particles(Particle).y = y
    Particles(Particle).z = z
    Particles(Particle).q = q
    Particles(Particle).m = m
End Sub

Private Sub Class_Initialize()
    ResetThreadPriority GetCurrentThread, THREAD_PRIORITY_TIME_CRITICAL
End Sub
