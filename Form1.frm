VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000008&
   Caption         =   "Form1"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   MousePointer    =   99  'Custom
   ScaleHeight     =   5595
   ScaleWidth      =   6045
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1680
      Top             =   3360
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const LifeSpan = 8      ' Life span of each particle
Const PartsNum = 256     ' Number of particles

Dim Parts(1 To PartsNum) As Part  ' Array of particles
Dim Mx As Integer   ' Mouse X-Axis position
Dim My As Integer   ' Mouse Y-Axis position

Sub Newpart(ByVal Num As Integer, ByVal X As Integer, ByVal Y As Integer)
' Creates a new particle numer Num at starting point (X,Y)
    Parts(Num).X = X
    ' Set starting X-Axis position
    Parts(Num).Y = Y
    ' Set starting Y-Axis position
    Parts(Num).b = 256
    ' Set staring remaining lifetime
    Parts(Num).drx = ((Rnd * 2) - 1) * 20
    ' Set particle's movement on the X-Axis
    Parts(Num).dry = ((Rnd * 4) - 1) * 20
    ' Set particle's movement on the Y-Axis
End Sub

Private Sub Form_Load()
    Randomize Timer
    For i = 1 To UBound(Parts)
        Parts(i).b = i * (256 / PartsNum)
        ' Sets when particle will first appear
    Next
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
    Case 1:
    ' If left-button is pressed, create new particles,
    ' all moving at the save time
        For i = 1 To UBound(Parts)
            Me.PSet (Parts(i).X, Parts(i).Y), &H0
            Newpart i, X, Y
        Next
    Case 2:
    ' If right-button is pressed, create new random particles
        For i = 1 To UBound(Parts)
            Me.PSet (Parts(i).X, Parts(i).Y), &H0
            Newpart i, X, Y
            Parts(i).b = i * (256 / PartsNum)
        Next
    End Select
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Mx = X
    My = Y
    'Get Mouse position
    doParts Mx, My   ' Do particle calculations and display
End Sub

Sub doParts(ByVal X As Integer, ByVal Y As Integer)
Dim t As Integer
    For i = 1 To UBound(Parts)
        Me.PSet (Parts(i).X, Parts(i).Y), &H0
        ' Display background
        ' If Parts(i).b = 0 Then Newpart i, X, Y
        ' Create new particle if existing one is dead
        Parts(i).X = Parts(i).X + Parts(i).drx
        Parts(i).Y = Parts(i).Y + Parts(i).dry
        ' Change particle's location
        t = Parts(i).b - LifeSpan
        If t <= 0 Then
            Newpart i, X, Y
        Else
            Parts(i).b = t
        End If
        ' Decrease particle's remaining lifetime
        Me.PSet (Parts(i).X, Parts(i).Y), Parts(i).b
        ' Display particle at new location
    Next
End Sub

Private Sub Timer1_Timer()
    doParts Mx, My
    ' Do particles calculations & display even if mouse's
    ' not moving.
End Sub
