VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   ClientHeight    =   4725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5490
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   5490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   2145
      Top             =   2115
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type stars      ' declare a user type stars
 X As Double
 Y As Double
 startx As Double
 starty As Double
 color As Long
 movex As Double
 movey As Double
End Type
Dim bCancel As Boolean
Private Declare Function GetTickCount Lib "kernel32" () As Long
Dim nColor
Dim bstart As Boolean
Dim star() As stars  ' Recording of the stars position, used in calculation of nstars() position

Private Sub Form_KeyPress(KeyAscii As Integer)
    Dim strPW As String
    If ScreenSaverMode <> ssPreview Then
        If strPassword <> "" Then
            ShowCursor True
            strPW = InputBox("Enter Password:", "Unlock Workstation", "")
            ShowCursor False
            If strPW <> strPassword Then Exit Sub
        End If
    End If
    bCancel = True
End Sub

Private Sub Form_Load()
    Dim nCount As Long
    Dim lCount As Long
    
    If maxstar = 0 Then maxstar = 500
    ReDim star(maxstar) As stars  ' Recording of the stars position, used in calculation of nstars() position
    
    ScaleMode = 3
    BackColor = QBColor(0)  ' Make sure that the background is black in color, because the star is white in color
    Me.Move 0, 0, Screen.Width, Screen.Height
    bstart = True
    lCount = GetTickCount()
    
    Randomize (lCount)
    
    nCount = Rnd * 10 + 1
    While nCount
        Randomize GetTickCount()
        nCount = nCount - 1
    Wend
    Me.Show
    
    Timer1.Interval = 10
    Timer1.Enabled = True
    
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strPW As String
    If ScreenSaverMode <> ssPreview Then
        If strPassword <> "" Then
            ShowCursor True
            strPW = InputBox("Enter Password:", "Unlock Workstation", "")
            ShowCursor False
            If strPW <> strPassword Then Exit Sub
        End If
    End If
    bCancel = True
End Sub
Private Sub DrawStar(X As Double, Y As Double, color As Long, colorCenter As Long)
    PSet (X, Y), colorCenter
    PSet (X + 1, Y), color
    PSet (X, Y + 1), color
    PSet (X, Y - 1), color
    PSet (X - 1, Y), color
End Sub
Private Sub DrawStars()
    Dim X As Double, Y As Double, i As Long
    Dim midx As Double, midy As Double, NewX As Double, NewY As Double
    Dim nNegx As Double, nNegy As Double
    Dim dScaleWIdth As Double, dScaleHeight As Double
    Dim dMidx As Long, dMidy As Long
    
    midx = ScaleWidth / 2
    midy = ScaleHeight / 2
    
    If (bstart) Then
        nColor = QBColor(15)
        For i = 1 To maxstar     'Initalises the variables
            NewX = Rnd * ScaleWidth
            NewY = Rnd * ScaleHeight
            NewX = (NewX - midx)
            NewY = (NewY - midy)
            nNegy = IIf(NewY < 0, -1, 1)
            nNegx = IIf(NewX < 0, -1, 1)
            ' we have change in y and x (a slope)
            ' Now reduce it so it moves at least 1 tick in the smallest direction
            If (NewY = 0) Then
                NewY = IIf(Rnd * 2 > 1, -1, 1)
            End If
            If (NewX = 0) Then
                NewX = IIf(Rnd * 2 > 1, -1, 1)
            End If
            NewX = Abs(NewX)
            NewY = Abs(NewY)
            If (NewX > NewY) Then
                ' Move in Y at least one tick for every X tick
                star(i).movex = nNegx * NewX / NewY
                star(i).movey = nNegy * NewY / NewX
            Else
                ' Move in X at least one tick for every Y tick
                star(i).movey = nNegy * NewY / NewX
                star(i).movex = nNegx * NewX / NewY
            End If
            star(i).color = RGB(128 + Rnd * 128, 128 + Rnd * 128, 128 + Rnd * 128)
            star(i).X = (nNegx * NewX) + midx
            star(i).Y = (nNegy * NewY) + midy
            star(i).startx = star(i).X
            star(i).starty = star(i).Y
            DrawStar star(i).X, star(i).Y, star(i).color, RGB(255, 255, 255)
        Next
        bstart = False
        Exit Sub
        
    End If
    
'    For i = 1 To maxstar
'        ' Erase All Old Star
'        PSet (star(i).x, star(i).y), 0
'    Next
    
    For i = 1 To maxstar
        DrawStar star(i).X, star(i).Y, 0, 0
        ' The newly calculated star position is recorded in nstar() for drawing
        star(i).X = star(i).X + star(i).movex
        star(i).Y = star(i).Y + star(i).movey
        ' If star has exceeded screen...
        If star(i).Y > ScaleHeight + 5 Or star(i).X > ScaleWidth + 5 Or star(i).X < -5 Or star(i).Y < -5 Then
            ' Put star back in center of screen
            star(i).X = midx
            star(i).Y = midy
            star(i).color = RGB(Rnd * 256, Rnd * 256, Rnd * 256)
        End If
        
        ' Draw star in new position
        DrawStar star(i).X, star(i).Y, star(i).color, RGB(255, 255, 255)
    Next
    
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    If bCancel Then
        ShowCursor True
        Unload Me
        Exit Sub
    End If
    DrawStars
    Timer1.Enabled = True
End Sub
