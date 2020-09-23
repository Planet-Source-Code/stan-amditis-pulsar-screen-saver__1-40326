Attribute VB_Name = "Module1"
Option Explicit

'Globals

Public gravityStrength As Double       'default 20
Public numOfBarsToDraw As Byte         'default 1
Public maxstar As Long
Public strPassword As String
Public ScreenSaverMode As ScreenSaverConstants

' WINAPI Declarations
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOZORDER = &H4
Public Const SWP_SHOWWINDOW = &H40
Public Const HWND_TOP = 0
Public Const WS_CHILD = &H40000000
Public Const GWL_HWNDPARENT = (-8)
Public Const GWL_STYLE = (-16)

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

' Application Enums
Public Enum ScreenSaverConstants
    ssSettings = 1
    ssNormal = 2
    ssPreview = 3
End Enum

' Private variables.
Public Const THIS_APPLICATION = "PulsarScreenSaver"


Private Function GetCommandLineWindowHandle(ByVal strCmdLine As String) As Long
    Dim nCmdLineSize As Integer, nCmdLinePos As Integer
    Dim strCmdLineChar As String

    ' Start from the end and backtrack until we find the first digit
    strCmdLine = Trim$(strCmdLine)
    nCmdLineSize = Len(strCmdLine)
    For nCmdLinePos = nCmdLineSize To 1 Step -1
        strCmdLineChar = Mid$(strCmdLine, nCmdLinePos, 1)
        If strCmdLineChar < "0" Or strCmdLineChar > "9" Then Exit For
    Next
    
    ' Found beginning of hwnd value, now return it as a long
    GetCommandLineWindowHandle = CLng(Mid$(strCmdLine, nCmdLinePos + 1))
    
End Function

Private Sub EnsureSingleInstance()
    ' If no other instance detected...exit sub
    If Not App.PrevInstance Then Exit Sub
    ' If the window is running as screen saver already, exit app
    If FindWindow(vbNullString, THIS_APPLICATION) Then End
    ' Change THIS instance's form caption to the screensaver title
    Form1.Caption = THIS_APPLICATION
End Sub

' Start the program.
Public Sub Main()
    Dim strCmdLine As String
    Dim hWndPreview As Long
    Dim rectPreview As RECT

    maxstar = CLng(GetSetting(THIS_APPLICATION, "Settings", "Stars", "500"))
    strPassword = GetSetting(THIS_APPLICATION, "Settings", "Password", "")
    
    strCmdLine = UCase$(Trim$(Command$))

    Select Case Mid$(strCmdLine, 1, 2)
        Case "/C"   ' Settings
            ScreenSaverMode = ssSettings
        Case "/P"   ' Preview
            ScreenSaverMode = ssPreview
        Case Else   ' Else Screensaver
            ScreenSaverMode = ssNormal
    End Select

    Select Case ScreenSaverMode
        Case ssSettings
            Form2.Show
        
        Case ssNormal
            
            EnsureSingleInstance
            ShowCursor False
            Form1.Show
            
        Case ssPreview
            
            ' Get the preview area hWnd from the command line
            hWndPreview = GetCommandLineWindowHandle(strCmdLine)
            ' Get the size of the preview window
            GetClientRect hWndPreview, rectPreview
            
            Load Form1
            Form1.Caption = "Preview"
            ' Set the window's new style so that it can act as a child of the preview controller.
            SetWindowLong Form1.hwnd, GWL_STYLE, GetWindowLong(Form1.hwnd, GWL_STYLE) Or WS_CHILD
            SetParent Form1.hwnd, hWndPreview
            SetWindowLong Form1.hwnd, GWL_HWNDPARENT, hWndPreview
            ' Show the window using the proper size
            SetWindowPos Form1.hwnd, HWND_TOP, 0&, 0&, rectPreview.Right, rectPreview.Bottom, SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_SHOWWINDOW
            
    End Select
End Sub
