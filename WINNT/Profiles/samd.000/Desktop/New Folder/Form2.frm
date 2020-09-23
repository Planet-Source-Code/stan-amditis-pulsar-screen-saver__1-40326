VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Pulsar Screensaver Configuration"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4215
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1305
      PasswordChar    =   "*"
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1560
      Width           =   1245
   End
   Begin VB.TextBox Text1 
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1305
      PasswordChar    =   "*"
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1140
      Width           =   1245
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   360
      Left            =   2700
      TabIndex        =   6
      Top             =   1560
      Width           =   1110
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   360
      Left            =   2700
      TabIndex        =   5
      Top             =   1200
      Width           =   1110
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   390
      Left            =   1290
      TabIndex        =   0
      Top             =   600
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   688
      _Version        =   393216
      Min             =   10
      Max             =   5000
      SelStart        =   500
      TickFrequency   =   1000
      Value           =   500
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Confirm:"
      Height          =   270
      Left            =   375
      TabIndex        =   3
      Top             =   1635
      Width           =   900
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Password:"
      Height          =   270
      Left            =   390
      TabIndex        =   1
      Top             =   1215
      Width           =   900
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Star Count:"
      Height          =   285
      Left            =   30
      TabIndex        =   8
      Top             =   195
      Width           =   1215
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "500"
      Height          =   330
      Left            =   1305
      TabIndex        =   7
      Top             =   165
      Width           =   1215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bInit As Boolean
Private Sub Command1_Click()
    maxstar = Slider1.Value
    If (Text1.Text <> Text2.Text) Then
        MsgBox "Password Mismatch", vbOKOnly + vbInformation, "Set Screensaver Password"
        Exit Sub
    End If
    SaveSetting THIS_APPLICATION, "Settings", "Stars", CStr(maxstar)
    SaveSetting THIS_APPLICATION, "Settings", "Password", Text1.Text
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If bInit Then
        bInit = False
        Text1.SetFocus
    End If
End Sub

Private Sub Form_Load()
    bInit = True
    If maxstar = 0 Then maxstar = 500
    Slider1.Value = maxstar
    Label1.Caption = maxstar
    Text1.Text = ""
    Text2.Text = ""
End Sub

Private Sub Slider1_Click()
    Label1.Caption = Slider1.Value
End Sub

Private Sub Text1_GotFocus()
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
End Sub


Private Sub Text2_GotFocus()
    Text2.SelStart = 0
    Text2.SelLength = Len(Text2.Text)
End Sub
