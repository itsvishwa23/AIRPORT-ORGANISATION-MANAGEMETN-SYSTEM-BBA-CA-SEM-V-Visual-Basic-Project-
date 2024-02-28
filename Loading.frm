VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BF38D12B-22A9-4B10-B26E-019F2B5F9C22}#1.0#0"; "Ani Gif.ocx"
Begin VB.Form Loading 
   Caption         =   "Form1"
   ClientHeight    =   10905
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   20010
   LinkTopic       =   "Form1"
   Picture         =   "Loading.frx":0000
   ScaleHeight     =   10905
   ScaleWidth      =   20010
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ProgressBar ProgressBar2 
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   10080
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   1085
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   7560
      Top             =   7800
   End
   Begin Project1.PictureG PictureG1 
      Height          =   9000
      Left            =   11280
      Top             =   720
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   15875
      GIF             =   "Loading.frx":7534
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Login SuccesFully"
      BeginProperty Font 
         Name            =   "@Malgun Gothic"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   480
      TabIndex        =   2
      Top             =   1320
      Width           =   6015
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "@Malgun Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   3360
      TabIndex        =   1
      Top             =   9120
      Width           =   525
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading....."
      BeginProperty Font 
         Name            =   "@Malgun Gothic"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   0
      Top             =   9120
      Width           =   2775
   End
End
Attribute VB_Name = "Loading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
ProgressBar2.Min = 0
Timer1.Enabled = True

End Sub


Private Sub Timer1_Timer()
ProgressBar2.Value = ProgressBar2.Value + 5
Label2.Caption = ProgressBar2.Value & "%"
If (ProgressBar2.Value = ProgressBar2.Max) Then
MDIForm1.Show
Timer1.Enabled = False
Unload Me
End If
End Sub
