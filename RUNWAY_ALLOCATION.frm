VERSION 5.00
Begin VB.Form RUNWAY_ALLOCATION 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Form1"
   ClientHeight    =   7200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   Picture         =   "RUNWAY_ALLOCATION.frx":0000
   ScaleHeight     =   7200
   ScaleWidth      =   11160
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "<<GO TO HOME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2760
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "ALLOCATE RUNWAY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   840
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1680
      Width           =   3255
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "RUNWAY_ALLOCATION.frx":3F1BD
      Left            =   2520
      List            =   "RUNWAY_ALLOCATION.frx":3F1D0
      TabIndex        =   0
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "RUNWAY ALLOCATION"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   4695
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2160
         TabIndex        =   4
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "RUNWAY"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "FNO"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "SOURCE "
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   1440
      Width           =   975
   End
End
Attribute VB_Name = "RUNWAY_ALLOCATION"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conf As New ADODB.Connection
Dim rsz As New ADODB.Recordset


Private Sub Command1_Click()
Dim rsb As New ADODB.Recordset
Dim st As String


st = "Insert into runway Values(" & Combo1.Text & ",'" & Combo2.Text & "')"
rsb.Open st, conf, adOpenDynamic, adLockOptimistic
Set rsb = Nothing

MsgBox "FLIGHT RUNWAY ALLOCATION REQUEST ACCEPTED"
End Sub

Private Sub Command2_Click()
Me.Hide

MDIForm1.Show

End Sub

Private Sub Form_Load()
conf.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\aoms.mdb;Persist Security Info=False"
Dim query As String
conf.Open
query = "select distinct fno from flights"
rsz.Open query, conf, adOpenDynamic, adLockOptimistic
rsz.MoveFirst
While Not rsz.EOF
Combo1.AddItem (rsz(0))
rsz.MoveNext
Wend
Set rsz = Nothing

End Sub


