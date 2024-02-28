VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form5 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Form5"
   ClientHeight    =   6405
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10530
   LinkTopic       =   "Form5"
   Picture         =   "Form5.frx":0000
   ScaleHeight     =   6405
   ScaleWidth      =   10530
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080FFFF&
      Caption         =   "<< GO TO HOMEPAGE"
      Height          =   495
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5040
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "UPDATE "
      Height          =   495
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "SEARCH FLIGHT"
      Height          =   495
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2760
      Width           =   1455
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "Form5.frx":56BC1
      Left            =   1440
      List            =   "Form5.frx":56BCE
      TabIndex        =   2
      Top             =   4320
      Width           =   2535
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   1320
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393216
      Format          =   120520705
      CurrentDate     =   43684
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "FILGHT DETAILS "
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   4335
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "FLIGHT NO"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "DATE"
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
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "FILGHT UPDATER "
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   240
      TabIndex        =   9
      Top             =   3720
      Width           =   4335
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "SELECT STATUS "
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
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "UPDATE FILGHT STATUS "
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   6
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cony As New ADODB.Connection
Dim RS As New ADODB.Recordset
Dim query As String
Private Sub Command1_Click()

query = "select distinct fno from flights where tdate='" & DTPicker1.Value & "'"
RS.Open query, cony, adOpenDynamic, adLockOptimistic
Combo1.Clear
RS.MoveFirst
While Not RS.EOF
Combo1.AddItem (RS(0))
RS.MoveNext
Wend
Set RS = Nothing
End Sub

Private Sub Command2_Click()
query = "update flights set [stats] = '" & Combo2.Text & "' where [fno] = '" & Combo1.Text & "'"
RS.Open query, cony, adOpenDynamic, adLockOptimistic
MsgBox "flight status updated"
End Sub

Private Sub Command4_Click()
Me.Hide
MDIForm1.Show
End Sub

Private Sub Form_Load()
cony.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\aoms.mdb;Persist Security Info=False"
cony.Open

End Sub

