VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form7 
   Caption         =   "Form7"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18915
   LinkTopic       =   "Form7"
   Picture         =   "Form7.frx":0000
   ScaleHeight     =   10935
   ScaleWidth      =   18915
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "<< GO TO HOMEPAGE"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3600
      TabIndex        =   10
      Top             =   1680
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "UPDATE SHIFT "
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3600
      TabIndex        =   4
      Top             =   720
      Width           =   2055
   End
   Begin MSComCtl2.DTPicker dt2 
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   3000
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393216
      Format          =   126484481
      CurrentDate     =   43684
   End
   Begin MSComCtl2.DTPicker dt1 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   3000
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393216
      Format          =   126484481
      CurrentDate     =   43684
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "Form7.frx":453C3
      Left            =   240
      List            =   "Form7.frx":453CD
      TabIndex        =   1
      Top             =   2040
      Width           =   2535
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form7.frx":453E1
      Left            =   240
      List            =   "Form7.frx":453E3
      TabIndex        =   0
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "EMPLOYEE SHIFT UPDATE PORTAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   6015
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "END DATE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   9
         Top             =   2520
         Width           =   2295
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "START DATE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2520
         Width           =   2295
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "SELECT SHIFT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "SELECT EMPLOYEE NO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   2775
      End
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conf As New ADODB.Connection
Dim RS As New ADODB.Recordset
Dim query As String
Private Sub Command2_Click()
Me.Hide
MDIForm1.Show

End Sub

Private Sub Command1_Click()
query = "insert into emp_shift values(" & Combo1.Text & ",'" & Combo2.Text & "','" & dt1.Value & "','" & dt2.Value & "')"
RS.Open query, conf, adOpenDynamic, adLockOptimistic
MsgBox "record inserted"
Set RS = Nothing
End Sub

Private Sub Form_Load()
conf.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\aoms.mdb;Persist Security Info=False"
conf.Open

query = "select empid from employee"
RS.Open query, conf, adOpenDynamic, adLockOptimistic
Combo1.Clear
RS.MoveFirst
While Not RS.EOF
Combo1.AddItem (RS(0))
RS.MoveNext
Wend
Set RS = Nothing


End Sub

Private Sub Frame1_DragDrop(SOURCE As Control, X As Single, Y As Single)
Me.Hide
MDIForm1.Show
End Sub
