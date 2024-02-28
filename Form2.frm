VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "form2"
   ClientHeight    =   10455
   ClientLeft      =   165
   ClientTop       =   840
   ClientWidth     =   20400
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   10455
   ScaleWidth      =   20400
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "SEARCH"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1800
      Width           =   2775
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "<<GO TO HOMEPAGE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   8160
      Width           =   2415
   End
   Begin MSDataGridLib.DataGrid dg1 
      Height          =   2895
      Left            =   240
      TabIndex        =   3
      Top             =   4320
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   5106
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "AVAILABLE FLIGHTS "
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   120
      TabIndex        =   15
      Top             =   3960
      Width           =   8175
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Economy"
      Height          =   615
      Left            =   3240
      TabIndex        =   2
      Top             =   2760
      Width           =   1815
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   2040
      TabIndex        =   12
      Top             =   2400
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   556
      _Version        =   393216
      CalendarBackColor=   12648447
      Format          =   120520705
      CurrentDate     =   43639
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Business"
      Height          =   615
      Left            =   2040
      TabIndex        =   13
      Top             =   2760
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Top             =   1200
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "MODIFY SEARCH"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8160
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "BOOK TICKET "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8160
      Width           =   2175
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   2040
      TabIndex        =   1
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "TRAVEL DETAILS "
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   8055
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "CLASS"
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
         TabIndex        =   14
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "DATE "
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
         TabIndex        =   11
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "DESTINATION"
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
         TabIndex        =   10
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label1 
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
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   240
      TabIndex        =   16
      Top             =   7920
      Width           =   8175
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
      Left            =   1080
      TabIndex        =   9
      Top             =   1800
      Width           =   975
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public abcxyz, query As String
Public s As String
Public rsa As New ADODB.Recordset
Public cona As New ADODB.Connection

Private Sub Command1_Click()
query = "select * from flights where sour = '" & Combo1.Text & "' and dest = '" & Combo2.Text & "' and cat = '" & abcxyz & "' and tdate = '" & DTPicker1.Value & "'"
rsa.CursorLocation = adUseClient
rsa.Open query, cona, adOpenForwardOnly, adLockPessimistic
Set dg1.DataSource = rsa
'Set rsa = Nothing

End Sub

Private Sub Command2_Click()

Form3.Label2.Caption = rsa.Fields(1)
Form3.Label3.Caption = rsa.Fields(2)
Form3.Label4.Caption = rsa.Fields(3)
Form3.Label5.Caption = rsa.Fields(4)
Form3.Label6.Caption = rsa.Fields(5)
Form3.Label7.Caption = rsa.Fields(6)
Form3.Label1.Caption = rsa.Fields(0)
Unload Me
Form3.Show
End Sub

Private Sub Command3_Click()
Set rsa = Nothing
Combo1.Text = ""
Combo2.Text = ""
Set dg1.DataSource = Nothing
End Sub

Private Sub Command4_Click()
Me.Hide
MDIForm1.Show

End Sub

Private Sub Form_Load()
cona.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\aoms.mdb;Persist Security Info=False"
cona.Open
query = "select distinct sour from flights"
rsa.Open query, cona, adOpenDynamic, adLockOptimistic
rsa.MoveFirst
While Not rsa.EOF
Combo1.AddItem (rsa(0))
rsa.MoveNext
Wend
Set rsa = Nothing


query = "select distinct dest from flights"
rsa.Open query, cona, adOpenDynamic, adLockOptimistic
rsa.MoveFirst
While Not rsa.EOF
Combo2.AddItem (rsa(0))
rsa.MoveNext
Wend
Set rsa = Nothing


End Sub

Private Sub Form_Unload(CANCEL As Integer)
cona.Close
End Sub

Private Sub Option1_Click()
abcxyz = "business"
End Sub

Private Sub Option2_Click()
abcxyz = "economy"
End Sub
