VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Form6"
   ClientHeight    =   7350
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10785
   LinkTopic       =   "Form6"
   Picture         =   "Form6.frx":0000
   ScaleHeight     =   7350
   ScaleWidth      =   10785
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "<< GO TO HOMEPAGE"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4440
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CHECK STATUS"
      Height          =   495
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4440
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "VIEW FLIGHT STATUS"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   4815
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ENTER ELECTRONIC TICKET NUMBER"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   4155
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "FLIGHT STATUS IS"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   5895
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   1335
      End
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conz As New ADODB.Connection
Dim RS As New ADODB.Recordset
Dim query As String
Private Sub Command1_Click()

query = "select stats from flights,booking_master where flights.fno=booking_master.fno and etn= " & Text1.Text & ""
RS.Open query, conz, adOpenDynamic, adLockOptimistic
Label2.Caption = RS(0)
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
conz.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\aoms.mdb;Persist Security Info=False"
conz.Open

End Sub

