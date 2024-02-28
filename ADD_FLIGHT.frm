VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ADD_FLIGHT 
   Caption         =   "Form1"
   ClientHeight    =   7920
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10710
   LinkTopic       =   "Form1"
   Picture         =   "ADD_FLIGHT.frx":0000
   ScaleHeight     =   7920
   ScaleWidth      =   10710
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "<<GO TO HOME"
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
      Left            =   1200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5400
      Width           =   2775
   End
   Begin VB.TextBox Text5 
      Height          =   315
      Left            =   2040
      TabIndex        =   14
      Top             =   3960
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      Height          =   315
      Left            =   2040
      TabIndex        =   13
      Top             =   3480
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   2040
      TabIndex        =   12
      Top             =   3000
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   2040
      TabIndex        =   8
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "pune"
      Top             =   960
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "ADD FLIGHT"
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
      Left            =   1920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4560
      Width           =   2775
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   2040
      TabIndex        =   1
      Top             =   2040
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   556
      _Version        =   393216
      CalendarBackColor=   12648447
      Format          =   41746433
      CurrentDate     =   43639
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "FLIGHT DETAILS "
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   4695
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "NEW"
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
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   4320
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "TIME"
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
         Top             =   3720
         Width           =   855
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "FARE"
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
         Top             =   3240
         Width           =   855
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
         TabIndex        =   9
         Top             =   2760
         Width           =   855
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
         TabIndex        =   5
         Top             =   720
         Width           =   1095
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
         TabIndex        =   4
         Top             =   1320
         Width           =   1815
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
         TabIndex        =   3
         Top             =   1800
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
      Left            =   1080
      TabIndex        =   6
      Top             =   1440
      Width           =   975
   End
End
Attribute VB_Name = "ADD_FLIGHT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim rsb As New ADODB.Recordset
Dim st As String
Dim newfare As Integer
st = "Insert into flights Values(" & Text3.Text & ",'" & Text1.Text & "','" & Text2.Text & "','economy'," & Text4.Text & ",'" & DTPicker1.Value & "','" & Text5.Text & "','ONTIME')"
rsb.Open st, conn, adOpenDynamic, adLockOptimistic
Set rsb = Nothing
newfare = Val(Text4.Text) * 2
st = "Insert into flights Values(" & Text3.Text & ",'" & Text1.Text & "','" & Text2.Text & "','business'," & newfare & ",'" & DTPicker1.Value & "','" & Text5.Text & "','ONTIME')"
rsb.Open st, conn, adOpenDynamic, adLockOptimistic
Set rsb = Nothing

MsgBox "FLIGHT ADDED"
End Sub

Private Sub Command2_Click()
Me.Hide
MDIForm1.Show

End Sub

Private Sub Command3_Click()
Dim intsupid As Variant
intsupid = NextIdVish("flights", 0)
Text3.Text = intsupid
End Sub

Private Sub Form_Load()
Call Connect
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call DisConnect
End Sub
Public Function NextIdVish(strTableName As String, intFieldIndex As Integer) As Integer
Dim intNextId  As Integer
Dim r As ADODB.Recordset
Set r = New ADODB.Recordset
r.Open "Select max(fno) from " + strTableName, conn, adOpenForwardOnly, adLockOptimistic
If Not r.EOF Then
    intNextId = r.Fields(0)
Else
    intNextId = 0
End If
intNextId = intNextId + 1
Set r = Nothing
NextIdVish = intNextId
End Function

