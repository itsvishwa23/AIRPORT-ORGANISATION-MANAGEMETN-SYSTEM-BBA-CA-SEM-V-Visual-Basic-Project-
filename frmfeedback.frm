VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FeedBack 
   BackColor       =   &H80000016&
   Caption         =   " "
   ClientHeight    =   10035
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15975
   BeginProperty Font 
      Name            =   "@Malgun Gothic"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form5"
   Picture         =   "frmfeedback.frx":0000
   ScaleHeight     =   10035
   ScaleWidth      =   15975
   WindowState     =   2  'Maximized
   Begin VB.OptionButton opt4 
      BackColor       =   &H80000001&
      ForeColor       =   &H00FFFFFF&
      Height          =   720
      Left            =   1080
      Picture         =   "frmfeedback.frx":370A3
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7320
      Width           =   2295
   End
   Begin VB.OptionButton opt5 
      BackColor       =   &H80000001&
      ForeColor       =   &H00FFFFFF&
      Height          =   720
      Left            =   3720
      Picture         =   "frmfeedback.frx":39D7F
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7320
      Width           =   2295
   End
   Begin VB.OptionButton opt3 
      BackColor       =   &H80000001&
      ForeColor       =   &H00FFFFFF&
      Height          =   720
      Left            =   5400
      Picture         =   "frmfeedback.frx":3CAF2
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6240
      Width           =   2295
   End
   Begin VB.OptionButton opt2 
      BackColor       =   &H80000001&
      ForeColor       =   &H00FFFFFF&
      Height          =   720
      Left            =   2760
      Picture         =   "frmfeedback.frx":3F6A3
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6240
      Width           =   2295
   End
   Begin VB.OptionButton opt1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   720
      Left            =   240
      Picture         =   "frmfeedback.frx":4224E
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6240
      Width           =   2295
   End
   Begin VB.TextBox txtfeed_id 
      Height          =   525
      Left            =   2640
      TabIndex        =   12
      Top             =   1200
      Width           =   2655
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   10080
      Top             =   9000
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\ayurvedic_shoppe\AyurvedicDatabase.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\ayurvedic_shoppe\AyurvedicDatabase.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "FeedBack_mst"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   2760
      TabIndex        =   11
      Top             =   2760
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "@Malgun Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   124583937
      CurrentDate     =   43328
   End
   Begin VB.CommandButton cmdaddnew 
      Caption         =   "Add New"
      Height          =   735
      Left            =   2040
      TabIndex        =   9
      ToolTipText     =   "Add the Feedback id"
      Top             =   8520
      Width           =   1935
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear"
      Height          =   735
      Left            =   6600
      TabIndex        =   8
      ToolTipText     =   "Clear all the text"
      Top             =   8520
      Width           =   1695
   End
   Begin VB.CommandButton cmdhome 
      Caption         =   "<< GO TO HOMEPAGE"
      Height          =   735
      Left            =   4320
      TabIndex        =   7
      ToolTipText     =   "Go to home page"
      Top             =   8520
      Width           =   1935
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
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
      Left            =   0
      TabIndex        =   6
      ToolTipText     =   "Save the feedback details"
      Top             =   8520
      Width           =   1815
   End
   Begin VB.TextBox txtsug 
      Height          =   1095
      Left            =   2640
      TabIndex        =   5
      Top             =   3480
      Width           =   5415
   End
   Begin VB.TextBox txtname 
      Height          =   510
      Left            =   2640
      TabIndex        =   3
      Top             =   1920
      Width           =   5415
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "RATING"
      BeginProperty Font 
         Name            =   "@Malgun Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   13
      Top             =   5280
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Feedback Date"
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Leave Suggestion"
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   3840
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Feedback id"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label Feedbaktitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "FEEDBACK FORM"
      BeginProperty Font 
         Name            =   "@Malgun Gothic"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   -720
      TabIndex        =   0
      Top             =   240
      Width           =   5175
   End
End
Attribute VB_Name = "FeedBack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rec As New ADODB.Recordset
Dim str As String
Dim feed_id As Integer
Private Sub cmdaddnew_Click()
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Lenovo\Desktop\vishwanath\aoms.mdb;Persist Security Info=False"
str = " Select max(FeedBack_ID) from FeedBack_mst"
rec.Open str, con, adOpenDynamic, adLockOptimistic
feed_id = rec.Fields(0)
feed_id = feed_id + 1
txtfeed_id.Text = feed_id
txtname = ""
txtsug = ""
con.Close
End Sub

Private Sub cmdclear_Click()
txtfeed_id.Text = ""
txtname = ""
txtsug = ""
opt1.Value = False
opt2.Value = False
opt3.Value = False
opt4.Value = False
opt5.Value = False

End Sub

Private Sub cmdhome_Click()
MDIForm1.Show
Me.Hide
End Sub

Private Sub cmdsave_Click()
Dim rating
If opt1.Value = True Then
rating = "1Star"
ElseIf opt2.Value = True Then
reting = "2Star"
ElseIf opt3.Value = True Then
rating = "3Star"
ElseIf opt4.Value = True Then
rating = "4Star"
Else
rating = "5Star"
End If


con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Lenovo\Desktop\vishwanath\aoms.mdb;Persist Security Info=False"
MsgBox "Connection Done"
str = "insert into FeedBack_mst values(" & txtfeed_id.Text & ",'" & txtname.Text & "','" & DTPicker1.Value & "','" & txtsug.Text & "','" & rating & "')"
rec.Open str, con, adOpenDynamic, adLockOptimistic
con.Close
MsgBox "FeedBack Saved"
End Sub



Private Sub Picture1_Click()

End Sub
