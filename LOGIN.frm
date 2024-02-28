VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "AOMS LOGIN PAGE"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19335
   LinkTopic       =   "Form1"
   Picture         =   "LOGIN.frx":0000
   ScaleHeight     =   10935
   ScaleWidth      =   19335
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   600
      Top             =   7320
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
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
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   5520
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   6840
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   5
      Top             =   5880
      Width           =   2295
   End
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   17280
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9840
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   17280
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3840
      TabIndex        =   8
      Top             =   6960
      Width           =   1545
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "USERNAME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3840
      TabIndex        =   7
      Top             =   6000
      Width           =   1515
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s As String
Dim unm As String
Dim pwd As String
Dim CUSTOMER, ticket, board, operation, atc As String
Dim rs As New ADODB.Recordset
Dim con As New ADODB.Connection
Private Sub Command1_Click()
On Error Resume Next


unm = Trim(Text1.Text)
pwd = Trim(Text2.Text)
If Len(unm) = 0 Or Len(pwd) = 0 Then
MsgBox "username or password is invalid", vbCritical
Exit Sub
End If

s = "select * from login where  userid = '" & unm & "' and pass = '" & pwd & "' and cat = '" & CUSTOMER & "'"
rs.Open s, con, adOpenDynamic, adLockOptimistic, adCmdText
 rs.MoveFirst
 If Not rs.EOF Then
 
   Unload Me
frmCUSTOMER.Show
   
 Else
   MsgBox "invalid user name or password ", vbCritical, "Error"
   Text2.Text = ""
   Text2.SetFocus
 End If
rs.Close



End Sub

Private Sub Command2_Click()
On Error Resume Next


unm = Trim(Text1.Text)
pwd = Trim(Text2.Text)
If Len(unm) = 0 Or Len(pwd) = 0 Then
MsgBox "username or password is invalid", vbCritical
Exit Sub
End If

s = "select * from login where  userid = '" & unm & "' and pass = '" & pwd & "' and cat = '" & board & "'"
rs.Open s, con, adOpenDynamic, adLockOptimistic, adCmdText
 rs.MoveFirst
 If Not rs.EOF Then
 
   Unload Me
frmBOARDING.Show
   
 Else
   MsgBox "invalid user name or password ", vbCritical, "Error"
   Text2.Text = ""
   Text2.SetFocus
 End If
rs.Close







End Sub

Private Sub Command3_Click()
On Error Resume Next


unm = Trim(Text1.Text)
pwd = Trim(Text2.Text)
If Len(unm) = 0 Or Len(pwd) = 0 Then
MsgBox "username or password is invalid", vbCritical
Exit Sub
End If

s = "select * from login where  userid = '" & unm & "' and pass = '" & pwd & "' and cat = '" & operation & "'"
rs.Open s, con, adOpenDynamic, adLockOptimistic, adCmdText
 rs.MoveFirst
 If Not rs.EOF Then
 
   Unload Me
frmBOARDING.Show
   
 Else
   MsgBox "invalid user name or password ", vbCritical, "Error"
   Text2.Text = ""
   Text2.SetFocus
 End If
rs.Close






frmAOF.Show

End Sub

Private Sub Command4_Click()
On Error Resume Next


unm = Trim(Text1.Text)
pwd = Trim(Text2.Text)
If Len(unm) = 0 Or Len(pwd) = 0 Then
MsgBox "username or password is invalid", vbCritical
Exit Sub
End If

s = "select * from login where  userid = '" & unm & "' and pass = '" & pwd & "' and cat = '" & atc & "'"
rs.Open s, con, adOpenDynamic, adLockOptimistic, adCmdText
 rs.MoveFirst
 If Not rs.EOF Then
 
   Unload Me
frmATC.Show
   
 Else
   MsgBox "invalid user name or password ", vbCritical, "Error"
   Text2.Text = ""
   Text2.SetFocus
 End If
rs.Close






frmATC.Show

End Sub

Private Sub Command6_Click()
On Error Resume Next


unm = Trim(Text1.Text)
pwd = Trim(Text2.Text)
If Len(unm) = 0 Or Len(pwd) = 0 Then
MsgBox "username or password is invalid", vbCritical
Exit Sub
End If

s = "select * from login where  userid = '" & unm & "' and pass = '" & pwd & "' and cat = '" & ticket & "'"
rs.Open s, con, adOpenDynamic, adLockOptimistic, adCmdText
 rs.MoveFirst
 If Not rs.EOF Then
 
   Unload Me
frmTIK_COUNT.Show
   
 Else
   MsgBox "invalid user name or password ", vbCritical, "Error"
   Text2.Text = ""
   Text2.SetFocus
 End If
rs.Close





End Sub

Private Sub Form_Load()
On Error Resume Next
CUSTOMER = "customer"
ticket = "ticket"
board = "board"
operation = "operation"
atc = "ATC"
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\aoms.mdb;Persist Security Info=False"
con.Open

End Sub

Private Sub Form_Unload(Cancel As Integer)
con.Close
End Sub
