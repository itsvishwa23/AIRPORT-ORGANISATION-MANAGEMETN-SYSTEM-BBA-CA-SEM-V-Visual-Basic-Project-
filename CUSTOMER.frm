VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCUSTOMER 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Form2"
   ClientHeight    =   10935
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   20895
   LinkTopic       =   "Form2"
   Picture         =   "CUSTOMER.frx":0000
   ScaleHeight     =   10935
   ScaleWidth      =   20895
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2295
      Left            =   240
      TabIndex        =   24
      Top             =   6720
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   4048
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
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0FFFF&
      Caption         =   "NEW"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0FFFF&
      Caption         =   "<< GO TO HOMEPAGE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "UPDATE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "DELETE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Height          =   1455
      Left            =   240
      TabIndex        =   22
      Top             =   5040
      Width           =   8415
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   840
      Width           =   2535
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   3000
      TabIndex        =   13
      Top             =   1440
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      Format          =   124583937
      CurrentDate     =   43641
   End
   Begin VB.TextBox Text6 
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   5160
      TabIndex        =   12
      Top             =   3720
      Width           =   1335
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Female"
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
      Left            =   3960
      TabIndex        =   11
      Top             =   3720
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Male"
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
      Left            =   3000
      TabIndex        =   10
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   4200
      Width           =   2535
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   3120
      Width           =   2535
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   2640
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   2040
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Customer Details "
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   120
      TabIndex        =   14
      Top             =   360
      Width           =   8535
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   2565
         Left            =   5760
         Picture         =   "CUSTOMER.frx":42569
         ScaleHeight     =   2565
         ScaleWidth      =   2295
         TabIndex        =   23
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer No "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   3960
         Width           =   1455
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gender "
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
         TabIndex        =   20
         Top             =   3480
         Width           =   825
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Email id"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Passport Number "
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
         TabIndex        =   18
         Top             =   2520
         Width           =   1860
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone No"
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
         Left            =   240
         TabIndex        =   17
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Birth Date "
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
         TabIndex        =   16
         Top             =   1200
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
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
         TabIndex        =   15
         Top             =   600
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmCUSTOMER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim rsnew As New ADODB.Recordset
Dim st As String
Private Sub Command1_Click()

st = "Insert into customer Values('" & Text1.Text & "','" & DTPicker1.Value & "','" & Text3.Text & "','" & Text4.Text & "','" & Text5.Text & "','" & Text6.Text & "'," & Text7.Text & ")"
RS.Open st, conn, adOpenDynamic, adLockOptimistic
MsgBox ("Record saved")
Set RS = Nothing
Dim oSmtp As New EASendMailObjLib.mail
    oSmtp.LicenseCode = "TryIt"
    
    ' Set your Gmail email address
    oSmtp.FromAddr = "rajashreestaware@gmail.com"   'Enter your Email ID here
    
    ' Add recipient email address
    oSmtp.AddRecipientEx "'" & Text5.Text & "'", 0   'Enter Reciver Email ID here
    
    ' Set email subject
    oSmtp.Subject = "This is a test mail with attachment"
    
    ' Set email body
    oSmtp.BodyText = "This is a Test Mail from VB 6.0 application"
       
       If oSmtp.AddAttachment("C:\Users\Lenovo\Desktop\vishwanath\EMAIL.txt") <> 0 Then 'Location of Attached File
        MsgBox "Failed to add attachment with error:" & oSmtp.GetLastErrDescription()
    End If
    ' Gmail SMTP server address
    oSmtp.ServerAddr = "smtp.gmail.com"
    
    ' set direct SSL 465 port,
    oSmtp.ServerPort = 465
    
    ' detect SSL/TLS automatically
    oSmtp.SSL_init

    ' Gmail user authentication should use your
    ' Gmail email address as the user name.
    ' For example: your email is "gmailid@gmail.com", then the user should be "gmailid@gmail.com"
    oSmtp.UserName = "rajashreestaware@gmail.com" 'Enter your Email ID here again
    oSmtp.Password = "vishwanathtaware"    'Enter Your Mail Password
    
    MsgBox "start to send email ..."

    If oSmtp.SendMail() = 0 Then
        MsgBox "email was sent successfully!"
    Else
        MsgBox "failed to send email with the following error:" & oSmtp.GetLastErrDescription()
    End If
End Sub

Private Sub Command10_Click()
Unload Me
MDIForm1.Show



End Sub

Private Sub Command11_Click()
intsupid = NextIdNew("customer", 6)
Text7.Text = Trim(str(intsupid))
End Sub

Private Sub Command2_Click()
Dim rdel As New ADODB.Recordset
st = "Delete from customer where cno='" + Text7.Text + "'"
rdel.Open st, conn, adOpenDynamic, adLockOptimistic
MsgBox ("Record deleted")
Set rdel = Nothing

End Sub

Private Sub Command3_Click()
conn.Execute "UPDATE customer SET [cname]='" & Text1.Text & "',[dob]='" & DTPicker1.Value & "',[phone]='" & Text3.Text & "',[pp]='" & Text4.Text & "',[email]='" & Text5.Text & "' where cno=" & Text7.Text & ""
MsgBox "Customer info updated successfully."

End Sub

Private Sub Command9_Click()
Text8.Text = ""
Text10.Text = ""
Text9.Text = ""

End Sub

Private Sub Text15_Change()

End Sub

Private Sub DataGrid1_Click()
Text1.Text = DataGrid1.Columns(0).Value
DTPicker1.Value = DataGrid1.Columns(1).Value
Text3.Text = DataGrid1.Columns(2).Value
Text4.Text = DataGrid1.Columns(3).Value
Text5.Text = DataGrid1.Columns(4).Value
Text7.Text = DataGrid1.Columns(6).Value





If (DataGrid1.Columns(0) = "male") Then
Option1.Value = True
Else
Option2.Value = True
End If
End Sub

Private Sub Form_Load()
Call Connect
rsnew.CursorLocation = adUseClient
rsnew.Open "select * from customer", conn, adOpenDynamic, adLockOptimistic

Set DataGrid1.DataSource = rsnew
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call DisConnect
End Sub

Private Sub Option1_Click()
Text6.Text = "MALE"
End Sub

Private Sub Option2_Click()
Text6.Text = "FEMALE"
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = character(KeyAscii)

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
KeyAscii = mobnumber(KeyAscii)
        

End Sub

Private Sub Text3_LostFocus()
myvar = Trim$(Me.Text3.Text)
     
    If myvar = "" Then
        MsgBox "PHONE NUMBER SHOULD NOT BE BLANK"
    ElseIf Len(myvar) > 10 Then
        MsgBox "PHONE NUMBER SHOULD NOT BE GREATER THAN 10"
        ElseIf Len(myvar) < 10 Then
        MsgBox "PHONE NUMBER SHOULD NOT BE LESS THAN 10"
    End If
End Sub

Private Sub Text4_LostFocus()
myvar = Trim$(Me.Text4.Text)
     
    If myvar = "" Then
        MsgBox "PASSPORT NUMBER SHOULD NOT BE BLANK"
    ElseIf Len(myvar) > 6 Then
        MsgBox "PASSPORT NUMBER SHOULD NOT BE GREATER THAN 6"
        ElseIf Len(myvar) < 6 Then
        MsgBox "PASSPORT NUMBER SHOULD NOT BE LESS THAN 6"
    End If
End Sub

Private Sub Text5_LostFocus()
valemail (Text5.Text)

End Sub
