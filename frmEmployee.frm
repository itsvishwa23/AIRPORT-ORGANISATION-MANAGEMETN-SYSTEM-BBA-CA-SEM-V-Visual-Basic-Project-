VERSION 5.00
Begin VB.Form frmEmployee 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Employee"
   ClientHeight    =   9960
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16020
   LinkTopic       =   "Form1"
   Picture         =   "frmEmployee.frx":0000
   ScaleHeight     =   9960
   ScaleWidth      =   16020
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Height          =   975
      Left            =   480
      TabIndex        =   11
      Top             =   3960
      Width           =   5295
      Begin VB.CommandButton cmdFind 
         Caption         =   "Find"
         Height          =   375
         Left            =   3000
         TabIndex        =   16
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2040
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1080
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   375
         Left            =   3960
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Employee Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Width           =   5295
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1680
         TabIndex        =   20
         Top             =   2520
         Width           =   2895
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1680
         TabIndex        =   5
         Top             =   2160
         Width           =   2895
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   4
         Top             =   1800
         Width           =   2895
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   3
         Top             =   1440
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox txtCustomerID 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   1
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Designation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label lblPatientID 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Employee"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      TabIndex        =   17
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "frmEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s1 As String
Dim custid As Integer
Public conn As ADODB.Connection
Public rsGlb As ADODB.Recordset
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
Dim response As Integer
 Dim sql As String
 Dim rsd As New ADODB.Recordset
    response = MsgBox("Do you want to Delete this Record", vbYesNo + vbExclamation, "Message")
    If response = vbNo Then Exit Sub
    sql = "DELETE FROM Employee WHERE EmpID =" & txtCustomerID.Text
    rsd.Open sql, conn, adOpenDynamic, adLockOptimistic
    MsgBox "Employee ID " & txtCustomerID.Text & " deleted..!", vbOKOnly + vbInformation, "Information"
    
    reset

End Sub

Private Sub cmdFind_Click()
Dim str As String
If MsgBox("Do You want to Search by Name?", vbYesNo + vbQuestion, "Information") = vbYes Then
    str = InputBox("Enter User Name", "Update", , 500, 2500)
    Call QuerySelect("select * from Employee where EmpName='" & str & "'")
    If rsGlb.EOF = True Then
    MsgBox "No Record Found for User Name=" & str, vbOKOnly, "Record Not Found"
    reset
    Exit Sub
    End If
Else
    custid = InputBox("Enter User ID for Update", "Update", , 500, 2500)
    Call QuerySelect("select * from Employee where EmpID=" & custid)
    If rsGlb.EOF = True Then
    MsgBox "No Record Found for User ID=" & custid, vbOKOnly, "Record Not Found"
    reset
    Exit Sub
    End If
End If


On Error Resume Next



txtCustomerID.Text = rsGlb.Fields(0)
Text1.Text = rsGlb.Fields(1)
Text2.Text = rsGlb.Fields(2)
Text3.Text = rsGlb.Fields(3)
Text4.Text = rsGlb.Fields(4)
Text5.Text = rsGlb.Fields(5)
cmdDelete.Enabled = True
cmdUpdate.Enabled = True

End Sub

Private Sub cmdsave_Click()
If txtCustomerID.Text = "" Or Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4 = "" Or Text5.Text = "" Then
    MsgBox "Please Enter Data..!", vbInformation + vbOKOnly, "Information"
Else
    Dim RS As New ADODB.Recordset
    RS.Open "select * from Employee", conn, adOpenDynamic, adLockOptimistic
    With RS
                If MsgBox("Are You Sure You Wish To Save This Record?", vbYesNo + vbQuestion, "Save This Record?") = vbYes Then
                    .AddNew
                    
                     .Fields(1) = Text1.Text
                    .Fields(2) = Text2.Text
                    .Fields(3) = Text3.Text
                    .Fields(4) = Text4.Text
                    .Fields(5) = Text5.Text
                    .Fields(6) = Text6.Text
                    .Update
                    .Requery
                    MsgBox "The Record Was Saved Successfully!", vbInformation, "Succesful Save Procedure"
                End If
    End With
End If
reset
End Sub

Private Sub cmdUpdate_Click()
If txtCustomerID.Text = "" Or Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4 = "" Or Text5.Text = "" Then
                MsgBox "Please fill all fields..!", vbOKOnly + vbInformation, "Information"
Else
        With rsGlb
                'Making sure that the user wants to save the record
                If MsgBox("Are You Sure You Wish To Update This Record?", vbYesNo + vbQuestion, "Update This Record?") = vbYes Then
                  '  .Fields(0) = txtCustomerID.Text
                    .Fields(1) = Text1.Text
                    .Fields(2) = Text2.Text
                    .Fields(3) = Text3.Text
                    .Fields(4) = Text4.Text
                    .Fields(5) = Text5.Text
                    .Fields(6) = Text6.Text
                    .Update
                    
                    MsgBox "The Record Was Saved Successfully!", vbInformation, "Succesful Save Procedure"
                End If
            End With
            rsGlb.Close
End If
reset
End Sub

Private Sub Form_Load()
Call Connection
Call reset

End Sub
Function reset()
cnt = 0
    Dim rsE As New ADODB.Recordset
    rsE.Open "select max(EmpID) from Employee", conn, adOpenDynamic, adLockOptimistic
    With rsE
        If .EOF = True Then
            txtCustomerID.Text = 1
        Else
           cnt = Val(.Fields(0)) + 1
        End If
        txtCustomerID.Text = cnt
    End With
    rsE.Close
    'txtCustomerID.Text = ""
    Text1.Text = Date
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
     Text6.Text = ""
    Text1.Enabled = False
    cmdDelete.Enabled = False
    cmdUpdate.Enabled = False
    
End Function

Private Sub Text2_KeyPress(KeyAscii As Integer)
KeyAscii = character(KeyAscii)
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
KeyAscii = character(KeyAscii)
End Sub

Private Sub txtCustomerID_KeyPress(KeyAscii As Integer)
KeyAscii = number(KeyAscii)
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
KeyAscii = address(KeyAscii)
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
KeyAscii = mobnumber(KeyAscii)
End Sub



Public Sub Connection()
    Set conn = New ADODB.Connection
    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\aoms.mdb;Persist Security Info=False"
    conn.Open
    'i = 1
End Sub


Public Sub QuerySelect(s1 As String)
    Set rsGlb = New ADODB.Recordset
   'MsgBox s1
    With rsGlb
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .ActiveConnection = conn
        .SOURCE = s1
        .CursorLocation = adUseClient
        .Open
    End With
End Sub

