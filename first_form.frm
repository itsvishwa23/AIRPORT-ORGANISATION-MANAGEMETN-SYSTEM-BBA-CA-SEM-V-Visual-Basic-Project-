VERSION 5.00
Begin VB.Form first_form 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form5"
   ClientHeight    =   8565
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   10830
   LinkTopic       =   "Form5"
   Picture         =   "first_form.frx":0000
   ScaleHeight     =   8565
   ScaleWidth      =   10830
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   4800
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   8400
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "Bahnschrift"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5040
      MaskColor       =   &H80000008&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9720
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      TabIndex        =   0
      Text            =   "USERNAME"
      Top             =   6360
      Width           =   3855
   End
End
Attribute VB_Name = "first_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "USERNAME" And Text2.Text = "" Then
Me.Hide
Loading.Show
 End If
 



End Sub

