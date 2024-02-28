VERSION 5.00
Begin VB.Form frmAOF 
   Caption         =   "AIRPORT OPERATIONS "
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15945
   LinkTopic       =   "Form2"
   Picture         =   "AOF.frx":0000
   ScaleHeight     =   10935
   ScaleWidth      =   15945
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "<< GO TO HOME PAGE"
      Height          =   735
      Left            =   4200
      TabIndex        =   6
      Top             =   9840
      Width           =   2655
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "AOF.frx":B6222
      Left            =   6360
      List            =   "AOF.frx":B6232
      TabIndex        =   5
      Top             =   8520
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   4680
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   7680
      Width           =   3615
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   4080
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   6840
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   4560
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   3000
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   3960
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   2160
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ALLOCATE RUNWAY "
      Height          =   735
      Left            =   1320
      TabIndex        =   0
      Top             =   9840
      Width           =   2655
   End
End
Attribute VB_Name = "frmAOF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Me.Hide
Form1.Show

End Sub

