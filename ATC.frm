VERSION 5.00
Begin VB.Form frmATC 
   Caption         =   "Form2"
   ClientHeight    =   8565
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18510
   LinkTopic       =   "Form2"
   Picture         =   "ATC.frx":0000
   ScaleHeight     =   8565
   ScaleWidth      =   18510
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "<< GO TO HOMEPAGE"
      Height          =   735
      Left            =   15720
      TabIndex        =   1
      Top             =   3360
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Request Runway"
      Height          =   1095
      Left            =   15600
      TabIndex        =   0
      Top             =   1920
      Width           =   2055
   End
End
Attribute VB_Name = "frmATC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Me.Hide
Form1.Show
End Sub
