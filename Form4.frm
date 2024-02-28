VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4485
   LinkTopic       =   "Form4"
   ScaleHeight     =   3030
   ScaleWidth      =   4485
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   0
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "close"
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   1320
      Width           =   1815
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim billno As Integer
billno = CInt(Text1.Text)

DataEnvironment1.Connection1.Open
DataEnvironment1.Command2 billno
DataReport2.Show vbModal
DataEnvironment1.Connection1.Close
End Sub

Private Sub Command2_Click()
Unload Me
Form1.Show
End Sub

