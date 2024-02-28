VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   7440
   ClientLeft      =   225
   ClientTop       =   1170
   ClientWidth     =   11280
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnu_cust 
      Caption         =   "CUSTOMER"
      Begin VB.Menu mnu_add_cust 
         Caption         =   "ADD CUSTOMER"
      End
      Begin VB.Menu nmu_srch 
         Caption         =   "SEARCH CUSTOMER"
      End
   End
   Begin VB.Menu mnu_ticket 
      Caption         =   "TICKET BOOKING "
      Begin VB.Menu mnu_book_ticket 
         Caption         =   "BOOK TIKCET "
      End
      Begin VB.Menu chk_stus 
         Caption         =   "CHECK STATUS"
      End
      Begin VB.Menu cnc_tik 
         Caption         =   "CANCEL TIKCET "
      End
   End
   Begin VB.Menu CFS 
      Caption         =   "UPDATE FLIGHT STATUS"
   End
   Begin VB.Menu ADD_FLI 
      Caption         =   "ADD FLIGHT "
   End
   Begin VB.Menu f_r 
      Caption         =   "FLIGHT WISE REPORT "
   End
   Begin VB.Menu mnu_emp 
      Caption         =   "EMPLOYEE"
      Begin VB.Menu add_emp 
         Caption         =   "ADD EMPLOYEE"
      End
      Begin VB.Menu U_E_R 
         Caption         =   "UPDATE EMPLOYEE SHIFT"
      End
      Begin VB.Menu E_s 
         Caption         =   "CHECK EMPLOYEE SHIFT"
      End
   End
   Begin VB.Menu mnu_brc 
      Caption         =   "BOARDING COUNTER"
      Begin VB.Menu mnu_bp 
         Caption         =   "GENERATE BOARDING PASS"
      End
   End
   Begin VB.Menu R_ALOO 
      Caption         =   "RUNWAY ALLOCATION"
   End
   Begin VB.Menu A_M 
      Caption         =   "AIRPLANE MAINTENANCE "
   End
   Begin VB.Menu FDB 
      Caption         =   "FEEDBACK"
   End
   Begin VB.Menu rpt_gen 
      Caption         =   "REPORT GENERATION"
      Begin VB.Menu emprpt 
         Caption         =   "GENERATE EMPLOYEE REPORT"
      End
      Begin VB.Menu cst_rpt 
         Caption         =   "GENERATE CUSTOMER REPORT"
      End
      Begin VB.Menu sr 
         Caption         =   "GENERATE EMPLOYEE SHIFT REPORT"
      End
      Begin VB.Menu FR 
         Caption         =   "GENERATE FLIGHT REPORT"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub A_M_Click()
Me.Hide
FLIGHT_MAIN.Show

End Sub

Private Sub add_emp_Click()
frmEmployee.Show

End Sub

Private Sub ADD_FLI_Click()
Me.Hide
ADD_FLIGHT.Show

End Sub

Private Sub CFS_Click()
Me.Hide
Form5.Show


End Sub

Private Sub chk_stus_Click()
Me.Hide
Form6.Show

End Sub

Private Sub cnc_tik_Click()
Me.Hide
Cancel.Show

End Sub

Private Sub cst_rpt_Click()
DataReport4.Show

End Sub

Private Sub E_s_Click()
Me.Hide
CHECK_SHIFT.Show

End Sub

Private Sub emprpt_Click()
DataReport5.Show

End Sub



Private Sub Command1_Click()
Me.Hide
MDIForm1.Show

End Sub






Private Sub Form_Load()
    
End Sub




Private Sub f_r_Click()
FLIGHT_WISE_REPORT.Show

End Sub

Private Sub FDB_Click()
Me.Hide
FeedBack.Show


End Sub

Private Sub frm_ATC_Click()
     Me.Hide
frmATC.Show


End Sub

Private Sub fr_Click()
Me.Hide

DataReport7.Show


End Sub

Private Sub mnu_add_cust_Click()
Me.Hide
frmCUSTOMER.Show
End Sub


Private Sub mnu_Aoc_Click()
Me.Hide
frmAOF.Show

End Sub

Private Sub mnu_book_ticket_Click()
Me.Hide
Form2.Show
End Sub

Private Sub mnu_bp_Click()
Me.Hide

frmBOARDING.Show

End Sub

Private Sub nmu_srch_Click()
Me.Hide
frmCustomersID.Show

End Sub

Private Sub R_ALOO_Click()
Me.Hide
RUNWAY_ALLOCATION.Show


End Sub

Private Sub sr_Click()
DataReport6.Show


End Sub

Private Sub U_E_R_Click()
Me.Hide
Form7.Show


End Sub
