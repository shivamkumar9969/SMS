VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   7815
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   17955
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu Enquery 
      Caption         =   "Enquery"
   End
   Begin VB.Menu Admission 
      Caption         =   "Admission"
   End
   Begin VB.Menu Teacher 
      Caption         =   "Teacher"
   End
   Begin VB.Menu Attendence 
      Caption         =   "Attendence"
      Begin VB.Menu Student_Attendence 
         Caption         =   "Student Attendence"
      End
      Begin VB.Menu Teacher_Attendence 
         Caption         =   "Teacher Attendence"
      End
   End
   Begin VB.Menu Reports 
      Caption         =   "Reports"
   End
   Begin VB.Menu aa 
      Caption         =   ""
   End
   Begin VB.Menu aaa 
      Caption         =   ""
   End
   Begin VB.Menu aaaa 
      Caption         =   ""
   End
   Begin VB.Menu aaaaaa 
      Caption         =   ""
   End
   Begin VB.Menu aaaaaaaaaaaa 
      Caption         =   ""
   End
   Begin VB.Menu b 
      Caption         =   ""
   End
   Begin VB.Menu bb 
      Caption         =   ""
   End
   Begin VB.Menu bbb 
      Caption         =   ""
   End
   Begin VB.Menu bbbb 
      Caption         =   ""
   End
   Begin VB.Menu bbbbbb 
      Caption         =   ""
   End
   Begin VB.Menu c 
      Caption         =   ""
   End
   Begin VB.Menu cc 
      Caption         =   ""
   End
   Begin VB.Menu ccc 
      Caption         =   ""
   End
   Begin VB.Menu ccccc 
      Caption         =   ""
   End
   Begin VB.Menu cccccc 
      Caption         =   ""
   End
   Begin VB.Menu exit 
      Caption         =   "EXIT"
   End
   Begin VB.Menu logout 
      Caption         =   "LOGOUT"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Admission_Click()
Unload Me
AdmissionForm.Show
End Sub

Private Sub Enquery_Click()
Unload Me
EnquiryForm.Show
End Sub


Private Sub Student_Click()
AdmissionPayForm.Show
End Sub


Private Sub exit_Click()
Unload Me
End Sub

Private Sub logout_Click()
LoginForm.Show
Unload Me
Unload MDIForm1
End Sub

Private Sub Teacher_Click()
TeacherReg.Show
End Sub
