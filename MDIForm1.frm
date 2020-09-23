VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Institute Management System [Copyright © Regider Software 2006]"
   ClientHeight    =   10710
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   15240
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":1708A
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   15180
      TabIndex        =   0
      Top             =   0
      Width           =   15240
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu New 
         Caption         =   "New Entry"
         Begin VB.Menu StudentDetails 
            Caption         =   "Student Details"
         End
         Begin VB.Menu CoursePackage 
            Caption         =   "Course Package"
         End
         Begin VB.Menu CourseDetails 
            Caption         =   "Course Details"
         End
         Begin VB.Menu Instructor_Details 
            Caption         =   "Instructor Details"
         End
         Begin VB.Menu Account_Head 
            Caption         =   "Account Head"
         End
      End
      Begin VB.Menu UserEntry 
         Caption         =   "Create User"
      End
      Begin VB.Menu change 
         Caption         =   "Change Password"
      End
      Begin VB.Menu separator 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "E&xit"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu Transaction_menu 
      Caption         =   "Transaction"
      Begin VB.Menu Payment 
         Caption         =   "Collection from Students"
      End
      Begin VB.Menu Cash_Transaction 
         Caption         =   "Cash Transaction"
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu Browse 
      Caption         =   "&Browse"
      Begin VB.Menu StudentsList 
         Caption         =   "Students List"
         Shortcut        =   ^S
      End
      Begin VB.Menu Instructor 
         Caption         =   "Instructor List"
      End
      Begin VB.Menu Courseslist 
         Caption         =   "Courses List"
      End
      Begin VB.Menu Collection 
         Caption         =   "Collection/Payment"
      End
   End
   Begin VB.Menu Issue 
      Caption         =   "&Issue"
      Begin VB.Menu Certificate 
         Caption         =   "Issue Certificate"
         Shortcut        =   ^I
      End
   End
   Begin VB.Menu Report 
      Caption         =   "&Report"
      Begin VB.Menu Students 
         Caption         =   "Student's Details"
      End
      Begin VB.Menu Courses 
         Caption         =   "Coures Details"
      End
      Begin VB.Menu Trainings 
         Caption         =   "Training Details"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu Cash 
         Caption         =   "Cash Transaction"
      End
      Begin VB.Menu Accounttype 
         Caption         =   "Account Type Wise"
      End
      Begin VB.Menu AccountHeadWise 
         Caption         =   "Account Head Wise"
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu Statistics_Report 
         Caption         =   "Statistics Report"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Account_Head_Click()
AccountHead.Show
End Sub


Private Sub AccountHeadWise_Click()
rams = 7
Dateselect.Show
End Sub

Private Sub Accounttype_Click()
rams = 8
Dateselect.Show
End Sub

Private Sub Cash_Click()
rams = 1
Dateselect.Show
End Sub

Private Sub Cash_Transaction_Click()
Transaction.Show
End Sub

Private Sub Certificate_Click()
List.Show
End Sub

Private Sub change_Click()
rams = 2
users.Show
End Sub

Private Sub Collection_Click()
Collect.Show
End Sub

Private Sub CourseDetails_Click()
SubCourse.Show
End Sub

Private Sub CoursePackage_Click()
CourseEntry.Show
End Sub

Private Sub Courses_Click()
DataEnvironment1.Connection1.ConnectionString = App.Path & "\MainData.mdb"
DataEnvironment1.Connection1.Open
DataEnvironment1.Command5_Grouping
DataReport5.Show
End Sub

Private Sub Courseslist_Click()
CourseList.Show
End Sub


Private Sub Exit_Click()
Unload Me
End Sub

Private Sub Instructor_Click()
InstructorList.Show
End Sub

Private Sub Instructor_Details_Click()
rams1 = 4
InstructorDetails.Show
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
Dim soc As String
soc = "" & App.Path & "\MainData.mdb"
FileSystem.SetAttr soc, vbReadOnly

End Sub



Private Sub Payment_Click()
rams1 = 1
Receive.Show
End Sub

Private Sub Status_Click()
DataReport2.Show
End Sub

Private Sub Statistics_Report_Click()
Statistics.Show
End Sub

Private Sub StudentDetails_Click()
studententry.Show
End Sub

Private Sub Students_Click()
DataEnvironment1.Connection1.ConnectionString = App.Path & "\MainData.mdb"
DataEnvironment1.Connection1.Open
DataEnvironment1.Command4_Grouping
DataReport4.Show
End Sub

Private Sub StudentsList_Click()
StudentList.Show
End Sub

Private Sub Trainings_Click()
DataEnvironment1.Connection1.ConnectionString = App.Path & "\MainData.mdb"
DataEnvironment1.Connection1.Open
DataEnvironment1.Command6_Grouping
DataReport6.Show
End Sub

Private Sub UserEntry_Click()
rams = 1
users.Show
End Sub
