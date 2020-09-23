VERSION 5.00
Begin VB.Form SubCourse 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sub Course"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4980
   Icon            =   "SubCourse.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "SubCourse.frx":1708A
   ScaleHeight     =   3315
   ScaleWidth      =   4980
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   3495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   7
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   6
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Top             =   840
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   960
      TabIndex        =   1
      Top             =   360
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Caption         =   "List of courses for Package"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   3015
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Courses:>>"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Package:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   360
      Width           =   1335
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   5040
      Y1              =   720
      Y2              =   720
   End
End
Attribute VB_Name = "SubCourse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection

Private Sub Combo1_lostfocus()
On Error GoTo myerror:
List1.Clear
If Combo1.Text = "" Then
MsgBox ("Select Course Name")
Exit Sub
End If

Dim strsql As String
Dim strsql1 As String

Set rs = New ADODB.Recordset
Set rs1 = New ADODB.Recordset

strsql = "Select * from Course where [Course Name]=" & "'" & Combo1.Text & "'"
rs.CursorLocation = adUseClient
rs.Open strsql, con, adOpenForwardOnly, adLockOptimistic
Text1.Text = rs.Fields("Course ID")
rs.Close
'*******
strsql1 = "Select * from [Sub course] where [Course ID]='" & Text1.Text & "'"
rs1.CursorLocation = adUseClient
rs1.Open strsql1, con, adOpenForwardOnly, adLockOptimistic
While Not rs1.EOF
List1.AddItem rs1.Fields("Courses")
rs1.MoveNext
Wend
rs1.Close

myerror:
Call error

End Sub

Private Sub Command1_Click()
Unload Me

End Sub

Private Sub Command2_Click()
On Error GoTo myerror:

If Text1.Text = "" Or Text2.Text = "" Then
MsgBox ("Select Package Course and then enter Course!!!")
Exit Sub
End If

Dim strsql As String
strsql = "Select * from [Sub Course]"
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open strsql, con, adOpenForwardOnly, adLockOptimistic

rs.AddNew
rs.Fields("Course ID") = UCase$(Text1.Text)
rs.Fields("Courses") = UCase$(Text2.Text)
rs.Update

Text2.Text = Clear
Call Combo1_lostfocus

myerror:
Call error

End Sub

Private Sub Command3_Click()
On Error GoTo myerror:

If Text1.Text = "" Or Text2.Text = "" Then
MsgBox ("Select Package Course and Course from list!!!")
Exit Sub
End If

Dim strsql As String
strsql = "Select * from [Sub Course] where [Course ID]='" & Text1.Text & "' and Courses='" & Text2.Text & "'"
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open strsql, con, adOpenForwardOnly, adLockOptimistic
rs.Delete

Text2.Text = Clear
Call Combo1_lostfocus

myerror:
Call error

End Sub

Private Sub Command4_Click()
Combo1.Text = Clear
Text1.Text = Clear
Text2.Text = Clear
List1.Clear
End Sub

Private Sub Form_Load()
On Error GoTo myerror:

Set con = CreateObject("Adodb.connection")
con.Provider = "Microsoft.jet.oledb.4.0"
con.Open App.Path & "\MainData.mdb"

Dim strsql As String
'strsql = "Select * from Course"
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open "Course", con, adOpenForwardOnly, adLockOptimistic

Combo1.Clear
While Not rs.EOF
Combo1.AddItem rs.Fields("Course Name")
rs.MoveNext
Wend
rs.Close

myerror:
Call error

End Sub

Private Sub Form_Unload(Cancel As Integer)
con.Close
End Sub

Private Sub List1_Click()
Text2.Text = List1.Text
End Sub

Private Sub error()
If Err.Number <> 0 Then
    MsgBox ("Error: " & Err.Description)
    Exit Sub
End If
End Sub
