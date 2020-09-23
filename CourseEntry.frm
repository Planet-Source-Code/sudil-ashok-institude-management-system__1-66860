VERSION 5.00
Begin VB.Form CourseEntry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Course Entry"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4650
   Icon            =   "CourseEntry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "CourseEntry.frx":1708A
   ScaleHeight     =   3450
   ScaleWidth      =   4650
   Begin VB.TextBox Text4 
      Height          =   350
      Left            =   1080
      TabIndex        =   4
      Top             =   1755
      Width           =   795
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   1080
      TabIndex        =   3
      Top             =   1335
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   1080
      TabIndex        =   2
      Top             =   900
      Width           =   3195
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1080
      TabIndex        =   1
      Top             =   495
      Width           =   1335
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   4920
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Image Command2 
      Height          =   675
      Left            =   2760
      MouseIcon       =   "CourseEntry.frx":2BC4A
      MousePointer    =   99  'Custom
      Picture         =   "CourseEntry.frx":2BF54
      Stretch         =   -1  'True
      ToolTipText     =   "Exit"
      Top             =   2640
      Width           =   720
   End
   Begin VB.Image Command1 
      Height          =   825
      Left            =   960
      MouseIcon       =   "CourseEntry.frx":2C6FE
      MousePointer    =   99  'Custom
      Picture         =   "CourseEntry.frx":2CA08
      ToolTipText     =   "Save"
      Top             =   2520
      Width           =   840
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Duration:"
      Height          =   345
      Left            =   0
      TabIndex        =   9
      Top             =   1755
      Width           =   1005
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Course Name:"
      Height          =   345
      Left            =   0
      TabIndex        =   8
      Top             =   900
      Width           =   1035
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Course ID:"
      Height          =   345
      Index           =   0
      Left            =   0
      TabIndex        =   7
      Top             =   540
      Width           =   1035
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Rate: (Rs.)"
      Height          =   345
      Index           =   1
      Left            =   0
      TabIndex        =   6
      Top             =   1335
      Width           =   1005
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Months"
      Height          =   240
      Left            =   1920
      TabIndex        =   5
      Top             =   1860
      Width           =   615
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Course Entry"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "CourseEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection
Dim rs As ADODB.Recordset

Private Sub Combo1_DropDown()
On Error GoTo myerror:

Dim strsql As String
strsql = "Select * from Course"
Set rs = New ADODB.Recordset
rs.Open strsql, con, adOpenForwardOnly, adLockOptimistic

Combo1.Clear
While Not rs.EOF
Combo1.AddItem rs.Fields("Course ID")
rs.MoveNext
Wend
rs.Close
myerror:
Call error
End Sub

Private Sub Command1_Click()
On Error GoTo myerror:

If Combo1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Then
MsgBox "Fill up all selected fields"
Exit Sub
End If

Dim strsql As String
strsql = "Select * from Course"
Set rs = New ADODB.Recordset
rs.Open strsql, con, adOpenForwardOnly, adLockOptimistic

If MsgBox("         Are you sure(Y/N)???", 4) = 7 Then Exit Sub
While Not rs.EOF
If rs.Fields("Course ID") = Combo1.Text Then
MsgBox ("Enter new 'Course ID' !!!!")
Exit Sub
End If
rs.MoveNext
Wend

rs.AddNew
If Not rs.EOF Then
rs.Fields("Course ID") = Combo1.Text
rs.Fields("Course Name") = UCase$(Text2.Text)
rs.Fields("Rate") = Text3.Text
rs.Fields("Duration") = UCase$(Text4.Text)
rs.Update
End If
rs.Close
Combo1.Text = Clear
Text2.Text = Clear
Text3.Text = Clear
Text4.Text = Clear
myerror:
Call error

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo myerror:

Set con = CreateObject("Adodb.connection")
con.Provider = "Microsoft.jet.oledb.4.0"
con.Open App.Path & "\MainData.mdb"
myerror:
Call error
End Sub

Private Sub error()
If Err.Number <> 0 Then
    MsgBox ("Error: " & Err.Description)
    Exit Sub
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
con.Close
End Sub

