VERSION 5.00
Begin VB.Form StudentEntry 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Student Entry"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5610
   Icon            =   "StudentEntry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "StudentEntry.frx":1708A
   ScaleHeight     =   3.938
   ScaleMode       =   5  'Inch
   ScaleWidth      =   3.896
   Begin VB.ComboBox Combo7 
      Height          =   315
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1920
      Width           =   1335
   End
   Begin VB.ComboBox Combo6 
      Height          =   315
      Left            =   4320
      TabIndex        =   11
      Top             =   4800
      Width           =   735
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      Left            =   3360
      TabIndex        =   10
      Top             =   4800
      Width           =   615
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   1920
      TabIndex        =   9
      Top             =   4800
      Width           =   975
   End
   Begin VB.TextBox Text9 
      Height          =   302
      Left            =   1920
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      Top             =   960
      Width           =   3615
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   1920
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   3840
      Width           =   3615
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1920
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   2880
      Width           =   3615
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1920
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   2400
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   317
      Left            =   360
      TabIndex        =   12
      Top             =   5235
      Width           =   1320
   End
   Begin VB.CommandButton Command3 
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
      Height          =   317
      Left            =   3960
      TabIndex        =   13
      Top             =   5235
      Width           =   1320
   End
   Begin VB.CommandButton Command2 
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
      Height          =   317
      Left            =   2160
      TabIndex        =   14
      Top             =   5235
      Width           =   1320
   End
   Begin VB.TextBox Text7 
      Height          =   315
      Left            =   1920
      TabIndex        =   8
      Top             =   4320
      Width           =   3615
   End
   Begin VB.TextBox Text5 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   315
      Left            =   1920
      TabIndex        =   6
      Top             =   3360
      Width           =   1440
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   1920
      TabIndex        =   2
      Top             =   1440
      Width           =   3600
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Gender:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   28
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label14 
      Caption         =   "Year"
      Height          =   255
      Left            =   5054
      TabIndex        =   27
      Top             =   4890
      Width           =   375
   End
   Begin VB.Label Label13 
      Caption         =   "Month"
      Height          =   255
      Left            =   2880
      TabIndex        =   26
      Top             =   4890
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Day"
      Height          =   255
      Left            =   3960
      TabIndex        =   25
      Top             =   4890
      Width           =   375
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Student ID:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   720
      TabIndex        =   24
      Top             =   540
      Width           =   1050
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "First Name:"
      Height          =   315
      Left            =   120
      TabIndex        =   23
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name:"
      Height          =   315
      Left            =   240
      TabIndex        =   22
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      Height          =   315
      Left            =   240
      TabIndex        =   21
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "City:"
      Height          =   255
      Left            =   360
      TabIndex        =   20
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Phone:"
      Height          =   315
      Left            =   360
      TabIndex        =   19
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Parent's Name:"
      Height          =   315
      Left            =   240
      TabIndex        =   18
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Academic Qualification:"
      Height          =   315
      Left            =   0
      TabIndex        =   17
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Join:"
      Height          =   315
      Left            =   360
      TabIndex        =   16
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00008080&
      BackStyle       =   1  'Opaque
      Height          =   450
      Left            =   120
      Top             =   5160
      Width           =   5415
   End
   Begin VB.Line Line1 
      X1              =   0.5
      X2              =   3.917
      Y1              =   0.25
      Y2              =   0.25
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Student Details"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "studententry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection

Private Sub Command1_Click()
On Error GoTo myerror:

If Text1 = "" Or Text2 = "" Or Combo7 = "" Or Combo1 = "" Or Combo2 = "" Or Text5 = "" Or Combo3 = "" Or Text7 = "" Or Text9 = "" Then
MsgBox "Fill up all fields"
Exit Sub
End If

Dim strsql As String
strsql = "Select * from Student"
Set rs = New ADODB.Recordset
rs.Open strsql, con, adOpenForwardOnly, adLockOptimistic
If MsgBox("           Are you sure(Y/N)???", 4) = 7 Then Exit Sub

rs.AddNew
If Not rs.EOF Then
    rs.Fields("Student ID") = Text9.Text
    rs.Fields("First Name") = UCase$(Text1.Text)
    rs.Fields("Last Name") = UCase$(Text2.Text)
    rs.Fields("Gender") = UCase$(Combo7.Text)
    rs.Fields("Address") = UCase$(Combo1.Text)
    rs.Fields("City") = UCase$(Combo2.Text)
    rs.Fields("Phone") = Text5.Text
    rs.Fields("Qualification") = UCase$(Combo3.Text)
    rs.Fields("Parent Name") = UCase$(Text7.Text)
    rs.Fields("Date of Join") = Combo4.Text & "/" & Combo5.Text & "/" & Combo6.Text
    rs.Fields("Months") = Combo4.Text
    rs.Update
End If
rs.Close

Set rs1 = New ADODB.Recordset
rs1.Open "select * from SumOfAmount", con, adOpenForwardOnly, adLockOptimistic
rs1.AddNew
rs1.Fields("Acc Head ID") = Text9.Text
rs1.Update
rs1.Close
Call Command2_Click
myerror:
Call error

End Sub

Private Sub Command2_Click()
Text1.Text = Clear
Text2.Text = Clear
'Combo7.Text = Clear
Combo1.Text = Clear
Combo2.Text = Clear
Text5.Text = Clear
Combo3.Text = Clear
Text7.Text = Clear
Combo4.Text = Clear
Combo5.Text = Clear
Combo6.Text = Clear
Text9.Text = Clear
Text1.SetFocus
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo myerror:

Set con = CreateObject("Adodb.connection")
con.Provider = "Microsoft.jet.oledb.4.0"
con.Open App.Path & "\MainData.mdb"

Dim strsql As String
strsql = "Select distinct Address from Student"
Set rs = New ADODB.Recordset
rs.Open strsql, con, adOpenForwardOnly, adLockOptimistic

Combo7.AddItem "FEMALE"
Combo7.AddItem "MALE"

While Not rs.EOF
Combo1.AddItem rs.Fields("Address")
rs.MoveNext
Wend
rs.Close

Dim strsql1 As String
strsql1 = "Select distinct City from Student"
Set rs1 = New ADODB.Recordset
rs1.Open strsql1, con, adOpenForwardOnly, adLockOptimistic

While Not rs1.EOF
Combo2.AddItem rs1.Fields("City")
rs1.MoveNext
Wend
rs1.Close

Dim strsql2 As String
strsql2 = "Select distinct Qualification from Student"
Set rs2 = New ADODB.Recordset
rs2.Open strsql2, con, adOpenForwardOnly, adLockOptimistic

While Not rs2.EOF
Combo3.AddItem rs2.Fields("Qualification")
rs2.MoveNext
Wend
rs2.Close
Text9.Enabled = False
'Date setting
    Combo4.AddItem "January"
    Combo4.AddItem "February"
    Combo4.AddItem "March"
    Combo4.AddItem "April"
    Combo4.AddItem "May"
    Combo4.AddItem "June"
    Combo4.AddItem "July"
    Combo4.AddItem "August"
    Combo4.AddItem "September"
    Combo4.AddItem "October"
    Combo4.AddItem "November"
    Combo4.AddItem "December"
For i = 1 To 31
    Combo5.AddItem i
Next i
For i = 2000 To 2099
    Combo6.AddItem i
Next i
myerror:
Call error

End Sub

Private Sub Form_Unload(Cancel As Integer)
con.Close
End Sub

Private Sub Text1_GotFocus()
On Error GoTo myerror:

Set rs3 = New ADODB.Recordset
rs3.Open "Student", con, adOpenForwardOnly, adLockOptimistic
If rs3.EOF = True Then
    Text9.Text = "S" & 1
    GoTo ext:
End If
rs3.MoveLast
Dim strcode As String
strcode = rs3.Fields("student ID")
Text9.Text = "S" & Val(Mid$(strcode, 2, Len(strcode) - 1)) + 1
ext:
rs3.Close
myerror:
Call error

End Sub
Private Sub error()
If Err.Number <> 0 Then
    MsgBox ("Error: " & Err.Description)
    Exit Sub
End If
End Sub

