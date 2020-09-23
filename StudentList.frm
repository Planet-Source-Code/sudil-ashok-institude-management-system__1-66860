VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form StudentList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Student List"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10185
   Icon            =   "StudentList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "StudentList.frx":1708A
   ScaleHeight     =   6585
   ScaleWidth      =   10185
   Begin VB.TextBox Text7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Left            =   9000
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   6240
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Coursed Offered"
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
      Left            =   4320
      TabIndex        =   12
      Top             =   75
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5175
      Left            =   75
      TabIndex        =   0
      Top             =   960
      Width           =   10020
      _ExtentX        =   17674
      _ExtentY        =   9128
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   14737632
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text9 
      Height          =   300
      Left            =   9090
      TabIndex        =   7
      Top             =   585
      Width           =   735
   End
   Begin VB.TextBox Text6 
      Height          =   300
      Left            =   8130
      TabIndex        =   6
      Top             =   585
      Width           =   945
   End
   Begin VB.TextBox Text5 
      Height          =   300
      Left            =   7290
      TabIndex        =   5
      Top             =   585
      Width           =   825
   End
   Begin VB.TextBox Text4 
      Height          =   300
      Left            =   4830
      TabIndex        =   4
      Top             =   585
      Width           =   2475
   End
   Begin VB.TextBox Text3 
      Height          =   300
      Left            =   3300
      TabIndex        =   3
      Top             =   585
      Width           =   1515
   End
   Begin VB.TextBox Text2 
      Height          =   300
      Left            =   1125
      TabIndex        =   2
      Top             =   585
      Width           =   2160
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   375
      TabIndex        =   1
      Top             =   585
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Close"
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
      Left            =   8280
      TabIndex        =   11
      Top             =   75
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Payment"
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
      Left            =   6360
      TabIndex        =   10
      Top             =   75
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Student Details"
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
      Left            =   2280
      TabIndex        =   9
      Top             =   75
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Find>>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   300
      TabIndex        =   8
      Top             =   75
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Total Balance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   220
      Left            =   7800
      TabIndex        =   14
      Top             =   6260
      Width           =   1095
   End
   Begin VB.Line Line1 
      X1              =   15
      X2              =   10095
      Y1              =   525
      Y2              =   525
   End
End
Attribute VB_Name = "StudentList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection

Private Sub Command1_Click()
CourseOffered.Show
End Sub

Private Sub Command2_Click()
On Error GoTo myerror:
Call datashow
myerror:
Call error

End Sub

Private Sub Command3_Click()
StudentEdit.Show
End Sub

Private Sub Command4_Click()
CourseOffered.Show
End Sub

Private Sub Command5_Click()
rams1 = 2
Receive.Show
End Sub

Private Sub Command6_Click()
On Error GoTo myerror:

If DataGrid1.Columns(0).Text = "" Then MsgBox ("Select one record to edit"): Exit Sub
Dim strsql As String
strsql = "SELECT * from Student where [Student ID]=" & DataGrid1.Columns(0).Text & ""
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open strsql, con, adOpenForwardOnly, adLockOptimistic
If MsgBox("Are you sure to delete 'S.ID. " & DataGrid1.Columns(0).Text & "' Student Record???", 4) = 7 Then GoTo nodelete:
rs.Delete
nodelete:
rs.Close
Call datashow
myerror:
Call error

End Sub

Private Sub Command7_Click()
Unload Me
End Sub

Private Sub Command8_Click()
rams = 4
Payments.Show
End Sub

Private Sub DataGrid1_DblClick()
StudentEdit.Show
Call datashow
End Sub

Private Sub Form_Activate()
Call datashow
End Sub

Private Sub Form_Load()
On Error GoTo myerror:

Set con = CreateObject("Adodb.connection")
con.Provider = "Microsoft.jet.oledb.4.0"
con.Open App.Path & "\MainData.mdb"
Call datashow
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
Sub datashow()
Dim strsql As String
If Not Text1.Text = "" Then
    strsql = "SELECT Student.[Student ID], Student.[First Name], Student.[Last Name], Student.Address, Student.Phone, Student.[Date Of Join], [TotalPayment]-[TotalPaid] AS Balance FROM Student LEFT JOIN SumOfAmount ON Student.[Student ID]=SumOfAmount.[Acc Head ID] Where Student.[Student ID] " & " Like '" & Text1.Text & "%" & "'"
ElseIf Not Text2.Text = "" Then
    strsql = "SELECT Student.[Student ID], Student.[First Name], Student.[Last Name], Student.Address, Student.Phone, Student.[Date Of Join], [TotalPayment]-[TotalPaid] AS Balance FROM Student LEFT JOIN SumOfAmount ON Student.[Student ID]=SumOfAmount.[Acc Head ID] WHERE Student.[First Name]" & " Like '" & Text2.Text & "%" & "'"
ElseIf Not Text3.Text = "" Then
    strsql = "SELECT Student.[Student ID], Student.[First Name], Student.[Last Name], Student.Address, Student.Phone, Student.[Date Of Join], [TotalPayment]-[TotalPaid] AS Balance FROM Student LEFT JOIN SumOfAmount ON Student.[Student ID]=SumOfAmount.[Acc Head ID] WHERE Student.[Last Name]" & " Like '" & Text3.Text & "%" & "'"
ElseIf Not Text4.Text = "" Then
    strsql = "SELECT Student.[Student ID], Student.[First Name], Student.[Last Name], Student.Address, Student.Phone, Student.[Date Of Join], [TotalPayment]-[TotalPaid] AS Balance FROM Student LEFT JOIN SumOfAmount ON Student.[Student ID]=SumOfAmount.[Acc Head ID] WHERE student.[Address]" & " Like '" & Text4.Text & "%" & "'"
ElseIf Not Text5.Text = "" Then
    strsql = "SELECT Student.[Student ID], Student.[First Name], Student.[Last Name], Student.Address, Student.Phone, Student.[Date Of Join], [TotalPayment]-[TotalPaid] AS Balance FROM Student LEFT JOIN SumOfAmount ON Student.[Student ID]=SumOfAmount.[Acc Head ID] WHERE student.[Phone]" & " Like '" & Text5.Text & "%" & "'"
ElseIf Not Text6.Text = "" Then
    strsql = "SELECT Student.[Student ID], Student.[First Name], Student.[Last Name], Student.Address, Student.Phone, Student.[Date Of Join], [TotalPayment]-[TotalPaid] AS Balance FROM Student LEFT JOIN SumOfAmount ON Student.[Student ID]=SumOfAmount.[Acc Head ID] WHERE student.[Date Of Join]" & " Like '" & Text6.Text & "%" & "'"
ElseIf Not Text9.Text = "" Then
    strsql = "SELECT Student.[Student ID], Student.[First Name], Student.[Last Name], Student.Address, Student.Phone, Student.[Date Of Join], [TotalPayment]-[TotalPaid] AS Balance FROM Student LEFT JOIN SumOfAmount ON Student.[Student ID]=SumOfAmount.[Acc Head ID] WHERE Balance=>" & Text9.Text & ""
Else
    strsql = "SELECT Student.[Student ID], Student.[First Name], Student.[Last Name], Student.Address, Student.Phone, Student.[Date Of Join], [TotalPayment]-[TotalPaid] AS Balance FROM Student LEFT JOIN SumOfAmount ON Student.[Student ID]=SumOfAmount.[Acc Head ID] order by student.sno desc"
End If

Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open strsql, con, adOpenForwardOnly, adLockOptimistic

Dim balanc As Double
balanc = 0
While Not rs.EOF
balanc = balanc + rs.Fields("Balance")
rs.MoveNext
Wend
Text7.Text = balanc

Set DataGrid1.DataSource = rs
DataGrid1.Columns(0).Caption = "S.ID"
DataGrid1.Columns(0).Width = 700
DataGrid1.Columns(1).Width = 2200
DataGrid1.Columns(2).Width = 1500
DataGrid1.Columns(3).Width = 2500
DataGrid1.Columns(4).Width = 800
DataGrid1.Columns(5).Width = 1000
DataGrid1.Columns(6).Width = 700
DataGrid1.Columns(6).Alignment = dbgRight
DataGrid1.Refresh
'rs.Close

End Sub
