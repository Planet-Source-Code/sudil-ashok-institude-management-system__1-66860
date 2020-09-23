VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form Statistics 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Statistics"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "Statistics.frx":0000
   ScaleHeight     =   5625
   ScaleWidth      =   11100
   Begin VB.CommandButton Command2 
      Caption         =   "<<Refresh>>"
      Height          =   375
      Left            =   9120
      TabIndex        =   1
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "M/d/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   320
      Left            =   8880
      TabIndex        =   0
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   375
      Left            =   9240
      TabIndex        =   2
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Frame Frame4 
      Caption         =   "Monthly Flow of Students"
      Height          =   5055
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   8295
      Begin MSChart20Lib.MSChart MSChart1 
         Height          =   4695
         Left            =   120
         OleObjectBlob   =   "Statistics.frx":14BC0
         TabIndex        =   3
         Top             =   240
         Width           =   8055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Gender"
      Height          =   1335
      Left            =   8520
      TabIndex        =   5
      Top             =   480
      Width           =   2415
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Students:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Female:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Male:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Shape Shape1 
      Height          =   1335
      Left            =   8520
      Top             =   3360
      Width           =   2415
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter year to be analysed"
      Height          =   315
      Left            =   8760
      TabIndex        =   13
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000013&
      Caption         =   "  Statistic Report  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   11175
   End
End
Attribute VB_Name = "Statistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()

Dim strsql As String
Dim strsql1 As String
Dim strsql2 As String

If Text4.Text = "" Then
strsql = "Select count([Student ID]) from student where Gender='MALE'"
strsql1 = "Select count([Student ID]) from student where Gender='FEMALE'"
strsql2 = "Select  Months, count([Student ID])as [No of Student] from student group by [months], month([Date Of Join]) order by month([Date Of Join])"

Else
strsql = "Select count([Student ID]) from student where Gender='MALE' AND year([Date Of Join])='" & Text4.Text & "'"
strsql1 = "Select count([Student ID]) from student where Gender='FEMALE' AND year([Date Of Join])= '" & Text4.Text & "'"
strsql2 = "Select [Months], count([Student ID])as [No of Student] from student where year([Date Of Join])= '" & Text4.Text & "'group by [months], month([Date Of Join]) order by month([Date Of Join])"

End If

Set rs = New ADODB.Recordset
rs.Open strsql, con, adOpenForwardOnly, adLockOptimistic
Text1 = rs.Fields(0)

Set rs1 = New ADODB.Recordset
rs1.Open strsql1, con, adOpenForwardOnly, adLockOptimistic
Text2 = rs1.Fields(0)
Text3 = rs.Fields(0) + rs1.Fields(0)

Set rs2 = New ADODB.Recordset
rs2.CursorLocation = adUseClient
rs2.Open strsql2, con, adOpenForwardOnly, adLockOptimistic
If rs2.EOF = False Then
rs2.MoveFirst
End If
Set MSChart1.DataSource = rs2

rs.Close
rs2.Close
End Sub

Private Sub Form_Load()
Set con = CreateObject("Adodb.connection")
con.Provider = "Microsoft.jet.oledb.4.0"
con.Open App.Path & "\maindata.mdb"
Call Command2_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
con.Close
End Sub
