VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form InstructorList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Instructors"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9975
   Icon            =   "InstructorList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "InstructorList.frx":1708A
   ScaleHeight     =   3750
   ScaleWidth      =   9975
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2685
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   4736
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   -2147483629
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
   Begin VB.Image Command4 
      Height          =   555
      Left            =   8280
      MouseIcon       =   "InstructorList.frx":2BC4A
      MousePointer    =   99  'Custom
      Picture         =   "InstructorList.frx":2BF54
      ToolTipText     =   "Exit"
      Top             =   3000
      Width           =   600
   End
   Begin VB.Image Command3 
      Height          =   720
      Left            =   5520
      MouseIcon       =   "InstructorList.frx":2C6FE
      MousePointer    =   99  'Custom
      Picture         =   "InstructorList.frx":2CA08
      ToolTipText     =   "Pay Percentage"
      Top             =   2880
      Width           =   720
   End
   Begin VB.Image Command2 
      Height          =   675
      Left            =   3120
      MouseIcon       =   "InstructorList.frx":2D2B5
      MousePointer    =   99  'Custom
      Picture         =   "InstructorList.frx":2D5BF
      Stretch         =   -1  'True
      ToolTipText     =   "Student Group"
      Top             =   2880
      Width           =   705
   End
   Begin VB.Image Command1 
      Height          =   675
      Left            =   1080
      MouseIcon       =   "InstructorList.frx":2DA87
      MousePointer    =   99  'Custom
      Picture         =   "InstructorList.frx":2DD91
      Stretch         =   -1  'True
      ToolTipText     =   "Instructors Details"
      Top             =   2880
      Width           =   705
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   720
      Left            =   105
      Top             =   2895
      Width           =   9750
   End
End
Attribute VB_Name = "InstructorList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection

Private Sub Command1_Click()
rams1 = 3
InstructorDetails.Show
End Sub

Private Sub Command2_Click()
Group.Show
End Sub

Private Sub Command3_Click()
Percentage.Show
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Call Form_Load
End Sub

Private Sub Form_Load()
Set con = CreateObject("Adodb.connection")
con.Provider = "Microsoft.jet.oledb.4.0"
con.Open App.Path & "\maindata.mdb"
Dim strsql As String
strsql = "SELECT Instructor.[Instructor ID], Instructor.Name, Instructor.Address, Instructor.Phone, SumOfAmount.TotalPayment, SumOfAmount.TotalPaid, Sum([TotalPayment]-[TotalPaid]) AS Balance FROM Instructor INNER JOIN SumOfAmount ON Instructor.[Instructor ID] = SumOfAmount.[Acc Head ID] GROUP BY Instructor.[Instructor ID], Instructor.Name, Instructor.Address, Instructor.Phone, SumOfAmount.TotalPayment, SumOfAmount.TotalPaid"
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open strsql, con, adOpenForwardOnly, adLockOptimistic
Set DataGrid1.DataSource = rs
DataGrid1.Columns(0).Caption = "I. ID"
DataGrid1.Columns(0).Width = 400
DataGrid1.Columns(1).Width = 2400
DataGrid1.Columns(2).Width = 2800
DataGrid1.Columns(3).Width = 800
DataGrid1.Columns(4).Width = 1100
DataGrid1.Columns(4).Alignment = dbgRight
DataGrid1.Columns(5).Width = 800
DataGrid1.Columns(5).Alignment = dbgRight
DataGrid1.Columns(6).Width = 800
DataGrid1.Columns(6).Alignment = dbgRight
DataGrid1.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
con.Close
End Sub

