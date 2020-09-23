VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form CourseList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Course List"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6360
   Icon            =   "CourseList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "CourseList.frx":1708A
   ScaleHeight     =   4515
   ScaleWidth      =   6360
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1575
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   2778
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
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text4 
      Height          =   300
      Left            =   4440
      TabIndex        =   3
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   300
      Left            =   1560
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   300
      Left            =   1560
      TabIndex        =   1
      Top             =   840
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   1560
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
   Begin VB.Image Command4 
      Height          =   720
      Left            =   960
      MouseIcon       =   "CourseList.frx":2BC4A
      MousePointer    =   99  'Custom
      Picture         =   "CourseList.frx":2BF54
      ToolTipText     =   "Edit"
      Top             =   1680
      Width           =   720
   End
   Begin VB.Image Command3 
      Height          =   555
      Left            =   5040
      MouseIcon       =   "CourseList.frx":2C6A3
      MousePointer    =   99  'Custom
      Picture         =   "CourseList.frx":2C9AD
      ToolTipText     =   "Exit"
      Top             =   1800
      Width           =   600
   End
   Begin VB.Image Command2 
      Height          =   720
      Left            =   3720
      MouseIcon       =   "CourseList.frx":2D157
      MousePointer    =   99  'Custom
      Picture         =   "CourseList.frx":2D461
      ToolTipText     =   "Delete"
      Top             =   1680
      Width           =   720
   End
   Begin VB.Image Command1 
      Height          =   720
      Left            =   2400
      MouseIcon       =   "CourseList.frx":2DC42
      MousePointer    =   99  'Custom
      Picture         =   "CourseList.frx":2DF4C
      ToolTipText     =   "Update"
      Top             =   1680
      Width           =   720
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "(months)"
      Height          =   255
      Left            =   5340
      TabIndex        =   10
      Top             =   1290
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Duration:"
      Height          =   255
      Left            =   3660
      TabIndex        =   9
      Top             =   1245
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Rate(Rs):"
      Height          =   375
      Left            =   600
      TabIndex        =   8
      Top             =   1245
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Course Name:"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Course ID:"
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Package Courses"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   -15
      TabIndex        =   5
      Top             =   0
      Width           =   2565
   End
End
Attribute VB_Name = "CourseList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection
Private Sub Command1_Click()
On Error GoTo myerror:

Dim strsql As String
strsql = "Select * from Course where [Course ID]='" & Text1.Text & "'"
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open strsql, con, adOpenForwardOnly, adLockOptimistic

If MsgBox("         Are you sure(Y/N)???", 4) = 7 Then Exit Sub

rs.Update("Course Name") = UCase$(Text2.Text)
rs.Update("Rate") = Text3.Text
rs.Update("Duration") = Text4.Text
rs.Close
Call refreshdata
myerror:
Call error

End Sub

Private Sub Command2_Click()
On Error GoTo myerror:

Dim strsql As String
strsql = "Select * from Course where [Course ID]='" & Text1.Text & "'"
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open strsql, con, adOpenForwardOnly, adLockOptimistic

If MsgBox("         Are you sure(Y/N)???", 4) = 7 Then Exit Sub
rs.Delete
rs.Close
Call refreshdata
myerror:
Call error

End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
Text1.Text = DataGrid1.Columns(0).Text
Text2.Text = DataGrid1.Columns(1).Text
Text3.Text = DataGrid1.Columns(2).Text
Text4.Text = DataGrid1.Columns(3).Text
End Sub



Private Sub Form_Load()
On Error GoTo myerror:

Set con = CreateObject("Adodb.connection")
con.Provider = "Microsoft.jet.oledb.4.0"
con.Open App.Path & "\maindata.mdb"

Dim strsql As String
strsql = "Select * from Course order by course.[Course ID]"
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open strsql, con, adOpenForwardOnly, adLockOptimistic
Set DataGrid1.DataSource = rs

DataGrid1.Columns(0).Caption = "S.No."
DataGrid1.Columns(0).Width = 500
DataGrid1.Columns(1).Width = 3500
DataGrid1.Columns(2).Width = 800
DataGrid1.Columns(3).Width = 800
DataGrid1.Refresh
myerror:
Call error

End Sub
Sub refreshdata()
On Error GoTo myerror:

Dim strsql As String
strsql = "Select * from Course order by course.[Course ID]"
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open strsql, con, adOpenForwardOnly, adLockOptimistic
Set DataGrid1.DataSource = rs

DataGrid1.Columns(0).Caption = "S.No."
DataGrid1.Columns(0).Width = 500
DataGrid1.Columns(1).Width = 3500
DataGrid1.Columns(2).Width = 800
DataGrid1.Columns(3).Width = 800
DataGrid1.Refresh
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

