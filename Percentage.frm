VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Percentage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pay Percentage"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6960
   Icon            =   "Percentage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "Percentage.frx":1708A
   ScaleHeight     =   4260
   ScaleWidth      =   6960
   Begin VB.CommandButton Command3 
      Caption         =   "Edit"
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
      Left            =   3120
      TabIndex        =   12
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
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
      Left            =   5520
      TabIndex        =   11
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1320
      TabIndex        =   9
      Top             =   1230
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   840
      Width           =   3855
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   510
      Width           =   1455
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
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   1680
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2055
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   6720
      _ExtentX        =   11853
      _ExtentY        =   3625
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
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   150
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Percentage:"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   1230
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Course Name:"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   870
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Course ID:"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   555
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Instructor ID:"
      Height          =   255
      Left            =   4320
      TabIndex        =   2
      Top             =   195
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pay Percentage Entry"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2535
   End
End
Attribute VB_Name = "Percentage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection

Private Sub Command1_Click()
If Text1.Text = "" Or Combo1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Then
MsgBox ("Enter all fields")
Exit Sub
End If
If Command1.Caption = "Update" Then
    Dim strsql1 As String
    strsql1 = "SELECT * from Percentage where SN=" & DataGrid1.Columns(0).Text & ""
    Set rs1 = New ADODB.Recordset
    rs1.CursorLocation = adUseClient
    rs1.Open strsql1, con, adOpenForwardOnly, adLockOptimistic
    rs1.Update("Course ID") = Combo1.Text
    rs1.Update("Percentage") = Text3.Text
    Command1.Caption = "Save"
    rs1.Close
Else
    Dim strsql As String
    strsql = "SELECT * from Percentage"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strsql, con, adOpenForwardOnly, adLockOptimistic
    rs.AddNew
    rs.Fields("Instructor ID") = Text1.Text
    rs.Fields("Course ID") = Combo1.Text
    rs.Fields("Percentage") = Text3.Text
    rs.Update
    rs.Close
End If
Text2.Text = ""
Text3.Text = ""
Call datashow
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Text1.Text = DataGrid1.Columns(1).Text
Combo1.Text = DataGrid1.Columns(2).Text
Text2.Text = DataGrid1.Columns(3).Text
Text3.Text = DataGrid1.Columns(4).Text
Command1.Caption = "Update"
End Sub

Private Sub Form_Load()
Set con = CreateObject("Adodb.connection")
con.Provider = "Microsoft.jet.oledb.4.0"
con.Open App.Path & "\MainData.mdb"
Text1.Text = InstructorList.DataGrid1.Columns(0).Text
Call datashow
Combo1.Clear
Set rs1 = New ADODB.Recordset
rs1.Open "select * from course", con, adOpenForwardOnly, adLockOptimistic
While Not rs1.EOF
Combo1.AddItem rs1.Fields("Course ID")
rs1.MoveNext
Wend
rs1.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
'For Instructor
Dim strsql3 As String
strsql3 = "SELECT Instructor.[Instructor ID], Sum(([Amount]*[percentage]/100)) AS TAmount FROM Instructor INNER JOIN (Percentage INNER JOIN Training ON Percentage.[Course ID] = Training.[Course ID]) ON (Instructor.[Instructor ID] = Training.[Instructor ID]) AND (Instructor.[Instructor ID] = Percentage.[Instructor ID]) GROUP BY Instructor.[Instructor ID]"
Set rs = New ADODB.Recordset
rs.Open strsql3, con, adOpenForwardOnly, adLockOptimistic
While Not rs.EOF
Dim AID1 As String
AID1 = rs.Fields("Instructor ID")
    Dim strsql2 As String
    strsql2 = "SELECT * From SumOfAmount where [Acc Head ID]='" & AID1 & "'"
    Set rs1 = New ADODB.Recordset
    rs1.Open strsql2, con, adOpenForwardOnly, adLockOptimistic
    rs1.Update("TotalPayment") = rs.Fields("TAmount")
    rs1.Close
rs.MoveNext
Wend
rs.Close
End Sub

Private Sub Text2_GotFocus()
If Combo1.Text = "" Then Exit Sub
Dim strsql As String
strsql = "select * from course where [course id]='" & Combo1.Text & "'"
Set rs1 = New ADODB.Recordset
rs1.Open strsql, con, adOpenForwardOnly, adLockOptimistic
Text2.Text = rs1.Fields("Course Name")
rs1.Close
End Sub
Sub datashow()
Dim strsql As String
strsql = "SELECT Percentage.SN, Percentage.[Instructor ID], Course.[Course ID], Course.[Course Name], Percentage.Percentage FROM Course LEFT JOIN Percentage ON Course.[Course ID] = Percentage.[Course ID] where Percentage.[Instructor ID]='" & Text1.Text & "'"
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open strsql, con, adOpenForwardOnly, adLockOptimistic
Set DataGrid1.DataSource = rs
DataGrid1.Columns(0).Width = 500
DataGrid1.Columns(1).Width = 1000
DataGrid1.Columns(2).Width = 800
DataGrid1.Columns(3).Width = 2800
DataGrid1.Columns(4).Alignment = dbgCenter
DataGrid1.Columns(4).Width = 1000
DataGrid1.Refresh
End Sub
