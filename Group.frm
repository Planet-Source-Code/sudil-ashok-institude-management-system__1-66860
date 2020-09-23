VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Group 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Student Group"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9015
   Icon            =   "Group.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "Group.frx":1708A
   ScaleHeight     =   3135
   ScaleWidth      =   9015
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3015
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   5318
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Group"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection
Private Sub Form_Load()
Set con = CreateObject("Adodb.connection")
con.Provider = "Microsoft.jet.oledb.4.0"
con.Open App.Path & "\maindata.mdb"
Dim strsql As String
strsql = "SELECT Instructor.[Instructor ID], Training.[Student ID], Course.[Course Name], Training.Amount, Training.Time, Training.[Started Date], Training.[Completed Date], Training.Certified FROM Course RIGHT JOIN (Instructor LEFT JOIN Training ON Instructor.[Instructor ID] = Training.[Instructor ID]) ON Course.[Course ID] = Training.[Course ID] where Instructor.[Instructor ID]='" & InstructorList.DataGrid1.Columns(0).Text & "' ORDER BY Training.[Started Date] DESC"
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open strsql, con, adOpenForwardOnly, adLockOptimistic
Set DataGrid1.DataSource = rs
DataGrid1.Columns(0).Caption = "I. ID"
DataGrid1.Columns(0).Width = 400
DataGrid1.Columns(1).Caption = "S. ID"
DataGrid1.Columns(1).Width = 600
DataGrid1.Columns(2).Width = 3000
DataGrid1.Columns(3).Alignment = dbgRight
DataGrid1.Columns(3).Width = 700
DataGrid1.Columns(4).Width = 700
DataGrid1.Columns(5).Width = 1000
DataGrid1.Columns(6).Width = 1200
DataGrid1.Columns(7).Width = 700
DataGrid1.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
'con.Close
End Sub
