VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form List 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Certificate List"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11775
   Icon            =   "List.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "List.frx":1708A
   ScaleHeight     =   6945
   ScaleWidth      =   11775
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "Selected Only"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7200
      TabIndex        =   10
      Top             =   195
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5655
      Left            =   30
      TabIndex        =   9
      Top             =   1290
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   9975
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text8 
      Height          =   300
      Left            =   10725
      TabIndex        =   7
      Top             =   975
      Width           =   735
   End
   Begin VB.TextBox Text7 
      Height          =   300
      Left            =   8250
      TabIndex        =   6
      Top             =   975
      Width           =   2460
   End
   Begin VB.TextBox Text6 
      Height          =   300
      Left            =   7035
      TabIndex        =   5
      Top             =   975
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   300
      Left            =   6015
      TabIndex        =   4
      Top             =   975
      Width           =   1020
   End
   Begin VB.TextBox Text4 
      Height          =   300
      Left            =   4080
      TabIndex        =   3
      Top             =   975
      Width           =   1965
   End
   Begin VB.TextBox Text3 
      Height          =   300
      Left            =   1575
      TabIndex        =   2
      Top             =   975
      Width           =   2475
   End
   Begin VB.TextBox Text2 
      Height          =   300
      Left            =   855
      TabIndex        =   1
      Top             =   975
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   255
      TabIndex        =   0
      Top             =   975
      Width           =   615
   End
   Begin VB.Image Command7 
      Height          =   555
      Left            =   10800
      MouseIcon       =   "List.frx":2BC4A
      MousePointer    =   99  'Custom
      Picture         =   "List.frx":2BF54
      ToolTipText     =   "Close"
      Top             =   240
      Width           =   600
   End
   Begin VB.Image Command1 
      Height          =   720
      Left            =   9240
      MouseIcon       =   "List.frx":2C6FE
      MousePointer    =   99  'Custom
      Picture         =   "List.frx":2CA08
      ToolTipText     =   "Print"
      Top             =   120
      Width           =   720
   End
   Begin VB.Image Command2 
      Height          =   720
      Left            =   3120
      MouseIcon       =   "List.frx":2D2C5
      MousePointer    =   99  'Custom
      Picture         =   "List.frx":2D5CF
      ToolTipText     =   "Search"
      Top             =   0
      Width           =   720
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   720
      Left            =   6990
      Top             =   75
      Width           =   3045
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fill in any fields and click on Search !"
      Height          =   255
      Left            =   270
      TabIndex        =   8
      Top             =   240
      Width           =   2700
   End
End
Attribute VB_Name = "List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection
Dim selec As Integer

Private Sub Command1_Click()
If Check1.Value = 1 Then
    If DataGrid1.Columns(4).Text = DataGrid1.Columns(5).Text Then
        MsgBox ("Completion date not valid!")
    Exit Sub
    End If
    DataEnvironment1.Connection1.ConnectionString = App.Path & "\MainData.mdb"
    DataEnvironment1.Connection1.Open
    DataEnvironment1.Command3_Grouping (DataGrid1.Columns(0).Text)
    DataReport3.Show
Else
    DataEnvironment1.Connection1.ConnectionString = App.Path & "\MainData.mdb"
    DataEnvironment1.Connection1.Open
    DataEnvironment1.StudentID_Grouping
    DataReport2.Show
End If
End Sub

Private Sub Command2_Click()
On Error GoTo myerror:
Dim strsql As String
If Not Text1.Text = "" Then
strsql = "SELECT Training.[Training ID], Student.[Student ID], Student.[First Name], Student.[Last Name],  Training.[Started Date], Training.[Completed Date], Course.[Course Name],  Training.Certified FROM Student INNER JOIN (Course INNER JOIN Training ON Course.[Course ID] = Training.[Course ID]) ON Student.[Student ID] = Training.[Student ID] where training.[Training ID]=" & Text1.Text & ""
ElseIf Not Text2.Text = "" Then
    strsql = "SELECT Training.[Training ID], Student.[Student ID], Student.[First Name], Student.[Last Name],  Training.[Started Date], Training.[Completed Date], Course.[Course Name],  Training.Certified FROM Student INNER JOIN (Course INNER JOIN Training ON Course.[Course ID] = Training.[Course ID]) ON Student.[Student ID] = Training.[Student ID] WHERE Student.[Student ID]" & " Like '" & Text2.Text & "%" & "'"
ElseIf Not Text3.Text = "" Then
    strsql = "SELECT Training.[Training ID], Student.[Student ID], Student.[First Name], Student.[Last Name],  Training.[Started Date], Training.[Completed Date], Course.[Course Name],  Training.Certified FROM Student INNER JOIN (Course INNER JOIN Training ON Course.[Course ID] = Training.[Course ID]) ON Student.[Student ID] = Training.[Student ID] WHERE Student.[First Name]" & " Like '" & Text3.Text & "%" & "'"
ElseIf Not Text4.Text = "" Then
    strsql = "SELECT Training.[Training ID], Student.[Student ID], Student.[First Name], Student.[Last Name],  Training.[Started Date], Training.[Completed Date], Course.[Course Name],  Training.Certified FROM Student INNER JOIN (Course INNER JOIN Training ON Course.[Course ID] = Training.[Course ID]) ON Student.[Student ID] = Training.[Student ID] WHERE Student.[Last Name]" & " Like '" & Text4.Text & "%" & "'"
ElseIf Not Text5.Text = "" Then
    strsql = "SELECT Training.[Training ID], Student.[Student ID], Student.[First Name], Student.[Last Name],  Training.[Started Date], Training.[Completed Date], Course.[Course Name],  Training.Certified FROM Student INNER JOIN (Course INNER JOIN Training ON Course.[Course ID] = Training.[Course ID]) ON Student.[Student ID] = Training.[Student ID] WHERE Training.[Started Date]" & " Like '" & Text5.Text & "%" & "'"
ElseIf Not Text6.Text = "" Then
    strsql = "SELECT Training.[Training ID], Student.[Student ID], Student.[First Name], Student.[Last Name],  Training.[Started Date], Training.[Completed Date], Course.[Course Name],  Training.Certified FROM Student INNER JOIN (Course INNER JOIN Training ON Course.[Course ID] = Training.[Course ID]) ON Student.[Student ID] = Training.[Student ID] WHERE Training.[Completed Date]" & " Like '" & Text6.Text & "%" & "'"
ElseIf Not Text7.Text = "" Then
    strsql = "SELECT Training.[Training ID], Student.[Student ID], Student.[First Name], Student.[Last Name],  Training.[Started Date], Training.[Completed Date], Course.[Course Name],  Training.Certified FROM Student INNER JOIN (Course INNER JOIN Training ON Course.[Course ID] = Training.[Course ID]) ON Student.[Student ID] = Training.[Student ID] WHERE Course.[Course Name]" & " Like '" & Text7.Text & "%" & "'"
ElseIf Not Text8.Text = "" Then
    strsql = "SELECT Training.[Training ID], Student.[Student ID], Student.[First Name], Student.[Last Name],  Training.[Started Date], Training.[Completed Date], Course.[Course Name],  Training.Certified FROM Student INNER JOIN (Course INNER JOIN Training ON Course.[Course ID] = Training.[Course ID]) ON Student.[Student ID] = Training.[Student ID] WHERE Training.Certified" & " Like '" & Text8.Text & "%" & "'"
Else
    strsql = "SELECT Training.[Training ID], Student.[Student ID], Student.[First Name], Student.[Last Name],  Training.[Started Date], Training.[Completed Date], Course.[Course Name],  Training.Certified FROM Student INNER JOIN (Course INNER JOIN Training ON Course.[Course ID] = Training.[Course ID]) ON Student.[Student ID] = Training.[Student ID] order by Training.[Training ID] desc"
End If

Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open strsql, con, adOpenForwardOnly, adLockOptimistic
Set DataGrid1.DataSource = rs
DataGrid1.Columns(0).Caption = "T.ID"
DataGrid1.Columns(0).Width = 500
DataGrid1.Columns(1).Caption = "S.ID"
DataGrid1.Columns(1).Width = 700
DataGrid1.Columns(2).Width = 2500
DataGrid1.Columns(3).Width = 2000
DataGrid1.Columns(4).Width = 1000
DataGrid1.Columns(5).Width = 1200
DataGrid1.Columns(6).Width = 2500
DataGrid1.Columns(7).Width = 700
DataGrid1.Refresh
'rs.Close
myerror:
Call error
End Sub

Private Sub Command7_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Dim strsql As String
strsql = "SELECT Training.[Training ID], Student.[Student ID], Student.[First Name], Student.[Last Name],  Training.[Started Date], Training.[Completed Date], Course.[Course Name],  Training.Certified FROM Student INNER JOIN (Course INNER JOIN Training ON Course.[Course ID] = Training.[Course ID]) ON Student.[Student ID] = Training.[Student ID] order by Training.[Training ID] desc"
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open strsql, con, adOpenForwardOnly, adLockOptimistic

Set DataGrid1.DataSource = rs
DataGrid1.Columns(0).Caption = "T.ID"
DataGrid1.Columns(0).Width = 500
DataGrid1.Columns(1).Caption = "S.ID"
DataGrid1.Columns(1).Width = 700
DataGrid1.Columns(2).Width = 2500
DataGrid1.Columns(3).Width = 2000
DataGrid1.Columns(4).Width = 1000
DataGrid1.Columns(5).Width = 1200
DataGrid1.Columns(6).Width = 2500
DataGrid1.Columns(7).Width = 700
DataGrid1.Refresh

End Sub

Private Sub Form_Load()
On Error GoTo myerror:

Set con = CreateObject("Adodb.connection")
con.Provider = "Microsoft.jet.oledb.4.0"
con.Open App.Path & "\maindata.mdb"
Dim strsql As String
strsql = "SELECT Training.[Training ID], Student.[Student ID], Student.[First Name], Student.[Last Name],  Training.[Started Date], Training.[Completed Date], Course.[Course Name],  Training.Certified FROM Student INNER JOIN (Course INNER JOIN Training ON Course.[Course ID] = Training.[Course ID]) ON Student.[Student ID] = Training.[Student ID] order by Training.[Training ID] desc"
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open strsql, con, adOpenForwardOnly, adLockOptimistic

Set DataGrid1.DataSource = rs
DataGrid1.Columns(0).Caption = "T.ID"
DataGrid1.Columns(0).Width = 500
DataGrid1.Columns(1).Caption = "S.ID"
DataGrid1.Columns(1).Width = 700
DataGrid1.Columns(2).Width = 2500
DataGrid1.Columns(3).Width = 2000
DataGrid1.Columns(4).Width = 1000
DataGrid1.Columns(5).Width = 1200
DataGrid1.Columns(6).Width = 2500
DataGrid1.Columns(7).Width = 700
DataGrid1.Refresh
'rs.Close
selec = 0
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
