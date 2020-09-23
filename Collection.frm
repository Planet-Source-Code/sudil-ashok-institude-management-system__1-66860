VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Collect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Collection/Payment"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7515
   Icon            =   "Collection.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "Collection.frx":1708A
   ScaleHeight     =   6585
   ScaleWidth      =   7515
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   5760
      TabIndex        =   2
      Top             =   6180
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   8916
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
   Begin VB.Image Command5 
      Height          =   555
      Left            =   5760
      MouseIcon       =   "Collection.frx":2BC4A
      MousePointer    =   99  'Custom
      Picture         =   "Collection.frx":2BF54
      ToolTipText     =   "Exit"
      Top             =   360
      Width           =   600
   End
   Begin VB.Image Command4 
      Height          =   675
      Left            =   4320
      MouseIcon       =   "Collection.frx":2C6FE
      MousePointer    =   99  'Custom
      Picture         =   "Collection.frx":2CA08
      Stretch         =   -1  'True
      ToolTipText     =   "Expences"
      Top             =   240
      Width           =   705
   End
   Begin VB.Image command1 
      Height          =   675
      Left            =   480
      MouseIcon       =   "Collection.frx":2CF06
      MousePointer    =   99  'Custom
      Picture         =   "Collection.frx":2D210
      Stretch         =   -1  'True
      ToolTipText     =   "Students"
      Top             =   240
      Width           =   705
   End
   Begin VB.Image command3 
      Height          =   675
      Left            =   3000
      MouseIcon       =   "Collection.frx":2D6D8
      MousePointer    =   99  'Custom
      Picture         =   "Collection.frx":2D9E2
      Stretch         =   -1  'True
      ToolTipText     =   "Income"
      Top             =   240
      Width           =   705
   End
   Begin VB.Image command2 
      Height          =   675
      Left            =   1800
      MouseIcon       =   "Collection.frx":2DEDA
      MousePointer    =   99  'Custom
      Picture         =   "Collection.frx":2E1E4
      Stretch         =   -1  'True
      ToolTipText     =   "Instructurs"
      Top             =   240
      Width           =   705
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   75
      Top             =   300
      Width           =   7335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      TabIndex        =   3
      Top             =   6240
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Click on buttons to view Acc Head ID Wise !"
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
      TabIndex        =   1
      Top             =   30
      Width           =   3375
   End
End
Attribute VB_Name = "Collect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection
Private Sub Command1_Click()
Dim strsql As String
strsql = "Select [Cash Transaction].[Transaction ID], [Cash Transaction].Date, [Cash Transaction].[Acc Head ID], [Cash Transaction].Description,[Cash Transaction].[Bill no], [Cash Transaction].Received FROM [Cash Transaction] Where [Cash Transaction].[Acc Head ID] " & " Like '" & "S%" & "'"
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open strsql, con, adOpenForwardOnly, adLockOptimistic
Set DataGrid1.DataSource = rs
DataGrid1.Columns(0).Caption = "T. ID"
DataGrid1.Columns(0).Width = 600
DataGrid1.Columns(1).Width = 1100
DataGrid1.Columns(2).Width = 1000
DataGrid1.Columns(3).Width = 2600
DataGrid1.Columns(4).Width = 510
DataGrid1.Columns(5).Alignment = dbgRight
DataGrid1.Columns(5).Width = 800
DataGrid1.Refresh
Dim totl As Double
While Not rs.EOF
totl = totl + rs.Fields("Received")
rs.MoveNext
Wend
Text1.Text = totl

End Sub



Private Sub Command2_Click()
Dim strsql As String
strsql = "Select [Cash Transaction].[Transaction ID], [Cash Transaction].Date, [Cash Transaction].[Acc Head ID], [Cash Transaction].Description, [Cash Transaction].Paid FROM [Cash Transaction] Where [Cash Transaction].[Acc Head ID] " & " Like '" & "T%" & "'"
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open strsql, con, adOpenForwardOnly, adLockOptimistic
Set DataGrid1.DataSource = rs
DataGrid1.Columns(0).Caption = "T. ID"
DataGrid1.Columns(0).Width = 600
DataGrid1.Columns(1).Width = 1100
DataGrid1.Columns(2).Width = 1000
DataGrid1.Columns(3).Width = 2800
DataGrid1.Columns(4).Alignment = dbgRight
DataGrid1.Columns(4).Width = 1000
DataGrid1.Refresh
Dim totl As Double
While Not rs.EOF
totl = totl + rs.Fields("Paid")
rs.MoveNext
Wend
Text1.Text = totl
End Sub

Private Sub Command3_Click()
Dim strsql As String
strsql = "Select [Cash Transaction].[Transaction ID], [Cash Transaction].Date, [Cash Transaction].[Acc Head ID], [Cash Transaction].Description, [Cash Transaction].Received FROM [Cash Transaction]"
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open strsql, con, adOpenForwardOnly, adLockOptimistic
Set DataGrid1.DataSource = rs
DataGrid1.Columns(0).Caption = "T. ID"
DataGrid1.Columns(0).Width = 600
DataGrid1.Columns(1).Width = 1100
DataGrid1.Columns(2).Width = 1000
DataGrid1.Columns(3).Width = 2800
DataGrid1.Columns(4).Alignment = dbgRight
DataGrid1.Columns(4).Width = 1000
DataGrid1.Refresh
Dim totl As Double
While Not rs.EOF
totl = totl + rs.Fields("Received")
rs.MoveNext
Wend
Text1.Text = totl
End Sub

Private Sub Command4_Click()
Dim strsql As String
strsql = "Select [Cash Transaction].[Transaction ID], [Cash Transaction].Date, [Cash Transaction].[Acc Head ID], [Cash Transaction].Description, [Cash Transaction].Paid FROM [Cash Transaction]"
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open strsql, con, adOpenForwardOnly, adLockOptimistic
Set DataGrid1.DataSource = rs
DataGrid1.Columns(0).Caption = "T. ID"
DataGrid1.Columns(0).Width = 600
DataGrid1.Columns(1).Width = 1100
DataGrid1.Columns(2).Width = 1000
DataGrid1.Columns(3).Width = 2800
DataGrid1.Columns(4).Alignment = dbgRight
DataGrid1.Columns(4).Width = 1000
DataGrid1.Refresh
Dim totl As Double
While Not rs.EOF
totl = totl + rs.Fields("Paid")
rs.MoveNext
Wend
Text1.Text = totl
End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Private Sub Form_Load()
Set con = CreateObject("Adodb.connection")
con.Provider = "Microsoft.jet.oledb.4.0"
con.Open App.Path & "\maindata.mdb"

End Sub

Private Sub Form_Unload(Cancel As Integer)
con.Close
End Sub

Private Sub Image3_Click()

End Sub
