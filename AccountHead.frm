VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form AccountHead 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Account Head Details"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5730
   Icon            =   "AccountHead.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "AccountHead.frx":1708A
   ScaleHeight     =   5220
   ScaleWidth      =   5730
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Top             =   840
      Width           =   3375
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3135
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   5530
      _Version        =   393216
      AllowUpdate     =   0   'False
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
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   480
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
   Begin VB.Image Command3 
      Height          =   555
      Left            =   4200
      MouseIcon       =   "AccountHead.frx":2BC4A
      MousePointer    =   99  'Custom
      Picture         =   "AccountHead.frx":2BF54
      ToolTipText     =   "Exit"
      Top             =   1320
      Width           =   600
   End
   Begin VB.Image Command2 
      Height          =   810
      Left            =   3000
      MouseIcon       =   "AccountHead.frx":2C6FE
      MousePointer    =   99  'Custom
      Picture         =   "AccountHead.frx":2CA08
      ToolTipText     =   "Edit"
      Top             =   1200
      Width           =   825
   End
   Begin VB.Image Command1 
      Height          =   825
      Left            =   1920
      MouseIcon       =   "AccountHead.frx":2D2EE
      MousePointer    =   99  'Custom
      Picture         =   "AccountHead.frx":2D5F8
      ToolTipText     =   "Save"
      Top             =   1200
      Width           =   840
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Type:"
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Head Name:"
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "AC Head ID:"
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "AccountHead"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection

Private Sub Command1_Click()
If Text1.Text = "" Or Text2.Text = "" Then
MsgBox ("Fill up all fields")
Exit Sub
End If
If Command1.ToolTipText = "Update" Then
    Dim strsql As String
    strsql = "Select * from [Account head] where [Head Id]='" & Text1.Text & "'"
    Set rs = New Recordset
    rs.Open strsql, con, adOpenForwardOnly, adLockOptimistic
    rs.Update("Head") = Text2.Text
    rs.Update("Type") = Text3.Text
    rs.Close
Else
    Dim strsql1 As String
    strsql1 = "Select * from [Account head]"
    Set rs = New Recordset
    rs.Open strsql1, con, adOpenForwardOnly, adLockOptimistic
    rs.AddNew
    rs.Fields("Head ID") = Text1.Text
    rs.Fields("Head") = Text2.Text
    rs.Fields("Type") = Text3.Text
    rs.Update
    rs.Close

    Set rs1 = New ADODB.Recordset
    rs1.Open "select * from SumOfAmount", con, adOpenForwardOnly, adLockOptimistic
    rs1.AddNew
    rs1.Fields("Acc Head ID") = Text1.Text
    rs1.Update
    rs1.Close
End If
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Call datashow
Call iddisplay
Command1.ToolTipText = "Save"
End Sub

Private Sub Command2_Click()
Text1.Text = DataGrid1.Columns(1).Text
Text2.Text = DataGrid1.Columns(2).Text
Text3.Text = DataGrid1.Columns(3).Text
Command1.ToolTipText = "Update"
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
Set con = CreateObject("Adodb.connection")
con.Provider = "Microsoft.jet.oledb.4.0"
con.Open App.Path & "\MainData.mdb"
Text1.Enabled = False
Call datashow
Call iddisplay
End Sub

Sub datashow()
Dim strsql As String
strsql = "Select * from [Account Head]"
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open strsql, con, adOpenForwardOnly, adLockOptimistic
Set DataGrid1.DataSource = rs
DataGrid1.Columns(0).Width = 500
DataGrid1.Columns(1).Width = 700
DataGrid1.Columns(2).Width = 2500
DataGrid1.Columns(3).Width = 1200
End Sub

Private Sub Form_Unload(Cancel As Integer)
con.Close
End Sub
Sub iddisplay()
Set rs3 = New ADODB.Recordset
rs3.Open "select * from [Account Head] order by [Sno]", con, adOpenForwardOnly, adLockOptimistic
If rs3.EOF = True Then
    Text1.Text = "AC" & 1
    GoTo ext:
End If
rs3.MoveLast
Dim strcode As String
strcode = rs3.Fields("Head ID")
Text1.Text = "AC" & Val(Mid$(strcode, 3, Len(strcode) - 1)) + 1
ext:
rs3.Close
End Sub

Private Sub Image1_Click()

End Sub

Private Sub Image3_Click()

End Sub
