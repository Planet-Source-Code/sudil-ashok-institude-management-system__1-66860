VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Receive 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Receive Payment"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6750
   Icon            =   "Payments.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "Payments.frx":1708A
   ScaleHeight     =   4680
   ScaleWidth      =   6750
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   280
      Left            =   3480
      TabIndex        =   9
      Top             =   360
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   1440
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker DPic1 
      Height          =   300
      Left            =   1440
      TabIndex        =   1
      Top             =   840
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   529
      _Version        =   393216
      Format          =   52756481
      CurrentDate     =   37938
   End
   Begin VB.TextBox Text3 
      Height          =   300
      Left            =   1440
      TabIndex        =   2
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   300
      Left            =   1440
      TabIndex        =   3
      Top             =   1560
      Width           =   5175
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1575
      Left            =   120
      TabIndex        =   8
      Top             =   2880
      Width           =   6570
      _ExtentX        =   11589
      _ExtentY        =   2778
      _Version        =   393216
      BackColor       =   -2147483624
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
      Left            =   840
      TabIndex        =   5
      Top             =   2400
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
      Left            =   4440
      TabIndex        =   6
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Receive Bill No."
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Paid Amount(Rs.):"
      Height          =   300
      Left            =   120
      TabIndex        =   13
      Top             =   1215
      Width           =   1275
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Student ID:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   420
      Width           =   1005
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
      Height          =   300
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   405
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   6480
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Caption         =   "Receive Payment"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   2055
   End
End
Attribute VB_Name = "Receive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection

Private Sub Command1_Click()
On Error GoTo myerror:

If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text5 = "" Then
MsgBox "Fill up all fields"
Exit Sub
End If
Dim strsql As String
Set rs = New ADODB.Recordset

strsql = "Select * from [Cash Transaction]"
rs.Open strsql, con, adOpenForwardOnly, adLockOptimistic
If MsgBox("           Are you sure(Y/N)???", 4) = 7 Then Exit Sub

Dim strsql1 As String
strsql1 = "Select * from Student"
Set rs1 = New ADODB.Recordset
rs1.Open strsql1, con, adOpenForwardOnly, adLockOptimistic

While Not rs1.EOF
If rs1.Fields("Student ID") = Text1.Text Then notfound = 1
rs1.MoveNext
Wend

If notfound <> 1 Then MsgBox ("Student ID is not found in Student List!!"): Exit Sub

rs.AddNew
If Not rs.EOF Then
    rs.Fields("Date") = Dpic1
    rs.Fields("Acc Head ID") = Text1.Text
    rs.Fields("Description") = Text2.Text
    rs.Fields("Received") = Text3.Text
    rs.Fields("Bill No") = Text5.Text
    rs.Update
End If

Text2.Text = Clear
Text3.Text = Clear
Text4.Text = Clear
Text5.Text = Clear
rs.Close
rs1.Close

Call datashow
myerror:
Call error

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo myerror:
Set con = CreateObject("Adodb.connection")
con.Provider = "Microsoft.jet.oledb.4.0"
con.Open App.Path & "\MainData.mdb"

Dim strsql As String
strsql = "Select * from Student"
If rams1 = 2 Then
    Text1.Text = StudentList.DataGrid1.Columns(0).Text
    Text4.Text = StudentList.DataGrid1.Columns(1).Text & " " & StudentList.DataGrid1.Columns(2).Text
    strsql = "Select * from Student where [student id]='" & Text1.Text & "'"
    Text1.Enabled = False
End If
Set rs = New ADODB.Recordset
rs.Open strsql, con, adOpenForwardOnly, adLockOptimistic
Dpic1.Value = Date
Text4.Enabled = False

myerror:
Call error

End Sub
Sub datashow()
On Error GoTo myerror:

Dim strsql As String
strsql = "Select * from [Cash Transaction] where [Acc Head ID]='" & Text1.Text & "'"
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open strsql, con, adOpenForwardOnly, adLockOptimistic
Set DataGrid1.DataSource = rs
DataGrid1.Columns(0).Caption = "T. ID"
DataGrid1.Columns(0).Width = 700
DataGrid1.Columns(1).Width = 1000
DataGrid1.Columns(2).Caption = "A. Head ID"
DataGrid1.Columns(2).Width = 900
DataGrid1.Columns(3).Width = 2200
DataGrid1.Columns(4).Width = 550
DataGrid1.Columns(5).Caption = "Amount"
DataGrid1.Columns(5).Width = 630
DataGrid1.Columns(6).Visible = False
DataGrid1.Columns(7).Visible = False
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

Dim strsql As String
strsql = "SELECT [Acc Head ID], Sum(Received) AS TReceived, Sum(Paid) AS TPaid From [Cash Transaction] GROUP BY [Acc Head ID]"
Set rs = New ADODB.Recordset
rs.Open strsql, con, adOpenForwardOnly, adLockOptimistic
While Not rs.EOF
Dim AID As String
AID = rs.Fields("Acc Head ID")
    
    Dim strsql1 As String
    strsql1 = "SELECT * from SumOfAmount where [Acc Head ID]='" & AID & "'"
    Set rs1 = New ADODB.Recordset
    rs1.Open strsql1, con, adOpenForwardOnly, adLockOptimistic
    If Left$(AID, 1) = "S" Then
        rs1.Update("TotalPaid") = rs.Fields("TReceived")
    ElseIf Left$(AID, 1) = "T" Then
          rs1.Update("TotalPaid") = rs.Fields("TPaid")
    Else
       rs1.Update("TotalPayment") = rs.Fields("TReceived")
       rs1.Update("TotalPaid") = rs.Fields("TPaid")
    End If
    rs1.Close
    rs.MoveNext
Wend
rs.Close
con.Close
End Sub

Private Sub Text1_Change()
Call datashow
End Sub

Private Sub Text1_DblClick()
IDList.Command2.Enabled = False
IDList.Command3.Enabled = False
IDList.Show
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
IDList.Command2.Enabled = False
IDList.Command3.Enabled = False
IDList.Show
End If
End Sub

Private Sub Text3_GotFocus()
If rams1 = 2 Then Text3.Text = StudentList.DataGrid1.Columns(6).Text
End Sub
