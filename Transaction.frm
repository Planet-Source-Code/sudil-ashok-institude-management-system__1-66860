VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Transaction 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cash Transaction"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9120
   FillStyle       =   0  'Solid
   ForeColor       =   &H00000000&
   Icon            =   "Transaction.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "Transaction.frx":1708A
   ScaleHeight     =   6840
   ScaleWidth      =   9120
   Begin VB.CommandButton Command14 
      Height          =   195
      Left            =   7935
      TabIndex        =   27
      Top             =   2445
      Width           =   795
   End
   Begin VB.CommandButton Command13 
      Height          =   195
      Left            =   7155
      TabIndex        =   26
      Top             =   2445
      Width           =   795
   End
   Begin VB.CommandButton Command12 
      Height          =   195
      Left            =   6345
      TabIndex        =   25
      Top             =   2445
      Width           =   825
   End
   Begin VB.CommandButton Command11 
      Height          =   200
      Left            =   5730
      TabIndex        =   24
      Top             =   2450
      Width           =   630
   End
   Begin VB.CommandButton Command10 
      Height          =   200
      Left            =   3255
      TabIndex        =   23
      Top             =   2450
      Width           =   2490
   End
   Begin VB.CommandButton Command9 
      Height          =   200
      Left            =   2160
      TabIndex        =   22
      Top             =   2450
      Width           =   1110
   End
   Begin VB.CommandButton Command8 
      Height          =   200
      Left            =   1155
      TabIndex        =   21
      Top             =   2450
      Width           =   1020
   End
   Begin VB.CommandButton Command4 
      Height          =   200
      Left            =   435
      TabIndex        =   20
      Top             =   2450
      Width           =   735
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   6840
      TabIndex        =   6
      Top             =   975
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   280
      Left            =   4080
      TabIndex        =   18
      Top             =   600
      Width           =   4695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1920
      TabIndex        =   4
      Top             =   600
      Width           =   2055
   End
   Begin MSComCtl2.DTPicker Dpic1 
      Height          =   315
      Left            =   6825
      TabIndex        =   3
      Top             =   210
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   556
      _Version        =   393216
      CalendarTitleBackColor=   -2147483638
      Format          =   52756481
      CurrentDate     =   37938
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3735
      Left            =   120
      TabIndex        =   12
      Top             =   2640
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   6588
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
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
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Print Receipt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   2880
      TabIndex        =   11
      Top             =   1920
      Width           =   1200
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "Text5"
      Top             =   6450
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Top             =   975
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
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
      Height          =   350
      Left            =   7440
      TabIndex        =   9
      Top             =   1920
      Width           =   1200
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   5160
      TabIndex        =   10
      Top             =   1920
      Width           =   1200
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
      Height          =   350
      Left            =   360
      TabIndex        =   8
      Top             =   1920
      Width           =   1200
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1920
      TabIndex        =   7
      Top             =   1350
      Width           =   6855
   End
   Begin VB.CommandButton Command6 
      Appearance      =   0  'Flat
      Caption         =   "Paid >>"
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
      Left            =   1440
      MaskColor       =   &H00008000&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Appearance      =   0  'Flat
      Caption         =   "Received >>"
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
      Left            =   120
      MaskColor       =   &H008080FF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C000&
      Caption         =   "Bill No :"
      Height          =   255
      Left            =   6240
      TabIndex        =   19
      Top             =   1035
      Width           =   615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      X1              =   0
      X2              =   9000
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label Label7 
      Caption         =   "Balance:"
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
      Left            =   6960
      TabIndex        =   16
      Top             =   6480
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "Amount:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   480
      TabIndex        =   15
      Top             =   1020
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "Discription:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   480
      TabIndex        =   14
      Top             =   1395
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "Account ID:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   480
      TabIndex        =   13
      Top             =   615
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "Date:"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   5385
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      FillColor       =   &H00FFFFFF&
      Height          =   1725
      Left            =   120
      Top             =   120
      Width           =   8895
   End
End
Attribute VB_Name = "Transaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection
Dim headselect As Integer

Private Sub Command1_Click()
If Label2.Caption = "Account ID:" Or Label4.Caption = "Amount:" Then
MsgBox ("Select the entry type 'Received' / 'Paid'!!!")
Exit Sub
End If
If Text2 = "" Or Text2 = "0" Or Text1 = "" Or Combo2 = "" Then
MsgBox ("Fill all the information!!!")
Exit Sub
End If
Dim strsql As String
If Command1.Caption = "Update" Then
strsql = "Select * From [Cash Transaction] where [Transaction ID]=" & DataGrid1.Columns(0).Text & ""
Set rs = New ADODB.Recordset
rs.Open strsql, con, adOpenForwardOnly, adLockOptimistic
    rs.Update("Date") = Dpic1
    rs.Update("Acc Head ID") = Text1.Text
    rs.Update("Description") = Combo2.Text
    rs.Fields("Bill No") = Text4.Text
    If Label4.Caption = "Received Amount:" Then
        rs.Update("Received") = Text2.Text
    ElseIf Label4.Caption = "Paid Amount:" Then
        rs.Update("Paid") = Text2.Text
    End If
Else
    strsql = "Select * From [Cash Transaction]"
    Set rs = New ADODB.Recordset
    rs.Open strsql, con, adOpenForwardOnly, adLockOptimistic
    rs.AddNew
    rs.Fields("Date") = Dpic1
    rs.Fields("Acc Head ID") = Text1.Text
    rs.Fields("Description") = Combo2.Text
    rs.Fields("Bill No") = Text4.Text
    If Label4.Caption = "Received Amount:" Then
        rs.Update("Received") = Text2.Text
    ElseIf Label4.Caption = "Paid Amount:" Then
        rs.Update("Paid") = Text2.Text
    End If
rs.Update
End If
rs.Close
Call Command2_Click
Call datashow

'call print_receipt_click
End Sub

Private Sub Command10_Click()
headselect = 4
Call datashow
End Sub

Private Sub Command11_Click()
headselect = 5
Call datashow
End Sub

Private Sub Command2_Click()
Dpic1.Enabled = True
Text2.Enabled = True
Text1.Enabled = True
Text4.Enabled = True
Combo2.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Command7.Enabled = True
Label2.Caption = "Account ID:"
Label4.Caption = "Amount:"
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Combo2.Text = ""
Text1.Text = ""
Command5.FontBold = False
Command6.FontBold = False
Command5.BackColor = &H8000000F
Command6.BackColor = &H8000000F
Command1.Caption = "Save"
End Sub

Private Sub Command3_Click()
Unload Me
End Sub


Private Sub Command4_Click()
headselect = 1
Call datashow
End Sub

Private Sub Command5_Click()
Call Command2_Click
Label2.Caption = "From Account ID:"
Label4.Caption = "Received Amount:"
Command5.FontBold = True
Command5.BackColor = &HC0FFC0
Command6.FontBold = False
Command6.BackColor = &H8000000F
Command1.Caption = "Save"
Text1.Text = ""
End Sub

Private Sub Command6_Click()
Call Command2_Click
Label2.Caption = "To Account ID:"
Label4.Caption = "Paid Amount:"
Command5.FontBold = False
Command5.BackColor = &H8000000F
Command6.FontBold = True
Command6.BackColor = &HC0FFC0
Command1.Caption = "Save"
Text1.Text = ""
End Sub

Private Sub Command8_Click()
headselect = 2
Call datashow
End Sub

Private Sub Command9_Click()
headselect = 3
Call datashow
End Sub

Private Sub DataGrid1_Click()
Call Command2_Click
Dim strsql As String
strsql = "Select * From [Cash Transaction] where [Transaction ID]=" & DataGrid1.Columns(0).Text & ""
Set rs = New ADODB.Recordset
rs.Open strsql, con, adOpenForwardOnly, adLockOptimistic
Dpic1 = rs.Fields("Date")
Dim strid As String
strid = rs.Fields("Acc Head ID")
If rs.Fields("Received") <> 0 Then
        Call Command5_Click
        Text1.Text = rs.Fields("Acc Head ID")
        Text2.Text = rs.Fields("Received")
        Combo2.Text = rs.Fields("Description")
        Text4.Text = rs.Fields("Bill No")
    ElseIf rs.Fields("Paid") <> 0 Then
        Call Command6_Click
        Text1.Text = rs.Fields("Acc Head ID")
        Text2.Text = rs.Fields("Paid")
        Combo2.Text = rs.Fields("Description")
        Text4.Text = rs.Fields("Bill No")

End If
Command1.Caption = "Update"
rs.Close
End Sub


Private Sub Form_Load()
Set con = CreateObject("Adodb.connection")
con.Provider = "Microsoft.jet.oledb.4.0"
con.Open App.Path & "\MainData.mdb"
Call datashow
Dpic1.Value = Date
Dpic1.Enabled = False
Text2.Enabled = False
Text1.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Combo2.Enabled = False
Command1.Enabled = False
Command2.Enabled = False
selection = 0
Command7.Enabled = False
Call datashow

End Sub

Sub datashow()

Dim strsqls1 As String
strsqls1 = "Select * From [Cash Transaction] order by [Transaction ID]"
Set rss1 = New ADODB.Recordset
rss1.CursorLocation = adUseClient
rss1.Open strsqls1, con, adOpenForwardOnly, adLockOptimistic

Dim balanc As Double
balanc = 0
While Not rss1.EOF
balanc = balanc + rss1.Fields("Received") - rss1.Fields("Paid")
rss1.Update("Balance") = balanc
rss1.MoveNext
Wend
rss1.Close

Dim strsqls As String
strsqls = "Select * From [Cash Transaction] order by [Transaction ID]DESC"
If headselect = 2 Then
strsqls = "Select * From [Cash Transaction] order by [Date]DESC"
ElseIf headselect = 3 Then
strsqls = "Select * From [Cash Transaction] order by [Acc Head ID]"
ElseIf headselect = 4 Then
strsqls = "Select * From [Cash Transaction] order by [Description]"
ElseIf headselect = 5 Then
strsqls = "Select * From [Cash Transaction] order by [Bill No]DESC"
End If

Set rss = New ADODB.Recordset
rss.CursorLocation = adUseClient
rss.Open strsqls, con, adOpenForwardOnly, adLockOptimistic

If rss.EOF = False Then rss.MoveFirst
Set DataGrid1.DataSource = rss
DataGrid1.Columns(0).Caption = "T. ID"
DataGrid1.Columns(0).Width = 700
DataGrid1.Columns(1).Width = 1000
DataGrid1.Columns(2).Width = 1100
DataGrid1.Columns(3).Width = 2500
DataGrid1.Columns(4).Width = 600
DataGrid1.Columns(5).Alignment = dbgRight
DataGrid1.Columns(5).Width = 800
DataGrid1.Columns(6).Alignment = dbgRight
DataGrid1.Columns(6).Width = 800
DataGrid1.Columns(7).Alignment = dbgRight
DataGrid1.Columns(7).Width = 800
DataGrid1.Refresh
Text5.Text = balanc
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call sumcalc
    con.Close
End Sub

Sub sumcalc()
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
End Sub

Private Sub Text1_DblClick()
IDList.Show
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then IDList.Show
End Sub
