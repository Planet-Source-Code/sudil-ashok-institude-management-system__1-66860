VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form CourseOffered 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Course Offered"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7590
   Icon            =   "CourseOffered.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "CourseOffered.frx":1708A
   ScaleHeight     =   4785
   ScaleWidth      =   7590
   Begin MSComCtl2.DTPicker DPic2 
      Height          =   285
      Left            =   6090
      TabIndex        =   21
      Top             =   1560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Format          =   24969217
      CurrentDate     =   37938
   End
   Begin MSComCtl2.DTPicker DPic1 
      Height          =   285
      Left            =   1200
      TabIndex        =   20
      Top             =   1560
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   503
      _Version        =   393216
      Format          =   24969217
      CurrentDate     =   37938
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   480
      Width           =   1335
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   6090
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1920
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1455
      Left            =   120
      TabIndex        =   7
      Top             =   3240
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   2566
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
   Begin VB.ComboBox Combo2 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1200
      TabIndex        =   5
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   6090
      TabIndex        =   4
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6090
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   840
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Top             =   840
      Width           =   3480
   End
   Begin VB.Image Command1 
      Height          =   825
      Left            =   480
      MouseIcon       =   "CourseOffered.frx":2BC4A
      MousePointer    =   99  'Custom
      Picture         =   "CourseOffered.frx":2BF54
      ToolTipText     =   "Savd"
      Top             =   2400
      Width           =   840
   End
   Begin VB.Image Command3 
      Height          =   720
      Left            =   2520
      MouseIcon       =   "CourseOffered.frx":2C9CD
      MousePointer    =   99  'Custom
      Picture         =   "CourseOffered.frx":2CCD7
      ToolTipText     =   "Edit"
      Top             =   2400
      Width           =   720
   End
   Begin VB.Image Command2 
      Height          =   555
      Left            =   6120
      MouseIcon       =   "CourseOffered.frx":2D426
      MousePointer    =   99  'Custom
      Picture         =   "CourseOffered.frx":2D730
      ToolTipText     =   "Exit"
      Top             =   2520
      Width           =   600
   End
   Begin VB.Image Command4 
      Height          =   720
      Left            =   4320
      MouseIcon       =   "CourseOffered.frx":2DEDA
      MousePointer    =   99  'Custom
      Picture         =   "CourseOffered.frx":2E1E4
      ToolTipText     =   "Cancel"
      Top             =   2400
      Width           =   720
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   7440
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   0
      X2              =   7320
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   0
      X2              =   7320
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Training ID:"
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
      Left            =   4920
      TabIndex        =   19
      Top             =   480
      Width           =   1095
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   120
      X2              =   7440
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Certified:"
      Height          =   255
      Left            =   4920
      TabIndex        =   18
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Instructor ID:"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   1920
      Width           =   1005
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Time:/Shift"
      Height          =   255
      Left            =   4920
      TabIndex        =   16
      Top             =   1245
      Width           =   810
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Completed Date:"
      Height          =   255
      Left            =   4920
      TabIndex        =   15
      Top             =   1605
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Amount(Rs.):"
      Height          =   345
      Left            =   120
      TabIndex        =   14
      Top             =   1200
      Width           =   1005
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date:"
      Height          =   345
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   1005
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Course Name:"
      Height          =   345
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   1005
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Course ID:"
      Height          =   345
      Left            =   120
      TabIndex        =   11
      Top             =   480
      Width           =   1005
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
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
      Height          =   225
      Left            =   4920
      TabIndex        =   10
      Top             =   840
      Width           =   1005
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Course Offered"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "CourseOffered"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection
Dim rs As ADODB.Recordset

Private Sub Combo1_Change()
On Error GoTo myerror:

Dim strsql As String
strsql = "Select * from Course where [Course ID] =" & "'" & Combo1.Text & "'"
Set rs = New ADODB.Recordset
rs.Open strsql, con, adOpenForwardOnly, adLockOptimistic

If Not rs.EOF Then
Text3.Text = rs.Fields("Course Name")
rs.MoveNext
End If
rs.Close
Text3.Locked = True

Combo2.Clear
Dim strsql1 As String
strsql1 = "Select * from Percentage where [Course ID] ='" & Combo1.Text & "'"
Set rs1 = New ADODB.Recordset
rs1.Open strsql1, con, adOpenForwardOnly, adLockOptimistic
While Not rs1.EOF
    Combo2.AddItem rs1.Fields("Instructor ID")
    rs1.MoveNext
Wend
rs1.Close
Call datashow
myerror:
Call error

End Sub

Private Sub Combo1_lostfocus()
Call Combo1_Change
End Sub



Private Sub Command1_Click()
On Error GoTo myerror:

If Text1 = "" Or Text3 = "" Or Combo1 = "" Or Text4 = "" Or Text5 = "" Or Combo2 = "" Then
MsgBox "Fill up all selected fields"
Exit Sub
End If
If MsgBox("             Are you sure(Y/N)???", 4) = 7 Then Exit Sub

If Command1.ToolTipText = "Update" Then
Dim AID As String
    Dim strsql1 As String
    strsql1 = "Select * from Training where [Training ID]=" & Text2.Text & ""
    Set rs1 = New ADODB.Recordset
    rs1.Open strsql1, con, adOpenForwardOnly, adLockOptimistic
        rs1.Update("Course ID") = Combo1.Text
        rs1.Update("Amount") = Text4.Text
        rs1.Update("Time") = Text5.Text
        rs1.Update("Started Date") = Dpic1
        rs1.Update("Completed Date") = DPic2
        rs1.Update("Instructor ID") = Combo2.Text
        rs1.Update("Certified") = Combo3.Text
    rs1.Close
Else


    Dim strsql As String
    strsql = "Select * from Training"
    Set rs = New ADODB.Recordset
    rs.Open strsql, con, adOpenForwardOnly, adLockOptimistic
    rs.AddNew
    If Not rs.EOF Then
        rs.Fields("Student ID") = Text1.Text
        rs.Fields("Training ID") = Text2.Text
        rs.Fields("Course ID") = Combo1.Text
        rs.Fields("Amount") = Text4.Text
        rs.Fields("Time") = Text5.Text
        rs.Fields("Started Date") = Dpic1
        rs.Fields("Completed Date") = DPic2
        rs.Fields("Instructor ID") = Combo2.Text
        rs.Update
    End If
    rs.Close
End If
Call Command4_Click
myerror:
Call error

End Sub

Private Sub Command2_Click()
Unload Me
StudentList.Show
End Sub

Private Sub Command3_Click()
Text2.Text = DataGrid1.Columns(0).Text
Combo1.Text = DataGrid1.Columns(2).Text
Text4.Text = DataGrid1.Columns(3).Text
Text5.Text = DataGrid1.Columns(4).Text
Dpic1 = DataGrid1.Columns(5).Text
DPic2 = DataGrid1.Columns(6).Text
Combo2.Text = DataGrid1.Columns(7).Text
Combo3.Text = DataGrid1.Columns(8).Text
Command1.ToolTipText = "Update"

End Sub

Private Sub Command4_Click()
Text2.Text = ""
Combo1.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Combo2.Text = ""
Command1.ToolTipText = "Save"
End Sub

Private Sub Form_Load()
On Error GoTo myerror:

Set con = CreateObject("Adodb.connection")
con.Provider = "Microsoft.jet.oledb.4.0"
con.Open App.Path & "\MainData.mdb"

Text1.Text = StudentList.DataGrid1.Columns(0).Text

Set rs3 = New ADODB.Recordset
rs3.Open "Training", con, adOpenForwardOnly, adLockOptimistic
If rs3.EOF = True Then
    Text2.Text = 300
    GoTo ext:
End If
rs3.MoveLast
Dim strcode As String
strcode = rs3.Fields("Training ID")
Text2.Text = strcode + 1
ext:
rs3.Close

Dim strsql As String
strsql = "Select * from Course"
Set rs = New ADODB.Recordset
rs.Open strsql, con, adOpenForwardOnly, adLockOptimistic

While Not rs.EOF
Combo1.AddItem rs.Fields("Course ID")
rs.MoveNext
Wend
rs.Close
Combo3.AddItem "Yes"
Combo3.AddItem "No"
Dpic1.Value = Date
DPic2.Value = Date
Call datashow

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
'For student
Dim strsql As String
strsql = "SELECT [Student ID], Sum(Amount) AS TAmount From Training GROUP BY [Student ID]"
Set rs = New ADODB.Recordset
rs.Open strsql, con, adOpenForwardOnly, adLockOptimistic
While Not rs.EOF
Dim AID As String
AID = rs.Fields("Student ID")
    Dim strsql1 As String
    strsql1 = "SELECT * From SumOfAmount where [Acc Head ID]='" & AID & "'"
    Set rs1 = New ADODB.Recordset
    rs1.Open strsql1, con, adOpenForwardOnly, adLockOptimistic
    rs1.Update("TotalPayment") = rs.Fields("TAmount")
    rs1.Close
rs.MoveNext
Wend
rs.Close

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
con.Close
StudentList.Show
End Sub

Sub datashow()
Dim strsql As String
strsql = "Select * from Training where [Student ID]='" & Text1.Text & "'"
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open strsql, con, adOpenForwardOnly, adLockOptimistic
Set DataGrid1.DataSource = rs
DataGrid1.Columns(0).Caption = "T.ID"
DataGrid1.Columns(0).Width = 600
DataGrid1.Columns(1).Visible = False
DataGrid1.Columns(2).Width = 800
DataGrid1.Columns(3).Width = 700
DataGrid1.Columns(4).Width = 700
DataGrid1.Columns(5).Width = 1000
DataGrid1.Columns(6).Width = 1200
DataGrid1.Columns(7).Width = 1000
DataGrid1.Columns(8).Width = 700
End Sub

