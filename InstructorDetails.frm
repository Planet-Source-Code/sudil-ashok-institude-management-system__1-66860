VERSION 5.00
Begin VB.Form InstructorDetails 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Instructor Details"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4830
   Icon            =   "InstructorDetails.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "InstructorDetails.frx":1708A
   ScaleHeight     =   3180
   ScaleWidth      =   4830
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   315
      Left            =   1080
      TabIndex        =   2
      Top             =   1200
      Width           =   3600
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   1080
      TabIndex        =   1
      Top             =   840
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   302
      Left            =   1080
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
   Begin VB.Image Command2 
      Height          =   555
      Left            =   3600
      MouseIcon       =   "InstructorDetails.frx":2BC4A
      MousePointer    =   99  'Custom
      Picture         =   "InstructorDetails.frx":2BF54
      ToolTipText     =   "Exit"
      Top             =   2280
      Width           =   600
   End
   Begin VB.Image Command3 
      Height          =   720
      Left            =   2040
      MouseIcon       =   "InstructorDetails.frx":2C6FE
      MousePointer    =   99  'Custom
      Picture         =   "InstructorDetails.frx":2CA08
      ToolTipText     =   "Clear"
      Top             =   2160
      Width           =   720
   End
   Begin VB.Image Command1 
      Height          =   825
      Left            =   360
      MouseIcon       =   "InstructorDetails.frx":2D1E9
      MousePointer    =   99  'Custom
      Picture         =   "InstructorDetails.frx":2D4F3
      ToolTipText     =   "Save"
      Top             =   2160
      Width           =   840
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   4800
      Y1              =   300
      Y2              =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Instructor Entry"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Phone:"
      Height          =   315
      Left            =   360
      TabIndex        =   7
      Top             =   1575
      Width           =   615
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      Height          =   315
      Left            =   240
      TabIndex        =   6
      Top             =   1230
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      Height          =   315
      Left            =   360
      TabIndex        =   5
      Top             =   885
      Width           =   615
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Instructor ID:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   510
      Width           =   930
   End
End
Attribute VB_Name = "InstructorDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection

Private Sub Command1_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Then
MsgBox "Fill up all fields"
Exit Sub
End If
If MsgBox("           Are you sure(Y/N)???", 4) = 7 Then Exit Sub

If rams1 = 3 Then
    Dim strsql As String
    strsql = "Select * From Instructor where [instructor id] ='" & Text1.Text & "'"
    Set rs = New ADODB.Recordset
    rs.Open strsql, con, adOpenForwardOnly, adLockOptimistic
    rs.Update("Name") = UCase$(Text2.Text)
    rs.Update("Address") = UCase$(Text3.Text)
    rs.Update("Phone") = Text4.Text
Else
    Dim strsql1 As String
    strsql1 = "Select * From Instructor"
    Set rs = New ADODB.Recordset
    rs.Open strsql1, con, adOpenForwardOnly, adLockOptimistic
    rs.AddNew
    If Not rs.EOF Then
        rs.Fields("Instructor ID") = Text1.Text
        rs.Fields("Name") = UCase$(Text2.Text)
        rs.Fields("Address") = UCase$(Text3.Text)
        rs.Fields("Phone") = Text4.Text
        rs.Update
    End If
    Set rs1 = New ADODB.Recordset
    rs1.Open "select * from SumOfAmount", con, adOpenForwardOnly, adLockOptimistic
    rs1.AddNew
    rs1.Fields("Acc Head ID") = Text1.Text
    rs1.Update
    rs1.Close
End If
rs.Close
Call Command3_Click
Call iddisplay
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
End Sub

Private Sub Form_Load()
Set con = CreateObject("Adodb.connection")
con.Provider = "Microsoft.jet.oledb.4.0"
con.Open App.Path & "\MainData.mdb"
Text1.Enabled = False
If rams1 = 3 Then
    Text1.Text = InstructorList.DataGrid1.Columns(0).Text
    Dim strsql As String
    strsql = "Select * from instructor where [instructor id]='" & Text1.Text & "'"
    Set rs4 = New ADODB.Recordset
    rs4.Open strsql, con, adopenforwardonlu, adLockOptimistic
    Text2.Text = rs4.Fields("Name")
    Text3.Text = rs4.Fields("Address")
    Text4.Text = rs4.Fields("Phone")
    Command1.ToolTipText = "Update"
ElseIf rams1 = 4 Then
    Call iddisplay
End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
con.Close
If rams1 = 3 Then InstructorList.Show
End Sub
Sub iddisplay()
Set rs3 = New ADODB.Recordset
    rs3.Open "Instructor", con, adOpenForwardOnly, adLockOptimistic
    If rs3.EOF = True Then
        Text1.Text = "T" & 1
        GoTo ext:
    End If
    rs3.MoveLast
    Dim strcode As String
    strcode = rs3.Fields("Instructor ID")
    Text1.Text = "T" & Val(Mid$(strcode, 2, Len(strcode) - 1)) + 1
ext:
    rs3.Close
End Sub
