VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form StudentEdit 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Student Edit"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5775
   Icon            =   "StudentEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "StudentEdit.frx":1708A
   ScaleHeight     =   5280
   ScaleWidth      =   5775
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
      Height          =   375
      Left            =   3720
      TabIndex        =   10
      Top             =   4800
      Width           =   1080
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Update"
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
      TabIndex        =   9
      Top             =   4800
      Width           =   1080
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Student Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   5535
      Begin MSComCtl2.DTPicker Dpic1 
         Height          =   315
         Left            =   1770
         TabIndex        =   22
         Top             =   4035
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   52756481
         CurrentDate     =   37938
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   1785
         TabIndex        =   3
         Top             =   1455
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   1800
         TabIndex        =   2
         Top             =   1050
         Width           =   3600
      End
      Begin VB.TextBox Text5 
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
         Left            =   1800
         TabIndex        =   6
         Top             =   2745
         Width           =   1440
      End
      Begin VB.TextBox Text7 
         Height          =   300
         Left            =   1785
         TabIndex        =   8
         Top             =   3615
         Width           =   3615
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1785
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   1890
         Width           =   3615
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1785
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   2310
         Width           =   3615
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   1785
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   3180
         Width           =   3615
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   1800
         TabIndex        =   1
         Top             =   630
         Width           =   3615
      End
      Begin VB.TextBox Text9 
         Height          =   302
         Left            =   1800
         TabIndex        =   0
         Top             =   225
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Gender:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   570
         TabIndex        =   21
         Top             =   1500
         Width           =   1095
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Join:"
         Height          =   315
         Left            =   240
         TabIndex        =   20
         Top             =   4065
         Width           =   1455
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Academic Qualification:"
         Height          =   315
         Left            =   15
         TabIndex        =   19
         Top             =   3210
         Width           =   1695
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Parent's Name:"
         Height          =   315
         Left            =   120
         TabIndex        =   18
         Top             =   3630
         Width           =   1575
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Phone:"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   2820
         Width           =   1455
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "City:"
         Height          =   255
         Left            =   255
         TabIndex        =   16
         Top             =   2340
         Width           =   1455
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   1965
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   1125
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "First Name:"
         Height          =   315
         Left            =   120
         TabIndex        =   13
         Top             =   705
         Width           =   1575
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Student ID:"
         Height          =   285
         Left            =   840
         TabIndex        =   12
         Top             =   300
         Width           =   810
      End
   End
End
Attribute VB_Name = "StudentEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection

Private Sub Command1_Click()
On Error GoTo myerror:

If Text1 = "" Or Text2 = "" Or Combo4 = "" Or Combo1 = "" Or Combo2 = "" Or Text5 = "" Or Combo3 = "" Or Text7 = "" Or Text9 = "" Then
    MsgBox "Fill up all fields"
Exit Sub
End If

Dim strsql As String
strsql = "Select * from Student where [Student ID]='" & Text9.Text & "'"
Set rs = New ADODB.Recordset
rs.Open strsql, con, adOpenForwardOnly, adLockOptimistic

If MsgBox("           Are you sure(Y/N)???", 4) = 7 Then Exit Sub
    
rs.Update("First Name") = UCase$(Text1.Text)
rs.Update("Last Name") = UCase$(Text2.Text)
rs.Update("Gender") = UCase$(Combo4.Text)
rs.Update("Address") = UCase$(Combo1.Text)
rs.Update("City") = UCase$(Combo2.Text)
rs.Update("Phone") = Text5.Text
rs.Update("Qualification") = UCase$(Combo3.Text)
rs.Update("Parent Name") = UCase$(Text7.Text)
rs.Update("Date of Join") = Dpic1

Dim strmth As String
Dim nmth As Integer
nmth = Month(Dpic1)
If nmth = 1 Then strmth = "January"
If nmth = 2 Then strmth = "February"
If nmth = 3 Then strmth = "March"
If nmth = 4 Then strmth = "April"
If nmth = 5 Then strmth = "May"
If nmth = 6 Then strmth = "June"
If nmth = 7 Then strmth = "July"
If nmth = 8 Then strmth = "August"
If nmth = 9 Then strmth = "September"
If nmth = 10 Then strmth = "October"
If nmth = 11 Then strmth = "November"
If nmth = 12 Then strmth = "December"
rs.Update("Months") = strmth

rs.Close
myerror:
Call error

End Sub

Private Sub Command3_Click()
Unload Me
StudentList.Show
End Sub

Private Sub Form_Load()
On Error GoTo myerror:

Set con = CreateObject("Adodb.connection")
con.Provider = "Microsoft.jet.oledb.4.0"
con.Open App.Path & "\MainData.mdb"

Dim strsql As String
strsql = "Select distinct Address from student"
Set rs = New ADODB.Recordset
rs.Open strsql, con, adOpenForwardOnly, adLockOptimistic

Combo4.AddItem "FEMALE"
Combo4.AddItem "MALE"

While Not rs.EOF
Combo1.AddItem rs.Fields("Address")
rs.MoveNext
Wend
rs.Close

Dim strsql1 As String
strsql1 = "Select distinct City from student"
Set rs1 = New ADODB.Recordset
rs1.Open strsql1, con, adOpenForwardOnly, adLockOptimistic

While Not rs1.EOF
Combo2.AddItem rs1.Fields("City")
rs1.MoveNext
Wend
rs1.Close

Dim strsql2 As String
strsql2 = "Select distinct Qualification from student"
Set rs2 = New ADODB.Recordset
rs2.Open strsql2, con, adOpenForwardOnly, adLockOptimistic

While Not rs2.EOF
Combo3.AddItem rs2.Fields("Qualification")
rs2.MoveNext
Wend
rs2.Close

Text9.Text = StudentList.DataGrid1.Columns(0).Text

Dim strsql10 As String
strsql10 = "Select * from Student where [Student ID]='" & Text9.Text & "'"
Set rs10 = New ADODB.Recordset
rs10.Open strsql10, con, adOpenForwardOnly, adLockOptimistic

If Not rs10.EOF Then
    Text1.Text = rs10.Fields("First Name")
    Text2.Text = rs10.Fields("Last Name")
    Combo4.Text = rs10.Fields("Gender")
    Combo1.Text = rs10.Fields("Address")
    Combo2.Text = rs10.Fields("City")
    Text5.Text = rs10.Fields("Phone")
    Combo3.Text = rs10.Fields("Qualification")
    Text7.Text = rs10.Fields("Parent Name")
    Dpic1 = rs10.Fields("Date of Join")
End If
rs10.Close
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
StudentList.Show
End Sub

