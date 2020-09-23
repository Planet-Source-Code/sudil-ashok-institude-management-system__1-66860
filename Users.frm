VERSION 5.00
Begin VB.Form Users 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Entry"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6675
   Icon            =   "Users.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "Users.frx":1708A
   ScaleHeight     =   2520
   ScaleWidth      =   6675
   Begin VB.CommandButton Command5 
      Caption         =   "Delete User"
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
      Left            =   5160
      TabIndex        =   14
      Top             =   2040
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   4680
      TabIndex        =   12
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "->>"
      Height          =   375
      Left            =   4200
      TabIndex        =   11
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox Text4 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   600
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1560
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Create User"
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
      TabIndex        =   4
      Top             =   2040
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
      Left            =   3120
      TabIndex        =   6
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change Password"
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
      TabIndex        =   5
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1080
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label5 
      Caption         =   "User List"
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
      Left            =   4680
      TabIndex        =   13
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Old Password"
      Height          =   255
      Left            =   480
      TabIndex        =   10
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Conform Password"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "New Password"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Users"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection

Private Sub Command1_Click()
Dim strsql As String
Dim users As String
users = UCase$(Text1.Text)
strsql = "Select * from Users where username='" & users & "'"
Set rs = New ADODB.Recordset
rs.Open strsql, con, adOpenForwardOnly, adLockOptimistic

If Text2.Text <> Text3.Text Then MsgBox ("Password do not match!"): Exit Sub

If rs.EOF = False Then
    If rs.Fields("Password") = encrypt(Text4.Text) Then
        rs.Fields("Password") = encrypt(Text3.Text)
        rs.Update
    Else
        MsgBox ("Old Password do not match!")
        Exit Sub
    End If
MsgBox ("Password successfully change!!!")
rs.Close
Unload Me
End If

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()

Dim strsql As String
strsql = "Select * from Users"
Set rs = New ADODB.Recordset
rs.Open strsql, con, adOpenForwardOnly, adLockOptimistic

If Text2.Text <> Text3.Text Then MsgBox ("Password do not match"): Exit Sub

rs.AddNew
If Not rs.EOF Then
rs.Fields("Username") = UCase$(Text1.Text)
rs.Fields("Password") = encrypt(Text3.Text)
rs.Update
End If
rs.Close
Text1.Text = Clear
Text2.Text = Clear
Text3.Text = Clear
Text4.Text = Clear
End Sub


Private Sub Command4_Click()
users.Width = 6700
Command5.Enabled = False
List1.Clear
Dim strsql As String
strsql = "Select * from Users"
Set rs = New ADODB.Recordset
rs.Open strsql, con, adOpenForwardOnly, adLockOptimistic
While Not rs.EOF
List1.AddItem rs.Fields("Username")
rs.MoveNext
Wend
rs.Close
End Sub

Private Sub Command5_Click()
Dim strsql As String
strsql = "Select * from Users where username='" & List1.Text & "'"
Set rs = New ADODB.Recordset
rs.Open strsql, con, adOpenForwardOnly, adLockOptimistic
rs.Delete
rs.Close
Call Command4_Click
End Sub

Private Sub Form_Load()
If rams = 1 Then
users.Width = 4300
Command1.Visible = False
Command4.Enabled = False
Text4.Enabled = False
ElseIf rams = 2 Then
users.Width = 4700
Command3.Enabled = False
End If
Set con = CreateObject("Adodb.connection")
con.Provider = "Microsoft.jet.oledb.4.0"
con.Open App.Path & "\maindata.mdb"
End Sub

Function encrypt(s As String)
Dim p As Integer
Dim a As String
a = ""
For i = 1 To Len(s)
p = Asc(Mid$(s, i, 1))
a = a + Chr(p Xor 6)
Next i
encrypt = a
End Function

Private Sub Form_Unload(Cancel As Integer)
con.Close
End Sub

Private Sub List1_Click()
Command5.Enabled = True
End Sub
