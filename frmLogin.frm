VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1710
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":1708A
   ScaleHeight     =   1010.324
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Top             =   135
      Width           =   2325
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   2325
   End
   Begin VB.Image cmdCancel 
      Height          =   555
      Left            =   2640
      MouseIcon       =   "frmLogin.frx":2BC4A
      MousePointer    =   99  'Custom
      Picture         =   "frmLogin.frx":2BF54
      ToolTipText     =   "Cancel"
      Top             =   1080
      Width           =   600
   End
   Begin VB.Image cmdOK 
      Height          =   720
      Left            =   1320
      MouseIcon       =   "frmLogin.frx":2C6FE
      MousePointer    =   99  'Custom
      Picture         =   "frmLogin.frx":2CA08
      ToolTipText     =   "OK"
      Top             =   960
      Width           =   720
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim con As ADODB.Connection
Dim rs As ADODB.Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
Set con = CreateObject("Adodb.connection")
con.Provider = "Microsoft.jet.oledb.4.0"
con.Open App.Path & "\MainData.mdb"
    
Dim strsql As String
strsql = "Select * from Users"
Set rs = New ADODB.Recordset
rs.Open strsql, con, adOpenForwardOnly, adLockOptimistic
Dim chkPassword

While Not rs.EOF
If UCase$(txtUserName) = rs.Fields("username") And encrypt(txtPassword) = rs.Fields("Password") Then chkPassword = "password"
rs.MoveNext
Wend
    
    If chkPassword = "password" Then
    Call MDIForm1.Show
    Unload Me
    con.Close
    
    Dim soc As String
    Dim dist As String
    soc = "" & App.Path & "\MainData.mdb"
    dist = "" & App.Path & "\data\MainData.mdb"
    FileSystem.SetAttr soc, vbNormal
    FileSystem.FileCopy soc, dist
    
    Else
        MsgBox "Invalid Password, try again!", , "PCI Login"
        SendKeys "{Home}+{End}"
    End If
End Sub

Function encrypt(s As String)
Dim p As Integer
Dim i As Integer
Dim a As String
a = ""
For i = 1 To Len(s)
p = Asc(Mid$(s, i, 1))
a = a + Chr(p Xor 6)
Next i
encrypt = a
End Function

