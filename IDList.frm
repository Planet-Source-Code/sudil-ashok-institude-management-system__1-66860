VERSION 5.00
Begin VB.Form IDList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ID list....."
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4350
   Icon            =   "IDList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "IDList.frx":1708A
   ScaleHeight     =   6135
   ScaleWidth      =   4350
   Begin VB.ListBox List1 
      BackColor       =   &H00C0C0C0&
      Height          =   5130
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   4335
   End
   Begin VB.Image Command1 
      Height          =   720
      Left            =   2880
      MouseIcon       =   "IDList.frx":2BC4A
      MousePointer    =   99  'Custom
      Picture         =   "IDList.frx":2BF54
      Stretch         =   -1  'True
      ToolTipText     =   "Srudents"
      Top             =   120
      Width           =   720
   End
   Begin VB.Image Command2 
      Height          =   720
      Left            =   1680
      MouseIcon       =   "IDList.frx":2C41C
      MousePointer    =   99  'Custom
      Picture         =   "IDList.frx":2C726
      Stretch         =   -1  'True
      ToolTipText     =   "Instructors"
      Top             =   120
      Width           =   720
   End
   Begin VB.Image Command3 
      Height          =   720
      Left            =   360
      MouseIcon       =   "IDList.frx":2CBBC
      MousePointer    =   99  'Custom
      Picture         =   "IDList.frx":2CEC6
      ToolTipText     =   "Account Head"
      Top             =   120
      Width           =   720
   End
End
Attribute VB_Name = "IDList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection


Private Sub Command1_Click()
List1.Clear
Dim strsql As String
strsql = "Select * From [Student]"
Set rs = New ADODB.Recordset
rs.Open strsql, con, adOpenForwardOnly, adLockOptimistic
While Not rs.EOF
List1.AddItem rs.Fields("Student ID") & " --> " & rs.Fields("First Name") & " " & rs.Fields("Last Name")
rs.MoveNext
Wend
rs.Close
End Sub

Private Sub Command2_Click()
List1.Clear
Dim strsql1 As String
strsql1 = "Select * From [Instructor]"
Set rs1 = New ADODB.Recordset
rs1.Open strsql1, con, adOpenForwardOnly, adLockOptimistic
While Not rs1.EOF
List1.AddItem rs1.Fields("Instructor ID") & " --> " & rs1.Fields("Name")
rs1.MoveNext
Wend
rs1.Close
End Sub

Private Sub Command3_Click()
List1.Clear
Dim strsql3 As String
strsql3 = "Select * From [Account Head]order by [sno]"
Set rs3 = New ADODB.Recordset
rs3.Open strsql3, con, adOpenForwardOnly, adLockOptimistic
While Not rs3.EOF
List1.AddItem rs3.Fields("Head ID") & " --> " & rs3.Fields("Head")
rs3.MoveNext
Wend
rs3.Close

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
Set con = CreateObject("Adodb.connection")
con.Provider = "Microsoft.jet.oledb.4.0"
con.Open App.Path & "\MainData.mdb"
If rams1 = 1 Then Call Command1_Click
End Sub

Private Sub List1_DblClick()
Dim idselected
idselected = IDList.List1.Text
For i = 1 To Len(idselected)
    If Mid$(idselected, i, 1) = " " Then Exit For
Next i
If Command2.Enabled = False And Command3.Enabled = False Then
Receive.Text1.Text = Left$(idselected, i - 1)
Receive.Text4.Text = Mid$(idselected, i + 5, Len(idselected) - i + 5)
Receive.Dpic1.SetFocus
Else
Transaction.Text1.Text = Left$(idselected, i - 1)
Transaction.Text3.Text = Mid$(idselected, i + 5, Len(idselected) - i + 5)
End If
Unload Me
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Dim idselected
idselected = IDList.List1.Text
For i = 1 To Len(idselected)
    If Mid$(idselected, i, 1) = " " Then Exit For
Next i
If Command2.Enabled = False And Command3.Enabled = False Then
Receive.Text1.Text = Left$(idselected, i - 1)
Receive.Text4.Text = Mid$(idselected, i + 5, Len(idselected) - i + 5)
Receive.Dpic1.SetFocus
Else
Transaction.Text1.Text = Left$(idselected, i - 1)
Transaction.Text3.Text = Mid$(idselected, i + 5, Len(idselected) - i + 5)
End If
Unload Me
End If
End Sub
