Attribute VB_Name = "InstructorList"
Dim con As ADODB.Connection

Private Sub Command1_Click()
rams1 = 3
InstructorDetails.Show
End Sub

Private Sub Command2_Click()
Group.Show
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
Percentage.Show
End Sub

Private Sub Form_Activate()
Call Form_Load
End Sub

Private Sub Form_Load()
Set con = CreateObject("Adodb.connection")
con.Provider = "Microsoft.jet.oledb.4.0"
con.Open App.Path & "\maindata.mdb"
Dim strsql As String
strsql = "SELECT Instructor.[Instructor ID], Instructor.Name, Instructor.Address, Instructor.Phone, SumOfAmount.TotalPayment, SumOfAmount.TotalPaid, Sum([TotalPayment]-[TotalPaid]) AS Balance FROM Instructor INNER JOIN SumOfAmount ON Instructor.[Instructor ID] = SumOfAmount.[Acc Head ID] GROUP BY Instructor.[Instructor ID], Instructor.Name, Instructor.Address, Instructor.Phone, SumOfAmount.TotalPayment, SumOfAmount.TotalPaid"
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open strsql, con, adOpenForwardOnly, adLockOptimistic
Set DataGrid1.DataSource = rs
DataGrid1.Columns(0).Caption = "I. ID"
DataGrid1.Columns(0).Width = 400
DataGrid1.Columns(1).Width = 2400
DataGrid1.Columns(2).Width = 2800
DataGrid1.Columns(3).Width = 800
DataGrid1.Columns(4).Width = 1100
DataGrid1.Columns(4).Alignment = dbgRight
DataGrid1.Columns(5).Width = 800
DataGrid1.Columns(5).Alignment = dbgRight
DataGrid1.Columns(6).Width = 800
DataGrid1.Columns(6).Alignment = dbgRight
End Sub

Private Sub Form_Unload(Cancel As Integer)
con.Close
End Sub
