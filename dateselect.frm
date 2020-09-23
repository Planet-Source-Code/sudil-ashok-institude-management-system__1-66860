VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Dateselect 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select Date"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2970
   Icon            =   "dateselect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "dateselect.frx":1708A
   ScaleHeight     =   1665
   ScaleWidth      =   2970
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.DTPicker DTP2 
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Top             =   480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      Format          =   52756481
      CurrentDate     =   37951
   End
   Begin MSComCtl2.DTPicker DTP1 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      Format          =   52756481
      CurrentDate     =   37951
   End
   Begin VB.Image Command2 
      Height          =   555
      Left            =   1920
      MouseIcon       =   "dateselect.frx":2BC4A
      MousePointer    =   99  'Custom
      Picture         =   "dateselect.frx":2BF54
      ToolTipText     =   "Exit"
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Command1 
      Height          =   720
      Left            =   480
      MouseIcon       =   "dateselect.frx":2C6FE
      MousePointer    =   99  'Custom
      Picture         =   "dateselect.frx":2CA08
      ToolTipText     =   "OK"
      Top             =   960
      Width           =   720
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Ending Date:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Starting Date:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Dateselect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If rams = 1 Then
DataEnvironment1.Connection1.ConnectionString = App.Path & "\MainData.mdb"
DataEnvironment1.Connection1.Open
DataEnvironment1.Command1 DTP1.Value, DTP2.Value
DataReport1.Show
ElseIf rams = 7 Then
DataEnvironment1.Connection1.ConnectionString = App.Path & "\MainData.mdb"
DataEnvironment1.Connection1.Open
DataEnvironment1.Command7_Grouping DTP1.Value, DTP2.Value
DataReport7.Show
ElseIf rams = 8 Then
DataEnvironment1.Connection1.ConnectionString = App.Path & "\MainData.mdb"
DataEnvironment1.Connection1.Open
DataEnvironment1.Command8_Grouping DTP1.Value, DTP2.Value
DataReport8.Show
End If
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
DTP1.Value = Date
DTP2.Value = Date
End Sub
