VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmResLoanMaster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resign Employee's LoanMaster"
   ClientHeight    =   6540
   ClientLeft      =   180
   ClientTop       =   1260
   ClientWidth     =   10230
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   10230
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   345
      Left            =   4080
      TabIndex        =   5
      Top             =   6120
      Width           =   855
   End
   Begin VB.TextBox txtEmpNo 
      Height          =   330
      Left            =   4560
      TabIndex        =   3
      Top             =   1440
      Width           =   1935
   End
   Begin MSComctlLib.ListView lstLoan 
      Height          =   3735
      Left            =   240
      TabIndex        =   1
      Top             =   2160
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   6588
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.ComboBox cbResignNo 
      Height          =   345
      Left            =   4560
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label lblCompanyName 
      AutoSize        =   -1  'True
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   4920
      TabIndex        =   7
      Top             =   240
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "CompanyName"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   3000
      TabIndex        =   6
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Empno"
      Height          =   225
      Left            =   3360
      TabIndex        =   4
      Top             =   1440
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ResignNo."
      Height          =   225
      Left            =   3360
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
End
Attribute VB_Name = "frmResLoanMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lt As ListItem
Private Sub cbResignNo_Click()
If cbResignNo <> "" Then
lstLoan.ListItems.clear
rs.Open "Select * from ResLoanMaster where CompanyName='" & lblCompanyName.Caption & "' and resignNo=" & cbResignNo, cnn, adOpenKeyset, adLockOptimistic
txtEmpNo = SubName & "/" & Format(rs.Fields("Empno"), "000")
While Not rs.EOF
    Set lt = lstLoan.ListItems.Add(, , rs.Fields("Loanid"))
    lt.SubItems(1) = Format(rs.Fields("LoanAmt"), "0.00")
    lt.SubItems(2) = rs.Fields("DateApp")
    lt.SubItems(3) = rs.Fields("NoInstall")
    lt.SubItems(4) = Format(rs.Fields("Balance"), "0.00")
    lt.SubItems(5) = rs.Fields("Reason")
rs.MoveNext
Wend
End If
rs.Close
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub Form_Activate()
frmResLoanMaster.Left = 0
frmResLoanMaster.Top = 0
End Sub
Private Sub Form_Load()
frmResLoanMaster.Icon = LoadPicture(App.Path & "\Jamaica my computer.ico")
lstLoan.ColumnHeaders.clear
lstLoan.ColumnHeaders.Add , , "LoanId", 1100
lstLoan.ColumnHeaders.Add , , "LoanAmount", 1700
lstLoan.ColumnHeaders.Add , , "DateApproved", 1100
lstLoan.ColumnHeaders.Add , , "No. of Instalments", 1100
lstLoan.ColumnHeaders.Add , , "Balance", 1700
lstLoan.ColumnHeaders.Add , , "Reason", 3000
lblCompanyName.Caption = CompanyName
AddResignNo
End Sub
Public Sub AddResignNo()
rs.Open "Select Distinct(ResignNo) from ResLoanMaster Where CompanyName='" & lblCompanyName.Caption & "'", cnn, adOpenKeyset, adLockOptimistic
cbResignNo.clear
While Not rs.EOF
cbResignNo.AddItem rs.Fields("ResignNo")
rs.MoveNext
Wend
rs.Close
End Sub
