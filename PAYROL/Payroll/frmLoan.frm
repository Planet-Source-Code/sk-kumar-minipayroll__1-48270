VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmLoan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loan Master"
   ClientHeight    =   6105
   ClientLeft      =   1320
   ClientTop       =   1230
   ClientWidth     =   8205
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   8205
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      Top             =   5280
      Width           =   735
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4020
      TabIndex        =   11
      Top             =   5280
      Width           =   735
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   24
      Top             =   5730
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9287
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "1:05 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "10/14/01"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   1140
      TabIndex        =   8
      Top             =   5280
      Width           =   735
   End
   Begin VB.ComboBox cboBranch 
      Height          =   345
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&New"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   5280
      Width           =   735
   End
   Begin VB.ListBox LstLoan 
      Height          =   3210
      Left            =   5520
      Sorted          =   -1  'True
      TabIndex        =   13
      Top             =   1320
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   375
      Left            =   5040
      TabIndex        =   12
      Top             =   5280
      Width           =   735
   End
   Begin VB.TextBox txtEmpName 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   2235
      Width           =   1815
   End
   Begin VB.CommandButton command1 
      Caption         =   "&Save"
      Height          =   375
      Left            =   3000
      TabIndex        =   10
      Top             =   5280
      Width           =   735
   End
   Begin MSMask.MaskEdBox mskDateApp 
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   3465
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtReason 
      Height          =   375
      Left            =   2040
      MaxLength       =   250
      TabIndex        =   6
      Top             =   4680
      Width           =   1815
   End
   Begin VB.TextBox txtInstall 
      Height          =   375
      Left            =   2040
      MaxLength       =   8
      TabIndex        =   5
      Top             =   4065
      Width           =   1815
   End
   Begin VB.TextBox txtLoan 
      Height          =   375
      Left            =   2040
      MaxLength       =   12
      TabIndex        =   3
      Top             =   2850
      Width           =   1815
   End
   Begin VB.ComboBox cboEmpNo 
      Height          =   345
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1665
      Width           =   1815
   End
   Begin VB.Label lblCompanyName 
      AutoSize        =   -1  'True
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   2760
      TabIndex        =   26
      Top             =   480
      Width           =   45
   End
   Begin VB.Label Label11 
      Caption         =   "CompanyName"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   1200
      TabIndex        =   25
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "MM/DD/YYYY"
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   4080
      TabIndex        =   23
      Top             =   3480
      Width           =   1200
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Branch"
      Height          =   225
      Left            =   360
      TabIndex        =   22
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Select (LoanId-LoanAmount)"
      Height          =   225
      Left            =   5400
      TabIndex        =   21
      Top             =   840
      Width           =   2340
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Emp name"
      Height          =   225
      Left            =   360
      TabIndex        =   20
      Top             =   2280
      Width           =   825
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Reason"
      Height          =   225
      Left            =   360
      TabIndex        =   19
      Top             =   4680
      Width           =   600
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "No Of Instalments"
      Height          =   225
      Left            =   360
      TabIndex        =   18
      Top             =   4080
      Width           =   1485
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Date Of Approval"
      Height          =   225
      Left            =   360
      TabIndex        =   17
      Top             =   3480
      Width           =   1395
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Loan Amount"
      Height          =   225
      Left            =   360
      TabIndex        =   16
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Employee No :"
      Height          =   225
      Left            =   360
      TabIndex        =   15
      Top             =   1680
      Width           =   1140
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Loan Master"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   285
      Left            =   2565
      TabIndex        =   14
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "frmLoan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***Dim cnn As New ADODB.Connection
'***Dim rs As New ADODB.Recordset
'***Dim rs1 As New ADODB.Recordset
Dim AddEditViewMode As String
Dim Msg As String
Private Sub cbobranch_Click()
clear
addcombo
End Sub
Private Sub cboEmpno_Click()
AddList
AddEditViewMode = "View"
StatusBar1.Panels(1) = AddEditViewMode
    sql = "select ename from admin where CompanyName='" & lblCompanyName.Caption & "' and empno=" & Mid(cboEmpno, InStr(1, cboEmpno, "/") + 1)
    rs.Open sql, cnn, adOpenKeyset, adLockOptimistic
    txtEmpName = rs(0)
    rs.Close
'    txtLoan.SetFocus

Command3.Enabled = NoOfLoans
If Command3.Enabled = False Then
    StatusBar1.Panels(1) = "Already gave Three Loans"
End If
End Sub
Private Sub CmdCancel_Click()
Form_Load
lstLoan.Enabled = True
cboBranch.Enabled = True
cboEmpno.Enabled = True
If cboEmpno.ListCount > 0 Then
    Command3.Enabled = True
Else
    Command3.Enabled = False
End If
AddEditViewMode = "View"
StatusBar1.Panels(1) = AddEditViewMode
End Sub
Private Sub CmdDelete_Click()
Msg = MsgBox("Are you sure to delete this Record", vbExclamation + vbYesNo, "Payroll")
If Msg = 6 Then

    cnn.Execute "Delete from LoanMaster Where CompanyName='" & lblCompanyName.Caption & "' and empno=" & Mid(cboEmpno, InStr(1, cboEmpno, "/") + 1) & " And Loanid = " & Mid(lstLoan, 1, InStr(1, lstLoan, "-") - 1)
    
End If
AddList
sql = "select ename from admin where CompanyName='" & lblCompanyName.Caption & "' and empno=" & Mid(cboEmpno, InStr(1, cboEmpno, "/") + 1)
    rs.Open sql, cnn, adOpenKeyset, adLockOptimistic
    txtEmpName = rs(0)
    rs.Close
End Sub
Private Sub cmdEdit_Click()
AddEditViewMode = "Edit"
lockText (False)
cmdEdit.Enabled = False
Command1.Enabled = True
mskDateApp.Enabled = True
cboBranch.Enabled = False
cboEmpno.Enabled = False
lstLoan.Enabled = False
txtLoan.SetFocus
StatusBar1.Panels(1) = AddEditViewMode
End Sub
Private Sub Command1_Click()   'Save Button
If Len(txtLoan) = 0 Or Len(txtInstall) = 0 Or Format(mskDateApp, "mm/dd/yyyy") = "__/__/____" Then
    MsgBox "Loan,Install and DateApproved should not be Empty Value", vbExclamation, "Payroll"
    Exit Sub
End If
If Val(txtLoan) < Val(txtInstall) Then
    MsgBox "LoanAmount Should be Greater Than No Of Installment Month", vbExclamation, "Payroll"
    Exit Sub
End If
   
    If AddEditViewMode = "Add" Then
        sql1 = "select  max(LoanId) as [pp] from loanMaster where CompanyName='" & lblCompanyName.Caption & "' and empno=" & Mid(cboEmpno, InStr(1, cboEmpno, "/") + 1)
        sql = "Select *from LoanMaster"
        rs.Open sql, cnn, adOpenKeyset, adLockOptimistic
        rs1.Open sql1, cnn, adOpenKeyset, adLockOptimistic
        rs.AddNew
        If IsNull(rs1.Fields(0)) Then
            rs(0) = 1
        Else
            rs.Fields(0) = Val(rs1(0)) + 1
        End If
        rs1.Close
     ElseIf AddEditViewMode = "Edit" Then

        sql = "Select *from LoanMaster Where CompanyName='" & lblCompanyName.Caption & "' and Empno=" & Mid(cboEmpno, InStr(1, cboEmpno, "/") + 1) & " and Loanid=" & Mid(lstLoan, 1, InStr(1, lstLoan, "-") - 1)
        rs.Open sql, cnn, adOpenKeyset, adLockOptimistic
        
     End If
     
        rs.Fields(1) = Format(Mid(cboEmpno, InStr(1, cboEmpno, "/") + 1), "000")
        rs.Fields(2) = txtLoan
        rs.Fields(3) = Format(mskDateApp, "mm/dd/yyyy")
        rs.Fields(4) = txtInstall
        rs.Fields(5) = txtReason
        rs.Fields(6) = txtLoan
        rs.Fields(7) = CompanyName
        rs.Update
    
        rs.Close
       
    AddEditViewMode = "View"
    Command1.Enabled = False
    cboBranch.Enabled = True
    cboEmpno.Enabled = True
    lstLoan.Enabled = True
    AddList
    lockText (True)
    mskDateApp.Enabled = False
    AddEditViewMode = "View"
    
  StatusBar1.Panels(1) = AddEditViewMode
   
End Sub
Private Sub Command2_Click()
   '*** cnn.Close
    Unload Me
End Sub
Private Sub Command3_Click()  'New
AddEditViewMode = "Add"
txtLoan = ""
txtInstall = ""
txtReason = ""
mskDateApp = "__/__/____"

Command3.Enabled = False
Command1.Enabled = True
cmdEdit.Enabled = False
CmdDelete.Enabled = False
cboBranch.Enabled = False
cboEmpno.Enabled = False
lstLoan.Enabled = False
lockText (False)
mskDateApp.Enabled = True
txtLoan.SetFocus
StatusBar1.Panels(1) = AddEditViewMode
End Sub
Private Sub Command4_Click()
Form_Load
End Sub
Private Sub Form_Activate()
frmLoan.Left = 0
frmLoan.Top = 0
End Sub
Private Sub Form_Load()
'***cnn.Provider = "Microsoft.Jet.OLEDB.4.0"
'*** cnn.Open app.path & "\Moolakadai.mdb"
frmLoan.Icon = LoadPicture(App.Path & "\Jamaica my computer.ico")
lblCompanyName.Caption = CompanyName
mskDateApp.Text = "__/__/____"
mskDateApp.Enabled = False
Command1.Enabled = False
cmdEdit.Enabled = False
CmdDelete.Enabled = False
rs.Open "Select BranchCode From Branch where CompanyName='" & lblCompanyName.Caption & "'", cnn, adOpenKeyset, adLockOptimistic
cboBranch.clear
While Not rs.EOF
cboBranch.AddItem rs.Fields(0)
rs.MoveNext
Wend
rs.Close

If cboBranch.ListCount > 0 Then
    cboBranch.ListIndex = 0
End If
If cboEmpno.ListCount > 0 Then
    Command3.Enabled = True
Else
    Command3.Enabled = False
End If
lockText (True)
AddEditViewMode = "View"
StatusBar1.Panels(1) = AddEditViewMode
End Sub
Public Sub AddList()
rs.Open "Select loanid,LoanAmt from LoanMaster where CompanyName='" & lblCompanyName.Caption & "' and empno=" & Mid(cboEmpno, InStr(1, cboEmpno, "/") + 1), cnn, adOpenKeyset, adLockOptimistic
lstLoan.clear
While Not rs.EOF
    lstLoan.AddItem rs.Fields(0) & "-" & rs.Fields(1)
    rs.MoveNext
Wend
rs.Close
If lstLoan.ListCount > 0 Then
    lstLoan.ListIndex = 0
Else
    clear
    cmdEdit.Enabled = False
    CmdDelete.Enabled = False
    
End If
End Sub
Private Sub LstLoan_Click()
rs.Open "Select * from LoanMaster where CompanyName='" & lblCompanyName.Caption & "' and empno=" & Mid(cboEmpno, InStr(1, cboEmpno, "/") + 1) & " and loanid=" & Left(lstLoan, InStr(1, lstLoan, "-") - 1), cnn, adOpenKeyset, adLockOptimistic
txtLoan = rs.Fields("LoanAmt")
mskDateApp.Text = Format(rs.Fields("dateapp"), "mm/dd/yyyy")
txtInstall = rs.Fields("Noinstall")
txtReason = rs.Fields("Reason")
rs.Close
rs.Open "Select count(*) from LoanDetails where CompanyName='" & lblCompanyName.Caption & "' and empno=" & Mid(cboEmpno, InStr(1, cboEmpno, "/") + 1) & " and loanid=" & Left(lstLoan, InStr(1, lstLoan, "-") - 1), cnn, adOpenKeyset, adLockOptimistic
If rs.Fields(0) > 0 Then
    cmdEdit.Enabled = False
    CmdDelete.Enabled = False
Else
    cmdEdit.Enabled = True
    CmdDelete.Enabled = True
End If
rs.Close
Command3.Enabled = NoOfLoans
If Command3.Enabled = False Then
    StatusBar1.Panels(1) = "Already Gave Three Loans"
Else
    StatusBar1.Panels(1) = AddEditViewMode
End If


End Sub
Public Sub addcombo()
sql = "Select Empno From Admin Where CompanyName='" & lblCompanyName.Caption & "' and Branchcode='" & cboBranch & "' Order by empno"
rs.Open sql, cnn, adOpenKeyset, adLockOptimistic
cboEmpno.clear
While Not rs.EOF
    cboEmpno.AddItem SubName & "/" & Format(rs.Fields(0), "000")
    rs.MoveNext
Wend
rs.Close
If cboEmpno.ListCount > 0 Then
    cboEmpno.ListIndex = 0
End If
End Sub
Public Sub lockText(Locked As Boolean)
txtInstall.Locked = Locked
txtLoan.Locked = Locked
txtReason.Locked = Locked
End Sub
Private Sub mskDateApp_LostFocus()
If mskDateApp.Text <> "__/__/____" Then
If CheckDate(mskDateApp) = False Then
    MsgBox "Invalid Date", vbExclamation, "Payroll"
    mskDateApp.Text = "__/__/____"
    mskDateApp.SetFocus
    mskDateApp.SetFocus
    Exit Sub
End If

If CheckMonth(mskDateApp) = False Then
    MsgBox "Invalid Month", vbExclamation, "Payroll"
    mskDateApp.Text = "__/__/____"
    mskDateApp.SetFocus
    Exit Sub
End If
End If

End Sub
Private Sub txtInstall_KeyPress(KeyAscii As Integer)
If KeyAscii > 46 And KeyAscii < 59 Or KeyAscii = 8 Then
Else
    KeyAscii = 0
End If
End Sub
Private Sub txtLoan_KeyPress(KeyAscii As Integer)
KeyAscii = NumericCheck(txtInstall, CInt(KeyAscii))
End Sub
Public Function NoOfLoans() As Boolean
rs.Open "Select count(*) from LoanMaster Where CompanyName='" & lblCompanyName.Caption & "' and empno=" & Mid(cboEmpno, InStr(1, cboEmpno, "/") + 1) & " And Balance <> 0"
If rs.Fields(0) < 3 Then
    NoOfLoans = True
Else
    NoOfLoans = False
End If
rs.Close
End Function
