VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Payroll"
   ClientHeight    =   6495
   ClientLeft      =   570
   ClientTop       =   285
   ClientWidth     =   9480
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuopen 
         Caption         =   "&Open"
         Begin VB.Menu mnucalc 
            Caption         =   "&Calculator"
         End
         Begin VB.Menu mnunote 
            Caption         =   "&Notepad"
         End
      End
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuMaster 
      Caption         =   "&Master"
      Begin VB.Menu mnuCompany 
         Caption         =   "&Company Master"
      End
      Begin VB.Menu mnuBranch 
         Caption         =   "&Branch Master"
      End
      Begin VB.Menu mnuUser 
         Caption         =   "&User Master"
      End
      Begin VB.Menu mnuDesignation 
         Caption         =   "&Designation"
      End
      Begin VB.Menu mnuEmployeeMaster 
         Caption         =   "&EmployeeMaster"
      End
   End
   Begin VB.Menu mnurangepfesi 
      Caption         =   "Range&OfPF&&ESI"
      Begin VB.Menu mnurangepf 
         Caption         =   "Range &PF"
      End
      Begin VB.Menu mnurangeesi 
         Caption         =   "Range &ESI"
      End
   End
   Begin VB.Menu mnuTransaction 
      Caption         =   "&Transaction"
      Begin VB.Menu mnuEntry 
         Caption         =   "Monthly Pay Entry"
      End
      Begin VB.Menu mnuLoan 
         Caption         =   "&Loan Master"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&Reports"
      Begin VB.Menu mnupay 
         Caption         =   "Monthly Pay Report"
      End
      Begin VB.Menu mnuslippay 
         Caption         =   "MonthlyPaySlip"
      End
      Begin VB.Menu mnuAdvance 
         Caption         =   "&Advance"
      End
   End
   Begin VB.Menu mnuresreport 
      Caption         =   "Re&signReports"
      Begin VB.Menu mnuresempDetails 
         Caption         =   "Res EmpDetails"
      End
      Begin VB.Menu mnuEmpPayDetails 
         Caption         =   "Resign Emp Pay Details"
      End
      Begin VB.Menu mnuResLoanMaster 
         Caption         =   "Res &Loan Master"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About  Payroll"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Activate()
MDIForm1.Caption = CompanyName & "-" & "Payroll"
End Sub
Private Sub MDIForm_Load()
'    MsgBox MDIForm1.Width
  '  MsgBox Picture1.Width
   ' Picture1.BorderStyle = 0
   MDIForm1.Icon = LoadPicture(App.Path & "\Jamaica my computer.ico")
   MDIForm1.Picture = LoadPicture(App.Path & "\Beach2.jpg")
   
End Sub
Private Sub MDIForm_Unload(Cancel As Integer)
    cnn.Close
End Sub
Private Sub mnuAbout_Click()
    frmAbout.Show
End Sub

Private Sub mnuAdvance_Click()
frmAdvance.Show
End Sub

Private Sub mnuBranch_Click()
    frmBranch.Show
End Sub
Private Sub mnucalc_Click()
On Error GoTo err
    Shell App.Path & "\calc.exe", vbNormalFocus
err:
    If err.Number = 53 Then
        MsgBox "File Not Found", vbExclamation, "Payroll"
    End If
    Exit Sub
End Sub

Private Sub mnuCompany_Click()
    frmCompany.Show
End Sub

Private Sub mnuDesignation_Click()
    frmdesignationMaster.Show
End Sub
Private Sub mnuEmployeeMaster_Click()
    frmEmployeeDetails.Show
End Sub
Private Sub mnuEmpPayDetails_Click()
    frmResEmppayDetails.Show
End Sub
Private Sub mnuEntry_Click()
    frmCalendar.Show
End Sub
Private Sub mnuExit_Click()
    End
End Sub
Private Sub mnuLoan_Click()
    frmLoan.Show
End Sub
Private Sub mnunote_Click()
On Error GoTo err
    Shell App.Path & "\Notepad.exe", vbNormalFocus
err:
    If err.Number = 53 Then
        MsgBox "File Not Found", vbExclamation, "Payroll"
    End If
    Exit Sub
End Sub

Private Sub mnuPay_Click()
'    rptPay.Show
cryrptpay.Show
End Sub

Private Sub mnuPFEsi_Click()
frmPFESI.Show
End Sub

Private Sub mnurangeesi_Click()
frmESI.Show
End Sub

Private Sub mnurangepf_Click()
frmPF.Show
End Sub

Private Sub mnuReports_Click()
 '   frmreport.Show
End Sub

Private Sub mnuresempDetails_Click()
    frmResEmpMaster.Show
End Sub

Private Sub mnuResLoanMaster_Click()
    frmResLoanMaster.Show
End Sub

Private Sub mnuslippay_Click()
    rptmonthly.Show
End Sub

Private Sub mnuUser_Click()
    'frmUser.Show
    User.Show
End Sub
