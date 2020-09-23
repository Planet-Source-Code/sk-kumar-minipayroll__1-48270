VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form rptmonthly 
   Caption         =   "Monthly PaySlip"
   ClientHeight    =   8025
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   10680
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
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   10680
   WindowState     =   2  'Maximized
   Begin VB.Frame fraselect 
      Caption         =   "Select"
      Height          =   855
      Left            =   240
      TabIndex        =   2
      Top             =   1920
      Width           =   3255
      Begin VB.OptionButton OptAllEmpno 
         Caption         =   "&All"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   615
      End
      Begin VB.OptionButton optSelectEmpno 
         Caption         =   "Par&ticular"
         Height          =   375
         Left            =   1560
         TabIndex        =   13
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame fraview 
      Caption         =   "View"
      Height          =   855
      Left            =   8400
      TabIndex        =   4
      Top             =   1920
      Width           =   3255
      Begin VB.OptionButton OptScreen 
         Caption         =   "&Screen"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton OptPrint 
         Caption         =   "&Print"
         Height          =   255
         Left            =   1920
         TabIndex        =   11
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame fraEmpno 
      Caption         =   "EmpNo."
      Height          =   855
      Left            =   4200
      TabIndex        =   3
      Top             =   1920
      Visible         =   0   'False
      Width           =   3615
      Begin VB.ComboBox cboEmpno 
         Height          =   345
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.ComboBox cbobranch 
      Height          =   345
      Left            =   4920
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1440
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   10200
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   7560
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Clear"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   9240
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   7560
      Width           =   735
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4470
      Left            =   120
      TabIndex        =   5
      Top             =   3000
      Width           =   11775
   End
   Begin MSMask.MaskEdBox mskDate 
      Height          =   375
      Left            =   9000
      TabIndex        =   0
      Top             =   1080
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      _Version        =   327680
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "mm/dd/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label lblCompanyName 
      AutoSize        =   -1  'True
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   4200
      TabIndex        =   17
      Top             =   720
      Width           =   45
   End
   Begin VB.Label Label3 
      Caption         =   "Company Name"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   2280
      TabIndex        =   16
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Branch"
      Height          =   225
      Left            =   3960
      TabIndex        =   10
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Pay Date"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8175
      TabIndex        =   9
      Top             =   1080
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Monthly PaySlip"
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
      Left            =   4320
      TabIndex        =   8
      Top             =   120
      Width           =   1665
   End
End
Attribute VB_Name = "rptmonthly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim total
Dim NetPay
Dim totalEarn
Dim totalDet
Dim others
Dim EarningSal
Dim EarningDA
Dim i
Dim Msg As String
Dim NumConvert
Dim PreMonth As Integer
Dim PreYear As Integer
Private Sub cbobranch_Click()
If cboBranch <> "" Then
    rs.Open "Select empno from admin Where CompanyName='" & lblCompanyName.Caption & "' and branchcode='" & cboBranch & "' order by empno", cnn, adOpenKeyset, adLockOptimistic
    cboEmpno.clear
    While Not rs.EOF
         cboEmpno.AddItem SubName & "/" & Format(rs.Fields(0), "000")
        rs.MoveNext
    Wend
End If
rs.Close
End Sub
Private Sub cboEmpno_Click()
If cboEmpno <> "" And optScreen.value = True And mskDate.Text <> "__/__/____" Then

    SelectScreen
End If
End Sub
Private Sub Command1_Click()
OptAllEmpno.value = False
OptPrint.value = False
optScreen.value = False
optSelectEmpno.value = False
mskDate.Text = "__/__/____"
List1.clear
fraEmpno.Visible = False
Cancel = True
mskDate.SetFocus
End Sub
Private Sub Command2_Click()
Unload Me
End Sub
Private Sub Form_Activate()
'rptmonthly.Left = 0
'rptmonthly.Top = 135
End Sub
Private Sub Form_Load()
rptmonthly.Icon = LoadPicture(App.Path & "\Jamaica my computer.ico")
lblCompanyName.Caption = CompanyName
rs.Open "Select BranchCode from Branch Where CompanyName='" & lblCompanyName.Caption & "'", cnn, adOpenKeyset, adLockOptimistic
cboBranch.clear
While Not rs.EOF
    cboBranch.AddItem rs.Fields(0)
    rs.MoveNext
Wend
rs.Close
End Sub
Private Sub OptAllEmpno_GotFocus()
fraEmpno.Visible = False
End Sub
Private Sub OptPrint_Click()
If cboEmpno <> "" Then
    SelectPrint
ElseIf OptAllEmpno.value = True Then
    AllEmpnoPrint
End If
End Sub
Public Sub SelectScreen()
If Format(mskDate, "mm") <> 1 Then
    PreMonth = Format(mskDate, "mm") - 1
    PreYear = Format(mskDate, "yyyy")
Else
    PreMonth = 12
    PreYear = Format(mskDate, "yyyy") - 1
End If

    rs.Open "Select * from daughters where CompanyName='" & lblCompanyName.Caption & "' and empno=" & Mid(cboEmpno, InStr(1, cboEmpno, "/") + 1) & " and  month(datemon)=" & PreMonth & " and Year(datemon)=" & PreYear, cnn, adOpenKeyset, adLockOptimistic
    List1.clear
   
   If rs.EOF = False Then
    EarningSal = Format((rs.Fields("Basic") * rs.Fields("NoDays")) / rs.Fields("totdays"), "###0.00")
    
    EarningDA = Format((rs.Fields("DA") * rs.Fields("NoDays")) / rs.Fields("totdays"), "####0.00")
    
    total = Format(rs("Basic") + rs("DA"), "####0.00")
    
    totalEarn = Format(CDbl(EarningSal) + CDbl(EarningDA) + rs(8) + rs.Fields("telephone"), "####0.00")
    totalEarn = AccurateCost(CDbl(totalEarn))
    
    totalDet = Format(rs(11) + rs(12) + rs(13) + rs(15) + rs.Fields("advance"), "####0.00")
    
    NetPay = Format(CDbl(totalEarn) - CDbl(totalDet), "####0.00")
    
    'NetPay = AccurateCost(Format(CDbl(NetPay), "0.00"))
    NumConvert = NumToWords(CDbl(NetPay))
    
    rs1.Open "Select * from admin where CompanyName='" & lblCompanyName.Caption & "' and empno=" & Mid(cboEmpno, InStr(1, cboEmpno, "/") + 1), cnn, adOpenKeyset, adLockOptimistic
    
    List1.AddItem ""
    List1.AddItem "Name            :" & UCase(rs.Fields("Ename")) & " [ " & rs.Fields("Empno") & " ] " & Space(50 - Len(rs.Fields("Ename"))) & "Designation     :" & rs.Fields("Designation")
    List1.AddItem "Days Paid For   :" & Format(rs.Fields("NoDays"), "0.0") & Space(57 - Len(rs.Fields("NoDays"))) & "PFNo.           :" & rs1.Fields("pfno")
    rs1.Close
    List1.AddItem ""
    
    List1.AddItem String(45, "-") & Space(5) & "ACTUALS" & Space(5) & String(45, "-")
    List1.AddItem ""
    List1.AddItem Space(10) & "Salary" & Space(20) & "DA" & Space(20) & "Others" & Space(20) & "Total"
        List1.AddItem ""
        List1.AddItem Space(8) & Format(rs.Fields("Basic"), "####00.00") & Space(26 - Len(Format(rs.Fields("Basic"), "####00.00"))) & Format(rs.Fields("DA"), "#00.00") & Space(24 - Len(Format(rs.Fields("DA"), "#00.00"))) & Format(rs.Fields("Others"), "#00.00") & Space(24 - Len(Format(rs.Fields("Others"), "#00.00"))) & total
    List1.AddItem String(108, "-")
    List1.AddItem Space(30) & "EARNINGS" & Space(40) & "DEDUCTIONS"
    List1.AddItem ""
    List1.AddItem Space(1) & "Salary" & Space(10) & "DA" & Space(7) & "OTHRES" & Space(7) & "Total" & Space(10) & "PF" & Space(7) & "ESIC" & Space(7) & "Loan" & Space(7) & "Others" & Space(5) & "Total"
    List1.AddItem ""
    List1.AddItem Space(1) & EarningSal & Space(15 - Len(EarningSal)) & EarningDA & Space(10 - Len(EarningDA)) & Format(rs.Fields("Others") + rs.Fields("telephone"), "#00.00") & Space(12 - Len(Format(rs.Fields("Others"), "#00.00"))) & totalEarn & Space(14 - Len(totalEarn)) & Format(rs.Fields("epf"), "####0.00") & Space(10 - Len(Format(rs.Fields("epf"), "####0.00"))) & Format(rs.Fields("Esic"), "####0.00") & Space(12 - Len(Format(rs.Fields("Esic"), "####0.00"))) & Format(rs.Fields("loandecl"), "####0.00") & Space(11 - Len(Format(rs.Fields("loandecl"), "####0.00"))) & Format(rs.Fields("othersded") + rs.Fields("advance"), "####0.00") & Space(11 - Len(Format(rs.Fields("othersded"), "###0.00"))) & totalDet
    List1.AddItem ""
    List1.AddItem Space(1) & "NetPay :" & " " & NetPay
    List1.AddItem ""
    List1.AddItem Space(1) & NumConvert
    
    List1.AddItem ""
    List1.AddItem ""
    List1.AddItem Space(15) & "Sign.Of Employee" & Space(25) & "Checked By" & Space(25) & "Manager"
    rs.Close
    
Else
MsgBox "No Records", vbExclamation, "Payroll"
rs.Close
End If
End Sub
Public Sub SelectPrint()
Dim tempEPFNo As String
Msg = MsgBox("Are you take Print", vbExclamation + vbYesNo, "Payroll")
If Msg = 6 Then
    On Error GoTo err
    If Format(mskDate, "mm") <> 1 Then
        PreMonth = Format(mskDate, "mm") - 1
        PreYear = Format(mskDate, "yyyy")
    Else
        PreMonth = 12
        PreYear = Format(mskDate, "yyyy") - 1
    End If

    rs.Open "Select * from daughters where CompanyName='" & lblCompanyName.Caption & "' and empno=" & Mid(cboEmpno, InStr(1, cboEmpno, "/") + 1) & " and  month(datemon)=" & PreMonth & " and Year(datemon)=" & PreYear, cnn, adOpenKeyset, adLockOptimistic
    
    If rs.EOF = False Then
    
        rs1.Open "Select * from admin where CompanyName='" & lblCompanyName.Caption & "' and empno=" & Mid(cboEmpno, InStr(1, cboEmpno, "/") + 1), cnn, adOpenKeyset, adLockOptimistic
           tempEPFNo = rs1.Fields("pfno")
        rs1.Close
        'Printer.NewPage
        Printer.Orientation = 1
        Printer.Height = 1175
        Printer.Font.Name = "Courier New"
        Printer.FontBold = True
        Printer.Print ""
        Printer.Print ""
        Printer.Font.Size = 14
        Printer.Print Space(20) & CompanyName
        Printer.Print ""
        Printer.Font.Size = 12
        Printer.Print Space(22) & "Salary For The Month Of " & MonthName(Month(PreMonth)) & "-" & year(PreYear)
        Printer.Font.Size = 10
        Printer.Print Space(20) & "Under rule 27(2) of the Mini Wages Madras Rules 1953"
        Printer.Print ""
        Printer.Print ""
        EarningSal = Format((rs.Fields("Basic") * rs.Fields("NoDays")) / rs.Fields("totdays"), "####0.00")
        
        EarningDA = Format((rs.Fields("DA") * rs.Fields("NoDays")) / rs.Fields("totdays"), "####0.00")
        
        total = Format(rs("Basic") + rs("DA"), "####0.00")
        
        totalEarn = Format(CDbl(EarningSal) + CDbl(EarningDA) + rs(8) + rs("telephone"), "####0.00")
        totalEarn = AccurateCost(CDbl(totalEarn))
        
        totalDet = Format(rs(11) + rs(12) + rs(13) + rs(15) + rs.Fields("advance"), "####0.00")
        
        NetPay = Format(CDbl(totalEarn) - CDbl(totalDet), "####0.00")
        'NetPay = AccurateCost(Format(CDbl(NetPay), ".00"))
        Printer.Font.Bold = True
        
        Printer.Print Space(5) & "Emp Name       :" & UCase(rs.Fields("eName")) & " [ " & rs.Fields("empno") & " ] " & Space(38 - Len(UCase(rs.Fields("eName")))) & "Designation : " & rs.Fields("Designation")
        Printer.Print Space(5) & "Days Paid For :" & Format(rs.Fields("Nodays"), "0.0") & Space(47 - Len(Format(rs.Fields("Nodays"), "0.0"))) & "EPFNO       :" & tempEPFNo
        Printer.Print String(40, "-") & Space(5) & "ACTUALS" & Space(5) & String(40, "-")
        Printer.Print ""
        Printer.Print Space(10) & "Salary" & Space(20) & "DA" & Space(20) & "Others" & Space(18) & "Total"
        Printer.Print ""
        Printer.Print Space(8) & Format(rs.Fields("Basic"), "####00.00") & Space(26 - Len(Format(rs.Fields("Basic"), "####00.00"))) & Format(rs.Fields("DA"), "#00.00") & Space(24 - Len(Format(rs.Fields("DA"), "#00.00"))) & Format(rs.Fields("Others"), "#00.00") & Space(24 - Len(Format(rs.Fields("Others"), "#00.00"))) & total
        Printer.Print String(98, "-")
        Printer.Print Space(27) & "EARNINGS" & Space(30) & "DEDUCTIONS"
        Printer.Print ""
        Printer.Print Space(5) & "Salary" & Space(7) & "DA" & Space(6) & "OTHRES" & Space(5) & "Total" & Space(7) & "PF" & Space(5) & "ESIC" & Space(7) & "Loan" & Space(5) & "Others" & Space(5) & "Total"
        Printer.Print ""
        Printer.Print Space(4) & EarningSal & Space(13 - Len(EarningSal)) & EarningDA & Space(9 - Len(EarningDA)) & Format(rs.Fields("Others"), "#00.00") & Space(11 - Len(Format(rs.Fields("Others"), "#00.00"))) & totalEarn & Space(11 - Len(totalEarn)) & Format(rs.Fields("epf"), "####0.00") & Space(8 - Len(Format(rs.Fields("epf"), "####0.00"))) & Format(rs.Fields("Esic"), "####0.00") & Space(11 - Len(Format(rs.Fields("Esic"), "####0.00"))) & Format(rs.Fields("loandecl"), "####0.00") & Space(9 - Len(Format(rs.Fields("loandecl"), "####0.00"))) & Format(rs.Fields("othersded"), "####0.00") & Space(10 - Len(Format(rs.Fields("othersded"), "###0.00"))) & totalDet
        Printer.Print ""
        
        Printer.Print Space(30) & "EARNINGS" & Space(40) & "DEDUCTIONS"
        Printer.Print ""
        Printer.Print Space(5) & "Salary" & Space(10) & "DA" & Space(7) & "OTHRES" & Space(7) & "Total" & Space(10) & "PF" & Space(7) & "ESIC" & Space(7) & "Loan" & Space(7) & "Others" & Space(7) & "Total"
        Printer.Print ""
        Printer.Print Space(4) & EarningSal & Space(15 - Len(EarningSal)) & EarningDA & Space(10 - Len(EarningDA)) & Format(rs.Fields("Others"), "#00.00") & Space(13 - Len(Format(rs.Fields("Others"), "#00.00"))) & totalEarn & Space(14 - Len(totalEarn)) & Format(rs.Fields("epf"), "####0.00") & Space(10 - Len(Format(rs.Fields("epf"), "####0.00"))) & Format(rs.Fields("Esic"), "####0.00") & Space(12 - Len(Format(rs.Fields("Esic"), "####0.00"))) & Format(rs.Fields("loandecl"), "####0.00") & Space(11 - Len(Format(rs.Fields("loandecl"), "####0.00"))) & Format(rs.Fields("othersded"), "####0.00") & Space(13 - Len(Format(rs.Fields("othersded"), "###0.00"))) & totalDet
        Printer.Print ""
        Printer.Font.Size = 12
        Printer.Font.Bold = True
        Printer.Print Space(4) & "NetPay :" & " " & NetPay
        Printer.Print ""
        NumConvert = NumToWords(CDbl(NetPay))
        Printer.Print Space(4) & NumConvert
        Printer.Print ""
        Printer.Print ""
        Printer.Print ""
        Printer.Print ""
        Printer.Font.Bold = False
        Printer.Font.Size = 10
        Printer.Print Space(15) & "Sign.Of Employee" & Space(20) & "Checked By" & Space(20) & "Manager"
        Printer.EndDoc
        
    Else
        MsgBox "No Records", vbExclamation, "Payroll"
        rs.Close
        Exit Sub
    End If
rs.Close
err:
        If err.Number = 482 Or err.Number = 484 Then
            rs.Close
            MsgBox "Make Sure The printer Is Ready", vbExclamation, "Payroll"
        End If
        Exit Sub
End If
End Sub
Private Sub mskDate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If mskDate.Text = "__/__/____" Then
        
        mskDate = Format(Date, "mm/dd/yyyy")
        Cancel = True
        
        
        Exit Sub
    End If
    
    If CheckMonth(mskDate) = False Then
        MsgBox "Invalid Month", vbExclamation, "Payroll"
        Cancel = True
        mskDate = "__/__/____"
        mskDate.SetFocus
        Exit Sub
    End If
    If CheckDate(mskDate) = False Then
        MsgBox "Invalid Date", vbExclamation, "Payroll"
        Cancel = True
        mskDate = "__/__/____"
        mskDate.SetFocus
        Exit Sub
    End If
End If
End Sub
Private Sub mskDate_Validate(Cancel As Boolean)
    If mskDate.Text = "__/__/____" Then
        MsgBox "Enter Pay Date", vbExclamation, "Payroll"
        Cancel = True
        mskDate.SetFocus
        Exit Sub
    End If
    
    If CheckMonth(mskDate) = False Then
        MsgBox "Invalid Month", vbExclamation, "Payroll"
        Cancel = True
        mskDate = "__/__/____"
        mskDate.SetFocus
        Exit Sub
    End If
    If CheckDate(mskDate) = False Then
        MsgBox "Invalid Date", vbExclamation, "Payroll"
        Cancel = True
        mskDate = "__/__/____"
        mskDate.SetFocus
        Exit Sub
    End If
End Sub
Private Sub OptScreen_GotFocus()
If cboEmpno <> "" And optSelectEmpno.value = True Then
    SelectScreen
ElseIf OptAllEmpno.value = True Then
    SelectAllScreen
End If
End Sub
Private Sub optSelectEmpno_GotFocus()
fraEmpno.Visible = True
End Sub
Public Sub SelectAllScreen()
List1.clear

If Format(mskDate, "mm") <> 1 Then
    PreMonth = Format(mskDate, "mm") - 1
    PreYear = Format(mskDate, "yyyy")
Else
    PreMonth = 12
    PreYear = Format(mskDate, "yyyy") - 1
End If

rs1.Open "Select Empno from admin where CompanyName='" & lblCompanyName.Caption & "' and Branchcode='" & cboBranch & "' order by empno", cnn, adOpenKeyset, adLockOptimistic
While Not rs1.EOF
    
    rs.Open "Select * from  daughters where CompanyName='" & lblCompanyName.Caption & "' and empno=" & rs1.Fields(0) & " and  month(datemon)=" & PreMonth & " and Year(datemon)=" & PreYear, cnn, adOpenKeyset, adLockOptimistic
    
    
    
    If rs.EOF = False Then
        EarningSal = Format((rs.Fields("Basic") * rs.Fields("NoDays")) / rs.Fields("totdays"), "####0.00")
        
        EarningDA = Format((rs.Fields("DA") * rs.Fields("NoDays")) / rs.Fields("totdays"), "####0.00")
        
        total = Format(rs("Basic") + rs("DA"), "####0.00")
        
        totalEarn = Format(CDbl(EarningSal) + CDbl(EarningDA) + rs(8) + rs.Fields("telephone"), "####0.00")
        totalEarn = AccurateCost(CDbl(totalEarn))
        
        totalDet = Format(rs(11) + rs(12) + rs(13) + rs(15) + rs.Fields("advance"), "####0.00")
        
        NetPay = Round(Format(CDbl(totalEarn) - CDbl(totalDet), "####0.00"), 2)
        'NetPay = AccurateCost(Format(CDbl(NetPay), "0.00"))
        NumConvert = NumToWords(CDbl(NetPay))
        
        rs2.Open "Select * from admin where CompanyName='" & lblCompanyName.Caption & "' and empno=" & rs1.Fields(0), cnn, adOpenKeyset, adLockOptimistic
        
        List1.AddItem ""
        List1.AddItem "Name            :" & UCase(rs.Fields("Ename")) & " [ " & rs.Fields("Empno") & " ] " & Space(50 - Len(rs.Fields("Ename"))) & "Designation     :" & rs.Fields("Designation")
        List1.AddItem "Days Paid For   :" & Format(rs.Fields("NoDays"), "0.0") & Space(57 - Len(rs.Fields("NoDays"))) & "PFNo.           :" & rs2.Fields("pfno")
        rs2.Close
        List1.AddItem ""
        List1.AddItem String(45, "-") & Space(5) & "ACTUALS" & Space(5) & String(45, "-")
        List1.AddItem ""
        List1.AddItem Space(10) & "Salary" & Space(20) & "DA" & Space(20) & "Others" & Space(20) & "Total"
        List1.AddItem ""
        List1.AddItem Space(8) & Format(rs.Fields("Basic"), "####00.00") & Space(26 - Len(Format(rs.Fields("Basic"), "####00.00"))) & Format(rs.Fields("DA"), "#00.00") & Space(24 - Len(Format(rs.Fields("DA"), "#00.00"))) & Format(rs.Fields("Others"), "#00.00") & Space(24 - Len(Format(rs.Fields("Others"), "#00.00"))) & total
        List1.AddItem String(108, "-")
        List1.AddItem Space(30) & "EARNINGS" & Space(40) & "DEDUCTIONS"
        List1.AddItem ""
        List1.AddItem Space(1) & "Salary" & Space(10) & "DA" & Space(7) & "OTHRES" & Space(7) & "Total" & Space(10) & "PF" & Space(7) & "ESIC" & Space(7) & "Loan" & Space(6) & "Others" & Space(5) & "Total"
        List1.AddItem ""
        List1.AddItem Space(1) & EarningSal & Space(15 - Len(EarningSal)) & EarningDA & Space(10 - Len(EarningDA)) & Format(rs.Fields("Others") + rs.Fields("telephone"), "#00.00") & Space(12 - Len(Format(rs.Fields("Others"), "#00.00"))) & totalEarn & Space(14 - Len(totalEarn)) & Format(rs.Fields("epf"), "####0.00") & Space(10 - Len(Format(rs.Fields("epf"), "####0.00"))) & Format(rs.Fields("Esic"), "####0.00") & Space(12 - Len(Format(rs.Fields("Esic"), "####0.00"))) & Format(rs.Fields("loandecl"), "####0.00") & Space(10 - Len(Format(rs.Fields("loandecl"), "####0.00"))) & Format(rs.Fields("othersded") + rs.Fields("advance"), "####0.00") & Space(11 - Len(Format(rs.Fields("othersded"), "###0.00"))) & totalDet
        List1.AddItem ""
        List1.AddItem Space(1) & "NetPay :" & " " & NetPay
        List1.AddItem ""
        List1.AddItem Space(1) & NumConvert
        List1.AddItem ""
        List1.AddItem ""
        List1.AddItem Space(15) & "Sign.Of Employee" & Space(25) & "Checked By" & Space(25) & "Manager"
        rs.Close
        
        
Else
    rs.Close
End If

rs1.MoveNext
Wend
rs1.Close
End Sub
Public Sub AllEmpnoPrint()
Dim tempEPFNo As String
Msg = MsgBox("Are you take Print", vbExclamation + vbYesNo, "Payroll")
If Msg = 6 Then
On Error GoTo err
'Printer.NewPage
Printer.Orientation = 1
'Printer.Width = 800
'Printer.Height = 1200
Printer.Font.Name = "Courier New"
    rs1.Open "Select Empno from admin where CompanyName='" & lblCompanyName.Caption & "' and Branchcode='" & cboBranch & "' order by empno", cnn, adOpenKeyset, adLockOptimistic
    j = 0
    
    If Format(mskDate, "mm") <> 1 Then
        PreMonth = Format(mskDate, "mm") - 1
        PreYear = Format(mskDate, "yyyy")
    Else
        PreMonth = 12
        PreYear = Format(mskDate, "yyyy") - 1
    End If
    While Not rs1.EOF
        
        rs.Open "Select * from daughters where CompanyName='" & lblCompanyName.Caption & "' and empno=" & rs1.Fields(0) & " and  month(datemon)=" & PreMonth & " and Year(datemon)=" & PreYear, cnn, adOpenKeyset, adLockOptimistic
    
        If rs.EOF = False Then
            j = j + 1
            rs2.Open "Select * from admin where CompanyName='" & lblCompanyName.Caption & "' and empno=" & rs1.Fields(0), cnn, adOpenKeyset, adLockOptimistic
            tempEPFNo = rs2.Fields("pfno")
            rs2.Close
            
            Printer.FontBold = True
            Printer.Print ""
            Printer.Print ""
            Printer.Font.Size = 14
            Printer.Print Space(20) & CompanyName
            Printer.Font.Size = 12
            Printer.Print ""
            Printer.Print Space(22) & "Salary For The Month Of " & MonthName(PreMonth) & " " & year(Month(PreMonth) & "- 01" & "-" & PreYear)
            Printer.FontBold = False
            Printer.Font.Size = 10
            Printer.Print Space(20) & "Under rule 27(2) of the Mini Wages Madras Rules 1953"
            Printer.Print
            Printer.Print ""

            
            EarningSal = Format((rs.Fields("Basic") * rs.Fields("NoDays")) / rs.Fields("totdays"), "####0.00")
            
            EarningDA = Format((rs.Fields("DA") * rs.Fields("NoDays")) / rs.Fields("totdays"), "####0.00")
            
            total = Format(rs("Basic") + rs("DA"), "####0.00")
            
            totalEarn = Format(CDbl(EarningSal) + CDbl(EarningDA) + rs(8) + rs.Fields("Telephone"), "####0.00")
            totalEarn = AccurateCost(CDbl(totalEarn))
            
            totalDet = Format(rs(11) + rs(12) + rs(13) + rs(15) + rs.Fields("Advance"), "####0.00")
            
            NetPay = Format(CDbl(totalEarn) - CDbl(totalDet), "####0.00")
            'NetPay = AccurateCost(Format(CDbl(NetPay), "0.00"))
            Printer.FontBold = True
            Printer.Print Space(5) & "Emp Name       :" & UCase(rs.Fields("eName")) & " [ " & rs.Fields("empno") & " ] " & Space(38 - Len(UCase(rs.Fields("eName")))) & "Designation : " & rs.Fields("Designation")
            Printer.FontBold = False
            Printer.Print Space(5) & "Days Paid For :" & Format(rs.Fields("Nodays"), "0.0") & Space(47 - Len(Format(rs.Fields("Nodays"), "0.0"))) & "EPFNO       :" & tempEPFNo
            Printer.Print String(40, "-") & Space(5) & "ACTUALS" & Space(5) & String(40, "-")
            Printer.Print ""
            Printer.Print Space(10) & "Salary" & Space(20) & "DA" & Space(20) & "Others" & Space(18) & "Total"
            Printer.Print ""
            Printer.Print Space(8) & Format(rs.Fields("Basic"), "####00.00") & Space(26 - Len(Format(rs.Fields("Basic"), "####00.00"))) & Format(rs.Fields("DA"), "#00.00") & Space(24 - Len(Format(rs.Fields("DA"), "#00.00"))) & Format(rs.Fields("Others"), "#00.00") & Space(24 - Len(Format(rs.Fields("Others"), "#00.00"))) & total
            Printer.Print String(98, "-")
            Printer.Print Space(27) & "EARNINGS" & Space(30) & "DEDUCTIONS"
            Printer.Print ""
            Printer.Print Space(5) & "Salary" & Space(7) & "DA" & Space(6) & "OTHRES" & Space(5) & "Total" & Space(7) & "PF" & Space(5) & "ESIC" & Space(7) & "Loan" & Space(5) & "Others" & Space(5) & "Total"
            Printer.Print ""
            Printer.Print Space(4) & EarningSal & Space(13 - Len(EarningSal)) & EarningDA & Space(9 - Len(EarningDA)) & Format(rs.Fields("Others"), "#0.00") & Space(11 - Len(Format(rs.Fields("Others"), "#0.00"))) & totalEarn & Space(11 - Len(totalEarn)) & Format(rs.Fields("epf"), "####0.00") & Space(8 - Len(Format(rs.Fields("epf"), "####0.00"))) & Format(rs.Fields("Esic"), "####0.00") & Space(11 - Len(Format(rs.Fields("Esic"), "####0.00"))) & Format(rs.Fields("loandecl"), "####0.00") & Space(9 - Len(Format(rs.Fields("loandecl"), "####0.00"))) & Format(rs.Fields("othersded"), "####0.00") & Space(10 - Len(Format(rs.Fields("othersded"), "###0.00"))) & totalDet
            Printer.Print ""
            Printer.Font.Size = 12
            Printer.FontBold = True
            Printer.Print Space(4) & "NetPay :" & " " & NetPay
            Printer.Print ""
            NumConvert = NumToWords(CDbl(NetPay))
            Printer.Print Space(4) & NumConvert
            Printer.Print ""
            Printer.Print ""
            Printer.Print ""
            Printer.Print ""
            Printer.FontBold = False
            Printer.Font.Size = 10
            Printer.Print Space(15) & "Sign.Of Employee" & Space(20) & "Checked By" & Space(20) & "Manager"
            Printer.Print ""
            Printer.Print ""
            Printer.Print ""
            
            If j = 2 Then
                j = 0
                'Printer.Print ""
                'Printer.Print ""
                'Printer.Print String(50, "-") & "End Of Page" & String(50, "-")
                 Printer.NewPage
                 Printer.Orientation = 1
                 'Printer.Width = 800
                 'Printer.Height = 1200
                 Printer.Font.Name = "Courier New"
                
            End If
            Printer.FontBold = False
'            Printer.EndDoc
        
    Else
'        MsgBox "No Records", vbInformation, "Payroll"
 '       rs.Close
  '      Exit Sub
    End If
    rs.Close
rs1.MoveNext
Wend
Printer.EndDoc
rs1.Close

err:
        If err.Number = 482 Or err.Number = 484 Then
        
            
            MsgBox "Make Sure The printer Is Ready", vbExclamation, "Payroll"
        End If
        Exit Sub
End If
End Sub
Public Function AccurateCost(Text As Double)
'If Right(Format(Text, ".00"), 1) > 5 Then
'    AccurateCost = Format((Round(Text, 1)), ".00")
'ElseIf Right(Format(Text, ".00"), 1) < 5 And Right(Format(Text, ".00"), 1) <> 0 Then
'    AccurateCost = Format((Round(Text, 2)), ".05")
'ElseIf Right(Format(Text, ".00"), 1) = 5 Or Right(Format(Text, ".00"), 1) = 0 Then
'    AccurateCost = Format(Text, ".00")
'
'End If
If Right(Format(Text, ".00"), 1) = 5 Then
    AccurateCost = Format(Text + 0.05, "0.00")
Else
    AccurateCost = Format((Round(Text, 1)), "0.00")
End If
End Function
