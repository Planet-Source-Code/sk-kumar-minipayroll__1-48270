VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form cryrptpay 
   Caption         =   "Pay Report"
   ClientHeight    =   6765
   ClientLeft      =   150
   ClientTop       =   765
   ClientWidth     =   11880
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
   ScaleHeight     =   6765
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Rep1 
      Height          =   480
      Left            =   2760
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   16
      Top             =   6360
      Width           =   1200
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10680
      TabIndex        =   11
      Top             =   6120
      Width           =   735
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Clear"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9840
      TabIndex        =   10
      Top             =   6120
      Width           =   735
   End
   Begin VB.Frame fraBranch 
      Caption         =   "Branch"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3360
      TabIndex        =   8
      Top             =   1440
      Width           =   3855
      Begin VB.ComboBox cboBranch 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         TabIndex        =   9
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "View"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8520
      TabIndex        =   5
      Top             =   1440
      Width           =   3135
      Begin VB.OptionButton optPrinter 
         Caption         =   "&Printer"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton optScreen 
         Caption         =   "&Screen"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   3135
      Begin VB.OptionButton optAll 
         Caption         =   "&All"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton optPart 
         Caption         =   "Par&ticular"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   3240
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   4895
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSMask.MaskEdBox mskDate 
      Height          =   375
      Left            =   9480
      TabIndex        =   1
      Top             =   960
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   661
      _Version        =   393216
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Monthly Salary Report"
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
      Left            =   4515
      TabIndex        =   15
      Top             =   240
      Width           =   2325
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
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
      Height          =   375
      Left            =   7800
      TabIndex        =   14
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Company Name"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   3120
      TabIndex        =   13
      Top             =   720
      Width           =   1260
   End
   Begin VB.Label lblCompanyName 
      AutoSize        =   -1  'True
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   4800
      TabIndex        =   12
      Top             =   720
      Width           =   45
   End
End
Attribute VB_Name = "cryrptpay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***Dim cnn As New ADODB.Connection
'***Dim rs As New ADODB.Recordset
'***Dim rs1 As New ADODB.Recordset
Dim lt As ListItem
Dim Msg As String

Dim TotCompActBasci As Double
Dim TotCompActDA As Double
Dim TotCompActTotal As Double

Dim TotCompEarnBasci As Double
Dim TotCompEarnDA As Double
Dim TotCompEarnOthers As Double
Dim TotCompEarnTotal As Double

Dim TotCompDedPF As Double
Dim TotCompDedEsic As Double
Dim TotCompDedLoan As Double
Dim TotCompDedOthers As Double
Dim TotCompDedTotal As Double

Dim CompTotal As Double
Dim PreMonth As Integer
Dim PreYear As Integer
Dim PreDate As Date
Private Sub cbobranch_Click()
'    fraBranch.Visible = False
End Sub
Private Sub cboBranch_Validate(Cancel As Boolean)
    If optPart Then
        If cboBranch.ListIndex = -1 Then
            MsgBox "Select Branch To View/Print", vbExclamation, "Payroll"
            Cancel = True
            SendKeys "{Home}+{End}"
            Exit Sub
        End If
    End If
End Sub
Private Sub CmdCancel_Click()
    clear
    ListView1.ListItems.clear
    fraBranch.Visible = False
    Cancel = True
    mskDate.SetFocus
    
End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub Form_Activate()
'cryrptpay.Left = 0
'cryrptpay.Top = 0
mskDate.SetFocus
ListView1.ColumnHeaders.clear
ListView1.ListItems.clear

ListView1.ColumnHeaders.Add , , "SNo", 500
ListView1.ColumnHeaders.Add , , "Employee Name"
ListView1.ColumnHeaders.Add , , "Days"
ListView1.ColumnHeaders.Add , , "Basic(A)"
ListView1.ColumnHeaders.Add , , "DA(A)"
ListView1.ColumnHeaders.Add , , "Total(A)"
ListView1.ColumnHeaders.Add , , "Basic(E)"
ListView1.ColumnHeaders.Add , , "DA(E)"
ListView1.ColumnHeaders.Add , , "Others"
ListView1.ColumnHeaders.Add , , "Total(E)"
ListView1.ColumnHeaders.Add , , "PF"
ListView1.ColumnHeaders.Add , , "ESIC"
ListView1.ColumnHeaders.Add , , "Loan"
ListView1.ColumnHeaders.Add , , "Others"
ListView1.ColumnHeaders.Add , , "Total(D)"
ListView1.ColumnHeaders.Add , , "Netpay"
End Sub
Private Sub Form_Load()
'***cnn.Provider = "Microsoft.Jet.OLEDB.4.0"
'***cnn.Open App.Path & "\moolakadai.mdb"
'lstPay.Font.Bold = True
'lstPay.Font.Size = 10

cryrptpay.Icon = LoadPicture(App.Path & "\Jamaica my computer.ico")
fraBranch.Visible = False
lblCompanyName.Caption = CompanyName
rs.Open "Select * From Branch Where CompanyName='" & lblCompanyName.Caption & "'", cnn, adOpenKeyset, adLockOptimistic
While Not rs.EOF
    cboBranch.AddItem rs(0)
    rs.MoveNext
Wend
rs.Close
End Sub
Private Sub Form_Unload(Cancel As Integer)
   '*** cnn.Close
End Sub
Private Sub mskDate_Change()
    optAll.value = False
    optPart.value = False
    optScreen.value = False
    optPrinter.value = False
End Sub
Private Sub mskDate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If mskDate.Text = "__/__/____" Then
        'MsgBox "Enter Pay Date", vbInformation, "Payroll"
        mskDate = Format(Date, "mm/dd/yyyy")
        Cancel = True
        'mskDate.SetFocus
        
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
Private Sub optAll_Click()
    fraBranch.Visible = False
End Sub
Private Sub optPart_Click()
    fraBranch.Visible = True
    cboBranch.SetFocus
   ' cboBranch.TabIndex = 3
End Sub
Private Sub optprinter_Click()
    If optAll Then
        PrintAll
    End If
    If optPart Then
        PrintPart
    End If
End Sub
Private Sub optScreen_Click()
    If optAll Then
        ScreenAll
    End If
    If optPart Then
      ScreenPart
     
    End If

End Sub
Private Sub PrintAll()
If Format(mskDate, "mm") <> 1 Then
        PreMonth = Format(mskDate, "mm") - 1
        PreYear = Format(mskDate, "yyyy")
    Else
        PreMonth = 12
        PreYear = Format(mskDate, "yyyy") - 1
    End If
    PreDate = Format(PreMonth & "-" & "01-" & PreYear, "mm/dd/yyyy")

rs.Open "Select Count(*) from Daughters Where CompanyName='" & lblCompanyName.Caption & "' and  DateMon=#" & PreDate & "#", cnn, adOpenKeyset, adLockOptimistic
If rs.Fields(0) = 0 Then
    MsgBox "No Records ", vbExclamation, "Payroll"
    rs.Close
    Exit Sub
End If

rs.Close
Dim Mon
Mon = Format(PreDate, "MMMM")
cnn.Execute "Delete From TempMonth"
cnn.Execute "Insert into TempMonth values('" & lblCompanyName.Caption & "','" & Mon & "'," & PreYear & " )"
    
Msg = MsgBox("Are you take Print", vbExclamation + vbYesNo, "Payroll")
If Msg = 6 Then
 
   
   rep1.ReportFileName = App.Path & "\PayReport.rpt"
   rep1.SelectionFormula = "{Daughters.Datemon}=date(" & year(PreDate) & "," & Month(PreDate) & "," & Day(PreDate) & ") and {Daughters.CompanyName}='" & lblCompanyName.Caption & "'"
   rep1.Action = True
   
End If
    
    
End Sub
Private Sub ScreenAllList()
'         List Control
    lstPay.clear
    lstPay.AddItem String(200, "-")
    lstPay.AddItem "Employee" & Space(3) & Space(5) & "Days" & Space(6) & "Total" & Space(5) & "Total" & Space(5) & "PF" & Space(9) & "ESIC" & Space(7) & "Loan" & Space(7) & "Total" & Space(6) & "Net Pay"
    lstPay.AddItem Space(26) & "(Actuals)" & Space(1) & "(Earnings)" & Space(33) & "(Deduction)"
    lstPay.AddItem String(200, "-")
    sql = "Select * from daughters where CompanyName='" & lblCompanyName.Caption & "' and month(datemon)=" & Format(mskDate, "mm") & " and Year(datemon)=" & Format(mskDate, "yyyy")
    'sql = "Select *From Daughters where datemon=#" & Format(mskDate, "mm/dd/yyyy") & "#" '& Format("datemon", "mm/yyyy") & "= '" & mskDate & "#"
    'MsgBox sql
    rs.Open sql, cnn, adOpenKeyset, adLockOptimistic
    If Not (rs.EOF Or rs.BOF) Then
        For i = 0 To rs.RecordCount - 1
            sql = "Select Basic,Father_Husband from admin where CompanyName='" & lblCompanyName.Caption & "' and empno =" & rs(0)
            rs1.Open sql, cnn, adOpenKeyset, adLockOptimistic
            If Not (rs1.EOF Or rs1.BOF) Then
                Dim total
                Dim NetPay
                Dim totalEarn
                Dim totalDet
                Dim Ename
                Dim Father
                Dim days
                Dim Abasic
                Dim Ada
                Dim Atot
                Dim Ebasic
                Dim Eda
                Dim Eothers
                Dim Etot
                Dim Epf
                Dim Eesic
                Dim Loan
                Dim Dtot
                Dim Net
                Ename = UCase(rs(1))
                
               ' Father = Space(25 - Len(Trim(rs(1)))) & rs1(1)
              ' Father = Space(18 - Len(Trim(rs(1)))) & UCase(rs1(1))
               days = Space(16 - Len(Trim(rs(1)))) & rs(10)
               'Abasic = Space(5 - Len(Trim(rs(10)))) & Format(rs(5), "####0.00")
               'Ada = Space(9 - Len(Trim(Format(rs(5), "####0.00")))) & Format(rs(7), "####0.00")
               total = Format(rs(5) + rs(7), "####0.00")
               totalEarn = Format(rs(5) + rs(7) + rs(8) + rs(16), "####0.00")
               totalDet = Format(rs(11) + rs(12) + rs(13) + rs(15) + rs(14), "####0.00")
               NetPay = Format(totalEarn - totalDet, "####0.00")
               Atot = Space(10 - Len(Trim(rs(10)))) & total
              ' total = Format(rs(5) + rs(7), "####0.00")
              ' totalEarn = Format(rs(5) + rs(7) + rs(8), "####0.00")
              ' totalDet = Format(rs(11) + rs(12) + rs(13) + rs(15), "####0.00")
               'Ebasic = Space(10 - Len(Trim(total))) & Format(rs(5), "####0.00")
               'Eda = Space(9 - Len(Trim(Format(rs(5), "####0.00")))) & Format(rs(7), "####0.00")
              ' Eothers = Space(10 - Len(Trim(total))) & Format(rs(8), "####0.00")
               Etot = Space(10 - Len(Trim(total))) & totalEarn
               Epf = Space(10 - Len(Trim(totalEarn))) & Format(rs(11), "####0.00")
               Eesic = Space(11 - Len(Trim(Format(rs(11), "####0.00")))) & Format(rs(12), "####0.00")
               Loan = Space(11 - Len(Trim(Format(rs(12), "####0.00")))) & Format(rs(12), "####0.00")
               Dtot = Space(11 - Len(Trim(Format(rs(12), "####0.00")))) & totalDet
               Net = Space(11 - Len(Trim(totalDet))) & NetPay
               NetPay = Format(totalEarn - totalDet, "####0.00")
               lstPay.AddItem Ename & days & Atot & Etot & Epf & Eesic & Loan & Dtot & Net
            Else
                MsgBox "No Records", vbExclamation, "Payroll"
            End If
            rs1.Close
            rs.MoveNext
        Next
    Else
        MsgBox "No Records", vbExclamation, "Payroll"
        rs.Close
        Exit Sub
    End If
    rs.Close
End Sub
Private Sub PrintPart()
If Format(mskDate, "mm") <> 1 Then
        PreMonth = Format(mskDate, "mm") - 1
        PreYear = Format(mskDate, "yyyy")
    Else
        PreMonth = 12
        PreYear = Format(mskDate, "yyyy") - 1
    End If
    PreDate = Format(PreMonth & "-" & "01-" & PreYear, "mm/dd/yyyy")

Dim Mon

cnn.Execute "Delete From TempMonth"
Mon = Format(PreDate, "mmmm")

cnn.Execute "Insert into TempMonth values('" & lblCompanyName.Caption & "','" & Mon & "'," & PreYear & ")"
Msg = MsgBox("Are you take Print", vbExclamation + vbYesNo, "Payroll")
If Msg = 6 Then
   rep1.ReportFileName = App.Path & "\PayReport.rpt"
   rep1.SelectionFormula = "{Daughters.Datemon}=date(" & year(PreDate) & "," & Month(PreDate) & "," & Day(PreDate) & ") and {Daughters.CompanyName}='" & lblCompanyName.Caption & "' and {Admin.Branchcode}='" & cboBranch & "'"
   rep1.Action = True
   
End If
End Sub
Public Sub ScreenPart1()


'***********Display of List Control *******

lstPay.clear
lstPay.AddItem Space(50) & "Actual" & Space(20) & "Earning" & Space(20) & "Deduction"
lstPay.AddItem String(170, "-")
lstPay.AddItem "Employee No" & Space(5) & "Employee Name" & Space(5) & "Days" & Space(5) & "Basic" & Space(5) & "DA" & Space(8) & "Total" & Space(5) & "Basic" & Space(5) & "DA" & Space(8) & "Others" & Space(5) & "Total" & Space(5) & "PF" & Space(8) & "ESIC" & Space(5) & "Loan" & Space(5) & "Others" & Space(5) & "Total" & Space(5) & "Net Pay"
lstPay.AddItem String(170, "-")
sql = "Select * from daughters where CompanyName='" & lblCompanyName.Caption & "' and month(datemon)=" & Format(mskDate, "mm") & " and Year(datemon)=" & Format(mskDate, "yyyy")
    'sql = "Select *From Daughters where datemon=#" & Format(mskDate, "mm/dd/yyyy") & "#" '& Format("datemon", "mm/yyyy") & "= '" & mskDate & "#" order by empno"
    rs.Open sql, cnn, adOpenKeyset, adLockOptimistic
    If Not (rs.EOF Or rs.BOF) Then
        For i = 0 To rs.RecordCount - 1
            sql = "Select *from Branch where CompanyName='" & lblCompanyName.Caption & "' and BranchCode='" & cboBranch & "'"
            rs1.Open sql, cnn, adOpenKeyset, adLockOptimistic
            If Not (rs1.EOF Or rs1.BOF) Then
                Branch = rs(5)
            Else
                MsgBox "No Employees", vbExclamation, "Payroll"
                rs.Close
                rs1.Close
                Exit Sub
            End If
            rs1.Close
            sql = ""
            sql = "Select Basic ,Father_Husband,BranchCode from admin where CompanyName='" & lblCompanyName.Caption & "' and branchcode='" & Branch & "'" & " And empno = " & rs(0)
            rs1.Open sql, cnn, adOpenKeyset, adLockOptimistic

           ' clear
           'MsgBox rs(5)
           If Not (rs1.EOF Or rs1.BOF) Then
            Dim total
            Dim NetPay
            Dim totalEarn
            Dim totalDet
            total = Format(rs(5) + rs(7), "####0.00")
            totalEarn = Format(rs(5) + rs(7) + rs(8) + rs(16), "####0.00")
            totalDet = Format(rs(11) + rs(12) + rs(13) + rs(15) + rs(14), "####0.00")
            NetPay = Format(totalEarn - totalDet, "####0.00")
         
            lstPay.AddItem rs(0) & Space(16 - Len(rs(0))) & rs(1) & Space(18 - Len(rs(1))) & rs(10) & Space(9 - Len(rs(10))) & Format(rs(5), "####0.00") & Space(10 - Len(Format(rs(5), "####0.00"))) & Format(rs(7), "####0.00") & Space(10 - Len(Format(rs(7), "####0.00"))) & total & Space(10 - Len(Format(total, "####0.00"))) & Format(rs(5), "####0.00") & Space(10 - Len(Format(rs(5), "####0.00"))) & Format(rs(7), "####0.00") & Space(10 - Len(Format(rs(7), "####0.00"))) & Format(rs(8), "####0.00") & Space(10 - Len(Format(rs(8), "####0.00"))) & totalEarn & Space(10 - Len(totalEarn)) & Format(rs(11), "####0.00") & Space(10 - Len(Format(rs(11), "####0.00"))) & Format(rs(12), "####0.00") & Space(10 - Len(Format(rs(12), "####0.00"))) & Format(rs(13), "####0.00") & Space(10 - Len(Format(rs(13), "####0.00"))) & Format(rs(15), "####0.00") & Space(10 - Len(Format(rs(15), "####0.00"))) & totalDet & Space(10 - Len(totalDet)) & NetPay & Space(10 - Len(NetPay)) & "------------"

            lstPay.AddItem Space(179 - Len(NetPay)) & "|" & Space(12) & "|"
            lstPay.AddItem Space(179 - Len(NetPay)) & "|" & Space(12) & "|"
            lstPay.AddItem Space(179 - Len(NetPay)) & "|" & Space(12) & "|"
            lstPay.AddItem Space(179 - Len(NetPay)) & "|" & Space(12) & "|"
            lstPay.AddItem Space(173) & "------------"
            lstPay.AddItem ""
            lstPay.AddItem ""
            lstPay.AddItem ""
            lstPay.AddItem ""
            'If i = 6 Then
                'lstPay.AddItem String(79, "-") & "End Of Page" & String(79, "-")
               ' printer.NewPage
               ' lstPay.AddItem Space(85) & "STP Services Private Limited"
               ' lstPay.AddItem Space(81) & "Salary For The Month Of " & MonthName(Month(Date)) & " " & year(Date)
               ' lstPay.AddItem String(170, "-")
               ' printer.Print "Employee No" & Space(5) & "Employee Name" & Space(5) & "Days" & Space(5) & "Basic" & Space(5) & "DA" & Space(8) & "Total" & Space(5) & "Basic" & Space(5) & "DA" & Space(8) & "Others" & Space(5) & "Total" & Space(5) & "PF" & Space(8) & "ESIC" & Space(5) & "Loan" & Space(5) & "Others" & Space(5) & "Total" & Space(5) & "Net Pay"
               ' printer.Print String(170, "-")
           ' End If
            End If
            rs1.Close
            rs.MoveNext
        Next
        
       ' printer.EndDoc
    End If
    rs.Close
End Sub
Public Sub ScreenPart()
ListView1.ListItems.clear
If Format(mskDate, "mm") <> 1 Then
    PreMonth = Format(mskDate, "mm") - 1
    PreYear = Format(mskDate, "yyyy")
Else
    PreMonth = 12
    PreYear = Format(mskDate, "yyyy") - 1
End If
sql = "Select * from daughters where CompanyName='" & lblCompanyName.Caption & "' and month(datemon)=" & PreMonth & " and Year(datemon)=" & PreYear & " order by empno"
    
    rs.Open sql, cnn, adOpenKeyset, adLockOptimistic
    If Not (rs.EOF Or rs.BOF) Then
        For i = 0 To rs.RecordCount - 1
            sql = "Select * from Branch where CompanyName='" & lblCompanyName.Caption & "' and BranchCode='" & cboBranch & "'"
            rs1.Open sql, cnn, adOpenKeyset, adLockOptimistic
            If Not (rs1.EOF Or rs1.BOF) Then
                Branch = rs1(0)
            Else
                MsgBox "No Employees", vbExclamation, "Payroll"
                rs.Close
                rs1.Close
                Exit Sub
            End If
            rs1.Close
            sql = ""
            sql = "Select Basic ,Father_Husband,BranchCode from admin where CompanyName='" & lblCompanyName.Caption & "' and branchcode='" & Branch & "'" & " And empno = " & rs(0)
            rs1.Open sql, cnn, adOpenKeyset, adLockOptimistic


           If Not (rs1.EOF Or rs1.BOF) Then
            Dim total
            Dim NetPay
            Dim totalEarn
            Dim totalDet
            total = Format(rs(5) + rs(7), "####0.00")
            totalEarn = Format((rs(5) / rs(10)) * rs(9) + (rs(7) / rs(10)) * rs(9) + rs(8) + rs(16), "####0.00")
            totalDet = Format(rs(11) + rs(12) + rs(13) + rs(15) + rs(14), "####0.00")
            NetPay = Format(totalEarn - totalDet, "####0.00")
           ' NetPay = AccurateCost(Format(CDbl(NetPay), "0.00"))
            
            Set lt = ListView1.ListItems.Add(, , i + 1)
            lt.SubItems(1) = UCase(rs(1))
            lt.SubItems(2) = Format(rs(9), "0.00")
            lt.SubItems(3) = Format(rs(5), "####0.00")
            lt.SubItems(4) = Format(rs(7), "####0.00")
            lt.SubItems(5) = total
            lt.SubItems(6) = Format((rs(5) / rs(10)) * rs(9), "####0.00")
            lt.SubItems(7) = Format((rs(7) / rs(10)) * rs(9), "####0.00")
            lt.SubItems(8) = Format(rs(8) + rs(16), "####0.00")
            lt.SubItems(9) = AccurateCost(CDbl(totalEarn))
            lt.SubItems(10) = Format(rs(11), "####0.00")
            lt.SubItems(11) = Format(rs(12), "####0.00")
            lt.SubItems(12) = Format(rs(13), "####0.00")
            lt.SubItems(13) = Format(rs(15) + rs(14), "####0.00")
            lt.SubItems(14) = totalDet
            lt.SubItems(15) = NetPay
            End If
            rs1.Close
            rs.MoveNext
        Next
       
    End If
    rs.Close

End Sub
Private Sub ScreenAll()
ListView1.ListItems.clear
If Format(mskDate, "mm") <> 1 Then
    PreMonth = Format(mskDate, "mm") - 1
    PreYear = Format(mskDate, "yyyy")
Else
    PreMonth = 12
    PreYear = Format(mskDate, "yyyy") - 1
End If

    sql = "Select * from daughters where CompanyName='" & lblCompanyName.Caption & "' and month(datemon)=" & PreMonth & " and Year(datemon)=" & PreYear & " order by empno"

    rs.Open sql, cnn, adOpenKeyset, adLockOptimistic
    If Not (rs.EOF Or rs.BOF) Then
        For i = 0 To rs.RecordCount - 1
            sql = "Select Basic,Father_Husband from admin where CompanyName='" & lblCompanyName.Caption & "' and empno =" & rs(0)
            rs1.Open sql, cnn, adOpenKeyset, adLockOptimistic
            Dim total
            Dim NetPay
            Dim totalEarn
            Dim totalDet
            total = Format(rs(5) + rs(7), "####0.00")
            totalEarn = Format((rs(5) / rs(10)) * rs(9) + (rs(7) / rs(10)) * rs(9) + rs(8) + rs(16), "####0.00")
            
            totalDet = Format(rs(11) + rs(12) + rs(13) + rs(15) + rs(14), "####0.00")
            NetPay = Format(totalEarn - totalDet, "####0.00")
            'NetPay = AccurateCost(Format(CDbl(NetPay), "0.00"))
            
            Set lt = ListView1.ListItems.Add(, , i + 1)
            lt.SubItems(1) = UCase(rs(1))
            lt.SubItems(2) = Format(rs(9), "0.0")
            lt.SubItems(3) = Format(rs(5), "####0.00")
            lt.SubItems(4) = Format(rs(7), "####0.00")
            lt.SubItems(5) = total
            lt.SubItems(6) = Format((rs(5) / rs(10)) * rs(9), "####0.00")
            lt.SubItems(7) = Format((rs(7) / rs(10)) * rs(9), "####0.00")
            lt.SubItems(8) = Format(rs(8) + rs(16), "####0.00")
            lt.SubItems(9) = AccurateCost(CDbl(totalEarn))
            lt.SubItems(10) = Format(rs(11), "####0.00")
            lt.SubItems(11) = Format(rs(12), "####0.00")
            lt.SubItems(12) = Format(rs(13), "####0.00")
            lt.SubItems(13) = Format(rs(15) + rs(14), "####0.00")
            lt.SubItems(14) = totalDet
            lt.SubItems(15) = NetPay
            rs1.Close
            rs.MoveNext
        Next
            
    Else
        MsgBox "No Records", vbExclamation, "Payroll"
        rs.Close
        Exit Sub
    End If
    
rs.Close
End Sub
Public Function AccurateCost(Text As Double)
If Right(Format(Text, ".00"), 1) = 5 Then
    AccurateCost = Format(Text + 0.05, "0.00")
Else
    AccurateCost = Format((Round(Text, 1)), "0.00")
End If
End Function


