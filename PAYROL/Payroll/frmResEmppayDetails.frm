VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmResEmppayDetails 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resign Employee Pay Details Report"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11715
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
   ScaleHeight     =   6270
   ScaleWidth      =   11715
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
      Top             =   5760
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
      Top             =   5760
      Width           =   735
   End
   Begin VB.Frame fraResno 
      Caption         =   "ResignNo"
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
      Top             =   1200
      Width           =   3855
      Begin VB.ComboBox cboResignNo 
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
      Top             =   1200
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
   Begin VB.Frame fraselect 
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
      Top             =   1200
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
      TabIndex        =   1
      Top             =   2760
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
      TabIndex        =   0
      Top             =   720
      Width           =   1935
      _ExtentX        =   3413
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
      Left            =   5160
      TabIndex        =   15
      Top             =   240
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "CompanyName"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   3360
      TabIndex        =   14
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Resign Employee PayDetails"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   3240
      TabIndex        =   13
      Top             =   720
      Width           =   2910
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
      TabIndex        =   12
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "frmResEmppayDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lt As ListItem
Dim Msg As String
Dim PreMonth As Integer
Dim PreYear As Integer
Private Sub cboResignNo_Validate(Cancel As Boolean)
If optPart Then
If cboResignNo.ListIndex = -1 Then
    MsgBox "Select ResgnNo To View/Print", vbExclamation, "Payroll"
    Cancel = True
    SendKeys "{Home}+{End}"
    Exit Sub
End If
End If
End Sub
Private Sub CmdCancel_Click()
    clear
    ListView1.ListItems.clear
    fraResno.Visible = False
    Cancel = True
    mskDate.SetFocus
End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub Form_Load()
frmResEmppayDetails.Icon = LoadPicture(App.Path & "\Jamaica my computer.ico")
lblCompanyName.Caption = CompanyName
rs.Open "Select Distinct(ResignNo) from ResEmppayDetails Where CompanyName='" & lblCompanyName.Caption & "'", cnn, adOpenKeyset, adLockOptimistic
cboResignNo.clear
While Not rs.EOF
cboResignNo.AddItem rs.Fields("ResignNo")
rs.MoveNext
Wend
rs.Close
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
    fraResno.Visible = False
End Sub
Private Sub optPart_Click()
    fraResno.Visible = True
    cboResignNo.SetFocus
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
Private Sub ScreenAll()
ListView1.ListItems.clear
If Format(mskDate, "mm") <> 1 Then
    PreMonth = Format(mskDate, "mm") - 1
    PreYear = Format(mskDate, "yyyy")
Else
    PreMonth = 12
    PreYear = Format(mskDate, "yyyy") - 1
End If
    sql = "Select * from ResEmpPayDetails where CompanyName='" & lblCompanyName.Caption & "' and month(datemon)=" & PreMonth & " and Year(datemon)=" & PreYear & " order by ResignNo"

    rs.Open sql, cnn, adOpenKeyset, adLockOptimistic
    If Not (rs.EOF Or rs.BOF) Then
        For i = 0 To rs.RecordCount - 1
            sql = "Select Basic,Father_Husband from ResEmpMaster where CompanyName='" & lblCompanyName.Caption & "' and ResignNo =" & rs(0)
            rs1.Open sql, cnn, adOpenKeyset, adLockOptimistic
            Dim total
            Dim NetPay
            Dim totalEarn
            Dim totalDet
            total = Format(rs(6) + rs(8), "####0.00")
            totalEarn = Format((rs(6) / rs(11)) * rs(10) + (rs(8) / rs(11)) * rs(10) + rs(9) + rs(17), "####0.00")
            totalEarn = AccurateCost(CDbl(totalEarn))
            totalDet = Format(rs(12) + rs(13) + rs(14) + rs(16) + rs(15), "####0.00")
            NetPay = Format(totalEarn - totalDet, "####0.00")
            'NetPay = AccurateCost(Format(CDbl(NetPay), "0.00"))
            
            Set lt = ListView1.ListItems.Add(, , "STP/" & Format(rs(1), "000"))
            lt.SubItems(1) = UCase(rs(2))
            lt.SubItems(2) = Format(rs(10), "0.0")
            lt.SubItems(3) = Format(rs(6), "####0.00")
            lt.SubItems(4) = Format(rs(8), "####0.00")
            lt.SubItems(5) = total
            lt.SubItems(6) = Format((rs(6) / rs(11)) * rs(10), "####0.00")
            lt.SubItems(7) = Format((rs(8) / rs(11)) * rs(10), "####0.00")
            lt.SubItems(8) = Format(rs(9) + rs(17), "####0.00")
            lt.SubItems(9) = totalEarn
            lt.SubItems(10) = Format(rs(12), "####0.00")
            lt.SubItems(11) = Format(rs(13), "####0.00")
            lt.SubItems(12) = Format(rs(14), "####0.00")
            lt.SubItems(13) = Format(rs(16) + rs(15), "####0.00")
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
Public Sub ScreenPart()
ListView1.ListItems.clear
If Format(mskDate, "mm") <> 1 Then
    PreMonth = Format(mskDate, "mm") - 1
    PreYear = Format(mskDate, "yyyy")
Else
    PreMonth = 12
    PreYear = Format(mskDate, "yyyy") - 1
End If
sql = "Select * from ResEmpPayDetails where CompanyName='" & lblCompanyName.Caption & "' and month(datemon)=" & PreMonth & " and Year(datemon)=" & PreYear & " and ResignNo=" & cboResignNo
    
    rs.Open sql, cnn, adOpenKeyset, adLockOptimistic
    If Not (rs.EOF Or rs.BOF) Then
        For i = 0 To rs.RecordCount - 1
            
            sql = "Select Basic ,Father_Husband from ResEmpMaster where CompanyName='" & lblCompanyName.Caption & "' and  ResignNo = " & rs(0)
            rs1.Open sql, cnn, adOpenKeyset, adLockOptimistic


           If Not (rs1.EOF Or rs1.BOF) Then
            Dim total
            Dim NetPay
            Dim totalEarn
            Dim totalDet
            total = Format(rs(6) + rs(8), "####0.00")
            totalEarn = Format((rs(6) / rs(11)) * rs(10) + (rs(8) / rs(11)) * rs(10) + rs(9) + rs(17), "####0.00")
            totalEarn = AccurateCost(CDbl(totalEarn))
            totalDet = Format(rs(12) + rs(13) + rs(14) + rs(16) + rs(15), "####0.00")
            NetPay = Format(totalEarn - totalDet, "####0.00")
            'NetPay = AccurateCost(Format(CDbl(NetPay), "0.00"))
            
            Set lt = ListView1.ListItems.Add(, , "STP/" & Format(rs(1), "000"))
            lt.SubItems(1) = UCase(rs(2))
            lt.SubItems(2) = Format(rs(10), "0.0")
            lt.SubItems(3) = Format(rs(6), "####0.00")
            lt.SubItems(4) = Format(rs(8), "####0.00")
            lt.SubItems(5) = total
            lt.SubItems(6) = Format((rs(6) / rs(11)) * rs(10), "####0.00")
            lt.SubItems(7) = Format((rs(7) / rs(11)) * rs(10), "####0.00")
            lt.SubItems(8) = Format(rs(9) + rs(17), "####0.00")
            lt.SubItems(9) = totalEarn
            lt.SubItems(10) = Format(rs(12), "####0.00")
            lt.SubItems(11) = Format(rs(13), "####0.00")
            lt.SubItems(12) = Format(rs(14), "####0.00")
            lt.SubItems(13) = Format(rs(16) + rs(15), "####0.00")
            lt.SubItems(14) = totalDet
            lt.SubItems(15) = NetPay
            End If
            rs1.Close
            rs.MoveNext
        Next
       
    End If
    rs.Close

End Sub
Private Sub PrintAll()
Msg = MsgBox("Are you take Print", vbExclamation + vbYesNo, "Payroll")
If Msg = 6 Then
    On Error GoTo err
    Printer.Orientation = 1
    Printer.FontName = "Courier New"
    Printer.Font.Size = 10
    Printer.Print ""
    Printer.FontBold = True
    Printer.Font.Size = 14
    Printer.Print Space(40) & CompanyName
    Printer.Font.Size = 12
    Printer.Print Space(50) & "Salary For The Month Of " & MonthName(PreMonth) & " " & year(Month(PreMonth) & "- 01" & "-" & PreYear)
    Printer.Font.Size = 10
    Printer.Print Space(50) & "See Rule 11(5) G.O.M.S.1216 dt 23.9.71 p.of.w.rules"
    Printer.Print ""
    Printer.Print ""
    
    Printer.FontBold = False
    Printer.Font.Size = 8
    Printer.Print String(220, "-")
    Printer.Print "EmpNo" & Space(11) & "Emp Name" & Space(30) & " Days " & Space(5) & " Basic " & Space(4) & " DA " & Space(5) & " Total " & Space(2) & " Basic " & Space(3) & " DA " & Space(6) & " Others " & Space(2) & " Total " & Space(4) & " PF " & Space(5) & " ESIC " & Space(5) & " Loan " & Space(3) & " Others " & Space(3) & " Total " & Space(3) & " NetPay"
    Printer.Print Space(66) & "<------  ACTUALS  ------->" & Space(3) & "<-----------  EARNINGS  ----------->" & Space(4) & "<---------------  DEDUCTIONS  --------------->"
    Printer.Print String(220, "-")
    
    If Format(mskDate, "mm") <> 1 Then
        PreMonth = Format(mskDate, "mm") - 1
        PreYear = Format(mskDate, "yyyy")
    Else
        PreMonth = 12
        PreYear = Format(mskDate, "yyyy") - 1
    End If
    
    sql = "Select * from ResEmpPayDetails where CompanyName='" & lblCompanyName.Caption & "' and month(datemon)=" & PreMonth & " and Year(datemon)=" & PreYear & " order by ResignNo"
    'sql = "Select *From Daughters where datemon=#" & Format(mskDate, "mm/dd/yyyy") & "#" '& Format("datemon", "mm/yyyy") & "= '" & mskDate & "#" order by empno"
    j = 0
    rs.Open sql, cnn, adOpenKeyset, adLockOptimistic
    If Not (rs.EOF Or rs.BOF) Then
        For i = 0 To rs.RecordCount - 1
            j = j + 1
            sql = "Select Basic,Father_Husband from ResEmpMaster where CompanyName='" & lblCompanyName.Caption & "' and ResignNo =" & rs(0)
            rs1.Open sql, cnn, adOpenKeyset, adLockOptimistic
            Dim total
            Dim NetPay
            Dim totalEarn
            Dim totalDet
            total = Format(rs(6) + rs(8), "0.00")
            totalEarn = Format((rs(6) / rs(11)) * rs(10) + (rs(8) / rs(11)) * rs(10) + rs(9) + rs(17), "0.00")
            totalEarn = AccurateCost(CDbl(totalEarn))
            totalDet = Format(rs(12) + rs(13) + rs(14) + rs(16) + rs(15), "0.00")
            NetPay = Format(totalEarn - totalDet, "0.00")
            'NetPay = AccurateCost(Format(CDbl(NetPay), "0.00"))
            
            
            Printer.Print "STP/" & Format(rs(1), "000") & Space(16 - Len("STP/" & Format(rs(1), "000"))) & UCase(rs(2)) & Space(40 - Len(UCase(rs(2)))) & Format(rs(10), "0.0") & Space(6 - Len(rs(10))) & Format(rs(6), "0.00") & Space(12 - Len(Format(rs(6), "0.00"))) & Format(rs(8), "0.00") & Space(10 - Len(Format(rs(8), "0.00"))) & total & Space(10 - Len(Format(total, "0.00"))) & Format((rs(6) / rs(11)) * rs(10), "0.00") & _
  Space(10 - Len(Format((rs(6) / rs(11)) * rs(10), "0.00"))) & Format((rs(8) / rs(11)) * rs(10), "0.00") & Space(10 - Len(Format((rs(8) / rs(11)) * rs(10), "0.00"))) & Format(rs(9) + rs(17), "0.00") & Space(10 - Len(Format(rs(9) + rs(17), "0.00"))) & totalEarn & Space(10 - Len(totalEarn)) & Format(rs(12), "0.00") & Space(10 - Len(Format(rs(12), "0.00"))) & Format(rs(13), "0.00") & Space(10 - Len(Format(rs(13), "0.00"))) & Format(rs(14), "0.00") & Space(10 - Len(Format(rs(14), "0.00"))) & Format(rs(16) + rs(15), "0.00") & Space(10 - Len(Format(rs(16) + rs(15), "0.00"))) & totalDet & Space(10 - Len(totalDet) _
            ) & NetPay & Space(10 - Len(NetPay))
            'Printer.print Space(171) & "------"
'            Printer.Print Space(179 - Len(NetPay)) & "|" & Space(12) & "|"
'            Printer.Print Space(179 - Len(NetPay)) & "|" & Space(12) & "|"
'            Printer.Print Space(179 - Len(NetPay)) & "|" & Space(12) & "|"
'            Printer.Print Space(179 - Len(NetPay)) & "|" & Space(12) & "|"
           ' Printer.Print Space(173) & "------------"
            Printer.Print ""
            Printer.Print
            If j = 8 Then
                j = 0
                'Printer.Print String(79, "-") & "End Of Page" & String(79, "-")
                Printer.NewPage
                Printer.Orientation = 2
                Printer.FontName = "Courier New"
                Printer.Font.Size = 10
                Printer.Print ""
                Printer.Print ""
                Printer.FontBold = True
                Printer.Font.Size = 14
                Printer.Print Space(50) & CompanyName
                Printer.Font.Size = 12
                Printer.Print Space(50) & "Salary For The Month Of " & MonthName(PreMonth) & " " & year(Month(PreMonth) & "- 01" & "-" & PreYear)
                Printer.Font.Size = 10
                Printer.Print Space(50) & "Under rule 27(2) of the Mini Wages Madras Rules 1953"
                Printer.Print ""
                Printer.Print ""
                Printer.Font.Size = 8
                Printer.Print String(220, "-")
                Printer.Print "EmpNo" & Space(11) & "Emp Name" & Space(30) & " Days " & Space(5) & " Basic " & Space(4) & " DA " & Space(5) & " Total " & Space(2) & " Basic " & Space(3) & " DA " & Space(6) & " Others " & Space(2) & " Total " & Space(4) & " PF " & Space(5) & " ESIC " & Space(5) & " Loan " & Space(3) & " Others " & Space(3) & " Total " & Space(3) & " NetPay"
                Printer.Print Space(49) & "<------  ACTUALS  ------->" & Space(3) & "<-----------  EARNINGS  ----------->" & Space(4) & "<---------------  DEDUCTIONS  --------------->"
                Printer.Print String(220, "-")
            End If
            rs1.Close
            rs.MoveNext
            
        Next
    End If
    Printer.EndDoc
    rs.Close
err:
        If err.Number = 482 Or err.Number = 484 Then
            MsgBox "Make Sure The printer Is Ready", vbExclamation, "Payroll"
        End If
        Exit Sub
End If
End Sub
Private Sub PrintPart()
Msg = MsgBox("Are you take Print", vbExclamation + vbYesNo, "Payroll")
If Msg = 6 Then
    On Error GoTo err
    Printer.Orientation = 1
    Printer.FontName = "Courier New"
    Printer.Font.Size = 14
    Printer.Print ""
    Printer.FontBold = True
    Printer.Print Space(40) & CompanyName
    Printer.Font.Size = 12
    Printer.Print Space(50) & "Salary For The Month Of " & MonthName(PreMonth) & " " & year(Month(PreMonth) & "- 01" & "-" & PreYear)
    Printer.Font.Size = 10
    Printer.Print Space(50) & "See Rule 11(5) G.O.M.S.1216 dt 23.9.71 p.of.w.rules"
    Printer.Print ""
    Printer.Print ""
    Printer.FontBold = False
    Printer.Font.Size = 8
    Printer.Print String(220, "-")
    Printer.Print "EmpNo" & Space(11) & "Emp Name" & Space(30) & " Days " & Space(3) & " Basic " & Space(4) & " DA " & Space(5) & " Total " & Space(4) & " Basic " & Space(3) & " DA " & Space(6) & " Others " & Space(2) & " Total " & Space(4) & " PF " & Space(5) & " ESIC " & Space(5) & " Loan " & Space(3) & " Others " & Space(3) & " Total " & Space(3) & " NetPay"
    Printer.Print Space(66) & "<------  ACTUALS  ------->" & Space(5) & "<-----------  EARNINGS  ----------->" & Space(4) & "<---------------  DEDUCTIONS  --------------->"
    Printer.Print String(220, "-")
    
    If Format(mskDate, "mm") <> 1 Then
        PreMonth = Format(mskDate, "mm") - 1
        PreYear = Format(mskDate, "yyyy")
    Else
        PreMonth = 12
        PreYear = Format(mskDate, "yyyy") - 1
    End If
    
    sql = "Select * from resEmpPayDetails where CompanyName='" & lblCompanyName.Caption & "' and month(datemon)=" & PreMonth & " and Year(datemon)=" & PreYear & " and ResignNo=" & cboResignNo
    'sql = "Select *From Daughters where datemon=#" & Format(mskDate, "mm/dd/yyyy") & "#" '& Format("datemon", "mm/yyyy") & "= '" & mskDate & "#" order by empno"
    j = 0
    rs.Open sql, cnn, adOpenKeyset, adLockOptimistic
    If Not (rs.EOF Or rs.BOF) Then
        For i = 0 To rs.RecordCount - 1
            j = j + 1
            sql = "Select Basic ,Father_Husband from ResEmpMaster where CompanyName='" & lblCompanyName.Caption & "' and ResignNo = " & rs(0)
            rs1.Open sql, cnn, adOpenKeyset, adLockOptimistic

           ' clear
           'MsgBox rs(6)
           If Not (rs1.EOF Or rs1.BOF) Then
            Dim total
            Dim NetPay
            Dim totalEarn
            Dim totalDet
            total = Format(rs(6) + rs(8), "####0.00")
            totalEarn = Format((rs(6) / rs(11)) * rs(10) + (rs(8) / rs(11)) * rs(10) + rs(9) + rs(17), "####0.00")
            totalEarn = AccurateCost(CDbl(totalEarn))
            totalDet = Format(rs(12) + rs(13) + rs(14) + rs(16) + rs(15), "####0.00")
            NetPay = Format(totalEarn - totalDet, "####0.00")
            'NetPay = AccurateCost(Format(CDbl(NetPay), "0.00"))
            
            Printer.Font.Size = 8
            Printer.Font.Bold = False
            Printer.Print "STP/" & Format(rs(1), "000") & Space(16 - Len("STP/" & Format(rs(1), "000") _
)) & UCase(rs(2)) & Space(40 - Len(UCase(rs(2)))) & Format(rs(10), "0.0") & Space(6 - Len(rs(10))) & Format(rs(6), "####0.00") & Space(10 - Len(Format(rs(6), "####0.00"))) & Format(rs(8), "####0.00") & Space(10 - Len(Format(rs(8), "####0.00"))) & total & Space(10 - Len(Format(total, "####0.00"))) & Format((rs(6) / rs(11)) * rs(10), "####0.00") & Space(10 - Len(Format(rs(6), "####0.00"))) & Format((rs(8) / rs(11)) * rs(10), "####0.00") & Space(10 - Len(Format((rs(8) / rs(11)) * rs(10), "####0.00"))) & _
Format(rs(9) + rs(17), "####0.00") & Space(10 - Len(Format(rs(9) + rs(17), "####0.00"))) & totalEarn & Space(10 - Len(totalEarn)) & Format(rs(12), "####0.00") & Space(10 - Len(Format(rs(12), "####0.00"))) & Format(rs(13), "####0.00") & Space(10 - Len(Format(rs(13), "####0.00"))) & Format(rs(14), "####0.00") & Space(10 - Len(Format(rs(14), "####0.00"))) & Format(rs(16) + rs(15), "####0.00") & Space(10 - Len(Format(rs(16) + rs(15), "####0.00"))) & totalDet & Space(10 - Len(totalDet)) & NetPay
            
            Printer.Print ""
            Printer.Print
            Printer.Print ""
            Printer.Print
            If j = 8 Then
                j = 0
                'Printer.Print String(94, "-") & "End Of Page" & String(94, "-")
                Printer.NewPage
                Printer.Orientation = 2
                Printer.FontName = "Courier New"
                Printer.Font.Size = 10
                Printer.Print ""
                Printer.Print ""
                Printer.FontBold = True
                Printer.Font.Size = 14
                Printer.Print Space(50) & CompanyName
                Printer.Font.Size = 12
                Printer.Print Space(50) & "Salary For The Month Of " & MonthName(PreMonth) & " " & year(Month(PreMonth) & "- 01" & "-" & PreYear)
                Printer.Font.Size = 10
                Printer.Print Space(50) & "Under rule 27(2) of the Mini Wages Madras Rules 1953"
                Printer.Print ""
                Printer.Print ""
                Printer.Font.Size = 8
                Printer.Print String(220, "-")
                Printer.Print "EmpNo" & Space(11) & "Emp Name" & Space(12) & " Days " & Space(3) & " Basic " & Space(4) & " DA " & Space(5) & " Total " & Space(4) & " Basic " & Space(3) & " DA " & Space(6) & " Others " & Space(2) & " Total " & Space(4) & " PF " & Space(5) & " ESIC " & Space(5) & " Loan " & Space(3) & " Others " & Space(3) & " Total " & Space(3) & " NetPay"
                Printer.Print Space(46) & "<------  ACTUALS  ------->" & Space(5) & "<-----------  EARNINGS  ----------->" & Space(4) & "<---------------  DEDUCTIONS  --------------->"
                Printer.Print String(220, "-")
            End If
            End If
            rs1.Close
            rs.MoveNext
        Next
        
        Printer.EndDoc
    End If
    rs.Close
err:
    If err.Number = 482 Then
        MsgBox "Make Sure The printer Is Ready", vbExclamation, "Payroll"
    End If
  '  Exit Sub
End If
End Sub
Private Sub Form_Activate()
frmResEmppayDetails.Left = 0
frmResEmppayDetails.Top = 0
fraResno.Visible = False
ListView1.ColumnHeaders.clear
ListView1.ListItems.clear

ListView1.ColumnHeaders.Add , , "Employee No"
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

Public Function AccurateCost(Text As Double)
If Right(Format(Text, ".00"), 1) = 5 Then
    AccurateCost = Format(Text + 0.05, "0.00")
Else
    AccurateCost = Format((Round(Text, 1)), "0.00")
End If
End Function
