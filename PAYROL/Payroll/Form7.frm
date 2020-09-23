VERSION 5.00
Begin VB.Form frmloanDetails 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loan Details"
   ClientHeight    =   8085
   ClientLeft      =   810
   ClientTop       =   330
   ClientWidth     =   9825
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   9825
   Begin VB.CommandButton cmddelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   1800
      TabIndex        =   45
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   3000
      TabIndex        =   44
      Top             =   7680
      Width           =   1095
   End
   Begin VB.ListBox lstloan 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   7200
      TabIndex        =   31
      Top             =   1920
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Deduction"
      Height          =   375
      Left            =   600
      TabIndex        =   30
      Top             =   7680
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Height          =   360
      Left            =   1605
      Locked          =   -1  'True
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   7080
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Pay"
      Height          =   240
      Index           =   2
      Left            =   5760
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   6120
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Pay"
      Height          =   240
      Index           =   1
      Left            =   5760
      TabIndex        =   26
      Top             =   3960
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Pay"
      Height          =   240
      Index           =   0
      Left            =   5760
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   360
      Index           =   2
      Left            =   1605
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   6000
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   360
      Index           =   2
      Left            =   1605
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3720
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   2
      Left            =   1605
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Index           =   1
      Left            =   1605
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5520
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   0
      Left            =   1605
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Index           =   0
      Left            =   1605
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   5040
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   1
      Left            =   1605
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   1605
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   1590
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label23 
      Caption         =   "( Press F2 Key)"
      ForeColor       =   &H8000000D&
      Height          =   465
      Left            =   2970
      TabIndex        =   50
      Top             =   4995
      Width           =   750
   End
   Begin VB.Label Label22 
      Caption         =   " (Press F2 Key)"
      ForeColor       =   &H8000000D&
      Height          =   465
      Left            =   2985
      TabIndex        =   49
      Top             =   2715
      Width           =   750
   End
   Begin VB.Label Label21 
      Caption         =   " (Press F2 Key)"
      ForeColor       =   &H8000000D&
      Height          =   465
      Left            =   2880
      TabIndex        =   48
      Top             =   540
      Width           =   750
   End
   Begin VB.Label lblCompanyName 
      AutoSize        =   -1  'True
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   3720
      TabIndex        =   47
      Top             =   120
      Width           =   45
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "Company Name"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   2040
      TabIndex        =   46
      Top             =   120
      Width           =   1260
   End
   Begin VB.Label lblBalMonth2 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5640
      TabIndex        =   43
      Top             =   5400
      Width           =   75
   End
   Begin VB.Label lblBalMonth1 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5640
      TabIndex        =   42
      Top             =   3120
      Width           =   75
   End
   Begin VB.Label lblbalAmount2 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5640
      TabIndex        =   41
      Top             =   4920
      Width           =   75
   End
   Begin VB.Label lblbalAmount1 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5640
      TabIndex        =   40
      Top             =   2640
      Width           =   75
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "BalanceAmount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3840
      TabIndex        =   39
      Top             =   4920
      Width           =   1635
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "Balance Dues"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3840
      TabIndex        =   38
      Top             =   5400
      Width           =   1470
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "BalanceAmount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3840
      TabIndex        =   37
      Top             =   2640
      Width           =   1635
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Balance Dues"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3840
      TabIndex        =   36
      Top             =   3120
      Width           =   1470
   End
   Begin VB.Label lblBalMonth 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5640
      TabIndex        =   35
      Top             =   1080
      Width           =   75
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "Balance Dues"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3840
      TabIndex        =   34
      Top             =   1080
      Width           =   1470
   End
   Begin VB.Label lblbalAmount 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5640
      TabIndex        =   33
      Top             =   600
      Width           =   75
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "BalanceAmount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3840
      TabIndex        =   32
      Top             =   600
      Width           =   1635
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Total"
      Height          =   225
      Left            =   120
      TabIndex        =   29
      Top             =   7200
      Width           =   420
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Date"
      Height          =   225
      Left            =   240
      TabIndex        =   24
      Top             =   6000
      Width           =   375
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Date"
      Height          =   225
      Left            =   240
      TabIndex        =   23
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Date"
      Height          =   225
      Left            =   120
      TabIndex        =   22
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label lblinstall3 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   1605
      TabIndex        =   21
      Top             =   6600
      Width           =   75
   End
   Begin VB.Label lblinstall2 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   1605
      TabIndex        =   20
      Top             =   4320
      Width           =   75
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "No of Instalment"
      Height          =   225
      Left            =   120
      TabIndex        =   19
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Instal/Month"
      Height          =   225
      Left            =   120
      TabIndex        =   18
      Top             =   6600
      Width           =   1050
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Instal/Month"
      Height          =   225
      Left            =   120
      TabIndex        =   17
      Top             =   4320
      Width           =   1050
   End
   Begin VB.Label lblbalance 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   6840
      TabIndex        =   16
      Top             =   600
      Width           =   75
   End
   Begin VB.Label lblinstall 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   1605
      TabIndex        =   15
      Top             =   2160
      Width           =   75
   End
   Begin VB.Label label8 
      AutoSize        =   -1  'True
      Caption         =   "Instal/Month"
      Height          =   225
      Left            =   120
      TabIndex        =   14
      Top             =   2160
      Width           =   1050
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "No of Instalment"
      Height          =   225
      Left            =   120
      TabIndex        =   13
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "No of Instalment "
      Height          =   225
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   1380
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Third Advance"
      Height          =   225
      Left            =   120
      TabIndex        =   11
      Top             =   5160
      Width           =   1185
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Second Advance"
      Height          =   225
      Left            =   120
      TabIndex        =   10
      Top             =   2760
      Width           =   1305
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "First Advance"
      Height          =   225
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1125
   End
End
Attribute VB_Name = "frmloanDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Tot As Integer
'***Dim cnn As New ADODB.Connection
'***Dim rs As New ADODB.Recordset
'***Dim rs1 As New ADODB.Recordset
Dim textindex As Integer
Public lid1 As Integer
Public lid2 As Integer
Public lid3 As Integer
Dim Msg As String
Dim Mon As String
Dim year As Integer
Private Sub Check1_Click(Index As Integer)
Tot = 0
If Check1(0).value = 1 Then
    Tot = Tot + Val(lblinstall)
End If
If Check1(1).value = 1 Then
    Tot = Tot + Val(lblinstall2)
End If
If Check1(2).value = 1 Then
    Tot = Tot + Val(lblinstall3)
End If
Text4 = Tot
End Sub
Private Sub CmdDelete_Click()
Msg = MsgBox("Are you sure to Delete this Record", vbExclamation + vbYesNo, "Payroll")
If Msg = 6 Then

rs.Open "select * from LoanDetails where CompanyName='" & lblCompanyName.Caption & "' and empno=" & Right(frmemployeepaydetails.Text1(0), Len(frmemployeepaydetails.Text1(0)) - InStr(1, frmemployeepaydetails.Text1(0), "/")) & "  and  monyear=#" & 1 & "-" & Mon & " - " & year & "# And paid =true", cnn, adOpenKeyset, adLockOptimistic

While Not rs.EOF
    rs1.Open "Select Balance from LoanMaster Where CompanyName='" & lblCompanyName.Caption & "' and Empno= " & Right(frmemployeepaydetails.Text1(0), Len(frmemployeepaydetails.Text1(0)) - InStr(1, frmemployeepaydetails.Text1(0), "/")) & " and Loanid=" & rs.Fields("LoanId"), cnn, adOpenKeyset, adLockOptimistic
    cnn.Execute "Update LoanMaster set Balance=" & CDbl(rs1.Fields("Balance")) & " +" & CDbl(rs.Fields("paidAmount")) & " Where CompanyName='" & lblCompanyName.Caption & "' and Empno= " & Right(frmemployeepaydetails.Text1(0), Len(frmemployeepaydetails.Text1(0)) - InStr(1, frmemployeepaydetails.Text1(0), "/")) & " and Loanid=" & rs.Fields("LoanId")
    rs1.Close
    rs.MoveNext
Wend
rs.Close
cnn.Execute "Delete From LoanDetails Where CompanyName='" & lblCompanyName.Caption & "' and Empno= " & Right(frmemployeepaydetails.Text1(0), Len(frmemployeepaydetails.Text1(0)) - InStr(1, frmemployeepaydetails.Text1(0), "/")) & " and monyear=#" & 1 & "-" & Mon & " - " & year & "#"
frmemployeepaydetails.Text2(2) = 0

cnn.Execute "Update Daughters set loandecl=0 where CompanyName='" & lblCompanyName.Caption & "' and Empno= " & Right(frmemployeepaydetails.Text1(0), Len(frmemployeepaydetails.Text1(0)) - InStr(1, frmemployeepaydetails.Text1(0), "/")) & " and datemon=#" & 1 & "-" & Mon & " - " & year & "#"
clear
Check1(0).Enabled = True
Check1(0).value = 0
Check1(1).Enabled = True
Check1(1).value = 0
Check1(2).Enabled = True
Check1(2).value = 0
lstLoan.Visible = True
Command1.Enabled = True
Command2.Enabled = True
cmdDelete.Enabled = False
addloan
End If
End Sub
Private Sub Command1_Click()

frmemployeepaydetails.Text2(2) = Val(Text4)
Cancel = True
frmemployeepaydetails.Text2(3).SetFocus
'Unload Me
End Sub
Private Sub Command2_Click()
clear
lblbalAmount.Caption = ""
lblbalAmount1.Caption = ""
lblbalAmount2.Caption = ""
lblBalMonth = ""
lblBalMonth1 = ""
lblBalMonth2 = ""
lblinstall = ""
lblinstall2 = ""
lblinstall3 = ""
addloan
lstLoan.Visible = False
End Sub
Private Sub Form_Activate()
Mon = frmemployeepaydetails.Mon
year = frmemployeepaydetails.year
If frmemployeepaydetails.AddEditViewMode = "Edit" Then
    rs.Open "select count(*) from LoanDetails where CompanyName='" & lblCompanyName.Caption & "' and empno=" & Right(frmemployeepaydetails.Text1(0), Len(frmemployeepaydetails.Text1(0)) - InStr(1, frmemployeepaydetails.Text1(0), "/")) & "  and  monyear=#" & 1 & "-" & Mon & " - " & year & "# And paid =true", cnn, adOpenKeyset, adLockOptimistic

    If rs.Fields(0) > 0 Then
        cmdDelete.Enabled = True
    Else
        cmdDelete.Enabled = False
    End If
    rs.Close
End If
If frmemployeepaydetails.AddEditViewMode = "Edit" And cmdDelete.Enabled = True Then
    Check1(0).Enabled = False
    Check1(1).Enabled = False
    Check1(2).Enabled = False
    lstLoan.Visible = False
    Command1.Enabled = False
    Command2.Enabled = False
ElseIf frmemployeepaydetails.AddEditViewMode = "Edit" And cmdDelete.Enabled = False Then
    Check1(0).Enabled = True
    Check1(1).Enabled = True
    Check1(2).Enabled = True
    lstLoan.Visible = True
    Command1.Enabled = True
    Command2.Enabled = True
ElseIf frmemployeepaydetails.AddEditViewMode = "View" Then
    Check1(0).Enabled = False
    Check1(1).Enabled = False
    Check1(2).Enabled = False
    lstLoan.Visible = False
    Command1.Enabled = False
    Command2.Enabled = False
    cmdDelete.Enabled = False
ElseIf frmemployeepaydetails.AddEditViewMode = "Add" Then
    Check1(0).Enabled = True
    Check1(1).Enabled = True
    Check1(2).Enabled = True
    lstLoan.Visible = True
    Command1.Enabled = True
    Command2.Enabled = True
    cmdDelete.Enabled = False
End If

rs.Open "select * from LoanDetails where CompanyName='" & lblCompanyName.Caption & "' and empno=" & Right(frmemployeepaydetails.Text1(0), Len(frmemployeepaydetails.Text1(0)) - InStr(1, frmemployeepaydetails.Text1(0), "/")) & "  and  monyear=#" & 1 & "-" & Mon & " - " & year & "# And paid =true", cnn, adOpenKeyset, adLockOptimistic
'MsgBox "select * from LoanDetails where empno=" & Right(frmemployeepaydetails.Text1(0), Len(frmemployeepaydetails.Text1(0)) - InStr(1, frmemployeepaydetails.Text1(0), "/")) & "  and  monyear=#" & 1 & "-" & mon & " - " & year & "# And paid =true"
Counter = 1
While Not rs.EOF
If Counter = 1 Then

    rs1.Open "Select * from LoanMaster where CompanyName='" & lblCompanyName.Caption & "' and empno=" & rs.Fields("Empno") & " and loanid=" & rs.Fields("LoanId"), cnn, adOpenKeyset, adLockOptimistic
    Text1(0) = rs1.Fields("LoanAmt")
    Text1(1) = rs1.Fields("NoInstall")
    Text1(2) = rs1.Fields("Dateapp")
    If rs.Fields("paid") = True Then
        Check1(0).value = 1
    Else
        Check1(0).value = 0
    End If
    rs1.Close
End If
If Counter = 2 Then
    rs1.Open "Select * from LoanMaster where CompanyName='" & lblCompanyName.Caption & "' and empno=" & rs.Fields("Empno") & " and loanid=" & rs.Fields("LoanId"), cnn, adOpenKeyset, adLockOptimistic
    Text2(0) = rs1.Fields("LoanAmt")
    Text2(1) = rs1.Fields("NoInstall")
    Text2(2) = rs1.Fields("Dateapp")
    If rs.Fields("paid") = True Then
        Check1(1).value = 1
    Else
        Check1(1).value = 0
    End If
    rs1.Close
End If
If Counter = 3 Then
    rs1.Open "Select * from LoanMaster where CompanyName='" & lblCompanyName.Caption & "' and empno=" & rs.Fields("Empno") & " and loanid=" & rs.Fields("LoanId"), cnn, adOpenKeyset, adLockOptimistic
    Text3(0) = rs1.Fields("LoanAmt")
    Text3(1) = rs1.Fields("NoInstall")
    Text3(2) = rs1.Fields("Dateapp")
    If rs.Fields("paid") = True Then
        Check1(2).value = 1
    Else
        Check1(2).value = 0
    End If
    rs1.Close
End If
Counter = Counter + 1
rs.MoveNext
Wend
rs.Close
End Sub
Private Sub Form_Load()
'***cnn.Provider = "Microsoft.Jet.OLEDB.4.0"
'***cnn.Open app.path & "\Moolakadai.mdb" ' payroll1.mdb"
frmloanDetails.Icon = LoadPicture(App.Path & "\Jamaica my computer.ico")
lblCompanyName.Caption = CompanyName
addloan
End Sub
Private Sub lstloan_DblClick()
rs.Open "Select *  from LoanMaster Where CompanyName='" & lblCompanyName.Caption & "' and EmpNo=" & Right(frmemployeepaydetails.Text1(0), Len(frmemployeepaydetails.Text1(0)) - InStr(1, frmemployeepaydetails.Text1(0), "/")) & " And loanid = " & Mid(lstLoan.Text, 1, InStr(1, lstLoan, "-") - 1), cnn, adOpenKeyset, adLockOptimistic
If rs.EOF = False Then
    If textindex = 0 Then
        Text1(0) = rs.Fields("LoanAmt")
        Text1(1) = rs.Fields("NoInstall")
        Text1(2) = rs.Fields("Dateapp")
        lid1 = Mid(lstLoan.Text, 1, InStr(1, lstLoan, "-") - 1)
    ElseIf textindex = 1 Then
        Text2(0) = rs.Fields("LoanAmt")
        Text2(1) = rs.Fields("NoInstall")
        Text2(2) = rs.Fields("Dateapp")
        lid2 = Mid(lstLoan.Text, 1, InStr(1, lstLoan, "-") - 1)
    ElseIf textindex = 2 Then
        Text3(0) = rs.Fields("LoanAmt")
        Text3(1) = rs.Fields("NoInstall")
        Text3(2) = rs.Fields("Dateapp")
        lid3 = Mid(lstLoan.Text, 1, InStr(1, lstLoan, "-") - 1)
    End If
End If
rs.Close

  rs.Open "select Balance,NoInstall from  LoanMaster where CompanyName='" & lblCompanyName.Caption & "' and EmpNo=" & Right(frmemployeepaydetails.Text1(0), Len(frmemployeepaydetails.Text1(0)) - InStr(1, frmemployeepaydetails.Text1(0), "/")) & " And loanid = " & Mid(lstLoan.Text, 1, InStr(1, lstLoan, "-") - 1), cnn, adOpenKeyset, adLockOptimistic
  If textindex = 0 Then
        lblbalAmount = rs.Fields(0)
  End If
  If textindex = 1 Then
        lblbalAmount1 = rs.Fields(0)
  End If
  If textindex = 2 Then
        lblbalAmount2 = rs.Fields(0)
  End If
  
rs1.Open "Select count(paid) from loanDetails where paid=yes and CompanyName='" & lblCompanyName.Caption & "' and EmpNo=" & Right(frmemployeepaydetails.Text1(0), Len(frmemployeepaydetails.Text1(0)) - InStr(1, frmemployeepaydetails.Text1(0), "/")) & " And loanid = " & Mid(lstLoan.Text, 1, InStr(1, lstLoan, "-") - 1), cnn, adOpenKeyset, adLockOptimistic

If textindex = 0 Then
    lblBalMonth = rs.Fields(1) - rs1.Fields(0)
End If
If textindex = 1 Then
    lblBalMonth1 = rs.Fields(1) - rs1.Fields(0)
End If
If textindex = 2 Then
    lblBalMonth2 = rs.Fields(1) - rs1.Fields(0)
End If

rs.Close
rs1.Close

If textindex <> -1 Then
      lstLoan.RemoveItem lstLoan.ListIndex
End If
textindex = -1
End Sub
Private Sub Text1_Change(Index As Integer)
If Text1(1) <> "" Then

    lblinstall.Caption = Round(Val(Text1(0)) / Val(Text1(1)), 2)
End If
End Sub
Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

Select Case Index
    Case 0
        If frmemployeepaydetails.AddEditViewMode <> "View" Then
        If KeyCode = vbKeyF2 Then
            textindex = 0
            lstLoan.Visible = True
           
        End If
        End If
End Select


End Sub
Private Sub Text2_Change(Index As Integer)
If Text2(1) <> "" Then
    lblinstall2.Caption = Round(Val(Text2(0)) / Val(Text2(1)), 2)
End If
End Sub
Private Sub Text2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case Index
    Case 0
        If frmemployeepaydetails.AddEditViewMode <> "View" Then
        If KeyCode = vbKeyF2 Then
             textindex = 1
            lstLoan.Visible = True
           
        End If
        End If
End Select
End Sub

Private Sub Text3_Change(Index As Integer)
If Text3(1) <> "" Then
    lblinstall3.Caption = Round(Val(Text3(0)) / Val(Text3(1)), 2)
End If
End Sub
Public Sub addloan()

rs.Open "select LoanId,LoanAmt from LoanMaster where balance<>0 and CompanyName='" & lblCompanyName.Caption & "' and empno=" & Mid(frmemployeepaydetails.cboEmpno, InStr(1, frmemployeepaydetails.cboEmpno, "/") + 1), cnn, adOpenKeyset, adLockOptimistic
'rs.Open "select LoanId,LoanAmt from LoanMaster where empno=" & Mid(frmemployeepaydetails.Text1(0), 5) & " And Balance = 0, cnn, adOpenKeyset, adLockOptimistic"
lstLoan.clear
While Not rs.EOF
    lstLoan.AddItem rs.Fields(0) & "-" & rs.Fields(1)
    rs.MoveNext
Wend
rs.Close
If ListCount > 0 Then
    lstLoan.ListIndex = 0
End If
End Sub
Private Sub Text3_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case Index
    Case 0
        If frmemployeepaydetails.AddEditViewMode <> "View" Then
        If KeyCode = vbKeyF2 Then
            lstLoan.Visible = True
            textindex = 2
        End If
        End If
End Select
End Sub
