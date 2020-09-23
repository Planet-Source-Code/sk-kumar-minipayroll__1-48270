VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmResEmpMaster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resign EmpMaster Report"
   ClientHeight    =   7335
   ClientLeft      =   975
   ClientTop       =   675
   ClientWidth     =   10320
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
   ScaleHeight     =   7335
   ScaleWidth      =   10320
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   345
      Left            =   2640
      TabIndex        =   22
      Top             =   6960
      Width           =   735
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   345
      Left            =   4920
      TabIndex        =   21
      Top             =   6960
      Width           =   735
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   345
      Left            =   3840
      TabIndex        =   20
      Top             =   6960
      Width           =   735
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6135
      Left            =   360
      TabIndex        =   23
      Top             =   600
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   10821
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "E&mployee Details"
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label22"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label17"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label4"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label5"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label6"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label7"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label18"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label19"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtDA"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtDoB"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cboResignNo"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtEmpno"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtBasic"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "TxtDoJ"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtEname"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "List1"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtDesignation"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtBranch"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtResDate"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).ControlCount=   22
      TabCaption(1)   =   "Pe&rsonal Details"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtOthers"
      Tab(1).Control(1)=   "txtAdd1"
      Tab(1).Control(2)=   "txtAdd2"
      Tab(1).Control(3)=   "txtCity"
      Tab(1).Control(4)=   "txtState"
      Tab(1).Control(5)=   "txtPincode"
      Tab(1).Control(6)=   "txtPfNo"
      Tab(1).Control(7)=   "txtPhone"
      Tab(1).Control(8)=   "txtFather"
      Tab(1).Control(9)=   "txtEsiNo"
      Tab(1).Control(10)=   "Label23"
      Tab(1).Control(11)=   "Label8"
      Tab(1).Control(12)=   "Label9"
      Tab(1).Control(13)=   "Label10"
      Tab(1).Control(14)=   "Label11"
      Tab(1).Control(15)=   "Label12"
      Tab(1).Control(16)=   "label13"
      Tab(1).Control(17)=   "label15(10)"
      Tab(1).Control(18)=   "Label16"
      Tab(1).Control(19)=   "label14"
      Tab(1).ControlCount=   20
      Begin VB.TextBox txtOthers 
         BackColor       =   &H00C0FFFF&
         Height          =   345
         Left            =   -71520
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   10
         Top             =   480
         Width           =   2055
      End
      Begin VB.TextBox txtResDate 
         BackColor       =   &H00C0FFFF&
         Height          =   345
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   5
         Top             =   3195
         Width           =   2055
      End
      Begin VB.TextBox txtBranch 
         BackColor       =   &H00C0FFFF&
         Height          =   345
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   3
         Top             =   2184
         Width           =   2055
      End
      Begin VB.TextBox txtDesignation 
         BackColor       =   &H00C0FFFF&
         Height          =   345
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   2640
         Width           =   3015
      End
      Begin VB.ListBox List1 
         Height          =   4560
         Left            =   6600
         TabIndex        =   24
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox txtEname 
         BackColor       =   &H00C0FFFF&
         Height          =   345
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   23
         TabIndex        =   2
         Top             =   1656
         Width           =   3015
      End
      Begin VB.TextBox TxtDoJ 
         BackColor       =   &H00C0FFFF&
         Height          =   345
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   7
         Top             =   4296
         Width           =   2055
      End
      Begin VB.TextBox txtBasic 
         BackColor       =   &H00C0FFFF&
         Height          =   345
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   8
         Top             =   4824
         Width           =   2055
      End
      Begin VB.TextBox txtEmpno 
         BackColor       =   &H00C0FFFF&
         Height          =   345
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   1
         Top             =   1128
         Width           =   2055
      End
      Begin VB.TextBox txtAdd1 
         BackColor       =   &H00C0FFFF&
         Height          =   345
         Left            =   -71520
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   12
         Top             =   1665
         Width           =   2055
      End
      Begin VB.TextBox txtAdd2 
         BackColor       =   &H00C0FFFF&
         Height          =   345
         Left            =   -71520
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   13
         Top             =   2250
         Width           =   2055
      End
      Begin VB.TextBox txtCity 
         BackColor       =   &H00C0FFFF&
         Height          =   345
         Left            =   -71520
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   14
         Top             =   2835
         Width           =   2055
      End
      Begin VB.TextBox txtState 
         BackColor       =   &H00C0FFFF&
         Height          =   345
         Left            =   -71520
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   15
         Top             =   3420
         Width           =   2055
      End
      Begin VB.TextBox txtPincode 
         BackColor       =   &H00C0FFFF&
         Height          =   345
         Left            =   -71520
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   16
         Top             =   4005
         Width           =   2055
      End
      Begin VB.TextBox txtPfNo 
         BackColor       =   &H00C0FFFF&
         Height          =   345
         Left            =   -71520
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   17
         Top             =   4590
         Width           =   2055
      End
      Begin VB.TextBox txtPhone 
         BackColor       =   &H00C0FFFF&
         Height          =   345
         Left            =   -71520
         Locked          =   -1  'True
         MaxLength       =   35
         TabIndex        =   19
         Top             =   5760
         Width           =   2055
      End
      Begin VB.TextBox txtFather 
         BackColor       =   &H00C0FFFF&
         Height          =   345
         Left            =   -71520
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   11
         Top             =   1080
         Width           =   2055
      End
      Begin VB.ComboBox cboResignNo 
         Height          =   345
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox txtDoB 
         BackColor       =   &H00C0FFFF&
         Height          =   345
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   6
         Top             =   3768
         Width           =   2055
      End
      Begin VB.TextBox txtEsiNo 
         BackColor       =   &H00C0FFFF&
         Height          =   345
         Left            =   -71520
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   18
         Top             =   5175
         Width           =   2055
      End
      Begin VB.TextBox txtDA 
         BackColor       =   &H00C0FFFF&
         Height          =   345
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   9
         Top             =   5352
         Width           =   2055
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Others"
         Height          =   225
         Left            =   -73920
         TabIndex        =   45
         Top             =   600
         Width           =   585
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "ResignDate"
         Height          =   225
         Left            =   240
         TabIndex        =   44
         Top             =   3240
         Width           =   960
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Branch Name"
         Height          =   225
         Left            =   240
         TabIndex        =   43
         Top             =   2160
         Width           =   1125
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "LIST Of EMPLOYEES :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6600
         TabIndex        =   42
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "BASIC"
         Height          =   225
         Left            =   240
         TabIndex        =   41
         Top             =   4920
         Width           =   585
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "DOJ"
         Height          =   225
         Left            =   240
         TabIndex        =   40
         Top             =   4455
         Width           =   390
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "DOB"
         Height          =   225
         Left            =   240
         TabIndex        =   39
         Top             =   3885
         Width           =   420
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Designation"
         Height          =   225
         Left            =   240
         TabIndex        =   38
         Top             =   2715
         Width           =   990
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "EName"
         Height          =   225
         Left            =   240
         TabIndex        =   37
         Top             =   1665
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "EmpNo."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   36
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Add1"
         Height          =   225
         Left            =   -73920
         TabIndex        =   35
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Add2"
         Height          =   225
         Left            =   -73920
         TabIndex        =   34
         Top             =   2400
         Width           =   420
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "City"
         Height          =   225
         Left            =   -73920
         TabIndex        =   33
         Top             =   2880
         Width           =   345
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "State"
         Height          =   225
         Left            =   -73920
         TabIndex        =   32
         Top             =   3480
         Width           =   420
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "PinCode"
         Height          =   225
         Left            =   -73920
         TabIndex        =   31
         Top             =   4080
         Width           =   690
      End
      Begin VB.Label label13 
         AutoSize        =   -1  'True
         Caption         =   "PFNo"
         Height          =   225
         Left            =   -73920
         TabIndex        =   30
         Top             =   4800
         Width           =   450
      End
      Begin VB.Label label15 
         AutoSize        =   -1  'True
         Caption         =   "Phone"
         Height          =   225
         Index           =   10
         Left            =   -73920
         TabIndex        =   29
         Top             =   5880
         Width           =   510
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Father/Husband Name"
         Height          =   225
         Left            =   -73920
         TabIndex        =   28
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "ResignNo"
         Height          =   225
         Left            =   240
         TabIndex        =   27
         Top             =   645
         Width           =   810
      End
      Begin VB.Label label14 
         AutoSize        =   -1  'True
         Caption         =   "ESI No."
         Height          =   225
         Left            =   -73920
         TabIndex        =   26
         Top             =   5280
         Width           =   600
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "DA"
         Height          =   225
         Left            =   240
         TabIndex        =   25
         Top             =   5400
         Width           =   270
      End
   End
   Begin VB.Label lblCompanyName 
      AutoSize        =   -1  'True
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   4680
      TabIndex        =   47
      Top             =   240
      Width           =   45
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "CompanyName"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   3000
      TabIndex        =   46
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmResEmpMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Msg As String
Public Sub AddResignNo()
rs.Open "Select * from ResEmpMaster Where CompanyName='" & lblCompanyName.Caption & "' order by resignNo", cnn, adOpenKeyset, adLockOptimistic
cboResignNo.clear
List1.clear
While Not rs.EOF
cboResignNo.AddItem rs.Fields("ResignNo")
List1.AddItem rs.Fields("ResignNo") & "-" & rs.Fields("Ename")
rs.MoveNext
Wend
rs.Close
If cboResignNo.ListCount > 0 Then
    cboResignNo.ListIndex = 0
End If
End Sub
Private Sub cboResignNo_Click()
rs.Open "Select * from ResEmpMaster where CompanyName ='" & CompanyName & "' and ResignNo=" & cboResignNo, cnn, adOpenKeyset, adLockOptimistic
txtEmpNo = SubName & Format(rs.Fields("Empno"), "000")
txtEname = rs.Fields("Ename")
txtDesignation = rs.Fields("Designation")
txtDoB = rs.Fields("Dob")
TxtDoJ = rs.Fields("Doj")
txtBasic = rs.Fields("Basic")
txtAdd1 = rs.Fields("Add1")
txtAdd2 = rs.Fields("Add2")
txtCity = rs.Fields("City")
txtState = rs.Fields("State")
txtPincode = rs.Fields("Pincode")
txtPfNo = rs.Fields("PfNo")
txtPhone = rs.Fields("Phone")
txtFather = rs.Fields("Father_Husband")
txtOthers = rs.Fields("Others")
txtEsiNo = rs.Fields("Esi")
txtBranch = rs.Fields("Branchcode")
txtResDate = rs.Fields("ResignDate")
rs.Close
rs.Open "Select Da from DesignationMaster where Designation='" & txtDesignation & "'", cnn, adOpenKeyset, adLockOptimistic
If Not rs.EOF Then
    txtDA = rs.Fields("Da")
End If

rs.Close
SSTab1.Tab = 0

End Sub
Private Sub Command3_Click()

End Sub
Private Sub CmdDelete_Click()
Msg = MsgBox("Are you sure To Delete This Record", vbExclamation + vbYesNo, "Payroll")
If Msg = 6 Then
    cnn.Execute "Delete from ResEmpMaster where CompanyName='" & lblCompanyName.Caption & "' and resignNo=" & cboResignNo
    clear
    AddResignNo
End If
If cboResignNo.ListCount > 0 Then
    cboResignNo.ListIndex = 0
    cmdDelete.Enabled = True
Else
    cmdDelete.Enabled = False
End If

End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub CmdPrint_Click()
Msg = MsgBox("Are you take Print", vbExclamation + vbYesNo, "Payroll")
On Error GoTo err
If Msg = 6 Then
'Printer.NewPage
Printer.Orientation = 1
Printer.Font.Size = 12
Printer.Font.Name = "Courier New"
Printer.FontBold = True
Printer.Print ""
Printer.Print ""
Printer.Print Space(10) & CompanyName

Printer.Print ""
Printer.Print Space(17) & "Employee Details"
Printer.FontBold = False
Printer.Font.Size = 10
Printer.Print Space(21) & String(18, "*")
Printer.Print ""
Printer.Print Space(5) & "Empno                 :" & Space(10) & txtEmpNo
Printer.Print ""
Printer.Print Space(5) & "Ename                 :" & Space(10) & UCase(txtEname)
Printer.Print ""
Printer.Print Space(5) & "Father/Husband Name   :" & Space(10) & UCase(txtFather)
Printer.Print ""
Printer.Print Space(5) & "Designation           :" & Space(10) & UCase(txtDesignation)
Printer.Print ""
Printer.Print Space(5) & "Date of Birth         :" & Space(10) & Format(txtDoB, "MMM/DD/YYYY")
Printer.Print ""
Printer.Print Space(5) & "Date of Joining       :" & Space(10) & Format(TxtDoJ, "MMM/DD/YYYY")
Printer.Print ""
Printer.Print Space(5) & "Resign Date           :" & Space(10) & Format(txtResDate, "MMM/DD/YYYY")
Printer.Print ""
Printer.Print Space(5) & "Basic                 :" & Space(10) & Format(txtBasic, "0.00")
Printer.Print ""
Printer.Print Space(5) & "DA                    :" & Space(10) & Format(txtDA, "0.00")
Printer.Print ""
Printer.Print Space(5) & "Others                :" & Space(10) & Format(txtOthers, "0.00")
Printer.Print ""
Printer.Print Space(5) & "Address1              :" & Space(10) & txtAdd1
Printer.Print ""
Printer.Print Space(5) & "Address2              :" & Space(10) & txtAdd2
Printer.Print ""
Printer.Print Space(5) & "City                  :" & Space(10) & UCase(txtCity)
Printer.Print ""
Printer.Print Space(5) & "State                 :" & Space(10) & UCase(txtState)
Printer.Print ""
Printer.Print Space(5) & "Pincode               :" & Space(10) & txtPincode
Printer.Print ""
Printer.Print Space(5) & "PFNo                  :" & Space(10) & txtPfNo
Printer.Print ""
Printer.Print Space(5) & "ESINo                 :" & Space(10) & txtEsiNo
Printer.Print ""
Printer.Print Space(5) & "Phone No.             :" & Space(10) & txtPhone

Printer.Print ""
Printer.EndDoc
err:
        If err.Number = 482 Or err.Number = 484 Then
            MsgBox "Make Sure The printer Is Ready", vbExclamation, "Payroll"
        End If
        Exit Sub
End If
End Sub
Private Sub Form_Activate()
frmResEmpMaster.Left = 0
frmResEmpMaster.Top = 0
If cboResignNo.ListCount > 0 Then
    cboResignNo.ListIndex = 0
    cmdDelete.Enabled = True
Else
    cmdDelete.Enabled = False
End If

End Sub
Private Sub Form_Load()
lblCompanyName.Caption = CompanyName
AddResignNo
frmResEmpMaster.Icon = LoadPicture(App.Path & "\Jamaica my computer.ico")

End Sub
Private Sub List1_Click()
cboResignNo = Mid(List1.Text, 1, InStr(1, List1.Text, "-") - 1)
End Sub

