VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form User 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Master"
   ClientHeight    =   3525
   ClientLeft      =   3960
   ClientTop       =   2460
   ClientWidth     =   4830
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
   ScaleHeight     =   3525
   ScaleWidth      =   4830
   Begin VB.CommandButton CmdCancel 
      Caption         =   "&Cancel"
      Enabled         =   0   'False
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
      Left            =   4080
      TabIndex        =   10
      Top             =   2400
      Width           =   735
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   18
      Top             =   3255
      Width           =   4830
      _ExtentX        =   8520
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3334
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "3:32 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "10/3/01"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtConfirmPwd 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1800
      Locked          =   -1  'True
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CommandButton CmdDelete 
      Caption         =   "&Delete"
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
      Left            =   4080
      TabIndex        =   8
      Top             =   1440
      Width           =   735
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1800
      TabIndex        =   5
      Top             =   1920
      Width           =   1935
      Begin VB.OptionButton optAdmin 
         Caption         =   "&Admin"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optUser 
         Caption         =   "&User"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   4
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.TextBox txtUser 
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
      Left            =   1800
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   0
      Top             =   600
      Width           =   1935
   End
   Begin VB.TextBox txtPass 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1800
      Locked          =   -1  'True
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton CmdNew 
      Caption         =   "&New"
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
      Left            =   4080
      TabIndex        =   6
      Top             =   480
      Width           =   735
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "&Edit"
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
      Left            =   4080
      TabIndex        =   7
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
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
      Left            =   4080
      TabIndex        =   9
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton CmdExit 
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
      Left            =   4080
      TabIndex        =   11
      Top             =   2880
      Width           =   735
   End
   Begin VB.ComboBox cboUser 
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
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Confirm Password"
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
      Left            =   120
      TabIndex        =   17
      Top             =   1560
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "User Name"
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
      Left            =   120
      TabIndex        =   16
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "PassWord"
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
      Left            =   120
      TabIndex        =   15
      Top             =   1080
      Width           =   780
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "User Master"
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
      Left            =   2400
      TabIndex        =   14
      Top             =   120
      Width           =   1290
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Access"
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
      Left            =   120
      TabIndex        =   13
      Top             =   2040
      Width           =   510
   End
End
Attribute VB_Name = "User"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AddEditViewMode As String
Dim EditUser As String
Dim Msg As String
Dim InputData As String
Private Sub cboUser_Click()
AddEditViewMode = "View"
If cboUser <> "" Then
    txtUser = cboUser
    rs.Open "Select Access  from Login where User ='" & cboUser & "'", cnn, adOpenKeyset, adLockOptimistic
    If Not (rs.EOF) Then
        If rs.Fields("Access") = 0 Then
            optUser.value = True
        Else
            optAdmin.value = True
        End If
    End If
        
    rs.Close
    
    CmdEdit.Enabled = True
    cmdDelete.Enabled = True
End If

End Sub

Private Sub CmdCancel_Click()
AddEditViewMode = "View"
cboUser.ZOrder 0
AddUser
CmdNew.Enabled = True
CmdEdit.Enabled = False
cmdDelete.Enabled = False
txtUser.Locked = True
txtPass.Locked = True
txtConfirmPwd.Locked = True
txtConfirmPwd.Visible = False
optAdmin.value = False
optUser.value = False
Label5.Visible = False
Frame1.Enabled = False
txtUser = ""
txtPass = ""
txtConfirmPwd = ""
Frame1.Top = 1680
Label4.Top = 1680
Cancel = True
cboUser.SetFocus
StatusBar1.Panels(1) = AddEditViewMode
End Sub

Private Sub cmdDelete_Click()
rs2.Open "Select Count(*) from Login where Access='1'", cnn, adOpenKeyset, adLockOptimistic

If rs2.Fields(0) > 1 Or optAdmin.value = False Then
    Msg = MsgBox("Are you sure to Delete this User", vbExclamation + vbYesNo, "Payroll")
    If Msg = 6 Then
        InputData = InputBox("Please Enter the Correct Password", "Payroll")
        rs1.Open "Select pwd from Login where user='" & cboUser & "'", cnn, adOpenKeyset, adLockOptimistic
        If Trim(InputData) <> "" Then
            If Trim(rs1.Fields(0)) = Trim(InputData) Then
                cnn.Execute "Delete from Login Where user='" & cboUser & "'"
                AddUser
                optAdmin.value = False
                optUser.value = False
                CmdNew.Enabled = True
                CmdEdit.Enabled = False
                cmdDelete.Enabled = False
            Else
                MsgBox "Please Enter the Correct Password", vbExclamation, "Payroll"
            
        End If
        End If
        rs1.Close
        
    End If
        
    
    
Else
    MsgBox "Atleast One User Access Admin", vbExclamation, "Payroll"
End If
rs2.Close

End Sub

Private Sub cmdEdit_Click()
AddEditViewMode = "Edit"
EditUser = cboUser
cboUser.ZOrder 1
txtConfirmPwd.Visible = True
Label5.Visible = True
Frame1.Top = 1920
Label4.Top = 1920
CmdNew.Enabled = False
CmdEdit.Enabled = False
cmdDelete.Enabled = False
txtUser.Locked = False
txtPass.Locked = False
txtConfirmPwd.Locked = False
Frame1.Enabled = True
StatusBar1.Panels(1) = AddEditViewMode
Cancel = True
txtUser.SetFocus

End Sub

Private Sub cmdExit_Click()
Unload Me

End Sub

Private Sub CmdNew_Click()
AddEditViewMode = "Add"
EditUser = ""
txtUser.Locked = False
txtPass.Locked = False
txtConfirmPwd.Locked = False
Frame1.Enabled = True
txtUser = ""
txtPass = ""
txtConfirmPwd.Visible = False
Label5.Visible = False
optAdmin.value = False
optUser.value = False
Cancel = True
txtUser.SetFocus
CmdNew.Enabled = False
CmdEdit.Enabled = False
cmdDelete.Enabled = False
cboUser.ZOrder 1
Frame1.Top = 1680
Label4.Top = 1680
StatusBar1.Panels(1) = AddEditViewMode
End Sub
Private Sub CmdSave_Click()

    'Choose Admin Or User
    
If optAdmin.value = False And optUser.value = False Then
    MsgBox "Please Choose Admin or User Access", vbExclamation, "Payroll"
    Exit Sub
End If

 'Check Exists User
 
If LCase(txtUser) <> LCase(EditUser) Then
    rs.Open "Select Count(*) from Login Where user='" & txtUser & "'", cnn, adOpenKeyset, adLockOptimistic
    If rs.Fields(0) > 0 Then
        MsgBox "User Already Exists", vbExclamation, "Payroll"
        Cancel = True
        txtUser.SetFocus
        rs.Close
        Exit Sub
     End If
        rs.Close
End If
 If AddEditViewMode = "Edit" Then
    rs.Open "Select count(*)  From Login where User='" & EditUser & "' and Pwd='" & txtPass & "'", cnn, adOpenKeyset, adLockOptimistic
    If rs.Fields(0) = 0 Then
        MsgBox "Please Enter the Correct Password", vbExclamation, "Payroll"
        rs.Close
        Exit Sub
    End If
    rs.Close
 End If
 
 

If AddEditViewMode = "Add" Then
    rs.Open "Select * from Login", cnn, adOpenKeyset, adLockOptimistic
    rs.AddNew
    rs.Fields("Pwd") = txtPass
ElseIf AddEditViewMode = "Edit" Then
    rs.Open "Select * from Login where user='" & EditUser & "'", cnn, adOpenKeyset, adLockOptimistic
    rs.Fields("Pwd") = txtConfirmPwd
End If

    rs.Fields("User") = txtUser
    
    If optAdmin.value = True Then
        rs.Fields("Access") = 1
    Else
        rs.Fields("Access") = 0
    End If
    rs.Update
    rs.Close

CmdNew.Enabled = True
CmdSave.Enabled = False
CmdEdit.Enabled = False
cmdDelete.Enabled = False
cboUser.ZOrder 0
AddUser
txtUser = ""
txtPass = ""
txtConfirmPwd = ""
txtUser.Locked = True
txtPass.Locked = True
Frame1.Enabled = False
optAdmin.value = False
optUser.value = False
txtConfirmPwd.Visible = False
Label5.Visible = False
Frame1.Top = 1680
Label4.Top = 1680
AddEditViewMode = "View"
StatusBar1.Panels(1) = AddEditViewMode

End Sub
Private Sub Form_Activate()
User.Left = 0
User.Top = 0
txtUser.Locked = True
txtPass.Locked = True
txtConfirmPwd.Locked = True
cboUser.ZOrder 0
AddUser
txtConfirmPwd.Visible = False
Label5.Visible = False
Frame1.Top = 1680
Label4.Top = 1680
CmdEdit.Enabled = False
cmdDelete.Enabled = False
cmdCancel.Enabled = True
cboUser.SetFocus
Frame1.Enabled = False
AddEditViewMode = "View"
StatusBar1.Panels(1) = AddEditViewMode
End Sub
Public Sub AddUser()
cboUser.clear
rs.Open "Select User From Login", cnn, adOpenKeyset, adLockOptimistic
While Not rs.EOF
    cboUser.AddItem rs.Fields("User")
    rs.MoveNext
Wend
rs.Close
End Sub

Private Sub Form_Load()
User.Icon = LoadPicture(App.Path & "\Jamaica my computer.ico")
End Sub

Private Sub txtConfirmPwd_Change()
txtConfirmPwd = Trim(txtConfirmPwd)
CmdSave.Enabled = CheckLengthField
End Sub

Private Sub txtPass_Change()
txtPass = Trim(txtPass)
CmdSave.Enabled = CheckLengthField
End Sub
Private Sub txtUser_Change()
txtUser = Trim(txtUser)
CmdSave.Enabled = CheckLengthField
End Sub
Public Function CheckLengthField() As Boolean
If AddEditViewMode <> "View" Then
If AddEditViewMode = "Add" Then
    If Len(Trim(txtUser)) = 0 Or Len(Trim(txtPass)) = 0 Then
        CheckLengthField = False
        Exit Function
    End If
        
ElseIf AddEditViewMode = "Edit" Then
     If Len(Trim(txtUser)) = 0 Or Len(Trim(txtPass)) = 0 Or Len(Trim(txtConfirmPwd)) = 0 Then
        CheckLengthField = False
        Exit Function
    End If
End If
CheckLengthField = True
End If
End Function

Private Sub txtUser_KeyPress(KeyAscii As Integer)
 gspAlphaNumeric txtUser, KeyAscii
End Sub
