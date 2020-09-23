VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBranch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Branch Master"
   ClientHeight    =   3600
   ClientLeft      =   3360
   ClientTop       =   2640
   ClientWidth     =   5430
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00404040&
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   5430
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
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
      Left            =   1800
      TabIndex        =   12
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Save"
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
      Left            =   2640
      TabIndex        =   5
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
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
      Left            =   3480
      TabIndex        =   6
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
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
      Left            =   4440
      TabIndex        =   7
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton Command0 
      Caption         =   "&New"
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
      TabIndex        =   3
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Edit"
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
      Left            =   960
      TabIndex        =   4
      Top             =   2760
      Width           =   735
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   3225
      Width           =   5430
      _ExtentX        =   9578
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4392
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "12:41 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "10/15/01"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtCode 
      Height          =   345
      Left            =   2040
      MaxLength       =   5
      TabIndex        =   1
      Top             =   2160
      Width           =   2175
   End
   Begin VB.TextBox txtBranch 
      Height          =   345
      Left            =   2040
      MaxLength       =   30
      TabIndex        =   0
      Top             =   1560
      Width           =   2175
   End
   Begin VB.ComboBox cbobranch 
      Height          =   345
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "CompanyName"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   120
      TabIndex        =   14
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lblCompanyName 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   1920
      TabIndex        =   13
      Top             =   840
      Width           =   60
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Branch Code"
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
      TabIndex        =   10
      Top             =   2160
      Width           =   1020
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Branch Master"
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
      Left            =   1890
      TabIndex        =   9
      Top             =   120
      Width           =   1545
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Branch Name"
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
      TabIndex        =   8
      Top             =   1560
      Width           =   1065
   End
End
Attribute VB_Name = "frmBranch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***Dim cnn As New ADODB.Connection
'***Dim rs As New ADODB.Recordset
'***Dim rs1 As New ADODB.Recordset
Dim AddEditViewMode As String
Dim EditBranchName As String
Dim EditBranchCode As String
Dim flag  As Boolean
Dim Msg As String
Private Sub cbobranch_Click()
If cboBranch <> "" Then
    txtBranch = cboBranch
    rs1.Open "Select BranchCode From Branch Where  CompanyName='" & lblCompanyName & "' and BranchName='" & cboBranch & "'", cnn, adOpenKeyset, adLockOptimistic
    txtCode = rs1.Fields(0)
    rs1.Close
    rs1.Open "Select count(*) from Admin where CompanyName='" & lblCompanyName & "' and  BranchCode='" & txtCode & "'", cnn, adOpenKeyset, adLockOptimistic
    If rs1.Fields(0) > 0 Then
       CmdDelete.Enabled = False
    Else
        CmdDelete.Enabled = True
    End If
    rs1.Close
End If
End Sub
Private Sub CmdDelete_Click()
AddEditViewMode = "View"
Msg = MsgBox("Are you sure Delete this Record", vbExclamation + vbYesNo, "Payroll")
If Msg = 6 Then
   ' cnn.Execute "Insert into Delbranch Select * from Branch Where BranchCode='" & txtCode & "'"
    'cnn.Execute "Insert into DelAdmin Select * from Admin Where BranchCode='" & txtCode & "'"
    
   ' rs.Open "Select empno from admin Where BranchCode='" & txtCode & "'", cnn, adOpenKeyset, adLockOptimistic
    'While Not rs.EOF
     '   cnn.Execute "Insert into DelDaughters Select * from Daughters Where Empno=" & rs.Fields(0)
      '  rs.MoveNext
    'Wend
    'rs.Close
    cnn.Execute "Delete From Branch where CompanyName='" & lblCompanyName & "' and Branchcode='" & txtCode & "'"
End If
txtCode = ""
txtBranch = ""
CountRecord
End Sub
Private Sub Command0_Click() 'New Button
AddEditViewMode = "Add"
Command0.Enabled = False
Command1.Enabled = False
Command4.Enabled = False
CmdDelete.Enabled = False
txtBranch.ZOrder 0
txtBranch.SetFocus
TextLock (False)
StatusBar1.Panels(1) = AddEditViewMode
clear
End Sub
Private Sub Command1_Click()     'Save Button
    
 If AddEditViewMode = "Add" Then
    rs.Open "Select * From Branch", cnn, adOpenKeyset, adLockOptimistic
    rs.AddNew
 ElseIf AddEditViewMode = "Edit" Then
    rs.Open "Select * From Branch Where CompanyName='" & lblCompanyName & "' and BranchName ='" & EditBranchName & "'", cnn, adOpenKeyset, adLockOptimistic
 End If
 
 rs.Fields(0) = txtCode
 rs.Fields(1) = txtBranch
 rs.Fields(2) = lblCompanyName.Caption
 rs.Update
 rs.Close
 'rs1.Close
 Command1.Enabled = False
 Command0.Enabled = True
 Command4.Enabled = True
 CmdDelete.Enabled = True
 AddEditViewMode = "View"
 cboBranch.ZOrder 0
 AddRecToCombo
 TextLock (True)
StatusBar1.Panels(1) = AddEditViewMode
End Sub
Private Sub Command2_Click() 'Clear Button
    AddEditViewMode = "View"
    clear
    TextLock (True)
    CountRecord
    Command1.Enabled = False
    StatusBar1.Panels(1) = AddEditViewMode
End Sub
Private Sub Command3_Click() 'Exit Button
    Unload Me
End Sub
Private Sub Command4_Click()  'Edit Button
    AddEditViewMode = "Edit"
    TextLock (False)
    Command0.Enabled = False
    Command1.Enabled = True
    Command4.Enabled = False
    CmdDelete.Enabled = False
    txtBranch.ZOrder 0
    txtBranch.SetFocus
    EditBranchName = txtBranch
    EditBranchCode = txtCode
End Sub
Private Sub Form_Load()
frmBranch.Top = 0
frmBranch.Left = 0
frmBranch.Icon = LoadPicture(App.Path & "\Jamaica my computer.ico")
AddEditViewMode = "View"
lblCompanyName.Caption = CompanyName
CountRecord
TextLock (True)
Command1.Enabled = False
 StatusBar1.Panels(1) = AddEditViewMode
End Sub
Private Sub Form_Unload(Cancel As Integer)
  '***  cnn.Close
End Sub
Public Sub AddRecToCombo()
rs1.Open "Select BranchName From Branch where CompanyName='" & lblCompanyName & "'", cnn, adOpenKeyset, adLockOptimistic
cboBranch.clear
While Not rs1.EOF
    cboBranch.AddItem rs1.Fields(0)
    rs1.MoveNext
Wend
rs1.Close
If cboBranch.ListCount > 0 Then
    cboBranch.ListIndex = 0
End If
End Sub
Private Sub txtBranch_Change()
txtBranch = LTrim(txtBranch)

If AddEditViewMode <> "View" Then
    
    If LCase(EditBranchName) <> LCase(txtBranch) Then
        If AddEditViewMode = "Add" Then
            rs.Open "Select Count(*) From Branch Where CompanyName='" & lblCompanyName.Caption & "' and BranchCode='" & txtCode & "' or  CompanyName='" & lblCompanyName.Caption & "' and BranchName='" & txtBranch & "'", cnn, adOpenKeyset, adLockOptimistic
        Else
            rs.Open "Select Count(*) From Branch Where CompanyName='" & lblCompanyName.Caption & "' and BranchCode='" & txtCode & "' and BranchName='" & txtBranch & "'", cnn, adOpenKeyset, adLockOptimistic
        End If
        
        If rs.Fields(0) > 0 Then
            StatusBar1.Panels(1) = "Already Exists-BranchName"
            Command1.Enabled = False
            rs.Close
            flag = True
            Exit Sub
        End If
    rs.Close
    flag = False
   Else
    flag = False
   End If
   If flag = False Then
    Command1.Enabled = CheckZeroLength(txtBranch, txtCode)
    StatusBar1.Panels(1) = AddEditViewMode
   End If
   End If
End Sub
Public Function CheckZeroLength(First As String, Second As String) As Boolean
If Trim(Len(First)) = 0 Or Trim(Len(Second)) = 0 Then
    CheckZeroLength = False
Else
    CheckZeroLength = True
End If

End Function
Private Sub txtBranch_KeyPress(KeyAscii As Integer)
gspAlphaNumeric txtBranch, KeyAscii
End Sub
Private Sub txtCode_Change()
txtCode = LTrim(txtCode)
If AddEditViewMode <> "View" Then

    If LCase(EditBranchCode) <> LCase(txtCode) Then
        If AddEditViewMode = "Add" Then
            rs.Open "Select Count(*) From Branch Where CompanyName='" & lblCompanyName & "' and BranchCode='" & txtCode & "' or CompanyName='" & lblCompanyName & "' and BranchName='" & txtBranch & "'", cnn, adOpenKeyset, adLockOptimistic
            Else
            rs.Open "Select Count(*) From Branch Where CompanyName='" & lblCompanyName & "' and BranchCode='" & txtCode & "'", cnn, adOpenKeyset, adLockOptimistic
        End If
     'rs.Open "Select Count(*) From Branch Where BranchCode='" & txtCode & "' or BranchName='" & txtBranch & "'", cnn, adOpenKeyset, adLockOptimistic
        If rs.Fields(0) > 0 Then
            StatusBar1.Panels(1) = "Already Exists -BranchCode or BranchName"
            Command1.Enabled = False
            flag = True
            rs.Close
            Exit Sub
        End If
      rs.Close
     flag = False
   Else
        flag = False
   End If
    If flag = False Then
        Command1.Enabled = CheckZeroLength(txtBranch, txtCode)
        StatusBar1.Panels(1) = AddEditViewMode
   End If
End If
End Sub
Public Sub CountRecord()
rs.Open "Select Count(*) From  Branch where CompanyName='" & lblCompanyName.Caption & "'", cnn, adOpenKeyset, adLockOptimistic
cboBranch.clear
If rs.Fields(0) > 0 Then
    cboBranch.ZOrder 0
    AddRecToCombo
    Command0.Enabled = True
    Command4.Enabled = True
    'cmdDelete.Enabled = True
Else
    txtBranch.ZOrder 0
    Command0.Enabled = True
    Command1.Enabled = False
    Command4.Enabled = False
    CmdDelete.Enabled = False
End If
rs.Close
End Sub
Public Sub TextLock(Locked As Boolean)
txtBranch.Locked = Locked
txtCode.Locked = Locked
End Sub
Private Sub txtCode_KeyPress(KeyAscii As Integer)
If AddEditViewMode <> "View" Then
    CheckSpecialChar txtCode, KeyAscii
End If
End Sub
