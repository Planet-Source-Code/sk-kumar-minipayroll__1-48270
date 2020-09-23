VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Master"
   ClientHeight    =   3420
   ClientLeft      =   2850
   ClientTop       =   2265
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   5400
   Begin VB.ListBox lstUser 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2205
      Left            =   2280
      TabIndex        =   9
      Top             =   600
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSComctlLib.StatusBar stbGeneral 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   3045
      Width           =   5400
      _ExtentX        =   9525
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin VB.CommandButton Command4 
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
      Left            =   4680
      TabIndex        =   7
      Top             =   2520
      Width           =   625
   End
   Begin VB.CommandButton Command3 
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
      Left            =   4680
      TabIndex        =   6
      Top             =   1560
      Width           =   625
   End
   Begin VB.CommandButton Command2 
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
      Left            =   4680
      TabIndex        =   5
      Top             =   1080
      Width           =   625
   End
   Begin VB.CommandButton Command1 
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
      Left            =   4680
      TabIndex        =   4
      Top             =   600
      Width           =   625
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
      Left            =   2640
      Locked          =   -1  'True
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1080
      Width           =   1935
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
      Height          =   330
      Left            =   2640
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   0
      Top             =   600
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   2640
      TabIndex        =   12
      Top             =   1440
      Width           =   1935
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
         TabIndex        =   14
         Top             =   600
         Width           =   1095
      End
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
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command5 
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
      Left            =   4680
      TabIndex        =   15
      Top             =   2040
      Width           =   625
   End
   Begin VB.Label Label4 
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
      Height          =   375
      Left            =   360
      TabIndex        =   11
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label3 
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
      Height          =   495
      Left            =   1710
      TabIndex        =   8
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label2 
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
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label1 
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
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   1455
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim cnn As New ADODB.Connection
'Dim rs As New ADODB.Recordset
Dim fzSelection As String

Private Sub Command1_Click()
    Command3.Enabled = True
    Command1.Enabled = False
    Command2.Enabled = False
    Command5.Enabled = False
    clear
    txtUser.Locked = False
    txtPass.Locked = False
    Cancel = True
    txtUser.SetFocus
    fzSelection = "N"
End Sub

Private Sub Command2_Click()
    Command3.Enabled = True
    Command1.Enabled = False
    Command2.Enabled = False
    Command5.Enabled = False
    clear
    txtUser.Locked = False
    txtPass.Locked = False
    txtUser.SetFocus
    fzSelection = "E"
    stbGeneral.Panels(2).Picture = LoadPicture(App.Path & "\" & "Lighton.ico")
    stbGeneral.Panels(2).Text = "Press F2 For List/Help"
End Sub

Private Sub Command3_Click()
    Dim flag As Boolean
    stbGeneral.Panels(2).Picture = LoadPicture(App.Path & "\" & "Lightoff.ico")
    If Len(Trim(txtUser)) = 0 Or Len(Trim(txtPass)) = 0 Then
        MsgBox "Enter Missing Data", vbInformation, "Payroll"
        Exit Sub
    End If
    If optAdmin.value = False And optUser.value = False Then
        MsgBox "Select Any Access", vbInformation, "Payroll"
        Exit Sub
    End If
    sql = "Select * From Login"
    If fzSelection = "N" Then
        rs.Open sql, cnn, adOpenKeyset, adLockOptimistic
        For i = 0 To rs.RecordCount - 1
            If UCase(Trim(rs(0))) = UCase(Trim(txtUser)) Then
                flag = False
                MsgBox "UserName Already Exists", vbInformation, "Payroll"
                txtUser = ""
                txtPass = ""
                txtUser.SetFocus
                rs.Close
                Exit Sub
            Else
                flag = True
            End If
            rs.MoveNext
        Next
        If flag = True Then
            rs.AddNew
            rs.Fields(0) = txtUser
            rs.Fields(1) = txtPass
            If optAdmin.value = True Then
                rs.Fields(2) = "1"
            ElseIf optUser.value = True Then
                rs.Fields(2) = "0"
            
            End If
            rs.Update
            rs.Close
            clear
        End If
    End If
    If fzSelection = "E" Then
        If Len(txtUser) <> 0 And Len(txtPass) <> 0 Then
            sql = "Select *From Login Where user='" & txtUser & "'"
            rs.Open sql, cnn, adOpenKeyset, adLockOptimistic
            
            If Not (rs.EOF Or rs.BOF) Then
                rs.Fields(0) = txtUser
                rs.Fields(1) = txtPass
                If optAdmin.value = True Then
                    rs.Fields(2) = "1"
                ElseIf optUser.value = True Then
                    rs.Fields(2) = "0"
                End If
             End If
            rs.Update
            rs.Close
            clear
        Else
            MsgBox "Enter Data ", vbInformation, "Payroll"
            Exit Sub
        End If
    End If
    If fzSelection = "D" Then
        If Len(txtUser) <> 0 And Len(txtPass) <> 0 Then
            sql = "Select *From Login Where user='" & txtUser & "'"
            rs.Open sql, cnn, adOpenKeyset, adLockOptimistic
            
            If Not (rs.EOF Or rs.BOF) Then
                rs.Delete
            End If
'                rs.Fields(0) = txtUser
'                rs.Fields(1) = txtPass
'                If optAdmin.value = True Then
'                    rs.Fields(2) = "1"
'                ElseIf optUser.value = True Then
'                    rs.Fields(2) = "0"
'                End If
'            Else
'                MsgBox ""
'            End If
            rs.Update
            rs.Close
            clear
        Else
            MsgBox "Enter Data ", vbInformation, "Payroll"
            Exit Sub
        End If
    End If
    Command1.Enabled = True
    Command2.Enabled = True
    Command3.Enabled = False
    Command5.Enabled = True
    txtUser.Locked = True
    txtPass.Locked = True
End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

Private Sub Command5_Click()
    Command3.Enabled = True
    Command1.Enabled = False
    Command2.Enabled = False
    Command5.Enabled = False
    clear
    txtUser.Locked = False
    txtPass.Locked = False
    txtUser.SetFocus
    fzSelection = "D"
End Sub

Private Sub Form_Load()
    'cnn.Provider = "Microsoft.Jet.OLEDB.4.0"
    'cnn.Open app.path & "\Moolakadai.mdb"
   ' cnn.Open App.Path & "\Moolakadai.mdb"
    frmUser.Top = 2110
    frmUser.Left = 3155
    stbGeneral.Panels(1).Text = Date
    stbGeneral.Panels(3).Text = Time
End Sub

Private Sub Form_Unload(Cancel As Integer)
fzSelection = ""
    'cnn.Close
End Sub

Private Sub lstUser_DblClick()
    txtUser = lstUser.Text
    sql = "Select *from Login where user='" & lstUser.Text & "'"
    rs.Open sql, cnn, adOpenKeyset, adLockOptimistic
    txtPass = rs(1)
    If rs(2) = "1" Then
        optAdmin.value = True
    ElseIf rs(2) = "0" Then
        optUser.value = True
    End If
    rs.Close
    lstUser.Visible = False
    stbGeneral.Panels(2).Text = ""
End Sub

Private Sub lstUser_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        lstUser.Visible = False
    End If
End Sub

Private Sub txtUser_KeyDown(KeyCode As Integer, Shift As Integer)
   
  'MsgBox App.Path
    lstUser.clear
    If fzSelection = "E" Then
         
        If KeyCode = vbKeyF2 Then
            lstUser.Visible = True
            sql = "Select *from Login"
            rs.Open sql, cnn, adOpenKeyset, adLockOptimistic
            If Not (rs.EOF Or rs.BOF) Then
                For i = 0 To rs.RecordCount - 1
                    lstUser.AddItem rs(0)
                    rs.MoveNext
                Next
            End If
            rs.Close
        End If
    End If
    If fzSelection = "D" Then
         
        If KeyCode = vbKeyF2 Then
            lstUser.Visible = True
            sql = "Select *from Login"
            rs.Open sql, cnn, adOpenKeyset, adLockOptimistic
            If Not (rs.EOF Or rs.BOF) Then
                For i = 0 To rs.RecordCount - 1
                    lstUser.AddItem rs(0)
                    rs.MoveNext
                Next
            End If
            rs.Close
        End If
    End If
    If KeyCode = vbKeyEscape And lstUser.Visible = True Then
        lstUser.Visible = False
        If Len(txtUser) = 0 Then
            Command1.Enabled = True
            Command2.Enabled = True
            Command5.Enabled = True
            Command3.Enabled = False
        End If
            
    End If
End Sub

Private Sub txtUser_KeyPress(KeyAscii As Integer)
   ' MsgBox KeyAscii
    
    gspAlphaNumeric txtUser, KeyAscii
    
End Sub

Private Sub txtUser_LostFocus()
'    If fzSelection = "E" Or fzSelection = "D" Then
'               ' txtUser = lstUser.Text
'                sql = "Select *from Login where user='" & txtUser & "'"
'                rs.Open sql, cnn, adOpenKeyset, adLockOptimistic
'                txtPass = rs(1)
'                If rs(2) = "1" Then
'                    optAdmin.value = True
'                ElseIf rs(2) = "0" Then
'                    optUser.value = True
'                End If
'                rs.Close
'                lstUser.Visible = False
'                stbGeneral.Panels(2).Text = ""
'    End If
End Sub
Private Sub txtUser_Validate(Cancel As Boolean)
    If fzSelection = "N" Then
        If Len(Trim(txtUser)) = 0 Then
            MsgBox "Enter User Name", vbInformation, "Payroll"
            'txtUser.SetFocus
            SendKeys "{Home}+{End}"
            Cancel = True
            Exit Sub
        End If
        
    End If
    If fzSelection = "E" Or fzSelection = "D" Then
        If Len(txtUser) <> 0 Then
               ' txtUser = lstUser.Text
                sql = "Select *from Login where user='" & txtUser & "'"
                rs.Open sql, cnn, adOpenKeyset, adLockOptimistic
                If Not (rs.EOF Or rs.BOF) Then
                    txtPass = rs(1)
                    If rs(2) = "1" Then
                        optAdmin.value = True
                    ElseIf rs(2) = "0" Then
                        optUser.value = True
                    End If
                    rs.Close
                    lstUser.Visible = False
                    stbGeneral.Panels(2).Text = ""
                Else
                    MsgBox "Invalid User Name", vbInformation, "Payroll"
                    txtUser = ""
                    txtUser.SetFocus
                    rs.Close
                    Exit Sub
                End If
            End If
    End If
End Sub
