VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmLoan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loan Master"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   5280
   Begin VB.CommandButton Command2 
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
      Height          =   495
      Left            =   4080
      TabIndex        =   14
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox txtEmpName 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   12
      Top             =   2040
      Width           =   1095
   End
   Begin MSMask.MaskEdBox mskDateApp 
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   2760
      Width           =   1815
      _ExtentX        =   3201
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
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtReason 
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
      TabIndex        =   5
      Top             =   3960
      Width           =   1815
   End
   Begin VB.TextBox txtInstall 
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
      TabIndex        =   4
      Top             =   3360
      Width           =   1815
   End
   Begin VB.TextBox txtLoan 
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
      TabIndex        =   2
      Top             =   2160
      Width           =   1815
   End
   Begin VB.ComboBox cboEmpNo 
      Height          =   315
      Left            =   1800
      TabIndex        =   0
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Emp name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Top             =   1560
      Width           =   885
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Reason"
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
      TabIndex        =   11
      Top             =   3960
      Width           =   570
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "No Of Instalments"
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
      Top             =   3360
      Width           =   1440
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Date Of Approval"
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
      TabIndex        =   9
      Top             =   2760
      Width           =   1350
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Loan Amount"
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
      Top             =   2160
      Width           =   1065
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Employee No :"
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
      TabIndex        =   7
      Top             =   960
      Width           =   1110
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
      Height          =   285
      Left            =   2760
      TabIndex        =   6
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmLoan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Private Sub cboEmpNo_Click()
    sql = "select ename from admin where empno=" & cboEmpNo
    rs.Open sql, cnn, adOpenKeyset, adLockOptimistic
    txtEmpName = rs(0)
    rs.Close
    txtLoan.SetFocus
End Sub

Private Sub Command1_Click()
    
    sql1 = "select  max(LoanId) as [pp] from loanMaster where empno=" & cboEmpNo
    sql = "Select *from LoanMaster"
    rs.Open sql, cnn, adOpenKeyset, adLockOptimistic
    rs1.Open sql1, cnn, adOpenKeyset, adLockOptimistic
    rs.AddNew
    If IsNull(rs1.Fields(0)) Then
        rs(0) = 1
    Else
        rs.Fields(0) = Val(rs1(0)) + 1
    End If
    
    rs.Fields(1) = cboEmpNo
    rs.Fields(2) = txtLoan
    rs.Fields(3) = mskDateApp
    rs.Fields(4) = txtInstall
    rs.Fields(5) = txtReason
    rs.Fields(6) = txtLoan
    
    rs.Update
    'rs1.Update
    rs.Close
    rs1.Close
    clear
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
cnn.Provider = "Microsoft.Jet.OLEDB.4.0"
cnn.Open "d:\project\Moolakadai.mdb"
sql = "Select * From Admin"
rs.Open sql, cnn, adOpenKeyset, adLockOptimistic
If Not (rs.EOF Or rs.BOF) Then
    For i = 0 To rs.RecordCount - 1
        cboEmpNo.AddItem rs.Fields(0)
        rs.MoveNext
    Next
End If
rs.Close
mskDateApp.text = "__/__/____"
End Sub
