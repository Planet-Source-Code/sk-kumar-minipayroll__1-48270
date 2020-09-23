VERSION 5.00
Begin VB.Form frmAdvance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Advance Report"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9900
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
   ScaleHeight     =   3405
   ScaleWidth      =   9900
   Begin VB.PictureBox rep1 
      Height          =   480
      Left            =   390
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   8
      Top             =   105
      Width           =   1200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show"
      Height          =   675
      Left            =   4020
      TabIndex        =   6
      Top             =   2385
      Width           =   2355
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
      Left            =   1080
      TabIndex        =   2
      Top             =   960
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
      Left            =   5550
      TabIndex        =   0
      Top             =   870
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
         TabIndex        =   1
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Label lblCompanyName 
      AutoSize        =   -1  'True
      Height          =   225
      Left            =   4920
      TabIndex        =   7
      Top             =   270
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Company Name"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   3000
      TabIndex        =   5
      Top             =   270
      Width           =   1260
   End
End
Attribute VB_Name = "frmAdvance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

rep1.ReportFileName = App.Path & "\Advance.rpt"
If optAll.value = True Then
    rep1.SelectionFormula = "{LoanMaster.CompanyName}='" & lblCompanyName.Caption & "'"
Else
    rep1.SelectionFormula = "{LoanMaster.CompanyName}='" & lblCompanyName.Caption & "' and {Branch.BranchCode}='" & cboBranch & "'"
    
End If

rep1.Action = True
   
End Sub

Private Sub Form_Activate()
frmAdvance.Left = 0
frmAdvance.Top = 0

End Sub

Private Sub Form_Load()
lblCompanyName.Caption = CompanyName
frmAdvance.Icon = LoadPicture(App.Path & "\Jamaica my computer.ico")

rs.Open "Select * From Branch Where CompanyName='" & lblCompanyName.Caption & "'", cnn, adOpenKeyset, adLockOptimistic
While Not rs.EOF
    cboBranch.AddItem rs(0)
    rs.MoveNext
Wend
rs.Close
optAll.value = True
End Sub
Private Sub optAll_Click()
fraBranch.Visible = False

End Sub
Private Sub optPart_Click()
fraBranch.Visible = True
End Sub
