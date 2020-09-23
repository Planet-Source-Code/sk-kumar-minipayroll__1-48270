VERSION 5.00
Begin VB.Form Frmadminlogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Admin"
   ClientHeight    =   1995
   ClientLeft      =   3750
   ClientTop       =   2775
   ClientWidth     =   4335
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   4335
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   2520
      TabIndex        =   6
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Clear"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   1440
      TabIndex        =   5
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   360
      TabIndex        =   4
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   1560
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "admin"
      Top             =   840
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1560
      MaxLength       =   25
      TabIndex        =   0
      Text            =   "admin"
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   900
   End
End
Attribute VB_Name = "Frmadminlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** Dim cnn As New ADODB.Connection
'*** Dim rs As New ADODB.Recordset
Private Sub Command1_Click(Index As Integer)

Select Case Index
    
    Case 0
        
        rs.Open "select * from adminlogin", cnn, adOpenKeyset, adLockOptimistic
            'While Not rs.EOF
            If Not (rs.EOF Or rs.BOF) Then
                For i = 0 To rs.RecordCount - 1
                    If Text1(0) = rs.Fields(0) And Text1(1) = rs.Fields(1) Then
                        
                        Unload Me
                        flag = True
                        rs.Close
                        frmEmployeeDetails.Show
                        Exit Sub
                    Else
                        flag = False
                        rs.MoveNext
                    End If
             Next
            '    Wend
            End If
            If Not flag Then
                MsgBox "Invalid UserName/Password", vbInformation, "Payroll"
                
            End If
            rs.Close
            
    Case 1
            
            Text1(0).Text = ""
            Text1(1).Text = ""
            Text1(0).SetFocus
            
    Case 2
           
            Unload Me
            
End Select
    
    'cnn.Close
End Sub
Private Sub Form_Load()
'***cnn.Provider = "Microsoft.Jet.OLEDB.4.0"
'***cnn.Open App.Path & "/Moolakadai.mdb" ' payroll1.mdb"
'Command1(0).Enabled = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
'***cnn.Close
End Sub

Private Sub Text1_Change(Index As Integer)
Text1(Index) = Trim(Text1(Index))
If Len(Trim(Text1(0))) = 0 Or Len(Trim(Text1(1))) = 0 Then
    Command1(0).Enabled = False
    Exit Sub
End If
Command1(0).Enabled = True

End Sub
Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Text1(Index).Index <> Text1.UBound Then
            Text1(Index + 1).SetFocus
    End If
End If
End Sub

