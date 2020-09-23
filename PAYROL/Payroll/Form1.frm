VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmlogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   1320
   ClientLeft      =   2595
   ClientTop       =   1875
   ClientWidth     =   4020
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   4020
   StartUpPosition =   2  'CenterScreen
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
      Height          =   330
      Index           =   0
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   0
      Top             =   120
      Width           =   1935
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
      Height          =   330
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   1320
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   600
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3240
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
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
      Index           =   1
      Left            =   3360
      TabIndex        =   3
      Top             =   600
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   240
      Top             =   960
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   960
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
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
      Height          =   375
      Index           =   0
      Left            =   3360
      TabIndex        =   2
      Top             =   120
      Width           =   615
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
      TabIndex        =   5
      Top             =   600
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
      TabIndex        =   4
      Top             =   120
      Width           =   900
   End
End
Attribute VB_Name = "frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'***Dim cnn As New ADODB.Connection
'***Dim rs As New ADODB.Recordset
Private Sub Command1_Click(Index As Integer)
Str1 = ""
Select Case Index
    
    Case 0
        
        rs.Open "select * from login", cnn, adOpenKeyset, adLockOptimistic
        flag = False
        While Not rs.EOF
            If Text1(0) = rs.Fields(0) And Text1(1) = rs.Fields(1) Then
                             
                ProgressBar1.value = 0
                Timer1.Enabled = True
                Timer1_Timer
                'Unload Me
               ' Form2.Show vbModal
               If rs.Fields(2) <> "1" Then
                    MDIForm1.mnuMaster.Visible = False
                    MDIForm1.mnurangepfesi.Visible = False
                End If
              flag = True
            End If
           rs.MoveNext
           Wend
           
        If flag = False Then
            MsgBox "Invalid UserName/Password Try Again", vbExclamation, "Payroll"
        End If
   

        rs.Close
            
    
    Case 1
            'cnn.Close
            Unload Me
            
End Select

End Sub
Private Sub Form_Load()

    frmlogin.Icon = LoadPicture(App.Path & "\Jamaica my computer.ico")
    cnn.Provider = "Microsoft.Jet.OLEDB.4.0"
    cnn.Open App.Path & "\Moolakadai.mdb" ' payroll1.mdb"
    Command1(0).Enabled = False
    Timer1.Enabled = False
    ProgressBar1.value = 0
    rs.Open "Select * From Company", cnn, adOpenKeyset, adLockOptimistic
        If Not (rs.EOF Or rs.BOF) Then
            CompanyName = rs("CompanyName")
            SubName = rs("SubName")
            
        End If
    rs.Close
End Sub

Private Sub Text1_Change(Index As Integer)
    Text1(Index) = LTrim(Text1(Index))
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

Private Sub Timer1_Timer()
    ProgressBar1.Visible = True
    Text3.Visible = True
    ProgressBar1.value = ProgressBar1.value + 10
    Text3 = ProgressBar1.value & "%"
    
    If ProgressBar1.value >= 100 Then
    Timer1.Enabled = False
    Unload Me
    MDIForm1.Show
    End If
End Sub
