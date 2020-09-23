VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Departments"
   ClientHeight    =   6195
   ClientLeft      =   960
   ClientTop       =   1380
   ClientWidth     =   10125
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   6195
   ScaleWidth      =   10125
   Begin VB.CommandButton Command1 
      Caption         =   "&Admin"
      Height          =   495
      Index           =   2
      Left            =   7080
      TabIndex        =   3
      Top             =   2760
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Future&plan"
      Height          =   495
      Index           =   1
      Left            =   4080
      TabIndex        =   2
      Top             =   2760
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Finance"
      Height          =   495
      Index           =   0
      Left            =   1080
      TabIndex        =   1
      Top             =   2760
      Width           =   2535
   End
   Begin VB.Line Line6 
      X1              =   2160
      X2              =   2160
      Y1              =   2640
      Y2              =   2760
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   5280
      X2              =   5280
      Y1              =   2280
      Y2              =   1440
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   5280
      X2              =   5280
      Y1              =   2280
      Y2              =   2760
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   8160
      X2              =   8160
      Y1              =   2280
      Y2              =   2760
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   2160
      X2              =   8160
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   2160
      X2              =   2160
      Y1              =   2760
      Y2              =   2280
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Departments"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4440
      TabIndex        =   0
      Top             =   1200
      Width           =   1575
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
Select Case Index
    
    Case 0
    
    Case 1
    
    Case 2
        Unload Me
        Form3.Show vbModal
    
End Select

End Sub

