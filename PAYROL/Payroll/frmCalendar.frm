VERSION 5.00
Begin VB.Form frmCalendar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calendar"
   ClientHeight    =   1950
   ClientLeft      =   4125
   ClientTop       =   3030
   ClientWidth     =   2640
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   2640
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   375
      Index           =   1
      Left            =   1200
      TabIndex        =   3
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   2
      Top             =   1440
      Width           =   615
   End
   Begin VB.ComboBox Combo2 
      Height          =   345
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   345
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Year"
      Height          =   225
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Month"
      Height          =   225
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   525
   End
End
Attribute VB_Name = "frmCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
Combo2_Click
End Sub

Private Sub Combo2_Click()
If Combo1.Text <> "" And Combo2.Text <> "" Then
    Command1(0).Enabled = True
Else
    Command1(0).Enabled = False
End If

End Sub
Private Sub Command1_Click(Index As Integer)
Select Case Index
    
    Case 0
        frmemployeepaydetails.Show
        
    Case 1
        Unload Me
        
End Select

End Sub
Private Sub Form_Activate()
frmCalendar.Left = 0
frmCalendar.Top = 0
Command1(0).Enabled = False
End Sub

Private Sub Form_Load()
Combo1.clear
For i = 1 To 12
    Combo1.AddItem Format(i & "-" & i & "-" & 2000, "MMM")
Next
Combo2.clear
For i = 0 To 9
    Combo2.AddItem 200 & i
Next
For i = 10 To 99
    Combo2.AddItem 20 & i
Next
frmCalendar.Icon = LoadPicture(App.Path & "\Jamaica my computer.ico")
End Sub
