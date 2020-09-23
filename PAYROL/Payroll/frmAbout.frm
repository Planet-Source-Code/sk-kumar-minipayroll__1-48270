VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Payroll 1.0"
   ClientHeight    =   3105
   ClientLeft      =   3240
   ClientTop       =   2385
   ClientWidth     =   5490
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   5490
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Ok"
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
      Left            =   4320
      TabIndex        =   2
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Kumar [sk_kumar2003@yahoo.com]"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1560
      TabIndex        =   3
      Top             =   1560
      Width           =   3735
   End
   Begin VB.Label Label2 
      Caption         =   "This Product is Licenced to :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   1695
      Left            =   120
      Picture         =   "frmAbout.frx":0000
      Stretch         =   -1  'True
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Payroll System 1.0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   1470
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   
    Unload Me
End Sub


Private Sub Form_Load()
frmAbout.Left = 0
frmAbout.Top = 0
frmAbout.Icon = LoadPicture(App.Path & "\Jamaica my computer.ico")
'Image1.Picture = LoadPicture(App.Path & "\Computer.wmf")
End Sub
