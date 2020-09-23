VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmemployeepaydetails 
   Caption         =   "Employee Pay Details"
   ClientHeight    =   7950
   ClientLeft      =   825
   ClientTop       =   210
   ClientWidth     =   10785
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form5"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   10785
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
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
      Left            =   1680
      TabIndex        =   53
      Top             =   6960
      Width           =   735
   End
   Begin VB.TextBox txtdesignation 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   52
      Top             =   2400
      Width           =   2535
   End
   Begin VB.TextBox txtMonYear 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9240
      Locked          =   -1  'True
      TabIndex        =   51
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&New"
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
      Left            =   240
      TabIndex        =   49
      Top             =   6960
      Width           =   735
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   48
      Top             =   7575
      Width           =   10785
      _ExtentX        =   19024
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13388
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "12:26 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "12/5/2000"
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cboEmpno 
      Height          =   360
      Left            =   4680
      Style           =   2  'Dropdown List
      TabIndex        =   47
      Top             =   1320
      Width           =   2055
   End
   Begin VB.ComboBox cboBranch 
      Height          =   360
      Left            =   4680
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   45
      Top             =   840
      Width           =   2175
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
      Height          =   360
      Index           =   0
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   43
      Top             =   1320
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Edit"
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
      Left            =   960
      TabIndex        =   42
      Top             =   6960
      Width           =   735
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   41
      Top             =   6960
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
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
      Index           =   5
      Left            =   7080
      TabIndex        =   38
      Top             =   6360
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   10
      Left            =   1920
      TabIndex        =   36
      Top             =   6360
      Width           =   1575
   End
   Begin VB.TextBox Text2 
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
      Index           =   3
      Left            =   7080
      MaxLength       =   10
      TabIndex        =   13
      Top             =   5160
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
      Height          =   375
      Index           =   11
      Left            =   1920
      MaxLength       =   10
      TabIndex        =   6
      Top             =   4320
      Width           =   1935
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
      Index           =   3
      Left            =   3840
      TabIndex        =   32
      Top             =   6960
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
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
      Index           =   2
      Left            =   3120
      TabIndex        =   31
      Top             =   6960
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Save"
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
      Left            =   2400
      TabIndex        =   30
      Top             =   6960
      Width           =   735
   End
   Begin VB.TextBox Text2 
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
      Index           =   4
      Left            =   7080
      MaxLength       =   10
      TabIndex        =   14
      Top             =   5640
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0C0FF&
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
      Index           =   2
      Left            =   7080
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "0"
      Top             =   4680
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
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
      Left            =   7080
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4200
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
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
      Left            =   7080
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3720
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
      Height          =   375
      Index           =   9
      Left            =   1920
      MaxLength       =   4
      TabIndex        =   8
      Top             =   5280
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
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
      Index           =   7
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4800
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
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
      Index           =   6
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3720
      Width           =   2055
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
      Index           =   5
      Left            =   1920
      TabIndex        =   4
      Text            =   "0"
      Top             =   3720
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   8
      Left            =   1920
      MaxLength       =   4
      TabIndex        =   9
      Top             =   5760
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
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
      Index           =   4
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2880
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
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
      Index           =   3
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2400
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
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
      Index           =   2
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1965
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
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
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1965
      Width           =   2535
   End
   Begin VB.Label lblCompanyName 
      AutoSize        =   -1  'True
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   5640
      TabIndex        =   55
      Top             =   120
      Width           =   75
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      Caption         =   "CompanyName"
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   3360
      TabIndex        =   54
      Top             =   120
      Width           =   1605
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Press F2 Key"
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
      Left            =   9600
      TabIndex        =   50
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Shape Shape2 
      Height          =   3255
      Left            =   5520
      Top             =   3600
      Width           =   3855
   End
   Begin VB.Shape Shape1 
      Height          =   3255
      Left            =   240
      Top             =   3600
      Width           =   4215
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "Branch"
      Height          =   240
      Left            =   3120
      TabIndex        =   46
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "Month"
      Height          =   240
      Left            =   8280
      TabIndex        =   44
      Top             =   840
      Width           =   630
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "NetTotal"
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
      Left            =   5760
      TabIndex        =   40
      Top             =   6960
      Width           =   900
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "Total"
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
      Left            =   5760
      TabIndex        =   39
      Top             =   6360
      Width           =   420
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "Total"
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
      Left            =   360
      TabIndex        =   37
      Top             =   6360
      Width           =   420
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "Advance"
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
      Left            =   5760
      TabIndex        =   35
      Top             =   5160
      Width           =   675
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "DEDUCTION :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5760
      TabIndex        =   34
      Top             =   3360
      Width           =   1200
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "EARNINGS :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   360
      TabIndex        =   33
      Top             =   3360
      Width           =   1065
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Telephone"
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
      Left            =   360
      TabIndex        =   29
      Top             =   4320
      Width           =   840
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "Others"
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
      Left            =   5760
      TabIndex        =   28
      Top             =   5640
      Width           =   585
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Loan due"
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
      Left            =   5760
      TabIndex        =   27
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "ESIC"
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
      Left            =   5760
      TabIndex        =   26
      Top             =   4200
      Width           =   420
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "EPF"
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
      Left            =   5760
      TabIndex        =   25
      Top             =   3720
      Width           =   330
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Worked Days"
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
      Left            =   360
      TabIndex        =   24
      Top             =   5760
      Width           =   1125
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Total Days"
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
      Left            =   360
      TabIndex        =   23
      Top             =   5280
      Width           =   870
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Others"
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
      Left            =   360
      TabIndex        =   22
      Top             =   4800
      Width           =   585
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "DA"
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
      Left            =   360
      TabIndex        =   21
      Top             =   3720
      Width           =   270
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "EmpNo."
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
      Left            =   3120
      TabIndex        =   20
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "EName"
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
      Left            =   480
      TabIndex        =   19
      Top             =   2025
      Width           =   570
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Designation"
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
      Left            =   480
      TabIndex        =   18
      Top             =   2475
      Width           =   990
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "DOB"
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
      Left            =   5640
      TabIndex        =   17
      Top             =   1965
      Width           =   420
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "DOJ"
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
      Left            =   5640
      TabIndex        =   16
      Top             =   2415
      Width           =   390
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "BASIC"
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
      Left            =   480
      TabIndex        =   15
      Top             =   2880
      Width           =   585
   End
End
Attribute VB_Name = "frmemployeepaydetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public AddEditViewMode As String
Dim sal As Double
Public Mon As String
Public year As Integer
Public frmshow As Boolean
Dim Msg As String
Dim TextDateVal As Double
Dim EarningSal As Double
Dim EarningDA As Double
Dim TotBasicDA As Double
Private Sub cbobranch_Click()
rs.Open "Select empno from admin where CompanyName='" & lblCompanyName.Caption & "' and branchcode='" & cboBranch & "' order by empno", cnn, adOpenKeyset, adLockOptimistic
cboEmpno.clear
While Not rs.EOF
    cboEmpno.AddItem SubName & "/" & Format(rs.Fields(0), "000")
    rs.MoveNext
Wend
rs.Close

For i = Text1.LBound To Text1.UBound
    Text1(i) = ""
Next
For i = Text2.LBound To Text2.UBound
    Text2(i) = ""
Next

Text3 = ""
txtDesignation = ""
Command1(0).Enabled = False
Command1(1).Enabled = False
cmdDelete.Enabled = False
Command2.Enabled = False
End Sub

Private Sub cboEmpno_Click()
AddEditViewMode = "View"
Text1(0) = cboEmpno
If cboEmpno <> "" Then
    DisplayRecords
End If
Command1(2).Enabled = True
End Sub
Private Sub CmdDelete_Click()
AddEditViewMode = "View"
Msg = MsgBox("Are you sure to delete this Record", vbExclamation + vbYesNo, "Payroll")
If Msg = 6 Then

cnn.Execute "Delete From Daughters Where CompanyName='" & lblCompanyName.Caption & "' and Empno= " & Mid(Text1(0), InStr(1, Text1(0), "/") + 1, Len(Text1(0)) - InStr(1, Text1(0), "/")) & " and datemon=#" & 1 & "-" & Mon & " - " & year & "#"

rs.Open "Select * from LoanDetails where CompanyName='" & lblCompanyName.Caption & "' and Empno= " & Mid(Text1(0), InStr(1, Text1(0), "/") + 1, Len(Text1(0)) - InStr(1, Text1(0), "/")) & " and monyear=#" & 1 & "-" & Mon & " - " & year & "#", cnn, adOpenKeyset, adLockOptimistic
While Not rs.EOF

rs1.Open "Select Balance from LoanMaster Where CompanyName='" & lblCompanyName.Caption & "' and Empno= " & Mid(Text1(0), InStr(1, Text1(0), "/") + 1, Len(Text1(0)) - InStr(1, Text1(0), "/")) & " and Loanid=" & rs.Fields("LoanId"), cnn, adOpenKeyset, adLockOptimistic

cnn.Execute "Update LoanMaster set Balance=" & CDbl(rs1.Fields("Balance")) & " +" & CDbl(rs.Fields("paidAmount")) & " Where CompanyName='" & lblCompanyName.Caption & "' and Empno= " & Mid(Text1(0), InStr(1, Text1(0), "/") + 1, Len(Text1(0)) - InStr(1, Text1(0), "/")) & " and Loanid=" & rs.Fields("LoanId")
rs1.Close
rs.MoveNext
Wend
rs.Close


cnn.Execute "Delete From LoanDetails Where CompanyName='" & lblCompanyName.Caption & "' and Empno= " & Mid(Text1(0), InStr(1, Text1(0), "/") + 1, Len(Text1(0)) - InStr(1, Text1(0), "/")) & " and monyear=#" & 1 & "-" & Mon & " - " & year & "#"
clear
txtMonYear = Mon & "-" & year
cboBranch.clear
cboEmpno.clear
AddBranch
Command2.Enabled = False
cmdDelete.Enabled = False
Command1(0).Enabled = False
End If
Cancel = True
cboBranch.SetFocus
StatusBar1.Panels(1) = AddEditViewMode
End Sub
Private Sub Command1_Click(Index As Integer)

Select Case Index
    
    Case 0                               'Edit Button
    
        AddEditViewMode = "Edit"
        TextLock (False)
        cboEmpno.Enabled = False
        cboBranch.Enabled = False
        Command1(0).Enabled = False
        Command1(1).Enabled = True
        Command2.Enabled = False
        cmdDelete.Enabled = False
        Text1(6).SetFocus
        StatusBar1.Panels(1) = AddEditViewMode
        
    Case 1                               'Save Button
    
        If AddEditViewMode = "Add" Then
            rs.Open "Select * from daughters ", cnn, adOpenKeyset, adLockPessimistic
            rs.AddNew
        Else
            rs.Open "Select * from daughters  where CompanyName='" & lblCompanyName.Caption & "' and empno= " & Mid(Text1(0), InStr(1, Text1(0), "/") + 1, Len(Text1(0)) - InStr(1, Text1(0), "/")) & " and datemon=#" & 1 & "-" & Mon & " - " & year & "#", cnn, adOpenKeyset, adLockOptimistic
        End If
            
        
        
            
        rs.Fields(0) = Mid(Text1(0), InStr(1, Text1(0), "/") + 1)
        rs.Fields(1) = Text1(1)
        rs.Fields(2) = txtDesignation
        rs.Fields(3) = Text1(2)
        rs.Fields(4) = Text1(3)
        rs.Fields(5) = Text1(4)
        rs.Fields(6) = Val(Text1(5))
        rs.Fields(7) = Text1(6)
        rs.Fields(8) = Text1(7)
        rs.Fields(9) = Text1(8)
        rs.Fields(10) = Text1(9)
        rs.Fields(11) = Text2(0)
        rs.Fields(12) = Text2(1)
        rs.Fields(13) = Val(Text2(2))
        rs.Fields(14) = Val(Text2(3))
        rs.Fields(15) = Val(Text2(4))
        rs.Fields(16) = Val(Text1(11))
        rs.Fields(17) = 1 & "-" & Mon & "-" & year
        rs.Fields(18) = CompanyName
        rs.Update
        
        rs.Close
        SaveLoan
        SaveLoan1
        SaveLoan2
        cboEmpno.Enabled = True
        cboBranch.Enabled = True
        Command2.Enabled = False
        Command1(1).Enabled = False
        Command1(0).Enabled = True
        cmdDelete.Enabled = True
        Unload frmloanDetails
        AddEditViewMode = "View"
        StatusBar1.Panels(1) = AddEditViewMode
        
        
    Case 2      'Clear button
        AddEditViewMode = "View"
        clear
        txtMonYear = Mon & "-" & year
        cboBranch.clear
        cboEmpno.clear
        AddBranch
        Command1(0).Enabled = False
        Command1(1).Enabled = False
        cboBranch.Enabled = True
        cboEmpno.Enabled = True
        Unload frmloanDetails
        Cancel = True
        cboBranch.SetFocus
        StatusBar1.Panels(1) = AddEditViewMode
   
    Case 3
       '***  cnn.Close
        Unload frmloanDetails
        Unload Me
End Select

End Sub
Private Sub Command2_Click()  'New Button
AddEditViewMode = "Add"
Command2.Enabled = False
Command1(2).Enabled = True
cboEmpno.Enabled = False
cboBranch.Enabled = False
TextLock (False)
Text1(11).SetFocus
StatusBar1.Panels(1) = AddEditViewMode
End Sub
Private Sub Form_Activate()
'cboBranch.SetFocus

End Sub
Private Sub Form_Load()
frmemployeepaydetails.Icon = LoadPicture(App.Path & "\Jamaica my computer.ico")
If frmCalendar.Combo1 <> "" Then
    Mon = frmCalendar.Combo1
End If
If frmCalendar.Combo2 <> "" Then
    year = frmCalendar.Combo2
End If
txtMonYear = Mon & "-" & year

'rs.Open "Select * from DesignationMaster", cnn, adOpenKeyset, adLockPessimistic
'    While Not rs.EOF
'        Combo1.AddItem rs.Fields("Designation")
'        rs.MoveNext
'    Wend
'
'rs.Close
lblCompanyName.Caption = CompanyName
AddBranch

Command2.Enabled = False
Command1(0).Enabled = False
Command1(1).Enabled = False
cmdDelete.Enabled = False
AddEditViewMode = "View"
StatusBar1.Panels(1) = AddEditViewMode
End Sub
Private Sub Text1_Change(Index As Integer)
If AddEditViewMode <> "View" Then
Select Case Index
    Case 5, 6, 7, 8, 9, 11

        If Val(Text1(9)) <> 0 Then
            EarningSal = Round(Val(Text1(4) * Val(Text1(8))) / Text1(9), 2)
            EarningDA = Round(Val(Text1(6) * Val(Text1(8))) / Text1(9), 2)
            Text1(10) = AccurateCost1(EarningSal + EarningDA + Val(Text1(7)) + Val(Text1(11)))
            
            rs1.Open "Select Permanent from Admin where CompanyName='" & lblCompanyName.Caption & "' and empno=" & Mid(cboEmpno, InStr(1, cboEmpno, "/") + 1), cnn, adOpenKeyset, adLockPessimistic
            If rs1.Fields("Permanent") = True Then
                rs.Open "Select pf from rangePF Where fromamount<=" & TotBasicDA & "  and toamount>=" & TotBasicDA, cnn, adOpenKeyset, adLockOptimistic
                If Not rs.EOF Then
                    
                    Text2(0) = Format(Val(Round((Val(Text1(4)) + Val(Text1(6))) * rs.Fields("PF") * Val(Text1(8)) / Val(Text1(9)))), "0.00")
                Else
                    StatusBar1.Panels(1) = "This Amout isn't Come to Range"
                End If
                rs.Close
                
                rs.Open "Select ESI from rangeESI Where fromamount<=" & TotBasicDA & "  and toamount>=" & TotBasicDA, cnn, adOpenKeyset, adLockOptimistic
                If Not rs.EOF Then
                    
                    Text2(1) = Val(Format(Round((Val(Text1(4)) + Val(Text1(6))) * rs.Fields("ESI") * Val(Text1(8)) / Val(Text1(9)), 2), "0.00"))
                    Text2(1) = AccurateCost(Text2(1))
                Else
                    StatusBar1.Panels(1) = "This Amout isn't Come to Range"
                End If
                rs.Close
            Else
                Text2(0) = "0.00"
                Text2(1) = "0.00"
            End If
            rs1.Close
            'Text2(0) = Val(Round((Val(Text1(4)) + Val(Text1(6))) * 0.12 * Val(Text1(8)) / Val(Text1(9))))
            'Text2(1) = Val(Format(Round((Val(Text1(4)) + Val(Text1(6))) * 0.0175 * Val(Text1(8)) / Val(Text1(9)), 2), ".00"))
            
        End If
    Case 10
        Text3 = Val(Text1(10)) - Val(Text2(5))
        'Text3 = Val(Format(Round(Val(Text1(10)) - Val(Text2(5)), 2), ".00"))
        
        Text3 = AccurateCost1(Text3)
End Select

StatusBar1.Panels(1) = AddEditViewMode
If Len(Trim(Text1(8))) = 0 Or Len(Trim(Text1(9))) = 0 Or Len(Trim(Text2(0))) = 0 Or Len(Trim(Text2(1))) = 0 Then
    Command1(1).Enabled = False
    Exit Sub
End If
If Val(Text1(8)) > Val(Text1(9)) Then
   Command1(1).Enabled = False
   StatusBar1.Panels(1) = "Total Days Should be greater then Worked Days"
   Exit Sub
End If
Command1(1).Enabled = True
End If
End Sub
Private Sub Text1_GotFocus(Index As Integer)
Select Case Index
    Case 8, 9
        If Text1(Index) <> "" Then
            TextDateVal = CDbl(Text1(Index))
        End If
        
    
End Select


End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If AddEditViewMode <> "View" Then
Select Case Index
    
    Case 5, 6, 7, 8, 9, 11
    
        KeyAscii = NumericCheck(Text1(Index), CInt(KeyAscii))
       

End Select
End If
End Sub
Private Sub Text1_LostFocus(Index As Integer)
If AddEditViewMode <> "View" Then
Select Case Index

 Case 8, 9
        If Text1(Index) <> "" Then
         If Mon = "Dec" Then
            If Text1(Index) > 31 Then
                MsgBox "Total Days Exceeds the Number of Days in the Current Month ", vbExclamation, "Payroll"
                If AddEditViewMode = "Add" Then
                    Text1(Index) = ""
                Else
                    Text1(Index) = TextDateVal
            End If
                    Cancel = True
                    Text1(Index).SetFocus
                    Exit Sub
            End If
         Else
         If Format(CDate(Format(Mon & "/" & 1 & "/" & year, "m") + 1 & "/" & 1 & "/" & year) - 1, "dd") < Val(Text1(Index)) Or Val(Text1(9)) = 0 Then
              MsgBox "Total Days Exceeds the Number of Days in the Current Month ", vbExclamation, "Payroll"
              If AddEditViewMode = "Add" Then
                Text1(Index) = ""
            Else
                Text1(Index) = TextDateVal
            End If
            Cancel = True
            Text1(Index).SetFocus
            Exit Sub
              
         End If
         End If
         End If
        'If Val(Text1(Index)) > 31 Or Val(Text1(Index)) = 0 Then
        '    MsgBox "Date between 1 to 31"
         '   If AddEditViewMode = "Add" Then
         '       Text1(Index) = ""
         '
         '   Else
          '      Text1(Index) = TextDateVal
          '  End If
           ' Cancel = True
           ' Text1(Index).SetFocus
            'Exit Sub
       ' End If
        
        If Val(Text1(9)) < Val(Text1(8)) Then
            MsgBox "Worked Day Should be Less Than Total Days", vbExclamation, "Payroll"
            If AddEditViewMode = "Add" Then
                Text1(Index) = ""
            Else
                Text1(Index) = TextDateVal
            End If
            Cancel = True
            Text1(Index).SetFocus
            Exit Sub
        End If
        
        
End Select
End If
End Sub

Private Sub Text2_Change(Index As Integer)
If AddEditViewMode <> "View" Then
Text2(5) = Val(Text2(0)) + Val(Text2(1)) + Val(Text2(2)) + Val(Text2(3)) + Val(Text2(4))
Text3 = Val(Text1(10)) - Val(Text2(5))

'Text3 = AccurateCost1(Text3)
StatusBar1.Panels(1) = AddEditViewMode
If Len(Trim(Text1(8))) = 0 Or Len(Trim(Text1(9))) = 0 Or Len(Trim(Text2(0))) = 0 Or Len(Trim(Text2(1))) = 0 Then
    Command1(1).Enabled = False
    Exit Sub
End If
End If
End Sub
'Public Sub findrecpos()
'rs.Open "select * from daughters where datemon=#" & 1 & "-" & mon & " - " & year & "#", cnn
'If rs.EOF = False Then
'    Text1(0) = "STP/" & Format(rs.Fields(0), "00")
'    Text1(1) = rs.Fields(1)
'    txtdesignation = rs.Fields(2)
'    Text1(2) = rs.Fields(3)
'    Text1(3) = rs.Fields(4)
'    Text1(4) = rs.Fields(5)
'    Text1(5) = rs.Fields(6)
'    Text1(6) = rs.Fields(7)
'    Text1(7) = rs.Fields(8)
'    Text1(8) = rs.Fields(9)
'    Text1(9) = rs.Fields(10)
'    Text2(0) = rs.Fields(11)
'    Text2(1) = rs.Fields(12)
'    Text2(2) = rs.Fields(13)
'    Text2(3) = rs.Fields(14)
'    Text2(4) = rs.Fields(15)
'    Text1(11) = rs.Fields(16)
'    Command1(0).Enabled = True
'Else
'
'For i = Text1.LBound To Text1.UBound
'    Text1(i) = ""
'Next
'For i = Text2.LBound To Text2.UBound
'    Text2(i) = ""
'Next
'Text2(2) = 0
'Text3 = ""
'
'rs1.Open "select * from admin", cnn, adOpenKeyset, adLockOptimistic
'Text1(0) = "STP/" & Format(rs1.Fields(0), "000")
'Text1(1) = rs1.Fields(1)
'txtdesignation = rs1.Fields(2)
'Text1(2) = rs1.Fields(3)
'Text1(3) = rs1.Fields(4)
'Text1(4) = rs1.Fields(5)
'Command1(0).Enabled = False
'rs1.Close
'
'
'End If
'rs.Close
'
'End Sub
Private Sub Text2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If cboEmpno <> "" Then
    If KeyCode = vbKeyF2 Then
        Select Case Index
            Case 2
                
                    frmloanDetails.Show
        End Select
    End If
    End If
End Sub
Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii > 47 And KeyAscii < 59 Or KeyAscii = 8 Or KeyAscii = 46 Then
Else
    KeyAscii = 0
End If
End Sub
Public Function AccurateCost(Text As Double)
' If Right(text, 1) > 5 Then
'
'    AccurateCost = Format((Round(text, 1)), ".00")
'ElseIf Right(text, 1) < 5 Then
'    AccurateCost = Format((Round(text, 2)), ".05")
'ElseIf Right(text, 1) = 5 Then
'    AccurateCost = text
'
'End If
 If Right(Format(Text, ".00"), 1) > 5 Then
    AccurateCost = Format((Round(Text, 1)), "0.00")
ElseIf Right(Format(Text, ".00"), 1) < 5 And Right(Format(Text, ".00"), 1) <> 0 Then
    AccurateCost = Format((Round(Text, 2)), "0.05")
ElseIf Right(Format(Text, ".00"), 1) = 5 Or Right(Format(Text, ".00"), 1) = 0 Then
    AccurateCost = Format(Text, "0.00")
End If
End Function
Public Sub SaveLoan()
If AddEditViewMode <> "View" Then
If Trim(frmloanDetails.Text1(0)) <> "" And Trim(frmloanDetails.Text1(0)) <> "" Then
rs.Open "Select * from LoanDetails", cnn, adOpenKeyset, adLockOptimistic
rs.AddNew
rs.Fields("Empno") = Val(Mid(Text1(0), InStr(1, Text1(0), "/") + 1, Len(Text1(0)) - InStr(1, Text1(0), "/")))
rs.Fields("loanid") = frmloanDetails.lid1
rs.Fields("monyear") = 1 & "-" & Mon & "-" & year
If frmloanDetails.Check1(0).value = 0 Then
    rs.Fields("paid") = False
    rs.Fields("paidamount") = 0
Else
    rs.Fields("paid") = True
    rs.Fields("paidamount") = Val(frmloanDetails.lblinstall)
End If
rs.Fields("CompanyName") = CompanyName
rs.Update
rs.Close

rs.Open "Select sum(paidamount) from LoanDetails Where CompanyName='" & lblCompanyName.Caption & "' and EmpNo=" & Right(Text1(0), Len(Text1(0)) - InStr(1, Text1(0), "/")) & " And loanid = " & frmloanDetails.lid1, cnn, adOpenKeyset, adLockOptimistic


rs1.Open "Select * from LoanMaster Where CompanyName='" & lblCompanyName.Caption & "' and EmpNo=" & Right(Text1(0), Len(Text1(0)) - InStr(1, Text1(0), "/")) & " And loanid = " & frmloanDetails.lid1, cnn, adOpenKeyset, adLockOptimistic
If rs1.EOF = False Then
    masterbalance = rs1.Fields("LoanAmt") - rs.Fields(0)
End If
rs1.Close
rs.Close

rs.Open "Select * from LoanMaster Where CompanyName='" & lblCompanyName.Caption & "' and EmpNo=" & Right(Text1(0), Len(Text1(0)) - InStr(1, Text1(0), "/")) & " And loanid = " & frmloanDetails.lid1, cnn, adOpenKeyset, adLockOptimistic
rs.Fields("Balance") = masterbalance
rs.Update
rs.Close
End If
End If
End Sub
Public Sub SaveLoan1()
If AddEditViewMode <> "View" Then
If Trim(frmloanDetails.Text2(0)) <> "" And Trim(frmloanDetails.Text2(0)) <> "" Then
rs.Open "Select * from LoanDetails", cnn, adOpenKeyset
rs.AddNew
rs.Fields("Empno") = Val(Mid(Text1(0), InStr(1, Text1(0), "/") + 1, Len(Text1(0)) - InStr(1, Text1(0), "/")))
rs.Fields("loanid") = frmloanDetails.lid2
rs.Fields("monyear") = 1 & "-" & Mon & "-" & year
If frmloanDetails.Check1(1).value = 0 Then
    rs.Fields("paid") = False
    rs.Fields("paidamount") = 0
Else
    rs.Fields("paid") = True
    rs.Fields("paidamount") = Val(frmloanDetails.lblinstall2)
End If
rs.Fields("CompanyName") = CompanyName
rs.Update
rs.Close
rs.Open "Select sum(paidamount) from LoanDetails Where CompanyName='" & lblCompanyName.Caption & "' and EmpNo=" & Right(Text1(0), Len(Text1(0)) - InStr(1, Text1(0), "/")) & " And loanid = " & frmloanDetails.lid2, cnn, adOpenKeyset, adLockOptimistic
rs1.Open "Select * from LoanMaster Where CompanyName='" & lblCompanyName.Caption & "' and EmpNo=" & Right(Text1(0), Len(Text1(0)) - InStr(1, Text1(0), "/")) & " And loanid = " & frmloanDetails.lid2, cnn, adOpenKeyset, adLockOptimistic
If rs1.EOF = False Then
    masterbalance = rs1.Fields("LoanAmt") - rs.Fields(0)
End If
rs1.Close
rs.Close
rs.Open "Select * from LoanMaster Where CompanyName='" & lblCompanyName.Caption & "' and EmpNo=" & Right(Text1(0), Len(Text1(0)) - InStr(1, Text1(0), "/")) & " And loanid = " & frmloanDetails.lid2, cnn, adOpenKeyset, adLockOptimistic
rs.Fields("Balance") = masterbalance
rs.Update
rs.Close
End If
End If
End Sub
Public Sub DisplayRecords()
rs.Open "select * from daughters where CompanyName='" & lblCompanyName.Caption & "' and empno=" & Mid(cboEmpno, InStr(1, cboEmpno, "/") + 1) & " and datemon=#" & 1 & "-" & Mon & " - " & year & "#", cnn, adOpenKeyset, adLockOptimistic
If rs.EOF = False Then
    Text1(0) = SubName & "/" & Format(rs.Fields(0), "000")
    Text1(1) = rs.Fields(1)
    txtDesignation = rs.Fields(2)
    Text1(2) = Format(rs.Fields(3), "mm/dd/yyyy")
    Text1(3) = Format(rs.Fields(4), "mm/dd/yyyy")
    Text1(4) = rs.Fields(5)
    Text1(5) = rs.Fields(6)
    Text1(6) = rs.Fields(7)
    Text1(7) = rs.Fields(8)
    Text1(8) = rs.Fields(9)
    Text1(9) = rs.Fields(10)
    Text2(0) = rs.Fields(11)
    Text2(1) = rs.Fields(12)
    Text2(2) = rs.Fields(13)
    Text2(3) = rs.Fields(14)
    Text2(4) = rs.Fields(15)
    Text1(11) = rs.Fields(16)
    
    EarningSal = Round(Val(Text1(4) * Val(Text1(8))) / Text1(9), 2)
    EarningDA = Round(Val(Text1(6) * Val(Text1(8))) / Text1(9), 2)
    
    
    Text1(10) = EarningSal + EarningDA + Val(Text1(7)) + Val(Text1(11))

        'If Val(Text1(9)) <> 0 Then
         '   Text1(10) = Val(Text1(10)) / Val(Text1(9))
          '  Text1(10) = Format((Round((Val(Text1(10)) * Val(Text1(8))), 2)), ".00")
         'End If
        Text2(5) = Val(Text2(0)) + Val(Text2(1)) + Val(Text2(2)) + Val(Text2(3)) + Val(Text2(4))
        Text3 = Val(Text1(10)) - Val(Text2(5))

        'Text3 = AccurateCost1(Text3)
    TextLock (True)
    Command1(0).Enabled = True
    Command2.Enabled = False
    Command1(2).Enabled = True
    cmdDelete.Enabled = True
    
Else

For i = Text1.LBound To Text1.UBound
    Text1(i) = ""
Next
For i = Text2.LBound To Text2.UBound
    Text2(i) = ""
Next
Text3 = ""
Text2(2) = 0
TextLock (True)
Text1(6).SetFocus
rs1.Open "select * from admin where CompanyName='" & lblCompanyName.Caption & "' and empno=" & Mid(cboEmpno, InStr(1, cboEmpno, "/") + 1), cnn, adOpenKeyset, adLockPessimistic
Text1(0) = SubName & "/" & Format(rs1.Fields(0), "000")
Text1(1) = rs1.Fields(1)
txtDesignation = rs1.Fields(2)
Text1(2) = Format(rs1.Fields(3), "mm/dd/yyyy")
Text1(3) = Format(rs1.Fields(4), "mm/dd/yyyy")
Text1(4) = rs1.Fields(5)
'text1(6)
rs1.Close
rs1.Open "Select Permanent from Admin where CompanyName='" & lblCompanyName.Caption & "' and empno=" & Mid(cboEmpno, InStr(1, cboEmpno, "/") + 1), cnn, adOpenKeyset, adLockPessimistic
If rs1.Fields("Permanent") = True Then
    rs2.Open "Select Da From DesignationMaster Where Designation='" & txtDesignation & "'", cnn, adOpenKeyset, adLockOptimistic
    Text1(6) = rs2.Fields(0)
    rs2.Close
Else
    Text1(6) = 0
End If
rs1.Close

rs1.Open "Select Others From Admin  Where CompanyName='" & lblCompanyName.Caption & "' and empno=" & Mid(cboEmpno, InStr(1, cboEmpno, "/") + 1), cnn, adOpenKeyset, adLockPessimistic
Text1(7) = rs1.Fields(0)
rs1.Close
TotBasicDA = Val(Text1(4)) + Val(Text1(6)) 'Basic+Da
Command1(0).Enabled = False
Command2.Enabled = True
cmdDelete.Enabled = flase
End If
rs.Close
End Sub
Public Sub TextLock(Locked As Boolean)
For i = Text1.LBound + 7 To Text1.UBound
    Text1(i).Locked = Locked
Next
For i = Text2.LBound + 3 To Text2.UBound
    Text2(i).Locked = Locked
Next
End Sub
Public Function CommandEnableDisable(Control As Object, StartIndex, EndIndex) As Boolean
CommandEnableDisable = True
For i = StartIndex To EndIndex
    If Len(Trim(Control(i))) = 0 Then
        If i <> 5 Then
            CommandEnableDisable = False
            Exit Function
        End If
        
    End If
Next



End Function
Public Sub SaveLoan2()
If AddEditViewMode <> "View" Then
If Trim(frmloanDetails.Text3(0)) <> "" And Trim(frmloanDetails.Text3(0)) <> "" Then
rs.Open "Select * from LoanDetails", cnn, adOpenKeyset, adLockOptimistic
rs.AddNew
rs.Fields("Empno") = Val(Mid(Text1(0), InStr(1, Text1(0), "/") + 1, Len(Text1(0)) - InStr(1, Text1(0), "/")))
rs.Fields("loanid") = frmloanDetails.lid3
rs.Fields("monyear") = Format(1 & "-" & Mon & "-" & year, "mm/dd/yyyy")
If frmloanDetails.Check1(2).value = 0 Then
    rs.Fields("paid") = False
    rs.Fields("paidamount") = 0
Else
    rs.Fields("paid") = True
    rs.Fields("paidamount") = Val(frmloanDetails.lblinstall3)
End If
rs.Fields("CompanyName") = CompanyName
rs.Update
rs.Close
rs.Open "Select sum(paidamount) from LoanDetails Where CompanyName='" & lblCompanyName.Caption & "' and EmpNo=" & Right(Text1(0), Len(Text1(0)) - InStr(1, Text1(0), "/")) & " And loanid = " & frmloanDetails.lid3, cnn, adOpenKeyset, adLockOptimistic
rs1.Open "Select * from LoanMaster Where CompanyName='" & lblCompanyName.Caption & "' and EmpNo=" & Right(Text1(0), Len(Text1(0)) - InStr(1, Text1(0), "/")) & " And loanid = " & frmloanDetails.lid3, cnn, adOpenKeyset, adLockOptimistic
If rs1.EOF = False Then
    masterbalance = rs1.Fields("LoanAmt") - rs.Fields(0)
End If
rs1.Close
rs.Close
rs.Open "Select * from LoanMaster Where CompanyName='" & lblCompanyName.Caption & "' and EmpNo=" & Right(Text1(0), Len(Text1(0)) - InStr(1, Text1(0), "/")) & " And loanid = " & frmloanDetails.lid3, cnn, adOpenKeyset, adLockOptimistic
rs.Fields("Balance") = masterbalance
rs.Update
rs.Close
End If
End If
End Sub
Public Sub AddBranch()
rs.Open "Select branchcode From Branch where CompanyName='" & lblCompanyName.Caption & "'", cnn, adOpenKeyset, adLockOptimistic
cboBranch.clear
While Not rs.EOF
    cboBranch.AddItem rs(0)
    rs.MoveNext
Wend
rs.Close
End Sub
Public Function AccurateCost1(Text As Double)
If Right(Format(Text, ".00"), 1) = 5 Then
    AccurateCost1 = Format(Text + 0.05, "0.00")
Else
    AccurateCost1 = Format((Round(Text, 1)), "0.00")
End If
End Function

