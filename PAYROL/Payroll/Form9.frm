VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmEmployeeDetails 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Details"
   ClientHeight    =   7515
   ClientLeft      =   690
   ClientTop       =   585
   ClientWidth     =   10335
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form9"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   10335
   Begin VB.CommandButton cmdPrint1 
      Caption         =   "Print&All"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5520
      TabIndex        =   69
      Top             =   6720
      Width           =   855
   End
   Begin VB.CommandButton CmdPrint 
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4680
      TabIndex        =   65
      Top             =   6720
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
      Height          =   345
      Index           =   3
      Left            =   2940
      TabIndex        =   22
      Top             =   6720
      Width           =   735
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5895
      Left            =   240
      TabIndex        =   25
      Top             =   600
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   10398
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "E&mployee Details"
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label15(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label15(1)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label15(2)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label15(3)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label15(4)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label15(5)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label15(9)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "label17"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label18"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label22"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label23"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label24"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label25"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label28"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "List1"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Combo1"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Text1(1)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Text1(3)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Text1(4)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Text1(0)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "cboBranch"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Text1(2)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Text1(14)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Text1(15)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Frame1"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).ControlCount=   32
      TabCaption(1)   =   "Pe&rsonal Details"
      TabPicture(1)   =   "Form9.frx":0000
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text1(13)"
      Tab(1).Control(1)=   "Text1(12)"
      Tab(1).Control(2)=   "Text1(11)"
      Tab(1).Control(3)=   "Text1(10)"
      Tab(1).Control(4)=   "Text1(9)"
      Tab(1).Control(5)=   "Text1(8)"
      Tab(1).Control(6)=   "Text1(7)"
      Tab(1).Control(7)=   "Text1(6)"
      Tab(1).Control(8)=   "Text1(5)"
      Tab(1).Control(9)=   "Label26"
      Tab(1).Control(10)=   "Label21"
      Tab(1).Control(11)=   "Label20"
      Tab(1).Control(12)=   "Label19"
      Tab(1).Control(13)=   "Label16"
      Tab(1).Control(14)=   "Label15(8)"
      Tab(1).Control(15)=   "Label15(7)"
      Tab(1).Control(16)=   "Label15(6)"
      Tab(1).Control(17)=   "Label14"
      Tab(1).Control(18)=   "Label13"
      Tab(1).Control(19)=   "Label12"
      Tab(1).Control(20)=   "Label11"
      Tab(1).Control(21)=   "Label10"
      Tab(1).Control(22)=   "Label9"
      Tab(1).Control(23)=   "Label8"
      Tab(1).ControlCount=   24
      Begin VB.Frame Frame1 
         Height          =   615
         Left            =   1800
         TabIndex        =   4
         Top             =   2640
         Width           =   3015
         Begin VB.OptionButton optTraniee 
            Caption         =   "Tr&aniee"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1800
            TabIndex        =   6
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton optPermanent 
            Caption         =   "&Permanent"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Value           =   -1  'True
            Width           =   1335
         End
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
         Height          =   345
         Index           =   15
         Left            =   1800
         MaxLength       =   12
         TabIndex        =   10
         Top             =   5280
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
         Height          =   345
         Index           =   14
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   59
         Top             =   4800
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
         Height          =   345
         Index           =   13
         Left            =   -71520
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   18
         Top             =   4680
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
         Height          =   345
         Index           =   2
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   7
         Top             =   3360
         Width           =   2055
      End
      Begin VB.ComboBox cboBranch 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   720
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
         Height          =   345
         Index           =   12
         Left            =   -71520
         Locked          =   -1  'True
         MaxLength       =   25
         TabIndex        =   11
         Top             =   600
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
         Height          =   345
         Index           =   11
         Left            =   -71520
         Locked          =   -1  'True
         MaxLength       =   35
         TabIndex        =   19
         Top             =   5280
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
         Height          =   345
         Index           =   10
         Left            =   -71520
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   17
         Top             =   4200
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
         Height          =   345
         Index           =   9
         Left            =   -71520
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   16
         Top             =   3600
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
         Height          =   345
         Index           =   8
         Left            =   -71520
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   15
         Top             =   3000
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
         Height          =   345
         Index           =   7
         Left            =   -71520
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   14
         Top             =   2400
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
         Height          =   345
         Index           =   6
         Left            =   -71520
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   13
         Top             =   1800
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
         Height          =   345
         Index           =   5
         Left            =   -71520
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   12
         Top             =   1200
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
         Height          =   345
         Index           =   0
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   1
         Top             =   1200
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
         Height          =   345
         Index           =   4
         Left            =   1800
         MaxLength       =   12
         TabIndex        =   9
         Top             =   4320
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
         Height          =   345
         Index           =   3
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   8
         Top             =   3855
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
         Height          =   345
         Index           =   1
         Left            =   1800
         MaxLength       =   40
         TabIndex        =   2
         Top             =   1665
         Width           =   3015
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   2235
         Width           =   3015
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4560
         Left            =   6360
         TabIndex        =   26
         Top             =   840
         Width           =   2895
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Status"
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
         Left            =   240
         TabIndex        =   68
         Top             =   2760
         Width           =   525
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   -72240
         TabIndex        =   64
         Top             =   4800
         Width           =   120
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1200
         TabIndex        =   63
         Top             =   5400
         Width           =   120
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1200
         TabIndex        =   62
         Top             =   4920
         Width           =   120
      End
      Begin VB.Label Label23 
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
         Left            =   240
         TabIndex        =   61
         Top             =   5400
         Width           =   585
      End
      Begin VB.Label Label22 
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
         Left            =   240
         TabIndex        =   60
         Top             =   4920
         Width           =   270
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "ESI No."
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
         Left            =   -73920
         TabIndex        =   58
         Top             =   4800
         Width           =   600
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "01-04228-30235"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   -69120
         TabIndex        =   57
         Top             =   5280
         Width           =   1170
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   -72240
         TabIndex        =   56
         Top             =   4200
         Width           =   120
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "MM/DD/YYYY"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   4320
         TabIndex        =   55
         Top             =   3600
         Width           =   1080
      End
      Begin VB.Label label17 
         AutoSize        =   -1  'True
         Caption         =   "Branch"
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
         Left            =   240
         TabIndex        =   54
         Top             =   765
         Width           =   615
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   9
         Left            =   1200
         TabIndex        =   53
         Top             =   720
         Width           =   120
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Father/Husband Name"
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
         Left            =   -73920
         TabIndex        =   52
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   8
         Left            =   -72120
         TabIndex        =   51
         Top             =   600
         Width           =   120
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   7
         Left            =   -72240
         TabIndex        =   50
         Top             =   3600
         Width           =   120
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   6
         Left            =   -72240
         TabIndex        =   49
         Top             =   2400
         Width           =   120
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   5
         Left            =   1200
         TabIndex        =   48
         Top             =   4200
         Width           =   120
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   4
         Left            =   1200
         TabIndex        =   47
         Top             =   3720
         Width           =   120
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   3
         Left            =   1200
         TabIndex        =   46
         Top             =   3240
         Width           =   120
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   2
         Left            =   1200
         TabIndex        =   45
         Top             =   2160
         Width           =   120
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   1
         Left            =   1200
         TabIndex        =   44
         Top             =   1560
         Width           =   120
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   43
         Top             =   1200
         Width           =   120
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Phone"
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
         Left            =   -73920
         TabIndex        =   42
         Top             =   5400
         Width           =   510
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "PFNo"
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
         Left            =   -73920
         TabIndex        =   41
         Top             =   4320
         Width           =   450
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "PinCode"
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
         Left            =   -73920
         TabIndex        =   40
         Top             =   3600
         Width           =   690
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "State"
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
         Left            =   -73920
         TabIndex        =   39
         Top             =   3000
         Width           =   420
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "City"
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
         Left            =   -73920
         TabIndex        =   38
         Top             =   2400
         Width           =   345
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Add2"
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
         Left            =   -73920
         TabIndex        =   37
         Top             =   1920
         Width           =   420
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Add1"
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
         Left            =   -73920
         TabIndex        =   36
         Top             =   1200
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "EmpNo."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   33
         Top             =   1200
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
         Left            =   240
         TabIndex        =   32
         Top             =   1665
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
         Left            =   240
         TabIndex        =   31
         Top             =   2235
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
         Left            =   240
         TabIndex        =   30
         Top             =   3285
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
         Left            =   240
         TabIndex        =   29
         Top             =   3855
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
         Left            =   240
         TabIndex        =   28
         Top             =   4320
         Width           =   585
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "LIST Of EMPLOYEES :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6600
         TabIndex        =   27
         Top             =   480
         Width           =   2415
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   35
      Top             =   7200
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12621
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "5:59 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "11/5/01"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command1 
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
      Height          =   345
      Index           =   0
      Left            =   360
      TabIndex        =   20
      Top             =   6720
      Width           =   735
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
      Height          =   345
      Index           =   1
      Left            =   1230
      TabIndex        =   21
      Top             =   6720
      Width           =   735
   End
   Begin VB.CommandButton Command1 
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
      Height          =   345
      Index           =   2
      Left            =   2100
      TabIndex        =   34
      Top             =   6720
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
      Height          =   345
      Index           =   4
      Left            =   3810
      TabIndex        =   23
      Top             =   6720
      Width           =   735
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
      Height          =   345
      Index           =   5
      Left            =   6480
      TabIndex        =   24
      Top             =   6720
      Width           =   735
   End
   Begin VB.Label Label27 
      Caption         =   "CompanyName"
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
      Left            =   2040
      TabIndex        =   67
      Top             =   120
      Width           =   1455
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
      Left            =   4920
      TabIndex        =   66
      Top             =   120
      Width           =   60
   End
End
Attribute VB_Name = "frmEmployeeDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***Dim cnn As New ADODB.Connection
'***Dim rs As New ADODB.Recordset
'***Dim rs1 As New ADODB.Recordset
Dim AddEditViewMode As String
Dim Dat As Date
Dim NextTab As Boolean
Dim EditPfNo As String
Dim EditESI As String
Dim EditEmpno
Dim Msg As String
Dim InBoxResDate As String
Dim InBoxBranch As String
Dim ResNo
Private Sub cbobranch_Click()
strBranch = cboBranch
If AddEditViewMode = "View" Then

'Combo1.ListIndex = 0
addcomboDesignation
CountRecord
'AddEditViewMode = "View"
End If
End Sub
Private Sub CmdPrint_Click()
Msg = MsgBox("Are you take Print", vbExclamation + vbYesNo, "Payroll")
If Msg = 6 Then
On Error GoTo err
'Printer.NewPage
Printer.Orientation = 1
Printer.Font.Size = 12
Printer.Font.Name = "Courier New"
Printer.FontBold = True
Printer.Print Space(20); CompanyName
Printer.Font.Size = 12
Printer.Print ""
Printer.Print Space(20); "Employee Personal Details"
Printer.Print Space(20); "-------------------------"
Printer.FontBold = False
Printer.Print ""
Printer.Print "Empno                 :" & Space(10); Text1(0)
Printer.Print ""
Printer.Print "EmpName               :" & Space(10); UCase(Text1(1))
Printer.Print ""
Printer.Print "Father/Husband Name   :" & Space(10); UCase(Text1(12))
Printer.Print ""
Printer.Print "Designation           :" & Space(10); UCase(Combo1)
Printer.Print ""
Printer.Print "Date Of Birth         :" & Space(10); Format(Text1(2), "MMM/DD/YYYY")
Printer.Print ""
Printer.Print "Date Of Joined        :" & Space(10); Format(Text1(3), "MMM/DD/YYYY")
Printer.Print ""
Printer.Print "Basic Salary          :" & Space(10); Format(Text1(4), "0.00")
Printer.Print ""
Printer.Print "DA                    :" & Space(10); Format(Text1(14), "0.00")
Printer.Print ""
Printer.Print "Others                :" & Space(10); Format(Text1(15), "0.00")
Printer.Print ""
Printer.Print "Address1              :" & Space(10); Text1(5)
Printer.Print ""
Printer.Print "Address2              :" & Space(10); Text1(6)
Printer.Print ""
Printer.Print "City                  :" & Space(10); UCase(Text1(7))
Printer.Print ""
Printer.Print "State                 :" & Space(10); UCase(Text1(8))
Printer.Print ""
Printer.Print "Pincode               :" & Space(10); Text1(9)
Printer.Print ""
Printer.Print "PFNo.                 :" & Space(10); Text1(10)
Printer.Print ""
Printer.Print "ESINo.                :" & Space(10); Text1(13)
Printer.Print ""
Printer.Print "Phone No.             :" & Space(10); Text1(11)
Printer.Print ""
Printer.EndDoc
err:
        If err.Number = 482 Or err.Number = 484 Then
            MsgBox "Make Sure The Printer Is Ready", vbExclamation, "Payroll"
        End If
        Exit Sub

End If
End Sub
Private Sub cmdPrint1_Click()
Dim PrintDA
Dim PrintTotal
Msg = MsgBox("Are you take Print", vbExclamation + vbYesNo, "Payroll")
If Msg = 6 Then

On Error GoTo err
'Printer.NewPage
Printer.Orientation = 1
Printer.Font.Size = 12
Printer.Font.Name = "Courier New"
Printer.FontBold = True
rs.Open "Select * from Admin Where CompanyName='" & lblCompanyName.Caption & "' and BranchCode='" & cboBranch & "' order by empno", cnn, adOpenKeyset, adLockOptimistic
Printer.Print Space(40); "Company Name:  " & CompanyName
Printer.Print ""
Printer.Print ""
rs1.Open "Select BranchName from Branch Where CompanyName='" & lblCompanyName.Caption & "' and BranchCode='" & cboBranch & "'", cnn, adOpenKeyset, adLockOptimistic
Printer.Font.Size = 12
Printer.Print Space(40) & "Branch Name:  " & rs1.Fields("BranchName")
rs1.Close
Printer.Print ""
Printer.Font.Size = 10
Printer.FontBold = False
Printer.Print Space(2) & String(160, "-")
Printer.Print Space(2) & "EmpNo" & Space(4) & "EName" & Space(35) & "Designation" & Space(31) & "DOJ" & Space(8) & "Basic" & Space(7) & "DA" & Space(7) & "Others" & Space(7) & "Total"
Printer.Print Space(2) & String(160, "-")
Printer.Print ""
While Not rs.EOF
rs1.Open "Select DA From DesignationMaster Where Designation='" & rs.Fields("Designation") & "'", cnn, adOpenKeyset, adLockOptimistic
PrintDA = rs1.Fields("DA")
rs1.Close
PrintTotal = Format(rs.Fields("Basic") + CDbl(PrintDA) + rs.Fields("Others"), "0.00")

Printer.Print Space(2) & Format(SubName & "/" & rs.Fields("Empno"), "000") & Space(9 - Len(Format(SubName & "/" & rs.Fields("Empno"), "000"))) & UCase(rs.Fields("Ename")) & Space(40 - Len(UCase(rs.Fields("Ename")))) & rs.Fields("Designation") & Space(40 - Len(rs.Fields("Designation"))) & _
Format(rs.Fields("DoJ"), "mm/dd/yyyy") & Space(12 - Len(Format(rs.Fields("DoJ"), "mm/dd/yyyy"))) & Format(rs.Fields("Basic"), "0.00") & Space(12 - Len(Format(rs.Fields("Basic"), "0.00"))) & Format(PrintDA, "0.00") & Space(12 - Len(Format(PrintDA, "0.00"))) & Format(rs.Fields("Others"), "0.00") & Space(12 - Len(Format(rs.Fields("Others"), "0.00"))) & PrintTotal
rs.MoveNext
Printer.Print ""
Wend
rs.Close
Printer.EndDoc

err:
        If err.Number = 482 Or err.Number = 484 Then
            MsgBox "Make Sure The Printer Is Ready", vbExclamation, "Payroll"
        End If
        Exit Sub

End If

End Sub

Private Sub Combo1_Click()
If Combo1 <> "" Then
    rs2.Open "Select Da from DesignationMaster where Designation='" & Combo1 & "'", cnn, adOpenKeyset, adLockOptimistic
    Text1(14) = rs2.Fields(0)
    rs2.Close
End If
End Sub
Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Text1(2).SetFocus
End If
End Sub
Private Sub Command1_Click(Index As Integer)
cboBranch.Enabled = True
List1.Enabled = True
SSTab1.Tab = 0

Select Case Index
    
    Case 0   'New Button
    
        AddEditViewMode = "Add"
        EditEmpno = ""
        EditESI = ""
        EditPfNo = ""
       
        cboBranch.Enabled = False
        List1.Enabled = False
        For i = Text1.LBound To Text1.UBound
            Text1(i) = ""
        Next
        addcomboDesignation
        
        
        Text1(0).SetFocus
        For i = 0 To 3
            Command1(i).Enabled = False
        Next
        cmdPrint.Enabled = False
        cmdPrint1.Enabled = False
        Frame1.Enabled = True
        TextLock (False)
        optPermanent.value = True
        
            
    Case 1
    
        AddEditViewMode = "Edit"
        Text1(0) = Val(Right(Text1(0), Len(Trim(Text1(0))) - InStr(1, Text1(0), "/")))
        EditEmpno = Text1(0)
        EditPfNo = Text1(10)
        EditESI = Text1(13)
        List1.Enabled = False
        Frame1.Enabled = True
        TextLock (False)
        
        For i = 0 To 2
            Command1(i).Enabled = False
        Next
        Command1(3).Enabled = True
        cmdPrint.Enabled = False
        cmdPrint1.Enabled = False
        cboBranch.SetFocus
        
    Case 2
        
        AddEditViewMode = "View"
        
        Msg = MsgBox("Are you Sure to Delete this Record", vbExclamation + vbYesNo, "Payroll")
        If Msg = 6 Then
        
            InBoxResDate = InputBox("Pleas Enter the Resign Date[MM/DD/YYYY]", "Payroll")
            
             If CheckDate(InBoxResDate) = False Then
                MsgBox "Invalid Date", vbExclamation, "Payroll"
               Exit Sub
            End If
            
            If CheckMonth(InBoxResDate) = False Then
                MsgBox "Invalid Month", vbExclamation, "Payroll"
                Exit Sub
            End If
            If InBoxResDate <> "" Then
                AddResEmpMasterTable
                AddResEmpDetails
                AddResEmpLoanMaster
                
                
                cnn.Execute "Delete from Admin Where CompanyName='" & lblCompanyName.Caption & "' and Empno=" & Val(Mid(Text1(0), InStr(1, Text1(0), "/") + 1, Len(Text1(0)) - InStr(1, Text1(0), "/")))
                                
                rs.Open "Select * from admin where  CompanyName='" & lblCompanyName.Caption & "' and EmpNo> " & Val(Mid(Text1(0), InStr(1, Text1(0), "/") + 1, Len(Text1(0)) - InStr(1, Text1(0), "/"))), cnn, adOpenKeyset, adLockOptimistic
        
                While Not rs.EOF
                cnn.Execute "Update Admin set empno=" & rs.Fields("Empno") - 1 & " where CompanyName='" & lblCompanyName.Caption & "' and empno=" & rs.Fields("Empno")
                rs.MoveNext
                Wend
                rs.Close
                CountRecord
           End If
        End If
        
    Case 3
        
       If CheckMonth(Text1(2)) = False Then
            MsgBox "Invalid Month Date Of Birth", vbExclamation, "Payroll"
            Exit Sub
        End If
               
        If CheckMonth(Text1(2)) = False Then
            MsgBox "Invalid Month Date Of Joining", vbExclamation, "Payroll"
            Exit Sub
        End If
          
        If DateDiffer(Text1(2), Text1(3)) = 0 Then
            MsgBox "DOJ Should be Greater Than DOB", vbExclamation, "Payroll"
            Exit Sub
         End If
        'Check  Duplicate PfNo
       If optPermanent.value = True Then
         
       If LCase(Text1(10)) <> LCase(EditPfNo) Then
                rs.Open "Select Count(*) from admin where CompanyName='" & lblCompanyName.Caption & "' and PFNo='" & Text1(10) & "'", cnn, adOpenKeyset, adLockOptimistic
                If rs.Fields(0) > 0 Then
                    MsgBox "PF Number Already Exists", vbExclamation, "Payroll"
                    Cancel = True
                    Text1(10).SetFocus
                    rs.Close
                    Exit Sub
                    
                End If
                rs.Close
            End If
            
            'Check  Duplicate ESI Number
            
        If LCase(Text1(13)) <> LCase(EditESI) Then
                rs.Open "Select Count(*) from admin where CompanyName='" & lblCompanyName.Caption & "' and ESI='" & Text1(13) & "'", cnn, adOpenKeyset, adLockOptimistic
                If rs.Fields(0) > 0 Then
                    MsgBox "ESI Number Already Exists", vbExclamation, "Payroll"
                    Cancel = True
                    Text1(13).SetFocus
                    rs.Close
                    Exit Sub
                End If
                rs.Close
           End If
       End If
       
        If AddEditViewMode = "Add" Then
            rs.Open "select * from admin", cnn, adOpenKeyset, adLockPessimistic
            rs.AddNew
        ElseIf AddEditViewMode = "Edit" Then
       
         rs.Open "select * from admin where CompanyName='" & lblCompanyName.Caption & "' and empno=" & Mid(List1.Text, InStr(1, List1.Text, "/") + 1, InStr(1, List1.Text, "-") - InStr(1, List1.Text, "/") - 1) & " and ename='" & Mid(List1.Text, InStr(1, List1.Text, "-") + 1) & "'", cnn, adOpenKeyset, adLockOptimistic
        End If

            rs.Fields(0) = CInt(Text1(0))
            rs.Fields("ename") = Text1(1)
            rs.Fields("Designation") = Combo1
            rs.Fields("dob") = Text1(2)
            rs.Fields("doj") = Text1(3)
            rs.Fields("basic") = Text1(4)
            rs.Fields("add1") = Text1(5)
            rs.Fields("add2") = Text1(6)
            rs.Fields("city") = Text1(7)
            rs.Fields("State") = Text1(8)
            rs.Fields("pincode") = Text1(9)
            rs.Fields("pfno") = Text1(10)
            rs.Fields("phone") = Text1(11)
            rs.Fields("father_husband") = Text1(12)
            rs.Fields("branchcode") = cboBranch
            rs.Fields("ESI") = Text1(13)
            rs.Fields("Others") = Text1(15)
            rs.Fields("CompanyName") = lblCompanyName.Caption
            If optPermanent.value = True Then
                rs.Fields("Permanent") = True
            Else
                rs.Fields("Permanent") = False
            End If
            
            rs.Update
            rs.Close
        
        AddEditViewMode = "View"
        listadditem
        
        List1.Text = SubName & "/" & Format(Text1(0), "000") & "-" & Text1(1)
        For i = 0 To 2
            Command1(i).Enabled = True
        Next
        Command1(3).Enabled = False
        cmdPrint.Enabled = True
        cmdPrint1.Enabled = True
        Frame1.Enabled = False
        TextLock (True)
        
        
    Case 4
            AddEditViewMode = "View"
            For i = Text1.LBound To Text1.UBound
                Text1(i) = ""
            Next
            If List1.ListCount > 0 Then
                List1.Text = List1.List(0)
            Else
                optPermanent.value = True
            End If
            cbobranch_Click
            Frame1.Enabled = False
            TextLock (True)
                        
    Case 5
       '*** cnn.Close
        Unload Me
        Exit Sub
        
End Select
StatusBar1.Panels(1) = AddEditViewMode
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If AddEditViewMode = "View" Then
If Shift = 4 And KeyCode = 46 And List1.ListCount > 0 Then
    Msg = MsgBox("Are you sure to delete  -  " & cboBranch & " - Branch Code", vbYesNo + vbExclamation, "Payroll")
    If Msg = 6 Then
    InBoxBranch = InputBox("Please Enter the Branch Code", "Payroll")
    If InBoxBranch <> "" Then
    If InBoxBranch = cboBranch Then
         cnn.Execute "Delete from Admin Where CompanyName='" & lblCompanyName.Caption & "' and Branchcode='" & cboBranch & "'"
         CountRecord
    Else
        MsgBox "Please Enter the Match Branch Code", vbExclamation, "Payroll"
    End If
    End If
    End If
End If
End If
End Sub
Private Sub Form_Load()
frmEmployeeDetails.Left = 0
frmEmployeeDetails.Top = 0
frmEmployeeDetails.Icon = LoadPicture(App.Path & "\Jamaica my computer.ico")

SSTab1.Tab = 0
lblCompanyName.Caption = CompanyName
'***cnn.Provider = "Microsoft.Jet.OLEDB.4.0"
'***cnn.Open app.path & "\Moolakadai.mdb" ' payroll1.mdb"
rs1.Open "Select * from Branch where CompanyName='" & lblCompanyName.Caption & "'", cnn, adOpenKeyset, adLockPessimistic
    While Not rs1.EOF
        cboBranch.AddItem rs1.Fields(0)
        rs1.MoveNext
    Wend
rs1.Close
AddEditViewMode = "View"
If cboBranch.ListCount > 0 Then
    cboBranch.ListIndex = 0
Else
    For i = Command1.LBound To Command1.UBound - 1
        Command1(i).Enabled = False
    Next
    cmdPrint.Enabled = False
    cmdPrint1.Enabled = False
End If
TextLock (True)
Frame1.Enabled = False

StatusBar1.Panels(1) = AddEditViewMode

'Combo1.ListIndex = 0
'
'rs.Open "Select count(*) from admin", cnn, adOpenKeyset, adLockPessimistic
'If rs.Fields(0) = 0 Then
'    Command1(0).Enabled = True
'    For i = 1 To 3
'        Command1(i).Enabled = False
'    Next
'Else
'    listadditem
'    List1.ListIndex = 0
'    For i = 0 To 2
'        Command1(i).Enabled = True
'    Next
'    Command1(3).Enabled = False
'End If
'rs.Close
End Sub
Private Sub Form_Unload(Cancel As Integer)

    'Cancel = vbFormControlMenu
End Sub

Private Sub List1_Click()
If AddEditViewMode = "View" Then
Dim EmpnoPos
EmpnoPos = Mid(List1.Text, InStr(1, List1.Text, "/") + 1, InStr(1, List1.Text, "-") - InStr(1, List1.Text, "/") - 1)
rs1.Open "select * from admin where CompanyName='" & lblCompanyName.Caption & "' and empno=" & EmpnoPos & " and ename='" & Mid(List1.Text, InStr(1, List1.Text, "-") + 1) & "'", cnn, adOpenKeyset, adLockOptimistic
'rs1.Open "select * from admin where empno=" & Mid(List1.text, 5, InStr(1, List1.text, "-") - 5) & " and ename='" & Mid(List1.text, InStr(1, List1.text, "-") + 1) & "'"
'rs1.Open "select * from admin where empno=" & Mid(List1, 1, InStr(1, List1.Text, "-") - 1) & " and ename='" & Mid(List1, InStr(1, List1.Text, "-") + 1, Len(List1.Text)) & "'"
Text1(0) = SubName & "/" & Format(rs1.Fields(0), "000")
Text1(1) = rs1.Fields("ename")
Combo1 = rs1.Fields("Designation")
Text1(2) = Format(rs1.Fields("dob"), "mm/dd/yyyy")
Text1(3) = Format(rs1.Fields("doj"), "mm/dd/yyyy")
Text1(4) = rs1.Fields("basic")
Text1(5) = rs1.Fields("add1")
Text1(6) = rs1.Fields("add2")
Text1(7) = rs1.Fields("city")
Text1(8) = rs1.Fields("state")
Text1(9) = rs1.Fields("pincode")
Text1(10) = rs1.Fields("pfno")
Text1(11) = rs1.Fields("phone")
Text1(12) = rs1.Fields("father_husband")
Text1(15) = rs1.Fields("others")
Text1(13) = rs1.Fields("ESI")
'cboBranch = rs1.Fields(14)
If rs1.Fields("Permanent") = True Then
    optPermanent.value = True
Else
    optTraniee.value = True
End If

rs1.Close
For i = 0 To 2
    Command1(i).Enabled = True
Next
Command1(3).Enabled = False
End If
End Sub

Private Sub optPermanent_Click()
If Combo1 <> "" Then
    rs2.Open "Select Da from DesignationMaster where Designation='" & Combo1 & "'", cnn, adOpenKeyset, adLockOptimistic
    Text1(14) = rs2.Fields(0)
    rs2.Close
End If
End Sub
Private Sub optTraniee_Click()
Text1(14).Text = 0


End Sub

Private Sub Text1_Change(Index As Integer)
If AddEditViewMode <> "View" Then
Text1(Index) = LTrim(Text1(Index))
For i = Text1.LBound To Text1.UBound

Select Case i
    
    Case 0 To 4, 7, 9, 10, 13, 15
        If Len(Trim(Text1(0))) <> 0 Then
        If EditEmpno <> Val(Text1(0)) Then
            rs.Open "Select Count(*) from Admin Where CompanyName='" & lblCompanyName.Caption & "' and Empno=" & Val(Text1(0)), cnn, adOpenKeyset, adLockOptimistic
            If rs.Fields(0) > 0 Or Val(Text1(0)) = 0 Then
                StatusBar1.Panels(1) = "Either Empno Already Exists or Empno not Allowed Zero Value"
                Command1(3).Enabled = False
                rs.Close
                Exit Sub
            End If
            rs.Close
            
        End If
        End If
            StatusBar1.Panels(1) = AddEditViewMode
            
            If Len(Text1(i)) = 0 Or Combo1.Text = "" Then
                Command1(3).Enabled = False
                Exit Sub
            End If
                    
End Select
Next
 Command1(3).Enabled = True
End If
End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If AddEditViewMode <> "View" Then

Select Case Index

    Case 1, 12

            If Len(Text1(Index)) > 0 And KeyAscii = 46 Then
            Else
                gspAlphaNumeric Text1(1), KeyAscii
            End If
            
    
    Case 7, 8
        
             If KeyAscii = 39 Then
                KeyAscii = 0
             End If
            
    Case 2, 3
    
            If KeyAscii > 46 And KeyAscii < 59 Or KeyAscii = 8 Then
            Else
                KeyAscii = 0
            End If
            
    Case 0
    
        KeyAscii = NumericCheck1(Text1(Index), CInt(KeyAscii))
        
    Case 9
        If KeyAscii <> 32 Then
            KeyAscii = NumericCheck1(Text1(Index), CInt(KeyAscii))
        End If
        
        
    Case 4, 15
    
        KeyAscii = NumericCheck(Text1(Index), CInt(KeyAscii))
        
    Case 10, 13
                
        'CheckSpecialChar Text1(Index), KeyAscii
        If (KeyAscii >= 33 And KeyAscii <= 44) Or (KeyAscii = 32) Or (KeyAscii = 64) Or (KeyAscii = 124) Or (KeyAscii = 92) Or (KeyAscii = 94) Or (KeyAscii = 96) Or (KeyAscii = 126) Then
            KeyAscii = 0
            MsgBox "Not Accepted Special Character", vbExclamation, "Payroll"
            Exit Sub
        End If
        
    Case 11

          If Len(Text1(11)) > 0 And KeyAscii = 45 Or Len(Text1(11)) > 0 And KeyAscii = 44 Then
          Else
          
            gspNumeric Text1(Index), KeyAscii
        End If
        
End Select
End If
End Sub
Public Sub listadditem()
rs1.Open "select * from admin  where CompanyName='" & lblCompanyName.Caption & "' and BranchCode='" & strBranch & "' order by empno", cnn, adOpenKeyset, adLockPessimistic
    List1.clear
    While Not rs1.EOF
        List1.AddItem SubName & "/" & Format(rs1.Fields(0), "000") & "-" & rs1.Fields(1)
        
        rs1.MoveNext
    Wend
    rs1.Close
End Sub
Private Sub Text1_LostFocus(Index As Integer)

If AddEditViewMode <> "View" Then

Select Case Index

    Case 2, 3
          
         If CheckDate(Text1(Index)) = False Then
            MsgBox "Invalid Date", vbExclamation, "Payroll"
            Text1(Index) = ""
            Cancel = True
            NextTab = False
            Text1(Index).SetFocus
            Exit Sub
        End If
        
        If CheckMonth(Text1(2)) = False Then
            MsgBox "Invalid Month Date Of Birth", vbExclamation, "Payroll"
            Text1(2) = ""
            Text1(2).SetFocus
            NextTab = False
            Exit Sub
        End If
        
        If Format(Text1(Index), "m/d/yyyy") <> Text1(Index) And Format(Text1(Index), "mm/dd/yyyy") <> Text1(Index) Then
            MsgBox "Not Date Format", vbExclamation, "Payroll"
            Text1(Index) = ""
            Text1(Index).SetFocus
             Exit Sub
        End If
        
               
        If CheckMonth(Text1(3)) = False Then
            MsgBox "Invalid Month Date Of Joining", vbExclamation, "Payroll"
            Text1(3) = ""
            Text1(3).SetFocus
            NextTab = False
            Exit Sub
        End If
        
        
         If Text1(2) <> "" And Text1(3) <> "" Then
         If DateDiffer(Text1(2), Text1(3)) <= 0 Then
         
            MsgBox "DOJ Should be Greater Than DOB", vbExclamation, "Payroll"
            NextTab = False
            Exit Sub
         End If
         End If
        
NextTab = True
End Select
End If

If Index = 15 And NextTab = True Then
   SSTab1.Tab = 1
   Text1(12).SetFocus
End If
End Sub
Public Sub TextLock(Locked As Boolean)
For i = Text1.LBound To Text1.UBound
    Text1(i).Locked = Locked
Next
Combo1.Locked = Locked
End Sub
Public Sub addcomboDesignation()
rs.Open "Select * from DesignationMaster", cnn, adOpenKeyset, adLockPessimistic
Combo1.clear
    While Not rs.EOF
        Combo1.AddItem rs.Fields("Designation")
        rs.MoveNext
    Wend
    
rs.Close
End Sub
Public Sub CountRecord()
rs.Open "Select count(*) from admin where CompanyName='" & lblCompanyName.Caption & "' and BranchCode='" & strBranch & "'", cnn, adOpenKeyset, adLockPessimistic
If rs.Fields(0) = 0 Then
    List1.clear
    Command1(0).Enabled = True
    For i = 1 To 3
        Command1(i).Enabled = False
        
    Next
    cmdPrint.Enabled = False
    cmdPrint1.Enabled = False
    clear
Else
    listadditem
    List1.ListIndex = 0
    For i = 0 To 2
        Command1(i).Enabled = True
    Next
    Command1(3).Enabled = False
    cmdPrint.Enabled = True
    cmdPrint1.Enabled = True
End If
rs.Close

End Sub
Public Sub AddResEmpMasterTable()
rs.Open "Select * from Admin Where CompanyName='" & lblCompanyName.Caption & "' and Empno=" & Val(Mid(Text1(0), InStr(1, Text1(0), "/") + 1, Len(Text1(0)) - InStr(1, Text1(0), "/"))), cnn, adOpenKeyset, adLockOptimistic

rs1.Open "Select max(resignNo) from ResEmpMaster Where CompanyName='" & lblCompanyName.Caption & "'", cnn, adOpenKeyset, adLockOptimistic
If rs1.Fields(0) <> "" Then
     ResNo = rs1.Fields(0) + 1
 Else
     ResNo = 1
 End If
 rs1.Close
 
 rs1.Open "Select * from ResEmpMaster", cnn, adOpenKeyset, adLockOptimistic
 
 If Not rs.EOF Then
 rs1.AddNew
 rs1.Fields("ResignNo") = ResNo
 
 For i = 0 To 16
    
    rs1.Fields(i + 1) = rs.Fields(i)
Next
rs1.Fields("ResignDate") = InBoxResDate
rs1.Fields("CompanyName") = CompanyName
rs1.Update
rs1.Close
rs.Close
 'cnn.Execute "Insert into ResEmpMaster values(" & Resno & "," & rs("Empno") & ",'" & rs("Ename") & "','" & rs("Designation") & "'," & rs("dob") & "," & rs("doj") & "," & rs("Basic") & ",'" & rs("Add1") & "','" & rs("Add2") & "','" & rs("City") & "','" & rs("State") & "'," & rs("Pincode") & ",'" & rs("PfNo") & "','" & rs("Phone") & "','" & rs("Father_Husband") & "','" & rs("BranchCode") & "'," & rs("Others") & ",'" & rs("ESI") & "'," & InBoxResDate & ")"
 
 End If
End Sub
Public Sub AddResEmpDetails()
'MsgBox Mid(Text1(0), , Len(Text1(0)) - 4)
rs.Open "Select * from Daughters Where CompanyName='" & lblCompanyName.Caption & "' and Empno=" & Val(Mid(Text1(0), InStr(1, Text1(0), "/") + 1, Len(Text1(0)) - InStr(1, Text1(0), "/"))), cnn, adOpenKeyset, adLockOptimistic
rs1.Open "Select * from ResEmpPayDetails", cnn, adOpenKeyset, adLockOptimistic

While Not rs.EOF
rs1.AddNew
rs1.Fields("ResignNo") = ResNo
For i = 0 To 17
rs1.Fields(i + 1) = rs.Fields(i)
Next
rs1.Fields("CompanyName") = CompanyName
rs1.Update
rs.MoveNext
Wend
rs1.Close
rs.Close
End Sub
Public Sub AddResEmpLoanMaster()
rs.Open "Select * from LoanMaster Where CompanyName='" & lblCompanyName.Caption & "' and Empno=" & Val(Mid(Text1(0), InStr(1, Text1(0), "/") + 1, Len(Text1(0)) - InStr(1, Text1(0), "/"))), cnn, adOpenKeyset, adLockOptimistic
rs1.Open "Select * from ResLoanMaster", cnn, adOpenKeyset, adLockPessimistic
While Not rs.EOF
rs1.AddNew
rs1.Fields("ResignNo") = ResNo
    For i = 0 To 6
    
        rs1.Fields(i + 1) = rs.Fields(i)
    Next
    rs1.Fields("CompanyName") = CompanyName
rs1.Update
rs.MoveNext
Wend
rs1.Close
rs.Close
End Sub
