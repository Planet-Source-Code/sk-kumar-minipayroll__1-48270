VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCompany 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Company Title"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7050
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   7050
   Begin VB.TextBox txtSubName 
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
      Left            =   1440
      MaxLength       =   5
      TabIndex        =   2
      Top             =   840
      Width           =   1815
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   2745
      Width           =   7050
      _ExtentX        =   12435
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4895
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   1
            Object.Width           =   4895
            TextSave        =   "3:14 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "10/18/01"
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
   Begin VB.ComboBox cboCompanyName 
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
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   4815
   End
   Begin VB.CommandButton CmdCancel 
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
      Left            =   4320
      TabIndex        =   9
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton CmdSave 
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
      Left            =   3360
      TabIndex        =   8
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton CmdDelete 
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
      Left            =   2400
      TabIndex        =   7
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton cmdNew 
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
      Left            =   480
      TabIndex        =   6
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton cmdExit 
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
      Left            =   5400
      TabIndex        =   5
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton cmdEdit 
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
      Left            =   1440
      TabIndex        =   4
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox txtCompanyName 
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
      Left            =   1440
      MaxLength       =   150
      TabIndex        =   1
      Top             =   240
      Width           =   4815
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Sub Name"
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
      TabIndex        =   11
      Top             =   840
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Company Name : "
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
      Left            =   0
      TabIndex        =   3
      Top             =   240
      Width           =   1410
   End
End
Attribute VB_Name = "frmCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AddEditViewMode As String
Dim EditCompanyName As String
Dim Msg As String
Private Sub cboCompanyName_Click()
AddEditViewMode = "View"
txtCompanyName = cboCompanyName
rs1.Open "Select SubName from Company where CompanyName='" & cboCompanyName & "'", cnn, adOpenKeyset, adLockOptimistic
txtSubName = rs1.Fields("SubName")
rs1.Close
If cboCompanyName.ListCount <= 1 Then
    cmdDelete.Enabled = False
    Exit Sub
End If
rs1.Open "Select count(*) from Branch where CompanyName='" & cboCompanyName & "'", cnn, adOpenKeyset, adLockOptimistic
    If rs1.Fields(0) > 0 Then
       cmdDelete.Enabled = False
    Else
        cmdDelete.Enabled = True
    End If
    rs1.Close
End Sub
Private Sub CmdCancel_Click()
AddEditViewMode = "View"
cboCompanyName.ZOrder 0
AddCompany
cmdNew.Enabled = True
cmdEdit.Enabled = True
cmdSave.Enabled = False
StatusBar1.Panels(1) = AddEditViewMode
End Sub
Private Sub CmdDelete_Click()
Msg = MsgBox("Are you Sure to Delete this Company", vbExclamation + vbYesNo, "Payroll")
If Msg = 6 Then
    cnn.Execute "Delete from Company where CompanyName='" & cboCompanyName & "'"
    AddCompany
End If

End Sub

Private Sub cmdEdit_Click()
AddEditViewMode = "Edit"
EditCompanyName = txtCompanyName
cmdNew.Enabled = False
cmdEdit.Enabled = False
cmdDelete.Enabled = False
cmdSave.Enabled = True
txtCompanyName.ZOrder 0
txtCompanyName.SetFocus
StatusBar1.Panels(1) = AddEditViewMode
End Sub
Private Sub cmdExit_Click()

Unload Me
End Sub
Private Sub cmdNew_Click()
AddEditViewMode = "Add"
cmdNew.Enabled = False
cmdEdit.Enabled = False
cmdDelete.Enabled = False
cmdSave.Enabled = False
txtCompanyName.ZOrder 0
txtCompanyName = ""
txtSubName = ""
txtCompanyName.SetFocus
StatusBar1.Panels(1) = AddEditViewMode
End Sub
Private Sub CmdSave_Click()
If AddEditViewMode = "Add" Then
    rs.Open "Select * from Company", cnn, adOpenKeyset, adLockOptimistic
    rs.AddNew
ElseIf AddEditViewMode = "Edit" Then
    rs.Open "Select * from Company where CompanyName='" & EditCompanyName & "'", cnn, adOpenKeyset, adLockOptimistic
End If

    rs.Fields("CompanyName") = txtCompanyName
    rs.Fields("SubName") = txtSubName
    rs.Update
    rs.Close
    AddEditViewMode = "View"
    cboCompanyName.ZOrder 0
    AddCompany
    cmdNew.Enabled = True
    cmdEdit.Enabled = True
    cmdSave.Enabled = False
    
    StatusBar1.Panels(1) = AddEditViewMode
End Sub
Private Sub Form_Activate()
frmCompany.Left = 0
frmCompany.Top = 0
AddEditViewMode = "View"
rs.Open "Select count(*) from Company", cnn, adOpenKeyset, adLockOptimistic
If rs.Fields(0) > 0 Then
   cboCompanyName.ZOrder 0
   AddCompany
   cmdNew.Enabled = True
   cmdEdit.Enabled = True
   cmdSave.Enabled = False
'Else
 '   txtCompanyName.ZOrder 0
  '  cboCompanyName.clear
   ' cmdNew.Enabled = True
   ' CmdDelete.Enabled = True
    
End If
rs.Close


StatusBar1.Panels(1) = AddEditViewMode

End Sub
Public Sub AddCompany()
rs1.Open "Select * from Company", cnn, adOpenKeyset
cboCompanyName.clear
While Not rs1.EOF
cboCompanyName.AddItem rs1.Fields("CompanyName")
rs1.MoveNext
Wend
rs1.Close
If cboCompanyName.ListCount > 0 Then
    cboCompanyName.ListIndex = 0
End If

End Sub
Private Sub Form_Unload(Cancel As Integer)
CompanyName = cboCompanyName
SubName = txtSubName
MDIForm1.Caption = CompanyName & "-" & "Payroll"
End Sub
Private Sub txtCompanyName_Change()
If AddEditViewMode <> "View" Then
    txtCompanyName = LTrim(txtCompanyName)
    
        cmdSave.Enabled = CheckEmptyText
       If cmdSave.Enabled = False Then
           Exit Sub
       End If
       
    
    If LCase(Trim(txtCompanyName)) <> LCase(Trim(EditCompanyName)) Then
        rs.Open "Select Count(*) from Company where CompanyName='" & txtCompanyName & "'", cnn, adOpenKeyset, adLockOptimistic
        If rs.Fields(0) > 0 Then
            StatusBar1.Panels(1) = "Already Records Exists"
            cmdSave.Enabled = False
            rs.Close
            Exit Sub
        End If
    rs.Close
    End If
    
    cmdSave.Enabled = True
End If

End Sub
Public Function CheckEmptyText() As Boolean
If Len(Trim(txtCompanyName)) = 0 Or Len(Trim(txtSubName)) = 0 Then
        CheckEmptyText = False
        Exit Function
End If
CheckEmptyText = True
End Function
Private Sub txtCompanyName_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Or KeyAscii = 45 Then
    KeyAscii = 0
End If
End Sub
Private Sub txtSubName_Change()
If AddEditViewMode <> "View" Then
    cmdSave.Enabled = CheckEmptyText
End If


End Sub
Private Sub txtSubName_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Or KeyAscii = 45 Then
    KeyAscii = 0
End If

End Sub
