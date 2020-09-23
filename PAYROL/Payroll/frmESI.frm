VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmESI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Range of ESI"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5220
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
   ScaleHeight     =   3105
   ScaleWidth      =   5220
   Begin VB.TextBox txtESI 
      Height          =   330
      Left            =   1680
      MaxLength       =   10
      TabIndex        =   2
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox txtFromAmt 
      Height          =   345
      Left            =   1680
      MaxLength       =   15
      TabIndex        =   0
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox txtToAmt 
      Height          =   330
      Left            =   1680
      MaxLength       =   15
      TabIndex        =   1
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   1755
      TabIndex        =   5
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3390
      TabIndex        =   7
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   375
      Left            =   4200
      TabIndex        =   8
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   2565
      TabIndex        =   6
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   930
      TabIndex        =   4
      Top             =   2160
      Width           =   735
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   2730
      Width           =   5220
      _ExtentX        =   9208
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4022
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "12:24 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "10/10/01"
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cboFromAmt 
      Height          =   345
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Range of  ESI"
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   1440
      TabIndex        =   16
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   " %"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   3360
      TabIndex        =   15
      Top             =   1560
      Width           =   270
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "ESI"
      Height          =   225
      Left            =   360
      TabIndex        =   14
      Top             =   1560
      Width           =   285
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "From Amount"
      Height          =   225
      Left            =   360
      TabIndex        =   13
      Top             =   600
      Width           =   1125
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "To Amount"
      Height          =   225
      Left            =   360
      TabIndex        =   12
      Top             =   1080
      Width           =   900
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "(Basic+DA)"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   3480
      TabIndex        =   11
      Top             =   600
      Width           =   960
   End
End
Attribute VB_Name = "frmESI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AddEditViewMode As String
Dim Msg As String
Dim EditFromAmt
Dim EditToAmt
Dim SetrangeFromAmt
Dim SetrangeToAmt
Dim i As Integer
Private Sub cboFromAmt_Click()
txtFromAmt = cboFromAmt
rs1.Open "Select * from RangeESI where FromAmount=" & cboFromAmt, cnn, adOpenKeyset, adLockOptimistic
txtFromAmt = rs1.Fields("FromAmount")
txtToAmt = rs1.Fields("ToAmount")
txtESI = rs1.Fields("ESI") * 100
rs1.Close

rs1.Open "Select max(FromAmount) from RangeESI", cnn, adOpenKeyset, adLockOptimistic
If rs1.Fields(0) = Val(cboFromAmt) Then
    cmdDelete.Enabled = True
Else
    cmdDelete.Enabled = False
End If
rs1.Close
End Sub
Private Sub CmdCancel_Click()
AddEditViewMode = "View"
CountRecord
textLocked True
StatusBar1.Panels(1) = AddEditViewMode
End Sub
Private Sub CmdDelete_Click()
Msg = MsgBox("Are you sure to delete this amount", vbExclamation + vbYesNo, "Payroll")
If Msg = 6 Then
    cnn.Execute "Delete from RangeEsi where fromamount=" & cboFromAmt
    CountRecord
End If
End Sub
Private Sub cmdEdit_Click()
AddEditViewMode = "Edit"
EditFromAmt = txtFromAmt
EditToAmt = txtToAmt
rs.Open "Select max(ToAmount) from RangeESI where fromAmount<" & EditFromAmt
If Not IsNull(rs.Fields(0)) Then
    SetrangeFromAmt = rs.Fields(0) + 0.01
Else
    SetrangeFromAmt = 0
End If
rs.Close
rs.Open "Select min(fromAmount) from RangeESI where ToAmount>" & EditToAmt
If Not IsNull(rs.Fields(0)) Then
    SetrangeToAmt = rs.Fields(0) - 0.01
Else
    SetrangeToAmt = " And Above"
End If
rs.Close
txtFromAmt.ZOrder 0
txtFromAmt.SetFocus
textLocked False
cmdNew.Enabled = False
cmdEdit.Enabled = False
cmdDelete.Enabled = False
cmdSave.Enabled = True
StatusBar1.Panels(1) = AddEditViewMode & "-" & SetrangeFromAmt & "-" & SetrangeToAmt
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdNew_Click()
AddEditViewMode = "Add"
clear
cmdNew.Enabled = False
cmdEdit.Enabled = False
cmdDelete.Enabled = False
cmdSave.Enabled = True
textLocked False
txtFromAmt.ZOrder 0
txtFromAmt.SetFocus
StatusBar1.Panels(1) = AddEditViewMode
End Sub
Private Sub CmdSave_Click()
If SaveEnableDisable = False Then
    MsgBox "Please Enter the All Amount"
    Exit Sub
End If
If Val(txtFromAmt) >= Val(txtToAmt) Then
    MsgBox "ToAmount Should be Greater then FromAmount", vbExclamation, "Payroll"
    Exit Sub
End If

If AddEditViewMode = "Add" Then
    rs.Open "Select max(toAmount) from RangeESI", cnn, adOpenKeyset, adLockOptimistic

        If rs.Fields(0) >= Val(txtFromAmt) Then
            MsgBox "This range of Amount already  Inserted", vbExclamation, "Payroll"
            rs.Close
            Exit Sub
        End If
        rs.Close
End If


If AddEditViewMode = "Edit" Then
If SetrangeToAmt = " And Above" Then
   SetrangeToAmt = Val(txtToAmt)
End If

If Val(SetrangeFromAmt) <= Val(txtFromAmt) And Val(SetrangeToAmt) >= Val(txtToAmt) Then
Else
MsgBox "Please Check Out of the range FromAmount or ToAmount", vbExclamation, "Payroll"
Exit Sub
End If
End If
If AddEditViewMode = "Add" Then
    rs.Open "Select * from RangeESI", cnn, adOpenKeyset, adLockOptimistic
    rs.AddNew
ElseIf AddEditViewMode = "Edit" Then
    rs.Open "Select * from RangeESI where FromAmount=" & EditFromAmt, cnn, adOpenKeyset, adLockOptimistic
End If
    rs.Fields("FromAmount") = txtFromAmt
    rs.Fields("ToAmount") = txtToAmt
    rs.Fields("ESI") = Val(txtESI) / 100
    rs.Update
    rs.Close
AddEditViewMode = "View"
cboFromAmt.ZOrder 0
AddListFromAmt
cboFromAmt.Text = Val(txtFromAmt)
cmdNew.Enabled = True
cmdEdit.Enabled = True
cmdDelete.Enabled = True
cmdSave.Enabled = False
textLocked True
StatusBar1.Panels(1) = AddEditViewMode
End Sub
Private Sub Form_Activate()
frmESI.Left = 0
frmESI.Top = 0
AddEditViewMode = "View"
CountRecord
textLocked True
StatusBar1.Panels(1) = AddEditViewMode
End Sub
Private Sub txtFromAmt_Change()
txtFromAmt = Trim(txtFromAmt)
End Sub
Private Sub txtFromAmt_KeyPress(KeyAscii As Integer)
KeyAscii = NumericCheck(txtFromAmt, CInt(KeyAscii))
End Sub
Private Sub txtESI_Change()
txtESI = Trim(txtESI)
End Sub
Private Sub txtESI_KeyPress(KeyAscii As Integer)
KeyAscii = NumericCheck(txtESI, CInt(KeyAscii))
End Sub
Private Sub txtToAmt_Change()
txtToAmt = Trim(txtToAmt)
End Sub
Private Sub txtToAmt_KeyPress(KeyAscii As Integer)
KeyAscii = NumericCheck(txtToAmt, CInt(KeyAscii))
End Sub
Public Sub textLocked(CheckLock As Boolean)
txtFromAmt.Locked = CheckLock
txtToAmt.Locked = CheckLock
txtESI.Locked = CheckLock
End Sub
Public Sub AddListFromAmt()
rs1.Open "Select fromAmount from RangeESI order by fromamount", cnn, adOpenKeyset, adLockOptimistic
cboFromAmt.clear
While Not rs1.EOF
cboFromAmt.AddItem rs1.Fields("FromAmount")
rs1.MoveNext
Wend
rs1.Close
End Sub
Public Function SaveEnableDisable() As Boolean
For i = 0 To Screen.ActiveForm.Count - 1
    If TypeOf Screen.ActiveForm.Controls(i) Is TextBox Then
      If Trim(Len(Screen.ActiveForm.Controls(i).Text)) = 0 Then
            SaveEnableDisable = False
            Exit Function
      End If
      End If
Next
SaveEnableDisable = True
End Function
Public Sub CountRecord()
rs.Open "Select count(*) from RangeESI", cnn, adOpenKeyset, adLockOptimistic
If rs.Fields(0) > 0 Then
    cboFromAmt.ZOrder 0
    AddListFromAmt
    cboFromAmt.ListIndex = 0
    cmdNew.Enabled = True
    cmdEdit.Enabled = True
    cmdSave.Enabled = False
Else
    txtFromAmt.ZOrder 0
    clear
    cmdNew.Enabled = True
    cmdEdit.Enabled = False
    cmdDelete.Enabled = False
    cmdSave.Enabled = False
    cmdNew.SetFocus
End If
rs.Close
End Sub

