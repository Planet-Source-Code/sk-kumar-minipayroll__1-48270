VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmdesignationMaster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Designation Master"
   ClientHeight    =   2895
   ClientLeft      =   3105
   ClientTop       =   2265
   ClientWidth     =   6165
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   6165
   Begin VB.CommandButton Command1 
      Caption         =   "&Delete"
      Height          =   375
      Index           =   2
      Left            =   1920
      TabIndex        =   12
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   375
      Index           =   5
      Left            =   4440
      TabIndex        =   7
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   4
      Left            =   3600
      TabIndex        =   6
      Top             =   2040
      Width           =   855
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   2520
      Width           =   6165
      _ExtentX        =   10874
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5689
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "5:02 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "11/2/01"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Save"
      Height          =   375
      Index           =   3
      Left            =   2760
      TabIndex        =   5
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Edit"
      Height          =   375
      Index           =   1
      Left            =   1080
      TabIndex        =   4
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&New"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Index           =   1
      Left            =   1560
      MaxLength       =   6
      TabIndex        =   2
      Text            =   "814"
      Top             =   1440
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      Height          =   345
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   840
      Width           =   2700
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Index           =   0
      Left            =   1560
      MaxLength       =   40
      TabIndex        =   1
      Top             =   840
      Width           =   2700
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Designation Master"
      ForeColor       =   &H00800080&
      Height          =   225
      Left            =   2272
      TabIndex        =   11
      Top             =   120
      Width           =   1620
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "DA"
      Height          =   225
      Left            =   360
      TabIndex        =   9
      Top             =   1440
      Width           =   270
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Designation"
      Height          =   225
      Left            =   360
      TabIndex        =   8
      Top             =   840
      Width           =   990
   End
End
Attribute VB_Name = "frmdesignationMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'***Dim cnn As New ADODB.Connection
'***Dim rs As New ADODB.Recordset
'*** Dim rs1 As New ADODB.Recordset
Dim AddEditViewMode As String
Dim i As Integer
Dim EditDesg  As String
Dim Msg As String
Private Sub Combo1_Click()
If Combo1 <> "" Then
    Text1(0) = Combo1
    rs1.Open "Select * from  DesignationMaster where Designation='" & Combo1 & "'", cnn, adOpenKeyset, adLockOptimistic
    Text1(1) = rs1.Fields("da")
    rs1.Close
    rs1.Open "Select count(*) from Admin Where Designation='" & Combo1 & "'", cnn, adOpenKeyset, adLockOptimistic
    If rs1.Fields(0) > 0 Then
        Command1(2).Enabled = False
    Else
         Command1(2).Enabled = True
    End If
    rs1.Close
End If
End Sub

Private Sub Command1_Click(Index As Integer)
Select Case Index
    
    Case 0          'Add Button
    
        AddEditViewMode = "Add"
        TextLock (False)
        Text1(0).ZOrder 0
        Text1(0).Text = ""
        Text1(1).Text = 814
        Text1(0).SetFocus
        For i = 0 To 3
            Command1(i).Enabled = False
        Next
        StatusBar1.Panels(1) = AddEditViewMode
        
        
    Case 1    'Edit Button
    
        AddEditViewMode = "Edit"
        TextLock (False)
        EditDesg = Text1(0)
        For i = 0 To 2
            Command1(i).Enabled = False
        Next
        Command1(3).Enabled = True
        Text1(0).ZOrder 0
        Text1(0).SetFocus
        StatusBar1.Panels(1) = AddEditViewMode
        
            
    Case 2    'Delete Button
        
        AddEditViewMode = "View"
        Msg = MsgBox("Are you sure to Delete This Record", vbExclamation + vbYesNo, "Payroll")
        If Msg = 6 Then
            cnn.Execute "Delete from DesignationMaster where Designation='" & Trim(Combo1) & "'"
        End If
        CountRecord
        StatusBar1.Panels(1) = AddEditViewMode
        
    Case 3   'Save Button
        
        If AddEditViewMode = "Add" Then
            rs.Open "Select * from  DesignationMaster", cnn, adOpenKeyset, adLockOptimistic
            rs.AddNew
        ElseIf AddEditViewMode = "Edit" Then
            rs.Open "Select * from  DesignationMaster Where Designation='" & EditDesg & "'", cnn, adOpenKeyset, adLockOptimistic
        End If
            rs.Fields("Designation") = Text1(0)
            rs.Fields("da") = Text1(1)
            rs.Update
            rs.Close
        AddEditViewMode = "View"
        For i = 0 To 2
            Command1(i).Enabled = True
        Next
        Command1(3).Enabled = False
        Combo1.ZOrder 0
        AddRecToCombo
        TextLock (True)
        StatusBar1.Panels(1) = AddEditViewMode
        
        
    Case 4
        
            AddEditViewMode = "View"
            CountRecord
            StatusBar1.Panels(1) = AddEditViewMode
    Case 5
       '*** cnn.Close
        Unload Me
        
    
End Select

    
End Sub
Private Sub Form_Activate()

AddEditViewMode = "View"
CountRecord
TextLock (True)
StatusBar1.Panels(1) = AddEditViewMode
End Sub

Private Sub Form_Load()
frmdesignationMaster.Left = 0
frmdesignationMaster.Top = 0
frmdesignationMaster.Icon = LoadPicture(App.Path & "\Jamaica my computer.ico")
'***cnn.Provider = "Microsoft.Jet.OLEDB.4.0"
'*** cnn.Open app.path & "\Moolakadai.mdb" ' payroll1.mdb"
End Sub
Private Sub Text1_Change(Index As Integer)
On Error GoTo err
If AddEditViewMode <> "View" Then
Text1(Index) = LTrim(Text1(Index))
If LCase(Text1(0)) <> LCase(EditDesg) Then
 rs.Open "Select count(*) from  DesignationMaster where Designation='" & Text1(0) & "'", cnn, adOpenKeyset, adLockOptimistic
 If rs.Fields(0) = 1 Then
    StatusBar1.Panels(1) = "Already exists"
    Command1(3).Enabled = False
    rs.Close
    Exit Sub
    
 End If
 rs.Close
End If


For i = Text1.LBound To Text1.UBound
    If Len(Text1(i)) = 0 Then
        Command1(3).Enabled = False
        Exit Sub
    End If
Next
Command1(3).Enabled = True
End If
   StatusBar1.Panels(1) = AddEditViewMode
err:
    Exit Sub
End Sub
Public Sub AddRecToCombo()
 rs1.Open "Select * from  DesignationMaster", cnn, adOpenKeyset, adLockOptimistic
 Combo1.clear
 While Not rs1.EOF
    Combo1.AddItem rs1.Fields(0)
    rs1.MoveNext
 Wend
 rs1.Close
 Combo1.ListIndex = 0
End Sub
Public Sub CountRecord()
rs.Open "Select count(*) from  DesignationMaster", cnn, adOpenKeyset, adLockOptimistic
Text1(0) = ""
Combo1.clear
If rs.Fields(0) > 0 Then
    For i = 0 To 2
        Command1(i).Enabled = True
    Next
    
    Command1(3).Enabled = False
    Combo1.ZOrder 0
    AddRecToCombo
Else
        Text1(0).ZOrder 0
        Command1(1).Enabled = False
        Command1(2).Enabled = False
        Command1(3).Enabled = False
        Command1(0).Enabled = True
    
End If
rs.Close
End Sub
Public Sub TextLock(Locked As Boolean)
Text1(0).Locked = Locked
Text1(1).Locked = Locked
End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If AddEditViewMode <> "View" Then
Select Case Index
    
    Case 0
            
            'gspAlphaNumeric Text1(0), KeyAscii
            If (KeyAscii >= 49 And KeyAscii <= 57) Or (KeyAscii = 48) Then
                's.Locked = True
                KeyAscii = 0
                MsgBox "Enter Character  Values", vbExclamation, "Payroll"
                Exit Sub
            End If
            If (KeyAscii >= 33 And KeyAscii <= 44) Or (KeyAscii = 64) Or (KeyAscii = 124) Or (KeyAscii = 92) Or (KeyAscii = 94) Or (KeyAscii = 96) Or (KeyAscii = 126) Then
                KeyAscii = 0
                MsgBox "Only String Values Accepted", vbExclamation, "Payroll"
                Exit Sub
            End If
    
    Case 1
            
            KeyAscii = NumericCheck(Text1(Index), CInt(KeyAscii))
    
'            If KeyAscii > 47 And KeyAscii < 59 Or KeyAscii = 8 Then
'            Else
'                KeyAscii = 0
'            End If
End Select
End If
End Sub
