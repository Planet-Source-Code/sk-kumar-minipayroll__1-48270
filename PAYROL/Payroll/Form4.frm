VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form4 
   Caption         =   "Employee Details"
   ClientHeight    =   7650
   ClientLeft      =   195
   ClientTop       =   480
   ClientWidth     =   11175
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   ScaleHeight     =   7650
   ScaleWidth      =   11175
   WindowState     =   2  'Maximized
   Begin VB.ListBox List1 
      Height          =   5340
      Left            =   6840
      TabIndex        =   19
      Top             =   960
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit{F7}"
      Height          =   450
      Index           =   5
      Left            =   7080
      TabIndex        =   18
      Top             =   6600
      Width           =   1335
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   17
      Top             =   7275
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11509
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "11:52 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "8/10/01"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear{F6}"
      Height          =   450
      Index           =   4
      Left            =   5760
      TabIndex        =   16
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save{F5}"
      Height          =   450
      Index           =   3
      Left            =   4500
      TabIndex        =   15
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Delete{F4}"
      Height          =   450
      Index           =   2
      Left            =   3240
      TabIndex        =   14
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Edit{F3}"
      Height          =   450
      Index           =   1
      Left            =   1980
      TabIndex        =   13
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New{F2}"
      Height          =   450
      Index           =   0
      Left            =   720
      TabIndex        =   12
      Top             =   6600
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2235
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   2160
      TabIndex        =   1
      Top             =   1419
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   2
      Left            =   2160
      TabIndex        =   3
      Top             =   3042
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   3
      Left            =   2160
      TabIndex        =   4
      Top             =   3861
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   4
      Left            =   2160
      TabIndex        =   5
      Top             =   4680
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   2055
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   450
      Left            =   480
      TabIndex        =   20
      Top             =   6600
      Width           =   8175
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   6840
      X2              =   8895
      Y1              =   720
      Y2              =   735
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "LIST Of EMPLOYEE :"
      Height          =   240
      Left            =   6840
      TabIndex        =   21
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "BASIC"
      Height          =   240
      Left            =   960
      TabIndex        =   11
      Top             =   4680
      Width           =   675
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "DOJ"
      Height          =   240
      Left            =   960
      TabIndex        =   10
      Top             =   3861
      Width           =   465
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "DOB"
      Height          =   240
      Left            =   960
      TabIndex        =   9
      Top             =   3042
      Width           =   495
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Class"
      Height          =   240
      Left            =   960
      TabIndex        =   8
      Top             =   2238
      Width           =   600
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "EName"
      Height          =   240
      Left            =   960
      TabIndex        =   7
      Top             =   1419
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "EmpNo."
      Height          =   240
      Left            =   960
      TabIndex        =   6
      Top             =   600
      Width           =   840
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim AddEditViewMode As String

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Text1(2).SetFocus
End If
End Sub

Private Sub Command1_Click(Index As Integer)

Select Case Index
    
    Case 0
    
        AddEditViewMode = "Add"
        For i = Text1.LBound To Text1.UBound
            Text1(i) = ""
        Next
        Combo1.ListIndex = 0
        
        rs.Open "select max(empno) from admin", cnn, adOpenKeyset, adLockPessimistic
        If IsNull(rs.Fields(0)) Then
            Text1(0) = "STP/" & "01"
        Else
            Text1(0) = "STP/" & Format(rs.Fields(0) + 1, "00")
        End If
        rs.Close
        
        Text1(1).SetFocus
        For i = 0 To 3
            Command1(i).Enabled = False
        Next
        lock1 (False)
        Combo1.locked = False
            
    Case 1
    
        AddEditViewMode = "Edit"
        lock1 (False)
        Text1(0).locked = True
        For i = 0 To 2
            Command1(i).Enabled = False
        Next
        Command1(3).Enabled = True
        Text1(1).SetFocus
    Case 2
        
        AddEditViewMode = "Add"
        
    Case 3
        If AddEditViewMode = "Add" Then
            rs.Open "select * from admin", cnn, adOpenKeyset, adLockPessimistic
            rs.AddNew
        ElseIf AddEditViewMode = "Edit" Then
        
        rs.Open "select * from admin where empno=" & Mid(List1.text, 5, InStr(1, List1.text, "-") - 5) & " and ename='" & Mid(List1.text, InStr(1, List1.text, "-") + 1) & "'", cnn, adOpenKeyset, adLockPessimistic
        
           ' rs.Open "Select * from admin where empno=" & Mid(List1, 1, InStr(1, List1.Text, "-") - 1) & " and ename='" & Mid(List1, InStr(1, List1.Text, "-") + 1, Len(List1.Text)) & "'", cnn, adOpenKeyset, adLockPessimistic
        End If
            'rs.Fields(0) = CInt(Right(Text1(0), Len(Text1(0)) - 4))
            rs.Fields(1) = Text1(1)
            rs.Fields(2) = Combo1
            rs.Fields(3) = Text1(2)
            rs.Fields(4) = Text1(3)
            rs.Fields(5) = Text1(4)
            rs.Update
            rs.Close
        
        listadditem
        List1.text = Text1(0) & "-" & Text1(1)
        For i = 0 To 2
            Command1(i).Enabled = True
        Next
        Command1(3).Enabled = False
        AddEditViewMode = "View"
        lock1 (True)
        Combo1.locked = True
        
            
    Case 4
            For i = Text1.LBound To Text1.UBound
                Text1(i) = ""
            Next
            If List1.ListCount > 0 Then
                List1.text = List1.List(0)
            End If
            AddEditViewMode = "View"
            
            
    Case 5
        cnn.Close
        Unload Me
        Exit Sub
        
End Select
StatusBar1.Panels(1) = AddEditViewMode
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 113 To 117
        If Command1(KeyCode - 113).Enabled = True Then
            Command1_Click (KeyCode - 113)
        End If
End Select
End Sub
Private Sub Form_Load()
cnn.Provider = "Microsoft.Jet.OLEDB.4.0"
cnn.Open "d:\project\Moolakadai.mdb" ' payroll1.mdb"
Combo1.AddItem "Manager"
Combo1.AddItem "Engg"
Combo1.AddItem "Supervisor"
Combo1.AddItem "Worker"
Combo1.ListIndex = 0

rs.Open "Select count(*) from admin", cnn, adOpenKeyset, adLockPessimistic
If rs.Fields(0) = 0 Then
    Command1(0).Enabled = True
    For i = 1 To 3
        Command1(i).Enabled = False
    Next
Else
    listadditem
    List1.ListIndex = 0
    For i = 0 To 2
        Command1(i).Enabled = True
    Next
    Command1(3).Enabled = False
End If
rs.Close
lock1 (True)
AddEditViewMode = "View"

End Sub

Private Sub List1_Click()
rs1.Open "select * from admin where empno=" & Mid(List1.text, 5, InStr(1, List1.text, "-") - 5) & " and ename='" & Mid(List1.text, InStr(1, List1.text, "-") + 1) & "'"
'rs1.Open "select * from admin where empno=" & Mid(List1, 1, InStr(1, List1.Text, "-") - 1) & " and ename='" & Mid(List1, InStr(1, List1.Text, "-") + 1, Len(List1.Text)) & "'"
Text1(0) = "STP/" & Format(rs1.Fields(0), "00")
Text1(1) = rs1.Fields(1)
Combo1 = rs1.Fields(2)
Text1(2) = rs1.Fields(3)
Text1(3) = rs1.Fields(4)
Text1(4) = rs1.Fields(5)
rs1.Close
For i = 0 To 2
    Command1(i).Enabled = True
Next
Command1(3).Enabled = False
End Sub

Private Sub mnudaughters_Click()
Form5.Show vbModal
End Sub

Private Sub mnureport_Click()
Form6.Show vbModal
End Sub

Private Sub Text1_Change(Index As Integer)
If AddEditViewMode <> "View" Then
Text1(Index) = LTrim(Text1(Index))
For i = Text1.LBound To Text1.UBound
    If Len(Text1(i)) = 0 Then
        Command1(3).Enabled = False
        Exit Sub
    End If
Next
Command1(3).Enabled = True
End If
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Select Case Index
    
    Case 1
        
        Combo1.SetFocus
        
    Case 2, 3
        
        If Not IsDate(Text1(Index).text) Then
            MsgBox "Please Enter Valid Date"
            Text1(Index).text = ""
        Else
            Text1(Index + 1).SetFocus
        End If
End Select
End If
End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)

Select Case Index
    
    Case 2, 3
    
            If KeyAscii > 46 And KeyAscii < 59 Or KeyAscii = 8 Then
            Else
                KeyAscii = 0
            End If
    
    Case 4
    
        If KeyAscii > 47 And KeyAscii < 59 Or KeyAscii = 8 Then
        Else
            KeyAscii = 0
        End If
End Select

End Sub

Public Sub listadditem()
rs1.Open "select * from admin", cnn, adOpenKeyset, adLockPessimistic
    List1.clear
    While Not rs1.EOF
        List1.AddItem "STP/" & Format(rs1.Fields(0), "00") & "-" & rs1.Fields(1)
        rs1.MoveNext
    Wend
    rs1.Close
    
End Sub

Public Sub lock1(locked As Boolean)
For i = Text1.LBound To Text1.UBound
         Text1(i).locked = locked
Next
Combo1.locked = locked
End Sub


