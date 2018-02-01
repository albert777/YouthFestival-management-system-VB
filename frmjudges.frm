VERSION 5.00
Begin VB.Form frmjudges 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Judjes"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12165
   Icon            =   "frmjudges.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   12165
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   5595
      Left            =   60
      TabIndex        =   10
      Top             =   0
      Width           =   12075
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   495
         Left            =   10740
         TabIndex        =   21
         Top             =   4140
         Width           =   1035
      End
      Begin VB.ComboBox cmbJudgeName 
         Height          =   1545
         Left            =   1080
         Style           =   1  'Simple Combo
         TabIndex        =   3
         Top             =   2280
         Width           =   4755
      End
      Begin VB.Frame Frame2 
         Caption         =   "Program Details"
         Height          =   1875
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   11655
         Begin VB.TextBox txtTime 
            Height          =   435
            Left            =   7140
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   1320
            Width           =   2175
         End
         Begin VB.TextBox txtPDescription 
            Height          =   1095
            Left            =   7140
            Locked          =   -1  'True
            TabIndex        =   1
            Top             =   180
            Width           =   4395
         End
         Begin VB.ComboBox cmbPName 
            Height          =   1545
            Left            =   960
            Style           =   1  'Simple Combo
            TabIndex        =   0
            Top             =   180
            Width           =   4815
         End
         Begin VB.Label Label9 
            Caption         =   "Time in Minutes"
            Height          =   315
            Left            =   5940
            TabIndex        =   20
            Top             =   1380
            Width           =   1275
         End
         Begin VB.Label Label8 
            Caption         =   "Description"
            Height          =   435
            Left            =   5940
            TabIndex        =   19
            Top             =   180
            Width           =   1995
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Program"
            Height          =   255
            Left            =   60
            TabIndex        =   18
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.TextBox txtAddress 
         Height          =   1575
         Left            =   1080
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   3900
         Width           =   4755
      End
      Begin VB.TextBox txtQualification 
         Height          =   435
         Left            =   7080
         MaxLength       =   50
         TabIndex        =   5
         Top             =   2220
         Width           =   4695
      End
      Begin VB.TextBox txtExperiance 
         Height          =   405
         Left            =   7080
         MaxLength       =   50
         TabIndex        =   6
         Top             =   2700
         Width           =   4695
      End
      Begin VB.TextBox txtContactNo 
         Height          =   435
         Left            =   7080
         MaxLength       =   10
         TabIndex        =   7
         Top             =   3180
         Width           =   4695
      End
      Begin VB.TextBox txtEmail 
         Height          =   405
         Left            =   7080
         MaxLength       =   50
         TabIndex        =   8
         Top             =   3660
         Width           =   4695
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   495
         Left            =   9660
         TabIndex        =   9
         Top             =   4140
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Judge Name"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   2340
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Address"
         Height          =   255
         Left            =   180
         TabIndex        =   15
         Top             =   3960
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Qualification"
         Height          =   255
         Left            =   6120
         TabIndex        =   14
         Top             =   2220
         Width           =   975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Experience"
         Height          =   195
         Left            =   6120
         TabIndex        =   13
         Top             =   2700
         Width           =   795
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Contact No."
         Height          =   255
         Left            =   6120
         TabIndex        =   12
         Top             =   3180
         Width           =   975
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Email"
         Height          =   195
         Left            =   6120
         TabIndex        =   11
         Top             =   3660
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmjudges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub AddProgram()
    If rsProgram.State = 1 Then rsProgram.Close
    rsProgram.Open "select * from tbl_program", cn, adOpenStatic, adLockPessimistic
    cmbPName.Clear
    While Not rsProgram.EOF
        cmbPName.AddItem rsProgram.Fields("program_name")
        cmbPName.ItemData(cmbPName.NewIndex) = rsProgram.Fields("program_id")
        rsProgram.MoveNext
    Wend
End Sub

Private Sub cmbJudgeName_Change()
    cmbJudgeName_Click
End Sub

Private Sub cmbJudgeName_Click()
On Error GoTo lbl
    If cmbJudgeName.ListIndex <> -1 Then
        If rsJudges.State = 1 Then rsJudges.Close
        rsJudges.Open "select * from tbl_judges where judge_id=" & cmbJudgeName.ItemData(cmbJudgeName.ListIndex), cn, adOpenStatic, adLockPessimistic
        If rsJudges.EOF = False Then
            txtAddress.Text = rsJudges.Fields("address")
            txtContactNo.Text = rsJudges.Fields("contact_no")
            Txtemail.Text = rsJudges.Fields("email")
            txtExperiance.Text = rsJudges.Fields("experience")
            txtQualification.Text = rsJudges.Fields("qualification")
        Else
            Clearing2
        End If
    Else
        Clearing2
    End If
Exit Sub
lbl:
    MsgBox Err.Description, vbInformation, "Message"
End Sub

Private Sub cmbJudgeName_DropDown()
If KeyCode = 13 Then txtAddress.SetFocus
End Sub

Private Sub cmbPName_Click()
On Error GoTo lbl
      If cmbPName.ListIndex <> -1 Then
        If rsProgram.State = 1 Then rsProgram.Close
        rsProgram.Open "select * from tbl_program where program_id=" & cmbPName.ItemData(cmbPName.ListIndex), cn, adOpenStatic, adLockPessimistic
        If rsProgram.EOF = False Then
            txtPDescription.Text = rsProgram.Fields("Program_description")
            txtTime.Text = rsProgram.Fields("program_time")
        Else
            Clearing1
        End If
        If rsJudges.State = 1 Then rsJudges.Close
        rsJudges.Open "select * from tbl_judges where program_id=" & cmbPName.ItemData(cmbPName.ListIndex), cn, adOpenStatic, adLockPessimistic
        cmbJudgeName.Clear
        While Not rsJudges.EOF
            cmbJudgeName.AddItem rsJudges.Fields("name")
            cmbJudgeName.ItemData(cmbJudgeName.NewIndex) = rsJudges.Fields("judge_id")
            rsJudges.MoveNext
        Wend
    Else
        Clearing1
    End If
    Exit Sub
lbl:
    MsgBox Err.Description, vbInformation, "Message"
    
    
End Sub

Private Sub cmbPName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtPDescription.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
  On Error GoTo lbl
    If cmbPName.ListIndex <> -1 And Len(Trim(cmbJudgeName.Text)) <> 0 Then
        If cmbJudgeName.ListIndex <> -1 Then
            If rsJudges.State = 1 Then rsJudges.Close
            rsJudges.Open "Select * from tbl_judges where judge_id=" & cmbPName.ItemData(cmbPName.ListIndex), cn, adOpenStatic, adLockPessimistic
            If rsJudges.EOF = True Then
                cn.Execute "insert into tbl_judges(program_id,name,address,qualification,experience,contact_no,email)Values(" & _
                cmbPName.ItemData(cmbPName.ListIndex) & ",'" & cmbJudgeName.Text & "','" & txtAddress.Text & _
                "','" & txtQualification.Text & "','" & txtExperiance.Text & "','" & txtContactNo.Text & "','" & Txtemail.Text & "')"
            Else
                If MsgBox("Already exists,Do u want to modify", vbYesNo + vbQuestion) = vbYes Then
                    If rsJudges.State = 1 Then rsJudges.Close
                    cn.Execute "update tbl_judges set program_id=" & cmbPName.ItemData(cmbPName.ListIndex) & _
                    ",name='" & cmbJudgeName.Text & "',address='" & txtAddress.Text & "',qualification='" & _
                    txtQualification.Text & "',experience='" & txtExperiance.Text & "',contact_no='" & txtContactNo.Text & _
                    "',email='" & Txtemail.Text & "' where judge_id=" & cmbJudgeName.ItemData(cmbJudgeName.ListIndex)
                End If
            End If
        Else
            cn.Execute "insert into tbl_judges(program_id,name,address,qualification,experience,contact_no,email)Values(" & _
            cmbPName.ItemData(cmbPName.ListIndex) & ",'" & cmbJudgeName.Text & "','" & txtAddress.Text & _
            "','" & txtQualification.Text & "','" & txtExperiance.Text & "','" & txtContactNo.Text & "','" & Txtemail.Text & "')"
        End If
        Clearing1
        Clearing2
        cmbPName.Text = ""
        cmbJudgeName.Text = ""
        cmbPName.ListIndex = -1
        cmbJudgeName.ListIndex = -1
        cmbPName.SetFocus
    Else
        MsgBox "Select Program and enter Judgename", vbInformation, "Message"
     End If
    Exit Sub
lbl:
    MsgBox Err.Description, vbInformation, "Message"
   
End Sub

Private Sub Form_Load()
    AddProgram
End Sub

Public Sub Clearing1()
    txtPDescription.Text = ""
    txtTime.Text = ""
End Sub

Public Sub Clearing2()
    txtAddress.Text = ""
    txtContactNo.Text = ""
    Txtemail.Text = ""
    txtExperiance.Text = ""
    txtPDescription.Text = ""
    txtQualification.Text = ""
    
    
End Sub




Private Sub txtAddress_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtQualification.SetFocus
End Sub




Private Sub txtContactNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Txtemail.SetFocus
End Sub



Private Sub Txtcontactno_LostFocus()
 If Not ValPhone(txtContactNo.Text) Then
        MsgBox "Not a valid contactno", vbInformation, "Message"
    End If
End Sub

Private Sub Txtemail_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdsave.SetFocus
End Sub

Private Sub txtEmail_LostFocus()
ValEmail (Txtemail.Text)
End Sub

Private Sub txtExperiance_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtContactNo.SetFocus
End Sub

Private Sub txtPDescription_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtTime.SetFocus

End Sub



Private Sub txtQualification_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtExperiance.SetFocus
End Sub

Private Sub txtTime_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmbJudgeName.SetFocus
End Sub
