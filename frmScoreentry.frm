VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmScoreentry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Score Entry"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11550
   Icon            =   "frmScoreentry.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   11550
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   6435
      Left            =   60
      TabIndex        =   8
      Top             =   0
      Width           =   11475
      Begin VB.TextBox txtMemberRegId 
         Height          =   495
         Left            =   6120
         TabIndex        =   18
         Top             =   5700
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox txtContactNo 
         Height          =   375
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   4560
         Width           =   4335
      End
      Begin VB.TextBox txtAddress 
         Height          =   855
         Left            =   7020
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   4080
         Width           =   4275
      End
      Begin VB.TextBox txtCandidate 
         Height          =   375
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   4080
         Width           =   4335
      End
      Begin VB.ComboBox cmbJudge 
         Height          =   1350
         Left            =   1560
         Style           =   1  'Simple Combo
         TabIndex        =   5
         Top             =   4980
         Width           =   4335
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   495
         Left            =   10080
         TabIndex        =   14
         Top             =   5820
         Width           =   1215
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   495
         Left            =   8820
         TabIndex        =   7
         Top             =   5820
         Width           =   1215
      End
      Begin VB.ComboBox cmbProgram 
         Height          =   1350
         Left            =   1380
         Style           =   1  'Simple Combo
         TabIndex        =   0
         Top             =   240
         Width           =   4455
      End
      Begin VB.TextBox txtPgmDescription 
         Height          =   375
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   4215
      End
      Begin VB.TextBox txtScore 
         Height          =   435
         Left            =   7020
         MaxLength       =   4
         TabIndex        =   6
         Top             =   4980
         Width           =   2115
      End
      Begin MSFlexGridLib.MSFlexGrid flxGrid 
         Height          =   2295
         Left            =   60
         TabIndex        =   9
         Top             =   1680
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   4048
         _Version        =   393216
         Cols            =   5
      End
      Begin VB.Label Label7 
         Caption         =   "Contact No"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   4560
         Width           =   1395
      End
      Begin VB.Label Label6 
         Caption         =   "Address"
         Height          =   435
         Left            =   6120
         TabIndex        =   16
         Top             =   4020
         Width           =   1635
      End
      Begin VB.Label Label5 
         Caption         =   "Candidate Name"
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   4080
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Program"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   3315
      End
      Begin VB.Label Label2 
         Caption         =   "Description"
         Height          =   435
         Left            =   5880
         TabIndex        =   12
         Top             =   180
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "Judge Name"
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Top             =   4980
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "Score"
         Height          =   435
         Left            =   6120
         TabIndex        =   10
         Top             =   4980
         Width           =   2115
      End
   End
End
Attribute VB_Name = "frmScoreentry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub AddProgram()
    If rsProgram.State = 1 Then rsProgram.Close
    rsProgram.Open "select * from tbl_program", cn, adOpenStatic, adLockPessimistic
    cmbProgram.Clear
    While Not rsProgram.EOF
        cmbProgram.AddItem rsProgram.Fields("program_name")
        cmbProgram.ItemData(cmbProgram.NewIndex) = rsProgram.Fields("program_id")
        rsProgram.MoveNext
    Wend
End Sub

Private Sub cmbJudge_Click()
On Error GoTo lbl
    If cmbProgram.ListIndex <> -1 And cmbJudge.ListIndex <> -1 And Val(txtMemberRegId.Text) <> 0 Then
        If rsScore.State = 1 Then rsScore.Close
        rsScore.Open "select * from tbl_score_entering where program_id=" & cmbProgram.ItemData(cmbProgram.ListIndex) & _
        " and judge_id=" & cmbJudge.ItemData(cmbJudge.ListIndex) & " and student_id=" & Val(txtMemberRegId.Text), cn, adOpenStatic, adLockPessimistic
        If rsScore.EOF = False Then
            txtScore.Text = rsScore.Fields("score")
        Else
            txtScore.Text = ""
        End If
    Else
        'Clearing
        txtScore.Text = ""
    End If
Exit Sub
lbl:
    MsgBox Err.Description, vbInformation, "Message"
End Sub

Private Sub cmbJudge_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtScore.SetFocus
End Sub

Private Sub cmbProgram_Change()
    cmbProgram_Click
End Sub

Private Sub cmbProgram_Click()
On Error GoTo lbl
     If cmbProgram.ListIndex <> -1 Then
        If rsProgram.State = 1 Then rsProgram.Close
        rsProgram.Open "select * from tbl_program where program_id=" & cmbProgram.ItemData(cmbProgram.ListIndex), cn, adOpenStatic, adLockPessimistic
        If rsProgram.EOF = False Then
            txtPgmDescription.Text = rsProgram.Fields("program_description")
        Else
            txtPgmDescription.Text = ""
        End If
        If rsMemberReg.State = 1 Then rsMemberReg.Close
        rsMemberReg.Open "select * from tbl_member_registration,tbl_MemberPgm where tbl_member_registration.member_registration_id=tbl_MemberPgm.MemberReg_id and tbl_MemberPgm.program_id=" & cmbProgram.ItemData(cmbProgram.ListIndex), cn, adOpenStatic, adLockPessimistic
        DispGrid
        While Not rsMemberReg.EOF
            flxGrid.TextMatrix(flxGrid.Rows - 1, 0) = flxGrid.Rows - 1
            flxGrid.TextMatrix(flxGrid.Rows - 1, 1) = rsMemberReg.Fields("name")
            flxGrid.TextMatrix(flxGrid.Rows - 1, 2) = rsMemberReg.Fields("address")
            flxGrid.TextMatrix(flxGrid.Rows - 1, 3) = rsMemberReg.Fields("contact_no")
            flxGrid.TextMatrix(flxGrid.Rows - 1, 4) = rsMemberReg.Fields("member_registration_id")
            rsMemberReg.MoveNext
            flxGrid.Rows = flxGrid.Rows + 1
        Wend
        If rsJudges.State = 1 Then rsJudges.Close
        rsJudges.Open "Select * from tbl_judges,tbl_PgmJudge,tbl_program_scheduling where tbl_program_scheduling.program_scheduling_id=tbl_PgmJudge.program_sheduling_id and tbl_PgmJudge.judge_id=tbl_judges.judge_id and tbl_program_scheduling.program_id=" & cmbProgram.ItemData(cmbProgram.ListIndex), cn, adOpenStatic, adLockPessimistic
        cmbJudge.Clear
        While Not rsJudges.EOF
            cmbJudge.AddItem rsJudges.Fields("name")
            cmbJudge.ItemData(cmbJudge.NewIndex) = rsJudges.Fields("judge_id")
            rsJudges.MoveNext
        Wend
    Else
        txtPgmDescription.Text = ""
        DispGrid
        cmbJudge.Clear
    End If
Exit Sub
lbl:
    MsgBox Err.Description, vbInformation, "Message"
End Sub

Private Sub cmbProgram_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtAddress.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo lbl
    If Val(txtMemberRegId.Text) <> 0 And cmbJudge.ListIndex <> -1 And cmbProgram.ListIndex <> -1 And Val(txtScore.Text) <> 0 Then
        If rsScore.State = 1 Then rsScore.Close
        rsScore.Open "select * from tbl_score_entering where program_id=" & cmbProgram.ItemData(cmbProgram.ListIndex) & _
        " and student_id=" & Val(txtMemberRegId.Text) & " and judge_id=" & cmbJudge.ItemData(cmbJudge.ListIndex), cn, adOpenStatic, adLockPessimistic
        If rsScore.EOF = True Then
            cn.Execute "insert into tbl_score_entering(program_id,student_id,judge_id,score) values(" & cmbProgram.ItemData(cmbProgram.ListIndex) & _
            "," & Val(txtMemberRegId.Text) & "," & cmbJudge.ItemData(cmbJudge.ListIndex) & "," & Val(txtScore.Text) & ")"
        Else
            If MsgBox("Already Exists,Do U Want to modify", vbYesNo + vbQuestion, "Warning") = vbYes Then
                If rsScore.State = 1 Then rsScore.Close
                cn.Execute "update tbl_score_entering set score=" & Val(txtScore.Text) & " where program_id=" & cmbProgram.ItemData(cmbProgram.ListIndex) & _
                " and student_id=" & Val(txtMemberRegId.Text) & " and judge_id=" & cmbJudge.ItemData(cmbJudge.ListIndex)
            End If
        End If
        Clearing
        cmbJudge.Text = ""
        txtScore.Text = ""
    Else
        MsgBox "Select Program,Candidate,Judge and enter score", vbInformation, "Message"
     End If
    Exit Sub
lbl:
    MsgBox Err.Description, vbInformation, "Message"
   
End Sub

Private Sub flxGrid_DblClick()
On Error GoTo lbl
    If flxGrid.RowSel > 0 Then
        txtCandidate.Text = flxGrid.TextMatrix(flxGrid.RowSel, 1)
        txtAddress.Text = flxGrid.TextMatrix(flxGrid.RowSel, 2)
        txtContactNo.Text = flxGrid.TextMatrix(flxGrid.RowSel, 3)
        txtMemberRegId.Text = flxGrid.TextMatrix(flxGrid.RowSel, 4)
    Else
        Clearing
      End If
        Exit Sub
lbl:
    MsgBox Err.Description, vbInformation, "Message"
  
End Sub

Private Sub Form_Load()
    AddProgram
End Sub

Public Sub DispGrid()
    flxGrid.Clear
    flxGrid.Rows = 2
    flxGrid.ColWidth(0) = 1000
    flxGrid.ColWidth(1) = 2500
    flxGrid.ColWidth(2) = 2500
    flxGrid.ColWidth(3) = 2000
    flxGrid.ColWidth(4) = 0
    
    flxGrid.TextMatrix(0, 0) = "Sl No"
    flxGrid.TextMatrix(0, 1) = "Student Name"
    flxGrid.TextMatrix(0, 2) = "Address"
    flxGrid.TextMatrix(0, 3) = "Contact No"
    flxGrid.TextMatrix(0, 4) = "MemberRegID"
    
End Sub

Public Sub Clearing()
    txtCandidate.Text = ""
    txtAddress.Text = ""
    txtContactNo.Text = ""
    txtMemberRegId.Text = ""
End Sub








Private Sub txtAddress_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtContactNo.SetFocus
End Sub

Private Sub txtCandidate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtAddress.SetFocus
End Sub



Private Sub txtContactNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmbJudge.SetFocus
End Sub

Private Sub Txtcontactno_LostFocus()
 If Not ValPhone(txtContactNo.Text) Then
        MsgBox "Not a valid contactno", vbInformation, "Message"
    End If
End Sub

Private Sub txtPgmDescription_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtCandidate.SetFocus
End Sub



Private Sub txtScore_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdSave.SetFocus
End Sub
