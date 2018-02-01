VERSION 5.00
Begin VB.Form frmProgramwise 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Programwise"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6150
   Icon            =   "frmProgramwise.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   6150
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   60
      TabIndex        =   3
      Top             =   0
      Width           =   6075
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   435
         Left            =   4800
         TabIndex        =   6
         Top             =   2220
         Width           =   1095
      End
      Begin VB.CommandButton cmdShow 
         Caption         =   "&Show"
         Height          =   435
         Left            =   3720
         TabIndex        =   2
         Top             =   2220
         Width           =   1095
      End
      Begin VB.TextBox txtDescription 
         Height          =   315
         Left            =   1560
         TabIndex        =   1
         Top             =   1860
         Width           =   4335
      End
      Begin VB.ComboBox cmbProgram 
         Height          =   1545
         Left            =   1560
         Style           =   1  'Simple Combo
         TabIndex        =   0
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label Label2 
         Caption         =   "Description"
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   1800
         Width           =   1515
      End
      Begin VB.Label Label1 
         Caption         =   "Program Name"
         Height          =   435
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1875
      End
   End
End
Attribute VB_Name = "frmProgramwise"
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

Private Sub cmbProgram_Change()
    cmbProgram_Click
End Sub

Private Sub cmbProgram_Click()
    If cmbProgram.ListIndex <> -1 Then
        If rsProgram.State = 1 Then rsProgram.Close
        rsProgram.Open "select * from tbl_program where program_id=" & cmbProgram.ItemData(cmbProgram.ListIndex), cn, adOpenStatic, adLockPessimistic
        If rsProgram.EOF = False Then
            txtDescription.Text = rsProgram.Fields("program_description")
        Else
            txtDescription.Text = ""
        End If
    Else
        txtDescription.Text = ""
    End If
End Sub

Private Sub cmbProgram_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtDescription.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdShow_Click()
    If MenuItem = "Member" Then
        rptMemberPgmwise.Show
    ElseIf MenuItem = "Score" Then
        rptScorePgmwise.Show
    End If
End Sub

Private Sub Form_Load()
    AddProgram
End Sub


Private Sub txtDescription_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdShow.SetFocus
End Sub
