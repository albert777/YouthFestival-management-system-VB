VERSION 5.00
Begin VB.Form frmPrize 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Prize"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11670
   Icon            =   "frmPrize.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   11670
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3795
      Left            =   60
      TabIndex        =   5
      Top             =   0
      Width           =   11595
      Begin VB.ComboBox cmbProgram 
         Height          =   1545
         Left            =   1440
         Style           =   1  'Simple Combo
         TabIndex        =   0
         Top             =   240
         Width           =   4455
      End
      Begin VB.ComboBox cmbPrizeCategory 
         Height          =   1740
         Left            =   1440
         Style           =   1  'Simple Combo
         TabIndex        =   1
         Top             =   1920
         Width           =   4455
      End
      Begin VB.TextBox txtPrize 
         Height          =   405
         Left            =   7020
         MaxLength       =   50
         TabIndex        =   2
         Top             =   240
         Width           =   4455
      End
      Begin VB.TextBox txtDescription 
         Height          =   1845
         Left            =   7020
         MaxLength       =   50
         TabIndex        =   3
         Top             =   720
         Width           =   4455
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   435
         Left            =   9240
         TabIndex        =   4
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   435
         Left            =   10380
         TabIndex        =   6
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Program"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Prize Category"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Prize"
         Height          =   255
         Left            =   6060
         TabIndex        =   8
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Description"
         Height          =   255
         Left            =   6060
         TabIndex        =   7
         Top             =   780
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmPrize"
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

Private Sub cmbPrizeCategory_Click()
    If cmbProgram.ListIndex <> -1 And cmbPrizeCategory.ListIndex <> -1 Then
        If rsPrize.State = 1 Then rsPrize.Close
        rsPrize.Open "select * from tbl_prize where prize_category_id=" & cmbPrizeCategory.ItemData(cmbPrizeCategory.ListIndex) & _
        " and program_id=" & cmbProgram.ItemData(cmbProgram.ListIndex), cn, adOpenStatic, adLockPessimistic
        If rsPrize.EOF = False Then
            txtPrize.Text = rsPrize.Fields("prize")
            txtDescription.Text = rsPrize.Fields("description")
        Else
            txtPrize.Text = ""
            txtDescription.Text = ""
        End If
    Else
         txtPrize.Text = ""
        txtDescription.Text = ""
    End If
    
End Sub

Private Sub cmbPrizeCategory_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyDown = 13 Then txtPrize.SetFocus
End Sub

Private Sub cmbProgram_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then cmbPrizeCategory.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo lbl
    If cmbProgram.ListIndex <> -1 And cmbPrizeCategory.ListIndex <> -1 And Len(Trim(txtPrize.Text)) <> 0 Then
        If rsPrize.State = 1 Then rsPrize.Close
        rsPrize.Open "select * from tbl_prize where prize_category_id=" & cmbPrizeCategory.ItemData(cmbPrizeCategory.ListIndex) & _
        " and program_id=" & cmbProgram.ItemData(cmbProgram.ListIndex), cn, adOpenStatic, adLockPessimistic
        If rsPrize.EOF = True Then
            cn.Execute "insert into tbl_prize(program_id,prize_category_id,prize,description) values(" & _
            cmbProgram.ItemData(cmbProgram.ListIndex) & "," & cmbPrizeCategory.ItemData(cmbPrizeCategory.ListIndex) & _
            ",'" & txtPrize.Text & "','" & txtDescription.Text & "')"
        Else
            If MsgBox("Already Exists,Do U want to modify", vbYesNo + vbQuestion, "Warning") = vbYes Then
                If rsPrize.State = 1 Then rsPrize.Close
                cn.Execute "update tbl_prize set(prize='" & txtPrize.Text & "',description='" & txtDescription.Text & _
                " where prize_category_id=" & cmbPrizeCategory.ItemData(cmbPrizeCategory.ListIndex) & _
                " and program_id=" & cmbProgram.ItemData(cmbProgram.ListIndex)
            End If
        
        End If
        Clearing
        cmbProgram.SetFocus
    Else
        MsgBox "Select Program,Prize Category and enter Prize", vbInformation, "Message"
    End If
        Exit Sub
lbl:
    MsgBox Err.Description, vbInformation, "Message"
    
End Sub

Private Sub Form_Load()
    AddProgram
    AddCategory
End Sub
Public Sub AddCategory()
    If rsPrizeCategory.State = 1 Then rsPrizeCategory.Close
    rsPrizeCategory.Open "select * from tbl_prize_category", cn, adOpenStatic, adLockPessimistic
    cmbPrizeCategory.Clear
    While Not rsPrizeCategory.EOF
        cmbPrizeCategory.AddItem rsPrizeCategory.Fields("category_name")
        cmbPrizeCategory.ItemData(cmbPrizeCategory.NewIndex) = rsPrizeCategory.Fields("prize_category_id")
        rsPrizeCategory.MoveNext
    Wend
End Sub

Public Sub Clearing()
    cmbProgram.ListIndex = -1
    cmbProgram.Text = ""
    cmbPrizeCategory.ListIndex = -1
    cmbPrizeCategory.Text = ""
    txtPrize.Text = ""
    txtDescription.Text = ""
End Sub



Private Sub txtDescription_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdSave.SetFocus
End Sub

Private Sub txtPrize_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtDescription.SetFocus
End Sub
