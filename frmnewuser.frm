VERSION 5.00
Begin VB.Form frmnewuser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New user"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6390
   Icon            =   "frmnewuser.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   6390
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   60
      TabIndex        =   7
      Top             =   60
      Width           =   6315
      Begin VB.CommandButton cmdcancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   4980
         TabIndex        =   6
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton cmdsave 
         Caption         =   "&Save"
         Height          =   375
         Left            =   3600
         TabIndex        =   5
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton optuser 
         Caption         =   "User"
         Height          =   495
         Left            =   2760
         TabIndex        =   4
         Top             =   1380
         Width           =   1215
      End
      Begin VB.OptionButton optadmin 
         Caption         =   "Admin"
         Height          =   495
         Left            =   1800
         TabIndex        =   3
         Top             =   1380
         Width           =   1215
      End
      Begin VB.TextBox txtconfirm 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1800
         MaxLength       =   50
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1020
         Width           =   4395
      End
      Begin VB.TextBox txtpassword 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1800
         MaxLength       =   50
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   600
         Width           =   4395
      End
      Begin VB.TextBox txtuser 
         Height          =   315
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   0
         Top             =   180
         Width           =   4395
      End
      Begin VB.Label lblstatus 
         Caption         =   "User status"
         Height          =   555
         Left            =   240
         TabIndex        =   11
         Top             =   1500
         Width           =   1455
      End
      Begin VB.Label lblconfirm 
         Caption         =   "Confirm password"
         Height          =   495
         Left            =   180
         TabIndex        =   10
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblpassword 
         Caption         =   "Password"
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lbluser 
         Caption         =   "User name"
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   180
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmnewuser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo lbl
If txtpassword.Text = txtconfirm.Text Then
Dim s As String
If optadmin.Value = True Then
    s = "admin"
Else
    s = "user"
End If
If rsNewUser.State = 1 Then rsNewUser.Close
rsNewUser.Open "select * from tbl_login where log_name='" & txtuser.Text & "'", cn, adOpenStatic, adLockPessimistic
If rsNewUser.EOF = True Then
cn.Execute "insert into tbl_login (log_name,log_password,log_status) values('" & _
txtuser.Text & "','" & txtpassword.Text & "','" & s & "')"
Else
 MsgBox "User Already Exist", vbInformation, "message"
 End If
 Else
 MsgBox "Password and Confirm password does not find a match", vbInformation, "message"
  End If
 Exit Sub
lbl:
    MsgBox Err.Description, vbInformation, "Message"


End Sub

Private Sub txtconfirm_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdSave.SetFocus
End Sub

Private Sub txtpassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtconfirm.SetFocus
End Sub

Private Sub txtuser_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtpassword.SetFocus
End Sub
