VERSION 5.00
Begin VB.Form frmChangePassword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ChangePassword"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6630
   Icon            =   "frmChangePassword.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   6630
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6495
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   435
         Left            =   5280
         TabIndex        =   10
         Top             =   1800
         Width           =   1035
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   435
         Left            =   4140
         TabIndex        =   9
         Top             =   1800
         Width           =   1035
      End
      Begin VB.TextBox txtCPassword 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2220
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   1380
         Width           =   4095
      End
      Begin VB.TextBox txtNewPassword 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2220
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   1005
         Width           =   4095
      End
      Begin VB.TextBox txtOldPassword 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2220
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   615
         Width           =   4095
      End
      Begin VB.TextBox txtUsername 
         Height          =   315
         Left            =   2220
         TabIndex        =   2
         Top             =   240
         Width           =   4095
      End
      Begin VB.Label Label4 
         Caption         =   "Confirm Password"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Width           =   1875
      End
      Begin VB.Label Label3 
         Caption         =   "New Password"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   1875
      End
      Begin VB.Label Label2 
         Caption         =   "Old Password"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   660
         Width           =   1875
      End
      Begin VB.Label Label1 
         Caption         =   "Username"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   1875
      End
   End
End
Attribute VB_Name = "frmChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo lbl
    If rsNewUser.State = 1 Then rsNewUser.Close
    rsNewUser.Open "select * from tbl_login where log_name='" & txtUsername.Text & "'and log_password='" & txtOldPassword.Text & "'", cn, adOpenStatic, adLockPessimistic
    If rsNewUser.EOF = False Then
        If txtNewPassword.Text = txtCPassword.Text Then
            If MsgBox("Are U sure to change password", vbYesNo + vbQuestion) = vbYes Then
                If rsNewUser.State = 1 Then rsNewUser.Close
                cn.Execute "update tbl_login set log_password='" & txtNewPassword.Text & "' where log_name='" & txtUsername.Text & "'and log_password='" & txtOldPassword.Text & "'"
            End If
        Else
            MsgBox "Password and confirmpassword doesn't find a match", vbInformation, "Message"
        End If
    Else
        MsgBox "User and Password doesn't find a match", vbInformation, "Message"
    End If
    
Exit Sub
lbl:
    MsgBox Err.Description, vbInformation, "Message"
End Sub

Private Sub txtCPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdsave.SetFocus
End Sub

Private Sub txtNewPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtCPassword.SetFocus
End Sub

Private Sub txtOldPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNewPassword.SetFocus
End Sub

Private Sub txtUsername_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtOldPassword.SetFocus
End Sub
