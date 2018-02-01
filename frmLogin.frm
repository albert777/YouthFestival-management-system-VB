VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5430
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   5430
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1875
      Left            =   0
      TabIndex        =   4
      Top             =   -60
      Width           =   5415
      Begin VB.CommandButton cmdcancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   3
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton cmdlogin 
         Caption         =   "Login"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   2
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtpassword 
         Height          =   390
         IMEMode         =   3  'DISABLE
         Left            =   1440
         MaxLength       =   50
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   720
         Width           =   3855
      End
      Begin VB.TextBox txtuser 
         Height          =   390
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   0
         Top             =   240
         Width           =   3855
      End
      Begin VB.Label lblpassword 
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lbluser 
         Caption         =   "User name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    End
End Sub

Private Sub cmdlogin_Click()
On Error GoTo lbl
    If rsNewUser.State = 1 Then rsNewUser.Close
    rsNewUser.Open "Select * from tbl_login where log_name='" & txtuser.Text & "' and log_password='" & txtpassword.Text & "'", cn, adOpenStatic, adLockPessimistic
    If rsNewUser.EOF = False Then
        If rsNewUser.Fields("log_status") = "user" Then
            MDIForm1.mnuAdmin.Enabled = False
        End If
        Unload Me
        MDIForm1.Show
    Else
        MsgBox "Username or Password doesn't find a match", vbInformation, "Message"
    End If
     Exit Sub
lbl:
    MsgBox Err.Description, vbInformation, "Message"
End Sub

Private Sub Form_Load()
    cn.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=db_YouthFestival;Data Source=HP-PC"

End Sub

Private Sub txtpassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdlogin.SetFocus
End Sub

Private Sub txtuser_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtpassword.SetFocus
End Sub
