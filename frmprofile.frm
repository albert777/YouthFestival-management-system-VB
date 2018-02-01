VERSION 5.00
Begin VB.Form frmprofile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Profile"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8235
   Icon            =   "frmprofile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   8235
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3495
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   8115
      Begin VB.TextBox txtMail 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   8
         Top             =   2880
         Width           =   4395
      End
      Begin VB.TextBox txtFax 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   7
         Top             =   2400
         Width           =   4395
      End
      Begin VB.TextBox txtPhone 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   6
         Top             =   1920
         Width           =   4395
      End
      Begin VB.TextBox txtTitle3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         MaxLength       =   50
         TabIndex        =   5
         Top             =   1380
         Width           =   7395
      End
      Begin VB.TextBox txtTitle2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         MaxLength       =   50
         TabIndex        =   4
         Top             =   840
         Width           =   7395
      End
      Begin VB.TextBox txtTitle1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         MaxLength       =   50
         TabIndex        =   3
         Top             =   300
         Width           =   7395
      End
      Begin VB.CommandButton cmdsave 
         Caption         =   "&Save"
         Height          =   375
         Left            =   5880
         TabIndex        =   2
         Top             =   2880
         Width           =   795
      End
      Begin VB.CommandButton cmdcancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   6720
         TabIndex        =   1
         Top             =   2880
         Width           =   795
      End
      Begin VB.Label Label6 
         Caption         =   "Email"
         Height          =   495
         Left            =   240
         TabIndex        =   11
         Top             =   2820
         Width           =   915
      End
      Begin VB.Label Label5 
         Caption         =   "Fax"
         Height          =   315
         Left            =   240
         TabIndex        =   10
         Top             =   2400
         Width           =   915
      End
      Begin VB.Label Label4 
         Caption         =   "Phone"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1980
         Width           =   915
      End
   End
End
Attribute VB_Name = "frmprofile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo lbl
If rsprofile.State = 1 Then rsprofile.Close
    rsprofile.Open "select * from tbl_profile1", cn, adOpenStatic, adLockPessimistic
    id = rsprofile.Fields("pro_id")
If MsgBox("Already exists,do u want to modify", vbYesNo + vbQuestion, "warning") = vbYes Then
    If rsprofile.State = 1 Then rsprofile.Close
        cn.Execute "update tbl_profile1 set title1='" & txtTitle1.Text & "',title2='" & txtTitle2.Text & "',title3='" & txtTitle3.Text & "',phone='" & txtPhone.Text & "',fax='" & txtFax.Text & "',email='" & txtMail.Text & "' where pro_id=" & id
    End If
    Exit Sub
lbl:
    MsgBox Err.Description, vbInformation, "Message"

   
End Sub

Private Sub Form_Load()
On Error GoTo errdesc
    If rsprofile.State = 1 Then rsprofile.Close
        rsprofile.Open "select * from tbl_profile1 ", cn, adOpenStatic, adLockPessimistic
    If rsprofile.EOF = False Then
            txtTitle1.Text = rsprofile.Fields("title1")
            txtTitle2.Text = rsprofile.Fields("title2")
            txtTitle3.Text = rsprofile.Fields("title3")
            txtPhone.Text = rsprofile.Fields("phone")
            txtFax.Text = rsprofile.Fields("fax")
            txtMail.Text = rsprofile.Fields("Email")
    Else
        Clearing
    End If
    Exit Sub
errdesc:
    MsgBox Err.Description, vbInformation, "Message"
End Sub

Public Sub Clearing()
On Error GoTo errdesc
            txtTitle1.Text = ""
            txtTitle2.Text = ""
            txtTitle3.Text = ""
            txtPhone.Text = ""
            txtFax.Text = ""
            txtMail.Text = ""
Exit Sub
errdesc:
    MsgBox Err.Description, vbInformation, "Message"
End Sub
Private Sub txtFax_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtMail.SetFocus
End Sub

Private Sub txtMail_LostFocus()
 ValEmail (txtMail.Text)
End Sub

Private Sub txtPhone_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtFax.SetFocus
End Sub

Private Sub txtPhone_LostFocus()
    If Not ValPhone(txtPhone.Text) Then
        MsgBox "Not a valid contactno", vbInformation, "Message"
    End If
End Sub

Private Sub txtTitle1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtTitle2.SetFocus
End Sub
Private Sub txtTitle2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtTitle3.SetFocus
End Sub
Private Sub txtTitle3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtPhone.SetFocus
End Sub

