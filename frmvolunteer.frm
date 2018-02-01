VERSION 5.00
Begin VB.Form frmvolunteer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Volanteer"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10605
   Icon            =   "frmvolunteer.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   10605
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Stage Details"
      Height          =   2295
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   10575
      Begin VB.ComboBox cmbStageNo 
         Height          =   1740
         Left            =   840
         Style           =   1  'Simple Combo
         TabIndex        =   0
         Top             =   240
         Width           =   4695
      End
      Begin VB.TextBox txtVenue 
         Height          =   375
         Left            =   6840
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   1
         Top             =   240
         Width           =   3615
      End
      Begin VB.TextBox txtLocation 
         Height          =   375
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   720
         Width           =   3615
      End
      Begin VB.TextBox txtContactPerson 
         Height          =   375
         Left            =   6840
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1200
         Width           =   3615
      End
      Begin VB.TextBox txtStageContact 
         Height          =   375
         Left            =   6840
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   4
         Top             =   1680
         Width           =   3615
      End
      Begin VB.Label Label1 
         Caption         =   "Stage"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Venue"
         Height          =   255
         Left            =   5640
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Location"
         Height          =   375
         Left            =   5640
         TabIndex        =   17
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Contact Person"
         Height          =   255
         Left            =   5640
         TabIndex        =   16
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Contact No"
         Height          =   255
         Left            =   5640
         TabIndex        =   15
         Top             =   1680
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Volunteer"
      Height          =   1815
      Left            =   0
      TabIndex        =   9
      Top             =   2400
      Width           =   10575
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   9240
         TabIndex        =   10
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   375
         Left            =   7920
         TabIndex        =   8
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtContactNo 
         Height          =   405
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   6
         Top             =   720
         Width           =   3735
      End
      Begin VB.TextBox txtAddress 
         Height          =   975
         Left            =   6480
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   3975
      End
      Begin VB.TextBox txtVName 
         Height          =   405
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   5
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label Label4 
         Caption         =   "Contact No."
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Address"
         Height          =   375
         Left            =   5640
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Volunteer Name"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmvolunteer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text4_Change()

End Sub
Public Sub addstageno()
    If rsstage.State = 1 Then rsstage.Close
    rsstage.Open "Select * from tbl_stage ", cn, adOpenStatic, adLockPessimistic
    cmbStageNo.Clear
    While Not rsstage.EOF
        cmbStageNo.AddItem rsstage.Fields("stage_no")
        cmbStageNo.ItemData(cmbStageNo.NewIndex) = rsstage.Fields("stage_id")
        rsstage.MoveNext
    Wend

End Sub

Private Sub cmbstageno_Change()
    cmbStageNo_Click
End Sub

Private Sub cmbStageNo_Click()
     If cmbStageNo.ListIndex <> -1 Then
        If rsstage.State = 1 Then rsstage.Close
        rsstage.Open "select * from tbl_stage where stage_id=" & cmbStageNo.ItemData(cmbStageNo.ListIndex), cn, adOpenStatic, adLockPessimistic
        If rsstage.EOF = False Then
            txtVenue.Text = rsstage.Fields("venue")
            txtLocation.Text = rsstage.Fields("location")
            txtStageContact.Text = rsstage.Fields("contact_no")
            txtContactPerson.Text = rsstage.Fields("contact_person")
        Else
            Clearing1
            Clearing
        End If
        If rsVolunteer.State = 1 Then rsVolunteer.Close
        rsVolunteer.Open "select * from tbl_volanteer where stage_id=" & cmbStageNo.ItemData(cmbStageNo.ListIndex), cn, adOpenStatic, adLockPessimistic
        If rsVolunteer.EOF = False Then
            txtAddress.Text = rsVolunteer.Fields("address")
            txtContactNo.Text = rsVolunteer.Fields("contact_no")
            txtVName.Text = rsVolunteer.Fields("volanteer_name")
        Else
            Clearing
        End If
    Else
        Clearing1
        Clearing
    End If
End Sub

Private Sub cmbStageNo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtVenue.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo lbl
    If cmbStageNo.ListIndex <> -1 And Len(Trim(txtVName.Text)) <> 0 Then
        If rsVolunteer.State = 1 Then rsVolunteer.Close
        rsVolunteer.Open "select * from tbl_volanteer where stage_id=" & cmbStageNo.ItemData(cmbStageNo.ListIndex), cn, adOpenStatic, adLockPessimistic
        If rsVolunteer.EOF = True Then
            cn.Execute "insert into tbl_volanteer(stage_id,volanteer_name,address,contact_no) values(" & _
            cmbStageNo.ItemData(cmbStageNo.ListIndex) & ",'" & txtVName.Text & "','" & txtAddress.Text & "','" & txtContactNo.Text & "')"
        Else
            If MsgBox("Already volanteer is assigned to this stage,Do u want to modify", vbYesNo + vbQuestion, "Warning") = vbYes Then
                If rsVolunteer.State = 1 Then rsVolunteer.Close
                cn.Execute "update tbl_volanteer set volanteer_name='" & txtVName.Text & "',address='" & txtAddress.Text & _
                "',contact_no='" & txtContactNo.Text & "' where stage_id=" & cmbStageNo.ItemData(cmbStageNo.ListIndex)
            End If
        End If
        Clearing
        Clearing1
    Else
        MsgBox "Select Stage and enter Volunteer", vbInformation, "Message"
      End If
        Exit Sub
lbl:
    MsgBox Err.Description, vbInformation, "Message"
        
  
End Sub

Private Sub Form_Load()
    addstageno
End Sub

Public Sub Clearing1()
    txtVenue.Text = ""
    txtLocation.Text = ""
    txtStageContact.Text = ""
    txtContactPerson.Text = ""
End Sub


Public Sub Clearing()
    txtVName.Text = ""
    txtContactNo.Text = ""
    txtAddress.Text = ""
End Sub







Private Sub txtAddress_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtContactNo.SetFocus
End Sub



Private Sub txtContactNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdSave.SetFocus
End Sub

Private Sub Txtcontactno_LostFocus()
 If Not ValPhone(txtContactNo.Text) Then
        MsgBox "Not a valid contactno", vbInformation, "Message"
    End If
End Sub

Private Sub txtContactPerson_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtContactNo.SetFocus
End Sub

Private Sub txtLocation_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtContactPerson.SetFocus
End Sub



Private Sub txtStageContact_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtVName.SetFocus
End Sub

Private Sub txtStageContact_LostFocus()
 If Not ValPhone(ttxtStageContact.Text) Then
        MsgBox "Not a valid contactno", vbInformation, "Message"
    End If
End Sub

Private Sub txtVenue_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtLocation.SetFocus
End Sub



Private Sub txtVName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtAddress.SetFocus
End Sub
