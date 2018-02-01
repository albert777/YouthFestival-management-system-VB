VERSION 5.00
Begin VB.Form frmgreen_room 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Green Room"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10890
   Icon            =   "frmgreen_room.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   10890
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   60
      TabIndex        =   6
      Top             =   0
      Width           =   10815
      Begin VB.TextBox txtGRoom 
         Height          =   405
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   14
         Top             =   2520
         Width           =   2295
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   375
         Left            =   8160
         TabIndex        =   5
         Top             =   2520
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   9420
         TabIndex        =   13
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         Caption         =   "Stage Details"
         Height          =   2295
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   10575
         Begin VB.TextBox txtContactNo 
            Height          =   375
            Left            =   6840
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   1680
            Width           =   3615
         End
         Begin VB.TextBox txtContactPerson 
            Height          =   375
            Left            =   6840
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   1200
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
         Begin VB.TextBox txtVenue 
            Height          =   375
            Left            =   6840
            Locked          =   -1  'True
            TabIndex        =   1
            Top             =   240
            Width           =   3615
         End
         Begin VB.ComboBox cmbStage 
            Height          =   1740
            Left            =   840
            Style           =   1  'Simple Combo
            TabIndex        =   0
            Top             =   240
            Width           =   4695
         End
         Begin VB.Label Label8 
            Caption         =   "Contact No"
            Height          =   255
            Left            =   5640
            TabIndex        =   12
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label Label7 
            Caption         =   "Contact Person"
            Height          =   255
            Left            =   5640
            TabIndex        =   11
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label Label6 
            Caption         =   "Location"
            Height          =   375
            Left            =   5640
            TabIndex        =   10
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "Venue"
            Height          =   255
            Left            =   5640
            TabIndex        =   9
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Stage"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Green Room No."
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   2640
         Width           =   1200
      End
   End
End
Attribute VB_Name = "frmgreen_room"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub addstageno()
    If rsstage.State = 1 Then rsstage.Close
    rsstage.Open "Select * from tbl_stage ", cn, adOpenStatic, adLockPessimistic
    cmbStage.Clear
    While Not rsstage.EOF
        cmbStage.AddItem rsstage.Fields("stage_no")
        cmbStage.ItemData(cmbStage.NewIndex) = rsstage.Fields("stage_id")
        rsstage.MoveNext
    Wend

End Sub

Private Sub cmbStage_Click()
On Error GoTo lbl
    If cmbStage.ListIndex <> -1 Then
        If rsstage.State = 1 Then rsstage.Close
        rsstage.Open "select * from tbl_stage where stage_id=" & cmbStage.ItemData(cmbStage.ListIndex), cn, adOpenStatic, adLockPessimistic
        If rsstage.EOF = False Then
            txtVenue.Text = rsstage.Fields("venue")
            txtLocation.Text = rsstage.Fields("location")
            txtContactNo.Text = rsstage.Fields("contact_no")
            txtContactPerson.Text = rsstage.Fields("contact_person")
        Else
            Clearing
        End If
    Else
        Clearing
    End If
    Exit Sub
lbl:
    MsgBox Err.Description, vbInformation, "Message"
    
End Sub

Private Sub cmbStage_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyDown = 13 Then txtVenue.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo lbl
    If cmbStage.ListIndex <> -1 And Val(txtGRoom.Text) <> 0 Then
        If rsGreenRoom.State = 1 Then rsGreenRoom.Close
        rsGreenRoom.Open "select * from tbl_GreenRoom where stage_id=" & cmbStage.ItemData(cmbStage.ListIndex) & _
        " and green_room_no=" & Val(txtGRoom.Text), cn, adOpenStatic, adLockPessimistic
        If rsGreenRoom.EOF = True Then
            cn.Execute "insert into tbl_greenroom(green_room_no,stage_id)values(" & Val(txtGRoom.Text) & "," & cmbStage.ItemData(cmbStage.ListIndex) & ")"
            txtGRoom.Text = ""
            txtGRoom.SetFocus
        Else
            MsgBox "Already added", vbInformation, "Message"
        End If
    Else
        MsgBox "Select Stage and enter Green Room Number", vbInformation, "Message"
    End If
   Exit Sub
lbl:
    MsgBox Err.Description, vbInformation, "Message"
   
End Sub

Private Sub Form_Load()
    addstageno
End Sub

Public Sub Clearing()
    txtVenue.Text = ""
    txtLocation.Text = ""
    txtContactNo.Text = ""
    txtContactPerson.Text = ""
End Sub




Private Sub txtContactNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtGRoom.SetFocus
End Sub

Private Sub Txtcontactno_LostFocus()
 If Not ValPhone(txtContactNo.Text) Then
        MsgBox "Not a valid contactno", vbInformation, "Message"
    End If
End Sub

Private Sub txtContactPerson_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then cmdcontactno.SetFocus
End Sub


Private Sub txtGRoom_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdsave.SetFocus
    Else
        NumCheck KeyAscii
    End If
End Sub

Private Sub txtLocation_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtContactPerson.SetFocus
End Sub



Private Sub txtVenue_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtLocation.SetFocus
End Sub
