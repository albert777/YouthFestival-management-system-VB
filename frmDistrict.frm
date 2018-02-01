VERSION 5.00
Begin VB.Form frmDistrict 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "District"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7305
   Icon            =   "frmDistrict.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   7305
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2355
      Left            =   60
      TabIndex        =   2
      Top             =   0
      Width           =   7215
      Begin VB.ComboBox Cmbname 
         Height          =   1545
         Left            =   1920
         Style           =   1  'Simple Combo
         TabIndex        =   0
         Top             =   240
         Width           =   5115
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   6120
         TabIndex        =   4
         Top             =   1860
         Width           =   915
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   375
         Left            =   5160
         TabIndex        =   1
         Top             =   1860
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "District name"
         Height          =   495
         Left            =   180
         TabIndex        =   3
         Top             =   240
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmDistrict"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cmbname_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdSave.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo lbl
    If Cmbname.ListIndex <> -1 Then
        If rsDistrict.State = 1 Then rsDistrict.Close
        rsDistrict.Open "Select * from tbl_district where district_id=" & Cmbname.ItemData(Cmbname.ListIndex), cn, adOpenStatic, adLockPessimistic
        If rsDistrict.EOF = True Then
            cn.Execute "insert into tbl_district(district_name) values('" & Cmbname.Text & "')"
        Else
            If MsgBox("Already Exists,Do U want to modify", vbYesNo + vbQuestion, "Warning") = vbYes Then
                If rsDistrict.State = 1 Then rsDistrict.Close
                cn.Execute "update tbl_district set district_name='" & Cmbname.Text & "' where district_id=" & Cmbname.ItemData(Cmbname.ListIndex)
            End If
        End If
    Else
        cn.Execute "insert into tbl_district(district_name) values('" & Cmbname.Text & "')"
    End If
    AddDistrict
Exit Sub
lbl:
    MsgBox Err.Description, vbInformation, "Message"
End Sub


Public Sub AddDistrict()
On Error GoTo lbl
    If rsDistrict.State = 1 Then rsDistrict.Close
    rsDistrict.Open "Select * from tbl_district", cn, adOpenStatic, adLockPessimistic
    Cmbname.Clear
    While Not rsDistrict.EOF
        Cmbname.AddItem rsDistrict.Fields("district_name")
        Cmbname.ItemData(Cmbname.NewIndex) = rsDistrict.Fields("district_id")
        rsDistrict.MoveNext
    Wend
Exit Sub
lbl:
    MsgBox Err.Description, vbInformation, "Message"
    
End Sub

Private Sub Form_Load()
    AddDistrict
End Sub
