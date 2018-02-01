VERSION 5.00
Begin VB.Form frmSchoolwise 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Schoolwise"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6165
   Icon            =   "frmSchoolwise.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   6165
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   4455
      Left            =   60
      TabIndex        =   2
      Top             =   0
      Width           =   6075
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   495
         Left            =   4920
         TabIndex        =   6
         Top             =   3840
         Width           =   975
      End
      Begin VB.CommandButton cmdShow 
         Caption         =   "&Show"
         Height          =   495
         Left            =   3900
         TabIndex        =   5
         Top             =   3840
         Width           =   975
      End
      Begin VB.ComboBox cmbSchool 
         Height          =   1740
         Left            =   1260
         Style           =   1  'Simple Combo
         TabIndex        =   1
         Top             =   2040
         Width           =   4695
      End
      Begin VB.ComboBox cmbDistrict 
         Height          =   1740
         Left            =   1260
         Style           =   1  'Simple Combo
         TabIndex        =   0
         Top             =   240
         Width           =   4695
      End
      Begin VB.Label Label2 
         Caption         =   "School Name"
         Height          =   435
         Left            =   60
         TabIndex        =   4
         Top             =   1980
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "District"
         Height          =   255
         Left            =   60
         TabIndex        =   3
         Top             =   300
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmSchoolwise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub AddDistrict()
    If rsDistrict.State = 1 Then rsDistrict.Close
    rsDistrict.Open "select * from tbl_district", cn, adOpenStatic, adLockPessimistic
    cmbDistrict.Clear
    While Not rsDistrict.EOF
        cmbDistrict.AddItem rsDistrict.Fields("district_name")
        cmbDistrict.ItemData(cmbDistrict.NewIndex) = rsDistrict.Fields("district_id")
        rsDistrict.MoveNext
    Wend
    
End Sub

Private Sub cmbDistrict_Change()
    cmbDistrict_Click
End Sub

Private Sub cmbDistrict_Click()
On Error GoTo lbl
     If cmbDistrict.ListIndex <> -1 Then
        If rsSchool.State = 1 Then rsSchool.Close
        rsSchool.Open "select * from tbl_school where district_id=" & cmbDistrict.ItemData(cmbDistrict.ListIndex), cn, adOpenStatic, adLockPessimistic
        cmbSchool.Clear
        While Not rsSchool.EOF
            cmbSchool.AddItem rsSchool.Fields("school_name")
            cmbSchool.ItemData(cmbSchool.NewIndex) = rsSchool.Fields("school_id")
            rsSchool.MoveNext
        Wend
    Else
        cmbSchool.Clear
    End If
Exit Sub
lbl:
    MsgBox Err.Description, vbInformation, "Message"
End Sub

Private Sub Cmbdistrict_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then cmbSchool.SetFocus
End Sub


Private Sub cmbSchool_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then cmdShow.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdShow_Click()
       rptMemberSchoolwise.Show
End Sub

Private Sub Form_Load()
    AddDistrict
End Sub
