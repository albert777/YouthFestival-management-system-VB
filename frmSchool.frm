VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmSchool 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "School"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11565
   Icon            =   "frmSchool.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   11565
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   6855
      Left            =   60
      TabIndex        =   7
      Top             =   0
      Width           =   11475
      Begin MSFlexGridLib.MSFlexGrid flxGrid 
         Height          =   3495
         Left            =   60
         TabIndex        =   15
         Top             =   3240
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   6165
         _Version        =   393216
         Cols            =   7
      End
      Begin VB.ComboBox Cmbdistrict 
         Height          =   1350
         Left            =   1140
         Style           =   1  'Simple Combo
         TabIndex        =   0
         Top             =   180
         Width           =   4155
      End
      Begin VB.ComboBox Cmbschoolname 
         Height          =   1350
         Left            =   1140
         Style           =   1  'Simple Combo
         TabIndex        =   1
         Top             =   1620
         Width           =   4155
      End
      Begin VB.TextBox Txtaddress 
         Height          =   1230
         Left            =   7080
         MaxLength       =   50
         TabIndex        =   2
         Top             =   180
         Width           =   4155
      End
      Begin VB.TextBox Txtcontactno1 
         Height          =   375
         Left            =   7080
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1440
         Width           =   4155
      End
      Begin VB.TextBox Txtcontactno2 
         Height          =   375
         Left            =   7080
         MaxLength       =   50
         TabIndex        =   4
         Top             =   1860
         Width           =   4155
      End
      Begin VB.TextBox Txtemail 
         Height          =   375
         Left            =   7080
         MaxLength       =   50
         TabIndex        =   5
         Top             =   2280
         Width           =   4155
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   435
         Left            =   10200
         TabIndex        =   8
         Top             =   2700
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Cancel          =   -1  'True
         Caption         =   "&Save"
         Height          =   435
         Left            =   9060
         TabIndex        =   6
         Top             =   2700
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Contact Number2"
         Height          =   435
         Left            =   5520
         TabIndex        =   14
         Top             =   1860
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Conatact Number1"
         Height          =   315
         Left            =   5520
         TabIndex        =   13
         Top             =   1500
         Width           =   1635
      End
      Begin VB.Label Label2 
         Caption         =   "School Name"
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   1560
         Width           =   1515
      End
      Begin VB.Label Label1 
         Caption         =   "District"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "Address"
         Height          =   315
         Left            =   5520
         TabIndex        =   10
         Top             =   180
         Width           =   1515
      End
      Begin VB.Label Label6 
         Caption         =   "Email"
         Height          =   315
         Left            =   5520
         TabIndex        =   9
         Top             =   2280
         Width           =   1635
      End
   End
End
Attribute VB_Name = "frmSchool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmbDistrict_Click()
On Error GoTo lbl
If cmbDistrict.ListIndex <> -1 Then
    addschool
    DispGrid
    If rsSchool.State = 1 Then rsSchool.Close
    rsSchool.Open "select * from tbl_school where district_id=" & cmbDistrict.ItemData(cmbDistrict.ListIndex), cn, adOpenStatic, adLockPessimistic
    While Not rsSchool.EOF
        flxGrid.TextMatrix(flxGrid.Rows - 1, 0) = flxGrid.Rows - 1
        flxGrid.TextMatrix(flxGrid.Rows - 1, 1) = rsSchool.Fields("School_name")
        flxGrid.TextMatrix(flxGrid.Rows - 1, 2) = rsSchool.Fields("address")
        flxGrid.TextMatrix(flxGrid.Rows - 1, 3) = rsSchool.Fields("contact_no1")
        flxGrid.TextMatrix(flxGrid.Rows - 1, 4) = rsSchool.Fields("contact_no2")
        flxGrid.TextMatrix(flxGrid.Rows - 1, 5) = rsSchool.Fields("email")
        flxGrid.TextMatrix(flxGrid.Rows - 1, 6) = rsSchool.Fields("School_id")
        rsSchool.MoveNext
        flxGrid.Rows = flxGrid.Rows + 1
    Wend
End If
Exit Sub
lbl:
    MsgBox Err.Description, vbInformation, "Message"
End Sub

Private Sub Cmbdistrict_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Cmbschoolname.SetFocus
End Sub

Private Sub Cmbschoolname_Click()
On Error GoTo lbl
    If rsSchool.State = 1 Then rsSchool.Close
    If rsSchool.State = 1 Then rsschool_name.Close
    rsSchool.Open "Select * from tbl_school where school_id=" & Cmbschoolname.ItemData(Cmbschoolname.ListIndex), cn, adOpenStatic, adLockPessimistic
    If rsSchool.EOF = False Then
        txtAddress.Text = rsSchool.Fields("address")
        Txtcontactno1.Text = rsSchool.Fields("contact_no1")
        Txtcontactno2.Text = rsSchool.Fields("contact_no2")
        txtEmail.Text = rsSchool.Fields("email")
    Else
       Clearing
    End If
        Exit Sub
lbl:
    MsgBox Err.Description, vbInformation, "Message"
    
        
End Sub

Private Sub Cmbschoolname_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtAddress.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    
End Sub

Private Sub cmdSave_Click()
On Error GoTo lbl
 If Cmbschoolname.ListIndex <> -1 Then
        If rsSchool.State = 1 Then rsSchool.Close
        rsSchool.Open "Select * from tbl_school where school_id=" & Cmbschoolname.ItemData(Cmbschoolname.ListIndex), cn, adOpenStatic, adLockPessimistic
        If rsSchool.EOF = True Then
             cn.Execute "insert into tbl_school(district_id,school_name,address,contact_no1,contact_no2,email)values(" & _
        cmbDistrict.ItemData(cmbDistrict.ListIndex) & ",'" & Cmbschoolname.Text & "','" & txtAddress.Text & "','" & Txtcontactno1.Text & "','" & Txtcontactno2.Text & "','" & txtEmail.Text & "') "
        Else
            If MsgBox("Already Exists,Do U want to modify", vbYesNo + vbQuestion, "Warning") = vbYes Then
            If rsSchool.State = 1 Then rsSchool.Close
                cn.Execute "update tbl_school set district_id= " & cmbDistrict.ItemData(cmbDistrict.ListIndex) & ",school_name='" & Cmbschoolname.Text & "',address='" & txtAddress.Text & "',contact_no1=" & Txtcontactno1.Text & ",contact_no2=" & Txtcontactno2.Text & ",email='" & txtEmail.Text & "' where school_id=" & Cmbschoolname.ItemData(Cmbschoolname.ListIndex) & " "
            End If
        End If
    Else
        cn.Execute "insert into tbl_school(district_id,school_name,address,contact_no1,contact_no2,email)values(" & _
        cmbDistrict.ItemData(cmbDistrict.ListIndex) & ",'" & Cmbschoolname.Text & "','" & txtAddress.Text & "','" & Txtcontactno1.Text & "','" & Txtcontactno2.Text & "','" & txtEmail.Text & "') "
    End If
    Clearing
    Cmbschoolname.Clear
    Cmbschoolname.ListIndex = -1
    cmbDistrict.Text = ""
    cmbDistrict.ListIndex = -1
    DispGrid
    Exit Sub
lbl:
    MsgBox Err.Description, vbInformation, "Message"
    
    
End Sub

Private Sub flxGrid_DblClick()
On Error GoTo lbl
     Dim id As Integer, i As Integer
    id = flxGrid.TextMatrix(flxGrid.RowSel, 6)
    For i = 0 To Cmbschoolname.ListCount - 1
        If Cmbschoolname.ItemData(i) = id Then
            Cmbschoolname.ListIndex = i
            Exit For
         End If
    Next
            Exit Sub
lbl:
    MsgBox Err.Description, vbInformation, "Message"
       
End Sub

Private Sub Form_Load()
    AddDistrict
    DispGrid
End Sub
Public Sub AddDistrict()
    If rsDistrict.State = 1 Then rsDistrict.Close
    rsDistrict.Open "Select * from tbl_district ", cn, adOpenStatic, adLockPessimistic
    cmbDistrict.Clear
    While Not rsDistrict.EOF
        cmbDistrict.AddItem rsDistrict.Fields("district_name")
        cmbDistrict.ItemData(cmbDistrict.NewIndex) = rsDistrict.Fields("district_id")
        rsDistrict.MoveNext
    Wend
    
End Sub


Public Sub addschool()
If rsSchool.State = 1 Then rsSchool.Close
    rsSchool.Open "Select * from tbl_school where district_id=" & cmbDistrict.ItemData(cmbDistrict.ListIndex), cn, adOpenStatic, adLockPessimistic
    Cmbschoolname.Clear
    While Not rsSchool.EOF
        Cmbschoolname.AddItem rsSchool.Fields("school_name")
        Cmbschoolname.ItemData(Cmbschoolname.NewIndex) = rsSchool.Fields("School_id")
        rsSchool.MoveNext
    Wend
End Sub

Public Sub DispGrid()
    flxGrid.Clear
    flxGrid.Rows = 2
    flxGrid.ColWidth(0) = 1000
    flxGrid.ColWidth(1) = 2500
    flxGrid.ColWidth(2) = 2500
    flxGrid.ColWidth(3) = 1000
    flxGrid.ColWidth(4) = 1000
    flxGrid.ColWidth(5) = 2000
    flxGrid.ColWidth(6) = 0
    
    flxGrid.TextMatrix(0, 0) = "Sl No"
    flxGrid.TextMatrix(0, 1) = "Name"
    flxGrid.TextMatrix(0, 2) = "Address"
    flxGrid.TextMatrix(0, 3) = "Contact No1"
    flxGrid.TextMatrix(0, 4) = "Contact No2"
    flxGrid.TextMatrix(0, 5) = "Email"
    flxGrid.TextMatrix(0, 6) = "Id"
    
End Sub


Private Sub txtAddress_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Txtcontactno1.SetFocus
End Sub




Private Sub Txtcontactno1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Txtcontactno2.SetFocus
End Sub



Private Sub Txtcontactno1_LostFocus()
 If Not ValPhone(Txtcontactno1.Text) Then
        MsgBox "Not a valid contactno", vbInformation, "Message"
    End If
End Sub

Private Sub Txtcontactno2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtEmail.SetFocus
End Sub


Private Sub Txtcontactno2_LostFocus()
 If Not ValPhone(Txtcontactno2.Text) Then
        MsgBox "Not a valid contactno", vbInformation, "Message"
    End If
End Sub

Private Sub Txtemail_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdsave.SetFocus
End Sub

Private Sub txtEmail_LostFocus()
ValEmail (txtEmail.Text)
End Sub

Public Sub Clearing()
      txtAddress.Text = ""
        Txtcontactno1.Text = ""
        Txtcontactno2.Text = ""
        txtEmail.Text = ""
End Sub
