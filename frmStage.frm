VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmStage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stage"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11040
   Icon            =   "frmStage.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   11040
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbstageno 
      Height          =   1545
      Left            =   1380
      Sorted          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   0
      Top             =   240
      Width           =   3675
   End
   Begin VB.Frame Frame1 
      Height          =   5475
      Left            =   60
      TabIndex        =   6
      Top             =   0
      Width           =   10935
      Begin MSFlexGridLib.MSFlexGrid flxGrid 
         Height          =   2955
         Left            =   120
         TabIndex        =   13
         Top             =   2460
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   5212
         _Version        =   393216
         Cols            =   7
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   495
         Left            =   9600
         TabIndex        =   12
         Top             =   1740
         Width           =   1155
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   495
         Left            =   8400
         TabIndex        =   5
         Top             =   1740
         Width           =   1155
      End
      Begin VB.TextBox Txtcontactno 
         Height          =   435
         Left            =   7080
         MaxLength       =   10
         TabIndex        =   4
         Top             =   1200
         Width           =   3675
      End
      Begin VB.TextBox Txtcontactperson 
         Height          =   435
         Left            =   7080
         MaxLength       =   50
         TabIndex        =   3
         Top             =   720
         Width           =   3675
      End
      Begin VB.TextBox Txtlocation 
         Height          =   435
         Left            =   7080
         MaxLength       =   50
         TabIndex        =   2
         Top             =   240
         Width           =   3675
      End
      Begin VB.TextBox Txtvenue 
         Height          =   435
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   1
         Top             =   1860
         Width           =   3675
      End
      Begin VB.Label Label5 
         Caption         =   "Contact Number"
         Height          =   495
         Left            =   5160
         TabIndex        =   11
         Top             =   1260
         Width           =   1395
      End
      Begin VB.Label Label4 
         Caption         =   "Contact Person"
         Height          =   495
         Left            =   5160
         TabIndex        =   10
         Top             =   720
         Width           =   1395
      End
      Begin VB.Label Label3 
         Caption         =   "Location"
         Height          =   495
         Left            =   5160
         TabIndex        =   9
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label Label2 
         Caption         =   "Venue"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   1395
      End
      Begin VB.Label Label1 
         Caption         =   "Stage Number"
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1395
      End
   End
End
Attribute VB_Name = "frmStage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
            txtContactNo.Text = rsstage.Fields("contact_no")
            txtContactPerson.Text = rsstage.Fields("contact_person")
        Else
            Clearing
        End If
    Else
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
If cmbStageNo.ListIndex <> -1 Then
    If rsstage.State = 1 Then rsstage.Close
    rsstage.Open "select * from tbl_stage where stage_id=" & cmbStageNo.ItemData(cmbStageNo.ListIndex), cn, adOpenStatic, adLockPessimistic
    If rsstage.EOF = True Then
        If rsstage.State = 1 Then rsstage.Close
        rsstage.Open "select * from tbl_stage where stage_no=" & Val(cmbStageNo.Text), cn, adOpenStatic, adLockPessimistic
        If rsstage.EOF = True Then
            cn.Execute "insert into tbl_stage(stage_no,venue,location,contact_person,contact_no) values(" & cmbStageNo.Text & ",'" & txtVenue.Text & "','" & txtLocation.Text & "','" & txtContactPerson.Text & "','" & txtContactNo.Text & "')"
        Else
            MsgBox "Stageno already exists", vbInformation, "Message"
        End If
    Else
        If MsgBox("Already Exists,Do U want to modify", vbYesNo + vbQuestion, "Warning") = vbYes Then
            If rsstage.State = 1 Then rsstage.Close
            cn.Execute "update tbl_stage set stage_no=" & cmbStageNo.Text & ",venue='" & txtVenue.Text & "',location= '" & txtLocation.Text & "',contact_person='" & txtContactPerson.Text & "',contact_no='" & txtContactNo.Text & "' where stage_id= " & cmbStageNo.ItemData(cmbStageNo.ListIndex) & ""
        End If
    End If
Else
    If rsstage.State = 1 Then rsstage.Close
    rsstage.Open "select * from tbl_stage where stage_no=" & Val(cmbStageNo.Text), cn, adOpenStatic, adLockPessimistic
    If rsstage.EOF = True Then
        cn.Execute "insert into tbl_stage(stage_no,venue,location,contact_person,contact_no) values(" & cmbStageNo.Text & ",'" & txtVenue.Text & "','" & txtLocation.Text & "','" & txtContactPerson.Text & "','" & txtContactNo.Text & "')"
    Else
        MsgBox "Stageno already exists", vbInformation, "Message"
    End If
    
  End If
  
    addstageno
DispGrid
Clearing
cmbStageNo.SetFocus

Exit Sub
lbl:
    MsgBox Err.Description, vbInformation, "Message"

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

Private Sub flxGrid_DblClick()
On Error GoTo lbl
    Dim id As Integer, i As Integer
    id = flxGrid.TextMatrix(flxGrid.RowSel, 6)
    For i = 0 To cmbStageNo.ListCount - 1
        If cmbStageNo.ItemData(i) = id Then
            cmbStageNo.ListIndex = i
            Exit For
         End If
    Next
       Exit Sub
lbl:
    MsgBox Err.Description, vbInformation, "Message"
      
    
End Sub

Private Sub Form_Load()
    addstageno
    DispGrid

End Sub

Public Sub DispGrid()
    flxGrid.Clear
    flxGrid.Rows = 2
    flxGrid.ColWidth(0) = 1000
    flxGrid.ColWidth(1) = 1000
    flxGrid.ColWidth(2) = 2000
    flxGrid.ColWidth(3) = 2000
    flxGrid.ColWidth(4) = 2000
    flxGrid.ColWidth(5) = 2000
    flxGrid.ColWidth(6) = 0
    
    flxGrid.TextMatrix(0, 0) = "Sl No"
    flxGrid.TextMatrix(0, 1) = "Stage No"
    flxGrid.TextMatrix(0, 2) = "Venue"
    flxGrid.TextMatrix(0, 3) = "Location"
    flxGrid.TextMatrix(0, 4) = "Contact Person"
    flxGrid.TextMatrix(0, 5) = "Contact No"
    flxGrid.TextMatrix(0, 6) = "Id"
    
    If rsstage.State = 1 Then rsstage.Close
    rsstage.Open "select * from tbl_stage", cn, adOpenStatic, adLockPessimistic
    While Not rsstage.EOF
        flxGrid.TextMatrix(flxGrid.Rows - 1, 0) = flxGrid.Rows - 1
        flxGrid.TextMatrix(flxGrid.Rows - 1, 1) = rsstage.Fields("stage_no")
        flxGrid.TextMatrix(flxGrid.Rows - 1, 2) = rsstage.Fields("venue")
        flxGrid.TextMatrix(flxGrid.Rows - 1, 3) = rsstage.Fields("location")
        flxGrid.TextMatrix(flxGrid.Rows - 1, 4) = rsstage.Fields("contact_person")
        flxGrid.TextMatrix(flxGrid.Rows - 1, 5) = rsstage.Fields("contact_no")
        flxGrid.TextMatrix(flxGrid.Rows - 1, 6) = rsstage.Fields("stage_id")
        rsstage.MoveNext
        flxGrid.Rows = flxGrid.Rows + 1
    Wend
    
    
End Sub

Public Sub Clearing()
    txtVenue.Text = ""
    txtLocation.Text = ""
    txtContactNo.Text = ""
    txtContactPerson.Text = ""
End Sub







Private Sub txtContactNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdsave.SetFocus
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

Private Sub txtVenue_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtLocation.SetFocus
End Sub
