VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmprogram 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Program"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11460
   Icon            =   "frmprogram.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   11460
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   6315
      Left            =   60
      TabIndex        =   4
      Top             =   0
      Width           =   11355
      Begin MSFlexGridLib.MSFlexGrid flxGrid 
         Height          =   3195
         Left            =   120
         TabIndex        =   9
         Top             =   3000
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   5636
         _Version        =   393216
         Cols            =   5
      End
      Begin VB.TextBox txtMinutes 
         Height          =   435
         Left            =   9300
         MaxLength       =   4
         TabIndex        =   2
         Top             =   180
         Width           =   1935
      End
      Begin VB.TextBox txtPDiscription 
         Height          =   1185
         Left            =   1740
         MaxLength       =   50
         TabIndex        =   1
         Top             =   1680
         Width           =   6135
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   435
         Left            =   8700
         TabIndex        =   3
         Top             =   2340
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   435
         Left            =   10020
         TabIndex        =   5
         Top             =   2340
         Width           =   1215
      End
      Begin VB.ComboBox cmbPName 
         Height          =   1350
         Left            =   1740
         Style           =   1  'Simple Combo
         TabIndex        =   0
         Top             =   240
         Width           =   3975
      End
      Begin VB.Label Label3 
         Caption         =   "Time in Minutes"
         Height          =   435
         Left            =   8040
         TabIndex        =   8
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Program Name"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1050
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Program Description"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   1740
         Width           =   1425
      End
   End
End
Attribute VB_Name = "frmprogram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbPName_Change()
    cmbPName_Click
End Sub

Private Sub cmbPName_Click()
    If cmbPName.ListIndex <> -1 Then
        If rsProgram.State = 1 Then rsProgram.Close
        rsProgram.Open "select * from tbl_program where program_id=" & cmbPName.ItemData(cmbPName.ListIndex), cn, adOpenStatic, adLockPessimistic
        If rsProgram.EOF = False Then
            txtPDiscription.Text = rsProgram.Fields("Program_description")
            txtMinutes.Text = rsProgram.Fields("program_time")
        Else
            Clearing
        End If
    Else
        Clearing
    End If
End Sub

Private Sub cmbPName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtPDiscription.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo lbl
    If Len(Trim(cmbPName.Text)) <> 0 And Val(txtMinutes.Text) <> 0 Then
        If cmbPName.ListIndex <> -1 Then
            If rsProgram.State = 1 Then rsProgram.Close
            rsProgram.Open "select * from tbl_program where program_id=" & cmbPName.ItemData(cmbPName.ListIndex), cn, adOpenStatic, adLockPessimistic
            If rsProgram.EOF = True Then
                cn.Execute "insert into tbl_program(program_name,program_description,program_time)values('" & _
                cmbPName.Text & "','" & txtPDiscription.Text & "'," & Val(txtMinutes.Text) & ")"
            Else
                If MsgBox("Program already added,Do u want to modify", vbYesNo + vbQuestion, "Warning") = vbYes Then
                    If rsProgram.State = 1 Then rsProgram.Close
                    cn.Execute "update tbl_program set program_description='" & txtPDiscription.Text & _
                    "',program_time=" & Val(txtMinutes.Text) & " where program_id=" & cmbPName.ItemData(cmbPName.ListIndex)
                End If
            End If
        Else
            cn.Execute "insert into tbl_program(program_name,program_description,program_time)values('" & _
            cmbPName.Text & "','" & txtPDiscription.Text & "'," & Val(txtMinutes.Text) & ")"
        End If
        Clearing
        cmbPName.Text = ""
        cmbPName.ListIndex = -1
        AddProgram
        DispGrid
    Else
        MsgBox "Enter programname and minutes", vbInformation, "Message"
    End If
       Exit Sub
lbl:
    MsgBox Err.Description, vbInformation, "Message"
    
End Sub

Public Sub Clearing()
    txtPDiscription.Text = ""
    txtMinutes.Text = ""
End Sub

Public Sub AddProgram()
    If rsProgram.State = 1 Then rsProgram.Close
    rsProgram.Open "select * from tbl_program", cn, adOpenStatic, adLockPessimistic
    cmbPName.Clear
    While Not rsProgram.EOF
        cmbPName.AddItem rsProgram.Fields("program_name")
        cmbPName.ItemData(cmbPName.NewIndex) = rsProgram.Fields("program_id")
        rsProgram.MoveNext
    Wend
End Sub

Private Sub flxGrid_DblClick()
On Error GoTo lbl
    Dim id As Integer, i As Integer
    id = flxGrid.TextMatrix(flxGrid.RowSel, 4)
    For i = 0 To cmbPName.ListCount - 1
        If cmbPName.ItemData(i) = id Then
            cmbPName.ListIndex = i
            Exit For
        End If
    Next
    Exit Sub
lbl:
    MsgBox Err.Description, vbInformation, "Message"
        
End Sub

Private Sub Form_Load()
    AddProgram
    DispGrid
End Sub

Public Sub DispGrid()
    flxGrid.Clear
    flxGrid.Rows = 2
    flxGrid.ColWidth(0) = 1000
    flxGrid.ColWidth(1) = 3500
    flxGrid.ColWidth(2) = 4500
    flxGrid.ColWidth(3) = 1500
    flxGrid.ColWidth(4) = 0
    
    flxGrid.TextMatrix(0, 0) = "Sl No"
    flxGrid.TextMatrix(0, 1) = "Program Name"
    flxGrid.TextMatrix(0, 2) = "Program Description"
    flxGrid.TextMatrix(0, 3) = "Time in Minutes"
    flxGrid.TextMatrix(0, 4) = "Id"
    
    If rsProgram.State = 1 Then rsProgram.Close
    rsProgram.Open "select * from tbl_program", cn, adOpenStatic, adLockPessimistic
    While Not rsProgram.EOF
        flxGrid.TextMatrix(flxGrid.Rows - 1, 0) = flxGrid.Rows - 1
        flxGrid.TextMatrix(flxGrid.Rows - 1, 1) = rsProgram.Fields("program_name")
        flxGrid.TextMatrix(flxGrid.Rows - 1, 2) = rsProgram.Fields("program_description")
        flxGrid.TextMatrix(flxGrid.Rows - 1, 3) = rsProgram.Fields("program_time")
        flxGrid.TextMatrix(flxGrid.Rows - 1, 4) = rsProgram.Fields("program_id")
        rsProgram.MoveNext
        flxGrid.Rows = flxGrid.Rows + 1
    Wend
End Sub





Private Sub txtMinutes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdsave.SetFocus
    Else
        NumCheck KeyAscii
    End If
End Sub

Private Sub txtPDiscription_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtMinutes.SetFocus
End Sub
