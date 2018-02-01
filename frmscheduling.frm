VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmscheduling 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scheduling"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12750
   Icon            =   "frmscheduling.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   12750
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   7275
      Left            =   60
      TabIndex        =   16
      Top             =   0
      Width           =   12675
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   375
         Left            =   9840
         TabIndex        =   12
         Top             =   6660
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   11280
         TabIndex        =   26
         Top             =   6660
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         Caption         =   "Program Details"
         Height          =   3975
         Left            =   120
         TabIndex        =   22
         Top             =   120
         Width           =   12435
         Begin MSFlexGridLib.MSFlexGrid flxGrid 
            Height          =   1815
            Left            =   7260
            TabIndex        =   32
            Top             =   2040
            Width           =   5115
            _ExtentX        =   9022
            _ExtentY        =   3201
            _Version        =   393216
            Cols            =   3
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete"
            Height          =   375
            Left            =   11340
            TabIndex        =   13
            Top             =   1680
            Width           =   975
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add"
            Height          =   375
            Left            =   10200
            TabIndex        =   3
            Top             =   1680
            Width           =   975
         End
         Begin VB.ComboBox cmbProgram 
            Height          =   1350
            Left            =   1680
            Style           =   1  'Simple Combo
            TabIndex        =   0
            Top             =   240
            Width           =   4935
         End
         Begin VB.TextBox txtPgmDescription 
            Height          =   435
            Left            =   1680
            MaxLength       =   50
            TabIndex        =   1
            Top             =   1680
            Width           =   4935
         End
         Begin VB.ComboBox cmbJudges 
            Height          =   1350
            Left            =   7320
            Style           =   1  'Simple Combo
            TabIndex        =   2
            Top             =   240
            Width           =   4935
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Program"
            Height          =   195
            Left            =   60
            TabIndex        =   25
            Top             =   240
            Width           =   585
         End
         Begin VB.Label Label7 
            Caption         =   "Program Description"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   1800
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "Judges"
            Height          =   255
            Left            =   6720
            TabIndex        =   23
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Stage Details"
         Height          =   2295
         Left            =   120
         TabIndex        =   8
         Top             =   4200
         Width           =   12435
         Begin VB.TextBox txtStageContact 
            Height          =   375
            Left            =   7680
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   11
            Top             =   1680
            Width           =   4515
         End
         Begin VB.TextBox txtContactPerson 
            Height          =   375
            Left            =   7680
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   10
            Top             =   1200
            Width           =   4515
         End
         Begin VB.TextBox txtLocation 
            Height          =   375
            Left            =   7680
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   9
            Top             =   720
            Width           =   4515
         End
         Begin VB.TextBox txtVenue 
            Height          =   375
            Left            =   7680
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   14
            Top             =   240
            Width           =   4515
         End
         Begin VB.ComboBox cmbStage 
            Height          =   1740
            Left            =   900
            Style           =   1  'Simple Combo
            TabIndex        =   4
            Top             =   240
            Width           =   4935
         End
         Begin VB.Label Label8 
            Caption         =   "Contact No"
            Height          =   255
            Left            =   6120
            TabIndex        =   21
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label Label5 
            Caption         =   "Contact Person"
            Height          =   255
            Left            =   6120
            TabIndex        =   20
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label Label9 
            Caption         =   "Location"
            Height          =   375
            Left            =   6120
            TabIndex        =   19
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label10 
            Caption         =   "Venue"
            Height          =   255
            Left            =   6120
            TabIndex        =   18
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label11 
            Caption         =   "Stage"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.TextBox txtTimeTo 
         Height          =   375
         Left            =   5640
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   7
         Top             =   6600
         Width           =   975
      End
      Begin VB.TextBox txtTimeFrom 
         Height          =   375
         Left            =   3960
         MaxLength       =   50
         TabIndex        =   6
         Top             =   6600
         Width           =   975
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1020
         TabIndex        =   5
         Top             =   6540
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   76808193
         CurrentDate     =   42142
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Date"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   6600
         Width           =   585
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Time From"
         Height          =   195
         Left            =   3000
         TabIndex        =   30
         Top             =   6600
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "To"
         Height          =   255
         Left            =   5160
         TabIndex        =   29
         Top             =   6600
         Width           =   375
      End
      Begin VB.Label lblMinutes 
         Height          =   435
         Left            =   6780
         TabIndex        =   28
         Top             =   6600
         Width           =   375
      End
      Begin VB.Label lblRequired 
         Height          =   375
         Left            =   7140
         TabIndex        =   27
         Top             =   6600
         Width           =   2535
      End
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   375
      Left            =   18000
      TabIndex        =   15
      Top             =   8220
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   661
      _Version        =   393216
      PromptChar      =   "_"
   End
End
Attribute VB_Name = "frmscheduling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub AddProgram()
    If rsProgram.State = 1 Then rsProgram.Close
    rsProgram.Open "select * from tbl_program", cn, adOpenStatic, adLockPessimistic
    cmbProgram.Clear
    While Not rsProgram.EOF
        cmbProgram.AddItem rsProgram.Fields("program_name")
        cmbProgram.ItemData(cmbProgram.NewIndex) = rsProgram.Fields("program_id")
        rsProgram.MoveNext
    Wend
End Sub


Private Sub cmbJudges_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then cmdAdd.SetFocus
End Sub

Private Sub cmbProgram_Click()
On Error GoTo lbl
    If cmbProgram.ListIndex <> -1 Then
        If rsProgram.State = 1 Then rsProgram.Close
        rsProgram.Open "select * from tbl_program where program_id=" & cmbProgram.ItemData(cmbProgram.ListIndex), cn, adOpenStatic, adLockPessimistic
        If rsProgram.EOF = False Then
            txtPgmDescription.Text = rsProgram.Fields("program_description")
        Else
            txtPgmDescription.Text = ""
        End If
        AddJudges
        
         Dim c As Integer
        If rsMemberPgm.State = 1 Then rsMemberPgm.Close
        rsMemberPgm.Open "select * from tbl_MemberPgm where Program_Id=" & cmbProgram.ItemData(cmbProgram.ListIndex), cn, adOpenStatic, adLockPessimistic
        If rsMemberPgm.EOF = True Then
            c = 0
        Else
            If rsMemberPgm.State = 1 Then rsMemberPgm.Close
            rsMemberPgm.Open "select count(program_id) as c from tbl_MemberPgm where Program_Id=" & cmbProgram.ItemData(cmbProgram.ListIndex), cn, adOpenStatic, adLockPessimistic
            c = rsMemberPgm.Fields("c")
        End If
         If rsProgram.State = 1 Then rsProgram.Close
        rsProgram.Open "select * from tbl_program where program_id=" & cmbProgram.ItemData(cmbProgram.ListIndex), cn, adOpenStatic, adLockPessimistic
        If rsProgram.EOF = False Then
            lblMinutes.Caption = c * rsProgram.Fields("program_time")
            lblRequired.Caption = " Minutes Required"
        Else
            lblMinutes.Caption = 0
            lblRequired.Caption = " Minutes Required"
        End If
        
    Else
        txtPgmDescription.Text = ""
        cmbJudges.Clear
        lblMinutes.Caption = ""
        lblRequired.Caption = ""
    End If
Exit Sub
lbl:
    MsgBox Err.Description, vbInformation, "Message"
End Sub

Private Sub cmbProgram_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtPgmDescription.SetFocus
End Sub

Private Sub cmbStage_Change()
    cmbStage_Click
End Sub

Private Sub cmbStage_Click()
On Error GoTo lbl
    If cmbStage.ListIndex <> -1 Then
        If rsstage.State = 1 Then rsstage.Close
        rsstage.Open "select * from tbl_stage where stage_id=" & cmbStage.ItemData(cmbStage.ListIndex), cn, adOpenStatic, adLockPessimistic
        If rsstage.EOF = False Then
            txtVenue.Text = rsstage.Fields("venue")
            txtLocation.Text = rsstage.Fields("location")
            txtStageContact.Text = rsstage.Fields("contact_no")
            txtContactPerson.Text = rsstage.Fields("contact_person")
            DTPicker1_Change
            txtTimeFrom_LostFocus
        Else
            Clearing1
        End If
    Else
        Clearing1
    End If
Exit Sub
lbl:
    MsgBox Err.Description, vbInformation, "Message"
End Sub

Private Sub cmbStage_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then DTPicker1.SetFocus
End Sub

Private Sub cmdAdd_Click()
    If cmbJudges.ListIndex <> -1 And cmbProgram.ListIndex <> -1 Then
        flxGrid.TextMatrix(flxGrid.Rows - 1, 0) = flxGrid.Rows - 1
        flxGrid.TextMatrix(flxGrid.Rows - 1, 1) = cmbJudges.Text
        flxGrid.TextMatrix(flxGrid.Rows - 1, 2) = cmbJudges.ItemData(cmbJudges.ListIndex)
        flxGrid.Rows = flxGrid.Rows + 1
        cmbJudges.Text = ""
        cmbJudges.SetFocus
    Else
        MsgBox "Select Program and Judge", vbInformation, "Message"
    End If
End Sub

Private Sub cmdAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then cmbStage.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo lbl
    Dim i As Integer, PId As Integer
    If cmbProgram.ListIndex <> -1 And cmbStage.ListIndex <> -1 And flxGrid.Rows > 2 And Len(Trim(txtTimeFrom.Text)) <> 0 And Len(Trim(txtTimeTo.Text)) <> 0 Then
        If rspschedule.State = 1 Then rspschedule.Close
        rspschedule.Open "select * from tbl_program_scheduling where program_id=" & cmbProgram.ItemData(cmbProgram.ListIndex), cn, adOpenStatic, adLockPessimistic
        If rspschedule.EOF = True Then
            cn.Execute "insert into tbl_program_scheduling(program_id,stage_id,pdate,time_from,time_to) values(" & _
            cmbProgram.ItemData(cmbProgram.ListIndex) & "," & cmbStage.ItemData(cmbStage.ListIndex) & ",'" & DTPicker1.Value & _
            "','" & txtTimeFrom.Text & "','" & txtTimeTo.Text & "')"
            If rspschedule.State = 1 Then rspschedule.Close
            rspschedule.Open "select max(program_scheduling_id) as m from tbl_program_scheduling", cn, adOpenStatic, adLockPessimistic
            PId = rspschedule.Fields("m")
            For i = 1 To flxGrid.Rows - 2
                cn.Execute "insert into tbl_PgmJudge(program_sheduling_id,Judge_Id) values(" & PId & "," & Val(flxGrid.TextMatrix(i, 2)) & ")"
            Next
        Else
            If MsgBox("Already scheduled, Do U want to modify", vbYesNo + vbQuestion) = vbYes Then
                PId = rspschedule.Fields("program_scheduling_id")
                If rspschedule.State = 1 Then rspschedule.Close
                cn.Execute "update tbl_program_scheduling set stage_id=" & cmbStage.ItemData(cmbStage.ListIndex) & ",pdate='" & _
                DTPicker1.Value & "',time_from='" & txtTimeFrom.Text & "',time_to='" & txtTimeTo.Text & "' where program_id=" & cmbProgram.ItemData(cmbProgram.ListIndex)
                cn.Execute "delete from tbl_PgmJudge where program_sheduling_id=" & PId
                For i = 1 To flxGrid.Rows - 2
                    cn.Execute "insert into tbl_PgmJudge(program_sheduling_id,Judge_Id) values(" & PId & "," & Val(flxGrid.TextMatrix(i, 2)) & ")"
                Next
            End If
        End If
          '''***********************Program Time scheduling ***************************'''
        Dim hfrom As Integer, mfrom As Integer, hto As Integer, mto As Integer, hdiff As Integer, mdiff As Integer, h As Integer, m As Integer, ptid As Integer, time As String
        If rsPgmTiming.State = 1 Then rsPgmTiming.Close
        rsPgmTiming.Open "select * from tbl_PgmTiming where schedule_id=" & PId, cn, adOpenStatic, adLockPessimistic
        If rsPgmTiming.EOF = False Then
            If rsPgmTiming.State = 1 Then rsPgmTiming.Close
            cn.Execute "delete from tbl_PgmTiming where schedule_id=" & PId
        End If
        
        hfrom = Mid(txtTimeFrom.Text, 1, 2)
        mfrom = Mid(txtTimeFrom.Text, 4, 5)
        hto = Mid(txtTimeTo.Text, 1, 2)
        mto = Mid(txtTimeTo.Text, 4, 5)
        If hfrom > hto Then
            hdiff = 12 - hfrom + hto
        Else
            hdiff = hto - hfrom
        End If
        If mfrom < mto Then
            mdiff = mto - mfrom
        Else
            mdiff = 60 - mfrom + mto
            hdiff = hdiff - 1
        End If
        mdiff = hdiff * 60 + mdiff
        While (mdiff > 0)
            time = hfrom & ":" & mfrom
            cn.Execute "insert into tbl_PgmTiming(pgm_id,schedule_id,time)values(" & cmbProgram.ItemData(cmbProgram.ListIndex) & _
            "," & PId & ",'" & time & "')"
            If rsProgram.State = 1 Then rsProgram.Close
            rsProgram.Open "select * from tbl_program where program_id=" & cmbProgram.ItemData(cmbProgram.ListIndex), cn, adOpenStatic, adLockPessimistic
            
            mdiff = mdiff - rsProgram.Fields("program_time")
            mfrom = mfrom + rsProgram.Fields("program_time")
            If mfrom >= 60 Then
                hfrom = hfrom + 1
                mfrom = mfrom - 60
                If hfrom > 12 Then hfrom = 1
            End If
        Wend
        
        
        
        
        
        
        DispGrid
        Clearing1
        txtPgmDescription.Text = ""
        cmbProgram.Text = ""
        
    Else
        MsgBox "Select program,stage,judges and date", vbInformation, "Message"
     End If
       Exit Sub
lbl:
    MsgBox Err.Description, vbInformation, "Message"
   
End Sub

Private Sub DTPicker1_Change()
On Error GoTo lbl
     Dim th As Integer, tm As Integer, tfrom As String, h As Integer, m As Integer
    If cmbStage.ListIndex <> -1 Then
        If rspschedule.State = 1 Then rspschedule.Close
        rspschedule.Open "select * from tbl_program_scheduling where pdate='" & Format(DTPicker1.Value, "mm/dd/yy") & "' and stage_id=" & cmbStage.ItemData(cmbStage.ListIndex), cn, adOpenStatic, adLockPessimistic
        If rspschedule.EOF = True Then
            txtTimeFrom.Text = "10:00"
            txtTimeFrom_LostFocus
        Else
            If rspschedule.State = 1 Then rspschedule.Close
            rspschedule.Open "Select * from tbl_program_scheduling where pdate='" & Format(DTPicker1.Value, "mm/dd/yy") & "' and stage_id=" & _
            cmbStage.ItemData(cmbStage.ListIndex) & " order by program_scheduling_id desc", cn, adOpenStatic, adLockPessimistic
            th = Mid(rspschedule.Fields("time_to"), 1, 2)
            tm = Mid(rspschedule.Fields("time_to"), 4, 2)
            tm = tm + 20
            If tm >= 60 Then
                h = tm / 60
                tm = tm Mod 60
            End If
            th = th + h
            If th > 12 Then
                th = th - 12
            End If
            If Len(Str(th)) < 2 Then
                tfrom = "0" & Str(th)
            Else
                tfrom = Str(th)
            End If
            If Len(Str(tm)) < 2 Then
                tfrom = Trim(tfrom) & ":" & "0" & Trim(Str(tm))
            Else
                 tfrom = Trim(tfrom) & ":" & Trim(Str(tm))
            End If
            txtTimeFrom.Text = Trim(tfrom)
            txtTimeFrom_LostFocus
        End If
    Else
        MsgBox "Select Stage", vbInformation, "Message"
    End If
Exit Sub
lbl:
    MsgBox Err.Description, vbInformation, "Message"
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then txtTimeFrom.SetFocus
End Sub

Private Sub Form_Load()
    AddProgram
    DispGrid
    AddStage
End Sub

Public Sub DispGrid()
    flxGrid.Clear
    flxGrid.Rows = 2
    flxGrid.ColWidth(0) = 1000
    flxGrid.ColWidth(1) = 2500
    flxGrid.ColWidth(2) = 0
    flxGrid.TextMatrix(0, 0) = "Sl No"
    flxGrid.TextMatrix(0, 1) = "Judge"
    flxGrid.TextMatrix(0, 2) = "Judge_Id"
    
End Sub

Public Sub AddJudges()
    If cmbProgram.ListIndex <> -1 Then
        If rsJudges.State = 1 Then rsJudges.Close
        rsJudges.Open "select * from tbl_judges where program_id=" & cmbProgram.ItemData(cmbProgram.ListIndex), cn, adOpenStatic, adLockPessimistic
        cmbJudges.Clear
        While Not rsJudges.EOF
            cmbJudges.AddItem rsJudges.Fields("name")
            cmbJudges.ItemData(cmbJudges.NewIndex) = rsJudges.Fields("judge_id")
            rsJudges.MoveNext
        Wend
    Else
        cmbJudges.Clear
    End If
End Sub

Public Sub AddStage()
    If rsstage.State = 1 Then rsstage.Close
    rsstage.Open "Select * from tbl_stage", cn, adOpenStatic, adLockPessimistic
    cmbStage.Clear
    While Not rsstage.EOF
        cmbStage.AddItem rsstage.Fields("stage_no")
        cmbStage.ItemData(cmbStage.NewIndex) = rsstage.Fields("stage_id")
        rsstage.MoveNext
    Wend
End Sub

Public Sub Clearing1()
    txtVenue.Text = ""
    txtLocation.Text = ""
    txtStageContact.Text = ""
    txtContactPerson.Text = ""
End Sub







Private Sub txtContactPerson_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtStageContact.SetFocus
End Sub

Private Sub txtLocation_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtContactPerson.SetFocus
End Sub

Private Sub txtPgmDescription_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmbJudges.SetFocus
End Sub


Private Sub txtStageContact_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdsave.SetFocus
End Sub

Private Sub txtStageContact_LostFocus()
 If Not ValPhone(txtStageContact.Text) Then
        MsgBox "Not a valid contactno", vbInformation, "Message"
    End If
End Sub

Private Sub txtTimeFrom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtTimeTo.SetFocus
End Sub

Private Sub txtTimeFrom_LostFocus()
    If Len(Trim(txtTimeFrom.Text)) = 5 Then
    Dim fh As Integer, fm As Integer, h As Integer, m As Integer, th As Integer, tm As Integer, timeto As String
    fh = Mid(txtTimeFrom.Text, 1, 2)
    fm = Mid(txtTimeFrom.Text, 4, 2)
    fm = fm + Val(lblMinutes.Caption)
    If fm >= 60 Then
        h = fm / 60
        m = fm Mod 60
    Else
        m = fm
    End If
    th = fh + h
    tm = m
    If th >= 13 Then th = th - 12
    If Len(Trim(th)) <> 2 Then
        timeto = "0" & th
    Else
        timeto = th
    End If
    If Len(Trim(tm)) <> 2 Then
        timeto = timeto & ":0" & Trim(Str(tm))
    Else
         timeto = timeto & ":" & Trim(Str(tm))
    End If
    txtTimeTo.Text = timeto
Else
    MsgBox "Time from is not in the correct format", vbInformation, "Message"
End If
End Sub



Private Sub txtTimeTo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtVenue.SetFocus
End Sub



Private Sub txtVenue_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtLocation.SetFocus
End Sub
