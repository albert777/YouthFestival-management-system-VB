VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPgmTiming 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Program Timing"
   ClientHeight    =   1200
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6270
   Icon            =   "frmPgmTiming.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   6270
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   900
      Top             =   120
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   60
      TabIndex        =   1
      Top             =   660
      Width           =   6075
      _ExtentX        =   10716
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdPgmTiming 
      Caption         =   "Click here to generate Program Timing"
      Height          =   555
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6195
   End
End
Attribute VB_Name = "frmPgmTiming"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPgmTiming_Click()
On Error GoTo lbl
    Dim rsTemp As New Recordset, PId As Integer
    If rsMemberPgm.State = 1 Then rsMemberPgm.Close
    rsMemberPgm.Open "Select memberreg_id,count(program_id) as c from tbl_MemberPgm group by MemberReg_id order by c desc", cn, adOpenStatic, adLockPessimistic
    While Not rsMemberPgm.EOF
        'MsgBox rsMemberProgram.Fields("c") & "-" & rsMemberProgram.Fields("memberregistrationid")
        If rsTemp.State = 1 Then rsTemp.Close
        rsTemp.Open "Select * from tbl_memberpgm where memberreg_id=" & rsMemberPgm.Fields("memberreg_id"), cn, adOpenStatic, adLockPessimistic
        While Not rsTemp.EOF
            If rspschedule.State = 1 Then rspschedule.Close
            rspschedule.Open "select * from tbl_program_scheduling where program_id=" & rsTemp.Fields("program_id"), cn, adOpenStatic, adLockPessimistic
            If rspschedule.EOF = False Then
                If rsPgmTiming.State = 1 Then rsPgmTiming.Close
                rsPgmTiming.Open "select * from tbl_pgmTiming where schedule_id=" & rspschedule.Fields("program_scheduling_id") & " and member_id is null", cn, adOpenStatic, adLockPessimistic
                If rsPgmTiming.EOF = False Then
                    PId = rsPgmTiming.Fields("pgmTimingId")
                    If rsPgmTiming.State = 1 Then rsPgmTiming.Close
                    cn.Execute "update tbl_pgmTiming set member_id=" & rsMemberPgm.Fields("memberreg_id") & " where pgmtimingid=" & PId
                End If
            End If
            rsTemp.MoveNext
        Wend
        rsMemberPgm.MoveNext
    Wend
    rptPgmTiming.Show
Exit Sub
lbl:
    MsgBox Err.Description, vbInformation, "Message"
End Sub

Private Sub Timer1_Timer()
    If ProgressBar1.Value >= 100 Then
        Timer1.Enabled = False
    Else
        ProgressBar1.Value = ProgressBar1.Value + 10
    End If
End Sub
