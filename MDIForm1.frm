VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "School Youth Festival"
   ClientHeight    =   5055
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   11400
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":0442
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "District"
            Object.ToolTipText     =   "District"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "School"
            Object.ToolTipText     =   "School"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Stage"
            Object.ToolTipText     =   "Stage"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "volunteer"
            Object.ToolTipText     =   "volunteer"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Judges"
            Object.ToolTipText     =   "Judges"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Greenroom"
            Object.ToolTipText     =   "Greenroom"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Program"
            Object.ToolTipText     =   "Program"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Prizecategory"
            Object.ToolTipText     =   "Prizecategory"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Prize"
            Object.ToolTipText     =   "Prize"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Memberregistration"
            Object.ToolTipText     =   "Memberregistration"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Programscheduling"
            Object.ToolTipText     =   "Programscheduling"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ScoreEntering"
            Object.ToolTipText     =   "ScoreEntering"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ProgramTiming"
            Object.ToolTipText     =   "ProgramTiming"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "Exit"
            ImageIndex      =   14
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   4740
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "2016-01-11"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   31751
            MinWidth        =   31751
            Text            =   "State Youth Festival"
            TextSave        =   "State Youth Festival"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "PM 01:12"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3720
      Top             =   1740
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":B8CB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":B9108
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":B955A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":BF7F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":BFC46
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":C0098
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":C04EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":C093C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":C0D8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":109060
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1094B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":109904
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":10FB9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":10FFF0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "Settings"
      Begin VB.Menu mnuDateandTime 
         Caption         =   "System Date"
      End
      Begin VB.Menu mnuCalculator 
         Caption         =   "Calculator"
      End
      Begin VB.Menu mnuNewUser 
         Caption         =   "New User"
      End
      Begin VB.Menu mnuChangePass 
         Caption         =   "Change Password"
      End
      Begin VB.Menu mnuProfile 
         Caption         =   "Profile"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu mnuRegistration 
      Caption         =   "Registration"
      Begin VB.Menu mnuAdmin 
         Caption         =   "Admin"
         Begin VB.Menu mnuDistrict 
            Caption         =   "District"
         End
         Begin VB.Menu mnuSchool 
            Caption         =   "School"
         End
         Begin VB.Menu mnuStage 
            Caption         =   "Stage"
         End
         Begin VB.Menu mnuVolunteer 
            Caption         =   "Volunteer"
         End
         Begin VB.Menu mnuJudges 
            Caption         =   "Judges"
         End
         Begin VB.Menu mnuGreenroom 
            Caption         =   "Greenroom"
         End
         Begin VB.Menu mnuProgram 
            Caption         =   "Program"
         End
         Begin VB.Menu mnuPrize_category 
            Caption         =   "Prize category"
         End
         Begin VB.Menu mnuPrize 
            Caption         =   "Prize"
         End
      End
      Begin VB.Menu mnuFestivalmanager 
         Caption         =   "Festivalmanager"
         Begin VB.Menu mnuMember_registration 
            Caption         =   "Member registration"
         End
         Begin VB.Menu mnuProgram_scheduling 
            Caption         =   "Program scheduling"
         End
      End
      Begin VB.Menu mnuPrize_manager 
         Caption         =   "Prize manager"
         Begin VB.Menu mnuScoreentering 
            Caption         =   "Score entering"
         End
      End
      Begin VB.Menu mnuProgramtiming 
         Caption         =   "Program timing"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "Reports"
      Begin VB.Menu mnuSchoolDetails 
         Caption         =   "School Details"
      End
      Begin VB.Menu mnuProgramDetails 
         Caption         =   "Program Details"
      End
      Begin VB.Menu mnuJudgesDetails 
         Caption         =   "Judges"
      End
      Begin VB.Menu mnuMemberDetails 
         Caption         =   "Member Details"
         Begin VB.Menu mnuSchoolwise 
            Caption         =   "Schoolwise"
         End
         Begin VB.Menu mnuProgramwise 
            Caption         =   "Programwise"
         End
      End
      Begin VB.Menu mnuScoreDetails 
         Caption         =   "Score Details"
         Begin VB.Menu mnuScorePgmwise 
            Caption         =   "Programwise"
         End
         Begin VB.Menu mnuScoreAll 
            Caption         =   "All"
         End
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
    If cn.State = 1 Then cn.Close
    cn.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=db_YouthFestival;Data Source=HP-PC"
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    If MsgBox("Are You Sure to Exit", vbYesNo + vbQuestion, "Warning") = vbYes Then
        Cancel = 0
    Else
        Cancel = 1
    End If
End Sub

Private Sub mnuCalculator_Click()
    Shell App.Path & "\calc.exe"
End Sub

Private Sub mnuChangePass_Click()
    frmChangePassword.Show
End Sub

Private Sub mnuDateandTime_Click()
    frmSystemDate.Show
End Sub

Private Sub mnuDistrict_Click()
    frmDistrict.Show
End Sub

Private Sub mnuExit_Click()
    If MsgBox("Are You sure to Exit", vbYesNo + vbQuestion, "Warning") = vbYes Then
        End
    End If
End Sub

Private Sub mnuGreenroom_Click()
    frmgreen_room.Show
End Sub

Private Sub mnuJudges_Click()
frmjudges.Show
End Sub

Private Sub mnuJudgesDetails_Click()
    rptJudges.Show
End Sub

Private Sub mnuMember_registration_Click()
    frmmemberregistration.Show
End Sub

Private Sub mnuNewUser_Click()
    frmnewuser.Show
End Sub





Private Sub mnuPrize_category_Click()
frmPrizeCategory.Show
End Sub

Private Sub mnuPrize_Click()
frmPrize.Show
End Sub

Private Sub mnuProfile_Click()
    frmprofile.Show
End Sub

Private Sub mnuProgram_Click()
frmprogram.Show
End Sub

Private Sub mnuProgram_scheduling_Click()
   frmscheduling.Show
End Sub

Private Sub mnuProgramDetails_Click()
    rptProgram.Show
End Sub

Private Sub mnuProgramtiming_Click()
    frmPgmTiming.Show
End Sub

Private Sub mnuProgramwise_Click()
    MenuItem = "Member"
    frmProgramwise.Show
End Sub

Private Sub mnuschool_Click()
frmSchool.Show
End Sub

Private Sub mnuSchoolDetails_Click()
    rptSchool.Show
End Sub

Private Sub mnuSchoolwise_Click()
    frmSchoolwise.Show
End Sub

Private Sub mnuScoreAll_Click()
    rptScore.Show
End Sub

Private Sub mnuScoreentering_Click()
    frmScoreentry.Show
End Sub

Private Sub mnuScorePgmwise_Click()
    MenuItem = "Score"
    frmProgramwise.Show
End Sub

Private Sub mnuStage_Click()
frmStage.Show
End Sub

Private Sub mnuVolunteer_Click()
frmvolunteer.Show
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
  Case "District"
       mnuDistrict_Click
  Case "School"
       mnuschool_Click
  Case "Stage"
       mnuStage_Click
  Case "volunteer"
       mnuVolunteer_Click
  Case "Judges"
       mnuJudges_Click
  Case "Greenroom"
       mnuGreenroom_Click
  Case "Program"
       mnuProgram_Click
  Case "Prizecategory"
       mnuPrize_category_Click
  Case "Prize"
       mnuPrize_Click
  Case "Memberregistration"
       mnuMember_registration_Click
  Case "Programscheduling"
       mnuProgram_scheduling_Click
  Case "ScoreEntering"
       mnuScoreentering_Click
  Case "ProgramTiming"
       mnuProgramtiming_Click
  Case "Exit"
       mnuExit_Click
  End Select
End Sub

