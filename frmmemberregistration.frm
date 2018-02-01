VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmmemberregistration 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Member Registration"
   ClientHeight    =   9720
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11925
   Icon            =   "frmmemberregistration.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9720
   ScaleWidth      =   11925
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   9675
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   11895
      Begin VB.ComboBox cmbName 
         Height          =   315
         Left            =   7560
         TabIndex        =   3
         Top             =   240
         Width           =   4215
      End
      Begin VB.TextBox txtSchoolAddress 
         Height          =   1215
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   4140
         Width           =   4215
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   495
         Left            =   10800
         TabIndex        =   25
         Top             =   9060
         Width           =   975
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   495
         Left            =   9720
         TabIndex        =   24
         Top             =   9060
         Width           =   975
      End
      Begin VB.TextBox txtContactNo 
         Height          =   405
         Left            =   7560
         MaxLength       =   10
         TabIndex        =   7
         Top             =   2640
         Width           =   2655
      End
      Begin VB.TextBox txtAddress 
         Height          =   975
         Left            =   7560
         MaxLength       =   50
         TabIndex        =   4
         Top             =   600
         Width           =   4215
      End
      Begin VB.ComboBox cmbSchool 
         Height          =   1740
         Left            =   1680
         Style           =   1  'Simple Combo
         TabIndex        =   1
         Top             =   2280
         Width           =   4215
      End
      Begin VB.ComboBox cmbDistrict 
         Height          =   1935
         Left            =   1680
         Style           =   1  'Simple Combo
         TabIndex        =   0
         Top             =   240
         Width           =   4215
      End
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   7560
         TabIndex        =   23
         Top             =   3480
         Width           =   3495
         Begin VB.OptionButton optMale 
            Caption         =   "Male"
            Height          =   315
            Left            =   240
            TabIndex        =   9
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton optFemale 
            Caption         =   "Female"
            Height          =   255
            Left            =   1800
            TabIndex        =   10
            Top             =   300
            Width           =   1215
         End
      End
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   7560
         TabIndex        =   22
         Top             =   4140
         Width           =   3495
         Begin VB.OptionButton optJunior 
            Caption         =   "Junior"
            Height          =   375
            Left            =   240
            TabIndex        =   11
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton optSenior 
            Caption         =   "Senior"
            Height          =   375
            Left            =   1320
            TabIndex        =   12
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.TextBox txtGname 
         Height          =   375
         Left            =   7560
         MaxLength       =   50
         TabIndex        =   5
         Top             =   1680
         Width           =   4215
      End
      Begin VB.Frame Frame4 
         Caption         =   "Program Details"
         Height          =   3735
         Left            =   120
         TabIndex        =   18
         Top             =   5280
         Width           =   11655
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete"
            Height          =   375
            Left            =   9240
            TabIndex        =   38
            Top             =   1440
            Width           =   1275
         End
         Begin VB.TextBox txtDescription 
            Height          =   315
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   1440
            Width           =   4215
         End
         Begin VB.ComboBox cmbProgram 
            Height          =   1155
            Left            =   1680
            Style           =   1  'Simple Combo
            TabIndex        =   13
            Top             =   240
            Width           =   4215
         End
         Begin VB.ComboBox cmbPrize 
            Height          =   1155
            Left            =   8040
            Style           =   1  'Simple Combo
            TabIndex        =   15
            Top             =   240
            Width           =   3495
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add"
            Height          =   375
            Left            =   8220
            TabIndex        =   16
            Top             =   1440
            Width           =   975
         End
         Begin MSFlexGridLib.MSFlexGrid flxGrid1 
            Height          =   1815
            Left            =   120
            TabIndex        =   19
            Top             =   1800
            Width           =   10455
            _ExtentX        =   18441
            _ExtentY        =   3201
            _Version        =   393216
            Cols            =   5
         End
         Begin VB.Label Label14 
            Caption         =   "Description"
            Height          =   375
            Left            =   180
            TabIndex        =   37
            Top             =   1380
            Width           =   1515
         End
         Begin VB.Label Label3 
            Caption         =   "Program"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label9 
            Caption         =   "District Level Prize Category"
            Height          =   375
            Left            =   6000
            TabIndex        =   20
            Top             =   360
            Width           =   2175
         End
      End
      Begin MSComCtl2.DTPicker dtpDOB 
         Height          =   375
         Left            =   7560
         TabIndex        =   8
         Top             =   3120
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         _Version        =   393216
         Format          =   76677121
         CurrentDate     =   42136
      End
      Begin VB.TextBox txtRelationship 
         Height          =   375
         Left            =   7560
         MaxLength       =   50
         TabIndex        =   6
         Top             =   2160
         Width           =   4215
      End
      Begin VB.Label Label13 
         Caption         =   "Address"
         Height          =   375
         Left            =   120
         TabIndex        =   36
         Top             =   4200
         Width           =   1575
      End
      Begin VB.Label Label10 
         Caption         =   "Member Category"
         Height          =   255
         Left            =   6000
         TabIndex        =   35
         Top             =   4320
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Gender"
         Height          =   255
         Left            =   6000
         TabIndex        =   34
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "DOB"
         Height          =   255
         Left            =   6000
         TabIndex        =   33
         Top             =   3180
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Contact No."
         Height          =   255
         Left            =   6000
         TabIndex        =   32
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Address"
         Height          =   255
         Left            =   6000
         TabIndex        =   31
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Name"
         Height          =   255
         Left            =   6000
         TabIndex        =   30
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "School"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "District"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Guardian's Name"
         Height          =   375
         Left            =   6000
         TabIndex        =   27
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label12 
         Caption         =   "Relationship"
         Height          =   375
         Left            =   6000
         TabIndex        =   26
         Top             =   2160
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmmemberregistration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub AddDistrict()
  On Error GoTo lbl
    If rsDistrict.State = 1 Then rsDistrict.Close
    rsDistrict.Open "select * from tbl_district", cn, adOpenStatic, adLockPessimistic
    cmbDistrict.Clear
    While Not rsDistrict.EOF
        cmbDistrict.AddItem rsDistrict.Fields("district_name")
        cmbDistrict.ItemData(cmbDistrict.NewIndex) = rsDistrict.Fields("district_id")
        rsDistrict.MoveNext
    Wend
   Exit Sub
lbl:
    MsgBox Err.Description, vbInformation, "Message"
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

Private Sub cmbDistrict_DropDown()
If KeyCode = 13 Then cmbSchool.SetFocus
End Sub

Private Sub Cmbdistrict_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then cmbSchool.SetFocus
End Sub

Private Sub cmbName_Click()
On Error GoTo lbl
    If Cmbname.ListIndex <> -1 Then
        If rsMemberReg.State = 1 Then rsMemberReg.Close
        rsMemberReg.Open "select * from tbl_member_registration where member_registration_id=" & Cmbname.ItemData(Cmbname.ListIndex), cn, adOpenStatic, adLockPessimistic
        If rsMemberReg.EOF = False Then
            txtAddress.Text = rsMemberReg.Fields("address")
            txtContactNo.Text = rsMemberReg.Fields("contact_no")
            txtGname.Text = rsMemberReg.Fields("g_name")
            txtRelationship.Text = rsMemberReg.Fields("relationship")
            If rsMemberReg.Fields("gender") = "M" Then
                optMale.Value = True
            Else
                optFemale.Value = True
            End If
            If rsMemberReg.Fields("member_category") = "J" Then
                optJunior.Value = True
            Else
                optSenior.Value = True
            End If
            If rsMemberPgm.State = 1 Then rsMemberPgm.Close
            rsMemberPgm.Open "select * from tbl_MemberPgm,tbl_Program,tbl_prize_Category where tbl_MemberPgm.Program_id=tbl_program.program_id and tbl_MemberPgm.DistLevelPrizeCat_id=tbl_prize_category.prize_category_id and  MemberReg_ID=" & Cmbname.ItemData(Cmbname.ListIndex), cn, adOpenStatic, adLockPessimistic
            DispGrid
            While Not rsMemberPgm.EOF
                flxGrid1.TextMatrix(flxGrid1.Rows - 1, 0) = flxGrid1.Rows - 1
                flxGrid1.TextMatrix(flxGrid1.Rows - 1, 1) = rsMemberPgm.Fields("program_name")
                flxGrid1.TextMatrix(flxGrid1.Rows - 1, 2) = rsMemberPgm.Fields("category_name")
                flxGrid1.TextMatrix(flxGrid1.Rows - 1, 3) = rsMemberPgm.Fields("program_id")
                flxGrid1.TextMatrix(flxGrid1.Rows - 1, 4) = rsMemberPgm.Fields("prize_category_id")
                rsMemberPgm.MoveNext
                flxGrid1.Rows = flxGrid1.Rows + 1
            Wend
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

Private Sub cmbName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtAddress.SetFocus
End Sub

Private Sub cmbPrize_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then cmdAdd.SetFocus
End Sub

Private Sub cmbProgram_Change()
    cmbProgram_Click
End Sub

Private Sub cmbProgram_Click()
On Error GoTo lbl
    If cmbProgram.ListIndex <> -1 Then
        If rsProgram.State = 1 Then rsProgram.Close
        rsProgram.Open "select * from tbl_program where program_id=" & cmbProgram.ItemData(cmbProgram.ListIndex), cn, adOpenStatic, adLockPessimistic
        If rsProgram.EOF = False Then
            txtDescription.Text = rsProgram.Fields("program_description")
        Else
            txtDescription.Text = ""
        End If
    Else
        txtDescription.Text = ""
     End If
       Exit Sub
lbl:
    MsgBox Err.Description, vbInformation, "Message"
   
End Sub

Private Sub cmbProgram_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtDescription.SetFocus
End Sub

Private Sub cmbSchool_Change()
    cmbSchool_Click
End Sub

Private Sub cmbSchool_Click()
On Error GoTo lbl
    If cmbSchool.ListIndex <> -1 Then
        If rsSchool.State = 1 Then rsSchool.Close
        rsSchool.Open "Select * from tbl_school where school_id=" & cmbSchool.ItemData(cmbSchool.ListIndex), cn, adOpenStatic, adLockPessimistic
        If rsSchool.EOF = False Then
            txtSchoolAddress.Text = rsSchool.Fields("address")
        Else
            txtSchoolAddress.Text = ""
        End If
        If rsMemberReg.State = 1 Then rsMemberReg.Close
        rsMemberReg.Open "select * from tbl_member_registration where school_id=" & cmbSchool.ItemData(cmbSchool.ListIndex), cn, adOpenStatic, adLockPessimistic
        Cmbname.Clear
        While Not rsMemberReg.EOF
            Cmbname.AddItem rsMemberReg.Fields("name")
            Cmbname.ItemData(Cmbname.NewIndex) = rsMemberReg.Fields("member_registration_id")
            rsMemberReg.MoveNext
        Wend
    Else
        txtSchoolAddress.Text = ""
        Cmbname.Clear
        Cmbname.Text = ""
     End If
      Exit Sub
lbl:
    MsgBox Err.Description, vbInformation, "Message"
   
End Sub

Private Sub cmbSchool_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtAddress.SetFocus
End Sub

Private Sub cmdAdd_Click()
On Error GoTo lbl
    If cmbProgram.ListIndex <> -1 And cmbPrize.ListIndex <> -1 Then
        flxGrid1.TextMatrix(flxGrid1.Rows - 1, 0) = flxGrid1.Rows - 1
        flxGrid1.TextMatrix(flxGrid1.Rows - 1, 1) = cmbProgram.Text
        flxGrid1.TextMatrix(flxGrid1.Rows - 1, 2) = cmbPrize.Text
        flxGrid1.TextMatrix(flxGrid1.Rows - 1, 3) = cmbProgram.ItemData(cmbProgram.ListIndex)
        flxGrid1.TextMatrix(flxGrid1.Rows - 1, 4) = cmbPrize.ItemData(cmbPrize.ListIndex)
        flxGrid1.Rows = flxGrid1.Rows + 1
    Exit Sub
lbl:
    MsgBox Err.Description, vbInformation, "Message"
    End If
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
On Error GoTo lbl
    Dim i As Integer
    If flxGrid1.RowSel > 0 Then
        If MsgBox("R U sure to delete???", vbYesNo + vbQuestion, "Warning") = vbYes Then
            For i = flxGrid1.RowSel To flxGrid1.Rows - 2
                flxGrid1.TextMatrix(i, 0) = flxGrid1.TextMatrix(i + 1, 0)
                flxGrid1.TextMatrix(i, 1) = flxGrid1.TextMatrix(i + 1, 1)
                flxGrid1.TextMatrix(i, 2) = flxGrid1.TextMatrix(i + 1, 2)
                flxGrid1.TextMatrix(i, 3) = flxGrid1.TextMatrix(i + 1, 3)
                flxGrid1.TextMatrix(i, 4) = flxGrid1.TextMatrix(i + 1, 4)
            Next
            flxGrid1.Rows = flxGrid1.Rows - 1
        End If
    Else
        MsgBox "Select one row to delete", vbInformation, "Message"
    End If
      Exit Sub
lbl:
    MsgBox Err.Description, vbInformation, "Message"
    
End Sub

Private Sub cmdSave_Click()
On Error GoTo lbl
    Dim gen As String, Cat As String, i As Integer, RegId As Integer
    If cmbSchool.ListIndex <> -1 And Len(Trim(Cmbname.Text)) <> 0 And flxGrid1.Rows > 2 Then
        If optMale.Value = True Then
            gen = "M"
        Else
            gen = "F"
        End If
        If optJunior.Value = True Then
            Cat = "J"
        Else
            Cat = "S"
        End If
        If Cmbname.ListIndex <> -1 Then
            If rsMemberReg.State = 1 Then rsMemberReg.Close
            rsMemberReg.Open "select * from tbl_member_registration where member_registration_id=" & Cmbname.ItemData(Cmbname.ListIndex), cn, adOpenStatic, adLockPessimistic
            If rsMemberReg.EOF = True Then
                cn.Execute "insert into tbl_member_registration(school_id,name,address,contact_no,dob,gender,member_category,g_name,relationship) values(" & _
                cmbSchool.ItemData(cmbSchool.ListIndex) & ",'" & Cmbname.Text & "','" & txtAddress.Text & "','" & txtContactNo.Text & "','" & dtpDOB.Value & _
                "','" & gen & "','" & Cat & "','" & txtGname.Text & "','" & txtRelationship.Text & "')"
                
                If rsMemberReg.State = 1 Then rsMemberReg.Close
                rsMemberReg.Open "Select max(member_registration_id) as m from tbl_member_registration", cn, adOpenStatic, adLockPessimistic
                RegId = rsMemberReg.Fields("m")
                For i = 1 To flxGrid1.Rows - 2
              
                    cn.Execute "insert into tbl_MemberPgm(MemberReg_Id,Program_Id,DistLevelPrizeCat_Id) values(" & RegId & _
                    "," & Val(flxGrid1.TextMatrix(i, 3)) & "," & Val(flxGrid1.TextMatrix(i, 4)) & ")"
                Next
            Else
                If MsgBox("Already Exists,Do U want to modify", vbYesNo + vbQuestion, "Warning") = vbYes Then
                    If rsMemberReg.State = 1 Then rsMemberReg.Close
                    If rsMemberPgm.State = 1 Then rsMemberPgm.Close
                    cn.Execute "update tbl_member_registration set name='" & Cmbname.Text & "',address='" & txtAddress.Text & "',contact_no='" & _
                    txtContactNo.Text & "',dob='" & dtpDOB.Value & "',gender='" & gen & "',member_category='" & Cat & "',g_name='" & _
                    txtGname.Text & "',relationship='" & txtRelationship.Text & "' where member_registration_id=" & Cmbname.ItemData(Cmbname.ListIndex)
                    RegId = Cmbname.ItemData(Cmbname.ListIndex)
                    If rsMemberReg.State = 1 Then rsMemberReg.Close
                    If rsMemberPgm.State = 1 Then rsMemberPgm.Close
                    cn.Execute "delete from tbl_memberpgm where MemberReg_Id=" & RegId
                     For i = 1 To flxGrid1.Rows - 2
                        cn.Execute "insert into tbl_MemberPgm(MemberReg_Id,Program_Id,DistLevelPrizeCat_Id) values(" & RegId & _
                        "," & Val(flxGrid1.TextMatrix(i, 3)) & "," & Val(flxGrid1.TextMatrix(i, 4)) & ")"
                    Next
                End If
            End If
        Else
            cn.Execute "insert into tbl_member_registration(school_id,name,address,contact_no,dob,gender,member_category,g_name,relationship) values(" & _
            cmbSchool.ItemData(cmbSchool.ListIndex) & ",'" & Cmbname.Text & "','" & txtAddress.Text & "','" & txtContactNo.Text & "','" & dtpDOB.Value & _
            "','" & gen & "','" & Cat & "','" & txtGname.Text & "','" & txtRelationship.Text & "')"
            If rsMemberReg.State = 1 Then rsMemberReg.Close
            rsMemberReg.Open "Select max(member_registration_id) as m from tbl_member_registration", cn, adOpenStatic, adLockPessimistic
            RegId = rsMemberReg.Fields("m")
            For i = 1 To flxGrid1.Rows - 2
                cn.Execute "insert into tbl_MemberPgm(MemberReg_Id,Program_Id,DistLevelPrizeCat_Id) values(" & RegId & _
                "," & Val(flxGrid1.TextMatrix(i, 3)) & "," & Val(flxGrid1.TextMatrix(i, 4)) & ")"
            Next
        End If
        Clearing
        cmbSchool.Text = ""
        cmbSchool.ListIndex = -1
        Cmbname.Text = ""
        Cmbname.ListIndex = -1
        cmbDistrict.ListIndex = -1
        cmbDistrict.Text = ""
        cmbProgram.ListIndex = -1
        cmbProgram.Text = ""
        cmbPrize.ListIndex = -1
        cmbPrize.Text = ""
    Else
        MsgBox "Select School,Enter Name and Programs", vbInformation, "Message"
    End If
Exit Sub
lbl:
    MsgBox Err.Description, vbInformation, "Message"

End Sub

Private Sub dtpDOB_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then optMale.SetFocus
End Sub

Private Sub Form_Load()
    AddDistrict
    AddProgram
    AddPrizeCategory
    DispGrid
End Sub

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

Public Sub AddPrizeCategory()
    If rsPrizeCategory.State = 1 Then rsPrizeCategory.Close
    rsPrizeCategory.Open "Select * from tbl_prize_category", cn, adOpenStatic, adLockPessimistic
    cmbPrize.Clear
    While Not rsPrizeCategory.EOF
        cmbPrize.AddItem rsPrizeCategory.Fields("category_name")
        cmbPrize.ItemData(cmbPrize.NewIndex) = rsPrizeCategory.Fields("prize_category_id")
        rsPrizeCategory.MoveNext
    Wend
End Sub

Public Sub DispGrid()
    flxGrid1.Clear
    flxGrid1.Rows = 2
    flxGrid1.ColWidth(0) = 1000
    flxGrid1.ColWidth(1) = 2500
    flxGrid1.ColWidth(2) = 2500
    flxGrid1.ColWidth(3) = 0
    flxGrid1.ColWidth(4) = 0
    flxGrid1.TextMatrix(0, 0) = "Sl No"
    flxGrid1.TextMatrix(0, 1) = "Program Name"
    flxGrid1.TextMatrix(0, 2) = "Dist. Level Prize Category"
    flxGrid1.TextMatrix(0, 3) = "Program ID"
    flxGrid1.TextMatrix(0, 4) = "PrizeCatId"
    
End Sub

Public Sub Clearing()
    txtAddress.Text = ""
    txtContactNo.Text = ""
    txtGname.Text = ""
    txtRelationship.Text = ""
    optMale.Value = False
    optFemale.Value = False
    optJunior.Value = False
    optSenior.Value = False
End Sub

Private Sub optFemale_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then optJunior.SetFocus
End Sub

Private Sub optJunior_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then cmbProgram.SetFocus
End Sub

Private Sub optMale_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then optJunior.SetFocus
End Sub

Private Sub optSenior_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then cmbProgram.SetFocus
End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtGname.SetFocus

End Sub

'If KeyAscii = 13 Then txtRelationship.SetFocus



Private Sub txtContactNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then dtpDOB.SetFocus
End Sub


Private Sub Txtcontactno_LostFocus()
 If Not ValPhone(txtContactNo.Text) Then
        MsgBox "Not a valid contactno", vbInformation, "Message"
    End If
End Sub

Private Sub txtDescription_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmbPrize.SetFocus
End Sub

Private Sub txtGname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtRelationship.SetFocus
End Sub

Private Sub txtRelationship_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtContactNo.SetFocus
End Sub

Private Sub txtSchoolAddress_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtame.SetFocus


End Sub
