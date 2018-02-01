VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPrizeCategory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Prize Category"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5700
   Icon            =   "frmPrizeCategory.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   5700
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   4515
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5655
      Begin MSFlexGridLib.MSFlexGrid flxGrid 
         Height          =   2235
         Left            =   60
         TabIndex        =   5
         Top             =   2220
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   3942
         _Version        =   393216
         Cols            =   3
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   375
         Left            =   2940
         TabIndex        =   1
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   4260
         TabIndex        =   3
         Top             =   1800
         Width           =   1215
      End
      Begin VB.ComboBox cmbCategory 
         Height          =   1545
         Left            =   1380
         Style           =   1  'Simple Combo
         TabIndex        =   0
         Top             =   180
         Width           =   4095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Category Name"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmPrizeCategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbCategory_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then cmdSave.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo lbl
    If Len(Trim(cmbCategory.Text)) <> 0 Then
        If cmbCategory.ListIndex <> -1 Then
            If rsPrizeCategory.State = 1 Then rsPrizeCategory.Close
            rsPrizeCategory.Open "select * from tbl_prize_category where prize_category_id=" & cmbCategory.ItemData(cmbCategory.ListIndex), cn, adOpenStatic, adLockPessimistic
            If rsPrizeCategory.EOF = True Then
                cn.Execute "insert into tbl_prize_category(category_name) values('" & cmbCategory.Text & "')"
            Else
                If MsgBox("Already exists,Do U want to modify", vbYesNo + vbQuestion, "Warning") = vbYes Then
                    If rsPrizeCategory.State = 1 Then rsPrizeCategory.Close
                    cn.Execute "update tbl_prize_category set category_name='" & cmbCategory.Text & "' where prize_category_id=" & cmbCategory.ItemData(cmbCategory.ListIndex)
                End If
            End If
            
        Else
            cn.Execute "insert into tbl_prize_category(category_name) values('" & cmbCategory.Text & "')"
        End If
        cmbCategory.ListIndex = -1
        cmbCategory.Text = ""
        cmbCategory.SetFocus
        AddCategory
        DispGrid
    Else
        MsgBox "Enter prizecategory", vbInformation, "Message"
     End If
        Exit Sub
lbl:
    MsgBox Err.Description, vbInformation, "Message"
   
End Sub


Public Sub AddCategory()
    If rsPrizeCategory.State = 1 Then rsPrizeCategory.Close
    rsPrizeCategory.Open "select * from tbl_prize_category", cn, adOpenStatic, adLockPessimistic
    cmbCategory.Clear
    While Not rsPrizeCategory.EOF
        cmbCategory.AddItem rsPrizeCategory.Fields("category_name")
        cmbCategory.ItemData(cmbCategory.NewIndex) = rsPrizeCategory.Fields("prize_category_id")
        rsPrizeCategory.MoveNext
    Wend
End Sub

Private Sub flxGrid_DblClick()
On Error GoTo lbl
    Dim id As Integer, i As Integer
    id = flxGrid.TextMatrix(flxGrid.RowSel, 2)
    For i = 0 To cmbCategory.ListCount - 1
        If cmbCategory.ItemData(i) = id Then
            cmbCategory.ListIndex = i
            Exit For
          End If
    Next
        Exit Sub
lbl:
    MsgBox Err.Description, vbInformation, "Message"
      
End Sub

Private Sub Form_Load()
    AddCategory
    DispGrid
End Sub

Public Sub DispGrid()
    flxGrid.Clear
    flxGrid.Rows = 2
    
    flxGrid.ColWidth(0) = 1000
    flxGrid.ColWidth(1) = 3000
    flxGrid.ColWidth(2) = 0
    
    flxGrid.TextMatrix(0, 0) = "Sl No"
    flxGrid.TextMatrix(0, 1) = "Prize Category"
    flxGrid.TextMatrix(0, 2) = "Id"
    If rsPrizeCategory.State = 1 Then rsPrizeCategory.Close
    rsPrizeCategory.Open "select * from tbl_prize_category", cn, adOpenStatic, adLockPessimistic
    While Not rsPrizeCategory.EOF
        flxGrid.TextMatrix(flxGrid.Rows - 1, 0) = flxGrid.Rows - 1
        flxGrid.TextMatrix(flxGrid.Rows - 1, 1) = rsPrizeCategory.Fields("Category_name")
        flxGrid.TextMatrix(flxGrid.Rows - 1, 2) = rsPrizeCategory.Fields("prize_Category_id")
        rsPrizeCategory.MoveNext
        flxGrid.Rows = flxGrid.Rows + 1
    Wend

End Sub
