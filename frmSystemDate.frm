VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmSystemDate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System Date"
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4275
   Icon            =   "frmSystemDate.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   4275
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1635
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   4155
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   3000
         TabIndex        =   6
         Top             =   1140
         Width           =   975
      End
      Begin VB.CommandButton cmdChange 
         Caption         =   "Chan&ge"
         Height          =   375
         Left            =   1980
         TabIndex        =   5
         Top             =   1140
         Width           =   975
      End
      Begin VB.TextBox txtSystemDate 
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   180
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         Top             =   660
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16580609
         CurrentDate     =   42335
      End
      Begin VB.Label Label2 
         Caption         =   "Select New Date"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Current Date Is"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   180
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmSystemDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdChange_Click()
    If MsgBox("Are You sure to change Ssytem Date", vbYesNo + vbQuestion, "Warning") = vbYes Then
        Date = DTPicker1.Value
        txtSystemDate.Text = Format(Date, "mm/dd/yy")
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then cmdChange.SetFocus
End Sub

Private Sub Form_Load()
    txtSystemDate.Text = Format(Date, "mm/dd/yy")
    DTPicker1.Value = Date
End Sub



Private Sub txtSystemDate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then DTPicker1.SetFocus
End Sub
