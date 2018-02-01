Attribute VB_Name = "Module1"
Public cn As New Connection
Public rsDistrict As New Recordset
Public rsSchool As New Recordset
Public rsstage As New Recordset
Public rsVolunteer As New Recordset
Public rsProgram As New Recordset
Public rsJudges As New Recordset
Public rsGreenRoom As New Recordset
Public rsPrizeCategory As New Recordset
Public rsPrize As New Recordset
Public rsMemberReg As New Recordset
Public rsMemberPgm As New Recordset
Public rspschedule As New Recordset
Public rsScore As New Recordset
Public rsPgmTiming As New Recordset
Public MenuItem As String
Public rsprofile As New Recordset
Public rsNewUser As New Recordset
Public Sub NumCheck(KeyAscii As Integer)
    If KeyAscii <> Asc(vbBack) And KeyAscii <> Asc(".") Then
        If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
            KeyAscii = 0
        End If
    End If
End Sub


Public Function ValEmail(mail As String)
    If Left(mail, 1) = "@" Then
GoTo notvalide
ElseIf Right(mail, 1) = "@" Then
    GoTo notvalide
ElseIf InStr(1, mail, "@") = False Then
    MsgBox "The @ is missing!"
ElseIf InStr(1, mail, ".") = False Then
    GoTo notvalide
ElseIf Right(mail, 1) = "." Then
    GoTo notvalide
ElseIf Left(mail, 1) = "." Then
    GoTo notvalide
notvalide:
MsgBox "This is Not a Valid Email Address!"
End If
End Function

Public Function ValPhone(phone As String)
isvalide = phone Like "[789]#########"
If Not isvalide Then
    isvalide = phone Like "[0-9]##########"
End If
ValPhone = isvalide
End Function


