Attribute VB_Name = "modMain"
Public oAppDriver As AppDriver
Public Declare Function GetFocus Lib "user32" () As Long

Private Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Private Sub Main()
10       Dim lsCommand As String
20       Dim lasParam() As String
   
30       On Error GoTo errProc
   
40       lsCommand = Command()
50       lasParam = Split(lsCommand)
   
60       Set oAppDriver = New AppDriver
70       If oAppDriver.LoadEnv(lasParam(0), lasParam(1)) = False Then
80          Exit Sub
90       End If
   
100      frmUserMaintenance.Show

endProc:
110      Exit Sub
errProc:
120      MsgBox "Line No:" & Erl & vbCrLf & Err.Description, vbCritical, "Error"
130      End
End Sub

Public Sub SetNextFocus()
10       keybd_event &H9, 0, 0, 0
20       keybd_event &H9, 0, &H2, 0
End Sub

Public Sub SetPreviousFocus()
10       keybd_event &H10, 0, 0, 0
20       keybd_event &H9, 0, 0, 0
30       keybd_event &H10, 0, &H2, 0
End Sub

Public Sub CenterChildForm(frmMDIForm As MDIForm, frmChild As Form)
10       Dim lbX As Long, lbY As Long
   
20       lbX = frmMDIForm.ScaleWidth
30       lbY = frmMDIForm.ScaleHeight
   
40       frmChild.Left = CLng((lbX - frmChild.Width) / 2)
50       frmChild.Top = CLng((lbY - frmChild.Height) / 2)
   
60       If frmChild.Left < 0 Then frmChild.Left = 0
70       If frmChild.Top < 0 Then frmChild.Top = 0
End Sub

Public Function TransStat(nStat As Integer) As String
10       Select Case nStat
   Case 0
20          TransStat = "Open"
30       Case 1
40          TransStat = "Closed"
50       Case 2
60          TransStat = "Posted"
70       Case 3
80          TransStat = "Cancelled"
90       Case 4
100         TransStat = "Unknown"
110      End Select
End Function

