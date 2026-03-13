VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmExport_Branch 
   BorderStyle     =   0  'None
   Caption         =   "Export Utility"
   ClientHeight    =   2745
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5970
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2745
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   0
      Left            =   4365
      TabIndex        =   7
      Top             =   1965
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
      Caption         =   "&Export"
      AccessKey       =   "E"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmExport_Branch.frx":0000
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   1
      Left            =   5130
      TabIndex        =   8
      Top             =   1965
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
      Caption         =   "&Close"
      AccessKey       =   "C"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmExport_Branch.frx":077A
      PicturePos      =   1
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   495
      Left            =   1770
      Tag             =   "wt0;fb0"
      Top             =   1245
      Width           =   4110
      _ExtentX        =   7250
      _ExtentY        =   873
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   1515
         TabIndex        =   5
         Top             =   60
         Width           =   2430
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Drive Destination"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   90
         Width           =   1260
      End
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   675
      Left            =   1770
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   4110
      _ExtentX        =   7250
      _ExtentY        =   1191
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   1515
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   345
         Width           =   2415
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   0
         Left            =   1515
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   75
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Reference No."
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1260
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Reference Date"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   2
         Top             =   90
         Width           =   1260
      End
   End
   Begin MSComCtl2.Animation Progress 
      Height          =   720
      Left            =   105
      TabIndex        =   6
      TabStop         =   0   'False
      Tag             =   "wt0;wb0"
      Top             =   1965
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   1270
      _Version        =   393216
      BackColor       =   1842204
      FullWidth       =   61
      FullHeight      =   48
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please Be "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   16
      Left            =   135
      TabIndex        =   0
      Top             =   585
      Width           =   930
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Very Accurate..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   0
      Left            =   315
      TabIndex        =   1
      Top             =   795
      Width           =   1395
   End
   Begin VB.Image Image1 
      Height          =   1200
      Left            =   90
      Picture         =   "frmExport_Branch.frx":0EF4
      Stretch         =   -1  'True
      Top             =   540
      Width           =   1635
   End
End
Attribute VB_Name = "frmExport_Branch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  Programmed By  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  Rosalyn Lazo Descallar  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  Date Started  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  July 03, 2007  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'

Option Explicit

Private WithEvents oDriver As FormDriver
Attribute oDriver.VB_VarHelpID = -1
Private oSkin As FormSkin
Private bLoaded As Boolean
Private poFileSys As FileSystemObject

Private pnCtr As Integer
Dim oRS As New ADODB.Recordset
Dim lrs As New ADODB.Recordset
Dim rRs As New ADODB.Recordset

Private Sub cmdButton_Click(Index As Integer)
   Select Case Index
      Case 0
         If txtfield(0).Text <> "" Then
            Export_Data
         Else
            MsgBox "Input Reference Date.!!!", vbCritical, "Warning"
         End If
      Case 1
         Unload Me
   End Select
End Sub

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0

   If bLoaded = False Then
      txtfield(0).Text = Format(oApp.ServerDate, "MMMM dd, yyyy")
      txtfield(1).Text = Trim(oApp.BranchCode & "-" & Format(txtfield(0).Text, "MMDDYY") _
                         & Format(Now, "HHNNSS"))
      txtfield(1).Enabled = False
      bLoaded = True
   End If
End Sub

Private Sub Form_Load()
Dim Table As String

   CenterChildForm mdiMain, Me
   bLoaded = False

   Set oDriver = New FormDriver
   Set oDriver.AppDriver = oApp
   Set oDriver.MainForm = Me

   Set oSkin = New FormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin

End Sub

Private Function Export_Data() As Boolean
Dim lrsTarget As ADODB.Recordset
Dim rsTarget As ADODB.Recordset
Dim rsSource As ADODB.Recordset
Dim rsUser As ADODB.Recordset
Dim Reference As String
Dim lsSQL As String
Dim lnrow As Long


Export_Data = True
oApp.Connection.BeginTrans
On Error GoTo errProc
      
Set poFileSys = New FileSystemObject

   If Not poFileSys.DriveExists(Drive1) Then
      MsgBox "Drive Does not Exist!!!" & vbCrLf & _
            "Please Insert Mobile Disk then Try again.", vbCritical, "Notice"
      Exit Function
   End If

   Reference = txtfield(1).Text
   
   'Search for last export date
   Set oRS = New ADODB.Recordset
   lsSQL = "SELECT * " _
         & " FROM xxxExportDate " _
         & " WHERE sBranchCd = '" & oApp.BranchCode & "'" _
         & " ORDER by dTransact Desc "
   If oRS.State = adStateOpen Then oRS.Close
   oRS.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   
   Progress.Open App.Path & "\images\FILECOPY.AVI"
   Progress.Play
   
   If oRS.RecordCount = 0 Then
      MsgBox "Export Table!!!" & vbCrLf & vbCrLf & _
      "Notify ROSALYN LAZO DESCALLAR for Assistance", vbCritical, "Warning"
      Export_Data = False
      Exit Function
   End If
   
   Set rsSource = New ADODB.Recordset
   rsSource.Open "SELECT Name FROM sysObjects where Type = 'U'" _
                  & " AND Name <> 'dtproperties' " _
                  & " AND left(Name,2) <> 'xx' " _
                  & " AND left(name,2) <> 'PO' " _
                  & " AND Name <> 'CP_Serial_Dummy' ", _
                   oApp.Connection, adOpenKeyset, adLockOptimistic, adCmdText
   
   Do While Not rsSource.EOF
      Set lrs = New ADODB.Recordset
      lsSQL = "SELECT * " _
            & " FROM " & rsSource("Name") & "  "
      
      Select Case rsSource("Name")
         Case "CP_Serial_Transfer_Master", "CP_Transfer_Master"
            lsSQL = lsSQL & " WHERE dReceived > '" & CDate(oRS("dTransact")) & "'" _
                        & " OR dModified > '" & CDate(oRS("dTransact")) & "'" _
                        & " OR dModified > '" & CDate(txtfield(0).Text) & "'"
         Case "CP_SO_Master"
            lsSQL = lsSQL & " WHERE dCancelxx > '" & CDate(oRS("dTransact")) & "'" _
                        & " OR dModified > '" & CDate(oRS("dTransact")) & "'" _
                        & " OR dModified > '" & CDate(txtfield(0).Text) & "'"
         Case Else
            lsSQL = lsSQL & " WHERE dModified > '" & CDate(oRS("dTransact")) & "'" _
                        & " OR dModified > '" & CDate(txtfield(0).Text) & "'"
      End Select
      If lrs.State = adStateOpen Then lrs.Close
      lrs.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
      
      If lrs.RecordCount <> 0 Then
         Set lrsTarget = New ADODB.Recordset
         lrsTarget.Open lrs
         lrsTarget.Save Drive1 & "\" & Reference & "_" & rsSource("Name")
         lrsTarget.Close
      End If
      rsSource.MoveNext
   Loop
   
   Set lrsTarget = Nothing
      
   Set rsUser = New ADODB.Recordset
   lsSQL = "SELECT * " _
         & " FROM xxxSysUser "
   If rsUser.State = adStateOpen Then rsUser.Close
   rsUser.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText

   Set lrsTarget = New ADODB.Recordset
   lrsTarget.Open rsUser
   lrsTarget.Save Drive1 & "\" & Reference & "_" & "xxxSysUser"
   lrsTarget.Close

   'Update Export Date
   lsSQL = "UPDATE xxxExportDate SET " _
            & " dTransact = '" & oApp.ServerDate & "', " _
            & " sModified = '" & Encrypt(oApp.UserID) & "' ," _
            & " dModified = getdate() " _
      & " WHERE sbranchcd = '" & oApp.BranchCode & "' "
   oApp.Connection.Execute lsSQL, lnrow, adCmdText
   
   Progress.Stop
   Progress.Close
   MsgBox "Export Successfully Completed!!!", vbInformation, "Information"

   Set oRS = Nothing
   Set lrs = Nothing
   Set lrsTarget = Nothing
   Set rsUser = Nothing
   
endProc:
   oApp.Connection.CommitTrans
   Exit Function
errProc:
   oApp.Connection.RollbackTrans
   Export_Data = False
   MsgBox Err.Description, vbCritical, "Warning"

End Function

Private Sub txtField_GotFocus(Index As Integer)
   txtfield(0).SelStart = 0
   txtfield(0).SelLength = Len(txtfield(0).Text)
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   txtfield(1).Text = Trim(oApp.BranchCode & "-" & Format(txtfield(0).Text, "MMDDYY") _
                      & Format(Now, "HHNNSS"))
End Sub

Private Sub txtfield_Validate(Index As Integer, Cancel As Boolean)
   If Not IsDate(txtfield(0).Text) Then
      txtfield(0).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
   ElseIf DateDiff("d", txtfield(0).Text, Date) < 0 Then
      MsgBox "Forward Date Not Permitted!!!", vbCritical, "Warning"
      txtfield(0).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
      txtfield(0).SetFocus
   Else
      txtfield(0).Text = Format(txtfield(0).Text, "MMMM DD, YYYY")
      txtfield(1).Text = Trim(oApp.BranchCode & "-" & Format(txtfield(0).Text, "MMDDYY") _
                   & Format(Now, "HHNNSS"))
   End If
End Sub

''¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  Date Tested  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
''¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  July 04, 2007  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'
''¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤    Version 1    ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
''¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  Date Finished  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
''¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  July 04, 2007  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'

