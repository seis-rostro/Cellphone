VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBackUP 
   Appearance      =   0  'Flat
   BorderStyle     =   0  'None
   Caption         =   "Back Up Utility"
   ClientHeight    =   2745
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4290
   Icon            =   "frmBackUP.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   4290
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   0
      Left            =   2715
      TabIndex        =   5
      Top             =   1950
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
      Caption         =   "&Ok"
      AccessKey       =   "O"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmBackUP.frx":0CCA
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   1
      Left            =   3480
      TabIndex        =   6
      Top             =   1950
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
      Picture         =   "frmBackUP.frx":1444
      PicturePos      =   1
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   600
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   1140
      Width           =   4110
      _ExtentX        =   7250
      _ExtentY        =   1058
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   1515
         TabIndex        =   4
         Top             =   120
         Width           =   2430
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Drive Destination"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   150
         Width           =   1260
      End
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   570
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   4110
      _ExtentX        =   7250
      _ExtentY        =   1005
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   1515
         TabIndex        =   2
         Text            =   "Cellphone_POS"
         Top             =   105
         Width           =   2415
      End
      Begin MSComCtl2.Animation Progress 
         Height          =   315
         Left            =   735
         TabIndex        =   1
         TabStop         =   0   'False
         Tag             =   "wt0;fb0"
         Top             =   -45
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   2334974
         FullWidth       =   39
         FullHeight      =   21
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Server"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   0
         Top             =   165
         Width           =   1260
      End
   End
End
Attribute VB_Name = "frmBackUP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  Programmed By  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  Rosalyn Lazo Descallar  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  Date Started  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  April 26, 2007  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'

Option Explicit

Private WithEvents oDriver As FormDriver
Attribute oDriver.VB_VarHelpID = -1
Private oSkin As FormSkin
Private bLoaded As Boolean
Private poFileSys As FileSystemObject

Private Sub cmdButton_Click(Index As Integer)
   Select Case Index
      Case 0
         If txtfield(0).Text <> "" Then
            Back_UP
         Else
            Exit Sub
         End If
      Case 1
         Unload Me
   End Select
End Sub

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0

   If bLoaded = False Then
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

Private Function Back_UP() As Boolean
Dim Reference As String
Dim lsSQL As String
Dim rsTarget As ADODB.Recordset
Dim oFolder As Folder
Dim oFiles As Files
Dim oFile As File
Dim oFileObject As New FileSystemObject
Dim rsSource As ADODB.Recordset
Dim lnrow As Long

'Utility
Dim rsMain As ADODB.Recordset
Dim rsNepo As ADODB.Recordset
Dim Quantity As Integer
Dim Entry As Integer
Dim oRS As New ADODB.Recordset
Dim lrs As New ADODB.Recordset
Dim lnEntry As Integer

Back_UP = True
oApp.Connection.BeginTrans
On Error GoTo errProc


   Set poFileSys = New FileSystemObject
   Set oFolder = oFileObject.GetFolder(Drive1 & "\" & "POS_BackUP\")
   Set oFiles = oFolder.Files

   If Not poFileSys.DriveExists(Drive1) Then
      MsgBox "Drive Does not Exist!!!" & vbCrLf & _
            "Please Insert Mobile Disk then Try again.", vbCritical, "Notice"
      Exit Function
   End If

   Progress.Open App.Path & "\images\BOOKS.AVI"
   Progress.Play

   Set rsSource = New ADODB.Recordset
   rsSource.Open "SELECT Name FROM sysObjects where Type = 'U'", _
                   oApp.Connection, adOpenKeyset, adLockOptimistic, adCmdText

   Do While Not rsSource.EOF

     Reference = Format(Date, "mmddyy") & "_" & rsSource("Name")
     Set rsTarget = New ADODB.Recordset

     rsTarget.Open Trim(rsSource("Name")), oApp.Connection

     txtfield(0).Text = Reference
     DoEvents

        If poFileSys.FileExists(Drive1 & "\" & "POS_BackUP\" & Reference) Then
           poFileSys.DeleteFile Drive1 & "\" & "POS_BackUP\" & Reference
        End If

     rsTarget.Save Drive1 & "\" & "POS_BackUP\" & Reference
     rsTarget.Close

     rsSource.MoveNext
   Loop

      'Delete Prev Back UP, >6 days
      For Each oFile In oFiles
        If DateDiff("D", oFile.DateLastModified, Date) > 6 Then
           oFileObject.DeleteFile oFile.Path
        End If
      Next

      Progress.Stop
      Progress.Close
      MsgBox "Back Up Successfully Completed!!!", vbInformation, "Information"

      rsSource.Close
      Set rsTarget = Nothing
      Set rsSource = Nothing

   MsgBox "tapos na po!!!"

endProc:
   oApp.Connection.CommitTrans
   Exit Function
errProc:
   Progress.Stop
   oApp.Connection.RollbackTrans
   Back_UP = False
   MsgBox Err.Description, vbCritical, "Warning"

End Function

''¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  Date Tested  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
''¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤   April 26, 2007  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'



