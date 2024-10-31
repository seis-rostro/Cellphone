VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmUploadfile 
   BorderStyle     =   0  'None
   Caption         =   "Upload File"
   ClientHeight    =   2760
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2760
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdButton 
      Caption         =   "CLEAR"
      Height          =   405
      Index           =   2
      Left            =   2355
      TabIndex        =   4
      Top             =   2115
      Width           =   1380
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "CLOSE"
      Height          =   405
      Index           =   1
      Left            =   3945
      TabIndex        =   3
      Top             =   2115
      Width           =   1380
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "EXECUTE"
      Height          =   405
      Index           =   0
      Left            =   690
      TabIndex        =   2
      Top             =   2115
      Width           =   1380
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1170
      Left            =   210
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   5580
      _ExtentX        =   9843
      _ExtentY        =   2064
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtFileName 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1350
         TabIndex        =   0
         Top             =   360
         Width           =   3765
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "File Name:"
         Height          =   300
         Left            =   420
         TabIndex        =   1
         Top             =   390
         Width           =   885
      End
      Begin VB.Shape Shape1 
         Height          =   705
         Left            =   270
         Top             =   225
         Width           =   5010
      End
   End
End
Attribute VB_Name = "frmUploadfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pxeMODULENAME = "frmUploadfile"
Private oSkin As clsFormSkin
Dim psFileName As String

Private Sub cmdButton_Click(Index As Integer)
   Select Case Index
    Case 0
        psFileName = txtFileName.Text
        Unload Me
    Case 1
        Unload Me
        psFileName = ""
    Case 2
        txtFileName.Text = ""
        psFileName = ""
    End Select
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String

   lsOldProc = "Form_Load"
   ''On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormMaintenance
   psFileName = ""

endProc:
   Exit Sub
errProc:
End Sub

Property Get FileName() As String
   FileName = psFileName
End Property
