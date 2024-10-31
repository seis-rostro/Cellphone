VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmTranStat 
   BorderStyle     =   0  'None
   Caption         =   "Select Order Status"
   ClientHeight    =   2775
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3600
   LinkTopic       =   "Form1"
   ScaleHeight     =   2775
   ScaleWidth      =   3600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin xrControl.xrButton cmdButton 
      Height          =   435
      Index           =   1
      Left            =   2235
      TabIndex        =   0
      Top             =   960
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   767
      Caption         =   "&Cancel"
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
      Picture         =   "frmTranStat.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Cancel          =   -1  'True
      Height          =   435
      Index           =   0
      Left            =   2235
      TabIndex        =   1
      Top             =   495
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   767
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
      Picture         =   "frmTranStat.frx":077A
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   2130
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   3757
      Begin VB.CheckBox Check1 
         Caption         =   "All"
         Height          =   285
         Index           =   4
         Left            =   330
         TabIndex        =   6
         Tag             =   "wt0;fb0"
         Top             =   315
         Width           =   1170
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Cancelled"
         Height          =   285
         Index           =   3
         Left            =   330
         TabIndex        =   5
         Tag             =   "wt0;fb0"
         Top             =   1575
         Width           =   1170
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Posted"
         Height          =   285
         Index           =   2
         Left            =   330
         TabIndex        =   4
         Tag             =   "wt0;fb0"
         Top             =   1260
         Width           =   1170
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Closed"
         Height          =   285
         Index           =   1
         Left            =   330
         TabIndex        =   3
         Tag             =   "wt0;fb0"
         Top             =   945
         Width           =   1170
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Open"
         Height          =   285
         Index           =   0
         Left            =   330
         TabIndex        =   2
         Tag             =   "wt0;fb0"
         Top             =   630
         Width           =   1170
      End
   End
End
Attribute VB_Name = "frmTranStat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private p_oSkin As clsFormSkin
Private p_oAppDrivr As clsAppDriver

Private pnTranStat As Integer
Private pbCancelled As Boolean

Property Set AppDriver(oAppDriver As clsAppDriver)
   Set p_oAppDrivr = oAppDriver
End Property

Property Get Cancelled() As Boolean
   Cancelled = pbCancelled
End Property

Property Get TranStatus() As Integer
   TranStatus = pnTranStat
End Property

Private Sub Check1_Click(Index As Integer)
   If Index = 4 Then
      Check1(0).Value = Check1(Index).Value
      Check1(1).Value = Check1(Index).Value
      Check1(2).Value = Check1(Index).Value
      Check1(3).Value = Check1(Index).Value
   End If
End Sub


Private Sub cmdButton_Click(Index As Integer)
   Dim lnCtr As Integer
   Dim lsVal As String
   
   lsVal = ""
   For lnCtr = 0 To 3
      If Check1(lnCtr).Value = 1 Then
         lsVal = Str(lnCtr) + lsVal
      End If
   Next
   pnTranStat = Val(IIf(lsVal = "", 4, lsVal))
   
   pbCancelled = Index = 1
   Me.Hide
End Sub

Private Sub Form_Load()
   Set p_oSkin = New clsFormSkin
   Set p_oSkin.AppDriver = p_oAppDrivr
   Set p_oSkin.Form = Me
   p_oSkin.ApplySkin xeFormTransEqualRight
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set p_oSkin = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyReturn, vbKeyUp, vbKeyDown
      Select Case KeyCode
      Case vbKeyReturn, vbKeyDown
         SetNextFocus
      Case vbKeyUp
         SetPreviousFocus
      End Select
   End Select
End Sub

