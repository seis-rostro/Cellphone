VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmLoad_Inventory 
   BorderStyle     =   0  'None
   Caption         =   "Load Inventory"
   ClientHeight    =   2280
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4125
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   4125
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   1
      Left            =   3225
      TabIndex        =   5
      Top             =   1500
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmLoad_Inventory.frx":0000
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   0
      Left            =   2445
      TabIndex        =   4
      Top             =   1500
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
      Caption         =   "&OK"
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
      Picture         =   "frmLoad_Inventory.frx":077A
      PicturePos      =   1
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   735
      Index           =   1
      Left            =   150
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   1296
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   2
         Left            =   1335
         TabIndex        =   3
         Top             =   360
         Width           =   2220
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   1335
         MaxLength       =   50
         TabIndex        =   1
         Top             =   105
         Width           =   2220
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         Height          =   285
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   375
         Width           =   900
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Reference No."
         Height          =   285
         Index           =   2
         Left            =   240
         TabIndex        =   0
         Top             =   150
         Width           =   1260
      End
   End
End
Attribute VB_Name = "frmLoad_Inventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents oDriver As FormDriver
Attribute oDriver.VB_VarHelpID = -1
Private oSkin As FormSkin
Private bLoaded As Boolean
Private pbnewitem As Boolean

Private Sub cmdButton_Click(Index As Integer)
Dim lnctr As Integer

With frmInventory
   Select Case Index
      Case 0 'OK
         .txtfield(7).Tag = txtfield(1).Text
         .txtothers(2).Tag = CDbl(txtfield(2).Text)
         .txtothers(2).Text = Format(CDbl(txtfield(2).Text), "#,##0.00")
         .txtothers(7).Text = Format(CDbl(txtfield(2).Text), "#,##0.00")
         
         Unload Me
         .txtfield(4).SetFocus
         
      Case 1 'Cancel
         Unload Me
   End Select
   
End With
   
End Sub
Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0

   If bLoaded = False Then
      bLoaded = True
      txtfield(1).SetFocus
   End If

End Sub

Private Sub Form_Load()
Dim lsSQL As String

   CenterChildForm mdiMain, Me
   bLoaded = False
   
   Set oDriver = New FormDriver
   Set oDriver.AppDriver = oApp
   Set oDriver.MainForm = Me
   
   Set oSkin = New FormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin
   
   txtfield(2).Text = "0.00"
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oDriver = Nothing
   Set oSkin = Nothing
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   oDriver.ColumnIndex = Index
   txtfield(Index).BackColor = &HE1FEFF
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   If Index = 2 Then
      If Not IsNumeric(txtfield(Index).Text) Then
         txtfield(Index).Text = 0#
         txtfield(Index).SetFocus
      Else
         txtfield(Index).Text = Format(txtfield(Index).Text, "#,##0.00")
      End If
   End If
   txtfield(Index).BackColor = &HFFFFFF
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
