VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmLoad_Tran 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2625
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4110
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2625
   ScaleWidth      =   4110
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1035
      Index           =   1
      Left            =   150
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   1826
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   0
         Left            =   1335
         MaxLength       =   50
         TabIndex        =   2
         Top             =   120
         Width           =   2220
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   2
         Left            =   1335
         TabIndex        =   1
         Top             =   630
         Width           =   2220
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   1335
         MaxLength       =   50
         TabIndex        =   0
         Top             =   375
         Width           =   2220
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Trace No."
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   165
         Width           =   1260
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         Height          =   285
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   645
         Width           =   900
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Reference No."
         Height          =   285
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   420
         Width           =   1260
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   1
      Left            =   3225
      TabIndex        =   6
      Top             =   1845
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
      Picture         =   "frmLoad_Trans.frx":0000
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   0
      Left            =   2445
      TabIndex        =   7
      Top             =   1845
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
      Picture         =   "frmLoad_Trans.frx":077A
      PicturePos      =   1
   End
End
Attribute VB_Name = "frmLoad_Tran"
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

With frmCP_POS
   Select Case Index
      Case 0 'OK
         .txtothers(2).Text = Format(txtField(2).Text, "#,##0.00")
         .txtothers(3).Text = 1
         .txtothers(5).Text = Format(txtField(2).Text, "#,##0.00")
         .txtField(3).Text = Format(txtField(2).Text, "#,##0.00")
         
         ShowGrid
         Unload Me
         .txtField(4).SetFocus
         
      Case 1 'Cancel
         Unload Me
         .txtothers(0).SetFocus
   End Select
   
End With
   
End Sub

Private Sub ShowGrid()
Dim temp As Integer
   
With frmCP_POS

   .MSFlexGrid1.TextMatrix(1, 1) = .txtothers(0).Text
   .MSFlexGrid1.TextMatrix(1, 2) = .txtothers(1).Text
   .MSFlexGrid1.TextMatrix(1, 3) = .txtothers(2).Text
   .MSFlexGrid1.TextMatrix(1, 4) = 1
   .MSFlexGrid1.TextMatrix(1, 5) = 0
   .MSFlexGrid1.TextMatrix(1, 6) = 0
   .MSFlexGrid1.TextMatrix(1, 7) = Format(txtField(2).Text, "#,##0.00")
   .MSFlexGrid1.TextMatrix(1, 8) = .txtothers(0).Tag
   .MSFlexGrid1.TextMatrix(1, 9) = txtField(0).Text
   .MSFlexGrid1.TextMatrix(1, 10) = txtField(1).Text
   
End With

End Sub

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0

   If bLoaded = False Then
      bLoaded = True
      txtField(0).SetFocus
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
   
   
   txtField(0).Tag = frmCP_POS.txtothers(0).Tag
   txtField(2).Text = "0.00"
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oDriver = Nothing
   Set oSkin = Nothing
End Sub

Private Sub txtfield_GotFocus(Index As Integer)
   oDriver.ColumnIndex = Index
End Sub

Private Sub txtField_LostFocus(Index As Integer)

If Index = 2 Then
   If Not IsNumeric(txtField(Index).Text) Then
      txtField(Index).Text = 0#
      txtField(Index).SetFocus
   Else
      txtField(Index).Text = Format(txtField(Index).Text, "#,##0.00")
   End If
End If

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




