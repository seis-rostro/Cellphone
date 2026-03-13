VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmWallet_POS 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Load Wallet POS"
   ClientHeight    =   1695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3990
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   1695
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrButton cmdButton 
      Height          =   375
      Index           =   1
      Left            =   2820
      TabIndex        =   7
      Top             =   1260
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
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
      Picture         =   "frmWallet_POS.frx":0000
      BackColor       =   14286077
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   375
      Index           =   0
      Left            =   1710
      TabIndex        =   6
      Top             =   1260
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
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
      Picture         =   "frmWallet_POS.frx":077A
      BackColor       =   14286077
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1020
      Index           =   1
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   90
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   1799
      BackColor       =   7716603
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   1335
         MaxLength       =   50
         TabIndex        =   3
         Top             =   375
         Width           =   2220
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   2
         Left            =   1335
         TabIndex        =   5
         Top             =   630
         Width           =   2220
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   0
         Left            =   1335
         TabIndex        =   1
         Top             =   120
         Width           =   2220
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Reference No."
         Height          =   285
         Index           =   2
         Left            =   240
         TabIndex        =   2
         Top             =   405
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
         Caption         =   "Cellphone No."
         Height          =   285
         Index           =   3
         Left            =   240
         TabIndex        =   0
         Top             =   135
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmWallet_POS"
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
Private psSelected() As String
Dim temp As Integer

Private Sub cmdButton_Click(Index As Integer)
Dim lnctr As Integer

With frmCP_POS
   Select Case Index
      Case 0 'OK
         If txtField(0).Text <> "" And txtField(1).Text <> "" And _
            txtField(2).Text <> "0.00" Then
            .txtothers(2).Text = Format(txtField(2).Text, "#,##0.00")
            .txtothers(5).Text = Format(txtField(2).Text, "#,##0.00")
            .txtField(2).Text = Format(txtField(2).Text, "#,##0.00")
            
            ShowGrid
            Unload Me
         Else
            Check_Input
         End If
         
      Case 1 'Cancel
         Unload Me
         .txtothers(0).Text = ""
         .txtothers(0).Tag = ""
         .txtothers(1).Text = ""
         .txtothers(0).SetFocus
   End Select
   
End With
   
End Sub

Private Sub Grand_Total()
Dim Total As Double
Dim lnctr As Integer

With frmCP_POS.MSFlexGrid1

   If .Rows <= 3 Then
      If .Rows = 2 Then
         frmCP_POS.txtField(2).Text = "0.00"
      Else
         frmCP_POS.txtField(2).Text = Format(.TextMatrix(1, 9), "#,##0.00")
         Total = txtField(2).Text
      End If
   Else
      For lnctr = 1 To .Rows - 2
         Total = Total + CDbl(.TextMatrix(lnctr, 9))
      Next
         frmCP_POS.txtField(2).Text = Format(CDbl(Total), "#,##0.00")
   End If

End With

End Sub

Private Sub ShowGrid()

   With frmCP_POS.MSFlexGrid1
                  
      .Rows = .Rows + 1
      
      If .Row = 1 Then
         .TextMatrix(.Row, 0) = 1
      Else
         .TextMatrix(.Row, 0) = .TextMatrix(.Row - 1, 0) + 1
      End If
      
      .TextMatrix(.Row, 1) = frmCP_POS.txtothers(0).Text
      .TextMatrix(.Row, 2) = frmCP_POS.txtothers(1).Text
      .TextMatrix(.Row, 3) = Format(txtField(2).Text, "#,##0.00")
      .TextMatrix(.Row, 4) = 1
      .TextMatrix(.Row, 5) = 0
      .TextMatrix(.Row, 6) = 0
      .TextMatrix(.Row, 7) = txtField(1).Text
      .TextMatrix(.Row, 8) = frmCP_POS.txtothers(0).Tag
      .TextMatrix(.Row, 9) = Format(txtField(2).Text, "#,##0.00")
      .TextMatrix(.Row, 10) = ""
      .TextMatrix(.Row, 11) = txtField(2).Text
      .TextMatrix(.Row, 12) = frmCP_POS.txtothers(1).Tag
      .TextMatrix(.Row, 13) = txtField(0).Text
            
      .Row = .Rows - 1
         
   End With

   Grand_Total

End Sub

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0

   If bLoaded = False Then
      bLoaded = True
      txtField(2).Text = "0.00"
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
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oDriver = Nothing
   Set oSkin = Nothing
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   
   oDriver.ColumnIndex = Index
   
   If txtField(Index).Text <> "" Then
      txtField(Index).SelStart = 0
      txtField(Index).SelLength = Len(txtField(Index).Text)
   End If

End Sub


Private Sub txtField_LostFocus(Index As Integer)

If Index = 2 Then
   If Not IsNumeric(txtField(Index).Text) Then
      txtField(Index).Text = 0#
      txtField(Index).SetFocus
   Else
      txtField(Index).Text = Format(txtField(Index).Text, "#,##0.00")
      cmdButton(0).SetFocus
   End If
End If

End Sub

Private Sub Check_Input()

   If txtField(0).Text = "" Then
      MsgBox "Invalid Phone Number Detected!!!", vbCritical, "Warning"
      txtField(0).SetFocus
   ElseIf txtField(1).Text = "" Then
      MsgBox "Invalid Input!!!", vbCritical, "Warning"
      txtField(1).SetFocus
   ElseIf txtField(2).Text = "0.00" Then
      MsgBox "Invalid Amount Detected!!!", vbCritical, "Warning"
      txtField(2).SetFocus
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

