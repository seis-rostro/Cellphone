VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmInstallment_POS 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Installment"
   ClientHeight    =   3150
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6450
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   ScaleHeight     =   3150
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1155
      Index           =   1
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   2037
      BackColor       =   7716603
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   4
         Left            =   4515
         MaxLength       =   50
         TabIndex        =   9
         Top             =   765
         Width           =   1570
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   3
         Left            =   4515
         MaxLength       =   50
         TabIndex        =   7
         Top             =   510
         Width           =   1570
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   2
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   5
         Top             =   765
         Width           =   1570
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   3
         Top             =   510
         Width           =   1570
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   240
         Index           =   0
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   1
         Top             =   135
         Width           =   1570
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Amortization"
         Height          =   270
         Index           =   3
         Left            =   3570
         TabIndex        =   8
         Top             =   750
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Term"
         Height          =   270
         Index           =   0
         Left            =   3570
         TabIndex        =   6
         Top             =   495
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Balance"
         Height          =   270
         Index           =   4
         Left            =   135
         TabIndex        =   4
         Top             =   750
         Width           =   1440
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Down Payment"
         Height          =   270
         Index           =   2
         Left            =   135
         TabIndex        =   2
         Top             =   495
         Width           =   1440
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Amount"
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   0
         Top             =   150
         Width           =   1500
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   450
      Index           =   0
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   1710
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   794
      BackColor       =   7716603
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   5
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   11
         Top             =   90
         Width           =   1570
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cash Given"
         Height          =   285
         Index           =   9
         Left            =   120
         TabIndex        =   10
         Top             =   105
         Width           =   1500
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   1
      Left            =   5250
      TabIndex        =   13
      Top             =   2355
      Width           =   1095
      _ExtentX        =   1931
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
      Picture         =   "frmInstallment.frx":0000
      BackColor       =   14286077
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   0
      Left            =   4140
      TabIndex        =   12
      Top             =   2355
      Width           =   1095
      _ExtentX        =   1931
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
      Picture         =   "frmInstallment.frx":077A
      BackColor       =   14286077
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
End
Attribute VB_Name = "frmInstallment_POS"
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
Dim lnrow As Long
Dim lssql As String

Dim psTransaction As String

Property Let Transaction(Transaction As String)
   psTransaction = Transaction
End Property

Private Sub cmdButton_Click(Index As Integer)
Dim lnCtr As Integer

   Select Case Index
      Case 0 'OK
         Select Case psTransaction
            Case "POS"
               If txtfield(2).Text <> 0# And txtfield(3).Text <> 0# Then
                  If CDbl(txtfield(5).Text) < CDbl(txtfield(1).Text) Then
                     MsgBox "Invalid Cash Amount Given!!!", vbCritical, "Warning"
                     txtfield(5).SetFocus
                  Else
                     frmCP_POS.Payment = "Installment"
                     frmCP_POS.txtfield(6) = "Bal: " & txtfield(2) & " Term: " & txtfield(3) & " Amort.: " & txtfield(4)
                     Me.Hide
                     frmCP_POS.txtfield(4).SetFocus
                  End If
               Else
                  MsgBox "Incomplete Data!!!", vbCritical, "Warning"
               End If
            Case "Change"
               If txtfield(0).Text <> "" And txtfield(5).Text <> "" Then
                  If CDbl(txtfield(5).Text) < CDbl(txtfield(1).Text) Then
                     MsgBox "Invalid Cash Amount Given!!!", vbCritical, "Warning"
                  Else
                     frmChange_Unit.Payment = "Installment"
                     frmChange_Unit.txtfield(5) = "Term: " & txtfield(3) & "Amort.: " & txtfield(4)
                     Me.Hide
                     frmChange_Unit.txtfield(5).SetFocus
                  End If
               Else
                  MsgBox "Incomplete Data!!!", vbCritical, "Warning"
               End If
         End Select
      Case 1 'Cancel
         Select Case psTransaction
            Case "POS"
               frmCP_POS.Payment = ""
            Case "Change"
               frmChange_Unit.Payment = ""
         End Select
         Unload Me
   End Select

End Sub

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0

   If bLoaded = False Then
      bLoaded = True
      oDriver.DisableTextbox 0
      oDriver.DisableTextbox 2
      oDriver.DisableTextbox 4
      txtfield(1).SetFocus
   End If

End Sub

Private Sub Form_Load()
Dim lssql As String

   CenterChildForm mdiMain, Me
   
   bLoaded = False
   
   Set oDriver = New FormDriver
   Set oDriver.AppDriver = oApp
   Set oDriver.MainForm = Me
   
   Set oSkin = New FormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin
    
   txtfield(1).Text = "0.00"
   txtfield(2).Text = "0.00"
   txtfield(3).Text = 0
   txtfield(4).Text = "0.00"
   txtfield(5).Text = "0.00"
       
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
Dim temp As Double
   Select Case Index
      Case 1
         If txtfield(Index).Text <> "" Then
            temp = CDbl(txtfield(0).Text) - CDbl(txtfield(Index).Text)
            txtfield(2).Text = Format(temp, "#,##0.00")
            Amortization
         End If
      Case 3
         Amortization
   End Select
   txtfield(Index).BackColor = &H80000005
End Sub

Private Sub Amortization()
Dim temp As Double
   If txtfield(3).Text <> 0# And txtfield(2).Text <> 0# Then
      temp = CDbl(txtfield(2).Text / txtfield(3).Text)
      txtfield(4).Text = Format(temp, "#,##0.00")
   Else
      txtfield(4).Text = "0.00"
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

Private Sub txtfield_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
   Case 1, 2, 4, 5
      If Not IsNumeric(txtfield(Index).Text) Then
         txtfield(Index).Text = 0#
         txtfield(Index).SetFocus
      Else
         txtfield(Index).Text = Format(txtfield(Index).Text, "#,##0.00")
      End If
      If Index = 5 Then
         Select Case psTransaction
            Case "POS"
               frmCP_POS.txtfield(3) = Format(txtfield(5), "#,##0.00")
               frmCP_POS.txtothers(6) = Format(CDbl(txtfield(5).Text - txtfield(1).Text), "#,##0.00")
            Case "Change"
               frmChange_Unit.txtfield(4) = Format(txtfield(5), "#,##0.00")
               frmChange_Unit.txtothers(0) = Format(CDbl(txtfield(5).Text - txtfield(1).Text), "#,##0.00")
         End Select
      End If
   Case 3
      If Not IsNumeric(txtfield(Index).Text) Then
         txtfield(Index).Text = 0#
         txtfield(Index).SetFocus
      Else
         txtfield(Index).Text = Format(txtfield(Index).Text, "#,##0")
      End If
   End Select
End Sub
