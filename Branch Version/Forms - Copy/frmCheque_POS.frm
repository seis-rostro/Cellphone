VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmCheque_POS 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Cheque Transaction"
   ClientHeight    =   3510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6735
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   ScaleHeight     =   3510
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   960
      Index           =   1
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   1545
      Width           =   6540
      _ExtentX        =   11536
      _ExtentY        =   1693
      BackColor       =   7716603
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   240
         Index           =   1
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   7
         Top             =   90
         Width           =   1570
      End
      Begin VB.TextBox txtfield 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000315F8&
         Height          =   495
         Index           =   3
         Left            =   4830
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   "Amount"
         Top             =   360
         Width           =   1560
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   2
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   9
         Top             =   345
         Width           =   1570
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   4
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   11
         Top             =   600
         Width           =   1570
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Amount"
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   105
         Width           =   1500
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cheque Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   3390
         TabIndex        =   12
         Top             =   375
         Width           =   1380
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cash Amount"
         Height          =   270
         Index           =   2
         Left            =   135
         TabIndex        =   8
         Top             =   360
         Width           =   1440
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cash Given"
         Height          =   270
         Index           =   3
         Left            =   135
         TabIndex        =   10
         Top             =   615
         Width           =   1440
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   960
      Index           =   0
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   6540
      _ExtentX        =   11536
      _ExtentY        =   1693
      BackColor       =   7716603
      Begin VB.TextBox txtothers 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   240
         Index           =   0
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   3
         Top             =   345
         Width           =   4695
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   5
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   5
         Top             =   600
         Width           =   3060
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   0
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   1
         Top             =   90
         Width           =   4695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cheque Number"
         Height          =   300
         Index           =   6
         Left            =   150
         TabIndex        =   4
         Top             =   615
         Width           =   1365
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   300
         Index           =   4
         Left            =   150
         TabIndex        =   2
         Top             =   360
         Width           =   885
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Name"
         Height          =   285
         Index           =   5
         Left            =   135
         TabIndex        =   0
         Top             =   90
         Width           =   1260
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   1
      Left            =   5535
      TabIndex        =   15
      Top             =   2715
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
      Picture         =   "frmCheque_POS.frx":0000
      BackColor       =   14286077
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   0
      Left            =   4425
      TabIndex        =   14
      Top             =   2715
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
      Picture         =   "frmCheque_POS.frx":077A
      BackColor       =   14286077
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
End
Attribute VB_Name = "frmCheque_POS"
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
Dim pnindex As Integer
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
               If txtfield(0).Text <> "" And txtfield(5).Text <> "" Then
                  frmCP_POS.Payment = "Cheque"
                  frmCP_POS.txtfield(6) = "Cheque No. : " & txtfield(5).Text
                  Me.Hide
                  frmCP_POS.txtfield(4).SetFocus
               Else
                  MsgBox "Incomplete Data!!!", vbCritical, "Warning"
               End If
            Case "Change"
               If txtfield(0).Text <> "" And txtfield(5).Text <> "" Then
                  frmChange_Unit.Payment = "Cheque"
                  frmChange_Unit.txtfield(5) = "Cheque No. : " & txtfield(5).Text
                  MsgBox txtfield(0).Tag
                  Me.Hide
                  frmChange_Unit.txtfield(5).SetFocus
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
      txtfield(0).SetFocus
      txtfield(1).Enabled = False
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
    
   txtfield(2).Text = "0.00"
   txtfield(3).Text = "0.00"
   txtfield(4).Text = "0.00"
      
End Sub

Private Sub Form_Unload(Cancel As Integer)

   Set oDriver = Nothing
   Set oSkin = Nothing

End Sub

Private Sub txtField_GotFocus(Index As Integer)
   oDriver.ColumnIndex = Index
   pnindex = Index
   txtfield(Index).BackColor = &HE1FEFF
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

   If KeyCode = vbKeyF3 Or KeyCode = 13 Then
      If Index = 0 Then
         SearchBank False
         If txtfield(Index).Text <> "" Then SetNextFocus
     End If
     KeyCode = 0
   End If

End Sub
Private Sub txtfield_Validate(Index As Integer, Cancel As Boolean)
Dim temp As Long
   Select Case Index
   Case 2, 4
      If Not IsNumeric(txtfield(Index).Text) Then
         txtfield(Index).Text = 0#
         txtfield(Index).SetFocus
      Else
         txtfield(Index).Text = Format(txtfield(Index).Text, "#,##0.00")
         temp = Format(CDbl(txtfield(1).Text) - CDbl(txtfield(2).Text), "#,##0.00")
         txtfield(3).Text = Format(temp, "#,##0.00")
         If Index = 4 Then
            Select Case psTransaction
               Case "POS"
                  frmCP_POS.txtfield(3) = Format(txtfield(4).Text, "#,##0.00")
                  frmCP_POS.txtothers(6) = Format(txtfield(4).Text - txtfield(2).Text, "#,##0.00")
               Case "Change"
                  frmChange_Unit.txtfield(4) = Format(txtfield(4).Text, "#,##0.00")
                  frmChange_Unit.txtothers(0) = Format(txtfield(4).Text - txtfield(2).Text, "#,##0.00")
            End Select
         End If
      End If
   End Select
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   txtfield(Index).BackColor = &H80000005
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

Private Sub SearchBank(ByVal SearchValue As Boolean)
   Dim lsSearch As String
   Dim lrs As ADODB.Recordset
   Dim lssql As String
   
   Set lrs = New ADODB.Recordset
   
   lssql = "SELECT" _
            & " sBankIDxx, " _
            & " sBankName, " _
            & " sAddressx " _
         & " FROM Banks " _
         & " WHERE cRecdStat = 1 " _

   If SearchValue Then
      lssql = lssql & " AND sBankName = '" & txtfield(0).Text & "'"
   Else
      lssql = lssql & " AND sBankName LIKE '" & txtfield(0).Text & "%' "
   End If
            
   lssql = lssql & " ORDER BY sBankName"
   lrs.Open lssql, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
      
   If lrs.RecordCount = 1 Then
      txtfield(0).Text = lrs("sBankName")
      txtfield(0).Tag = lrs("sBankIDxx")
      txtothers(0).Text = lrs("sAddressx")
      ElseIf lrs.RecordCount > 1 Then
           lsSearch = KwikBrowse(oApp, lrs, _
                             "sBankIDxx�" _
                           & "sBankName�" _
                           & "sAddressx", _
                             "Bank ID�" _
                           & "Bank Name�" _
                           & "Address")
                           
           If lsSearch <> "" Then
               psSelected = Split(lsSearch, "�")
               txtfield(0).Text = psSelected(1)
               txtfield(0).Tag = psSelected(0)
               txtothers(0).Text = psSelected(2)
           End If
   Else
      txtfield(0).Text = ""
      txtfield(0).Tag = ""
      txtothers(0).Text = ""
   End If
   Set lrs = Nothing

End Sub


