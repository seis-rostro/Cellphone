VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmSupplies_Request_Register 
   BorderStyle     =   0  'None
   Caption         =   "Supply Request-Register"
   ClientHeight    =   5340
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8565
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5340
   ScaleWidth      =   8565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin xrControl.xrFrame xrFrame2 
      Height          =   3570
      Left            =   1500
      Tag             =   "wt0;fb0"
      Top             =   1470
      Width           =   6780
      _ExtentX        =   11959
      _ExtentY        =   6297
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   495
         Index           =   10
         Left            =   1170
         TabIndex        =   17
         Top             =   2865
         Width           =   5445
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   9
         Left            =   1170
         TabIndex        =   16
         Top             =   1440
         Width           =   1830
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   7
         Left            =   1170
         MaxLength       =   512
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   1080
         Width           =   4080
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   1020
         TabIndex        =   14
         Top             =   150
         Width           =   1620
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   1170
         TabIndex        =   13
         Top             =   735
         Width           =   1815
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   3
         Left            =   4800
         TabIndex        =   12
         Top             =   720
         Width           =   1845
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   2
         Left            =   1170
         MaxLength       =   512
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   1785
         Width           =   1830
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   4
         Left            =   1170
         MaxLength       =   512
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   2145
         Width           =   1830
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   8
         Left            =   1170
         MaxLength       =   512
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   2505
         Width           =   1830
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   10
         Left            =   255
         TabIndex        =   26
         Top             =   1140
         Width           =   795
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stock ID"
         Height          =   195
         Index           =   9
         Left            =   255
         TabIndex        =   25
         Top             =   1500
         Width           =   645
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trans. #"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   24
         Top             =   210
         Width           =   735
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   195
         Index           =   0
         Left            =   255
         TabIndex        =   23
         Top             =   795
         Width           =   345
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   315
         Left            =   1110
         Tag             =   "et0;ht2"
         Top             =   240
         Width           =   1620
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Entry No."
         Height          =   195
         Index           =   6
         Left            =   4125
         TabIndex        =   22
         Top             =   780
         Width           =   660
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   195
         Index           =   2
         Left            =   255
         TabIndex        =   21
         Top             =   2925
         Width           =   675
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
         Height          =   195
         Index           =   4
         Left            =   255
         TabIndex        =   20
         Top             =   1860
         Width           =   585
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qty on Hnd"
         Height          =   195
         Index           =   7
         Left            =   255
         TabIndex        =   19
         Top             =   2220
         Width           =   810
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Act. on Hnd"
         Height          =   195
         Index           =   8
         Left            =   270
         TabIndex        =   18
         Top             =   2580
         Width           =   855
      End
      Begin VB.Shape Shape4 
         Height          =   420
         Index           =   0
         Left            =   4400
         Top             =   120
         Width           =   2220
      End
      Begin VB.Shape Shape3 
         Height          =   360
         Index           =   0
         Left            =   4425
         Top             =   150
         Width           =   2160
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4470
         TabIndex        =   0
         Tag             =   "eb0;et0"
         Top             =   180
         Width           =   2070
      End
   End
   Begin xrControl.xrFrame xrFrame3 
      Height          =   1125
      Left            =   1500
      Tag             =   "wt0;fb0"
      Top             =   300
      Width           =   6780
      _ExtentX        =   11959
      _ExtentY        =   1984
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   5
         Left            =   1125
         MaxLength       =   512
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   570
         Width           =   4080
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   1125
         TabIndex        =   5
         Top             =   225
         Width           =   1815
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   5
         Left            =   210
         TabIndex        =   8
         Top             =   630
         Width           =   795
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trans #"
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   7
         Top             =   285
         Width           =   555
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   3345
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
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
      Picture         =   "frmSupplies_Request_Register.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   1455
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Browse"
      AccessKey       =   "B"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmSupplies_Request_Register.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   0
      TabIndex        =   3
      Top             =   2085
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Approve"
      AccessKey       =   "A"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmSupplies_Request_Register.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   0
      TabIndex        =   4
      Top             =   2715
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&DisApprv"
      AccessKey       =   "D"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmSupplies_Request_Register.frx":166E
   End
End
Attribute VB_Name = "frmSupplies_Request_Register"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pxeMODULENAME = "frmSupplies_Request_Register"

Private WithEvents oTrans As clsShiftSchedule
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin
Private bLoaded As Boolean

Dim pnIndex As Integer
Dim pbSearched As Boolean
Private psTransNox As String
Public Property Let TransNox(value As String)
   psTransNox = value
End Property

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   Dim lnRep As Integer

   lsOldProc = "cmdButton_Click"
   On Error GoTo errProc

   Select Case Index
   Case 0   'close
      Unload Me
   Case 1   'search
      If pnIndex = 0 Or pnIndex = 1 Then
         If pnIndex = 0 Then
            If oTrans.SearchTransaction(txtSearch(pnIndex).Text, False) Then
               ClearFields
               LoadMaster
               Call InitFields
            End If
         Else
            If oTrans.SearchTransaction(txtSearch(pnIndex).Text) Then
               ClearFields
               LoadMaster
               Call InitFields
            End If
         End If
         pnIndex = 3
      Else
         If oTrans.SearchTransaction("") Then
            ClearFields
            LoadMaster
            Call InitFields
         End If
      End If
   Case 2   'approve
      If oTrans.Master(0) <> "" Then
         If oTrans.CloseTransaction(oTrans.Master(0)) Then
            MsgBox "Transaction was closed successfuly!!!", vbInformation, "Notice"
         Else
            MsgBox "Closing/Posting transaction failed!!!", vbInformation, "Notice"
         End If
         Call ClearFields
      End If
      GoTo endWithFocus
   Case 3   'disapprove
      If oTrans.Master(0) <> "" And txtField(0) <> "" Then
         If oTrans.CancelTransaction Then
            MsgBox "Transaction was cancelled!!!", vbInformation, "Notice"
         Else
            MsgBox "Transaction cancellation failed!!!", vbInformation, "Notice"
         End If
         ClearFields
      End If
      GoTo endWithFocus
   End Select

endProc:
   Exit Sub
endWithFocus:
   txtSearch(0) = ""
   txtSearch(1) = ""
   txtSearch(0).SetFocus
   GoTo endProc
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Sub Form_Activate()
   Dim lsOldProc As String
   
   lsOldProc = "Form_Activate"
   On Error GoTo errProc
   
   oApp.MenuName = Me.Tag
   Me.ZOrder 0
   
   If bLoaded = False Then
      bLoaded = True
   End If

   pbSearched = False
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
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

Private Sub Form_Load()
   Dim lsOldProc As String

   lsOldProc = "Form_Load"
'   On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oTrans = New clsShiftSchedule
   Set oTrans.AppDriver = oApp
   
   If LCase(oApp.ProductID) = "petmgr" Then
      oTrans.TransStatus = 10
   Else
      oTrans.TransStatus = 0
   End If
   
   oTrans.Branch = oApp.BranchCode
   oTrans.InitTransaction
   If psTransNox <> "" Then
      '@@@ soft-monitor
      Call oTrans.OpenTransaction(psTransNox)
   End If

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransaction
   
   ClearFields
   InitFields
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oTrans = Nothing
   Set oSkin = Nothing
End Sub

Private Sub LoadMaster()
   Dim loTxt As TextBox
   
   For Each loTxt In txtField
      Select Case loTxt.Index
         Case 1, 3
            loTxt.Text = strLongDate(oTrans.Master(loTxt.Index))
         Case Else
            loTxt.Text = oTrans.Master(loTxt.Index)
      End Select
   Next
   
   txtSearch(0) = txtField(0)
   txtSearch(1) = txtField(2)
   
   If oTrans.Master("cTranStat") = "4" Then
      Label2.Caption = "APPLIED"
   Else
      Label2.Caption = TransStat(CInt(oTrans.Master("cTranStat")))
   End If
      
   pbSearched = True
End Sub

Private Sub ClearFields()
   Dim loTxt As TextBox
   
   For Each loTxt In txtField
      loTxt = ""
   Next
   
   txtSearch(0) = ""
   txtSearch(1) = ""
   Label2.Caption = ""
   
   pbSearched = False
End Sub

Private Sub oTrans_MasterRetrieved(ByVal Index As Variant, ByVal value As Variant)
   Select Case Index
      Case 1, 3
      txtField(Index) = strLongDate(oTrans.Master(Index))
      Case 9
         Label2.Caption = TransStat(CInt(value))
   End Select
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   
   Select Case Index
      Case 1, 3
         txtField(Index) = strShortDate(oTrans.Master(Index))
   End Select
   
   With txtField(Index)
      .BackColor = oApp.getColor("HT1")
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
   
   pnIndex = Index
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtSearch_LostFocus(Index As Integer)
   With txtSearch(Index)
      .BackColor = oApp.getColor("EB")
   End With
   
   pnIndex = Index
End Sub
Private Sub txtSearch_GotFocus(Index As Integer)
   With txtSearch(Index)
      .BackColor = oApp.getColor("HT1")
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
   
   pnIndex = Index
End Sub
Private Sub txtSearch_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
      Select Case Index
      Case 0
         If oTrans.SearchTransaction(txtSearch(Index).Text, False) Then
            ClearFields
            LoadMaster
            Call InitFields
         End If
      Case 1
         If oTrans.SearchTransaction(txtSearch(Index).Text) Then
            ClearFields
            LoadMaster
            Call InitFields
         End If
      End Select
   End If
End Sub

Private Sub InitFields()
   xrFrame2.Enabled = (pbSearched = True)
End Sub

Private Sub ShowError(ByVal lsProcName As String, Optional bEnd As Boolean = False)
   With oApp
      .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
      If bEnd Then
         .xShowError
         End
      Else
         With Err
            .Raise .Number, .Source, .Description
         End With
      End If
   End With
End Sub


