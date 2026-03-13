VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmMPProductInquiry 
   BorderStyle     =   0  'None
   Caption         =   "MP Product Inquiry"
   ClientHeight    =   5535
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9315
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   9315
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   4875
      Index           =   0
      Left            =   1575
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   8599
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   645
         Index           =   9
         Left            =   1590
         MultiLine       =   -1  'True
         TabIndex        =   19
         Text            =   "frmMPProductInquiry.frx":0000
         Top             =   4050
         Width           =   5850
      End
      Begin VB.ComboBox cmbField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   5
         ItemData        =   "frmMPProductInquiry.frx":0006
         Left            =   1590
         List            =   "frmMPProductInquiry.frx":001C
         TabIndex        =   17
         Top             =   3705
         Width           =   2355
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   6
         Left            =   1590
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   3015
         Width           =   2355
      End
      Begin VB.ComboBox cmbField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   8
         ItemData        =   "frmMPProductInquiry.frx":0074
         Left            =   1590
         List            =   "frmMPProductInquiry.frx":007E
         TabIndex        =   15
         Top             =   3375
         Width           =   2370
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   645
         Index           =   81
         Left            =   1590
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   1560
         Width           =   5850
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Index           =   0
         Left            =   1590
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   150
         Width           =   2355
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   1
         Left            =   1590
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   855
         Width           =   2355
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   80
         Left            =   1590
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   1215
         Width           =   5850
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   82
         Left            =   1590
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   2325
         Width           =   3000
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   83
         Left            =   1590
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   2670
         Width           =   3000
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Source Info"
         Height          =   195
         Index           =   9
         Left            =   180
         TabIndex        =   16
         Top             =   3765
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   2
         Top             =   923
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   18
         Top             =   4275
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transact No"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   0
         Top             =   218
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Color"
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   10
         Top             =   2730
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   8
         Top             =   2385
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   4
         Top             =   1275
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   6
         Top             =   1785
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Purchase Type"
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   14
         Top             =   3435
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Target Date"
         Height          =   195
         Index           =   8
         Left            =   180
         TabIndex        =   12
         Top             =   3090
         Width           =   855
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   1
         Left            =   1695
         Tag             =   "et0;ht2"
         Top             =   270
         Width           =   2355
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   90
      TabIndex        =   24
      Top             =   1830
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
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
      Picture         =   "frmMPProductInquiry.frx":0095
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   90
      TabIndex        =   23
      Top             =   1200
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Save"
      AccessKey       =   "S"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmMPProductInquiry.frx":080F
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   90
      TabIndex        =   22
      Top             =   1830
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
      Picture         =   "frmMPProductInquiry.frx":0F89
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   90
      TabIndex        =   20
      Top             =   570
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "Searc&h"
      AccessKey       =   "h"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmMPProductInquiry.frx":1703
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   90
      TabIndex        =   21
      Top             =   1200
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&New"
      AccessKey       =   "N"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmMPProductInquiry.frx":1E7D
   End
End
Attribute VB_Name = "frmMPProductInquiry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pxeMODULENAME = "frmMPProductnquiry"

Private WithEvents oTrans As clsMPProductInquiry
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin
Private bLoaded As Boolean

Dim psSelected() As String
Dim pnIndex As Integer

Private Sub cmbField_LostFocus(Index As Integer)
   With cmbField(Index)
      Select Case Index
      Case 8
         oTrans.Master(8) = .ListIndex
      Case 5
         Select Case .ListIndex
         Case 0
            oTrans.Master(5) = "FB"
         Case 1
            oTrans.Master(5) = "WS"
         Case 2
            oTrans.Master(5) = "WI"
         Case 3
            oTrans.Master(5) = "ER"
         Case 4
            oTrans.Master(5) = "BR"
         Case 4
            oTrans.Master(5) = "MD"
         End Select
      End Select
   End With
End Sub

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   Dim lnRep As Integer

   lsOldProc = "cmdButton_Click"
   '''On Error GoTo errProc

   Select Case Index
   Case 0   'cancel
      lnRep = MsgBox("Transaction is in Update Mode!!!" & vbCrLf & _
                     "Do you want to Cancel Transaction!!!", vbYesNo + vbQuestion, "Confirm")
      If lnRep = vbYes Then
         InitForm 1
      End If
   Case 1   'search
      Call oTrans.SearchMaster(2, txtField(2))
   Case 2   'save
      If oTrans.SaveTransaction() Then
         MsgBox "Transaction Updated Successfully!!!", vbInformation, "Confirm"
         oTrans.NewTransaction
         InitForm 0
      Else
         MsgBox "Unable to Update Transaction!!!", vbCritical, "Warning"
      End If
   Case 3   'new
      oTrans.NewTransaction
      InitForm 0
   Case 4   'close
      Unload Me
   End Select

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Sub Form_Activate()
   Dim lsOldProc As String

   lsOldProc = "Form_Activate"
   '''On Error GoTo errProc

   oApp.MenuName = Me.Tag
   Me.ZOrder 0

   If bLoaded = False Then
      bLoaded = True
   End If

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String

   lsOldProc = "Form_Load"
   '''On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oTrans = New clsMPProductInquiry
   Set oTrans.AppDriver = oApp

   oTrans.Branch = oApp.BranchCode
   oTrans.InitTransaction

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualLeft

   oTrans.NewTransaction
   Call InitForm(0)
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oTrans = Nothing
   Set oSkin = Nothing
End Sub

Private Sub InitForm(ByVal fnEdit As Integer)
   Dim loTxt As TextBox

   xrFrame1(0).Enabled = (fnEdit = 0)
   cmdButton(4).Visible = Not (fnEdit = 0)
   cmdButton(3).Visible = Not (fnEdit = 0)

   cmdButton(0).Visible = (fnEdit = 0)
   cmdButton(1).Visible = (fnEdit = 0)
   cmdButton(2).Visible = (fnEdit = 0)

   For Each loTxt In txtField
      loTxt = ""
   Next

   cmbField(8).ListIndex = 0
   cmbField(5).ListIndex = 2
   oTrans.Master(5) = "WI"

   If fnEdit = 0 Then LoadMaster
End Sub

Private Sub LoadMaster()
   With oTrans
      txtField(0).Text = .Master(0)
      txtField(1).Text = strLongDate(.Master(1))
   End With
End Sub

Private Sub oTrans_MasterRetrieved(ByVal Index As Variant)
   Select Case Index
      Case 1
         txtField(Index) = strLongDate(oTrans.Master(Index))
      Case 6
         If IsDate(oTrans.Master(Index)) Then
            txtField(Index) = strLongDate(oTrans.Master(Index))
         Else
            txtField(Index) = ""
         End If
      Case Else
         txtField(Index) = oTrans.Master(Index)
   End Select
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   With txtField(Index)
      oTrans.Master(Index) = .Text
   End With
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("HT1")
      .SelStart = 0
      .SelLength = Len(.Text)
   End With

   Select Case Index
   Case 1, 6
      If IsDate(oTrans.Master(Index)) Then
         txtField(Index) = strShortDate(oTrans.Master(Index))
      End If
   End Select

   pnIndex = Index
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyF3
      Select Case Index
      Case 80, 82, 83
         If oTrans.SearchMaster(Index, txtField(Index).Text) Then
            SetNextFocus
         End If
      End Select
   Case vbKeyReturn
      Select Case Index
      Case 80, 82, 83
         If txtField(Index) <> "" Then
            If oTrans.SearchMaster(Index, txtField(Index).Text) Then SetNextFocus
         Else
            oTrans.Master(Index) = txtField(Index).Text
         End If
      End Select
   End Select
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
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
