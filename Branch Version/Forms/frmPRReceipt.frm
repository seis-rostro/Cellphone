VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmPRReceipt 
   BorderStyle     =   0  'None
   Caption         =   "Temporary Delivery Receipt"
   ClientHeight    =   3840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10110
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   10110
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   3135
      Index           =   0
      Left            =   1575
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   8400
      _ExtentX        =   14817
      _ExtentY        =   5530
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   6
         Left            =   5655
         MaxLength       =   128
         TabIndex        =   13
         Tag             =   "ht0;ft0"
         Top             =   2370
         Width           =   2430
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
         Height          =   285
         Index           =   0
         Left            =   1665
         TabIndex        =   1
         Top             =   165
         Width           =   2310
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   5685
         MaxLength       =   8
         TabIndex        =   5
         Top             =   900
         Width           =   2565
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   1665
         MaxLength       =   50
         TabIndex        =   3
         Top             =   600
         Width           =   2310
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   1665
         TabIndex        =   7
         Top             =   1290
         Width           =   6420
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   1665
         MaxLength       =   128
         TabIndex        =   11
         Top             =   1890
         Width           =   6420
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   1665
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1590
         Width           =   6420
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   285
         Left            =   1755
         Tag             =   "et0;ht2"
         Top             =   270
         Width           =   2310
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   9
         Left            =   165
         TabIndex        =   0
         Top             =   210
         Width           =   1425
      End
      Begin VB.Shape Shape3 
         Height          =   615
         Index           =   1
         Left            =   120
         Top             =   2310
         Width           =   8130
      End
      Begin VB.Shape Shape2 
         Height          =   1035
         Index           =   0
         Left            =   120
         Top             =   1215
         Width           =   8130
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "P.R. No."
         Height          =   285
         Index           =   2
         Left            =   4905
         TabIndex        =   4
         Top             =   945
         Width           =   930
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   285
         Index           =   12
         Left            =   315
         TabIndex        =   10
         Top             =   1935
         Width           =   690
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Amt. Tendered"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   8
         Left            =   3465
         TabIndex        =   12
         Top             =   2490
         Width           =   2160
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cust. Address"
         Height          =   195
         Index           =   11
         Left            =   315
         TabIndex        =   8
         Top             =   1635
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Date"
         Height          =   285
         Index           =   1
         Left            =   165
         TabIndex        =   2
         Top             =   645
         Width           =   1245
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name"
         Height          =   285
         Index           =   3
         Left            =   315
         TabIndex        =   6
         Top             =   1320
         Width           =   1200
      End
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   600
      Index           =   3
      Left            =   90
      TabIndex        =   18
      Top             =   3105
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
      Picture         =   "frmPRReceipt.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   90
      TabIndex        =   14
      Top             =   1215
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
      Picture         =   "frmPRReceipt.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   90
      TabIndex        =   19
      Top             =   3105
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
      Picture         =   "frmPRReceipt.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   600
      Index           =   1
      Left            =   90
      TabIndex        =   15
      Top             =   1845
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
      Picture         =   "frmPRReceipt.frx":166E
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   90
      TabIndex        =   16
      Top             =   1845
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
      Picture         =   "frmPRReceipt.frx":1DE8
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   90
      TabIndex        =   17
      Top             =   2475
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Receipt"
      AccessKey       =   "R"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmPRReceipt.frx":2562
   End
End
Attribute VB_Name = "frmPRReceipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Private Const pxeMODULENAME = "frmCashierTrans"
'
'Private WithEvents oTrans As clsTDR
'Private oSkin As clsFormSkin
'
'Dim pnIndex As Integer, pnCtr As Integer
'
'Private Sub cmdButton_Click(Index As Integer)
'   Dim lsOldProc As String
'   Dim lnRep As Integer
'
'   lsOldProc = "cmdButton_Click"
'   ''On Error GoTo errProc
'
'   Select Case Index
'   Case 0
'      If isEntryOK Then
'         oTrans.Master("sSystemCd") = "CP"
'
'         If oTrans.SaveTransaction = True Then
'            MsgBox "Record Updated Successfully!!!", vbInformation, "Notice"
'            Call cmdButton_Click(4) 'new
'         Else
'            MsgBox "Unable to Update Record!!!", vbCritical, "Warning"
'         End If
'      End If
'   Case 1
'      Select Case pnIndex
'      Case 3
'         oTrans.Master(pnIndex) = txtField(pnIndex).Text
'      Case Else
'         oTrans.SearchMaster pnIndex
'         txtField(pnIndex).SetFocus
'      End Select
'      txtField(pnIndex).SetFocus
'   Case 2
'   Case 3
'      lnRep = MsgBox("Transaction is in Update Mode!!!" & vbCrLf & _
'                     "Do you want to Cancel Transaction!!!", vbYesNo + vbQuestion, "Confirm")
'
'      If lnRep = vbYes Then
'         ClearFields
'         InitButton xeModeReady
'      Else
'         txtField(pnIndex).SetFocus
'      End If
'   Case 4
'      oTrans.NewTransaction
'      InitButton xeModeAddNew
'      ClearFields
'
'      txtField(1).SetFocus
'   Case 5
'      Unload Me
'   End Select
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " & Index & " )", True
'End Sub
'
'Private Sub Form_Activate()
'   oApp.MenuName = Me.Tag
'   Me.ZOrder 0
'End Sub
'
'Private Sub Form_Load()
'   Dim lsOldProc As String
'
'   lsOldProc = "Form_Load"
'   ''On Error GoTo errProc
'
'   CenterChildForm mdiMain, Me
'
'   Set oTrans = New clsTDR
'   Set oTrans.AppDriver = oApp
'
'   oTrans.Branch = oApp.BranchCode
'   oTrans.InitTransaction
'   oTrans.NewTransaction
'
'   Set oSkin = New clsFormSkin
'   Set oSkin.AppDriver = oApp
'   Set oSkin.Form = Me
'   oSkin.ApplySkin xeFormTransEqualLeft
'
'   ClearFields
'   InitButton xeModeAddNew
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " & " )", True
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'   Set oTrans = Nothing
'   Set oSkin = Nothing
'End Sub
'
'Private Sub oTrans_MasterRetrieved(ByVal Index As Variant)
'   txtField(Index).Text = IIf(IsNull(oTrans.Master(Index)), "", oTrans.Master(Index))
'End Sub
'
'Private Sub ClearFields()
'   Dim lotxt As TextBox
'
'   For Each lotxt In txtField
'      pnCtr = lotxt.Index
'      Select Case pnCtr
'      Case 0
'         txtField(pnCtr).Text = Format(oTrans.Master(pnCtr), IIf(Len(oApp.BranchCode) = 2, "@@@@-@@@@@@", "@@@@@@-@@@@@@"))
'      Case 1
'         txtField(pnCtr).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
'      Case 6
'         txtField(pnCtr).Text = Format(oTrans.Master(pnCtr), "#,##0.00")
'      Case Else
'         txtField(pnCtr).Text = oTrans.Master(pnCtr)
'      End Select
'   Next
'End Sub
'
'Private Sub txtField_GotFocus(Index As Integer)
'   With txtField(Index)
'      If Index = 1 Then .Text = Format(.Text, "MM/DD/YYYY")
'      .SelStart = 0
'      .SelLength = Len(.Text)
'      .BackColor = oApp.getColor("HT1")
'   End With
'   pnIndex = Index
'End Sub
'
'Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'   Dim lsOldProc As String
'
'   lsOldProc = "txtField_KeyDown"
'   ''On Error GoTo errProc
'
'   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
'      With txtField(Index)
'         Select Case Index
'         Case 3
'            oTrans.Master(Index) = txtField(Index).Text
'         End Select
'      End With
'      KeyCode = 0
'   End If
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " _
'                       & "  " & Index _
'                       & ", " & KeyCode _
'                       & ", " & Shift _
'                       & " )", True
'End Sub
'
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'   Select Case KeyCode
'   Case vbKeyReturn, vbKeyUp, vbKeyDown
'      Select Case KeyCode
'      Case vbKeyReturn, vbKeyDown
'         SetNextFocus
'      Case vbKeyUp
'         SetPreviousFocus
'      End Select
'   End Select
'End Sub
'
'Private Sub InitButton(lnStat As Integer)
'   Dim lbShow As Boolean
'
'   lbShow = IIf(lnStat = 0, False, True)
'   cmdButton(0).Visible = lbShow
'   cmdButton(1).Visible = lbShow
'   cmdButton(3).Visible = lbShow
'
'   cmdButton(4).Visible = Not lbShow
'   cmdButton(5).Visible = Not lbShow
'
'   xrFrame1(0).Enabled = lbShow
'
'   If Not lbShow Then cmdButton(4).SetFocus
'End Sub
'
'Private Function isEntryOK() As Boolean
'   If txtField(3).Text = "" Then
'      MsgBox "Customer not found!!!", vbCritical, "Warning"
'      txtField(3).SetFocus
'      GoTo EntryNotOK
'   End If
'
'   If CDbl(txtField(6)) = 0# Then
'      MsgBox "No Amount to Paid!!!" & vbCrLf & _
'             "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
'      GoTo EntryNotOK
'   End If
'
'   oTrans.Master("nTranTotl") = CDbl(txtField(6))
'
'EntryOK:
'   isEntryOK = True
'   Exit Function
'EntryNotOK:
'   isEntryOK = False
'End Function
'
'Private Sub txtField_LostFocus(Index As Integer)
'   With txtField(Index)
'      .BackColor = oApp.getColor("EB")
'   End With
'End Sub
'
'Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
'   Dim lsOldProc As String
'
'   lsOldProc = "txtField_Validate"
'   ''On Error GoTo errProc
'
'   With txtField(Index)
'      .Text = TitleCase(.Text)
'
'      Select Case Index
'      Case 1
'         If Not IsDate(.Text) Then .Text = oApp.ServerDate
'         .Text = Format(.Text, "MMMM DD, YYYY")
'      Case 2
'         If Not IsNumeric(.Text) Then txtField(Index).Text = ""
'      Case 6
'         If Not IsNumeric(.Text) Then txtField(Index).Text = "0.00"
'         .Text = Format(.Text, "#,##0.00")
'      End Select
'      oTrans.Master(Index) = .Text
'   End With
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " _
'                       & "  " & Index _
'                       & ", " & Cancel _
'                       & " )", True
'End Sub
'
'Private Sub ShowError(ByVal lsProcName As String, Optional bEnd As Boolean = False)
'   With oApp
'      .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
'      If bEnd Then
'         .xShowError
'         End
'      Else
'         With Err
'            .Raise .Number, .Source, .Description
'         End With
'      End If
'   End With
'End Sub
