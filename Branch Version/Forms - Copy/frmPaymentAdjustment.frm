VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmPaymentAdjustment 
   BorderStyle     =   0  'None
   Caption         =   "Payment Adjustment"
   ClientHeight    =   6210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8130
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6210
   ScaleWidth      =   8130
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   5535
      Index           =   1
      Left            =   1605
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   9763
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   585
         Index           =   3
         Left            =   1155
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   1680
         Width           =   4905
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Index           =   7
         Left            =   4230
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   2850
         Width           =   1515
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   6
         Left            =   1395
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   3075
         Width           =   1515
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   5
         Left            =   1395
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   2670
         Width           =   1515
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   8
         Left            =   1155
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   4950
         Width           =   4905
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   1155
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   960
         Width           =   1905
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   795
         Index           =   4
         Left            =   1155
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   3870
         Width           =   4905
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
         Index           =   0
         Left            =   1155
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   315
         Width           =   1905
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   1155
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1320
         Width           =   4905
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   195
         Index           =   1
         Left            =   195
         TabIndex        =   23
         Top             =   1725
         Width           =   930
      End
      Begin VB.Shape Shape3 
         Height          =   1215
         Left            =   330
         Top             =   2445
         Width           =   5745
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Balance"
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
         Index           =   2
         Left            =   3345
         TabIndex        =   20
         Top             =   2880
         Width           =   705
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amt. Out"
         Height          =   195
         Index           =   0
         Left            =   510
         TabIndex        =   18
         Top             =   3105
         Width           =   615
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amt. In"
         Height          =   195
         Index           =   1
         Left            =   495
         TabIndex        =   17
         Top             =   2730
         Width           =   495
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Approve By"
         Height          =   195
         Index           =   7
         Left            =   210
         TabIndex        =   16
         Top             =   4995
         Width           =   825
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Trans. No."
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
         Left            =   270
         TabIndex        =   15
         Top             =   285
         Width           =   765
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date Tran"
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   1
         Top             =   1005
         Width           =   810
      End
      Begin VB.Shape Shape2 
         Height          =   5295
         Index           =   0
         Left            =   105
         Top             =   90
         Width           =   6165
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   195
         Index           =   5
         Left            =   195
         TabIndex        =   9
         Top             =   3900
         Width           =   630
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   315
         Left            =   1245
         Tag             =   "et0;ht2"
         Top             =   420
         Width           =   1920
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier"
         Height          =   195
         Index           =   4
         Left            =   195
         TabIndex        =   8
         Top             =   1380
         Width           =   930
      End
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   600
      Index           =   3
      Left            =   105
      TabIndex        =   12
      Top             =   1785
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
      Picture         =   "frmPaymentAdjustment.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   105
      TabIndex        =   10
      Top             =   525
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
      Picture         =   "frmPaymentAdjustment.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   105
      TabIndex        =   13
      Top             =   1785
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
      Picture         =   "frmPaymentAdjustment.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   600
      Index           =   1
      Left            =   105
      TabIndex        =   11
      Top             =   1155
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
      Picture         =   "frmPaymentAdjustment.frx":166E
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   105
      TabIndex        =   14
      Top             =   525
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
      Picture         =   "frmPaymentAdjustment.frx":1DE8
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   6
      Left            =   105
      TabIndex        =   21
      Top             =   1155
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Register"
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
      Picture         =   "frmPaymentAdjustment.frx":2562
   End
End
Attribute VB_Name = "frmPaymentAdjustment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmPaymentAdjustment"

Private WithEvents oTrans As clsAdjustment
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin
Private pnIndex As Integer

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   Dim lnRep As Integer
   Dim loObj As Object

   lsOldProc = "cmdButton_Click"
   'On Error GoTo errProc

   Select Case Index
   Case 0 'save
      If oTrans.SaveTransaction(True) Then
         MsgBox "Record Updated Successfully!!!", vbInformation, "Notice"
         Clearfields
         initButton xeModeReady
      Else
         MsgBox "Unable to Update Record!!!", vbCritical, "Warning"
         txtField(pnIndex).SetFocus
      End If
   Case 1 'search
      oTrans.SearchTransaction (""), False
      txtField(pnIndex).SetFocus
   Case 3 'cancel
      Clearfields
      initButton xeModeReady
   Case 4 'New
      oTrans.NewTransaction
      initButton xeModeAddNew
      txtField(1).SetFocus
   Case 5 'Close
      Unload Me
   Case 6 'Register
      frmPaymentAdjustmentReg.Show
   End Select

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String

   lsOldProc = "Form_Load"
   'On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oTrans = New clsAdjustment
   Set oTrans.AppDriver = oApp
   
   If oApp.MenuName = "mnuARAdjustment" Then
      oTrans.AccountType = "AR"
      frmPaymentAdjustment.Caption = "AR Payment Adjustment"
   Else
      oTrans.AccountType = "AP"
      frmPaymentAdjustment.Caption = "AP Payment Adjustment"
   End If
   
   oTrans.Branch = oApp.BranchCode
   oTrans.InitTransaction
   oTrans.NewTransaction

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransaction
   
   Clearfields
   initButton xeModeAddNew
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oTrans = Nothing
   Set oSkin = Nothing
End Sub

Private Sub oTrans_MasterRetrieved(ByVal Index As Variant, ByVal Value As Variant)
   Select Case txtField(Index)
   Case 5, 6, 7
      txtField(Index) = Format(oTrans.Master(Index), "#,##0.00")
   Case Else
      txtField(Index) = oTrans.Master(Index)
   End Select
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      If Index = 1 Then .Text = Format(.Text, "MM/DD/YYYY")
      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = oApp.getColor("HT1")
   End With
   pnIndex = Index
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String

   lsOldProc = "txtField_KeyDown"
   'On Error GoTo errProc

   With txtField(Index)
      Select Case Index
      Case 2
         If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
            oTrans.SearchMaster Index, txtField(2).Text
         End If
      Case 5
         If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
            If txtField(6).Text <> 0# Then
               txtField(5).Text = 0#
            End If
         End If
      Case 6
         If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
            If txtField(5).Text <> 0# Then
               txtField(6).Text = 0#
            End If
         End If
      End Select
   End With
   KeyCode = 0

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " _
                       & "  " & Index _
                       & ", " & KeyCode _
                       & ", " & Shift _
                       & " )", True
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

Private Sub initButton(lnStat As Integer)
   Dim lbShow As Boolean

   lbShow = IIf(lnStat = 0, False, True)
   cmdButton(0).Visible = lbShow
   cmdButton(1).Visible = lbShow
   cmdButton(3).Visible = lbShow

   cmdButton(4).Visible = Not lbShow
   cmdButton(5).Visible = Not lbShow
   cmdButton(6).Visible = Not lbShow

   xrFrame1(1).Enabled = lbShow

   If Not lbShow Then cmdButton(4).SetFocus
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   Dim lsOldProc As String

   lsOldProc = "txtField_Validate"
   'On Error GoTo errProc

   With txtField(Index)
      .Text = TitleCase(.Text)

      Select Case Index
      Case 1
         If Not IsDate(.Text) Then .Text = oApp.ServerDate
         .Text = Format(.Text, "MMMM DD, YYYY")
      Case 5
         If Not IsNumeric(.Text) Then txtField(Index).Text = 0
            .Text = Format(.Text, "#,##0.00")
            oTrans.Master(Index) = .Text
      Case 6
         If Not IsNumeric(.Text) Then txtField(Index).Text = 0
            .Text = Format(.Text, "#,##0.00")
            oTrans.Master(Index) = .Text
      Case 4
         oTrans.Master(Index) = .Text
      Case 8
         oTrans.Master(8) = .Text
      End Select
   End With
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " _
                       & "  " & Index _
                       & ", " & Cancel _
                       & " )", True
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

Private Sub Clearfields()
   txtField(0).Text = oTrans.Master(0)
   txtField(1).Text = ""
   txtField(2).Text = ""
   txtField(3).Text = ""
   txtField(4).Text = ""
   txtField(5).Text = 0#
   txtField(6).Text = 0#
   txtField(7).Text = 0#
   txtField(8).Text = ""
End Sub
