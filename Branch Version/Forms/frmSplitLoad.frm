VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmSplitLoad 
   BorderStyle     =   0  'None
   Caption         =   "Load Transfer"
   ClientHeight    =   6930
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6930
   ScaleWidth      =   7920
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame2 
      Height          =   2475
      Left            =   1590
      Tag             =   "wt0;fb0"
      Top             =   4320
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   4366
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtFied 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   7
         Left            =   4035
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox txtFied 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   6
         Left            =   4035
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   435
         Width           =   2055
      End
      Begin VB.TextBox txtFied 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Index           =   10
         Left            =   4035
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   1245
         Width           =   1290
      End
      Begin VB.TextBox txtFied 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   8
         Left            =   1035
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   1245
         Width           =   1290
      End
      Begin VB.TextBox txtFied 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   5
         Left            =   1035
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox txtFied 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   4
         Left            =   1035
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   450
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Supersede"
         Height          =   210
         Index           =   10
         Left            =   3165
         TabIndex        =   25
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   210
         Index           =   6
         Left            =   75
         TabIndex        =   24
         Top             =   870
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Supersede"
         Height          =   210
         Index           =   9
         Left            =   3165
         TabIndex        =   23
         Top             =   465
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Trnsfer Amt"
         Height          =   210
         Index           =   8
         Left            =   3165
         TabIndex        =   17
         Top             =   1290
         Width           =   840
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Load Amt"
         Height          =   210
         Index           =   7
         Left            =   75
         TabIndex        =   16
         Top             =   1290
         Width           =   675
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Orig Code"
         Height          =   210
         Index           =   5
         Left            =   75
         TabIndex        =   15
         Top             =   450
         Width           =   840
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   3690
      Left            =   1590
      Tag             =   "wt0;fb0"
      Top             =   600
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   6509
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtFied 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   14
         Left            =   1200
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   1470
         Width           =   1740
      End
      Begin VB.TextBox txtFied 
         Appearance      =   0  'Flat
         Height          =   765
         Index           =   9
         Left            =   1200
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   2745
         Width           =   4455
      End
      Begin VB.TextBox txtFied 
         Appearance      =   0  'Flat
         Height          =   450
         Index           =   3
         Left            =   1200
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   2235
         Width           =   4455
      End
      Begin VB.TextBox txtFied 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   2
         Left            =   1200
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1845
         Width           =   4455
      End
      Begin VB.TextBox txtFied 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   1
         Left            =   1200
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   1095
         Width           =   1740
      End
      Begin VB.TextBox txtFied 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Index           =   0
         Left            =   1095
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   225
         Width           =   1740
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Refer No"
         Height          =   210
         Index           =   11
         Left            =   285
         TabIndex        =   26
         Top             =   1530
         Width           =   885
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   210
         Index           =   4
         Left            =   315
         TabIndex        =   14
         Top             =   2805
         Width           =   885
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   210
         Index           =   3
         Left            =   315
         TabIndex        =   13
         Top             =   2295
         Width           =   885
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Destination"
         Height          =   210
         Index           =   2
         Left            =   300
         TabIndex        =   12
         Top             =   1890
         Width           =   885
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   210
         Index           =   1
         Left            =   285
         TabIndex        =   11
         Top             =   1155
         Width           =   885
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Trans. No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   180
         TabIndex        =   9
         Top             =   270
         Width           =   885
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   90
      TabIndex        =   18
      Top             =   615
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
      Picture         =   "frmSplitLoad.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   600
      Index           =   2
      Left            =   90
      TabIndex        =   19
      Top             =   1875
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
      Picture         =   "frmSplitLoad.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   90
      TabIndex        =   20
      Top             =   615
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
      Picture         =   "frmSplitLoad.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   90
      TabIndex        =   21
      Top             =   1875
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
      Picture         =   "frmSplitLoad.frx":166E
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   90
      TabIndex        =   22
      Top             =   1245
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
      Picture         =   "frmSplitLoad.frx":1DE8
   End
End
Attribute VB_Name = "frmSplitLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pxeMODULENAME = "frmSplitLoad"

Private WithEvents oTrans As clsCPWallet2Retail
Attribute oTrans.VB_VarHelpID = -1

Private oSkin As clsFormSkin
Private bLoaded As Boolean

Dim pbFormLoad As Boolean

Private Sub Form_Activate()
   Dim lsOldProc As String

   lsOldProc = "Form_Activate"
   ''On Error GoTo errProc

   oApp.MenuName = Me.Tag
   Me.ZOrder 0

   If Not pbFormLoad = False Then
      pbFormLoad = True
   End If

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String

   lsOldProc = "Form_Load"
   'On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oTrans = New clsCPWallet2Retail
   Set oTrans.AppDriver = oApp

   oTrans.InitTransaction
   oTrans.NewTransaction

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualLeft

   ClearOthers
   InitFields
   
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oTrans = Nothing
   Set oSkin = Nothing
   
   pbFormLoad = False
End Sub

Private Sub InitForm(ByVal fnEdit As Integer)
   Dim loTxt As TextBox
   
   cmdButton(4).Visible = Not (fnEdit = 0)
   cmdButton(3).Visible = Not (fnEdit = 0)
   
   cmdButton(0).Visible = (fnEdit = 0)
   cmdButton(1).Visible = (fnEdit = 0)
   cmdButton(2).Visible = (fnEdit = 0)
   
   For Each loTxt In txtField
      loTxt = ""
   Next
   
   If fnEdit = 0 Then LoadMaster
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

Private Sub InitFields()
   With oTrans
      txtFied(0) = Format(.Master("sTransNox"), "@@@@@@-@@@@@@")
      txtFied(1) = Format(.Master("dTransact"), "MMMM DD, YYYY")
      txtFied(8) = Format(.Master("nAmountxx"), "#,##0.00")
      txtFied(10) = Format(.Master("nAmountxx"), "#,##0.00")

   End With
End Sub

Private Sub initButton(lnStat As Integer)
   Dim lbShow As Boolean

   lbShow = IIf(lnStat = 0, False, True)
   cmdButton(1).Visible = lbShow
   cmdButton(3).Visible = Not lbShow
   
   cmdButton(0).Visible = Not lbShow
   cmdButton(2).Visible = lbShow

End Sub

Private Sub ClearOthers()
   Dim loTxt As TextBox
   Dim lnCtr As Integer
   
   For lnCtr = 2 To 10
      txtFied(lnCtr) = ""
   Next
   
   txtFied(14).Text = ""
   
   For Each loTxt In txtFied
      loTxt.BackColor = oApp.getColor("EB")
   Next
End Sub

Private Sub cmdButton_Click(Index As Integer)
   Dim lnRep As Integer
   
   Select Case Index
   Case 0 'New
      oTrans.NewTransaction
      initButton (0)
   Case 1 'save
      If txtFied(0).Text <> "" Then
         If oTrans.SaveTransaction Then
            MsgBox "Transaction Save Succesfully!!!", vbInformation
            initButton (0)
         End If
      End If
   Case 2 'Cancel
      lnRep = MsgBox("Do you want to cancel Transaction???", vbYesNo, "Confirmation")
         If lnRep = vbYes Then
            oTrans.CancelTransaction
         End If
      initButton (0)
   Case 3 'Close
      Unload Me
   Case 4 'Register
   End Select
End Sub

Private Sub oTrans_MasterRetrieved(ByVal Index As Variant, ByVal Value As Variant)
   Select Case Index
   Case 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 14
      txtFied(Index) = Value
   End Select
End Sub

Private Sub txtFied_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
      Select Case Index
      Case 2, 4, 6
         oTrans.SearchMaster Index, txtFied(Index).Text
         SetNextFocus
      Case 8
         txtFied(8).Text = Format(txtFied(8).Text, "#,##0.00")
         txtFied(10).Text = Format(txtFied(8).Text, "#,##0.00")
      Case Else
         SetNextFocus
      End Select
   End If
End Sub

Private Sub txtFied_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
   Case 8
      If Not IsNumeric(txtFied(Index).Text) Then txtFied(Index).Text = 0#
      oTrans.Master(Index) = Trim(txtFied(Index).Text)
   Case 10
      txtFied(Index) = txtFied(8)
   Case 14
      txtFied(Index) = txtFied(14)
      oTrans.Master(Index) = Trim(txtFied(Index).Text)
   Case Else
      oTrans.Master(Index) = txtFied(Index)
   End Select
End Sub
