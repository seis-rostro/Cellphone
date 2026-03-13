VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmLoadWallet 
   BorderStyle     =   0  'None
   Caption         =   "CP Load Wallet"
   ClientHeight    =   5070
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6945
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5070
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin xrControl.xrFrame xrFrame1 
      Height          =   4425
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   7805
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   1320
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   705
         Width           =   2265
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   7
         Left            =   2850
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   3045
         Width           =   2070
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
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
         Height          =   495
         Index           =   6
         Left            =   2850
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   2535
         Width           =   2070
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   1320
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   1035
         Width           =   2265
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
         Left            =   1320
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   150
         Width           =   2265
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   3
         Left            =   1320
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   1365
         Width           =   3720
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   4
         Left            =   1320
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   1695
         Width           =   3720
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   5
         Left            =   1320
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   2025
         Width           =   3720
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transact Date"
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   2
         Top             =   720
         Width           =   1020
      End
      Begin VB.Shape Shape2 
         Height          =   1815
         Left            =   150
         Top             =   2415
         Width           =   4875
      End
      Begin VB.Label lblChangeAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   19.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   2850
         TabIndex        =   17
         Top             =   3480
         Width           =   2085
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Change Amt."
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
         Index           =   10
         Left            =   1320
         TabIndex        =   16
         Top             =   3540
         Width           =   1095
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amt. Tendered"
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
         Index           =   8
         Left            =   1320
         TabIndex        =   14
         Top             =   3075
         Width           =   1260
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Load Amount"
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
         Index           =   6
         Left            =   1320
         TabIndex        =   12
         Top             =   2580
         Width           =   1125
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reference No."
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   4
         Top             =   1080
         Width           =   1050
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   315
         Left            =   1410
         Tag             =   "et0;ht2"
         Top             =   255
         Width           =   2265
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transact. No."
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
         Index           =   11
         Left            =   90
         TabIndex        =   0
         Top             =   210
         Width           =   1185
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Barcode"
         Height          =   195
         Index           =   9
         Left            =   150
         TabIndex        =   6
         Top             =   1425
         Width           =   600
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   7
         Left            =   150
         TabIndex        =   8
         Top             =   1755
         Width           =   795
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile Number"
         Height          =   195
         Index           =   5
         Left            =   150
         TabIndex        =   10
         Top             =   2085
         Width           =   1065
      End
   End
   Begin VB.Label lblField 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "F5-OK"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   420
      Index           =   4
      Left            =   5580
      TabIndex        =   20
      Top             =   1050
      Width           =   1275
   End
   Begin VB.Label lblField 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Escape"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   420
      Index           =   3
      Left            =   5580
      TabIndex        =   19
      Top             =   1545
      Width           =   1275
   End
   Begin VB.Label lblField 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "F1-Help"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   420
      Index           =   0
      Left            =   5580
      TabIndex        =   18
      Top             =   555
      Width           =   1275
   End
End
Attribute VB_Name = "frmLoadWallet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmEload"

Private WithEvents oTrans As clsCPLoadWallet
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin

Dim pnIndex As Integer

Private Sub Form_Activate()
   'column width
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
   Case vbKeyF1
   Case vbKeyF5
'      If Not isEntryOK Then Exit Sub
      Call txtField_Validate(pnIndex, False)
      If Not oTrans.SaveTransaction Then
         MsgBox "Unable to Save Transaction!!!" & vbCrLf & _
                  "Please contanct GMC/GGC SEG for assistance!!!", vbCritical, "Warning"
         Exit Sub
      End If
      
      oTrans.NewTransaction
      ClearFields
      txtField(1).SetFocus
   Case vbKeyEscape
      Unload Me
   End Select
End Sub

Private Sub Form_Load()
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.DisableClose = True
   oSkin.ApplySkin xeFormTransDetail
   
   Set oTrans = New clsCPLoadWallet
   Set oTrans.AppDriver = oApp

   oTrans.InitTransaction
   oTrans.NewTransaction
   
   ClearFields
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
   Set oTrans = Nothing
End Sub

Private Sub ClearFields()
   Dim lnCtr As Integer
   
   For lnCtr = 0 To txtField.Count - 1
      Select Case lnCtr
      Case 0
         txtField(lnCtr).Text = Format(oTrans.Master("sTransNox"), "@@@@-@@@@@@")
      Case 1
         txtField(lnCtr).Text = Format(oTrans.Master("sReferNox"), ">")
      Case 2
         txtField(lnCtr).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
      Case 6, 7
         txtField(lnCtr).Text = "0.00"
      Case Else
         txtField(lnCtr).Text = ""
      End Select
   Next
   
   lblChangeAmount.Caption = "0.00"
End Sub

Private Sub oTrans_MasterRetrieved(ByVal Index As Integer)
   With oTrans
      Select Case Index
      Case 3
         txtField(3).Text = oTrans.Master("sBarrCode")
      Case 4
         txtField(4).Text = oTrans.Master("sDescript")
      End Select
   End With
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
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
   
   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
      With txtField(pnIndex)
         Select Case Index
         Case 3
            If KeyCode = vbKeyF3 Then
               If oTrans.searchBarrcode(1, .Text) Then
                  If .Text <> "" Then SetNextFocus
               End If
            Else
               If .Text <> "" Then oTrans.searchBarrcode 1, .Text
            End If
         End Select
      End With
      KeyCode = 0
   End If

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " _
                       & "  " & Index _
                       & ", " & KeyCode _
                       & ", " & Shift _
                       & " )", True
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   Dim lnChange As Currency
   
   With txtField(Index)
      Select Case Index
      Case 1
         .Text = Format(.Text, ">")
      Case 2
         If Not IsDate(.Text) Then .Text = oApp.ServerDate
         .Text = Format(.Text, "MMMM DD, YYYY")
      Case 6, 7
         If Not IsNumeric(.Text) Then .Text = 0#
         .Text = Format(.Text, "#,##0.00")
         lnChange = CDbl(txtField(7).Text) - CDbl(txtField(6).Text)
         lblChangeAmount.Caption = Format(IIf(lnChange > 0#, lnChange, 0#), "#,##0.00")
      End Select
      oTrans.Master(Index) = .Text
   End With
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
