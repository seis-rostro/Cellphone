VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmCPLoadAdjPosted 
   BorderStyle     =   0  'None
   Caption         =   "CP Inventory Adjustment"
   ClientHeight    =   5415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8775
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5415
   ScaleWidth      =   8775
   ShowInTaskbar   =   0   'False
   Tag             =   "et0;eb0;et0;bc2"
   Begin xrControl.xrFrame xrFrame1 
      Height          =   4185
      Index           =   1
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   1125
      Width           =   7110
      _ExtentX        =   12541
      _ExtentY        =   7382
      BackColor       =   12632256
      Enabled         =   0   'False
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   1275
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   930
         Width           =   1905
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   10
         Left            =   1275
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   3510
         Width           =   2955
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   6
         Left            =   1275
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   3180
         Width           =   5550
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   4
         Left            =   1275
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   2340
         Width           =   1515
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   5
         Left            =   1275
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   2670
         Width           =   1515
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   8
         Left            =   1275
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   1590
         Width           =   5550
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
         Left            =   1275
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   345
         Width           =   1905
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   5325
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   930
         Width           =   1515
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   7
         Left            =   1275
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   1260
         Width           =   2805
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
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
         Index           =   9
         Left            =   5325
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   2340
         Width           =   1515
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "unknown"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   4380
         TabIndex        =   29
         Tag             =   "eb0;et0"
         Top             =   405
         Width           =   2400
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H8000000F&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   285
         Index           =   2
         Left            =   4395
         Tag             =   "et0;et0"
         Top             =   405
         Width           =   2400
      End
      Begin VB.Shape Shape3 
         Height          =   390
         Index           =   0
         Left            =   4335
         Top             =   345
         Width           =   2505
      End
      Begin VB.Shape Shape4 
         Height          =   330
         Index           =   0
         Left            =   4365
         Top             =   375
         Width           =   2445
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date Tran"
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   8
         Top             =   975
         Width           =   810
      End
      Begin VB.Shape Shape2 
         Height          =   1815
         Index           =   1
         Left            =   105
         Top             =   2190
         Width           =   6855
      End
      Begin VB.Shape Shape2 
         Height          =   1980
         Index           =   0
         Left            =   105
         Top             =   135
         Width           =   6855
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Approved"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   22
         Top             =   3525
         Width           =   690
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   20
         Top             =   3210
         Width           =   630
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AmtOnHnd"
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
         Index           =   4
         Left            =   4350
         TabIndex        =   18
         Top             =   2385
         Width           =   915
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Doc #"
         Height          =   195
         Index           =   3
         Left            =   4770
         TabIndex        =   27
         Top             =   1005
         Width           =   450
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AmtIn"
         Height          =   195
         Index           =   1
         Left            =   255
         TabIndex        =   14
         Top             =   2385
         Width           =   405
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AmtOut"
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   16
         Top             =   2700
         Width           =   525
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   315
         Left            =   1365
         Tag             =   "et0;ht2"
         Top             =   450
         Width           =   1920
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Barrcode"
         Height          =   195
         Index           =   4
         Left            =   195
         TabIndex        =   10
         Top             =   1320
         Width           =   1125
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
         Left            =   195
         TabIndex        =   6
         Top             =   375
         Width           =   765
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   2
         Left            =   195
         TabIndex        =   12
         Top             =   1650
         Width           =   810
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   7425
      TabIndex        =   26
      Top             =   1800
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
      Picture         =   "frmCPLoadAdjPosted.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   7425
      TabIndex        =   24
      Top             =   540
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Memo"
      AccessKey       =   "M"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCPLoadAdjPosted.frx":077A
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   525
      Index           =   0
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   7110
      _ExtentX        =   12541
      _ExtentY        =   926
      Begin VB.TextBox txtSearch 
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
         Index           =   2
         Left            =   4305
         MaxLength       =   50
         TabIndex        =   5
         Top             =   90
         Width           =   2715
      End
      Begin VB.TextBox txtSearch 
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
         Index           =   1
         Left            =   2625
         MaxLength       =   50
         TabIndex        =   3
         Top             =   90
         Width           =   870
      End
      Begin VB.TextBox txtSearch 
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
         Index           =   0
         Left            =   810
         TabIndex        =   1
         Top             =   90
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Barcode"
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
         Index           =   1
         Left            =   3540
         TabIndex        =   4
         Top             =   135
         Width           =   780
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Doc #"
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
         Index           =   8
         Left            =   2055
         TabIndex        =   2
         Top             =   135
         Width           =   675
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Trans #"
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
         Left            =   90
         TabIndex        =   0
         Top             =   135
         Width           =   765
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   7425
      TabIndex        =   25
      Top             =   1170
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
      Picture         =   "frmCPLoadAdjPosted.frx":0EF4
   End
End
Attribute VB_Name = "frmCPLoadAdjPosted"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmCPLoadAdjPosted"

Private WithEvents oTrans As clsCPLoadAdj
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin
Private pnIndex As Integer
Private pbPosted As Boolean

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   Dim lnRep As Integer
   Dim loObj As Object

   lsOldProc = "cmdButton_Click"
   'On Error GoTo errProc

   Select Case Index
   Case 0
      If oTrans.Master(0) <> "" Then oTrans.GetMemo loObj
   Case 1
      oTrans.SearchTransaction
   Case 2
      Unload Me
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

   Set oTrans = New clsCPLoadAdj
   Set oTrans.AppDriver = oApp

   oTrans.Branch = oApp.BranchCode
   oTrans.InitTransaction
   oTrans.OpenTransaction ""

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualRight

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oTrans = Nothing
   Set oSkin = Nothing
End Sub

Private Sub oTrans_LoadData()
   Dim pnCtr As Integer

   For pnCtr = 0 To 10
      Select Case pnCtr
      Case 0
         txtField(pnCtr).Text = Format(oTrans.Master(pnCtr), "@@@@-@@@@@@")
         txtSearch(0).Text = oTrans.Master(pnCtr)
      Case 1
         txtField(pnCtr).Text = Format(oTrans.Master(pnCtr), "MMMM DD, YYYY")
      Case 2
         txtField(pnCtr).Text = IFNull(oTrans.Master(pnCtr))
         txtSearch(1).Text = IFNull(oTrans.Master(pnCtr))
      Case 3
      Case 4, 5, 9
         txtField(pnCtr).Text = Format(oTrans.Master(pnCtr), "#,##0.00")
      Case 7
         txtField(pnCtr).Text = IFNull(oTrans.Master(pnCtr))
         txtSearch(2).Text = txtField(pnCtr).Text
      Case Else
         txtField(pnCtr).Text = IFNull(oTrans.Master(pnCtr))
      End Select
   Next

   Label2.Caption = TransStat(oTrans.Master("cTranStat"))
   pbPosted = oTrans.Master("cTranStat") = xeStatePosted
End Sub

Private Sub oTrans_MasterRetrieved(ByVal Index As Variant)
   Select Case Index
   Case 7 To 10
      txtField(Index) = IFNull(oTrans.Master(Index))
   Case 12
      Label2.Caption = TransStat(oTrans.Master("cTranStat"))
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

   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
      With txtField(Index)
         Select Case Index
         Case 7, 8
            If KeyCode = vbKeyF3 Then
               oTrans.SearchMaster Index, txtField(Index).Text
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyReturn, vbKeyUp, vbKeyDown
      Select Case KeyCode
      Case vbKeyReturn, vbKeyDown
         SetNextFocus
      Case vbKeyUp
         SetPreviousFocus
      End Select
   Case vbKeyF8
      Call oTrans.DeleteTransaction
   Case vbKeyF12
   End Select
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
      Case 4, 5
         If Not IsNumeric(.Text) Then txtField(Index).Text = ""
         .Text = Format(.Text, "#,##0")
      End Select
      oTrans.Master(Index) = .Text
   End With

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " _
                       & "  " & Index _
                       & ", " & Cancel _
                       & " )", True
End Sub

Private Sub txtSearch_GotFocus(Index As Integer)
   With txtSearch(Index)
      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = oApp.getColor("HT1")
   End With
End Sub

Private Sub txtSearch_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String

   lsOldProc = "txtSearch_KeyDown"
   'On Error GoTo errProc

   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
      With txtSearch(Index)
         Select Case Index
         Case 0
            oTrans.OpenTransaction .Text
         Case Else
            If .Text <> "" Then
               oTrans.SearchTransaction .Text, IIf(Index = 1, True, False)
            Else
               oTrans.OpenTransaction ""
            End If
         End Select
      End With
      KeyCode = 0
   ElseIf KeyCode = vbKeyF8 Then
      If oTrans.DeleteTransaction = False Then
         MsgBox "Unable to delete transaction!"
      End If
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

Private Sub txtSearch_LostFocus(Index As Integer)
   With txtSearch(Index)
      .BackColor = oApp.getColor("EB")
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
