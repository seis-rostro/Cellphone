VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmCP_PO_ReplacementReg 
   BorderStyle     =   0  'None
   Caption         =   "PO Replacement Reg"
   ClientHeight    =   7215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11625
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   11625
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3765
      Left            =   3630
      TabIndex        =   14
      Top             =   3345
      Width           =   6390
      _ExtentX        =   11271
      _ExtentY        =   6641
      _Version        =   393216
      FocusRect       =   0
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   3735
      Left            =   120
      Tag             =   "wt0;fb0"
      Top             =   3360
      Width           =   3480
      _ExtentX        =   6138
      _ExtentY        =   6588
      BackColor       =   12632256
      ClipControls    =   0   'False
      BorderStyle     =   3
      Begin VB.TextBox txtFieldDetail 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   360
         Index           =   3
         Left            =   1230
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   1755
         Width           =   1260
      End
      Begin VB.TextBox txtFieldDetail 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   360
         Index           =   2
         Left            =   1230
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   1335
         Width           =   1260
      End
      Begin VB.TextBox txtFieldDetail 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   360
         Index           =   1
         Left            =   1230
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   930
         Width           =   2130
      End
      Begin VB.TextBox txtFieldDetail 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   360
         Index           =   0
         Left            =   1230
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   540
         Width           =   2130
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Price"
         Height          =   195
         Index           =   9
         Left            =   195
         TabIndex        =   21
         Top             =   1815
         Width           =   690
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QTY"
         Height          =   195
         Index           =   8
         Left            =   180
         TabIndex        =   19
         Top             =   1410
         Width           =   330
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   17
         Top             =   1005
         Width           =   795
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Code/IMEI"
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   15
         Top             =   615
         Width           =   780
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   2235
      Index           =   0
      Left            =   120
      Tag             =   "wt0;fb0"
      Top             =   1095
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   3942
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   6
         Left            =   6570
         TabIndex        =   29
         Text            =   "Text1"
         Top             =   1170
         Width           =   3090
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   660
         Index           =   4
         Left            =   1215
         MultiLine       =   -1  'True
         TabIndex        =   13
         Text            =   "frmCP_PO_ReplacementReg.frx":0000
         Top             =   1410
         Width           =   4560
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   6570
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   840
         Width           =   3090
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   465
         Index           =   3
         Left            =   1230
         MultiLine       =   -1  'True
         TabIndex        =   7
         Text            =   "frmCP_PO_ReplacementReg.frx":0008
         Top             =   915
         Width           =   4545
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   1230
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   90
         Width           =   2265
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   1230
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   585
         Width           =   4545
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
         Left            =   6945
         TabIndex        =   31
         Tag             =   "eb0;et0"
         Top             =   315
         Width           =   2385
      End
      Begin VB.Shape Shape4 
         Height          =   330
         Index           =   0
         Left            =   6915
         Top             =   270
         Width           =   2445
      End
      Begin VB.Shape Shape3 
         Height          =   390
         Index           =   0
         Left            =   6885
         Top             =   240
         Width           =   2505
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Refer #"
         Height          =   195
         Index           =   10
         Left            =   5955
         TabIndex        =   30
         Top             =   1215
         Width           =   540
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   6570
         TabIndex        =   11
         Tag             =   "ht0;ft0"
         Top             =   1515
         Width           =   3090
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   195
         Index           =   4
         Left            =   315
         TabIndex        =   12
         Top             =   1425
         Width           =   735
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   195
         Index           =   1
         Left            =   6030
         TabIndex        =   8
         Top             =   885
         Width           =   345
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   195
         Index           =   6
         Left            =   315
         TabIndex        =   6
         Top             =   945
         Width           =   570
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   315
         Left            =   1320
         Tag             =   "et0;ht2"
         Top             =   195
         Width           =   2265
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trans. No."
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   0
         Top             =   120
         Width           =   750
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier"
         Height          =   195
         Index           =   2
         Left            =   315
         TabIndex        =   4
         Top             =   630
         Width           =   570
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         Height          =   195
         Index           =   3
         Left            =   6030
         TabIndex        =   10
         Top             =   1545
         Width           =   360
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   10260
      TabIndex        =   23
      Top             =   2445
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
      Picture         =   "frmCP_PO_ReplacementReg.frx":0017
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   10260
      TabIndex        =   24
      Top             =   555
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
      Picture         =   "frmCP_PO_ReplacementReg.frx":0791
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   10260
      TabIndex        =   25
      Top             =   1815
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Void"
      AccessKey       =   "V"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCP_PO_ReplacementReg.frx":0F0B
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   525
      Index           =   1
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   9885
      _ExtentX        =   17436
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
         Height          =   330
         Index           =   1
         Left            =   4665
         MaxLength       =   50
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   90
         Width           =   5115
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
         Left            =   945
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   90
         Width           =   2610
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier"
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
         Left            =   3825
         TabIndex        =   27
         Top             =   135
         Width           =   705
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Trans.#"
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
         Index           =   10
         Left            =   90
         TabIndex        =   26
         Top             =   135
         Width           =   660
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   10275
      TabIndex        =   28
      Top             =   1185
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "Print"
      AccessKey       =   "Print"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCP_PO_ReplacementReg.frx":1685
   End
End
Attribute VB_Name = "frmCP_PO_ReplacementReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pxeMODULENAME = "frmCP_PO_ReplacementReg"

Private WithEvents oTrans As clsCPPOReplacement
Attribute oTrans.VB_VarHelpID = -1
Private oBranch As clsBranch
Private oSkin As clsFormSkin

Dim pnIndex As Integer
Dim pnCtr As Integer
Dim pnRow As Integer
Dim pbGridGotFocus As Boolean
Dim pbMasterGotFocus As Boolean
Dim pbFormLoad As Boolean
Dim pbPosted As Boolean

Private Sub cmdButton_Click(Index As Integer)
   Dim lnMsg As Integer
   Dim lsRep As String
   
   With MSFlexGrid1
      Select Case Index
      Case 0 'Browse
         If oTrans.SearchTransaction() = True Then
            LoadMaster
            LoadDetail
         End If
      Case 1 'Void
         If txtField(0).Text <> "" Then
            If oTrans.Master("cTranStat") <> xeStatePosted Then
               lsRep = MsgBox("Cancel Transaction?", vbYesNo + vbQuestion, "Confirm")
                  If lsRep = vbYes Then
                    Call oTrans.CancelTransaction
                  End If
            End If
         End If
      Case 2 'Close
         Unload Me
      Case 3 'Close
         If txtField(0).Text <> "" Then
            lsRep = MsgBox("Print Transaction!!!", vbYesNo + vbQuestion, "Confirm")
               If lsRep = vbYes Then
                  If Not PrintTrans Then MsgBox "Unable to Print Transaction!!!", vbCritical, "Warning"
               End If
         End If
      End Select
   End With
End Sub

Private Sub Form_Activate()
   If Not pbFormLoad Then pbFormLoad = True
   oApp.MenuName = Me.Tag
   Me.ZOrder 0
   txtSearch(1).SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyReturn, vbKeyUp, vbKeyDown
      Select Case KeyCode
      Case vbKeyReturn, vbKeyDown
         If GetFocus = MSFlexGrid1.hwnd Then Exit Sub
         SetNextFocus
      Case vbKeyUp
         SetPreviousFocus
      End Select
   End Select
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String

   lsOldProc = "Form_Load"
   'On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oTrans = New clsCPPOReplacement
   Set oTrans.AppDriver = oApp

   oTrans.InitTransaction
   oTrans.NewTransaction

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualRight

   Set oBranch = New clsBranch
   Set oBranch.AppDriver = oApp
   oBranch.InitRecord
   oBranch.NewRecord
   
   InitGrid
   InitFields
   ClearFields
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oTrans = Nothing
   Set oSkin = Nothing
   Set oBranch = Nothing
   pbFormLoad = False
End Sub

Private Sub MSFlexGrid1_GotFocus()
   pbGridGotFocus = True
End Sub

Private Sub MSFlexGrid1_RowColChange()
   pnRow = MSFlexGrid1.Row
End Sub

Private Sub oTrans_DetailRetrieved(ByVal Index As Integer)
    With MSFlexGrid1
      Select Case Index
      Case 1
         txtFieldDetail(0) = oTrans.Detail(pnRow - 1, "xReferNox")
      Case 2
         txtFieldDetail(1) = oTrans.Detail(pnRow - 1, "sDescript")
      Case 7
         txtFieldDetail(3) = oTrans.Detail(pnRow - 1, "nUnitPrce")
      End Select
      .TextMatrix(pnRow, 3) = IFNull(oTrans.Detail(pnRow - 1, "sModelNme"), "No Model")
      .TextMatrix(pnRow, 5) = oTrans.Detail(pnRow - 1, "sReferNox")
   End With
End Sub

Private Sub oTrans_MasterRetrieved(ByVal Index As Integer)
   txtField(Index) = oTrans.Master(Index)
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      Select Case Index
      Case 1
         .Text = Format(.Text, "MM/DD/YY")
      End Select
      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = oApp.getColor("HT1")
   End With

   pbGridGotFocus = False
   pbMasterGotFocus = True
   pnIndex = Index
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
      With txtField(Index)
         Select Case Index
         Case 2
            If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
               oTrans.SearchMaster Index, .Text
               If .Text <> "" Then SetNextFocus
            Else
               If .Text <> "" Then oTrans.SearchMaster Index, .Text
               If .Text <> "" Then SetNextFocus
            End If
         Case Else
            If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
               If .Text <> "" Then SetNextFocus
            End If
         End Select
      End With
      KeyCode = 0
   End If
End Sub

Private Sub InitGrid()
   With MSFlexGrid1
      .Rows = 2
      .Cols = 7
      .Font = "MS Sans Serif"

      'column title
      .TextMatrix(0, 1) = "BarrCode"
      .TextMatrix(0, 2) = "Description"
      .TextMatrix(0, 3) = "Model"
      .TextMatrix(0, 4) = "Unit Price"
      .TextMatrix(0, 5) = "ReferNo"
      .TextMatrix(0, 6) = "Qty"

      .Row = 0

      'column alignment
      For pnCtr = 0 To .Cols - 1
         .Col = pnCtr
         .CellFontBold = True
         .CellAlignment = 3
      Next

      'column width
      .ColWidth(0) = 330
      .ColWidth(1) = 2000
      .ColWidth(2) = 2000
      .ColWidth(4) = 1020
      .ColWidth(5) = 1300
      .ColWidth(6) = 500

      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      .ColAlignment(4) = 6
      .ColAlignment(5) = 1
      .ColAlignment(6) = 6

      .Row = 1
      .Col = 1
      .ColSel = .Cols - 1
   End With
End Sub

Private Sub InitFields()

   With MSFlexGrid1
      .TextMatrix(1, 0) = 1
      .TextMatrix(1, 1) = ""
      .TextMatrix(1, 2) = ""
      .TextMatrix(1, 3) = "0"
      .TextMatrix(1, 4) = "0.00"
      
      .Row = 1
      .Col = 1
      .ColSel = .Cols - 1
   End With
   
   pnRow = 1
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   oTrans.Master(Index) = txtField(Index)
End Sub

Private Sub ComputeTotal()
   Dim lnCtr As Integer
   Dim lnTotalAmt As Currency
   
   With MSFlexGrid1
      lnTotalAmt = 0#
      For lnCtr = 1 To .Rows - 1
         lnTotalAmt = Format(lnTotalAmt + CDbl(.TextMatrix(lnCtr, 3)) * (.TextMatrix(lnCtr, 4)), "#0.00")
      Next
   End With
   
   txtField(6) = Format(lnTotalAmt, "#0.00")
   oTrans.Master("nTranTotl") = lnTotalAmt
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

Private Sub LoadMaster()
   Dim pnCtr As Integer
   
   For pnCtr = 0 To txtField.Count
      Select Case pnCtr
         Case 0, 1, 2, 3, 4
            txtField(pnCtr).Text = oTrans.Master(pnCtr)
      End Select
   Next
   
   txtSearch(0).Text = oTrans.Master(0)
   txtSearch(1).Text = oTrans.Master(2)
   lblTotal.Caption = Format(oTrans.Master(5), "#,##0.00")
   
   Label2.Caption = Format(TransStat(oTrans.Master("cTranStat")), ">")
   
End Sub

Private Sub LoadDetail()
   txtFieldDetail(0) = oTrans.Detail(pnRow - 1, "xReferNox")
   txtFieldDetail(1) = oTrans.Detail(pnRow - 1, "sDescript")
   txtFieldDetail(2) = oTrans.Detail(pnRow - 1, "nQuantity")
   txtFieldDetail(3) = oTrans.Detail(pnRow - 1, "nUnitPrce")
   With MSFlexGrid1
      .TextMatrix(pnRow, 1) = oTrans.Detail(pnRow - 1, "xReferNox")
      .TextMatrix(pnRow, 2) = oTrans.Detail(pnRow - 1, "sDescript")
      .TextMatrix(pnRow, 3) = IFNull(oTrans.Detail(pnRow - 1, "sModelNme"), "No Model")
      .TextMatrix(pnRow, 4) = oTrans.Detail(pnRow - 1, "nUnitPrce")
      .TextMatrix(pnRow, 6) = oTrans.Detail(pnRow - 1, "nQuantity")
      .TextMatrix(pnRow, 5) = oTrans.Detail(pnRow - 1, "sReferNox")
   End With
End Sub

Private Sub ClearFields()
   For pnCtr = 0 To txtField.Count - 1
      Select Case pnCtr
      Case 0
         txtField(pnCtr).Text = ""
      Case 1
         txtField(pnCtr).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
      Case 2, 3, 4
         txtField(pnCtr).Text = Empty
      End Select
   Next
   
      txtField(6).Text = ""
      
  For pnCtr = 0 To txtFieldDetail.Count - 1
      Select Case pnCtr
      Case 0, 1, 2
         txtFieldDetail(pnCtr).Text = ""
      Case 3
         txtFieldDetail(pnCtr).Text = "0.00"
      End Select
   Next
   txtSearch(0).Text = ""
   txtSearch(1).Text = ""

   With MSFlexGrid1
      .Rows = 2
      .ColWidth(3) = 2300

      'empty row
      .TextMatrix(1, 1) = ""
      .TextMatrix(1, 2) = ""
      .TextMatrix(1, 3) = ""
      .TextMatrix(1, 4) = "0.00"
      .TextMatrix(1, 5) = ""
      .TextMatrix(1, 6) = "0"
   End With
   pbPosted = False
End Sub

Private Sub txtSearch_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
      
      With txtSearch(Index)
         Select Case Index
         Case 0, 1
            If KeyCode = vbKeyF3 Then
               oTrans.SearchTransaction .Text, False
               If .Text <> "" Then SetNextFocus
            Else
               If .Text <> "" Then oTrans.SearchMaster Index, .Text
            End If
         End Select
      End With
      
      KeyCode = 0
      LoadMaster
      LoadDetail
   End If
   
End Sub

Private Function PrintTrans() As Boolean
   Dim loreport As frmRepViewer

   Dim lrs As ADODB.Recordset
   Dim lors As ADODB.Recordset
   Dim lnCtr As Integer
   Dim lsOldProc As String
   Dim lsStockIDx As String

   lsOldProc = "PrintTrans"
   'On Error GoTo errProc

   PrintTrans = True
   Set lrs = New ADODB.Recordset
   lrs.Fields.Append "lField01", adCurrency
   lrs.Fields.Append "nField01", adInteger, 3
   lrs.Fields.Append "nField02", adChar, 1
   lrs.Fields.Append "sField01", adVarChar, 20
   lrs.Fields.Append "sField02", adVarChar, 128
   lrs.Fields.Append "sField03", adVarChar, 20
   lrs.Fields.Append "sField04", adVarChar, 128
   lrs.Fields.Append "sField05", adVarChar, 100
   lrs.Fields.Append "sField06", adVarChar, 30
   lrs.Fields.Append "sField07", adVarChar, 25
   lrs.Fields.Append "sField08", adVarChar, 25
   
   lrs.Open

   With oTrans
      For lnCtr = 0 To .ItemCount - 1
         lrs.AddNew
         lrs.Fields("lField01") = 0#  'oTrans.Detail(lnCtr, "nUnitPrce") 'she 2017-12-22 validation of unit price is accounting
         lrs.Fields("nField01") = oTrans.Detail(lnCtr, "nQuantity")
         lrs.Fields("nField02") = oTrans.Detail(lnCtr, "cHsSerial")
         lrs.Fields("sField01") = oTrans.Detail(lnCtr, "sTransNox")
         lrs.Fields("sField02") = oTrans.Detail(lnCtr, "sStockIDx")
         lrs.Fields("sField03") = oTrans.Detail(lnCtr, "sBarrCode")
         lrs.Fields("sField04") = oTrans.Detail(lnCtr, "sDescript")
         lrs.Fields("sField05") = oTrans.Detail(lnCtr, "sBrandNme")
         lrs.Fields("sField06") = IFNull(oTrans.Detail(lnCtr, "sSerialNo"), "")
         lrs.Fields("sField07") = IFNull(oTrans.Detail(lnCtr, "sReferNox"), "")
         lrs.Fields("sField08") = IFNull(oTrans.Master("sReferNox"), "")
      Next
      lrs.Sort = "nField02,sField05,sField03,sField06"
   End With

   Set lors = New ADODB.Recordset
   If lors.State = adStateOpen Then lors.Close

   lors.Open "SELECT" _
               & "  CONCAT(a.sAddressx, ', ', b.sTownName, ', ', c.sProvName, ' ', b.sZippCode) as xAddressx" _
               & ", a.sBranchNm" _
            & " From Branch a" _
               & ", TownCity b" _
               & ", Province c" _
            & " WHERE a.sBranchCd = " & strParm(Left(oTrans.Master("sTransNox"), Len(oApp.BranchCode))) _
               & " AND a.sTownIDxx = b.sTownIDxx" _
               & " AND b.sProvIDxx = c.sProvIDxx" _
            , oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText

   Set lors = New ADODB.Recordset
   If lors.State = adStateOpen Then lors.Close

   lors.Open "SELECT" _
               & "  CONCAT(a.sAddressx, ', ', b.sTownName, ', ', c.sProvName, ' ', b.sZippCode) as xAddressx" _
               & ", a.sBranchNm" _
            & " From Branch a" _
               & ", TownCity b" _
               & ", Province c" _
            & " WHERE a.sBranchCd = " & strParm(Left(oTrans.Master("sTransNox"), Len(oApp.BranchCode))) _
               & " AND a.sTownIDxx = b.sTownIDxx" _
               & " AND b.sProvIDxx = c.sProvIDxx" _
            , oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText

   Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\CPPurchaseReplacement.rpt")
   'assign important info to the report
   oReport.DiscardSavedData
   oReport.FieldMappingType = crAutoFieldMapping
   oReport.Database.SetDataSource lrs

   oReport.Sections("PHa").ReportObjects("txtTransNox").SetText "CP-" & Right(oTrans.Master("sTransNox"), 10)
   oReport.Sections("PHa").ReportObjects("txtDate").SetText txtField(1).Text
   oReport.Sections("PHb").ReportObjects("txtTo").SetText txtField(2).Text
   oReport.Sections("PHb").ReportObjects("txtToAddress").SetText txtField(3).Text
   oReport.Sections("PHb").ReportObjects("txtFrom").SetText lors("sBranchNm")
   oReport.Sections("PHb").ReportObjects("txtFromAddress").SetText lors("xAddressx")
   oReport.Sections("RFb").ReportObjects("txtRemarks").SetText txtField(4).Text
   oReport.Sections("PF").ReportObjects("txtRptUser").SetText oApp.UserName

   Set loreport = New frmRepViewer
   Set loreport.ReportSource = oReport
   loreport.Show

endPoc:
   If Not pbPosted Then
      oTrans.CloseTransaction (oTrans.Master("sTransNox"))
      pbPosted = True
   End If
   Set oReport = Nothing
   Set lrs = Nothing
   Set lors = Nothing
   Set loreport = Nothing
   Exit Function
errProc:
   PrintTrans = False
   ShowError lsOldProc & "( " & " )"
End Function
