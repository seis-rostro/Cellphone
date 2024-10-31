VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0"; "xrGridControl.ocx"
Begin VB.Form frmCP_BranchOrder_Posting 
   BorderStyle     =   0  'None
   Caption         =   "Branch Spareparts Reservation Posting"
   ClientHeight    =   6960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11595
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6960
   ScaleWidth      =   11595
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   10230
      TabIndex        =   19
      Top             =   3480
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
      Picture         =   "frmCP_BranchOrder_Posting.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   10230
      TabIndex        =   16
      Top             =   1590
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
      Picture         =   "frmCP_BranchOrder_Posting.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   10230
      TabIndex        =   17
      Top             =   2220
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "Con&firm"
      AccessKey       =   "f"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCP_BranchOrder_Posting.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   10230
      TabIndex        =   18
      Top             =   2850
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "Cance&l"
      AccessKey       =   "l"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCP_BranchOrder_Posting.frx":166E
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   525
      Index           =   1
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   926
      Begin VB.TextBox txtField 
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
         Index           =   6
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   90
         Width           =   4845
      End
      Begin VB.TextBox txtField 
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
         Left            =   7500
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   90
         Width           =   2265
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Source/Origin"
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
         Left            =   90
         TabIndex        =   0
         Top             =   135
         Width           =   1245
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Transact. No."
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
         Left            =   6285
         TabIndex        =   2
         Top             =   135
         Width           =   1200
      End
   End
   Begin xrGridEditor.GridEditor GridEditor1 
      Height          =   4275
      Left            =   90
      TabIndex        =   15
      Tag             =   "et0;eb0;et0;bc2"
      Top             =   2625
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   7541
      AllowBigSelection=   -1  'True
      AutoAdd         =   0   'False
      AutoNumber      =   -1  'True
      BACKCOLOR       =   -2147483643
      BACKCOLORBKG    =   8421504
      BACKCOLORFIXED  =   -2147483633
      BACKCOLORSEL    =   -2147483635
      BORDERSTYLE     =   1
      COLS            =   2
      FILLSTYLE       =   0
      FIXEDCOLS       =   1
      FIXEDROWS       =   1
      FOCUSRECT       =   1
      EDITORBACKCOLOR =   -2147483643
      EDITORFORECOLOR =   -2147483640
      FORECOLOR       =   -2147483640
      FORECOLORFIXED  =   -2147483630
      FORECOLORSEL    =   -2147483634
      FORMATSTRING    =   ""
      Object.HEIGHT          =   4275
      GRIDCOLOR       =   12632256
      GRIDCOLORFIXED  =   0
      BeginProperty GRIDFONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GRIDLINES       =   1
      GRIDLINESFIXED  =   2
      GRIDLINEWIDTH   =   1
      MOUSEICON       =   "frmCP_BranchOrder_Posting.frx":1DE8
      MOUSEPOINTER    =   0
      REDRAW          =   -1  'True
      RIGHTTOLEFT     =   0   'False
      ROWS            =   2
      SCROLLBARS      =   3
      SCROLLTRACK     =   0   'False
      SELECTIONMODE   =   0
      Object.TOOLTIPTEXT     =   ""
      WORDWRAP        =   0   'False
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1485
      Index           =   0
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   1080
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   2619
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   3
         Left            =   1320
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   675
         Width           =   5460
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   1320
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   150
         Width           =   2265
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   8115
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   675
         Width           =   1620
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   1320
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   675
         Width           =   5460
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   5
         Left            =   1320
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   1005
         Width           =   5460
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   4
         Left            =   8115
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   1005
         Width           =   1620
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   7
         Top             =   735
         Width           =   570
      End
      Begin VB.Shape Shape4 
         Height          =   330
         Index           =   0
         Left            =   6945
         Top             =   180
         Width           =   2775
      End
      Begin VB.Shape Shape3 
         Height          =   390
         Index           =   0
         Left            =   6915
         Top             =   150
         Width           =   2835
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
         Left            =   6975
         TabIndex        =   20
         Tag             =   "eb0;et0"
         Top             =   225
         Width           =   2715
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
         Caption         =   "Transaction No"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   4
         Top             =   210
         Width           =   1095
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transact. Date"
         Height          =   195
         Index           =   1
         Left            =   6885
         TabIndex        =   11
         Top             =   735
         Width           =   1065
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   195
         Index           =   5
         Left            =   90
         TabIndex        =   9
         Top             =   1065
         Width           =   630
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Requested By"
         Height          =   195
         Index           =   6
         Left            =   6885
         TabIndex        =   13
         Top             =   1065
         Width           =   1005
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H8000000F&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   285
         Index           =   0
         Left            =   6975
         Tag             =   "et0;et0"
         Top             =   210
         Width           =   2730
      End
   End
End
Attribute VB_Name = "frmCP_BranchOrder_Posting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Private Const pxeMODULENAME = "frmSP_BranchOrder_Posting"
'
'Private WithEvents oTrans As clsBranchOrder
'Private oBranch As clsBranch
'Private oSkin As clsFormSkin
'
'Dim pnIndex As Integer
'Dim pnCtr As Integer
'Dim psBranchNm As String
'Dim pbLoad As Boolean
'
'Private Sub cmdButton_Click(Index As Integer)
'   Dim lsOldProc As String
'   Dim lsRep As String
'   Dim lsTransNox As String
'
'   lsOldProc = "cmdButton_Click"
'   ''On Error GoTo errProc
'
'   With GridEditor1
'      Select Case Index
'      Case 0 'Browse
'         If oTrans.SearchTransaction() Then
'            lsTransNox = oTrans.Master("sTransNox")
'            oTrans.BranchAsDestination = True
'            oTrans.Branch = Left(lsTransNox, Len(oApp.BranchCode))
'            oTrans.InitTransaction
'            If oTrans.OpenTransaction(lsTransNox) Then
'               LoadMaster
'               LoadDetail
'            End If
'         Else
'            If txtField(0).Text = "" Then ClearFields
'         End If
'
'         txtField(6).SetFocus
'      Case 1 'Confirm
'         If pbLoad Then
'            If oTrans.Master("cTranStat") <> xeStatePosted And oTrans.Master("cTranStat") <> xeStateCancelled Then
'               lsRep = MsgBox("Post Transaction!!!", vbYesNo + vbQuestion, "Confrim")
'
'               If lsRep = vbYes Then
'                  If Not oTrans.PostTransaction() Then
'                     MsgBox "Unable to Post Transaction!!!", vbCritical, "Warning"
'                  Else
'                     MsgBox "Transaction Post Successfully!!!", vbInformation, "Confirm"
'                     Label2.Caption = Format(TransStat(oTrans.Master("cTranStat")), ">")
'                     GridEditor1.ColEnabled(6) = False
'                  End If
'               End If
'            Else
'               MsgBox "Unable to Post Transaction!!!" & vbCrLf & _
'                        "Transaction already posted/cancelled!!!", vbCritical, "Warning"
'            End If
'         Else
'            MsgBox "Unable to Post Transaction!!!" & vbCrLf & _
'                   "No Transaction Loaded!!!", vbCritical, "Warning"
'         End If
'      Case 2 'Cancel
'           If pbLoad Then
'            If oTrans.Master("cTranStat") <> xeStatePosted And oTrans.Master("cTranStat") <> xeStateCancelled Then
'               lsRep = MsgBox("Cancel Transaction!!!", vbYesNo + vbQuestion, "Confirm")
'
'               If lsRep = vbYes Then
'                  If Not oTrans.CancelTransaction Then
'                     MsgBox "Unable to Cancel Transaction!!!", vbCritical, "Warning"
'                  Else
'                     MsgBox "Transaction Cancelled Successfully!!!", vbInformation, "Confirm"
'                     If oTrans.OpenTransaction(oTrans.Master("sTransNox")) Then
'                        LoadMaster
'                        LoadDetail
'                     Else
'                        ClearFields
'                     End If
'                  End If
'               End If
'            Else
'               MsgBox "Unable to Cancel Transaction!!!" & vbCrLf & _
'                        "Transaction already cancelled/posted!!!", vbCritical, "Warning"
'            End If
'         Else
'            MsgBox "Unable to Post Transaction!!!" & vbCrLf & _
'                   "No Transaction Loaded!!!", vbCritical, "Warning"
'         End If
'      Case 3 'Close
'         Unload Me
'      End Select
'   End With
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
'
'   GridEditor1.Refresh
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
'   Set oTrans = New clsBranchOrder
'   Set oTrans.AppDriver = oApp
'   oTrans.TransStatus = 10
'   oTrans.BranchAsDestination = True
'   oTrans.InitTransaction
'
'   Set oSkin = New clsFormSkin
'   Set oSkin.AppDriver = oApp
'   Set oSkin.Form = Me
'   oSkin.ApplySkin xeFormTransMaintenance
'
'   Set oBranch = New clsBranch
'   Set oBranch.AppDriver = oApp
'   oBranch.InitRecord
'   oBranch.NewRecord
'
'   InitForm
'   ClearFields
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
'Private Sub GridEditor1_EditorValidate(Cancel As Boolean)
'   With GridEditor1
'      If .Col = 6 Then
'         If CDbl(.TextMatrix(.Row, .Col - 1)) < CDbl(.TextMatrix(.Row, .Col)) Then
'            .TextMatrix(.Row, .Col) = oTrans.Detail(.Row - 1, "nQuantity")
'         End If
'
'         oTrans.Detail(.Row - 1, "nApproved") = CDbl(.TextMatrix(.Row, .Col))
'      End If
'   End With
'End Sub
'
'Private Sub GridEditor1_GotFocus()
'   With GridEditor1
'      .EditorBackColor = oApp.getColor("HT1")
'   End With
'End Sub
'
'Private Sub GridEditor1_LostFocus()
'   With GridEditor1
'      .EditorBackColor = oApp.getColor("EB")
'   End With
'End Sub
'
'Private Sub oTrans_DetailRetrieved(ByVal Index As Integer)
'   With GridEditor1
'      .TextMatrix(.Row, Index) = oTrans.Detail(.Row - 1, IIf(Index = 4, Index + 1, Index))
'   End With
'End Sub
'
'Private Sub oTrans_MasterRetrieved(ByVal Index As Integer)
'   txtField(Index).Text = oTrans.Master(Index)
'End Sub
'
'Private Sub ClearFields()
'   For pnCtr = 0 To txtField.Count - 1
'      txtField(pnCtr).Text = ""
'      txtField(pnCtr).Tag = ""
'   Next
'
'   Label2.Caption = "UNKNOWN"
'
'   With GridEditor1
'      .Rows = 2
'      .Row = 1
'      .Col = 1
'
'      .ColWidth(2) = 3800
'      .ColWidth(3) = 2050
'
'      'empty row
'      .TextMatrix(1, 1) = ""
'      .TextMatrix(1, 2) = ""
'      .TextMatrix(1, 3) = ""
'      .TextMatrix(1, 4) = 0
'      .TextMatrix(1, 5) = 0
'      .TextMatrix(1, 6) = 0
'      .TextMatrix(1, 7) = 0
'   End With
'   pbLoad = False
' End Sub
'
'Private Sub InitForm()
'   Dim lnCtr As Integer
'
'   With GridEditor1
'      .Cols = 8
'      .Rows = 2
'      .Font = "MS Sans Serif"
'
'      'Column Title
'      .TextMatrix(0, 1) = "Barcode"
'      .TextMatrix(0, 2) = "Description"
'      .TextMatrix(0, 3) = "Model"
'      .TextMatrix(0, 4) = "QOH"
'      .TextMatrix(0, 5) = "Qty"
'      .TextMatrix(0, 6) = "Apv"
'      .TextMatrix(0, 7) = "BO"
'
'      .Row = 0
'
'      'Column Alignment
'      For lnCtr = 0 To .Cols - 1
'         .Col = lnCtr
'         .CellFontBold = True
'         .CellAlignment = 3
'         .ColEnabled(lnCtr) = False
'      Next
'
'      .ColEnabled(6) = True
'
'      'Column Width
'      .ColWidth(0) = 300
'      .ColWidth(1) = 2200
'      .ColWidth(4) = 480
'      .ColWidth(5) = 450
'      .ColWidth(6) = 450
'      .ColWidth(7) = 450
'
'      .ColAlignment(1) = 1
'      .ColAlignment(2) = 1
'      .ColAlignment(3) = 1
'      .ColAlignment(4) = flexAlignRightCenter
'      .ColAlignment(5) = flexAlignRightCenter
'      .ColAlignment(6) = flexAlignRightCenter
'      .ColAlignment(7) = flexAlignRightCenter
'
'      .ColNumberOnly(6) = True
'      .EditorBackColor = oApp.getColor("HT1")
'
'      .Col = 1
'      .Row = 1
'   End With
'End Sub
'
'Private Sub txtField_GotFocus(Index As Integer)
'   With txtField(Index)
'      .SelStart = 0
'      .SelLength = Len(.Text)
'      .BackColor = oApp.getColor("HT1")
'   End With
'
'   pnIndex = Index
'End Sub
'
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'   Select Case KeyCode
'   Case vbKeyReturn, vbKeyUp, vbKeyDown
'      Select Case KeyCode
'      Case vbKeyReturn, vbKeyDown
'         If GetFocus = GridEditor1.hwnd Then Exit Sub
'         SetNextFocus
'      Case vbKeyUp
'         SetPreviousFocus
'      End Select
'   End Select
'End Sub
'
'Private Sub LoadMaster()
'   For pnCtr = 0 To txtField.Count - 1
'      Select Case pnCtr
'      Case 0, 7
'         txtField(pnCtr).Text = Format(oTrans.Master(0), "@@@@-@@@@@@")
'         txtField(pnCtr).Tag = txtField(pnCtr).Text
'      Case 1
'         txtField(pnCtr).Text = Format(oTrans.Master(pnCtr), "MMMM DD, YYYY")
'      Case 2
'         txtField(pnCtr).Text = oTrans.Master(2)
'         txtField(pnCtr).Tag = txtField(pnCtr).Text
'      Case 3
'          txtField(pnCtr).Text = oBranch.Master("sAddressx")
'      Case 6
'         txtField(pnCtr).Tag = txtField(pnCtr).Text
'      Case Else
'         txtField(pnCtr).Text = IIf(IsNull(oTrans.Master(pnCtr)), "", oTrans.Master(pnCtr))
'      End Select
'   Next
'   Label2.Caption = Format(TransStat(oTrans.Master("cTranStat")), ">")
'   pbLoad = True
'End Sub
'
'Private Sub LoadDetail()
'   If Not oTrans.UpdateTransaction Then Exit Sub
'   With GridEditor1
'      .Rows = IIf(oTrans.ItemCount = 0, 2, oTrans.ItemCount + 1)
'
'      If .Rows > 22 Then
'         .ColWidth(2) = 3700
'         .ColWidth(3) = 1950
'      Else
'         .ColWidth(2) = 3800
'         .ColWidth(3) = 2050
'      End If
'
'      For pnCtr = 0 To oTrans.ItemCount - 1
'         .TextMatrix(pnCtr + 1, 1) = oTrans.Detail(pnCtr, "sBarrCode")
'         .TextMatrix(pnCtr + 1, 2) = oTrans.Detail(pnCtr, "sDescript")
'         .TextMatrix(pnCtr + 1, 3) = oTrans.Detail(pnCtr, "sModelNme")
'         .TextMatrix(pnCtr + 1, 4) = oTrans.Detail(pnCtr, "nQtyOnHnd")
'         .TextMatrix(pnCtr + 1, 5) = oTrans.Detail(pnCtr, "nQuantity")
'         .TextMatrix(pnCtr + 1, 6) = oTrans.Detail(pnCtr, "nQuantity")
'         .TextMatrix(pnCtr + 1, 7) = oTrans.Detail(pnCtr, "nBackOrdr")
'
'         oTrans.Detail(pnCtr, "nApproved") = CDbl(.TextMatrix(pnCtr + 1, 6))
'      Next
'      .ColEnabled(6) = True
'   End With
'End Sub
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
'      Select Case Index
'      Case 6
'         If Trim(.Text) = "" Then
'            oTrans.Branch = ""
'            oTrans.BranchAsDestination = True
'            oTrans.InitTransaction
'
'            ClearFields
'            psBranchNm = ""
'            Exit Sub
'         End If
'
'         If .Text <> .Tag Then
'            If oBranch.SearchRecord(.Text, False) Then
'               oTrans.Branch = oBranch.Master("sBranchCd")
'               oTrans.BranchAsDestination = True
'               oTrans.InitTransaction
'               oTrans.NewTransaction
'               ClearFields
'
'               .Text = oBranch.Master("sBranchNm")
'               psBranchNm = .Text
'            Else
'               If Trim(.Tag) <> "" Then
'                  .Text = .Tag
'                  Exit Sub
'               End If
'
'               ClearFields
'               .SetFocus
'            End If
'         End If
'
'         .Tag = .Text
'      Case 7
'         If .Text = "" Then
'            ClearFields
'
'            txtField(6).Text = psBranchNm
'            txtField(6).Tag = txtField(6).Text
'            Exit Sub
'         End If
'
'         If .Text <> .Tag Then
'            If oTrans.OpenTransaction(CodeFormat(oApp.BranchCode, .TabIndex)) Then
'               LoadMaster
'               LoadDetail
'            Else
'               ClearFields
'               .SetFocus
'            End If
'         End If
'
'         txtField(6).Text = psBranchNm
'         txtField(6).Tag = txtField(6).Text
'      End Select
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
