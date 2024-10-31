VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmCP_CustomerOrder_Cancelling 
   BorderStyle     =   0  'None
   Caption         =   "Spareparts Purchasing Register"
   ClientHeight    =   7800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11940
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7800
   ScaleWidth      =   11940
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   10590
      TabIndex        =   20
      Top             =   4335
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
      Picture         =   "frmCP_CustomerOrder_Cancelling.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   10590
      TabIndex        =   17
      Top             =   2445
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
      Picture         =   "frmCP_CustomerOrder_Cancelling.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   10590
      TabIndex        =   18
      Top             =   3075
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
      Picture         =   "frmCP_CustomerOrder_Cancelling.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   10590
      TabIndex        =   19
      Top             =   3705
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Refresh"
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
      Picture         =   "frmCP_CustomerOrder_Cancelling.frx":166E
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   525
      Index           =   1
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   10245
      _ExtentX        =   18071
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
         Left            =   4725
         MaxLength       =   50
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   90
         Width           =   5415
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
         Left            =   915
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   90
         Width           =   2265
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Company Name"
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
         Left            =   3315
         TabIndex        =   2
         Top             =   135
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Trans.No."
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
         Width           =   1365
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1755
      Index           =   0
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   1080
      Width           =   10260
      _ExtentX        =   18098
      _ExtentY        =   3096
      BackColor       =   12632256
      Enabled         =   0   'False
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   915
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   150
         Width           =   2265
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   7305
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   675
         Width           =   1905
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   915
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   675
         Width           =   5235
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   3
         Left            =   915
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   1005
         Width           =   5235
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   4
         Left            =   915
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   1335
         Width           =   5235
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   645
         Index           =   5
         Left            =   7305
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   1005
         Width           =   2835
      End
      Begin VB.Shape Shape4 
         Height          =   330
         Index           =   0
         Left            =   7335
         Top             =   180
         Width           =   2775
      End
      Begin VB.Shape Shape3 
         Height          =   390
         Index           =   0
         Left            =   7305
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
         Left            =   7365
         TabIndex        =   21
         Tag             =   "eb0;et0"
         Top             =   225
         Width           =   2715
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   315
         Left            =   1005
         Tag             =   "et0;ht2"
         Top             =   255
         Width           =   2265
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trans. No."
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   4
         Top             =   210
         Width           =   750
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transact. Date"
         Height          =   195
         Index           =   1
         Left            =   6195
         TabIndex        =   12
         Top             =   735
         Width           =   1065
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   6
         Top             =   735
         Width           =   660
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   10
         Top             =   1380
         Width           =   570
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   195
         Index           =   5
         Left            =   6195
         TabIndex        =   14
         Top             =   1050
         Width           =   630
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Full Name"
         Height          =   195
         Index           =   6
         Left            =   90
         TabIndex        =   8
         Top             =   1065
         Width           =   705
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H8000000F&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   285
         Index           =   0
         Left            =   7365
         Tag             =   "et0;et0"
         Top             =   210
         Width           =   2730
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4800
      Left            =   90
      TabIndex        =   16
      Tag             =   "et0;eb0;et0;bc2"
      Top             =   2880
      Width           =   10260
      _ExtentX        =   18098
      _ExtentY        =   8467
      _Version        =   393216
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "frmCP_CustomerOrder_Cancelling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Private Const pxeMODULENAME = "frmSP_CustomerOrder_Cancelling"
'
'Private WithEvents oTrans As clsCustomerOrder
'Private oSkin As clsFormSkin
'
'Dim pnCtr As Integer
'Dim pbLoaded As Boolean
'
'Private Sub cmdButton_Click(Index As Integer)
'   Dim lsOldProc As String
'   Dim lsRep As String
'
'   lsOldProc = "cmdButton_Click"
'   On Error GoTo errProc
'
'   With MSFlexGrid1
'      Select Case Index
'      Case 0 'Browse
'         If oTrans.SearchTransaction() Then
'            LoadMaster
'            LoadDetail
'         End If
'         txtField(7).SetFocus
'      Case 1 'Save
'         If pbLoaded Then
'            lsRep = MsgBox("Are you sure you want to save this transaction!!!", vbYesNo + vbQuestion, "Confrim")
'
'            If lsRep = vbYes Then
'               If Not SaveTransaction Then
'                  MsgBox "Unable to Save Transaction!!!", vbCritical, "Warning"
'               Else
'                  MsgBox "Transaction Save Successfully!!!", vbInformation, "Confirm"
'                  oTrans.InitTransaction
'                  oTrans.NewTransaction
'                  txtField(7).SetFocus
'                  InitEntry
'               End If
'            End If
'         Else
'            MsgBox "Unable to Save Transaction!!!" & vbCrLf & _
'                   "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
'         End If
'      Case 2 'Refresh
'         If pbLoaded Then
'            InitEntry
'            If oTrans.OpenTransaction(oTrans.Master("sTransNox")) Then
'               LoadMaster
'               LoadDetail
'            End If
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
'  oApp.MenuName = Me.Tag
'   Me.ZOrder 0
'
'   MSFlexGrid1.Refresh
'End Sub
'
'Private Sub Form_Load()
'   Dim lsOldProc As String
'
'   lsOldProc = "Form_Load"
'   On Error GoTo errProc
'
'   CenterChildForm mdiMain, Me
'
'   Set oTrans = New clsCustomerOrder
'   Set oTrans.AppDriver = oApp
'
'   oTrans.TransStatus = xeStateClosed
'   oTrans.InitTransaction
'
'   InitGrid
'   InitEntry
'
'   Set oSkin = New clsFormSkin
'   Set oSkin.AppDriver = oApp
'   Set oSkin.Form = Me
'   oSkin.ApplySkin xeFormTransMaintenance
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
'Private Sub oTrans_DetailRetrieved(ByVal Index As Integer)
'   With MSFlexGrid1
'      .TextMatrix(.Row, Index) = oTrans.Detail(.Row - 1, IIf(Index = 4, Index + 1, Index))
'   End With
'End Sub
'
'Private Sub oTrans_MasterRetrieved(ByVal Index As Integer)
'   txtField(Index).Text = oTrans.Master(Index)
'End Sub
'
'Private Sub InitEntry()
'   For pnCtr = 0 To txtField.Count - 1
'     txtField(pnCtr).Text = ""
'   Next
'
'   txtField(6).Tag = ""
'   txtField(7).Tag = ""
'   Label2.Caption = "UNKNOWN"
'
'   With MSFlexGrid1
'      .Rows = 2
'
'      .ColWidth(2) = 3550
'      .ColWidth(3) = 2000
'
'      'empty row
'      .TextMatrix(1, 0) = 1
'      .TextMatrix(1, 1) = ""
'      .TextMatrix(1, 2) = ""
'      .TextMatrix(1, 3) = ""
'      .TextMatrix(1, 4) = 0
'      .TextMatrix(1, 5) = 0
'      .TextMatrix(1, 6) = 0
'      .TextMatrix(1, 7) = "NO"
'      .TextMatrix(1, 8) = 0
'
'      .Row = 1
'      .Col = 1
'      .ColSel = .Cols - 1
'   End With
'
'   pbLoaded = False
' End Sub
'
'Private Sub InitGrid()
'   Dim lnCtr As Integer
'
'   With MSFlexGrid1
'      .Cols = 9
'      .Rows = 2
'      .Font = "MS Sans Serif"
'
'      'Column Title
'      .TextMatrix(0, 1) = "Barcode"
'      .TextMatrix(0, 2) = "Description"
'      .TextMatrix(0, 3) = "Model"
'      .TextMatrix(0, 4) = "QOH"
'      .TextMatrix(0, 5) = "Qty."
'      .TextMatrix(0, 6) = "Iss."
'      .TextMatrix(0, 7) = "Cncl"
'
'      .Row = 0
'
'      'Column Alignment
'      For lnCtr = 0 To .Cols - 1
'         .Col = lnCtr
'         .CellFontBold = True
'         .CellAlignment = 3
'      Next
'
'      'Column Width
'      .ColWidth(0) = 300
'      .ColWidth(1) = 1900
'      .ColWidth(4) = 600
'      .ColWidth(5) = 600
'      .ColWidth(6) = 600
'      .ColWidth(7) = 600
'      .ColWidth(8) = 0
'
'      .ColAlignment(1) = 1
'      .ColAlignment(2) = 1
'      .ColAlignment(3) = 1
'      .ColAlignment(7) = 3
'   End With
'End Sub
'
'Private Sub txtField_GotFocus(Index As Integer)
'   With txtField(Index)
'      .SelStart = 0
'      .SelLength = Len(.Text)
'      .BackColor = oApp.getColor("HT1")
'   End With
'End Sub
'
'Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'   Dim lsOldProc As String
'
'   lsOldProc = "txtField_KeyDown"
'   On Error GoTo errProc
'
'   Select Case Index
'   Case 6, 7
'      If KeyCode = vbKeyF3 Then
'         If oTrans.SearchTransaction() = True Then
'            LoadMaster
'            LoadDetail
'         Else
'            If txtField(0).Text <> "" Then InitEntry
'         End If
'         txtField(7).SetFocus
'         KeyCode = 0
'      End If
'   End Select
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
'         If GetFocus = MSFlexGrid1.hwnd Then Exit Sub
'         SetNextFocus
'      Case vbKeyUp
'         SetPreviousFocus
'      End Select
'   End Select
'End Sub
'
'Private Sub LoadMaster()
'   For pnCtr = 2 To 5
'      txtField(pnCtr).Text = IIf(IsNull(oTrans.Master(pnCtr)), "", oTrans.Master(pnCtr))
'   Next
'
'   txtField(0).Text = Format(oTrans.Master(0), "@@@@-@@@@@@")
'   txtField(1).Text = Format(oTrans.Master(1), "MMMM DD, YYYY")
'   Label2.Caption = Format(TransStat(oTrans.Master("cTranStat")), ">")
'
'   txtField(7).Text = txtField(0).Text
'   txtField(6).Text = txtField(2).Text
'
'   txtField(7).Tag = txtField(0).Text
'   txtField(6).Tag = txtField(2).Text
'
'   pbLoaded = True
'End Sub
'
'Private Sub LoadDetail()
'   Dim lnCtr As Integer
'
'   With MSFlexGrid1
'      .Rows = IIf(oTrans.ItemCount = 0, 2, oTrans.ItemCount + 1)
'
'      If .Rows > 22 Then
'         .ColWidth(2) = 3450
'         .ColWidth(3) = 1900
'      Else
'         .ColWidth(2) = 3550
'         .ColWidth(3) = 2000
'      End If
'
'      For pnCtr = 0 To oTrans.ItemCount - 1
'         .TextMatrix(pnCtr + 1, 0) = oTrans.Detail(pnCtr, "nEntryNox")
'         .TextMatrix(pnCtr + 1, 1) = oTrans.Detail(pnCtr, "sBarrCode")
'         .TextMatrix(pnCtr + 1, 2) = oTrans.Detail(pnCtr, "sDescript")
'         .TextMatrix(pnCtr + 1, 3) = IIf(IsNull(oTrans.Detail(pnCtr, "sModelNme")), "", oTrans.Detail(pnCtr, "sModelNme"))
'         .TextMatrix(pnCtr + 1, 4) = oTrans.Detail(pnCtr, "nQtyOnHnd")
'         .TextMatrix(pnCtr + 1, 5) = oTrans.Detail(pnCtr, "nQuantity")
'         .TextMatrix(pnCtr + 1, 6) = oTrans.Detail(pnCtr, "nIssuedxx")
'         .TextMatrix(pnCtr + 1, 7) = IIf(oTrans.Detail(pnCtr, "cCanceled") = 0, "NO", "YES")
'         .TextMatrix(pnCtr + 1, 8) = oTrans.Detail(pnCtr, "cCanceled")
'      Next
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
'   On Error GoTo errProc
'
'   With txtField(Index)
'      Select Case Index
'      Case 6, 7
'         If .Text = "" Then
'            InitEntry
'            GoTo endProc
'         End If
'
'         If .Text <> .Tag Then
'            If oTrans.SearchTransaction(IIf(Index = 7, CodeFormat(oApp.BranchCode, .Text) _
'               , .Text) _
'               , IIf(Index = 7, True, False)) = True Then
'
'               LoadMaster
'               LoadDetail
'            Else
'               InitEntry
'               .SetFocus
'            End If
'         End If
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
'Private Sub MSFlexGrid1_DblClick()
'   If Not pbLoaded Then Exit Sub
'   With MSFlexGrid1
'      If .TextMatrix(.Row, 8) = 1 Then Exit Sub
'      If oTrans.Detail(.Row - 1, "nQuantity") = oTrans.Detail(.Row - 1, "nIssuedxx") Then Exit Sub
'      .TextMatrix(.Row, 7) = IIf(.TextMatrix(.Row, 7) = "NO", "YES", "NO")
'      oTrans.Detail(.Row - 1, "cCanceled") = IIf(.TextMatrix(.Row, 7) = "NO", 0, 1)
'   End With
'End Sub
'
'Private Sub MSFlexGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
'   If Not pbLoaded Then Exit Sub
'   If KeyCode = vbKeySpace Then
'      With MSFlexGrid1
'         If .TextMatrix(.Row, 8) = 1 Then Exit Sub
'         If oTrans.Detail(.Row - 1, "nQuantity") = oTrans.Detail(.Row - 1, "nIssuedxx") Then Exit Sub
'         .TextMatrix(.Row, 7) = IIf(.TextMatrix(.Row, 7) = "NO", "YES", "NO")
'         oTrans.Detail(.Row - 1, "cCanceled") = IIf(.TextMatrix(.Row, 7) = "NO", 0, 1)
'      End With
'      KeyCode = 0
'   End If
'End Sub
'
'Private Sub MSFlexGrid1_GotFocus()
'   MSFlexGrid1.BackColorSel = &HA4A36A
'End Sub
'
'Private Sub MSFlexGrid1_LostFocus()
'   MSFlexGrid1.BackColorSel = &H8000000D
'End Sub
'
'Private Function SaveTransaction() As Boolean
'   Dim lbIssued As Boolean
'   Dim lnCtr As Integer
'   Dim lnRow As Long
'   Dim lsOldProc As String
'
'   lsOldProc = "SaveTransaction"
'   On Error GoTo errProc
'
'   SaveTransaction = False
'
'   oApp.BeginTrans
'   lbIssued = True
'   For pnCtr = 0 To oTrans.ItemCount - 1
'      If oTrans.Detail(pnCtr, "cCanceled") = 1 Then
'         If oApp.Execute("Update SP_Customer_Order_Detail Set" _
'                                    & "  cCanceled = " & strParm(oTrans.Detail(pnCtr, "cCanceled")) _
'                                    & ", dModified = " & dateParm(oApp.ServerDate) _
'                                 & " Where sTransNox = " & strParm(oTrans.Master("sTransNox")) _
'                                    & " And nEntryNox = " & strParm(oTrans.Detail(pnCtr, "nEntryNox")) _
'                                    & " And sPartsIDx = " & strParm(oTrans.Detail(pnCtr, "sPartsIDx")) _
'                                 , "SP_Customer_Order_Detail") = 0 Then
'            MsgBox "Unable to Update SP_Branch_Order_Detail", vbCritical, "Warning"
'            GoTo endProcWithRoll
'         End If
'      Else
'         If oTrans.Detail(pnCtr, "nQuantity") <> oTrans.Detail(pnCtr, "nIssuedxx") Then lbIssued = False
'      End If
'   Next
'
'   If lbIssued Then
'      If oApp.Execute("Update SP_Customer_Order_Master Set" _
'                                 & "  cTranStat = " & strParm(xeStatePosted) _
'                                 & ", dModified = " & dateParm(oApp.ServerDate) _
'                              & " Where sTransNox = " & strParm(oTrans.Master("sTransNox")) _
'                              , "SP_Customer_Order_Master") = 0 Then
'         MsgBox "Unable to Update SP_Customer_Order_Master", vbCritical, "Warning"
'         GoTo endProcWithRoll
'      End If
'   End If
'
'   oApp.CommitTrans
'   SaveTransaction = True
'
'endProc:
'   Exit Function
'endProcWithRoll:
'   oApp.RollbackTrans
'   GoTo endProc
'errProc:
'   oApp.RollbackTrans
'   ShowError lsOldProc & "( " & " )"
'End Function
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
