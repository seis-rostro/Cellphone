VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmIssuance 
   BorderStyle     =   0  'None
   Caption         =   "Issuance of Order"
   ClientHeight    =   6570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6570
   ScaleWidth      =   11160
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   495
      Index           =   1
      Left            =   75
      Tag             =   "wt0;fb0"
      Top             =   525
      Width           =   9525
      _ExtentX        =   16801
      _ExtentY        =   873
      BackColor       =   12632256
      ClipControls    =   0   'False
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
         Index           =   1
         Left            =   4350
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   90
         Width           =   5070
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
         Index           =   0
         Left            =   1365
         MaxLength       =   11
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   90
         Width           =   2085
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Branch"
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
         Index           =   1
         Left            =   3600
         TabIndex        =   2
         Top             =   150
         Width           =   615
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   0
         Top             =   150
         Width           =   1185
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   9810
      TabIndex        =   7
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
      Picture         =   "frmIssuance.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   9810
      TabIndex        =   5
      Top             =   540
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Load"
      AccessKey       =   "L"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmIssuance.frx":077A
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   5475
      Index           =   0
      Left            =   75
      Tag             =   "wt0;fb0"
      Top             =   1035
      Width           =   9525
      _ExtentX        =   16801
      _ExtentY        =   9657
      BackColor       =   12632256
      BorderStyle     =   1
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   5310
         Left            =   60
         TabIndex        =   4
         Tag             =   "et0;eb0;et0;bc2"
         Top             =   45
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   9366
         _Version        =   393216
         FocusRect       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   9810
      TabIndex        =   6
      Top             =   1170
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Retrieve"
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
      Picture         =   "frmIssuance.frx":0EF4
   End
End
Attribute VB_Name = "frmIssuance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Private Const pxeMODULENAME = "frmIssuance"
'
'Private oFormBranch As frmCP_TransferOrder
'Private poProgress As clsSpeedometer
'Private oSkin As clsFormSkin
'
'Dim pnCtr As Integer
'Dim pbDisplayed As Boolean
'Dim pbSortTrans As Boolean
'Dim pbSortBranch As Boolean
'
'Private Sub cmdButton_Click(Index As Integer)
'   Dim lsOldProc As String
'
'   lsOldProc = "cmdButton_Click"
'   On Error GoTo errProc
'
'   Select Case Index
'   Case 0
'      If Not pbDisplayed Then Exit Sub
'      With MSFlexGrid1
'         If .TextMatrix(.Row, 5) = 0 Then
'            Select Case .TextMatrix(.Row, 3)
''            Case "Customer"
''               InitSPWholeSale
'            Case "Branch"
'               InitSPTransfer
'            End Select
'         Else
'            MsgBox "Transaction has been modified!!!" & vbCrLf & _
'                   "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
'         End If
'      End With
'   Case 1
'      Unload Me
'   Case 2
'      LoadData
'   End Select
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " & Index & " )", True
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
'   Set oFormBranch = New frmSP_TransferOrder
'
'   Set oSkin = New clsFormSkin
'   Set oSkin.AppDriver = oApp
'   Set oSkin.Form = Me
'   oSkin.ApplySkin xeFormTransDetail
'
'   InitGrid
'   LoadData
'
'   txtField(0).Text = ""
'   txtField(1).Text = ""
'
'   pbSortTrans = False
'   pbSortBranch = False
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " & " )", True
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'   Set oSkin = Nothing
'   Set poProgress = Nothing
'   Set oFormBranch = Nothing
'   Set oFormCustomer = Nothing
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
'Private Sub InitGrid()
'   With MSFlexGrid1
'      .Cols = 5
'      .Font = "MS Sans Serif"
'
'      'column title
'      .TextMatrix(0, 1) = "Trans. No."
'      .TextMatrix(0, 2) = "Branch"
'      .TextMatrix(0, 3) = "Date of Order"
'      .Row = 0
'
'      'column alignment
'      For pnCtr = 0 To .Cols - 1
'         .Col = pnCtr
'         .CellFontBold = True
'         .CellAlignment = 1
'      Next
'
'      'column width
'      .ColWidth(0) = 330
'      .ColWidth(1) = 1500
'      .ColWidth(2) = 1000
'      .ColWidth(3) = 2000
'      .ColWidth(4) = 0
'
'      .ColAlignment(1) = 1
'      .ColAlignment(2) = 1
'      .ColAlignment(3) = 1
'      .ColAlignment(4) = 1
'   End With
'End Sub
'
'Private Sub LoadData()
'   Dim lors As ADODB.Recordset
'   Dim lsSQL As String
'   Dim lsOldProc As String
'
'   lsOldProc = "LoadData"
'   On Error GoTo errProc
'
'   Set poProgress = New clsSpeedometer
'   poProgress.InitProgress
'   poProgress.PrimaryRemarks = "Loading Data"
'   DoEvents
'
'   lsSQL = "SELECT Distinct" _
'               & "  a.sTransNox" _
'               & ", b.sBranchNm" _
'               & ", a.dTransact" _
'            & " FROM CP_Branch_Order_Master a" _
'               & ", Branch b" _
'               & ", CP_Branch_Order_Detail c" _
'            & " WHERE LEFT(a.sTransNox, 2) = b.sBranchCd" _
'               & " AND a.cTranStat = " & strParm(xeStatePosted) _
'               & " AND a.sTransNox = c.sTransNox" _
'               & " AND c.cCanceled = '0'" _
'               & " AND c.nApproved > c.nIssuedxx)" _
'            & " ORDER BY sTransNox"
'
'   Set lors = New ADODB.Recordset
'   lors.CacheSize = 1
'   lors.Open lsSQL, oApp.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText Or adAsyncFetchNonBlocking
'
'   pbDisplayed = False
'   With MSFlexGrid1
'      .Rows = 2
'      .Row = 1
'      .ColWidth(2) = 5240
'      .TextMatrix(1, 0) = 1
'      .TextMatrix(1, 5) = 0
'
'      For pnCtr = 1 To .Cols - 1
'         .TextMatrix(1, pnCtr) = Empty
'      Next
'
'      If Not lors.EOF Then
'         poProgress.SecMaxValue = lors.RecordCount + 1
'         poProgress.PriMaxValue = 1
'
'         Do While lors.EOF = False
'            .TextMatrix(.Row, 0) = .Row
'            .TextMatrix(.Row, 1) = Format(lors("sTransNox"), "@@@@-@@@@@@")
'            .TextMatrix(.Row, 2) = lors("sBranchNm")
'            .TextMatrix(.Row, 3) = Format(lors("dTransact"), "MMMM DD, YYYY")
'            .TextMatrix(.Row, 4) = 0
'
'            .Rows = .Rows + 1
'            .Row = .Rows - 1
'            If .Rows > 29 Then .ColWidth(2) = 5000
'            poProgress.MoveProgress lors("sTransNox"), lors("sBranchNm")
'            lors.MoveNext
'         Loop
'         .Rows = .Rows - 1
'         pbDisplayed = True
'      End If
'
'      .Row = 1
'      .Col = 1
'      .ColSel = .Cols - 1
'   End With
'
'   lors.Close
'   poProgress.CloseProgress
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " & " )"
'End Sub
'
'Private Function ResultingGrid(iKeyAscii%) As String
'   Dim sLeft As String
'   Dim sRight As String
'   Dim sResult As String
'
'   On Error Resume Next
'
'   With MSFlexGrid1
'      sLeft = Left$(.TextMatrix(.Row, 1), 0)
'      sRight = Mid$(.TextMatrix(.Row, 1), Len(.TextMatrix(.Row, 1)) + 1)
'   End With
'
'   sResult = sLeft & Chr$(iKeyAscii) & sRight
'   ResultingGrid = sResult
'End Function
'
'Private Function SearchOn(ByVal lsSeek, ByVal lnIndex) As Boolean
'   Dim lnCtr As Long
'   Dim lbFound As Boolean
'
'   lbFound = False
'   With MSFlexGrid1
'      For lnCtr = 1 To .Rows
'         If StrComp(Left(.TextMatrix(lnCtr, lnIndex + 1), Len(lsSeek)), lsSeek, vbTextCompare) >= 0 Then
'            .TopRow = lnCtr
'            .Row = lnCtr
'            .RowSel = lnCtr
'            .ColSel = .Cols - 1
'            lbFound = True
'            Exit For
'         End If
'      Next
'   End With
'   SearchOn = lbFound
'End Function
'
'Private Sub MSFlexGrid1_DblClick()
'   Dim lsOldProc As String
'
'   lsOldProc = "MSFlexGrid1_DblClick"
'   On Error GoTo errProc
'
'   With MSFlexGrid1
'      If .TextMatrix(.Row, 5) = 0 Then
'         Select Case .TextMatrix(.Row, 3)
'         Case "Branch"
'            InitCPTransfer
'         End Select
'      Else
'         MsgBox "Transaction has been modified!!!" & vbCrLf & _
'                "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
'      End If
'   End With
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " & " )", True
'End Sub
'
'Private Sub MSFlexGrid1_GotFocus()
'   MSFlexGrid1.BackColorSel = &HA4A36A
'End Sub
'
'Private Sub MSFlexGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
'   Dim lsOldProc As String
'
'   lsOldProc = "MSFlexGrid1_KeyDown"
'   On Error GoTo errProc
'
'   If KeyCode = vbKeySpace Then
'      With MSFlexGrid1
'         If .TextMatrix(.Row, 5) = 0 Then
'            Select Case .TextMatrix(.Row, 3)
'            Case "Branch"
'               InitSPTransfer
'            End Select
'         Else
'            MsgBox "Transaction has been modified!!!" & vbCrLf & _
'                   "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
'         End If
'      End With
'      KeyCode = 0
'   End If
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " _
'                       & " " & KeyCode _
'                       & ", " & Shift _
'                       & " )", True
'End Sub
'
'Private Sub MSFlexGrid1_LostFocus()
'   MSFlexGrid1.BackColorSel = &H8000000D
'End Sub
'
'Private Function ResultingText(iKeyAscii%, lnIndex) As String
'   Dim sLeft As String
'   Dim sSel As String
'   Dim sRight As String
'   Dim sResult As String
'
'   On Error Resume Next
'
'   With txtField(lnIndex)
'      sLeft = Left$(.Text, .SelStart)
'      sSel = Mid$(.Text, .SelStart + 1, .SelLength)
'      sRight = Mid$(.Text, .SelStart + .SelLength + 1)
'   End With
'
'   Select Case iKeyAscii
'   Case vbKeyBack
'      If Len(sSel) = 0 Then
'         sResult = MinusRightChar(sLeft) & sRight
'      Else
'         sResult = sLeft & sRight
'      End If
'   Case vbKeyDelete
'      If Len(sSel) = 0 Then
'         sResult = sLeft & MinusLeftChar(sRight)
'      Else
'         sResult = sLeft & sRight
'      End If
'   Case Else
'      sResult = sLeft & Chr$(iKeyAscii) & sRight
'   End Select
'   ResultingText = sResult
'End Function
'
'Private Function MinusLeftChar(ByVal sGiven As String) As String
'   On Error Resume Next
'
'   If Len(sGiven) = 0 Then
'      MinusLeftChar = ""
'   Else
'      MinusLeftChar = Mid$(sGiven, 2)
'   End If
'End Function
'
'Private Function MinusRightChar(ByVal sGiven As String) As String
'   On Error Resume Next
'
'   If Len(sGiven) = 0 Then
'      MinusRightChar = ""
'   Else
'      MinusRightChar = Left$(sGiven, Len(sGiven) - 1)
'   End If
'End Function
'
'Private Sub txtField_GotFocus(Index As Integer)
'   With MSFlexGrid1
'      Select Case Index
'      Case 0
'         If Not pbSortTrans And pbDisplayed Then
'            .Col = 1
'            .ColSel = .Cols - 1
'            .Sort = 7
'            pbSortTrans = True
'            pbSortBranch = False
'
'            For pnCtr = 1 To .Rows - 1
'               .TextMatrix(pnCtr, 0) = pnCtr
'            Next
'
'            .TopRow = .Row
'         End If
'      Case 1
'         If Not pbSortBranch And pbDisplayed Then
'            .Col = 2
'            .Sort = 7
'            .Col = 1
'            .ColSel = .Cols - 1
'            pbSortBranch = True
'            pbSortTrans = False
'
'            For pnCtr = 1 To .Rows - 1
'               .TextMatrix(pnCtr, 0) = pnCtr
'            Next
'
'            .TopRow = .Row
'         End If
'      End Select
'   End With
'End Sub
'
'Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'   Dim lsSearchOn As String
'
'   On Error Resume Next
'
'   With MSFlexGrid1
'      If pbDisplayed = False Then Exit Sub
'      Select Case KeyCode
'      Case vbKeyDown, vbKeyUp, vbKeyPageDown, vbKeyPageUp
''            MSFlexGrid1.SetFocus
'         Exit Sub
'      Case vbKeyReturn
'         If Shift = 2 Then
'            If .TextMatrix(.Row, 5) = 0 Then
'               Select Case .TextMatrix(.Row, 3)
''               Case "Customer"
''                  InitSPWholeSale
'               Case "Branch"
'                  InitSPTransfer
'               End Select
'            Else
'               MsgBox "Transaction has been modified!!!" & vbCrLf & _
'                      "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
'            End If
'         End If
'      Case Else
'         If KeyCode <> vbKeyDelete Then Exit Sub
'      End Select
'
'      If vbKeyReturn Then Exit Sub
'      lsSearchOn = ResultingText(KeyCode, Index)
'      SearchOn lsSearchOn, Index
'   End With
'End Sub
'
'Private Sub txtField_KeyPress(Index As Integer, KeyAscii As Integer)
'   Dim lsSearchOn As String
'
'   On Error Resume Next
'
'   If pbDisplayed = False Then Exit Sub
'   If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then Exit Sub
'
'   lsSearchOn = ResultingText(KeyAscii, Index)
'   If SearchOn(lsSearchOn, Index) = False Then KeyAscii = 0
'End Sub
'
'Private Sub InitCPTransfer()
'   Dim lsSQL As String
'   Dim lrs As ADODB.Recordset
'   Dim lsOldProc As String
'
'   lsOldProc = "InitSPTransfer"
'   On Error GoTo errProc
'
'   With MSFlexGrid1
'      lsSQL = "Select" _
'                  & "  a.sTransNox" _
'                  & ", LEFT(a.sTransNox, 2) sBranchCd" _
'                  & ", b.nApproved - b.nIssuedxx xQuantity" _
'                  & ", c.sBranchNm" _
'                  & ", CONCAT(c.sAddressx, ', ', f.sTownName, ', ', g.sProvName) xAddressx" _
'                  & ", d.sBarrCode" _
'                  & ", e.nQtyOnHnd" _
'                  & ", e.nSelPrice"
'
'      lsSQL = lsSQL _
'               & " From CP_Branch_Order_Master a" _
'                  & ", CP_Branch_Order_Detail b" _
'                  & ", Branch c" _
'                  & ", CP_Inventory d" _
'                  & ", CP_Inventory_Master e" _
'                  & ", TownCity f" _
'                  & ", Province g" _
'               & " Where a.sTransNox = b.sTransNox" _
'                  & " And LEFT(a.sTransNox, 2) = c.sBranchCd" _
'                  & " And b.sStockIDx = d.sStockIDx" _
'                  & " And a.sTransNox = " & strParm(Left(.TextMatrix(.Row, 1), 4) & Right(.TextMatrix(.Row, 1), 6)) _
'                  & " And d.sStockIDx = e.sStockIDx" _
'                  & " And e.sBranchCd = " & strParm(oApp.BranchCode) _
'                  & " And c.sTownIDxx = f.sTownIDxx" _
'                  & " And f.sProvIDxx = g.sProvIDxx" _
'                  & " And b.cCanceled = '0'" _
'                  & " And b.nApproved > b.nIssuedxx" _
'                  & " And e.nQtyOnHnd > 0" _
'                  & " Order By b.nEntryNox"
'
'      Set lrs = New ADODB.Recordset
'      lrs.Open lsSQL, oApp.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'      Load oFormBranch
'      With oFormBranch.CPTransfer
'         .Master("sDestinat") = lrs("sBranchCd")
'         .Master("xAddressx") = lrs("xAddressx")
'         .Master(2) = lrs("sBranchNm")
'
'         For pnCtr = 0 To lrs.RecordCount - 1
'            .Detail(pnCtr, "sBarrCode") = lrs("sBarrCode")
'            .Detail(pnCtr, "nQuantity") = IIf(lrs("xQuantity") > lrs("nQtyOnHnd"), lrs("nQtyOnHnd"), lrs("xQuantity"))
'
'            lrs.MoveNext
'            Call .AddDetail
'         Next
'
'         Call .DeleteDetail(pnCtr)
'      End With
'
'      Me.Hide
'      oFormBranch.Tag = "mnuWholeSale"
'      oFormBranch.Show 1
'      If Not oFormBranch.Cancel Then .TextMatrix(.Row, 5) = 1
'      Me.Show
'   End With
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " & " )"
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
