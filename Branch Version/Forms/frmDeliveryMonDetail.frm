VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmDeliveryMonDetail 
   BorderStyle     =   0  'None
   Caption         =   "Guanzon Delivery (Delivery Detail)"
   ClientHeight    =   8160
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11880
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame3 
      Height          =   7440
      Left            =   4875
      Tag             =   "wt0;fb0"
      Top             =   570
      Width           =   5445
      _ExtentX        =   9604
      _ExtentY        =   13123
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtDetail 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   1905
         TabIndex        =   9
         Text            =   "M0W115000102"
         Top             =   1125
         Width           =   1800
      End
      Begin VB.CommandButton cmdAdd 
         Appearance      =   0  'Flat
         Caption         =   "Add Other Detail"
         Height          =   315
         Left            =   3750
         TabIndex        =   10
         Top             =   1140
         Width           =   1590
      End
      Begin VB.TextBox txtDetail 
         Appearance      =   0  'Flat
         Height          =   975
         Index           =   1
         Left            =   1125
         TabIndex        =   7
         Text            =   "UEMI Laoag"
         Top             =   120
         Width           =   4215
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   5790
         Left            =   60
         TabIndex        =   11
         Top             =   1500
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   10213
         _Version        =   393216
         FocusRect       =   0
         SelectionMode   =   1
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Refer No"
         Height          =   195
         Index           =   1
         Left            =   1155
         TabIndex        =   8
         Top             =   1185
         Width           =   645
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
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
         Left            =   120
         TabIndex        =   6
         Top             =   195
         Width           =   750
      End
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   1350
      Left            =   210
      Tag             =   "wt0;fb0"
      Top             =   690
      Width           =   4530
      _ExtentX        =   7990
      _ExtentY        =   2381
      BackColor       =   12632256
      ClipControls    =   0   'False
      BorderStyle     =   2
      Begin VB.TextBox txtRoute 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1200
         TabIndex        =   1
         Text            =   "UEMI Laoag"
         Top             =   105
         Width           =   3165
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Include ALL in Delivery"
         Height          =   375
         Left            =   1200
         TabIndex        =   4
         Tag             =   "et0;fb0"
         Top             =   825
         Width           =   2250
      End
      Begin VB.ComboBox cmbSource 
         Height          =   315
         ItemData        =   "frmDeliveryMonDetail.frx":0000
         Left            =   1200
         List            =   "frmDeliveryMonDetail.frx":001C
         TabIndex        =   3
         Text            =   "Motorcycle"
         Top             =   480
         Width           =   1800
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Source"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   2
         Top             =   525
         Width           =   510
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Route"
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   0
         Top             =   180
         Width           =   435
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   7440
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   570
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   13123
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   5790
         Left            =   75
         TabIndex        =   5
         Top             =   1515
         Width           =   4560
         _ExtentX        =   8043
         _ExtentY        =   10213
         _Version        =   393216
         FocusRect       =   0
         SelectionMode   =   1
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   10530
      TabIndex        =   13
      Top             =   1200
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
      Picture         =   "frmDeliveryMonDetail.frx":0070
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   10530
      TabIndex        =   12
      Top             =   570
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Ok"
      AccessKey       =   "O"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmDeliveryMonDetail.frx":07EA
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle No"
      Height          =   195
      Index           =   9
      Left            =   375
      TabIndex        =   14
      Top             =   900
      Width           =   780
   End
End
Attribute VB_Name = "frmDeliveryMonDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pxeMODULENAME = "frmDeliveryMonDetail"
Private WithEvents oTrans As clsDeliveryMonitoring
Attribute oTrans.VB_VarHelpID = -1

Private oSkin As clsFormSkin
Private p_nEntryNo As Integer
Private p_oRSRoutex As Recordset
Private p_sBranchCd As String
Private pbLoaded As Boolean
Private p_nRow As Integer

Private Sub Check2_Click()
   Dim lnCtr As Integer
   
   If Check2.Value = Checked Then
      If p_oRSRoutex.EOF Then Exit Sub
      With MSFlexGrid1
         .Rows = p_oRSRoutex.RecordCount + 1
         
         p_oRSRoutex.MoveFirst
         For lnCtr = 0 To p_oRSRoutex.RecordCount - 1
            p_oRSRoutex("cIncluded") = xeYes
            .TextMatrix(lnCtr + 1, 3) = IIf(p_oRSRoutex("cIncluded") = 0, "NO", "YES")
            p_oRSRoutex.MoveNext
         Next
      End With
   Else
      With MSFlexGrid1
         .Rows = p_oRSRoutex.RecordCount + 1
         
         p_oRSRoutex.MoveFirst
         For lnCtr = 0 To p_oRSRoutex.RecordCount - 1
            p_oRSRoutex("cIncluded") = xeNo
            .TextMatrix(lnCtr + 1, 3) = IIf(p_oRSRoutex("cIncluded") = 0, "NO", "YES")
            p_oRSRoutex.MoveNext
         Next
      End With
   End If
   Call loadSelected
End Sub

Private Sub cmbSource_Click()
   If Not pbLoaded Then Exit Sub
   LoadRoute cmbSource.ListIndex
   loadSelected
End Sub

Private Sub cmdAdd_Click()
   Call searchDetail(Trim(txtDetail(0)))
End Sub

Private Sub searchDetail(ByVal sReferNox As String)
   Dim lors As Recordset
   Dim lsSQLMC As String
   Dim lsSQLSP As String
   Dim lsSQLCP As String
   Dim lsSQLAs As String
   Dim lsSQLSU As String
   Dim lsSQLRg As String
   Dim lasSelect() As String
   Dim lsSQL As String
   Dim lsCriteria As String
   
   lsSQLMC = "SELECT" & _
                  "  a.sTransNox" & _
                  ", b.sBranchNm" & _
                  ", 'MC Transfer' xDescript" & _
                  ", 'MC' xSourcexx" & _
                  ", 'MCDl' xSourceCd" & _
                  ", a.dTransact" & _
               " FROM MC_Transfer_Master a" & _
                  ", Branch b" & _
               " WHERE a.sDestinat = b.sBranchCd" & _
                  " AND b.sBranchCd <> " & strParm(p_sBranchCd) & _
                  " AND a.sTransNox LIKE " & strParm("%" & sReferNox) & _
                  " AND (a.cTranStat = " & strParm(xeStateOpen) & _
                     " OR a.cTranStat = " & strParm(xeStateClosed) & ")" & _
                  " AND (a.cDeliverx IS NULL OR a.cDeliverx = " & strParm(xeNo) & ")"
   lsSQLSP = "SELECT" & _
                  "  a.sTransNox" & _
                  ", b.sBranchNm" & _
                  ", 'SP Transfer' xDescript" & _
                  ", 'SP' xSourcexx" & _
                  ", 'SPDl' xSourceCd" & _
                  ", a.dTransact" & _
               " FROM SP_Transfer_Master a" & _
                  ", Branch b" & _
               " WHERE a.sDestinat = b.sBranchCd" & _
                  " AND b.sBranchCd <> " & strParm(p_sBranchCd) & _
                  " AND a.sTransNox LIKE " & strParm("%" & sReferNox) & _
                  " AND (a.cTranStat = " & strParm(xeStateOpen) & _
                     " OR a.cTranStat = " & strParm(xeStateClosed) & ")" & _
                  " AND (a.cDeliverx IS NULL OR a.cDeliverx = " & strParm(xeNo) & ")"
   lsSQLSP = lsSQLSP & _
               " UNION SELECT" & _
                  "  a.sTransNox" & _
                  ", b.sBranchNm" & _
                  ", 'SP Waranty Transfer' xDescript" & _
                  ", 'SP' xSourcexx" & _
                  ", 'SPWT' xSourceCd" & _
                  ", a.dTransact" & _
               " FROM SP_Warranty_Transfer_Master a" & _
                  ", Branch b" & _
               " WHERE a.sDestinat = b.sBranchCd" & _
                  " AND b.sBranchCd <> " & strParm(p_sBranchCd) & _
                  " AND a.sTransNox LIKE " & strParm("%" & sReferNox) & _
                  " AND (a.cTranStat = " & strParm(xeStateOpen) & _
                     " OR a.cTranStat = " & strParm(xeStateClosed) & ")" & _
                  " AND (a.cDeliverx IS NULL OR a.cDeliverx = " & strParm(xeNo) & ")"
   lsSQLCP = "SELECT" & _
                  "  a.sTransNox" & _
                  ", b.sBranchNm" & _
                  ", 'CP Transfer' xDescript" & _
                  ", 'CP' xSourcexx" & _
                  ", 'CPDl' xSourceCd" & _
                  ", a.dTransact" & _
               " FROM CP_Transfer_Master a" & _
                  ", Branch b" & _
               " WHERE a.sDestinat = b.sBranchCd" & _
                  " AND b.sBranchCd <> " & strParm(p_sBranchCd) & _
                  " AND a.sTransNox LIKE " & strParm("%" & sReferNox) & _
                  " AND (a.cTranStat = " & strParm(xeStateOpen) & _
                     " OR a.cTranStat = " & strParm(xeStateClosed) & ")" & _
                  " AND (a.cDeliverx IS NULL OR a.cDeliverx = " & strParm(xeNo) & ")"
   lsSQLCP = lsSQLCP & _
               " UNION SELECT" & _
                  "  a.sTransNox" & _
                  ", b.sBranchNm" & _
                  ", 'CP Job Order' xDescript" & _
                  ", 'CP' xSourcexx" & _
                  ", 'CPJT' xSourceCd" & _
                  ", a.dTransact" & _
               " FROM CP_JobOrder_Transfer_Master a" & _
                  ", Branch b" & _
               " WHERE a.sDestinat = b.sBranchCd" & _
                  " AND b.sBranchCd <> " & strParm(p_sBranchCd) & _
                  " AND a.sTransNox LIKE " & strParm("%" & sReferNox) & _
                  " AND (a.cTranStat = " & strParm(xeStateOpen) & _
                     " OR a.cTranStat = " & strParm(xeStateClosed) & ")" & _
                  " AND (a.cDeliverx IS NULL OR a.cDeliverx = " & strParm(xeNo) & ")"
   lsSQLAs = "SELECT" & _
                  "  a.sTransNox" & _
                  ", b.sBranchNm" & _
                  ", 'Asset Transfer' xDescript" & _
                  ", 'As' xSourcexx" & _
                  ", 'AsDl' xSourceCd" & _
                  ", a.dTransact" & _
               " FROM Asset_Transfer_Master a" & _
                  ", Branch b" & _
               " WHERE a.sDestinat = b.sBranchCd" & _
                  " AND b.sBranchCd <> " & strParm(p_sBranchCd) & _
                  " AND a.sTransNox LIKE " & strParm("%" & sReferNox) & _
                  " AND (a.cTranStat = " & strParm(xeStateOpen) & _
                     " OR a.cTranStat = " & strParm(xeStateClosed) & ")" & _
                  " AND (a.cDeliverx IS NULL OR a.cDeliverx = " & strParm(xeNo) & ")"
   lsSQLSU = "SELECT" & _
                  "  a.sTransNox" & _
                  ", b.sBranchNm" & _
                  ", 'Supplies Transfer' xDescript" & _
                  ", 'SU' xSourcexx" & _
                  ", 'SUDl' xSourceCd" & _
                  ", a.dTransact" & _
               " FROM Supplies_Transfer_Master a" & _
                  ", Branch b" & _
               " WHERE a.sDestinat = b.sBranchCd" & _
                  " AND b.sBranchCd <> " & strParm(p_sBranchCd) & _
                  " AND a.sTransNox LIKE " & strParm("%" & sReferNox) & _
                  " AND (a.cTranStat = " & strParm(xeStateOpen) & _
                     " OR a.cTranStat = " & strParm(xeStateClosed) & ")" & _
                  " AND (a.cDeliverx IS NULL OR a.cDeliverx = " & strParm(xeNo) & ")"
   lsSQLRg = "SELECT" & _
                  "  a.sTransNox" & _
                  ", b.sBranchNm" & _
                  ", 'MC Registration Transfer' xDescript" & _
                  ", 'Rg' xSourcexx" & _
                  ", 'RgDl' xSourceCd" & _
                  ", a.dTransact" & _
               " FROM MC_Reg_Transfer_Master a" & _
                  ", Branch b" & _
               " WHERE a.sBranchCd = b.sBranchCd" & _
                  " AND a.sBranchCd <> " & strParm(p_sBranchCd) & _
                  " AND a.sTransNox LIKE " & strParm("%" & sReferNox) & _
                  " AND (a.cTranStat = " & strParm(xeStateOpen) & _
                     " OR a.cTranStat = " & strParm(xeStateClosed) & ")" & _
                  " AND (a.cDeliverx IS NULL OR a.cDeliverx = " & strParm(xeNo) & ")"
                  
   Select Case cmbSource.ListIndex
   Case 0
      lsSQL = lsSQLMC & _
               " UNION " & _
               lsSQLSP & _
               " UNION " & _
               lsSQLCP & _
               " UNION " & _
               lsSQLRg & _
               " UNION " & _
               lsSQLAs & _
               " UNION " & _
               lsSQLSU
      lsCriteria = "b.sBranchNm»xDescript»a.sTransNox»a.dTransact"
   Case 1
      lsSQL = lsSQLMC
      lsCriteria = "b.sBranchNm»'MC Transfer' xDescript»a.sTransNox»a.dTransact"
   Case 2
      lsSQL = lsSQLSP
      lsCriteria = "b.sBranchNm»'SP Transfer' xDescript»a.sTransNox»a.dTransact"
   Case 3
      lsSQL = lsSQLCP
      lsCriteria = "b.sBranchNm»'CP Transfer' xDescript»a.sTransNox»a.dTransact"
   Case 4
      lsSQL = lsSQLRg
      lsCriteria = "b.sBranchNm»'MC Registration Transfer' xDescript»a.sTransNox»a.dTransact"
   Case 5
      lsSQL = lsSQLAs
      lsCriteria = "b.sBranchNm»'Asset Transfer' xDescript»a.sTransNox»a.dTransact"
   Case 6
      lsSQL = lsSQLSU
      lsCriteria = "b.sBranchNm»'Supplies Transfer' xDescript»a.sTransNox»a.dTransact"
   End Select
'   MsgBox lsSQL
'   lsSQL = "SELECT a.sBranchNm,a.xDescript,a.sTransNox,a.dTransact FROM(" & lsSQL & ") a"
'   Debug.Print lsSQL
   Set lors = New Recordset
   lors.Open lsSQL, oApp.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
   If lors.EOF Then GoTo endProc
   lsSQL = KwikBrowse(oApp, lors, _
                           "sBranchNm»xDescript»sTransNox»dTransact", _
                           "Branch»Source»Refer No»Date", _
                           "@»@»@@@@-@@@@@@@@»@")

'   lsSQL = KwikSearch(oApp, lsSQL, _
'                           "sBranchNm»xDescript»sTransNox»dTransact", _
'                           "Branch»Source»Refer No»Date", _
'                           "@»@»@@@@-@@@@@@@@»@")
   If lsSQL = "" Then GoTo endProc
   lasSelect = Split(lsSQL, "»")
   With p_oRSRoutex
      .AddNew
      .Fields("sReferNox") = lasSelect(0)
      .Fields("sDestinat") = lasSelect(1)
      .Fields("sDescript") = lasSelect(2)
      .Fields("cIncluded") = xeYes
      .Fields("sSourcexx") = lasSelect(3)
      .Fields("sSourceCd") = lasSelect(4)
      .Fields("dTransact") = lasSelect(5)
      .Fields("sRemarksx") = ""
   End With
   Call loadSelected

endProc:
   txtDetail(0) = ""
   Exit Sub
errProc:
   MsgBox Err.Description
End Sub

Private Sub cmdButton_Click(Index As Integer)
   Dim lnCtr As Integer
   
   Select Case Index
   Case 0   'Cancel
      pbLoaded = False
      Unload Me
   Case 1   'Okey
      If oTrans.delBranchDetail(p_sBranchCd) Then
         With p_oRSRoutex
            .Filter = adFilterNone
            .MoveFirst
            lnCtr = 0
            Do Until .EOF
               If .Fields("cIncluded") = xeYes Then
                  Call oTrans.addDetail(p_sBranchCd)
                  oTrans.Detail(lnCtr, "sReferNox", p_sBranchCd) = .Fields("sReferNox")
                  oTrans.Detail(lnCtr, "sDescript", p_sBranchCd) = .Fields("sDescript")
                  oTrans.Detail(lnCtr, "sSourcexx", p_sBranchCd) = .Fields("sSourcexx")
                  oTrans.Detail(lnCtr, "sSourceCd", p_sBranchCd) = .Fields("sSourceCd")
                  oTrans.Detail(lnCtr, "sRemarksx", p_sBranchCd) = .Fields("sRemarksx")
                  oTrans.Detail(lnCtr, "cIncluded", p_sBranchCd) = .Fields("cIncluded")
                  oTrans.Detail(lnCtr, "sDestinat", p_sBranchCd) = .Fields("sDestinat")
                  lnCtr = lnCtr + 1
               End If
               .MoveNext
            Loop
         End With
         pbLoaded = False
         Unload Me
      End If
   End Select
End Sub

Property Set Delivery(ByVal oDelivery As clsDeliveryMonitoring)
   Set oTrans = oDelivery
End Property

Property Let EntryNo(ByVal nEntryNo As Integer)
   p_nEntryNo = nEntryNo
End Property

Private Sub Form_Activate()
   If pbLoaded = False Then
      oApp.MenuName = Me.Tag
      Me.ZOrder 0
      
      createRoute
      loadFields
      retreiveRecords
      LoadRoute cmbSource.ListIndex
      loadSelected
      pbLoaded = True
      
      MSFlexGrid1.Refresh
   End If
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String
   Dim loTxt As TextBox
   lsOldProc = "Form_Load"
   ''On Error GoTo errProc
   
   CenterChildForm mdiMain, Me

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualRight
   
   InitGrid
   cmbSource.ListIndex = 0
   txtDetail(0) = ""
   txtDetail(1) = ""
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
End Sub

Private Sub InitGrid()
   Dim lnCtr As Integer
   
   With MSFlexGrid1
      .Rows = 2
      .Cols = 4
      .Font = "MS Sans Serif"

      'column title
      .TextMatrix(0, 0) = "No"
      .TextMatrix(0, 1) = "Transfer"
      .TextMatrix(0, 2) = "Refer No"
      .TextMatrix(0, 3) = "INC"
      .Row = 0

      'column alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = 1
      Next

      'column width
      .ColWidth(0) = 450
      .ColWidth(1) = 1900
      .ColWidth(2) = 1400
      .ColWidth(3) = 400
      
      .Row = 1
      .Col = 1
      .ColSel = .Cols - 1
   End With
   
   With MSFlexGrid2
      .Rows = 2
      .Cols = 5
      .Font = "MS Sans Serif"

      'column title
      .TextMatrix(0, 0) = "No"
      .TextMatrix(0, 1) = "Transfer"
      .TextMatrix(0, 2) = "Refer No"
      .TextMatrix(0, 3) = "Remarks"
      .TextMatrix(0, 4) = "Destination"
      .Row = 0

      'column alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = 1
      Next

      'column width
      .ColWidth(0) = 450
      .ColWidth(1) = 1900
      .ColWidth(2) = 1400
      .ColWidth(3) = 3000
      .ColWidth(4) = 3000
      
      .Row = 1
      .Col = 1
      .ColSel = .Cols - 1
   End With
End Sub

Private Sub loadFields()
   txtRoute = oTrans.Route(p_nEntryNo - 1, "sBranchNm")
   p_sBranchCd = oTrans.Route(p_nEntryNo - 1, "sBranchCd")
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

Private Sub MSFlexGrid1_DblClick()
   With MSFlexGrid1
      If .Row = 0 Then Exit Sub
      p_nRow = .Row
      p_oRSRoutex.Move MSFlexGrid1.Row - 1, adBookmarkFirst
      If .TextMatrix(.Row, 3) = "NO" Then
         .TextMatrix(.Row, 3) = "YES"
         p_oRSRoutex("cIncluded") = xeYes
      Else
         .TextMatrix(.Row, 3) = "NO"
         p_oRSRoutex("cIncluded") = xeNo
      End If
      Call loadSelected
   End With
End Sub

Private Sub MSFlexGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
   Call MSFlexGrid1_DblClick
End Sub

Private Sub MSFlexGrid1_RowColChange()
   If Not pbLoaded Then Exit Sub
   If p_oRSRoutex.RecordCount = 0 Then Exit Sub
   p_oRSRoutex.Move MSFlexGrid1.Row - 1, adBookmarkFirst
   p_nRow = MSFlexGrid1.Row
   
'   Dim lnCtr As Integer
'
'   With MSFlexGrid2
'      For lnCtr = 1 To .Rows - 1
'         If .TextMatrix(lnCtr, 1) = MSFlexGrid1.TextMatrix(MSFlexGrid2.Row, 1) And _
'            .TextMatrix(lnCtr, 2) = MSFlexGrid1.TextMatrix(MSFlexGrid2.Row, 2) Then
'
'            .TopRow = lnCtr
'            .Row = lnCtr
'            .Col = 1
'            .ColSel = .Cols - 1
'         End If
'      Next
'   End With
End Sub

Private Sub MSFlexGrid2_RowColChange()
   Dim lnCtr As Integer
   
   If Not pbLoaded Then Exit Sub
   With MSFlexGrid1
      For lnCtr = 1 To .Rows - 1
         If .TextMatrix(lnCtr, 1) = MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 1) And _
            .TextMatrix(lnCtr, 2) = MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 2) Then

            p_oRSRoutex.Move lnCtr - 1, adBookmarkFirst
            txtDetail(1) = p_oRSRoutex("sRemarksx")
         End If
      Next
   End With
End Sub

Private Sub txtDetail_GotFocus(Index As Integer)
   Call HighlightOn(Me.txtDetail(Index))
End Sub

Private Sub txtDetail_LostFocus(Index As Integer)
   Call HighlightOff(Me.txtDetail(Index))
End Sub

Private Sub txtDetail_Validate(Index As Integer, Cancel As Boolean)
   With txtDetail(Index)
      Select Case Index
      Case 1
         p_oRSRoutex("sRemarksx") = .Text
      End Select
   End With
End Sub

Private Sub txtRoute_GotFocus()
   Call HighlightOn(Me.txtRoute)
End Sub

Private Sub txtRoute_LostFocus()
   Call HighlightOff(Me.txtRoute)
End Sub

Private Sub retreiveRecords()
   Dim lors As Recordset
   Dim lsSQLMC As String
   Dim lsSQLSP As String
   Dim lsSQLCP As String
   Dim lsSQLAs As String
   Dim lsSQLSU As String
   Dim lsSQLRg As String
   Dim lsSQL As String
   Dim lnCtr As Integer
   
   lsSQLMC = "SELECT" & _
                  "  a.sTransNox" & _
                  ", b.sBranchNm" & _
                  ", 'MC Transfer' xDescript" & _
                  ", 'MC' xSourcexx" & _
                  ", 'MCDl' xSourceCd" & _
                  ", a.dTransact" & _
               " FROM MC_Transfer_Master a" & _
                  ", Branch b" & _
               " WHERE a.sDestinat = b.sBranchCd" & _
                  " AND b.sBranchCd = " & strParm(p_sBranchCd) & _
                  " AND (a.cTranStat = " & strParm(xeStateClosed) & _
                     " OR a.cTranStat = " & strParm(xeStatePosted) & ")" & _
                  " AND (a.cDeliverx IS NULL OR a.cDeliverx = " & strParm(xeNo) & ")"
   lsSQLSP = "SELECT" & _
                  "  a.sTransNox" & _
                  ", b.sBranchNm" & _
                  ", 'SP Transfer' xDescript" & _
                  ", 'SP' xSourcexx" & _
                  ", 'SPDl' xSourceCd" & _
                  ", a.dTransact" & _
               " FROM SP_Transfer_Master a" & _
                  ", Branch b" & _
               " WHERE a.sDestinat = b.sBranchCd" & _
                  " AND b.sBranchCd = " & strParm(p_sBranchCd) & _
                  " AND (a.cTranStat = " & strParm(xeStateClosed) & _
                     " OR a.cTranStat = " & strParm(xeStatePosted) & ")" & _
                  " AND (a.cDeliverx IS NULL OR a.cDeliverx = " & strParm(xeNo) & ")"
   lsSQLSP = lsSQLSP & _
               " UNION SELECT" & _
                  "  a.sTransNox" & _
                  ", b.sBranchNm" & _
                  ", 'SP Waranty Transfer' xDescript" & _
                  ", 'SP' xSourcexx" & _
                  ", 'SPWT' xSourceCd" & _
                  ", a.dTransact" & _
               " FROM SP_Warranty_Transfer_Master a" & _
                  ", Branch b" & _
               " WHERE a.sDestinat = b.sBranchCd" & _
                  " AND b.sBranchCd = " & strParm(p_sBranchCd) & _
                  " AND (a.cTranStat = " & strParm(xeStateClosed) & _
                     " OR a.cTranStat = " & strParm(xeStatePosted) & ")" & _
                  " AND (a.cDeliverx IS NULL OR a.cDeliverx = " & strParm(xeNo) & ")"
   lsSQLCP = "SELECT" & _
                  "  a.sTransNox" & _
                  ", b.sBranchNm" & _
                  ", 'CP Transfer' xDescript" & _
                  ", 'CP' xSourcexx" & _
                  ", 'CPDl' xSourceCd" & _
                  ", a.dTransact" & _
               " FROM CP_Transfer_Master a" & _
                  ", Branch b" & _
               " WHERE a.sDestinat = b.sBranchCd" & _
                  " AND b.sBranchCd = " & strParm(p_sBranchCd) & _
                  " AND (a.cTranStat = " & strParm(xeStateClosed) & _
                     " OR a.cTranStat = " & strParm(xeStatePosted) & ")" & _
                  " AND (a.cDeliverx IS NULL OR a.cDeliverx = " & strParm(xeNo) & ")"
   lsSQLCP = lsSQLCP & _
               " UNION SELECT" & _
                  "  a.sTransNox" & _
                  ", b.sBranchNm" & _
                  ", 'CP Job Order' xDescript" & _
                  ", 'CP' xSourcexx" & _
                  ", 'CPJT' xSourceCd" & _
                  ", a.dTransact" & _
               " FROM CP_JobOrder_Transfer_Master a" & _
                  ", Branch b" & _
               " WHERE a.sDestinat = b.sBranchCd" & _
                  " AND b.sBranchCd = " & strParm(p_sBranchCd) & _
                  " AND (a.cTranStat = " & strParm(xeStateClosed) & _
                     " OR a.cTranStat = " & strParm(xeStatePosted) & ")" & _
                  " AND (a.cDeliverx IS NULL OR a.cDeliverx = " & strParm(xeNo) & ")"
   lsSQLAs = "SELECT" & _
                  "  a.sTransNox" & _
                  ", b.sBranchNm" & _
                  ", 'Asset Transfer' xDescript" & _
                  ", 'As' xSourcexx" & _
                  ", 'AsDl' xSourceCd" & _
                  ", a.dTransact" & _
               " FROM Asset_Transfer_Master a" & _
                  ", Branch b" & _
               " WHERE a.sDestinat = b.sBranchCd" & _
                  " AND b.sBranchCd = " & strParm(p_sBranchCd) & _
                  " AND (a.cTranStat = " & strParm(xeStateClosed) & _
                     " OR a.cTranStat = " & strParm(xeStatePosted) & ")" & _
                  " AND (a.cDeliverx IS NULL OR a.cDeliverx = " & strParm(xeNo) & ")"
   lsSQLSU = "SELECT" & _
                  "  a.sTransNox" & _
                  ", b.sBranchNm" & _
                  ", 'Supplies Transfer' xDescript" & _
                  ", 'SU' xSourcexx" & _
                  ", 'SUDl' xSourceCd" & _
                  ", a.dTransact" & _
               " FROM Supplies_Transfer_Master a" & _
                  ", Branch b" & _
               " WHERE a.sDestinat = b.sBranchCd" & _
                  " AND b.sBranchCd = " & strParm(p_sBranchCd) & _
                  " AND (a.cTranStat = " & strParm(xeStateClosed) & _
                     " OR a.cTranStat = " & strParm(xeStatePosted) & ")" & _
                  " AND (a.cDeliverx IS NULL OR a.cDeliverx = " & strParm(xeNo) & ")"
   lsSQLRg = "SELECT" & _
                  "  a.sTransNox" & _
                  ", b.sBranchNm" & _
                  ", 'MC Registration Transfer' xDescript" & _
                  ", 'Rg' xSourcexx" & _
                  ", 'RgDl' xSourceCd" & _
                  ", a.dTransact" & _
               " FROM MC_Reg_Transfer_Master a" & _
                  ", Branch b" & _
               " WHERE a.sBranchCd = b.sBranchCd" & _
                  " AND a.sBranchCd = " & strParm(p_sBranchCd) & _
                  " AND (a.cTranStat = " & strParm(xeStateClosed) & _
                     " OR a.cTranStat = " & strParm(xeStatePosted) & ")" & _
                  " AND (a.cDeliverx IS NULL OR a.cDeliverx = " & strParm(xeNo) & ")"
   lsSQL = lsSQLMC & _
            " UNION " & _
            lsSQLSP & _
            " UNION " & _
            lsSQLCP & _
            " UNION " & _
            lsSQLRg & _
            " UNION " & _
            lsSQLAs & _
            " UNION " & _
            lsSQLSU
   
   Set lors = New Recordset
   lors.Open lsSQL, oApp.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
   
   Do Until lors.EOF
      p_oRSRoutex.AddNew
      p_oRSRoutex("sReferNox") = lors("sTransNox")
      p_oRSRoutex("sSourcexx") = lors("xSourcexx")
      p_oRSRoutex("sSourceCd") = lors("xSourceCd")
      p_oRSRoutex("sDescript") = lors("xDescript")
      p_oRSRoutex("dTransact") = lors("dTransact")
      p_oRSRoutex("sDestinat") = lors("sBranchNm")
      p_oRSRoutex("cIncluded") = IIf(oTrans.SearchSerial(p_sBranchCd, lors("xSourceCd"), lors("sTransNox")), 1, 0)
      lors.MoveNext
   Loop
   
   For lnCtr = 0 To oTrans.ItemCount(p_sBranchCd) - 1
      If txtRoute <> oTrans.Detail(lnCtr, "sDestinat", p_sBranchCd) Then
         p_oRSRoutex.AddNew
         p_oRSRoutex("sReferNox") = oTrans.Detail(lnCtr, "sReferNox", p_sBranchCd)
         p_oRSRoutex("sSourcexx") = oTrans.Detail(lnCtr, "sSourcexx", p_sBranchCd)
         p_oRSRoutex("sSourceCd") = oTrans.Detail(lnCtr, "sSourceCd", p_sBranchCd)
         p_oRSRoutex("sDescript") = oTrans.Detail(lnCtr, "sDescript", p_sBranchCd)
         p_oRSRoutex("dTransact") = oTrans.Detail(lnCtr, "dTransact", p_sBranchCd)
         p_oRSRoutex("sDestinat") = oTrans.Detail(lnCtr, "sDestinat", p_sBranchCd)
         p_oRSRoutex("cIncluded") = xeYes
      End If
   Next
End Sub

Private Sub LoadRoute(ByVal Index As Integer)
   Dim lnCtr As Integer
   
   With p_oRSRoutex
      .Filter = adFilterNone
      .Sort = "dTransact,sSourceCd,sReferNox"
   
      Select Case Index
      Case 1
         .Filter = " sSourcexx = 'MC'"
      Case 2
         .Filter = " sSourcexx = 'SP'"
      Case 3
         .Filter = " sSourcexx = 'CP'"
      Case 4
         .Filter = " sSourcexx = 'Rg'"
      Case 5
         .Filter = " sSourcexx = 'As'"
      Case 6
         .Filter = " sSourcexx = 'SU'"
      End Select
      
      With MSFlexGrid1
         If Not p_oRSRoutex.EOF Then
            p_oRSRoutex.MoveFirst
            Do Until p_oRSRoutex.EOF
               If txtRoute = p_oRSRoutex("sDestinat") Then
                  .Rows = lnCtr + 1 + 1
                  .TextMatrix(lnCtr + 1, 0) = lnCtr + 1
                  .TextMatrix(lnCtr + 1, 1) = p_oRSRoutex("sDescript")
                  .TextMatrix(lnCtr + 1, 2) = p_oRSRoutex("sReferNox")
                  .TextMatrix(lnCtr + 1, 3) = IIf(p_oRSRoutex("cIncluded") = 0, "NO", "YES")
                  lnCtr = lnCtr + 1
               End If
               p_oRSRoutex.MoveNext
            Loop
         Else
            .Rows = 2
            .TextMatrix(lnCtr + 1, 0) = 1
            .TextMatrix(lnCtr + 1, 1) = ""
            .TextMatrix(lnCtr + 1, 2) = ""
            .TextMatrix(lnCtr + 1, 3) = ""
         End If
         
         .Row = 1
         .Col = 1
         .ColSel = .Cols - 1
      End With
   End With
End Sub

Private Sub createRoute()
   Set p_oRSRoutex = New Recordset
   With p_oRSRoutex
      .Fields.Append "sReferNox", adChar, 12
      .Fields.Append "sDescript", adVarChar, 50
      .Fields.Append "cIncluded", adChar, 1
      .Fields.Append "sSourcexx", adVarChar, 2
      .Fields.Append "sSourceCd", adChar, 4
      .Fields.Append "dTransact", adDate
      .Fields.Append "sRemarksx", adVarChar, 1256
      .Fields.Append "sDestinat", adVarChar, 50
      .Open
   End With
End Sub

Private Sub loadSelected()
   Dim lnCtr As Integer

   With MSFlexGrid2
      p_oRSRoutex.Filter = "cIncluded = " & strParm(xeYes)
      
      If p_oRSRoutex.EOF Then
         .Rows = 2
         
         .TextMatrix(1, 0) = "1"
         .TextMatrix(1, 1) = ""
         .TextMatrix(1, 2) = ""
         .TextMatrix(1, 3) = ""
         .TextMatrix(1, 4) = ""
      Else
         .Rows = p_oRSRoutex.RecordCount + 1
         
         p_oRSRoutex.MoveFirst
         For lnCtr = 0 To p_oRSRoutex.RecordCount - 1
            .TextMatrix(lnCtr + 1, 0) = lnCtr + 1
            .TextMatrix(lnCtr + 1, 1) = p_oRSRoutex("sDescript")
            .TextMatrix(lnCtr + 1, 2) = p_oRSRoutex("sReferNox")
            .TextMatrix(lnCtr + 1, 3) = p_oRSRoutex("sRemarksx")
            .TextMatrix(lnCtr + 1, 4) = p_oRSRoutex("sDestinat")
            p_oRSRoutex.MoveNext
         Next
      End If
      
      p_oRSRoutex.Filter = adFilterNone
      p_oRSRoutex.Sort = "dTransact,sSourceCd,sReferNox"
   
      Select Case cmbSource.ListIndex
      Case 1
         p_oRSRoutex.Filter = " sSourcexx = 'MC'"
      Case 2
         p_oRSRoutex.Filter = " sSourcexx = 'SP'"
      Case 3
         p_oRSRoutex.Filter = " sSourcexx = 'CP'"
      Case 4
         p_oRSRoutex.Filter = " sSourcexx = 'Rg'"
      Case 5
         p_oRSRoutex.Filter = " sSourcexx = 'As'"
      Case 6
         p_oRSRoutex.Filter = " sSourcexx = 'SU'"
      End Select
      
      .Row = .Rows - 1
      .Col = 1
      .ColSel = .Cols - 1
   End With
End Sub
