VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmCPClusterDeliveryDetail 
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
         TabIndex        =   10
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
         ItemData        =   "frmCPClusterDeliveryDetail.frx":0000
         Left            =   1200
         List            =   "frmCPClusterDeliveryDetail.frx":001C
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
      TabIndex        =   12
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
      Picture         =   "frmCPClusterDeliveryDetail.frx":0070
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   10530
      TabIndex        =   11
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
      Picture         =   "frmCPClusterDeliveryDetail.frx":07EA
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle No"
      Height          =   195
      Index           =   9
      Left            =   375
      TabIndex        =   13
      Top             =   900
      Width           =   780
   End
End
Attribute VB_Name = "frmCPClusterDeliveryDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pxeMODULENAME = "frmCPClusterDeliveryDetail"

Private oSkin As clsFormSkin
Private pbLoaded As Boolean
Private p_nRow As Integer

Private p_sClustrID As String
Private p_sBranchCd As String
Private p_sBranchNm As String
Private p_sCancelld As String
Private p_oRSOthers As Recordset
Private p_oRSDetail As Recordset
Private p_oRSRemove As Recordset

Private oTrans As clsCPClusterDelivery

Property Let Delivery(ByVal oValue As clsCPClusterDelivery)
   Set oTrans = oValue
End Property

Property Let Cluster(ByVal Value As String)
   p_sClustrID = Value
End Property

Property Let Branch(ByVal Value As String)
   p_sBranchNm = Value
End Property

Property Get Transfer() As Recordset
   Set Transfer = p_oRSDetail
End Property

Property Get UnTransfer() As Recordset
   Set UnTransfer = p_oRSRemove
End Property

Property Get Cancelled() As String
   Cancelled = p_sCancelld
End Property

Private Sub Check2_Click()
   Dim lnCtr As Integer
   
   If Check2.Value = Checked Then
      If p_oRSOthers.EOF Then Exit Sub
      With MSFlexGrid1
         .Rows = p_oRSOthers.RecordCount + 1
         
         p_oRSOthers.MoveFirst
         For lnCtr = 0 To p_oRSOthers.RecordCount - 1
            p_oRSOthers("cIncluded") = xeYes
            .TextMatrix(lnCtr + 1, 3) = IIf(p_oRSOthers("cIncluded") = 0, "NO", "YES")
            p_oRSOthers.MoveNext
         Next
      End With
   Else
      With MSFlexGrid1
         .Rows = p_oRSOthers.RecordCount + 1
         
         p_oRSOthers.MoveFirst
         For lnCtr = 0 To p_oRSOthers.RecordCount - 1
            p_oRSOthers("cIncluded") = xeNo
            .TextMatrix(lnCtr + 1, 3) = IIf(p_oRSOthers("cIncluded") = 0, "NO", "YES")
            p_oRSOthers.MoveNext
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

Private Sub cmdButton_Click(Index As Integer)
   Dim lnCtr As Integer
   Dim lbAdd As Boolean
   
   Select Case Index
   Case 0   'Cancel
      p_sCancelld = True
      pbLoaded = False
      Unload Me
   Case 1   'Okey
      With p_oRSOthers
         .Filter = "cIncluded = " & strParm(xeYes)
         If Not .EOF Then .MoveFirst
         Do Until .EOF
            lbAdd = True
            For lnCtr = 0 To oTrans.ItemCount - 1
               If .Fields("sReferNox") = oTrans.Detail(lnCtr, "sReferNox") And _
                  .Fields("sSourceCd") = oTrans.Detail(lnCtr, "sSourceCd") Then
                  
                  lbAdd = False
                  Exit For
               End If
            Next
            
            If lbAdd Then
               p_oRSDetail.AddNew
               p_oRSDetail("sReferNox") = .Fields("sReferNox")
               p_oRSDetail("sSourcexx") = .Fields("sSourcexx")
               p_oRSDetail("sSourceCd") = .Fields("sSourceCd")
               p_oRSDetail("sDescript") = .Fields("sDescript")
               p_oRSDetail("dTransact") = .Fields("dTransact")
               p_oRSDetail("sDestinat") = .Fields("sDestinat")
               p_oRSDetail("sRemarksx") = .Fields("sRemarksx")
               p_oRSDetail("sBranchCd") = .Fields("sBranchCd")
               p_oRSDetail("cIncluded") = .Fields("cIncluded")
            End If
            .MoveNext
         Loop
         
         .Filter = ""
         .Filter = "cIncluded = " & strParm(xeNo)
         If Not .EOF Then .MoveFirst

         Do Until .EOF
            lbAdd = False
            For lnCtr = 0 To oTrans.ItemCount - 1
               If .Fields("sReferNox") = oTrans.Detail(lnCtr, "sReferNox") And _
                  .Fields("sSourceCd") = oTrans.Detail(lnCtr, "sSourceCd") Then
                  
                  lbAdd = True
                  Exit For
               End If
            Next
            
            If lbAdd Then
               p_oRSRemove.AddNew
               p_oRSRemove("sReferNox") = .Fields("sReferNox")
               p_oRSRemove("sSourcexx") = .Fields("sSourcexx")
               p_oRSRemove("sSourceCd") = .Fields("sSourceCd")
               p_oRSRemove("sDescript") = .Fields("sDescript")
               p_oRSRemove("dTransact") = .Fields("dTransact")
               p_oRSRemove("sDestinat") = .Fields("sDestinat")
               p_oRSRemove("sRemarksx") = .Fields("sRemarksx")
               p_oRSRemove("sBranchCd") = .Fields("sBranchCd")
               p_oRSRemove("cIncluded") = .Fields("cIncluded")
            End If
            .MoveNext
         Loop
      End With
   
      p_sCancelld = False
      pbLoaded = False
      Me.Hide
   End Select
End Sub

Private Sub Form_Activate()
   If pbLoaded = False Then
      oApp.MenuName = Me.Tag
      Me.ZOrder 0
   
      pbLoaded = True
      
      createRoute
      
      If p_sBranchNm <> "" Then
         If getBranch(p_sBranchNm, False) Then
            Call retreiveRecords
            Call LoadRoute(cmbSource.ListIndex)
            Call loadSelected
         End If
      End If
      
      MSFlexGrid1.Refresh
   End If
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String
   Dim loTxt As TextBox
   lsOldProc = "Form_Load"
   'On Error GoTo errProc
   
   CenterChildForm mdiMain, Me

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualRight
   
   InitGrid
   ClearFields
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

Private Sub ClearFields()
   txtRoute.Text = ""
   txtRoute.Tag = ""
   Check2.Value = 0
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
      .ColWidth(1) = 2200
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
      p_oRSOthers.Move MSFlexGrid1.Row - 1, adBookmarkFirst
      If .TextMatrix(.Row, 3) = "NO" Then
         .TextMatrix(.Row, 3) = "YES"
         p_oRSOthers("cIncluded") = xeYes
      Else
         .TextMatrix(.Row, 3) = "NO"
         p_oRSOthers("cIncluded") = xeNo
      End If
      Call loadSelected
   End With
End Sub

Private Sub MSFlexGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
   Call MSFlexGrid1_DblClick
End Sub

Private Sub MSFlexGrid1_RowColChange()
   If Not pbLoaded Then Exit Sub
   If p_oRSOthers.RecordCount = 0 Then Exit Sub
   p_oRSOthers.Move MSFlexGrid1.Row - 1, adBookmarkFirst
   p_nRow = MSFlexGrid1.Row
End Sub

Private Sub MSFlexGrid2_RowColChange()
   Dim lnCtr As Integer
   
   If Not pbLoaded Then Exit Sub
   With MSFlexGrid1
      For lnCtr = 1 To .Rows - 1
         If .TextMatrix(lnCtr, 1) = MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 1) And _
            .TextMatrix(lnCtr, 2) = MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 2) Then

            p_oRSOthers.Move lnCtr - 1, adBookmarkFirst
            txtDetail(0) = p_oRSOthers("sReferNox")
            txtDetail(1) = p_oRSOthers("sRemarksx")
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

Private Sub txtRoute_GotFocus()
   Call HighlightOn(Me.txtRoute)
End Sub

Private Sub txtRoute_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyReturn, vbKeyF3
      If getBranch(IIf(txtRoute = "", "%", txtRoute), True) Then
         Call retreiveRecords
         Call LoadRoute(cmbSource.ListIndex)
         Call loadSelected
      End If
   End Select
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
   Dim lsSQLDc As String 'document
   Dim lsSQLCk As String 'check
   Dim lsSQL As String
   Dim lnCtr As Integer
   
   lsSQLMC = "SELECT" & _
                  "  a.sTransNox" & _
                  ", b.sBranchNm" & _
                  ", 'MC Transfer' xDescript" & _
                  ", 'MC' xSourcexx" & _
                  ", 'MCDv' xSourceCd" & _
                  ", a.dTransact" & _
                  ", a.sRemarksx" & _
                  ", a.sDestinat" & _
               " FROM MC_Transfer_Master a" & _
                  ", Branch b" & _
               " WHERE a.sDestinat = b.sBranchCd" & _
                  " AND b.sBranchCd = " & strParm(p_sBranchCd) & _
                  " AND (a.cTranStat = " & strParm(xeStateClosed) & _
                     " OR a.cTranStat = " & strParm(xeStatePosted) & ")" & _
                  " AND (a.cDeliverx IS NULL OR a.cDeliverx = " & strParm(xeNo) & ")" & _
                  " AND a.sTransNox NOT IN(SELECT sReferNox FROM Cluster_Delivery_Detail)"
                  
   lsSQLSP = "SELECT" & _
                  "  a.sTransNox" & _
                  ", b.sBranchNm" & _
                  ", 'SP Transfer' xDescript" & _
                  ", 'SP' xSourcexx" & _
                  ", 'SPDv' xSourceCd" & _
                  ", a.dTransact" & _
                  ", a.sRemarksx" & _
                  ", a.sDestinat" & _
               " FROM SP_Transfer_Master a" & _
                  ", Branch b" & _
               " WHERE a.sDestinat = b.sBranchCd" & _
                  " AND b.sBranchCd = " & strParm(p_sBranchCd) & _
                  " AND (a.cTranStat = " & strParm(xeStateClosed) & _
                     " OR a.cTranStat = " & strParm(xeStatePosted) & ")" & _
                  " AND (a.cDeliverx IS NULL OR a.cDeliverx = " & strParm(xeNo) & ")" & _
                  " AND a.sTransNox NOT IN(SELECT sReferNox FROM Cluster_Delivery_Detail)"
                               
   lsSQLSP = lsSQLSP & _
               " UNION SELECT" & _
                  "  a.sTransNox" & _
                  ", b.sBranchNm" & _
                  ", 'SP Waranty Transfer' xDescript" & _
                  ", 'SP' xSourcexx" & _
                  ", 'SPWT' xSourceCd" & _
                  ", a.dTransact" & _
                  ", a.sRemarksx" & _
                  ", a.sDestinat" & _
               " FROM SP_Warranty_Transfer_Master a" & _
                  ", Branch b" & _
               " WHERE a.sDestinat = b.sBranchCd" & _
                  " AND b.sBranchCd = " & strParm(p_sBranchCd) & _
                  " AND (a.cTranStat = " & strParm(xeStateClosed) & _
                     " OR a.cTranStat = " & strParm(xeStatePosted) & ")" & _
                  " AND (a.cDeliverx IS NULL OR a.cDeliverx = " & strParm(xeNo) & ")" & _
                  " AND a.sTransNox NOT IN(SELECT sReferNox FROM Cluster_Delivery_Detail)"
                  
   lsSQLCP = "SELECT" & _
                  "  a.sTransNox" & _
                  ", b.sBranchNm" & _
                  ", 'CP Transfer' xDescript" & _
                  ", 'CP' xSourcexx" & _
                  ", 'CPDv' xSourceCd" & _
                  ", a.dTransact" & _
                  ", a.sRemarksx" & _
                  ", a.sDestinat" & _
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
                  ", a.sRemarksx" & _
                  ", a.sDestinat" & _
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
                  ", 'AsDv' xSourceCd" & _
                  ", a.dTransact" & _
                  ", a.sRemarksx" & _
                  ", a.sDestinat" & _
               " FROM Asset_Transfer_Master a" & _
                  ", Branch b" & _
               " WHERE a.sDestinat = b.sBranchCd" & _
                  " AND b.sBranchCd = " & strParm(p_sBranchCd) & _
                  " AND (a.cTranStat = " & strParm(xeStateClosed) & _
                     " OR a.cTranStat = " & strParm(xeStatePosted) & ")" & _
                  " AND (a.cDeliverx IS NULL OR a.cDeliverx = " & strParm(xeNo) & ")" & _
                  " AND a.sTransNox NOT IN(SELECT sReferNox FROM Cluster_Delivery_Detail)"
                  
   lsSQLSU = "SELECT" & _
                  "  a.sTransNox" & _
                  ", b.sBranchNm" & _
                  ", 'Supplies Transfer' xDescript" & _
                  ", 'SU' xSourcexx" & _
                  ", 'SUDv' xSourceCd" & _
                  ", a.dTransact" & _
                  ", a.sRemarksx" & _
                  ", a.sDestinat" & _
               " FROM Supplies_Transfer_Master a" & _
                  ", Branch b" & _
               " WHERE a.sDestinat = b.sBranchCd" & _
                  " AND b.sBranchCd = " & strParm(p_sBranchCd) & _
                  " AND (a.cTranStat = " & strParm(xeStateClosed) & _
                     " OR a.cTranStat = " & strParm(xeStatePosted) & ")" & _
                  " AND (a.cDeliverx IS NULL OR a.cDeliverx = " & strParm(xeNo) & ")" & _
                  " AND a.sTransNox NOT IN(SELECT sReferNox FROM Cluster_Delivery_Detail)"
                  
   lsSQLRg = "SELECT" & _
                  "  a.sTransNox" & _
                  ", b.sBranchNm" & _
                  ", 'MC Registration Transfer' xDescript" & _
                  ", 'Rg' xSourcexx" & _
                  ", 'RgDv' xSourceCd" & _
                  ", a.dTransact" & _
                  ", '' sRemarksx" & _
                  ", a.sBranchCd sDestinat" & _
               " FROM MC_Reg_Transfer_Master a" & _
                  ", Branch b" & _
               " WHERE a.sBranchCd = b.sBranchCd" & _
                  " AND a.sBranchCd = " & strParm(p_sBranchCd) & _
                  " AND (a.cTranStat = " & strParm(xeStateClosed) & _
                     " OR a.cTranStat = " & strParm(xeStatePosted) & ")" & _
                  " AND (a.cDeliverx IS NULL OR a.cDeliverx = " & strParm(xeNo) & ")" & _
                  " AND a.sTransNox NOT IN(SELECT sReferNox FROM Cluster_Delivery_Detail)"

   lsSQLDc = "SELECT" & _
                  "  a.sTransNox" & _
                  ", b.sBranchNm" & _
                  ", 'Doc Transfer' xDescript" & _
                  ", 'DC' xSourcexx" & _
                  ", 'DCDv' xSourceCd" & _
                  ", a.dTransact" & _
                  ", a.sRemarksx" & _
                  ", a.sDestinat" & _
               " FROM General_Transfer_Master a" & _
                  ", Branch b" & _
               " WHERE a.sDestinat = b.sBranchCd" & _
                  " AND b.sBranchCd = " & strParm(p_sBranchCd) & _
                  " AND (a.cTranStat = " & strParm(xeStateClosed) & _
                     " OR a.cTranStat = " & strParm(xeStatePosted) & ")" & _
                  " AND (a.cDeliverx IS NULL OR a.cDeliverx = " & strParm(xeNo) & ")" & _
                  " AND a.sTransNox NOT IN(SELECT sReferNox FROM Cluster_Delivery_Detail)"
                  

   lsSQLCk = "SELECT" & _
                  "  a.sTransNox" & _
                  ", b.sBranchNm" & _
                  ", 'Check Transfer' xDescript" & _
                  ", 'CK' xSourcexx" & _
                  ", 'CKDv' xSourceCd" & _
                  ", a.dTransact" & _
                  ", a.sRemarksx" & _
                  ", a.sDestinat" & _
               " FROM Check_Transfer_Master a" & _
                  ", Branch b" & _
               " WHERE a.sDestinat = b.sBranchCd" & _
                  " AND b.sBranchCd = " & strParm(p_sBranchCd) & _
                  " AND (a.cTranStat = " & strParm(xeStateClosed) & _
                     " OR a.cTranStat = " & strParm(xeStatePosted) & ")" & _
                  " AND (a.cDeliverx IS NULL OR a.cDeliverx = " & strParm(xeNo) & ")" & _
                  " AND a.sTransNox NOT IN(SELECT sReferNox FROM Cluster_Delivery_Detail)"

   lsSQL = lsSQLSP & _
            " UNION " & _
            lsSQLCP & _
            " UNION " & _
            lsSQLRg & _
            " UNION " & _
            lsSQLAs & _
            " UNION " & _
            lsSQLSU & _
            " UNION " & _
            lsSQLDc & _
            " UNION " & _
            lsSQLCk & _
            " ORDER BY sBranchNm, sTransNox"
   Debug.Print lsSQL
   Set lors = New Recordset
   lors.Open lsSQL, oApp.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
   
   Do Until lors.EOF
      p_oRSOthers.AddNew
      p_oRSOthers("sReferNox") = lors("sTransNox")
      p_oRSOthers("sSourcexx") = lors("xSourcexx")
      p_oRSOthers("sSourceCd") = lors("xSourceCd")
      p_oRSOthers("sDescript") = lors("xDescript")
      p_oRSOthers("dTransact") = lors("dTransact")
      p_oRSOthers("sDestinat") = lors("sBranchNm")
      p_oRSOthers("sRemarksx") = lors("sRemarksx")
      p_oRSOthers("sBranchCd") = lors("sDestinat")
      p_oRSOthers("cIncluded") = xeNo
      lors.MoveNext
   Loop
   For lnCtr = 0 To oTrans.ItemCount - 1
      With p_oRSOthers
         .Filter = "sReferNox = " & strParm(oTrans.Detail(lnCtr, "sReferNox")) & _
                     " AND sSourceCd = " & strParm(oTrans.Detail(lnCtr, "sSourceCd"))
      
         If Not .EOF Then
            p_oRSOthers("cIncluded") = xeYes
         End If
         
         .Filter = ""
      End With
   Next
   
'   For lnCtr = 0 To oTrans.ItemCount - 1
'      If txtRoute <> oTrans.Detail(lnCtr, "sBranchNm") Then
'         If oTrans.Detail(lnCtr, "sBranchNm") <> "" And oTrans.Detail(lnCtr, "sSourceCd") <> "MCDv" Then
'            p_oRSOthers.AddNew
'            p_oRSOthers("sReferNox") = oTrans.Detail(lnCtr, "sReferNox")
'            p_oRSOthers("sSourcexx") = Left(oTrans.Detail(lnCtr, "sSourceCd"), 2)
'            p_oRSOthers("sSourceCd") = oTrans.Detail(lnCtr, "sSourceCd")
'            p_oRSOthers("sDestinat") = oTrans.Detail(lnCtr, "sBranchNm")
'            p_oRSOthers("sDescript") = getSource(oTrans.Detail(lnCtr, "sSourceCd"))
'            p_oRSOthers("cIncluded") = xeYes
'         End If
'      End If
'   Next
End Sub

Private Sub LoadRoute(ByVal Index As Integer)
   Dim lnCtr As Integer
   
   With p_oRSOthers
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
         If Not p_oRSOthers.EOF Then
            p_oRSOthers.MoveFirst
            Do Until p_oRSOthers.EOF
               If txtRoute = p_oRSOthers("sDestinat") Then
                  .Rows = lnCtr + 1 + 1
                  .TextMatrix(lnCtr + 1, 0) = lnCtr + 1
                  .TextMatrix(lnCtr + 1, 1) = p_oRSOthers("sDescript")
                  .TextMatrix(lnCtr + 1, 2) = p_oRSOthers("sReferNox")
                  .TextMatrix(lnCtr + 1, 3) = IIf(p_oRSOthers("cIncluded") = 0, "NO", "YES")
                  lnCtr = lnCtr + 1
               End If
               p_oRSOthers.MoveNext
            Loop
         Else
            .Rows = 2
            .TextMatrix(lnCtr + 1, 0) = 1
            .TextMatrix(lnCtr + 1, 1) = ""
            .TextMatrix(lnCtr + 1, 2) = ""
            .TextMatrix(lnCtr + 1, 3) = ""
         End If
         
         If .Rows > 23 Then
            .ColWidth(1) = 1950
         Else
            .ColWidth(1) = 2200
         End If
         
         .Row = 1
         .Col = 1
         .ColSel = .Cols - 1
      End With
   End With
End Sub

Private Sub createRoute()
   Set p_oRSOthers = New Recordset
   With p_oRSOthers
      .Fields.Append "sReferNox", adChar, 12
      .Fields.Append "sDescript", adVarChar, 50
      .Fields.Append "cIncluded", adChar, 1
      .Fields.Append "sSourcexx", adVarChar, 2
      .Fields.Append "sSourceCd", adChar, 4
      .Fields.Append "dTransact", adDate
      .Fields.Append "sDestinat", adVarChar, 50
      .Fields.Append "sRemarksx", adVarChar, 1024
      .Fields.Append "sBranchCd", adVarChar, 4
      .Open
   End With
   
   Set p_oRSDetail = New Recordset
   With p_oRSDetail
      .Fields.Append "sReferNox", adChar, 12
      .Fields.Append "sDescript", adVarChar, 50
      .Fields.Append "cIncluded", adChar, 1
      .Fields.Append "sSourcexx", adVarChar, 2
      .Fields.Append "sSourceCd", adChar, 4
      .Fields.Append "dTransact", adDate
      .Fields.Append "sDestinat", adVarChar, 50
      .Fields.Append "sRemarksx", adVarChar, 1024
      .Fields.Append "sBranchCd", adVarChar, 4
      .Open
   End With
   
   Set p_oRSRemove = New Recordset
   With p_oRSRemove
      .Fields.Append "sReferNox", adChar, 12
      .Fields.Append "sDescript", adVarChar, 50
      .Fields.Append "cIncluded", adChar, 1
      .Fields.Append "sSourcexx", adVarChar, 2
      .Fields.Append "sSourceCd", adChar, 4
      .Fields.Append "dTransact", adDate
      .Fields.Append "sDestinat", adVarChar, 50
      .Fields.Append "sRemarksx", adVarChar, 1024
      .Fields.Append "sBranchCd", adVarChar, 4
      .Open
   End With
End Sub

Private Sub loadSelected()
   Dim lnCtr As Integer

   With MSFlexGrid2
      p_oRSOthers.Filter = "cIncluded = " & strParm(xeYes)
      
      If p_oRSOthers.EOF Then
         .Rows = 2
         
         .TextMatrix(1, 0) = "1"
         .TextMatrix(1, 1) = ""
         .TextMatrix(1, 2) = ""
         .TextMatrix(1, 3) = ""
         .TextMatrix(1, 4) = ""
      Else
         .Rows = p_oRSOthers.RecordCount + 1
         
         p_oRSOthers.MoveFirst
         For lnCtr = 0 To p_oRSOthers.RecordCount - 1
            .TextMatrix(lnCtr + 1, 0) = lnCtr + 1
            .TextMatrix(lnCtr + 1, 1) = p_oRSOthers("sDescript")
            .TextMatrix(lnCtr + 1, 2) = p_oRSOthers("sReferNox")
            .TextMatrix(lnCtr + 1, 3) = p_oRSOthers("sRemarksx")
            .TextMatrix(lnCtr + 1, 4) = p_oRSOthers("sDestinat")
            p_oRSOthers.MoveNext
         Next
      End If
      
      p_oRSOthers.Filter = adFilterNone
      p_oRSOthers.Sort = "dTransact,sSourceCd,sReferNox"
   
      Select Case cmbSource.ListIndex
      Case 1
         p_oRSOthers.Filter = " sSourcexx = 'MC'"
      Case 2
         p_oRSOthers.Filter = " sSourcexx = 'SP'"
      Case 3
         p_oRSOthers.Filter = " sSourcexx = 'CP'"
      Case 4
         p_oRSOthers.Filter = " sSourcexx = 'Rg'"
      Case 5
         p_oRSOthers.Filter = " sSourcexx = 'As'"
      Case 6
         p_oRSOthers.Filter = " sSourcexx = 'SU'"
      End Select
      
      .Row = .Rows - 1
      .Col = 1
      .ColSel = .Cols - 1
   End With
End Sub

Private Function getBranch(ByVal lsValue As String, ByVal lbSearch As Boolean) As Boolean
   Dim lors As Recordset
   Dim lsSQL As String
   Dim lsMaster As String
   Dim lasMaster() As String
   Dim lsOldProc As String

   lsOldProc = "getBranch"
   'On Error GoTo errProc
   
   getBranch = False
   
   If lsValue <> "" Then
      If lsValue = txtRoute.Tag Then
         txtRoute.Text = txtRoute.Tag
         GoTo endProc
      End If
      
      If lbSearch Then
         lsMaster = "a.sBranchNm LIKE " & strParm(Trim(lsValue) & "%")
      Else
         lsMaster = "a.sBranchNm = " & strParm(Trim(lsValue))
      End If
   ElseIf lbSearch = False Then
      GoTo endWithClear
   End If
   
   If lsMaster <> "" Then
      lsSQL = AddCondition(getSQLBranch, lsMaster)
   Else
      GoTo endWithClear
   End If
   
   Set lors = New Recordset
   With lors
      Debug.Print lsSQL
      .Open lsSQL, oApp.Connection, , , adCmdText
      If .EOF Then
         GoTo endWithClear
      ElseIf .RecordCount = 1 Then
         p_sBranchCd = lors(0)
         
         txtRoute.Text = lors(1)
         txtRoute.Tag = lors(1)
      Else
         lsSQL = KwikBrowse(oApp, lors, _
                              "sBranchCd»sBranchNm", _
                              "Code»Branch")
         
         If lsSQL = Empty Then GoTo endWithClear
         lasMaster = Split(lsSQL, "»")
         
         p_sBranchCd = lasMaster(0)
         
         txtRoute.Text = lasMaster(1)
         txtRoute.Tag = lasMaster(1)
      End If
   End With

   txtRoute.Enabled = False
   getBranch = True
endProc:
   Exit Function
endWithClear:
   p_sBranchCd = ""
   txtRoute = ""
   txtRoute.Enabled = True
   GoTo endProc
errProc:
   ShowError lsOldProc & "( " & " )", True
End Function

Private Function getSQLBranch() As String
   getSQLBranch = "SELECT" & _
                     "  a.sBranchCd" & _
                     ", a.sBranchNm" & _
                     ", CONCAT(a.sAddressx, ', ', b.sTownName, ', ', c.sProvName) xAddressx" & _
                  " FROM Branch a" & _
                     ", TownCity b" & _
                     ", Province c" & _
                     ", Branch_Others d" & _
                  " WHERE a.sTownIDxx = b.sTownIDxx" & _
                     " AND b.sProvIDxx = c.sProvIDxx" & _
                     " AND a.cRecdStat = " & strParm(xeRecStateActive) & _
                     " AND a.sBranchCd = d.sBranchCd" & _
                     " AND d.sClustrID = " & strParm(p_sClustrID)
End Function

Private Function getSource(ByVal lsSourceCd As String) As String
   Select Case LCase(lsSourceCd)
   Case "mcdv"
      getSource = "MC Transfer"
   Case "spdl"
      getSource = "SP Transfer"
   Case "cpdl"
      getSource = "CP Transfer"
   Case "spwt"
      getSource = "SP Warranty Transfer"
   Case "cpjt"
      getSource = "CP Job Order"
   Case "asdl"
      getSource = "Asset Transfer"
   Case "sudl"
      getSource = "Supplies Transfer"
   Case "rgdl"
      getSource = "MC Reg Transfer"
   Case Else
      getSource = "Oth Transfer"
   End Select
End Function
