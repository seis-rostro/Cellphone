VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmMCARLedger 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Client Ledger"
   ClientHeight    =   7680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9060
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   9060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   285
      Left            =   135
      TabIndex        =   10
      Top             =   7260
      Visible         =   0   'False
      Width           =   8520
      _ExtentX        =   15028
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5265
      Left            =   120
      TabIndex        =   9
      Tag             =   "et0;eb0;et0;bc2"
      Top             =   2295
      Width           =   8550
      _ExtentX        =   15081
      _ExtentY        =   9287
      _Version        =   393216
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1665
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   6885
      _ExtentX        =   12144
      _ExtentY        =   2937
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   1590
         TabIndex        =   8
         Top             =   1230
         Width           =   5145
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   1590
         TabIndex        =   6
         Top             =   930
         Width           =   5145
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
         Height          =   285
         Index           =   0
         Left            =   1590
         TabIndex        =   2
         Top             =   150
         Width           =   1950
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   1590
         TabIndex        =   4
         Top             =   630
         Width           =   5145
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client ID"
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
         Left            =   210
         TabIndex        =   1
         Top             =   195
         Width           =   750
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1665
         Tag             =   "et0;ht2"
         Top             =   240
         Width           =   1950
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name"
         Height          =   195
         Index           =   10
         Left            =   210
         TabIndex        =   3
         Top             =   690
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Address"
         Height          =   195
         Index           =   11
         Left            =   210
         TabIndex        =   7
         Top             =   1005
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company "
         Height          =   195
         Index           =   12
         Left            =   210
         TabIndex        =   5
         Top             =   1305
         Width           =   705
      End
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   1665
      Left            =   7035
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   2937
      Begin xrControl.xrButton cmdButton 
         Height          =   450
         Left            =   120
         TabIndex        =   0
         Top             =   1125
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   794
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
         Picture         =   "frmMCARLedger.frx":0000
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00E0E0E0&
         Height          =   945
         Index           =   1
         Left            =   150
         Top             =   135
         Width           =   1290
      End
      Begin VB.Shape Shape2 
         Height          =   1005
         Index           =   0
         Left            =   120
         Top             =   105
         Width           =   1350
      End
   End
End
Attribute VB_Name = "frmMCARLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmMCARLedger"

Private oSkin As clsFormSkin
Private oRS As ADODB.Recordset
Private oRSMaster As ADODB.Recordset

Dim psAcctNo As String
Dim pbLoaded As Boolean

Property Let AccountNo(lsAcctNo As String)
   psAcctNo = lsAcctNo
End Property

Private Sub cmdButton_Click()
   Unload Me
End Sub

Public Function browseLedger() As Boolean
   Dim lsSQL As String, lsOldProc As String
   Dim lnCtr As Integer

   lsOldProc = "browseLedger"
   browseLedger = False
   '''On Error GoTo errProc
   
   If pbLoaded = False Then GoTo endProc
   Set oRS = New ADODB.Recordset
   If oRS.State = adStateOpen Then oRS.Close

   lsSQL = "SELECT * FROM (SELECT" _
               & "  a.sClientID" _
               & ", c.sCompnyNm" _
               & ", CONCAT(c.sLastName, ', ', c.sFrstName, ' ', c.sMiddName) xFullName" _
               & ", CONCAT(c.sAddressx, ', ', d.sTownName, ', ' , e.sProvName, ' ', d.sZippCode) xAddressx" _
               & ", b.dTransact" _
               & ", b.cOffPaymx" _
               & ", b.cTrantype" _
               & ", b.sORNoxxxx" _
               & ", b.nTranAmtx" _
               & ", b.nRebatesx" _
               & ", b.nOthersxx" _
               & ", b.nABalance" _
               & ", b.nDebitAmt" _
               & ", b.nMonDelay" _
               & ", CONCAT(f.sLastName, ', ', f.sFrstName, ' ', f.sMiddName) xCollectr" _
               & ", b.sRemarksx" _
               & ", g.sBranchNm" _
               & ", b.nEntryNox" _
               & ", b.sAcctNmbr"
   
   lsSQL = lsSQL _
            & " FROM MC_AR_Master a" _
               & ", MC_AR_Ledger b" _
                  & " LEFT JOIN Employee_Master f" _
                     & " ON b.sCollIDxx = f.sEmployID" _
               & ", Client_Master c" _
                  & " LEFT JOIN TownCity d" _
                     & " ON c.sTownIDxx = d.sTownIDxx" _
                  & " LEFT JOIN Province e" _
                     & " ON d.sProvIDxx = e.sProvIDxx" _
               & ", Branch g" _
            & " WHERE a.sAcctNmbr = b.sAcctNmbr" _
               & " And a.sClientID = c.sClientID" _
               & " And a.sAcctNmbr = " & strParm(psAcctNo) _
               & " And b.sBranchCd = g.sBranchCd" _
               & " AND (((b.cOffPaymx = '0' or b.cOffPaymx = '2') AND NOT CONCAT(f.sLastName, ', ', f.sFrstName, ' ', f.sMiddName) IS NULL)" _
                  & " OR ((b.cOffPaymx = '1' or b.cOffPaymx = '3') AND NOT g.sBranchNm IS NULL))"
   
   lsSQL = lsSQL _
            & " UNION SELECT" _
               & "  a.sClientID" _
               & ", c.sCompnyNm" _
               & ", CONCAT(c.sLastName, ', ', c.sFrstName, ' ', c.sMiddName) xFullName" _
               & ", CONCAT(c.sAddressx, ', ', d.sTownName, ', ' , e.sProvName, ' ', d.sZippCode) xAddressx" _
               & ", b.dTransact" _
               & ", b.cOffPaymx" _
               & ", b.cTrantype" _
               & ", b.sORNoxxxx" _
               & ", b.nTranAmtx" _
               & ", b.nRebatesx" _
               & ", b.nOthersxx" _
               & ", b.nABalance" _
               & ", b.nDebitAmt" _
               & ", b.nMonDelay" _
               & ", CONCAT(h.sLastName, ', ', h.sFrstName, ' ', h.sMiddName) xCollectr" _
               & ", b.sRemarksx" _
               & ", g.sBranchNm" _
               & ", b.nEntryNox" _
               & ", b.sAcctNmbr"
   
   lsSQL = lsSQL _
            & " FROM MC_AR_Master a" _
               & ", MC_AR_Ledger b" _
                  & " LEFT JOIN Employee_Master001 f" _
                     & " LEFT JOIN Client_Master h" _
                        & " ON f.sEmployID = h.sClientID" _
                     & " ON b.sCollIDxx = f.sEmployID" _
               & ", Client_Master c" _
                  & " LEFT JOIN TownCity d" _
                     & " ON c.sTownIDxx = d.sTownIDxx" _
                  & " LEFT JOIN Province e" _
                     & " ON d.sProvIDxx = e.sProvIDxx" _
               & ", Branch g" _
            & " WHERE a.sAcctNmbr = b.sAcctNmbr" _
               & " And a.sClientID = c.sClientID" _
               & " And a.sAcctNmbr = " & strParm(psAcctNo) _
               & " And b.sBranchCd = g.sBranchCd" _
               & " AND (((b.cOffPaymx = '0' or b.cOffPaymx = '2') AND NOT CONCAT(h.sLastName, ', ', h.sFrstName, ' ', h.sMiddName) IS NULL)" _
                  & " OR ((b.cOffPaymx = '1' or b.cOffPaymx = '3') AND NOT g.sBranchNm IS NULL))" _
            & " ORDER BY dTransact) xSourceTable GROUP BY sAcctNmbr, nEntryNox, sORNoxxxx"

   oRS.Open lsSQL, oApp.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
   If oRS.EOF Then GoTo endProc
   browseLedger = True
Debug.Print lsSQL
endProc:
   'Set oRS = Nothing
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lsUserID As String, lsUserName As String, lsOldProc As String
Dim lnUserRights As Integer, lnRep As String

   lsUserID = oApp.UserID
   If KeyCode = vbKeyF12 Then
      If oApp.UserLevel < xeManager Then
         If Not GetApproval(oApp, lnUserRights, lsUserID, lsUserName, "mnuActiveAccount") Then
            KeyCode = 0
            Exit Sub
         End If
         
         If lnUserRights < xeManager Then
            MsgBox "Approving User is not Authorized" & vbCrLf & _
            "Please Contact SSG/SEG!!!", vbCritical, "Warning"
            Exit Sub
         End If
      End If
         
      If SearchTransaction(psAcctNo, True, False) Then
         If reCalc() Then
         Call LoadDetail
         MsgBox "Transaction Updated Successfully!!!", vbInformation, "Notice"
         End If
      End If
      
      KeyCode = 0
   End If
End Sub

Private Sub Form_Load()
   Dim lnCtr As Integer
   Dim lsOldProc As String
   
   lsOldProc = "Form_Load"
   '''On Error GoTo errProc

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   
   oSkin.ApplySkin xeFormLedger
   
   For lnCtr = 0 To txtField.Count - 1
      txtField(lnCtr).Locked = True
   Next
   
   pbLoaded = True
   
   
'      ProgressBar1.Visible = True
'      ProgressBar1.Max = IIf(oRS.RecordCount = 0, 1, oRS.RecordCount)
   
   With MSFlexGrid1
      .Cols = 13
      .Font = "MS Sans Serif"

      'column title
      .TextMatrix(0, 1) = "Date"
      .TextMatrix(0, 2) = "PC"
      .TextMatrix(0, 3) = "TC"
      .TextMatrix(0, 4) = "OR No"
      .TextMatrix(0, 5) = "Amount"
      .TextMatrix(0, 6) = "Rebates"
      .TextMatrix(0, 7) = "Others"
      .TextMatrix(0, 8) = "ABalance"
      .TextMatrix(0, 9) = "Debit"
      .TextMatrix(0, 10) = "MonDelay"
      .TextMatrix(0, 11) = "Collector"
      .TextMatrix(0, 12) = "Remarks"
      
      
      
      
      'column alignment
      .Row = 0
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = 1
      Next
      
      'column width
      .ColWidth(0) = 330
      .ColWidth(1) = 1100
      .ColWidth(2) = 330
      .ColWidth(3) = 330
      .ColWidth(4) = 1000
      .ColWidth(5) = 1100
      .ColWidth(6) = 1060
      .ColWidth(7) = 1000
      .ColWidth(8) = 1000
      .ColWidth(9) = 1100
      .ColWidth(10) = 1060
      .ColWidth(11) = 2500
      .ColWidth(12) = 2500
      
      .ColAlignment(1) = 1
      .ColAlignment(2) = 3
      .ColAlignment(3) = 3
      .ColAlignment(4) = 1
      
      .Row = 1
      .Col = 1
      .ColSel = .Cols - 1
   End With
   LoadDetail

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
End Sub

Private Sub MSFlexGrid1_GotFocus()
   MSFlexGrid1.BackColorSel = &HA4A36A
End Sub

Private Sub MSFlexGrid1_LostFocus()
   MSFlexGrid1.BackColorSel = &H8000000D
End Sub

Private Function SearchTransaction(Optional sValue As Variant, Optional bByCode As Variant, Optional bSearch As Variant) As Boolean
   Dim lsOldProc As String
   Dim lsBrowse As String
   Dim lsSQL As String

   lsOldProc = "SearchTransaction"
'   '''On Error GoTo errProc
   SearchTransaction = False

   Set oRSMaster = New ADODB.Recordset

   lsSQL = "Select" _
               & "  a.sAcctNmbr" _
               & ", CONCAT(b.sLastName, ', ', b.sFrstName, ' ', b.sMiddName) xFullName" _
               & ", CONCAT(g.sBrandNme, ' ', f.sModelNme, ' ( ', RTrim(e.sEngineNo), ' )') xModelNme" _
               & ", CONCAT(i.sLastName, ', ', i.sFrstName, ' ', i.sMiddName) xCollectr" _
               & ", a.dPurchase" _
               & ", a.dFirstPay" _
               & ", a.nAcctTerm" _
               & ", a.dDueDatex" _
               & ", a.nMonAmort" _
               & ", a.cAcctStat" _
               & ", a.nGrossPrc" _
               & ", a.nDownPaym" _
               & ", a.nCashBalx" _
               & ", a.nPNValuex" _
               & ", a.nPenaltyx" _
               & ", a.nRebatesx" _
               & ", j.sBranchNm" _
               & ", a.nPaymTotl" _
               & ", a.nRebTotlx" _
               & ", a.nPenTotlx" _
               & ", CONCAT(b.sAddressx, ', ', c.sTownName, ', ', d.sProvName, ' ', c.sZippCode) xAddressx" _
               & ", h.sBranchCd" _
               & ", a.nABalance"
   lsSQL = lsSQL _
               & ", a.nDebtTotl" _
               & ", a.nCredTotl" _
               & ", a.nDownTotl" _
               & ", a.nCashTotl" _
               & ", a.cRatingxx" _
               & ", a.nLastPaym" _
               & ", a.dLastPaym" _
               & ", a.nAmtDuexx" _
               & ", a.nDelayAvg" _
               & ", a.nLedgerNo" _
               & ", a.sModified" _
               & ", a.dModified" _
               & ", a.dClosedxx" _
               & ", CONCAT(k.sLastName, ', ', k.sFrstName, ' ', k.sMiddName) xCoCltNm1" _
               & ", CONCAT(l.sLastName, ', ', l.sFrstName, ' ', l.sMiddName) xCoCltNm2" _
               & ", a.cLoanType" _
               & ", CONCAT(n.sLastName, ', ', n.sFrstName, ' ', n.sMiddName) xCoMakrNm" _
               & ", a.nLgrLinex" _
               & ", a.nPassLine" _
               & ", CONCAT(p.sLastName, ', ', p.sFrstName, ' ', p.sMiddName) zCollectr" _
               & ", a.cLoanType"
               
   lsSQL = lsSQL _
            & " From MC_AR_Master a" _
               & " LEFT JOIN MC_Serial e" _
                  & " On a.sSerialID = e.sSerialID" _
               & " Left Join MC_Model f" _
                  & " On e.sModelIDx = f.sModelIDx" _
               & " Left Join Brand g" _
                  & " On f.sBrandIDx = g.sBrandIDx" _
               & " Left Join Client_Master k" _
                  & " On a.sCoCltID1 = k.sClientID" _
               & " Left Join Client_Master l" _
                  & " On a.sCoCltID2 = l.sClientID" _
               & " Left Join MC_Credit_Application m" _
                  & " Left Join Client_Master n" _
                     & " On m.sCoMakrID = n.sClientID" _
                  & " On a.sApplicNo = m.sTransNox" _
                  & " And a.sClientID = m.sClientID" _
               & ", Client_Master b" _
                  & " Left Join TownCity c" _
                     & " On b.sTownIDxx = c.sTownIDxx" _
                  & " Left Join Province d" _
                     & " On c.sProvIDxx = d.sProvIDxx"
   lsSQL = lsSQL _
               & ", Route_Area h" _
                  & " LEFT JOIN Employee_Master i" _
                     & " ON h.sCollctID = i.sEmployID" _
                  & " LEFT JOIN Employee_Master001 o" _
                     & " LEFT JOIN Client_Master p" _
                        & " ON o.sEmployID = p.sClientID" _
                     & " ON h.sCollctID = o.sEmployID" _
               & ", Branch j"

   lsSQL = lsSQL _
            & " Where a.sClientID = b.sClientID" _
               & " And a.sRouteIDx = h.sRouteIDx" _
               & " And h.sBranchCd = j.sBranchCd" _
               & " And a.cAcctstat = '0'"
      If Not IsMissing(sValue) Then
      If Not IsMissing(bByCode) Then
         If bByCode Then
            lsSQL = lsSQL & " And a.sAcctNmbr = " & strParm(sValue)
         Else
            lsSQL = lsSQL & " And CONCAT(b.sLastName, ', ', b.sFrstName, ' ', b.sMiddName) Like " & strParm(Trim(sValue) & "%")
         End If
      Else
         lsSQL = lsSQL & " And CONCAT(b.sLastName, ', ', b.sFrstName, ' ', b.sMiddName) = " & strParm(Trim(sValue))
      End If
   End If
   lsSQL = lsSQL & " Order By sAcctNmbr, xFullName"

   oRSMaster.Open lsSQL, oApp.Connection, adOpenStatic, adLockOptimistic, adCmdText

   If oRSMaster.EOF Then GoTo endProc
   
   SearchTransaction = True
   Set oRSMaster.ActiveConnection = Nothing

endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " _
                       & "  " & IFNull(sValue) _
                       & ", " & IFNull(bByCode) _
                       & ", " & IFNull(bSearch) _
                       & " )"
End Function

Private Function reCalc() As Boolean
   Dim lsOldProc As String
   Dim lsSQL As String

   lsOldProc = "reCalc"
   '''On Error GoTo errProc

   oApp.BeginTrans

   If Recalculate(oRSMaster, oApp) = False Then GoTo endProc
   Debug.Print oRSMaster("nLedgerNo")
   lsSQL = ADO2SQL(oRSMaster, "MC_AR_Master", _
                     "sAcctNmbr = " & strParm(oRSMaster("sAcctNmbr")), _
                     Encrypt(oApp.UserID), _
                     oApp.ServerDate, _
                     "xFullName»xModelNme»xCollectr»xAddressx»sBranchCd")
   Debug.Print lsSQL
   If lsSQL <> "" Then
      If oApp.Execute(lsSQL, "MC_AR_Master") = 0 Then
         MsgBox "Unable to Save Loan Receivable Master!!!", vbCritical, "Warning"
         GoTo endWithRoll
      End If
   End If

   oApp.CommitTrans
   reCalc = True

endProc:
   Exit Function
endWithRoll:
   oApp.RollbackTrans
   GoTo endProc
errProc:
   oApp.RollbackTrans
   ShowError lsOldProc & "( " & " )"
End Function

Private Sub LoadDetail()
   Dim lnCtr As Integer
   
   If browseLedger Then
      txtField(0).Text = Format(oRS("sClientID"), IIf(Len(oApp.BranchCode) = 2, "@@-@@@@@@", "@@@@-@@@@@@"))
      txtField(1).Text = oRS("xFullName")
      txtField(2).Text = oRS("xAddressx")
      txtField(3).Text = IIf(IsNull(oRS("sCompnyNm")), "", oRS("sCompnyNm"))

      With MSFlexGrid1
          .Rows = IIf(oRS.RecordCount = 0, 2, oRS.RecordCount + 1)
            For lnCtr = 0 To oRS.RecordCount - 1
               .TextMatrix(lnCtr + 1, 0) = lnCtr + 1
               .TextMatrix(lnCtr + 1, 1) = Format(oRS("dTransact"), "MMM-DD-YYYY")
               .TextMatrix(lnCtr + 1, 2) = IIf((oRS("cOffPaymx")) = 0 Or (oRS("cOffPaymx")) = 2, "F", "O")
               Select Case LCase(oRS("cTrantype"))
               Case "p"
                  .TextMatrix(lnCtr + 1, 3) = "MP"
               Case "d"
                  .TextMatrix(lnCtr + 1, 3) = "DP"
               Case "m"
                  .TextMatrix(lnCtr + 1, 3) = "Dm"
               Case "c"
                  .TextMatrix(lnCtr + 1, 3) = "Cm"
               Case "b"
                  .TextMatrix(lnCtr + 1, 3) = "CB"
               End Select
               
               .TextMatrix(lnCtr + 1, 4) = oRS("sORNoxxxx")
               .TextMatrix(lnCtr + 1, 5) = Format(oRS("nTranAmtx"), "#,##0.00")
               .TextMatrix(lnCtr + 1, 6) = Format(oRS("nRebatesx"), "#,##0.00")
               .TextMatrix(lnCtr + 1, 7) = Format(oRS("nOthersxx"), "#,##0.00")
               .TextMatrix(lnCtr + 1, 8) = Format(oRS("nABalance"), "#,##0.00")
               .TextMatrix(lnCtr + 1, 9) = Format(oRS("nDebitAmt"), "#,##0.00")
               .TextMatrix(lnCtr + 1, 10) = Format(oRS("nMonDelay"), "#,##0.00")
               .TextMatrix(lnCtr + 1, 11) = IIf(IsNull(oRS("xCollectr")), oRS("sBranchNm"), IIf(Trim(oRS("xCollectr")) = "", oRS("sBranchNm"), oRS("xCollectr")))
               .TextMatrix(lnCtr + 1, 12) = oRS("sRemarksx")
      
               oRS.MoveNext
            Next
      End With
   End If
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

