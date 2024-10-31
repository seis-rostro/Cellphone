VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Receipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type CheckInfo
   CheckNox As String
   AcctNoxx As String
   BankName As String
   CheckDte As Date
   Amountxx As Double
End Type

Private p_oAppDrivr As AppDriver
Private p_oForm As frmReceipt
Private p_oMod As New MainModules

Private p_bCancelxx As Boolean
Private p_dTransact As Date
Private p_sORNoxxxx As String
Private p_sReceived As String
Private p_sAddressx As String
Private p_nTranTotl As Double
Private p_nAmtPaidx As Double
Private p_nCashAmtx As Double
Private p_nCheckAmt As Double
Private p_sRemarksx As String
Private p_bEmptyORx As Boolean
Private p_axChecks(4) As CheckInfo

Private bLoaded As Boolean

Dim pnCtr As Integer

Property Set AppDriver(Value As AppDriver)
10       Set p_oAppDrivr = Value
End Property

Property Get AllowEmptyOR() As Boolean
10       AllowEmptyOR = p_bEmptyORx
End Property

Property Let AllowEmptyOR(ByVal Value As Boolean)
10       p_bEmptyORx = Value
End Property

Property Get Checks(ByVal Row As Integer, ByVal Index As Variant) As Variant
10       If Row > 4 Then Exit Property

20       Select Case Index
      Case 0, "CheckNox"
30             Checks = p_axChecks(Row).CheckNox
40          Case 1, "sAcctNoxx"
50             Checks = p_axChecks(Row).AcctNoxx
60          Case 2, "sBankName"
70             Checks = p_axChecks(Row).BankName
80          Case 3, "dCheckDte"
90             Checks = p_axChecks(Row).CheckDte
100         Case 4, "nAmountxx"
110            Checks = p_axChecks(Row).Amountxx
120      End Select
End Property

Property Let Checks(ByVal Row As Integer, ByVal Index As Variant, ByVal Value As Variant)
10       If Row > 4 Then Exit Property

20       Select Case Index
      Case 0, "CheckNox"
30             p_axChecks(Row).CheckNox = Value
40          Case 1, "sAcctNoxx"
50             p_axChecks(Row).AcctNoxx = Value
60          Case 2, "sBankName"
70             p_axChecks(Row).BankName = Value
80          Case 3, "dCheckDte"
90             p_axChecks(Row).CheckDte = Value
100         Case 4, "nAmountxx"
110            p_axChecks(Row).Amountxx = Value
120      End Select
End Property

Property Get TranTotal() As Double
10       TranTotal = p_nTranTotl
End Property

Property Let TranTotal(ByVal Value As Double)
10       p_nTranTotl = Value
End Property

Property Get AmountPaid() As Double
10       AmountPaid = p_nAmtPaidx
End Property

Property Let AmountPaid(ByVal Value As Double)
10       p_nAmtPaidx = Value
End Property

Property Get CashAmount() As Double
10       CashAmount = p_nCashAmtx
End Property

Property Let CashAmount(ByVal Value As Double)
10       p_nCashAmtx = Value
End Property

Property Get CheckAmount() As Double
10       CheckAmount = p_nCheckAmt
End Property

Property Let CheckAmount(ByVal Value As Double)
10       p_nCheckAmt = Value
End Property

Property Get Remarks() As String
10       Remarks = p_sRemarksx
End Property

Property Let Remarks(ByVal Value As String)
10       p_sRemarksx = Value
End Property

Property Let Address(ByVal Value As String)
10       p_sAddressx = Value
End Property

Property Let TranDate(ByVal Value As Date)
10       p_dTransact = Value
End Property

Property Let ReceiveFrom(ByVal Value As String)
10       p_sReceived = Value
End Property

Property Get Cancelled() As Boolean
10       Cancelled = p_bCancelxx
End Property

Property Get ORNo() As String
10       ORNo = p_sORNoxxxx
End Property

Property Let ORNo(ByVal Value As String)
10       p_sORNoxxxx = Value
End Property

Sub InitReceipt()
10       If p_oAppDrivr Is Nothing Then Exit Sub
   
20       Set p_oForm.AppDriver = p_oAppDrivr
30       p_bCancelxx = False
40       p_dTransact = Date
50       p_sORNoxxxx = ""
60       p_sReceived = ""
70       p_sAddressx = ""
80       p_nTranTotl = 0#
90       p_nCashAmtx = 0#
100      p_nCheckAmt = 0#
110      p_sRemarksx = ""
   
120      If p_bEmptyORx = False Then
130         p_sORNoxxxx = p_oForm.GetNextOR
140      Else
150         p_sORNoxxxx = ""
160      End If
   
170      InitChecks
End Sub

Sub ShowReceipt()
10       If p_oAppDrivr Is Nothing Then Exit Sub
   
20       Load p_oForm
30       With p_oForm
40          .lblHead(0) = p_oAppDrivr.ClientName
50          .lblHead(1) = Trim(p_oAppDrivr.Address) & ", " & Trim(p_oAppDrivr.TownCity) & ", " & _
            Trim(p_oAppDrivr.Province) & " " & Trim(p_oAppDrivr.ZippCode) & ", "
60          .lblHead(2) = "Tel No: " & p_oAppDrivr.TelNo
      
70          .AllowEmptyOR = p_bEmptyORx
80          .txtField(0) = p_sORNoxxxx
90          .txtField(1) = Format(p_dTransact, "MMMM DD, YYYY")
100         .txtField(2) = p_sReceived
110         .txtField(3) = p_sAddressx
120         .txtField(4) = p_sRemarksx
130         .txtField(5) = Format(p_nCashAmtx, "#,##0.00")
140         .AmountPaid = Format(p_nAmtPaidx, "#,##0.00")
150         .txtField(6) = Format(p_nCheckAmt, "#,##0.00")
160         .lblPayment = Format(p_nTranTotl, "#,##0.00")
170         .lblChange = Format(p_nAmtPaidx - p_nTranTotl, "#,##0.00")
      
180         For pnCtr = 0 To 4
190            .GridEditor1.TextMatrix(pnCtr + 1, 1) = p_axChecks(pnCtr).CheckNox
200            .GridEditor1.TextMatrix(pnCtr + 1, 2) = p_axChecks(pnCtr).AcctNoxx
210            .GridEditor1.TextMatrix(pnCtr + 1, 3) = p_axChecks(pnCtr).BankName
220            .GridEditor1.TextMatrix(pnCtr + 1, 5) = p_axChecks(pnCtr).Amountxx
230            .GridEditor1.TextMatrix(pnCtr + 1, 4) = p_axChecks(pnCtr).CheckDte
240         Next
      
250         .Show 1
      
260         p_bCancelxx = .Cancelled
      
270         p_sORNoxxxx = .txtField(0)
280         p_dTransact = CDate(.txtField(1))
290         p_sReceived = .txtField(2)
300         p_sAddressx = .txtField(3)
310         p_sRemarksx = .txtField(4)
320         p_nCashAmtx = CDbl(.txtField(5))
330         p_nCheckAmt = CDbl(.txtField(6))

340         For pnCtr = 0 To 4
350            p_axChecks(pnCtr).CheckNox = .GridEditor1.TextMatrix(pnCtr + 1, 1)
360            p_axChecks(pnCtr).AcctNoxx = .GridEditor1.TextMatrix(pnCtr + 1, 2)
370            p_axChecks(pnCtr).BankName = .GridEditor1.TextMatrix(pnCtr + 1, 3)
380            p_axChecks(pnCtr).CheckDte = CDate(.GridEditor1.TextMatrix(pnCtr + 1, 4))
390            p_axChecks(pnCtr).Amountxx = CDbl(.GridEditor1.TextMatrix(pnCtr + 1, 5))
400         Next
410      End With
420      Unload p_oForm
End Sub

Private Sub InitChecks()
10       For pnCtr = 0 To 4
20          p_axChecks(pnCtr).CheckNox = ""
30          p_axChecks(pnCtr).AcctNoxx = ""
40          p_axChecks(pnCtr).BankName = ""
50          p_axChecks(pnCtr).CheckDte = p_dTransact
60          p_axChecks(pnCtr).Amountxx = 0#
70       Next
End Sub

Private Sub Class_Initialize()
10       Set p_oForm = New frmReceipt
End Sub

Private Sub Class_Terminate()
10       Set p_oForm = Nothing
20       Set p_oMod = Nothing
End Sub


