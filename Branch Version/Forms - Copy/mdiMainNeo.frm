VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm mdiMain 
   BackColor       =   &H8000000C&
   Caption         =   "Guanzon Telecom Point-Of-Sale & Inventory System Branch Version"
   ClientHeight    =   7365
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   11280
   Icon            =   "mdiMainNeo.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "mdiMainNeo.frx":424A
   ScrollBars      =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer tmeLog 
      Interval        =   1000
      Left            =   675
      Top             =   405
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   75
      Top             =   390
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMainNeo.frx":769FB
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   7065
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6668
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1764
            MinWidth        =   1764
            Object.ToolTipText     =   "Edit Mode"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Text            =   "Dl"
            TextSave        =   "Dl"
            Object.ToolTipText     =   "Motorcycle Monitoring"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Text            =   "Dl"
            TextSave        =   "Dl"
            Object.ToolTipText     =   "Spareparts Monitoring"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Text            =   "Ck"
            TextSave        =   "Ck"
            Object.ToolTipText     =   "Other Monitoring"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
            Object.ToolTipText     =   "Branch"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2646
            MinWidth        =   2646
            Object.ToolTipText     =   "Current User"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1850
            MinWidth        =   1850
            Object.ToolTipText     =   "System Date"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lnGrayText As Long
Dim pbProcClassify As Boolean
Private Const pxeJavaPath As String = "D:\GGC_Java_Systems\"


Private Sub MDIForm_Click()
'   Dim lsSQL As String
'   Dim loControl As Control
'
'   For Each loControl In mdiMain
'      If TypeName(loControl) = "Menu" Then
'         lsSQL = "INSERT INTO xxxMenuObject " & _
'                     "( sMenuIDxx" & _
'                     ", sMenuName" & _
'                     ", sProdctID" & _
'                     ", sMenuDesc" & _
'                     ", sRemarksx" & _
'                     ", nUserRght" & _
'                     ", nAddRight" & _
'                     ", nUpdRight" & _
'                     ", nDelRight" & _
'                     ", nCanRight" & _
'                  " ) VALUES ( " & _
'                     strParm(GetNextCode("xxxMenuObject", "sMenuIDxx", True, oApp.Connection, True, oApp.BranchCode)) & _
'                     ", " & strParm(loControl.Name) & _
'                     ", " & strParm(oApp.ProductID) & _
'                     ", " & strParm(loControl.Caption) & _
'                     ", " & strParm("") & _
'                     ", " & 255 & _
'                     ", " & 231 & _
'                     ", " & 240 & _
'                     ", " & 192 & _
'                     ", " & 224 & " )"
'
'         oApp.Execute lsSQL, "xxxMenuObject", oApp.BranchCode
'      End If
'   Next
'
'   MsgBox "Tapos Na Po!!!"
End Sub

Private Sub MDIForm_Load()
   lnGrayText = GetSysColor(17)
   setGrayText oApp.getColor("ET0")
   
   If oApp.IsWarehouse Then
      mdiMain.mnuCPSerial.Visible = oApp.UserLevel = xeManager
   Else
      mdiMain.mnuCPSerial.Visible = oApp.UserLevel = xeEngineer
   End If
   mdiMain.mnuManagerRep.Visible = oApp.IsMainOffice = True Or oApp.IsWarehouse = True
   mdiMain.mnuWholeSale.Visible = oApp.IsMainOffice = True Or oApp.IsWarehouse = True
   mdiMain.mnuWholeSaleReturn.Visible = oApp.IsMainOffice = True Or oApp.IsWarehouse = True
   mdiMain.mnuChargeInvoice.Visible = oApp.IsMainOffice = True Or oApp.IsWarehouse = True
   mdiMain.mnuMarketingSupport.Visible = oApp.IsMainOffice = True Or oApp.IsWarehouse = True
   mdiMain.mnuPriceProtection.Visible = oApp.IsMainOffice = True Or oApp.IsWarehouse = True
   mdiMain.mnuCPClustering.Visible = LCase(oApp.ProductID) = "telecom1" And oApp.IsWarehouse = True
   mdiMain.mnuDelSched.Visible = (oApp.UserLevel = xeManager Or oApp.UserLevel = xeSupervisor Or oApp.UserLevel = xeEngineer)
'   mdiMain.mnuAppApproval.Visible = oApp.UserLevel = xeEngineer
   
   'mac 2021.10.20
   '  service phone tagging menu visibility
   mdiMain.mnuServicePhoneTransfer.Visible = (LCase(oApp.BranchCode) = "c0w6" Or LCase(oApp.BranchCode) = "c0w2")
   mdiMain.mnuServicePhoneTransferReg.Visible = (LCase(oApp.BranchCode) = "c0w6" Or LCase(oApp.BranchCode) = "c0w2")
   mdiMain.mnuReceiveServicePhone.Visible = LCase(oApp.BranchCode) = "c0w6" Or LCase(oApp.BranchCode) = "c0a9"
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
   Dim oForm As Object
'   If UnloadMode = 0 Or UnloadMode = 1 Then oApp.LogOutUser

   For Each oForm In Forms
      Unload oForm
   Next

   setGrayText lnGrayText
   Set oApp = Nothing
End Sub

Private Sub mnuAForwarded_Click()
   frmCP_AccessJobOrderPosting.Tag = "mnuAForwarded"
   frmCP_AccessJobOrderPosting.Show
End Sub

Private Sub mnuAppApproval_Click()
   frmMPCreditApproval.Tag = "mnuAppApproval"
   frmMPCreditApproval.Show
End Sub

Private Sub mnuAppEntry_Click()
   frmMPCreditApp.Tag = "mnuAppEntry"
   frmMPCreditApp.Show
End Sub

Private Sub mnuARAdjustment_Click()
'   frmPaymentAdjustment.Tag = "mnuARAdjustment"
'   frmPaymentAdjustment.Show
End Sub

Private Sub mnuARPayment_Click()
'   frmPaymentAdjustment.Tag = "mnuARPayment"
'   frmPaymentAdjustment.Show
End Sub

Private Sub mnuAServiceCenter_Click()
   frmCP_AccessJobOrder.Tag = "mnuAServiceCenter"
   frmCP_AccessJobOrder.Show
End Sub

Private Sub mnuAssetRequest_Click()
   frmAssetStockRequest.Tag = "mnuAssetRequest"
   frmAssetStockRequest.Show
End Sub

Private Sub mnuAssets_Click()
   frmAssets.Tag = "mnuAssets"
   frmAssets.Show
End Sub

Private Sub mnuAssetsBrand_Click()
   frmAssetsBrand.Tag = "mnuAssetsBrand"
   frmAssetsBrand.Show
End Sub

Private Sub mnuAssetsMaintenance_Click()
   frmAssetMaintenance.Tag = "mnuAssetsMaintenance"
   frmAssetMaintenance.Show
End Sub

Private Sub mnuAssetsPORec_Click()
   frmAssetPOReceiving.Tag = "mnuAssetsPORec"
   frmAssetPOReceiving.Show
End Sub

Private Sub mnuAssetsTransfer_Click()
   frmAssetStockTransferRec.Tag = "mnuAssetsTransfer"
   frmAssetStockTransferRec.Show
End Sub

Private Sub mnuAssetTransfer_Click()
   frmAssetStockTransfer.Tag = "mnuAssetTransfer"
   frmAssetStockTransfer.Show
End Sub

Private Sub mnuAssetTransferAcceptance_Click()
   frmAssetStockTransferRec.Tag = "mnuAssetTransfer"
   frmAssetStockTransferRec.Show
End Sub

Private Sub mnuAuditReports_Click()
   Dim loReports As clsCPAuditRep
   Dim loRepViewer As frmRepViewer

   Set loReports = New clsCPAuditRep
   With loReports
      Set .AppDriver = oApp
      If .ShowReport Then
         Set loRepViewer = New frmRepViewer
         Set loRepViewer.ReportSource = .Source

         loRepViewer.Show
         .CloseReport
      End If
   End With
End Sub

Private Sub mnuCardRate_Click()
   If oApp.UserLevel >= xeManager Then
      frmCPCardRate.Tag = "mnuCardRate"
      frmCPCardRate.Show
   End If
End Sub

Private Sub mnuCardRateModel_Click()
'disble the user level condition. the encoder of the promo is from MP Executive
'   If oApp.UserLevel >= xeAudit Then
      frmCPCardRatePromo.Tag = "mnuCardRateModel"
      frmCPCardRatePromo.Show
'   End If
End Sub

Private Sub mnuCashDep_Click()
   frmCashDeposit.Tag = "mnuCashDep"
   frmCashDeposit.Show
End Sub

Private Sub mnuCashDepReg_Click()
   frmCashDepositReg.Tag = "mnuCashDepReg"
   frmCashDepositReg.Show
End Sub

Private Sub mnuCashDisbursement_Click()
   frmCashDisbursement.Tag = "mnuCashDisbursement"
   frmCashDisbursement.Show
End Sub

Private Sub mnuCashDisbursementApprvl_Click()
   frmCashDisbursementApprvl.Tag = "mnuCashDisbursementApprvl"
   frmCashDisbursementApprvl.Show
End Sub

Private Sub mnuCashDisReg_Click(Index As Integer)
   frmCashDisbursementReg.Tag = "mnuCashDisReg"
   frmCashDisbursementReg.Show
End Sub

Private Sub mnuCharge_Invoice_Reg_Click()
   frmCP_Charge_Invoice_Reg.Tag = "mnuCharge_Invoice_Reg"
   frmCP_Charge_Invoice_Reg.Show
End Sub

Private Sub mnuChargeInvoicePosting_Click()
   frmCP_Charge_Invoice_Posting.Tag = "mnuChargeInvoice"
   frmCP_Charge_Invoice_Posting.Show
End Sub

Private Sub mnuCheckDep_Click()
   frmBranchCheckDeposit.Tag = "mnuCheckDep"
   frmBranchCheckDeposit.Show
End Sub

Private Sub mnuCheckDepReg_Click()
   frmBranchCheckDepositReg.Tag = "mnuCheckDepReg"
   frmBranchCheckDepositReg.Show
End Sub

Private Sub mnuCloseDay2Day_Click()
   frmCloseDay2Day.Tag = "mnuCloseDay2Day"
   frmCloseDay2Day.Show
End Sub

Private Sub mnuCompOff_Click()
   frmCompensationOffApplication.Tag = "mnuCompOff"
   frmCompensationOffApplication.Show
End Sub

Private Sub mnuCompOffApp_Click()
   frmCompensationOffApproval.Tag = "mnuCompOffApp"
   frmCompensationOffApproval.Show
End Sub

Private Sub mnuCompOffReg_Click()
   frmCompensationOffReg.Tag = "mnuCompOffReg"
   frmCompensationOffReg.Show
End Sub

Private Sub mnuCP2MCTransfer_Click()
   frmCPTransfer2MC.Tag = "mnuCP2MCTransfer"
   frmCPTransfer2MC.Show
End Sub

Private Sub mnuCPConsignment_Click()
   frmCpConsignment.Tag = "mnuCPConsignment"
   frmCpConsignment.Show
End Sub

Private Sub mnuCPGiveaway_Click()
   frmCPTransGAways.Tag = "mnuCPGiveaway"
   frmCPTransGAways.Show
End Sub

Private Sub mnuCPInvStockReqReg_Click()
   frmCPInvStockReqReg.Tag = "mnuCPInvStockReqReg"
   frmCPInvStockReqReg.Show
End Sub

Private Sub mnuCPInvUnitReqReg_Click()
   frmCPInvUnitReqReg.Tag = "mnuCPInvUnitReqReg"
   frmCPInvUnitReqReg.Show
End Sub

Private Sub mnuCPJobOrder_Click()
   frmCP_JobOrder.Tag = "mnuCPJobOrder"
   frmCP_JobOrder.Show
End Sub

Private Sub mnuCPJOTransfer_Click()
   frmCP_JO_Branch_Transfer.Tag = "mnuCPJOTransfer"
   frmCP_JO_Branch_Transfer.Show
End Sub

Private Sub mnuCPPriceList_Click()
'   frmCPCashPrice.Show
   frmAppliancesCashPrice.Show
End Sub

Private Sub mnuCPPriceUpdate_Click()
   If oApp.UserLevel >= xeEngineer Then
      frmCPPriceUpdate.Tag = "mnuCPPriceUpdate"
      frmCPPriceUpdate.Show
   End If
End Sub

Private Sub mnuCPSellPriceEntry_Click()
   frmCP_SellPrice_Entry.Tag = "mnuCPSellPriceEntry"
   frmCP_SellPrice_Entry.Show
End Sub

Private Sub mnuCPSRP_Click()
   frmCP_SRP.Tag = "mnuCPSRP"
   frmCP_SRP.Show
End Sub

Private Sub mnuCPStockOrder_Click()
   frmCPInvStockRequest.Tag = "mnuCPStockOrder"
   frmCPInvStockRequest.Show
End Sub

Private Sub mnuCPUnitStockOrder_Click()
   frmCPInvUnitRequest.Tag = "mnuCPUnitStockOrder"
   frmCPInvUnitRequest.Show
End Sub

Private Sub mnuCustomerAccesories_Click()
   frmCustomerAccessories.Tag = "mnuCustomerAccesories"
   frmCustomerAccessories.Show
End Sub

Private Sub mnuDealer_Click()
   frmCP_Dealer.Tag = "mnuDealer"
   frmCP_Dealer.Show
End Sub

'Private Sub mnuDefect_Click()
'   frmCP_Defect.Tag = "mnuDefect"
'   frmCP_Defect.Show
'End Sub

Private Sub mnuDelSched_Click()
   frmCPDelSched.Tag = "mnuDelSched"
   frmCPDelSched.Show
End Sub

Private Sub mnuDocumentAccept_Click()
   frmGeneralTransferPosting.Tag = "mnuDocumentAccept"
   frmGeneralTransferPosting.Show
End Sub

Private Sub mnuDocumentTransfer_Click()
   frmGeneralTransfer.Tag = "mnuDocumentTransfer"
   frmGeneralTransfer.Show
End Sub

Private Sub mnuDTRPosting_Click()
   frmDtrSummary.Tag = "mnuDTRPosting"
   frmDtrSummary.Show
End Sub

Private Sub mnuEloadPosting_Click()
   frmELoadTagging.Tag = "mnuEloadPosting"
   frmELoadTagging.Show
End Sub

Private Sub mnuEloadReg_Click()
   frmEloadReg.Tag = "mnuEloadReg"
   frmEloadReg.Show
End Sub

Private Sub mnuEmployee_Click()
   frmEmployee.Tag = "mnuEmployee"
   frmEmployee.Show
End Sub

Private Sub mnuExport_Click()
'   Dim lo As clsExport
'
'   Set lo = New clsExport
'   Set lo.AppDriver = oApp
'   If lo.Export = False Then
'      MsgBox "Unable to Export Records!!!", vbCritical, "Warning"
'   End If
'   Set lo = Nothing
End Sub

Private Sub mnuFinancer_Click()
   frmCP_Financer.Tag = "mnuFinancer"
   frmCP_Financer.Show
End Sub

Private Sub mnuGenAccountsRep_Click()
   Dim loReports As clsCPAuditRep
   Dim loRepViewer As frmRepViewer

   Set loReports = New clsCPAuditRep
   With loReports
      Set .AppDriver = oApp
      If .ShowReport Then
         Set loRepViewer = New frmRepViewer
         Set loRepViewer.ReportSource = .Source

         loRepViewer.Show
         .CloseReport
      End If
   End With
End Sub

Private Sub mnuGSCMCode_Click()
   frmGSCMCode.Tag = "mnuGSCMCode"
   frmGSCMCode.Show
End Sub

Private Sub mnuICCreditCard_Click()
   frmCPCredCardCalc.Tag = "mnuICCreditCard"
   frmCPCredCardCalc.Show
End Sub

Private Sub mnuICFinancing_Click()
   frmInhouseFinCalculator.Tag = "mnuICFinancing"
   frmInhouseFinCalculator.Show
End Sub

Private Sub mnuImport_Click()
'   Dim lo As clsImport
'
'   Set lo = New clsImport
'   Set lo.AppDriver = oApp
'   If lo.Import = False Then
'      MsgBox "Unable to Export Records!!!", vbCritical, "Warning"
'   End If
'   Set lo = Nothing
End Sub

Private Sub mnuInsCalculator_Click()
'   frmCPInsCalculator.Show
End Sub

Private Sub mnuInvClass_Click()
   Dim loCPClassify As clsCPABCClassify

   Set loCPClassify = New clsCPABCClassify
   Set loCPClassify.AppDriver = oApp
   If Not loCPClassify.InitTransaction Then Exit Sub
   If Not loCPClassify.ClassifyABC Then Exit Sub
   MsgBox "Inventory was Classified Successfully!", vbInformation, "Notice"
End Sub

Private Sub mnuInvClassifyUnit_Click()
   pbProcClassify = True
   Dim loCPClassifyUnit As clsCPUnitClassify
   
   Set loCPClassifyUnit = New clsCPUnitClassify
   Set loCPClassifyUnit.AppDriver = oApp
   If Not loCPClassifyUnit.InitTransaction Then Exit Sub
   If Not loCPClassifyUnit.ClassifyABC Then Exit Sub
   MsgBox "Cellphone Units was Classified Successfully!", vbInformation, "Notice"
   pbProcClassify = False
End Sub

Private Sub mnuInvTypeTransfer_Click()
   frmCPInvTypeTrans.Tag = "mnuInvTypeTransfer"
   frmCPInvTypeTrans.Show
End Sub

Private Sub mnuJobOrderReg_Click()
   frmCP_JobOrderReg.Tag = "mnuJobOrderReg"
   frmCP_JobOrderReg.Show
End Sub

Private Sub mnuJobOrderReports_Click()
   Dim loReports As clsJobOrderRep
   Dim loRepViewer As frmRepViewer

   Set loReports = New clsJobOrderRep
   With loReports
      Set .AppDriver = oApp
      If .ShowReport Then
         Set loRepViewer = New frmRepViewer
         Set loRepViewer.ReportSource = .Source

         loRepViewer.Show
         .CloseReport
      End If
   End With
End Sub

Private Sub mnuJobOrderTransferReg_Click()
   frmCP_JO_Branch_Transfer_Reg.Tag = "mnuJobOrderTransferReg"
   frmCP_JO_Branch_Transfer_Reg.Show
End Sub

Private Sub mnuJOReceivingBranch_Click()
   frmPostCPJODelivery.Tag = "mnuJOReceivingBranch"
   frmPostCPJODelivery.Show
End Sub

Private Sub mnuJOReceivingSrvcCntr_Click()
   frmPostCPJOForwarded.Tag = "mnuJOReceivingSrvcCntr"
   frmPostCPJOForwarded.Show
End Sub

'Private Sub mnuLabor_Click()
'   frmCP_Labor.Tag = "mnuLabor"
'   frmCP_Labor.Show
'End Sub

Private Sub mnuLiquadationReg_Click(Index As Integer)
   frmLiquidationReg.Tag = "mnuLiquadationReg"
   frmLiquidationReg.Show
End Sub

Private Sub mnuLiquidationApprvl_Click()
   frmLiquidationPosting.Tag = "mnuLiquidationApprvl"
   frmLiquidationPosting.Show
End Sub

Private Sub mnuLiquidationEntry_Click()
   frmLiquidationEntry.Tag = "mnuLiquidationEntry"
   frmLiquidationEntry.Show
End Sub

Private Sub mnuLoadAdjPos_Click()
   frmCPLoadAdjReg.Tag = "mnuLoadAdjPos"
   frmCPLoadAdjReg.Show
End Sub

Private Sub mnuLoadAdjustment_Click()
   frmCPLoadAdj.Tag = "mnuLoadAdjustment"
   frmCPLoadAdj.Show
End Sub

Private Sub mnuLoadAdjustmentReg_Click()
   frmCPLoadAdjPosted.Tag = "mnuLoadAdjustmentReg"
   frmCPLoadAdjPosted.Show
End Sub

Private Sub mnuLoadReceiving_Click()
   frmCP_Load_Receiving.Tag = "mnuLoadReceiving"
   frmCP_Load_Receiving.Show
   MsgBox oApp.MenuName
End Sub

Private Sub mnuLoadReceivingReg_Click()
   frmCP_Load_Receiving_Reg.Tag = "mnuLoadReceivingReg"
   frmCP_Load_Receiving_Reg.Show
End Sub

Private Sub mnuLoadSplitAdj_Click()
   If oApp.BranchCode = "C001" Or oApp.BranchCode = "C0A9" Then
      frmSplitLoadPosting.Tag = "mnuPostLoadTransfer"
      frmSplitLoadPosting.Show
   End If
End Sub

Private Sub mnuLoadTransfer_Click()
   frmCP_Load_Transfer.Tag = "mnuLoadTransfer"
   frmCP_Load_Transfer.Show
End Sub

Private Sub mnuLoadTransferReg_Click()
   frmCP_Load_Transfer_Reg.Tag = "mnuLoadTransferReg"
   frmCP_Load_Transfer_Reg.Show
End Sub

Private Sub mnuLoadWalletReg_Click()
   frmLoad_WalletReg.Tag = "mnuLoadWalletReg"
   frmLoad_WalletReg.Show
End Sub

Private Sub mnuLRActive_Click()
   Dim oFormMCActRecMP As frmMPActRecMP

   Set oFormMCActRecMP = New frmMPActRecMP
   Set oFormMCActRecMP.FormMCActRec = oFormMCActRecMP
   
   oFormMCActRecMP.TranStatus = xeActStatActive
   oFormMCActRecMP.Caption = "Accounts Receivable(Active)"
   oFormMCActRecMP.Tag = "mnuLRActive"
   oFormMCActRecMP.Show
End Sub

Private Sub mnuLRInactive_Click()
   Dim oFormMCActRecMP As frmMPActRecMP

   Set oFormMCActRecMP = New frmMPActRecMP
   Set oFormMCActRecMP.FormMCActRec = oFormMCActRecMP
   
   oFormMCActRecMP.TranStatus = xeActStatClosed
   oFormMCActRecMP.Caption = "Accounts Receivable(Inactive)"
   oFormMCActRecMP.Tag = "mnuLRInactive"
   oFormMCActRecMP.Show
End Sub

Private Sub mnuManagerRep_Click()
   Dim loReports As clsManagerRep
   Dim loRepViewer As frmRepViewer

   Set loReports = New clsManagerRep
   With loReports
      Set .AppDriver = oApp
      If .ShowReport Then
         Set loRepViewer = New frmRepViewer
         Set loRepViewer.ReportSource = .Source

         loRepViewer.Show
         .CloseReport
      End If
   End With
End Sub

Private Sub mnuMarketingSupport_Click()
   frmMarketingSupport.Tag = "mnuMarketingSupport"
   frmMarketingSupport.Show
End Sub

Private Sub mnuMCSOverride_Click()
   frmCP_MCS_Override.Tag = "mnuMCSOverride"
   frmCP_MCS_Override.Show
End Sub


Private Sub mnuNPRate_Click()
   frmMPPromoCat.Tag = "mnuRateMP_Click"
   frmMPPromoCat.Show
End Sub

Private Sub mnuPAAdvances_Click()
   frmEmployeeAdvances.Tag = "frmEmployeeAdvances"
   frmEmployeeAdvances.Show
End Sub

Private Sub mnuPABusinessTrip_Click()
   frmOBApplication.Tag = "mnuPABusinessTrip"
   frmOBApplication.Show
End Sub

Private Sub mnuPackageModel_Click()
   frmPackageModel.Tag = "mnuPackageModel"
   frmPackageModel.Show
End Sub

Private Sub mnuPALeave_Click()
   frmLeaveApplication.Tag = "mnuPALeave"
   frmLeaveApplication.Show
End Sub

Private Sub mnuPALoan_Click()
   frmEmployeeLoans.Tag = "mnuPALoan"
   frmEmployeeLoans.Show
End Sub

Private Sub mnuPAOBTripWLog_Click()
   frmOBWithLogApp.Tag = "mnuPAOBTripWLog"
   frmOBWithLogApp.Show
End Sub

Private Sub mnuPAOvertime_Click()
   frmOTApplication.Tag = "mnuPAOvertime"
   frmOTApplication.Show
End Sub

Private Sub mnuPATardiness_Click()
   frmTardiness.Tag = "mnuPATardiness"
   frmTardiness.Show
End Sub

Private Sub mnuPAUndertime_Click()
   frmUndertime.Tag = "mnuPAUndertime"
   frmUndertime.Show
End Sub

Private Sub mnuPCCashAdvancApprvl_Click()
   'frmCashAdvanceApproval.Tag = "mnuPCCashAdvancApprvl"
   'frmCashAdvanceApproval.Show
End Sub

Private Sub mnuPCCashAdvance_Click()
   'frmCashAdvanceEntry.Tag = "mnuPCCashAdvance"
   'frmCashAdvanceEntry.Show
End Sub

Private Sub mnuPettyCash_Click()
   frmPettyCash.Tag = "mnuPettyCash"
   frmPettyCash.Show
End Sub

Private Sub mnuPOSReg_Click()
   frmCP_POSReg.Tag = "mnuPOSReg"
   frmCP_POSReg.Show
End Sub

Private Sub mnuPostLoadTransfer_Click()
   frmPostLoadTransfer.Tag = "mnuPostLoadTransfer"
   frmPostLoadTransfer.Show
End Sub

Private Sub mnuPRDayoffShifting_Click()
   frmDayOffShftApplication.Tag = "mnuPRDayoffShifting"
   frmDayOffShftApplication.Show
End Sub

Private Sub mnuPriceProtection_Click()
   frmCP_Price_Protection.Tag = "mnuPriceProtection"
   frmCP_Price_Protection.Show
End Sub

Private Sub mnuPrintBarcodeLX310_Click()
'   frmBarrCodeLX310.Tag = "mnuPrintBarcodeLX31"
'   frmBarrCodeLX310.Show
End Sub

Private Sub mnuPRMovement_Click()
   frmEmployeeMovement.Tag = "mnuPRMovement"
   frmEmployeeMovement.Show
End Sub

Private Sub mnuProductInquiry_Click()
   frmMPProductInquiry.Tag = "mnuProductInquiry"
   frmMPProductInquiry.Show
End Sub

Private Sub mnuPRShiftMovement_Click()
   frmShiftSchedApplication.Tag = "mnuPRShiftMovement"
   frmShiftSchedApplication.Show
End Sub

Private Sub mnuPRSuspension_Click()
   frmSuspensionApplication.Tag = "mnuPRSuspension"
   frmSuspensionApplication.Show
End Sub

Private Sub mnuPTForgot_Click()
   frmForgot2Swipe.Tag = "mnuPTForgot"
   frmForgot2Swipe.Show
End Sub

Private Sub mnuPTManualLog_Click()
'   frmLogManual2.Tag = "mnuPTManualLog"
'   frmLogManual2.ByBranch = True
'   frmLogManual2.Show
   frmLogManualWR.Tag = "mnuPTManualLog"
   frmLogManualWR.ByBranch = True
   frmLogManualWR.Show
End Sub

Private Sub mnuPTProcessLog_Click()
   frmLogProcess.Tag = "mnuPTProcessLog"
   frmLogProcess.Show
End Sub

Private Sub mnuPUExport_Click()
   Dim loCls As clsLogCapture
   Set loCls = New clsLogCapture
   Set loCls.AppDriver = oApp
   
   If loCls.Export Then
      MsgBox "Timesheet exported successfully!"
   Else
      MsgBox "Unable to export timesheet!"
   End If
End Sub

Private Sub mnuPurchaseReplacement_Click()
   frmCP_PO_Replacement.Tag = "mnuPurchaseReplacement"
   frmCP_PO_Replacement.Show
End Sub

Private Sub mnuPurchaseReplacementReg_Click()
   frmCP_PO_ReplacementReg.Tag = "mnuPurchaseReplacement"
   frmCP_PO_ReplacementReg.Show
End Sub

Private Sub mnuPVAdvances_Click()
   frmEmployeeAdvancesApprvl.Tag = "mnuPVAdvances"
   frmEmployeeAdvancesApprvl.Show
End Sub

Private Sub mnuPVBusinessTrip_Click()
   frmOBApproval.Tag = "mnuPVBusinessTrip"
   frmOBApproval.Show
End Sub

Private Sub mnuPVDayoffShifting_Click()
   frmDayOffShftApproval.Tag = "mnuPVDayoffShifting"
   frmDayOffShftApproval.Show
End Sub

Private Sub mnuPVForgot2Swipe_Click()
   frmForgot2SwipeApprvl.Tag = "mnuPVForgot2Swipe"
   frmForgot2SwipeApprvl.Show
End Sub

Private Sub mnuPVLeave_Click()
   frmLeaveApproval.Tag = "mnuPVLeave"
   frmLeaveApproval.Show
End Sub

Private Sub mnuPVManualLog_Click()
'   frmLogManualApprvl2.Tag = "mnuPVManualLog"
'   frmLogManualApprvl2.Show
   frmLogManualApprvlWR.Tag = "mnuPVManualLog"
   frmLogManualApprvlWR.Show
End Sub

Private Sub mnuPVOBTripWLog_Click()
   frmOBWithLogApprvl.Tag = "mnuPVOBTripWLog"
   frmOBWithLogApprvl.Show
End Sub

Private Sub mnuPVOvertime_Click()
   frmOTApproval.Tag = "mnuPVOvertime"
   frmOTApproval.Show
End Sub

Private Sub mnuPVShiftMovement_Click()
   frmShiftSchedApproval.Tag = "mnuPVShiftMovement"
   frmShiftSchedApproval.Show
End Sub

Private Sub mnuPVTardiness_Click()
   frmTardinessApproval.Tag = "mnuPVTardiness"
   frmTardinessApproval.Show
End Sub

Private Sub mnuPVUndertime_Click()
   frmUndertimeApproval.Tag = "mnuPVUndertime"
   frmUndertimeApproval.Show
End Sub

Private Sub mnuRaffle_Click()
Dim lnResult As Long

        If (Dir(pxeJavaPath & "raffle.bat") <> "") Then
                    lnResult = (RMJExecute(pxeJavaPath & "raffle.bat" & "argrument"))
                    If (lnResult = 0) Then
                       MsgBox "Raffle successfully retreive/create !" & vbCrLf & vbCrLf & _
                   "Thank you ", vbInformation, "Notice"
                        End If
                    
                    
                    If (lnResult = 1) Then
                        MsgBox "Image Does'nt Exist!! Please Inform MIS Department for uploading image!!", vbInformation, "Notice"
                    End If
                 Else 'path check
                     MsgBox "File Path Does'nt Exist  " & pxeJavaPath & "raffle.bat" & "   Please Inform MIS Dept !!", vbInformation, "Notice"
                End If
End Sub

Private Sub mnuRaffleEntryScanner_Click()
    Dim lsArguments As String
        Dim lnResult As Long
            If (Dir(pxeJavaPath & "readpanalo.bat") <> "") Then
            lsArguments = oApp.ProductID & " " & oApp.UserID
                lnResult = (RMJExecute(pxeJavaPath & "readpanalo.bat " & lsArguments))
                    If (lnResult = 0) Then
                        MsgBox "Raffle entry created successfully." _
                        & vbCrLf & vbCrLf & "The customer should expect a notification regarding his raffle coupons on his Guanzon Connect within the day. Thank you." _
                        , vbInformation, "Notice"
                    End If
                    If (lnResult = 1) Then
                        MsgBox "Unable to Retrieve Information. Please Inform MIS !! ", vbInformation, "Notice"
                    End If
            Else 'path check
                 MsgBox "File Path Does'nt Exist  " & pxeJavaPath & "readpanalo.bat" & "   Please Inform MIS Dept !!", vbInformation, "Notice"
            End If
End Sub

Private Sub mnuRateExtreme_Click()
   frmExtremePromoCat.Tag = "mnuRateExtreme_Click"
   frmExtremePromoCat.Show
End Sub

Private Sub mnuReceipt_Click()
   frmCashierTrans.Tag = "mnuReceipt"
   frmCashierTrans.Show
End Sub

Private Sub mnuReceiptReg_Click()
   frmCashierTransReg.Tag = "mnuReceiptReg"
   frmCashierTransReg.Show
End Sub

Private Sub mnuReceiveServicePhone_Click()
   frmServicePhonePosting.Tag = "mnuReceiveServicePhone"
   frmServicePhonePosting.Show
End Sub

Private Sub mnuRefundableApp_Click()
   frmRefundableDepConfirm.Tag = "mnuRefundableApp"
   frmRefundableDepConfirm.Show
End Sub

Private Sub mnuRefundableEntry_Click()
   frmRefundableDep.Tag = "mnuRefundableEntry"
   frmRefundableDep.Show
End Sub

Private Sub mnuRefundableHist_Click()
   frmRefundableHist.Tag = "mnuRefundableHist"
   frmRefundableHist.Show
End Sub

Private Sub mnuRegPurchaseReturn_Click()
   frmCP_PO_Return_Reg.Tag = "mnuRegPurchaseReturn"
   frmCP_PO_Return_Reg.Show
End Sub

Private Sub mnuAccessories_Click()
   frmAccessories.Tag = "mnuAccessories"
   frmAccessories.Show
End Sub

Private Sub mnuBranch_Click()
   frmBranch.Tag = "mnuBranch"
   frmBranch.Show
End Sub

Private Sub mnuBranchReceived_Click()
   frmCP_BranchReceived.Tag = "mnuBranchReceived"
   frmCP_BranchReceived.Show
End Sub

Private Sub mnuBrand_Click()
   frmBrand.Tag = "mnuBrand"
   frmBrand.Show
End Sub

Private Sub mnuCard_Click()
   frmCreditCard.Tag = "mnuCard"
   frmCreditCard.Show
End Sub

Private Sub mnuCategory_Click()
   If oApp.UserLevel > xeManager Then
      frmCategory.Tag = "mnuCategory"
      frmCategory.Show
   End If
End Sub

Private Sub mnuChargeInvoice_Click()
   frmCP_Charge_Invoice.Tag = "mnuChargeInvoice"
   frmCP_Charge_Invoice.Show
End Sub

Private Sub mnuClientMaster_Click()
   frmClientInfo.Tag = "mnuClientMaster"
   frmClientInfo.Show
End Sub

Private Sub mnuColor_Click()
   frmColor.Tag = "mnuColor"
   frmColor.Show
End Sub

Private Sub mnuCPInventory_Click()
   frmCP_Inventory.Tag = "mnuCPInventory"
   frmCP_Inventory.Show
End Sub

Private Sub mnuCPInventoryBranch_Click()
   frmCP_Inventory_Branch.Tag = "mnuCPInventoryBranch"
   frmCP_Inventory_Branch.Show
End Sub

Private Sub mnuCPSerial_Click()
   frmCP_Serial.Tag = "mnuCPSerial"
   frmCP_Serial.Show
End Sub

Private Sub mnuCpSerialStatus_Click()
   frmCP_Serial_Status.Tag = "mnuCpSerialStatus"
   frmCP_Serial_Status.Show
End Sub

Private Sub mnuDummySerial_Click()
   frmDASerial.Tag = "mnuDummySerial"
   frmDASerial.Show
End Sub

Private Sub mnuEloadMatrix_Click()
   frmEload_Matrix.Tag = "mnuEloadMatrix"
   frmEload_Matrix.Show
End Sub

Private Sub mnuEloadMatrixInventory_Click()
   frmCP_Load_Matrix.Tag = "mnuEloadMatrixInventory"
   frmCP_Load_Matrix.Show
End Sub

Private Sub mnuExit_Click()
   Unload Me
End Sub

Private Sub mnuInvAdjPosting_Click()
   frmCPInvAdjReg.Tag = "mnuInvAdjPosting"
   frmCPInvAdjReg.Show
End Sub

Private Sub mnuInventoryAdjustment_Click()
   frmCPInvAdj.Tag = "mnuInventoryAdjustment"
   frmCPInvAdj.Show
End Sub

Private Sub mnuModel_Click()
   frmCP_Model.Tag = "mnuModel"
   frmCP_Model.Show
End Sub

Private Sub mnuPOS_Click()
   frmCP_POSOld.Tag = "mnuPOS"
   frmCP_POSOld.Show
End Sub

Private Sub mnuPrintBarcode_Click()
   frmBarrCode.Tag = "mnuPrintBarcode"
   frmBarrCode.Show
End Sub

Private Sub mnuPurchaseOrder_Click()
   frmCP_Purchasing.Tag = "mnuPurchaseOrder"
   frmCP_Purchasing.Show
End Sub

Private Sub mnuPurchaseReceiving_Click()
   frmCP_PO_Receiving.Tag = "mnuPurchaseReceiving"
   frmCP_PO_Receiving.Show
End Sub

Private Sub mnuPurchaseReturn_Click()
   frmCP_PO_Return.Tag = "mnuPurchaseReturn"
   frmCP_PO_Return.Show
End Sub

Private Sub mnuReceiveTransfer_Click()
   frmPostCPDelivery.Tag = "mnuReceiveTransfer"
   frmPostCPDelivery.Show
End Sub

Private Sub mnuRegInvAdjustment_Click()
   frmCPInvAdjPosted.Tag = "mnuRegInvAdjustment"
   frmCPInvAdjPosted.Show
End Sub

Private Sub mnuRegPurchaseOrder_Click()
   frmCP_Purchasing_Post.Tag = "mnuRegPurchaseOrder"
   frmCP_Purchasing_Post.Show
End Sub

Private Sub mnuRegPurchaseReceiving_Click()
   frmCP_PO_Receiving_Reg.Tag = "mnuRegPurchaseReceiving"
   frmCP_PO_Receiving_Reg.Show
End Sub

Private Sub mnuRegServiceCenter_Click()
   frmCP_JobOrderReg.Tag = "mnuRegServiceCenter"
   frmCP_JobOrderReg.Show
End Sub

Private Sub mnuRegStockIssue_Click()
   frmCP_Branch_Transfer_Reg.Tag = "mnuRegStockIssue"
   frmCP_Branch_Transfer_Reg.Show
End Sub

Private Sub mnuReplenishment_Click()
   frmReplenishment.Tag = "mnuReplenishment"
   frmReplenishment.Show
End Sub

Private Sub mnuReplenishmentApprvl_Click()
   frmReplenishmentApprvl.Tag = "mnuReplenishmentApprvl"
   frmReplenishmentApprvl.Show
End Sub

Private Sub mnuReplinishmentReg_Click(Index As Integer)
   frmReplenishmentReg.Tag = "mnuReplinishmentReg"
   frmReplenishmentReg.Show
End Sub

Private Sub mnuRequestExport_Click()
'   Dim lo As clsExportRequest
'
'   Set lo = New clsExportRequest
'   Set lo.AppDriver = oApp
'   If lo.ExportRequest = False Then
'      MsgBox "Unable to Request Export!!!", vbCritical, "Warning"
'   End If
'   Set lo = Nothing
End Sub

Private Sub mnuRPAAdvances_Click()
   frmEmployeeAdvancesReg.Tag = "mnuRPAAdvances"
   frmEmployeeAdvancesReg.Show
End Sub

Private Sub mnuRPABusinessTrip_Click()
   frmOBReg.Tag = "mnuRPABusinessTrip"
   frmOBReg.Show
End Sub

Private Sub mnuRPALeave_Click()
   frmLeaveReg.Tag = "mnuRPALeave"
   frmLeaveReg.Show
End Sub

Private Sub mnuRPALoan_Click()
   frmEmployeeLoansReg.Tag = "mnuRPALoan"
   frmEmployeeLoansReg.Show
End Sub

Private Sub mnuRPAOBTripWLog_Click()
   frmOBWithLogReg.Tag = "mnuRPAOBTripWLog"
   frmOBWithLogReg.Show
End Sub

Private Sub mnuRPAOvertime_Click()
   frmOTReg.Tag = "mnuRPAOvertime"
   frmOTReg.Show
End Sub

Private Sub mnuRPATardiness_Click()
   frmTardinessReg.Tag = "mnuRPATardiness"
   frmTardinessReg.Show
End Sub

Private Sub mnuRPRDayoffShifting_Click()
   frmDayOffShftReg.Tag = "mnuRPRDayoffShifting"
   frmDayOffShftReg.Show
End Sub

Private Sub mnuRPRMovement_Click()
   frmEmployeeMovementReg.Tag = "mnuRPRMovement"
   frmEmployeeMovementReg.Show
End Sub

Private Sub mnuRPRShiftMovement_Click()
   frmShiftSchedReg.Tag = ""
   frmShiftSchedReg.Show
End Sub

Private Sub mnuRPTForgot_Click()
   frmForgot2SwipeReg.Tag = "mnuRPTForgot"
   frmForgot2SwipeReg.Show
End Sub

Private Sub mnuRPTManualLog_Click()
'   frmLogManualReg.Tag = "mnuRPTManualLog"
'   frmLogManualReg.Show
   frmLogManualRegWR.Tag = "mnuRPTManualLog"
   frmLogManualRegWR.Show
End Sub

Private Sub mnuSales_Return_Reg_Click()
   frmCP_Sales_Return_Reg.Tag = "mnuSales_Return_Reg"
   frmCP_Sales_Return_Reg.Show
End Sub

Private Sub mnuSalesByDate_Click()
   frmSalesTagging.Tag = "mnuSalesByDate"
   frmSalesTagging.Show
End Sub

Private Sub mnuSalesman_Click()
   frmSalesman.Tag = "mnuSalesman"
   frmSalesman.Show
End Sub

Private Sub mnuSalesReturn_Click()
   frmCP_Sales_Return.Tag = "mnuSalesReturn"
   frmCP_Sales_Return.Show
End Sub

Private Sub mnuServiceCenter_Click()
   frmCP_Service_Center.Tag = "mnuServiceCenter"
   frmCP_Service_Center.Show
End Sub

Private Sub mnuServicePhoneTransfer_Click()
   frmServicePhoneTagging.Tag = "mnuServicePhoneTransfer"
   frmServicePhoneTagging.Show
End Sub

Private Sub mnuServicePhoneTransferReg_Click()
   frmServicePhoneHistory.Tag = "mnuServicePhoneTransferReg"
   frmServicePhoneHistory.Show
End Sub

Private Sub mnuSetGiveaways_Click()
   frmSalesGiveaways.Tag = "mnuSetGiveaways"
   frmSalesGiveaways.Show
End Sub

Private Sub mnuSetPackage_Click()
   frmSalesPackage.Tag = "mnuSetPackage"
   frmSalesPackage.Show
End Sub

Private Sub mnuSize_Click()
   frmSize.Tag = "mnuSize"
   frmSize.Show
End Sub

Private Sub mnuSplitLoadTransfer_Click()
   frmSplitLoad.Tag = "mnuSplitLoadTransfer"
   frmSplitLoad.Show
End Sub

Private Sub mnuStandardReports_Click()
   Dim loReports As clsCPBranchRep
   Dim loRepViewer As frmRepViewer

   Set loReports = New clsCPBranchRep
   With loReports
      Set .AppDriver = oApp
      If .ShowReport Then
         Set loRepViewer = New frmRepViewer
         Set loRepViewer.ReportSource = .Source

         loRepViewer.Show
         .CloseReport
      End If
   End With
End Sub

Private Sub mnuStockInquiry_Click()
   frmCP_Inquiry.Tag = "mnuStockInquiry"
   frmCP_Inquiry.Show
End Sub

Private Sub mnuStockTransfer_Click()
   frmCPDeliveryOthers.Tag = "mnuStockTransfer"
   frmCPDeliveryOthers.Show
End Sub

Private Sub mnuSupplier_Click()
   frmCP_Supplier.Tag = "mnuSupplier"
   frmCP_Supplier.Show
End Sub

Private Sub mnuSupplies_Click()
   frmSupplies.Tag = "mnuSupplies"
   frmSupplies.Show
End Sub

Private Sub mnuSuppliesPosting_Click()
   frmSuppliesTransferPosting.Tag = "mnuSuppliesPosting"
   frmSuppliesTransferPosting.Show
End Sub

Private Sub mnuSuppliesRequest_Click()
   frmSuppliesRequest.Tag = "mnuSuppliesRequest"
   frmSuppliesRequest.Show
End Sub

Private Sub mnuSuppliesTransfer_Click()
   frmSuppliesTransfer.Tag = "mnuSuppliesTransfer"
   frmSuppliesTransfer.Show
End Sub

Private Sub mnuSuppliesTransferPost_Click()
   frmSuppliesTransferPosting.Tag = "mnuSuppliesTransferPost"
   frmSuppliesTransferPosting.Show
End Sub

'Private Sub mnuSymptom_Click()
'   frmCP_Symptom.Tag = "mnuSymptom"
'   frmCP_Symptom.Show
'End Sub

'Private Sub mnuTechnician_Click()
'   frmTechnician.Tag = "mnuTechnician"
'   frmTechnician.Show
'End Sub

Private Sub mnuTradeIn_Click()
'   frmTradeInTransfer.Tag = "mnuTradeIn"
'   frmTradeInTransfer.Show
End Sub

Private Sub mnuTradeInTransfer_Click()
   frmTradeInTransfer.Tag = "mnuTradeInTransfer"
   frmTradeInTransfer.Show
End Sub

Private Sub mnuUActiveAccounts_Click()
   frmMPCustomerLedger.Tag = "mnuUActiveAccounts"
   frmMPCustomerLedger.Show
End Sub

Private Sub mnuUInactiveAccounts_Click()
   frmMPCustomerLedgerClosed.Tag = "mnuUInactiveAccounts"
   frmMPCustomerLedgerClosed.Show
End Sub

Private Sub mnuUnencodedTrans_Click()
   frmUnpostedTransaction.Tag = "mnuUnencodedTrans"
   frmUnpostedTransaction.Show
End Sub

Private Sub mnuUnitClusterDel_Click()
   frmCPClusterDelivery.Tag = "mnuUnitClusterDel"
   frmCPClusterDelivery.Show
End Sub

Private Sub mnuUnitClusterDelHist_Click()
   frmCPClusterDeliveryReg.Tag = "mnuUnitClusterDelHist"
   frmCPClusterDeliveryReg.Show
End Sub

Private Sub mnuUnitReqApp_Click()
   frmCP_Request_Approval.Tag = "mnuUnitReqApp"
   frmCP_Request_Approval.Show
End Sub

Private Sub mnuUnitTransfer_Click()
   frmCPDelivery.Tag = "mnuUnitTransfer"
   frmCPDelivery.Show
End Sub

Private Sub mnuWAccessoriesReg_Click()
   frmCP_AccessJobOrderReg.Tag = "mnuWAccessoriesReg"
   frmCP_AccessJobOrderReg.Show
End Sub

Private Function getLastPeriod(ByVal fsEmployID As String) As Date
   Dim lsSQL As String
   Dim lors As Recordset
   
   lsSQL = "SELECT" & _
                  " a.dCovergTo" & _
          " FROM Payroll_Period a" & _
              " LEFT JOIN Payroll_Summary b ON a.sPayPerID = b.sPayPerID" & _
          " WHERE b.sEmployID = " & strParm(fsEmployID) & _
          " ORDER BY a.dCovergTo DESC LIMIT 1"
   Set lors = oApp.Connection.Execute(lsSQL, , adCmdText)
   
   If lors.EOF Then
      getLastPeriod = Format(oApp.ServerDate, "yyyy-mm-dd")
   Else
      getLastPeriod = lors("dCovergTo") + 1
   End If
End Function

Private Sub quickTransfer(ByVal fsTable As String, ByVal fsFilter As String, ByVal fsBranchCD As String)
'   Dim lors As Recordset
'   Dim lsSQL As String
'
'   Set lors = GetRecordSet(oApp.Connection, fsTable, fsFilter)
'   Do Until lors.EOF
'      lsSQL = ADO2SQL(lors, fsTable)
'      lsSQL = Replace(lsSQL, "INSERT INTO", "REPLACE INTO")
'
'      Call send2Log( _
'         oApp.Connection, _
'         oApp.BranchCode, _
'         oApp.BranchCode, _
'         lsSQL, _
'         fsTable, _
'         fsBranchCD, _
'         oApp.UserID, _
'         oApp.ServerDate, _
'         True)
'
'      lors.MoveNext
'   Loop
End Sub

Private Sub mnuWholeSale_Click()
   frmCPWholeSale.Tag = "mnuWholeSale"
   frmCPWholeSale.Show
End Sub

Private Sub mnuWholeSaleReturn_Click()
   frmCPWholeSaleReturn.Tag = "mnuWholeSaleReturn"
   frmCPWholeSaleReturn.Show
End Sub

Private Sub mnuYearEndBunos_Click()
   frmEmp13thMonth.Tag = "mnuYearEndBunos"
   frmEmp13thMonth.Show
End Sub

Private Sub tmeLog_Timer()
   Dim lsSQL As String
   Dim lors As Recordset
   Dim loCls As clsEmployeeMovement
   Dim ldDateFrom As Date
   
   If pbProcClassify = False Then
      
      DoEvents
      lsSQL = "SELECT" & _
                     "  sTransNox" & _
                     ", sEmployID" & _
                     ", sBranchCD" & _
                     ", xBranchCD" & _
                     ", dEffectve" & _
             " FROM Employee_Movement" & _
             " WHERE cTranStat = " & strParm(xeStateClosed) & _
               " AND dEffectve < " & dateParm(oApp.ServerDate) & _
             " ORDER BY dEffectve DESC"
         
      If InStr("C001»C0A1»C0CW»C0W1", oApp.BranchCode) > 0 Or oApp.UserLevel = xeAudit Then
         Exit Sub
      End If
      
      Set lors = oApp.Connection.Execute(lsSQL, , adCmdText)
      DoEvents
      
      If lors.EOF Then Exit Sub
      
      Set loCls = New clsEmployeeMovement
      'Set loCls.AppDriver = oApp
      'loCls.HasParent = True
      loCls.InitTransaction

      DoEvents
      Do Until lors.EOF
         DoEvents
         If LCase(oApp.ProductID) = "petmgr" Then
            'Its from the main office so send updates to all branches...
            If loCls.OpenTransaction(lors("sTransNox")) Then
               loCls.PostTransaction (lors("sTransNox"))
            End If
         Else
            'if monitor is not from main office then just post the movement
            lsSQL = "UPDATE Employee_Movement" & _
                   " SET cTranStat = " & strParm(xeStatePosted) & _
                   " WHERE sTransNox = " & strParm(lors("sTransNox"))
            oApp.Connection.Execute lsSQL, , adCmdText
         End If

          'From this Branch employee is assigned to other branch
         If IFNull(loCls.Master("sBranchCD")) <> "" _
            And loCls.Master("sBranchCD") <> oApp.BranchCode _
            And IFNull(loCls.Master("xBranchCD"), "") = oApp.BranchCode _
            And InStr(1, "M001»M0W1", lors("sBranchCD")) = 0 Then

            DoEvents
            ldDateFrom = getLastPeriod(loCls.Master("sEmployID"))
            Call quickTransfer("Employee_Log", _
                                "sEmployID = " & strParm(loCls.Master("sEmployID")) & _
                           " AND dTransact BETWEEN " & dateParm(ldDateFrom) & " AND " & dateParm(loCls.Master("dEffectve")), _
                                loCls.Master("sBranchCD"))
            DoEvents

            Call quickTransfer("Employee_Timesheet", _
                                "sEmployID = " & strParm(loCls.Master("sEmployID")) & _
                           " AND dTransact BETWEEN " & dateParm(ldDateFrom) & " AND " & dateParm(loCls.Master("dEffectve")), _
                                loCls.Master("sBranchCD"))
            DoEvents

            Call quickTransfer("Employee_Leave", _
                                "sEmployID = " & strParm(loCls.Master("sEmployID")) & _
                           " AND dApproved BETWEEN " & dateParm(ldDateFrom) & " AND " & dateParm(loCls.Master("dEffectve")), _
                                loCls.Master("sBranchCD"))
            DoEvents

            Call quickTransfer("Employee_Business_Trip", _
                                "sEmployID = " & strParm(loCls.Master("sEmployID")) & _
                           " AND dApproved BETWEEN " & dateParm(ldDateFrom) & " AND " & dateParm(loCls.Master("dEffectve")), _
                                loCls.Master("sBranchCD"))
         End If

         lors.MoveNext
         DoEvents
      Loop
   End If
End Sub
