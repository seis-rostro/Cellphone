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
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuLoanReceivable 
         Caption         =   "Loan Receivable"
         Begin VB.Menu mnuLRActive 
            Caption         =   "Active Account"
         End
         Begin VB.Menu mnuLRInactive 
            Caption         =   "Inactive Account"
         End
         Begin VB.Menu mnuTradeIn 
            Caption         =   "CP Trade IN"
         End
      End
      Begin VB.Menu mnuCPInventory 
         Caption         =   "CP Inventory"
      End
      Begin VB.Menu mnuCPPriceList 
         Caption         =   "CP Price List"
      End
      Begin VB.Menu mnuEloadMatrixInventory 
         Caption         =   "Eload Matrix Inventory"
      End
      Begin VB.Menu mnuCpSerialStatus 
         Caption         =   "CP Serial Status"
      End
      Begin VB.Menu mnuClientMaster 
         Caption         =   "Client Master"
      End
      Begin VB.Menu mnuEloadMatrix 
         Caption         =   "Eload Matrix"
      End
      Begin VB.Menu mnuInsCalculator 
         Caption         =   "Installment Calculator"
         Begin VB.Menu mnuICCreditCard 
            Caption         =   "Credit Card"
         End
         Begin VB.Menu mnuICFinancing 
            Caption         =   "Financing"
         End
      End
      Begin VB.Menu mnuAssetsMaintenance 
         Caption         =   "Assets Maintenance"
      End
      Begin VB.Menu mnuSupplies 
         Caption         =   "Supplies Maintenance"
      End
      Begin VB.Menu mnuFinancer 
         Caption         =   "Financer Maintenance"
      End
      Begin VB.Menu mnuPettyCash 
         Caption         =   "Petty Cash"
      End
      Begin VB.Menu mnuParameters 
         Caption         =   "Parameters"
         Begin VB.Menu mnuBrand 
            Caption         =   "Brand"
         End
         Begin VB.Menu mnuColor 
            Caption         =   "Color"
         End
         Begin VB.Menu mnuAccessories 
            Caption         =   "Accessories"
         End
         Begin VB.Menu mnuSize 
            Caption         =   "Size"
         End
      End
      Begin VB.Menu mnuOthers 
         Caption         =   "Others"
         Begin VB.Menu mnuEmployee 
            Caption         =   "Employee"
         End
         Begin VB.Menu mnuSupplier 
            Caption         =   "Supplier"
         End
         Begin VB.Menu mnuServiceCenter 
            Caption         =   "Service Center"
         End
         Begin VB.Menu mnuModel 
            Caption         =   "Model"
         End
         Begin VB.Menu mnuCard 
            Caption         =   "Card"
         End
         Begin VB.Menu mnuCardRate 
            Caption         =   "CP Card Rate"
         End
         Begin VB.Menu mnuCardRateModel 
            Caption         =   "CP Card Rate Model"
         End
         Begin VB.Menu mnuNPRate 
            Caption         =   "Mobile Phone NortPoint Rate"
         End
         Begin VB.Menu mnuRateExtreme 
            Caption         =   "Extreme Appliances NorthPoint  Rate"
         End
         Begin VB.Menu mnuCategory 
            Caption         =   "Category"
         End
         Begin VB.Menu mnuSalesman 
            Caption         =   "Salesman"
         End
         Begin VB.Menu mnuDealer 
            Caption         =   "Dealer"
         End
         Begin VB.Menu mnuSetGiveaways 
            Caption         =   "Set Giveaways"
         End
         Begin VB.Menu mnuSetPackage 
            Caption         =   "Set Package"
         End
         Begin VB.Menu mnuAssets 
            Caption         =   "Other Assets"
         End
         Begin VB.Menu mnuAssetsBrand 
            Caption         =   "Assets Brand"
         End
      End
      Begin VB.Menu mnuRaffle 
         Caption         =   "Raffle Entry"
      End
      Begin VB.Menu mnuLogOff 
         Caption         =   "Log Off"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuTransaction 
      Caption         =   "&Transaction"
      Begin VB.Menu mnuPurchasing 
         Caption         =   "Purchasing"
         Begin VB.Menu mnuPurchaseReceiving 
            Caption         =   "Purchase Receiving"
         End
         Begin VB.Menu mnuPurchaseReturn 
            Caption         =   "Purchase Return"
         End
         Begin VB.Menu mnuPurchaseOrder 
            Caption         =   "Purchase Order"
         End
         Begin VB.Menu mnuPurchaseReplacement 
            Caption         =   "Purchase Replacement"
         End
         Begin VB.Menu mnuLoadReceiving 
            Caption         =   "Load Receiving"
         End
         Begin VB.Menu mnuCPConsignment 
            Caption         =   "CP Consignment"
         End
      End
      Begin VB.Menu mnuTransfer 
         Caption         =   "Transfer"
         Begin VB.Menu mnuStockTransfer 
            Caption         =   "Stock Transfer"
         End
         Begin VB.Menu mnuUnitTransfer 
            Caption         =   "Unit Transfer"
         End
         Begin VB.Menu mnuLoadTransfer 
            Caption         =   "Load Transfer"
         End
         Begin VB.Menu mnuSplitLoadTransfer 
            Caption         =   "Load Transfer New"
         End
         Begin VB.Menu mnuTradeInTransfer 
            Caption         =   "Trade In Transfer"
         End
         Begin VB.Menu mnuServicePhoneTransfer 
            Caption         =   "Service Phone Transfer"
         End
         Begin VB.Menu mnuCP2MCTransfer 
            Caption         =   "Stock Transfer to MC"
         End
      End
      Begin VB.Menu mnuCPClustering 
         Caption         =   "CP Clustering"
         Begin VB.Menu mnuDelSched 
            Caption         =   "Delivery Schedule"
         End
         Begin VB.Menu mnuUnitReqApp 
            Caption         =   "Unit Request Approval"
         End
         Begin VB.Menu mnuUnitClusterDel 
            Caption         =   "Unit Cluster Delivery"
         End
         Begin VB.Menu mnuUnitClusterDelHist 
            Caption         =   "Unit Cluster Delivery History"
         End
      End
      Begin VB.Menu mnuJobOrderMnu 
         Caption         =   "Job Order"
         Begin VB.Menu mnuJOCellphone 
            Caption         =   "Cellphone"
            Begin VB.Menu mnuCPJobOrder 
               Caption         =   "Job Order"
            End
            Begin VB.Menu mnuCPJOTransfer 
               Caption         =   "Job Order Transfer"
            End
            Begin VB.Menu mnuJobOrderReceiving 
               Caption         =   "Job Order Receiving"
               Begin VB.Menu mnuJOReceivingBranch 
                  Caption         =   "Branch"
               End
               Begin VB.Menu mnuJOReceivingSrvcCntr 
                  Caption         =   "Service Center"
               End
            End
         End
         Begin VB.Menu mnuWrtAccessories 
            Caption         =   "Accessories"
            Begin VB.Menu mnuAServiceCenter 
               Caption         =   "Service Center"
            End
            Begin VB.Menu mnuAForwarded 
               Caption         =   "Forwarded"
            End
         End
      End
      Begin VB.Menu mnuPos 
         Caption         =   "POS"
      End
      Begin VB.Menu mnuReceipt 
         Caption         =   "Receipt"
      End
      Begin VB.Menu mnuWholeSale 
         Caption         =   "CP Whole Sale"
      End
      Begin VB.Menu mnuWholeSaleReturn 
         Caption         =   "CP Whole Sale Return"
      End
      Begin VB.Menu mnuChargeInvoice 
         Caption         =   "Charge Invoice"
      End
      Begin VB.Menu mnuSalesReturn 
         Caption         =   "Sales Return"
      End
      Begin VB.Menu mnuMarketingSupport 
         Caption         =   "Marketing Support"
      End
      Begin VB.Menu mnuPriceProtection 
         Caption         =   "Price Protection"
      End
      Begin VB.Menu mnuMCSOverride 
         Caption         =   "MCS Override"
      End
      Begin VB.Menu mnuPayment 
         Caption         =   "Payment"
         Begin VB.Menu mnuARAdjustment 
            Caption         =   "AR Adjustment"
         End
         Begin VB.Menu mnuARPayment 
            Caption         =   "AR Payment"
         End
      End
      Begin VB.Menu mnuPosting 
         Caption         =   "Posting"
         Begin VB.Menu mnuReceiveTransfer 
            Caption         =   "Receive Transfer"
         End
         Begin VB.Menu mnuPostLoadTransfer 
            Caption         =   "Load Transfer"
         End
         Begin VB.Menu mnuLoadSplitAdj 
            Caption         =   "Load Split Adj"
         End
         Begin VB.Menu mnuBranchReceived 
            Caption         =   "Branch Received"
         End
         Begin VB.Menu mnuInvAdjPosting 
            Caption         =   "Inventory Adjustment"
         End
         Begin VB.Menu mnuLoadAdjPos 
            Caption         =   "Load Adjustment "
         End
         Begin VB.Menu mnuChargeInvoicePosting 
            Caption         =   "Charge Invoice"
         End
         Begin VB.Menu mnuAssetsTransfer 
            Caption         =   "Receive Assets Transfer"
         End
         Begin VB.Menu mnuSuppliesTransferPost 
            Caption         =   "Receive Supplies Transfer"
         End
         Begin VB.Menu mnuReceiveServicePhone 
            Caption         =   "Receive Service Phone"
         End
      End
      Begin VB.Menu mnuInvClass 
         Caption         =   "Inventory Classification"
      End
      Begin VB.Menu mnuInvClassifyUnit 
         Caption         =   "Inventory Classification Unit"
      End
      Begin VB.Menu mnuInventoryCount 
         Caption         =   "Inventory Count"
      End
      Begin VB.Menu mnuInvTypeTransfer 
         Caption         =   "Inventory Type Transfer"
      End
      Begin VB.Menu mnuCPStockOrder 
         Caption         =   "Stock Order w/ ROQ"
      End
      Begin VB.Menu mnuCPUnitStockOrder 
         Caption         =   "CP Unit Stock Order w/ ROQ"
      End
      Begin VB.Menu mnuAppEntry 
         Caption         =   "Application Entry"
      End
      Begin VB.Menu mnuAppApproval 
         Caption         =   "Application Approval"
      End
      Begin VB.Menu mnuRaffleEntryScanner 
         Caption         =   "Raffle Entry Scanner"
      End
   End
   Begin VB.Menu mnuGenActng 
      Caption         =   "General Accounts"
      Begin VB.Menu mnuCashierManager 
         Caption         =   "&Cashier Manager"
         Begin VB.Menu mnuPCCashAdvance 
            Caption         =   "Cash Advance Entry"
         End
         Begin VB.Menu mnuLiquidationEntry 
            Caption         =   "Liquidation Entry"
         End
         Begin VB.Menu mnuCashDisbursement 
            Caption         =   "Cash Disbursement Entry"
         End
         Begin VB.Menu mnuReplenishment 
            Caption         =   "Replinishment Entry"
         End
      End
      Begin VB.Menu mnuCashManApproval 
         Caption         =   "Cash Manager Approval"
         Begin VB.Menu mnuPCCashAdvancApprvl 
            Caption         =   "Cash Advance Approval"
         End
         Begin VB.Menu mnuLiquidationApprvl 
            Caption         =   "Liquidation Approval"
         End
         Begin VB.Menu mnuCashDisbursementApprvl 
            Caption         =   "Cash Disbursement Aproval"
         End
         Begin VB.Menu mnuReplenishmentApprvl 
            Caption         =   "Replenishment Approval"
         End
         Begin VB.Menu mnuGASep01 
            Caption         =   "-"
         End
      End
      Begin VB.Menu mnuCashDep 
         Caption         =   "Cash Deposit"
      End
      Begin VB.Menu mnuCheckDep 
         Caption         =   "Check Deposit"
      End
      Begin VB.Menu mnuSupplyManager 
         Caption         =   "&Supplies Manager"
         Begin VB.Menu mnuSuppliesRequest 
            Caption         =   "Supplies Stock Request"
         End
         Begin VB.Menu mnuSuppliesTransfer 
            Caption         =   "Supplies Stock Transfer"
         End
         Begin VB.Menu mnuSuppliesPosting 
            Caption         =   "Supplies Transfer Acceptance"
         End
      End
      Begin VB.Menu mnuOtherAssets 
         Caption         =   "Other Ass&ets"
         Index           =   0
         Begin VB.Menu mnuAssetsPORec 
            Caption         =   "Assets PO Receiving"
         End
         Begin VB.Menu mnuAssetRequest 
            Caption         =   "Assets Stock Request"
         End
         Begin VB.Menu mnuAssetTransfer 
            Caption         =   "Assets Stock Transfer"
         End
         Begin VB.Menu mnuAssetTransferAcceptance 
            Caption         =   "Assets Teansfer Acceptance"
         End
      End
      Begin VB.Menu mnuDocument 
         Caption         =   "Document Transfer"
         Begin VB.Menu mnuDocumentTransfer 
            Caption         =   "Document Transfer"
         End
         Begin VB.Menu mnuDocumentAccept 
            Caption         =   "Document Acceptance"
         End
      End
      Begin VB.Menu mnuPetMgr 
         Caption         =   "&PET Manager"
         Begin VB.Menu mnuPApplications 
            Caption         =   "Applications"
            Begin VB.Menu mnuPALoan 
               Caption         =   "Loan"
            End
            Begin VB.Menu mnuPAAdvances 
               Caption         =   "Advances"
            End
            Begin VB.Menu mnuPALeave 
               Caption         =   "Leave"
            End
            Begin VB.Menu mnuPABusinessTrip 
               Caption         =   "Business Trip"
            End
            Begin VB.Menu mnuPAOBTripWLog 
               Caption         =   "Business Trip W/ Log"
            End
            Begin VB.Menu mnuPTForgot 
               Caption         =   "Forgot to Log"
            End
            Begin VB.Menu mnuPAOvertime 
               Caption         =   "Overtime"
            End
            Begin VB.Menu mnuPAUndertime 
               Caption         =   "Undertime"
            End
            Begin VB.Menu mnuPATardiness 
               Caption         =   "Tardiness"
            End
         End
         Begin VB.Menu mnuPRequest 
            Caption         =   "Request"
            Begin VB.Menu mnuPRMovement 
               Caption         =   "Employee Movement"
            End
            Begin VB.Menu mnuPRHiring 
               Caption         =   "Employee Hiring"
            End
            Begin VB.Menu mnuPRTermination 
               Caption         =   "Employee Termination"
            End
            Begin VB.Menu mnuPRSuspension 
               Caption         =   "Suspension"
            End
            Begin VB.Menu mnuPRShiftMovement 
               Caption         =   "Schedule Shifting"
            End
            Begin VB.Menu mnuPRDayoffShifting 
               Caption         =   "Day-off Shifting"
            End
         End
         Begin VB.Menu mnuPApprovals 
            Caption         =   "Approvals"
            Begin VB.Menu mnuPVLeave 
               Caption         =   "Leave"
            End
            Begin VB.Menu mnuPVAdvances 
               Caption         =   "Advances"
            End
            Begin VB.Menu mnuPVOvertime 
               Caption         =   "Overtime"
            End
            Begin VB.Menu mnuPVUndertime 
               Caption         =   "Undertime"
            End
            Begin VB.Menu mnuPVTardiness 
               Caption         =   "Tardiness"
            End
            Begin VB.Menu mnuPVBusinessTrip 
               Caption         =   "Business Trip"
            End
            Begin VB.Menu mnuPVOBTripWLog 
               Caption         =   "Business Trip W/ Log"
            End
            Begin VB.Menu mnuPVForgot2Swipe 
               Caption         =   "Forgot to Log"
            End
            Begin VB.Menu mnuPVShiftMovement 
               Caption         =   "Schedule Shifting"
            End
            Begin VB.Menu mnuPVDayoffShifting 
               Caption         =   "Day-off Shifting"
            End
            Begin VB.Menu mnuPVManualLog 
               Caption         =   "Manual Log"
            End
         End
         Begin VB.Menu mnuPAttendance 
            Caption         =   "Attendance"
            Begin VB.Menu mnuPTProcessLog 
               Caption         =   "Process Log"
            End
            Begin VB.Menu mnuPTManualLog 
               Caption         =   "Manual Log"
            End
            Begin VB.Menu mnuPUExport 
               Caption         =   "Export Attendance"
            End
         End
         Begin VB.Menu mnuYearEndBunos 
            Caption         =   "Year End Bonus Entry"
         End
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&Reports"
      Begin VB.Menu mnuStandardReports 
         Caption         =   "Standard Reports"
      End
      Begin VB.Menu mnuJobOrderReports 
         Caption         =   "Job Order Reports"
      End
      Begin VB.Menu mnuAuditReports 
         Caption         =   "Audit Reports"
      End
      Begin VB.Menu mnuGenAccountsRep 
         Caption         =   "General Accounts Report"
      End
      Begin VB.Menu mnuManagerRep 
         Caption         =   "Manager Reports"
      End
   End
   Begin VB.Menu mnuRegisters 
      Caption         =   "&Registers"
      Begin VB.Menu mnuPOSReg 
         Caption         =   "POS"
      End
      Begin VB.Menu mnuReceiptReg 
         Caption         =   "Receipt"
      End
      Begin VB.Menu mnuCharge_Invoice_Reg 
         Caption         =   "Charge Invoice"
      End
      Begin VB.Menu mnuSales_Return_Reg 
         Caption         =   "Sales Return"
      End
      Begin VB.Menu mnuEloadReg 
         Caption         =   "Eload"
      End
      Begin VB.Menu mnuLoadWalletReg 
         Caption         =   "Load Wallet"
      End
      Begin VB.Menu mnuCPInvStockReqReg 
         Caption         =   "CP Stock Request"
      End
      Begin VB.Menu mnuCPInvUnitReqReg 
         Caption         =   "CP Unit Request"
      End
      Begin VB.Menu mnuRegPurchasing 
         Caption         =   "Purchasing"
         Begin VB.Menu mnuRegPurchaseReceiving 
            Caption         =   "Purchase Receiving"
         End
         Begin VB.Menu mnuRegPurchaseReturn 
            Caption         =   "Purchase Return"
         End
         Begin VB.Menu mnuPurchaseReplacementReg 
            Caption         =   "Purchase Replacement"
         End
         Begin VB.Menu mnuRegPurchaseOrder 
            Caption         =   "Purchase Order"
         End
         Begin VB.Menu mnuLoadReceivingReg 
            Caption         =   "Load Receiving"
         End
      End
      Begin VB.Menu mnuJobOrder 
         Caption         =   "Job Order"
         Begin VB.Menu mnuJobOrderReg 
            Caption         =   "Job Order"
         End
         Begin VB.Menu mnuJobOrderTransferReg 
            Caption         =   "Job Order Transfer"
         End
      End
      Begin VB.Menu mnuTransferReg 
         Caption         =   "Transfer"
         Begin VB.Menu mnuRegStockIssue 
            Caption         =   "Stock Transfer"
         End
         Begin VB.Menu mnuLoadTransferReg 
            Caption         =   "Load Transfer"
         End
         Begin VB.Menu mnuServicePhoneTransferReg 
            Caption         =   "Service Phone Transfer"
         End
      End
      Begin VB.Menu mnuRegWarranty 
         Caption         =   "Warranty"
         Begin VB.Menu mnuRegServiceCenter 
            Caption         =   "Service Center"
         End
         Begin VB.Menu mnuWAccessoriesReg 
            Caption         =   "Accessories"
         End
      End
      Begin VB.Menu mnuRegInvAdjustment 
         Caption         =   "Inventory Adjustment"
      End
      Begin VB.Menu mnuLoadAdjustmentReg 
         Caption         =   "Load Adjustment"
      End
      Begin VB.Menu mnuSalesByDate 
         Caption         =   "Sales By Date"
      End
      Begin VB.Menu mnuEloadPosting 
         Caption         =   "ELoad By Date"
      End
      Begin VB.Menu mnuRPetMgr 
         Caption         =   "PET Manager"
         Begin VB.Menu mnuRPApplications 
            Caption         =   "Applications"
            Begin VB.Menu mnuRPALoan 
               Caption         =   "Loan"
            End
            Begin VB.Menu mnuRPAAdvances 
               Caption         =   "Advances"
            End
            Begin VB.Menu mnuRPALeave 
               Caption         =   "Leave"
            End
            Begin VB.Menu mnuRPABusinessTrip 
               Caption         =   "Business Trip"
            End
            Begin VB.Menu mnuRPAOBTripWLog 
               Caption         =   "Business Trip W/ Log"
            End
            Begin VB.Menu mnuRPAOvertime 
               Caption         =   "Overtime"
            End
            Begin VB.Menu mnuRPAUndertime 
               Caption         =   "Undertime"
            End
            Begin VB.Menu mnuRPATardiness 
               Caption         =   "Tardiness"
            End
         End
         Begin VB.Menu mnuRPRequest 
            Caption         =   "Request"
            Begin VB.Menu mnuRPRMovement 
               Caption         =   "Employee Movement"
            End
            Begin VB.Menu mnuRPRHiring 
               Caption         =   "Employee Hiring"
            End
            Begin VB.Menu mnuRPRTermination 
               Caption         =   "Employee Termination"
            End
            Begin VB.Menu mnuRPRSuspension 
               Caption         =   "Suspension"
            End
            Begin VB.Menu mnuRPRShiftMovement 
               Caption         =   "Schedule Shifting"
            End
            Begin VB.Menu mnuRPRDayoffShifting 
               Caption         =   "Day-off Shifting"
            End
         End
         Begin VB.Menu mnuRPAttendance 
            Caption         =   "Attendance"
            Begin VB.Menu mnuRPTManualLog 
               Caption         =   "Manual Log"
            End
         End
      End
      Begin VB.Menu mnuAssetsReg 
         Caption         =   "Ass&ets"
         Index           =   0
         Begin VB.Menu mnuEAssets 
            Caption         =   "Asset Transfer"
            Index           =   3
         End
      End
      Begin VB.Menu mnuRSuppliesManager 
         Caption         =   "Supplies Manager"
         Begin VB.Menu mnuRSupplies 
            Caption         =   "Supply Request"
            Index           =   0
         End
         Begin VB.Menu mnuRSupplies 
            Caption         =   "Supply Adjustment"
            Index           =   1
         End
      End
      Begin VB.Menu mnuRCashierManager 
         Caption         =   "Cashier Manager"
         Begin VB.Menu mnuReplinishmentReg 
            Caption         =   "Replenishment"
            Index           =   0
         End
         Begin VB.Menu mnuCashAdvReg 
            Caption         =   "Cash Advance"
            Index           =   1
         End
         Begin VB.Menu mnuLiquadationReg 
            Caption         =   "Liquidation"
            Index           =   2
         End
         Begin VB.Menu mnuCashDisReg 
            Caption         =   "Cash Disbursement"
            Index           =   3
         End
         Begin VB.Menu mnuCashDepReg 
            Caption         =   "Cash Deposit"
         End
         Begin VB.Menu mnuCheckDepReg 
            Caption         =   "Check Deposit"
         End
      End
   End
   Begin VB.Menu mnuUtilities 
      Caption         =   "&Utilities"
      Begin VB.Menu mnuProductInquiry 
         Caption         =   "Product Inquiry"
      End
      Begin VB.Menu mnuPrintBarcode 
         Caption         =   "Print Barcode"
      End
      Begin VB.Menu mnuPrintBarcodeLX310 
         Caption         =   "Print Barcode by Laser Jet"
      End
      Begin VB.Menu mnuCPSerial 
         Caption         =   "CP Serial"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCPSellPriceEntry 
         Caption         =   "CP Selling Price Entry"
      End
      Begin VB.Menu mnuInventoryAdjustment 
         Caption         =   "Inventory Adjustment"
      End
      Begin VB.Menu mnuLoadAdjustment 
         Caption         =   "Load Adjustment"
      End
      Begin VB.Menu mnuCPSRP 
         Caption         =   "CP SRP"
      End
      Begin VB.Menu mnuStockInquiry 
         Caption         =   "Stock Inquiry"
      End
      Begin VB.Menu mnuUActiveAccounts 
         Caption         =   "Active Accounts"
      End
      Begin VB.Menu mnuUInactiveAccounts 
         Caption         =   "Inactive Accounts"
      End
      Begin VB.Menu mnuUnencodedTrans 
         Caption         =   "UNENCODED TRANSACTION(s)"
      End
      Begin VB.Menu mnuDTRPosting 
         Caption         =   "DTR POSTING"
      End
   End
   Begin VB.Menu mnuAdministrator 
      Caption         =   "&Administrator"
      Begin VB.Menu mnuGSCMCode 
         Caption         =   "Samsung Model Code"
      End
      Begin VB.Menu mnuCPPriceUpdate 
         Caption         =   "CP Price Update"
      End
      Begin VB.Menu mnuCloseDay2Day 
         Caption         =   "Close Day-To-Day Transaction"
      End
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
   
   mdiMain.mnuCPSerial.Visible = oApp.UserLevel = xeEngineer
   mdiMain.mnuManagerRep.Visible = oApp.isMainOffice = True Or oApp.IsWarehouse = True
   mdiMain.mnuWholeSale.Visible = oApp.isMainOffice = True Or oApp.IsWarehouse = True
   mdiMain.mnuWholeSaleReturn.Visible = oApp.isMainOffice = True Or oApp.IsWarehouse = True
   mdiMain.mnuChargeInvoice.Visible = oApp.isMainOffice = True Or oApp.IsWarehouse = True
   mdiMain.mnuMarketingSupport.Visible = oApp.isMainOffice = True Or oApp.IsWarehouse = True
   mdiMain.mnuPriceProtection.Visible = oApp.isMainOffice = True Or oApp.IsWarehouse = True
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
   frmCPCashPrice.Show
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
   frmCashAdvanceApproval.Tag = "mnuPCCashAdvancApprvl"
   frmCashAdvanceApproval.Show
End Sub

Private Sub mnuPCCashAdvance_Click()
   frmCashAdvanceEntry.Tag = "mnuPCCashAdvance"
   frmCashAdvanceEntry.Show
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
   frmBarrCodeLX310.Tag = "mnuPrintBarcodeLX31"
   frmBarrCodeLX310.Show
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
   Dim loRS As Recordset
   
   lsSQL = "SELECT" & _
                  " a.dCovergTo" & _
          " FROM Payroll_Period a" & _
              " LEFT JOIN Payroll_Summary b ON a.sPayPerID = b.sPayPerID" & _
          " WHERE b.sEmployID = " & strParm(fsEmployID) & _
          " ORDER BY a.dCovergTo DESC LIMIT 1"
   Set loRS = oApp.Connection.Execute(lsSQL, , adCmdText)
   
   If loRS.EOF Then
      getLastPeriod = Format(oApp.ServerDate, "yyyy-mm-dd")
   Else
      getLastPeriod = loRS("dCovergTo") + 1
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
   Dim loRS As Recordset
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
      
      Set loRS = oApp.Connection.Execute(lsSQL, , adCmdText)
      DoEvents
      
      If loRS.EOF Then Exit Sub
      
      Set loCls = New clsEmployeeMovement
      'Set loCls.AppDriver = oApp
      'loCls.HasParent = True
      loCls.InitTransaction

      DoEvents
      Do Until loRS.EOF
         DoEvents
         If LCase(oApp.ProductID) = "petmgr" Then
            'Its from the main office so send updates to all branches...
            If loCls.OpenTransaction(loRS("sTransNox")) Then
               loCls.PostTransaction (loRS("sTransNox"))
            End If
         Else
            'if monitor is not from main office then just post the movement
            lsSQL = "UPDATE Employee_Movement" & _
                   " SET cTranStat = " & strParm(xeStatePosted) & _
                   " WHERE sTransNox = " & strParm(loRS("sTransNox"))
            oApp.Connection.Execute lsSQL, , adCmdText
         End If

          'From this Branch employee is assigned to other branch
         If IFNull(loCls.Master("sBranchCD")) <> "" _
            And loCls.Master("sBranchCD") <> oApp.BranchCode _
            And IFNull(loCls.Master("xBranchCD"), "") = oApp.BranchCode _
            And InStr(1, "M001»M0W1", loRS("sBranchCD")) = 0 Then

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

         loRS.MoveNext
         DoEvents
      Loop
   End If
End Sub
