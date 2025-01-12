if exists (select * from sysobjects where id = object_id(N'[dbo].[Banks]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Banks]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[Branch]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Branch]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[Brand]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Brand]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[Card]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Card]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[Category]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Category]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[Client_Ledger]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Client_Ledger]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[Client_Master]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Client_Master]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[Color]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Color]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[Company]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Company]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[Country]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Country]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[CP_Expense_Detail]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CP_Expense_Detail]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[CP_Expense_Master]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CP_Expense_Master]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[CP_Inventory]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CP_Inventory]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[CP_Inventory_Ledger]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CP_Inventory_Ledger]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[CP_Inventory_Master]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CP_Inventory_Master]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[CP_JobOrder_Detail]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CP_JobOrder_Detail]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[CP_JobOrder_Master]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CP_JobOrder_Master]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[CP_Serial_Dummy]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CP_Serial_Dummy]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[CP_Serial_Ledger]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CP_Serial_Ledger]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[CP_Serial_Master]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CP_Serial_Master]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[CP_Serial_Transfer_Detail]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CP_Serial_Transfer_Detail]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[CP_Serial_Transfer_Master]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CP_Serial_Transfer_Master]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[CP_SO_Cheque]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CP_SO_Cheque]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[CP_SO_Credit]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CP_SO_Credit]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[CP_SO_Detail]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CP_SO_Detail]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[CP_SO_Installment]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CP_SO_Installment]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[CP_SO_Master]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CP_SO_Master]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[CP_SO_Serial]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CP_SO_Serial]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[CP_Transfer_Detail]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CP_Transfer_Detail]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[CP_Transfer_Master]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CP_Transfer_Master]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[Credit_Card]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Credit_Card]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[ELoad_Ledger]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ELoad_Ledger]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[ELoad_Matrix]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ELoad_Matrix]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[Employee_Log(2005)]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Employee_Log(2005)]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[Made]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Made]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[Model]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Model]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[PO_Receiving_Detail]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[PO_Receiving_Detail]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[PO_Receiving_Master]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[PO_Receiving_Master]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[PO_Receiving_Serial]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[PO_Receiving_Serial]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[PO_Return_Detail]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[PO_Return_Detail]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[PO_Return_Master]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[PO_Return_Master]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[PO_Return_Serial]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[PO_Return_Serial]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[Province]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Province]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[Results]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Results]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[Sales_Person]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Sales_Person]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[Supplier]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Supplier]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[Supplier2]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Supplier2]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[Term]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Term]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[TownCity]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[TownCity]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[xxxAppObject]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[xxxAppObject]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[xxxComputer]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[xxxComputer]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[xxxDeletedRec]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[xxxDeletedRec]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[xxxDepartment]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[xxxDepartment]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[xxxImportTable]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[xxxImportTable]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[xxxMenuObject]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[xxxMenuObject]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[xxxReport]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[xxxReport]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[xxxReportDetail]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[xxxReportDetail]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[xxxReportMaster]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[xxxReportMaster]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[xxxReportsLog]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[xxxReportsLog]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[xxxSkin]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[xxxSkin]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[xxxSysMonitor]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[xxxSysMonitor]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[xxxSysObject]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[xxxSysObject]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[xxxSysUser]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[xxxSysUser]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[xxxSysUserLog]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[xxxSysUserLog]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[xxxTransactionSource]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[xxxTransactionSource]
GO

CREATE TABLE [dbo].[Banks] (
	[sBankIDxx] [varchar] (7) NOT NULL ,
	[sBankName] [varchar] (30) NULL ,
	[sContactP] [varchar] (30) NULL ,
	[sAddressx] [varchar] (40) NULL ,
	[sTownIdxx] [varchar] (4) NULL ,
	[sTelNoxxx] [varchar] (30) NULL ,
	[sFaxNoxxx] [varchar] (15) NULL ,
	[cRecdStat] [char] (1) NULL ,
	[sModified] [varchar] (8) NULL ,
	[dModified] [datetime] NULL ,
	[vTimestmp] [binary] (8) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Branch] (
	[sBranchCd] [varchar] (2) NOT NULL ,
	[sBranchNm] [varchar] (50) NULL ,
	[sCompnyID] [varchar] (2) NULL ,
	[sAddressx] [varchar] (50) NULL ,
	[sTownIDxx] [varchar] (4) NULL ,
	[sManagerx] [varchar] (8) NULL ,
	[cRecdStat] [char] (1) NULL ,
	[cMainOffc] [char] (1) NULL ,
	[sContactx] [varchar] (50) NULL ,
	[sModified] [varchar] (8) NULL ,
	[dModified] [datetime] NULL ,
	[vTimeStmp] [timestamp] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Brand] (
	[sBrandIDx] [varchar] (5) NOT NULL ,
	[sBrandNme] [varchar] (25) NULL ,
	[cRecdStat] [char] (1) NULL ,
	[sModified] [varchar] (20) NULL ,
	[dModified] [datetime] NULL ,
	[vTimeStmp] [timestamp] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Card] (
	[sCardIDxx] [varchar] (5) NOT NULL ,
	[sCardName] [varchar] (25) NULL ,
	[cRecdStat] [char] (1) NULL ,
	[sModified] [varchar] (20) NULL ,
	[dModified] [datetime] NULL ,
	[vTimeStmp] [timestamp] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Category] (
	[sCategIDx] [varchar] (5) NOT NULL ,
	[sCategNme] [varchar] (20) NULL ,
	[cRecdStat] [char] (1) NULL ,
	[sModified] [varchar] (20) NULL ,
	[dModified] [datetime] NULL ,
	[vTimeStmp] [timestamp] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Client_Ledger] (
	[sClientID] [varchar] (10) NOT NULL ,
	[sBranchCd] [varchar] (2) NOT NULL ,
	[dTransact] [datetime] NULL ,
	[sSourceCd] [varchar] (4) NULL ,
	[sSourceNo] [varchar] (10) NULL ,
	[nCreditxx] [decimal](10, 2) NULL ,
	[nDebitxxx] [decimal](10, 2) NULL ,
	[nABalance] [decimal](10, 2) NULL ,
	[dModified] [datetime] NULL ,
	[vTimeStmp] [binary] (8) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Client_Master] (
	[sClientID] [varchar] (10) NOT NULL ,
	[sLastName] [varchar] (30) NULL ,
	[sFrstName] [varchar] (30) NULL ,
	[sMiddName] [varchar] (30) NULL ,
	[cGenderCd] [char] (1) NULL ,
	[cCvilStat] [char] (1) NULL ,
	[sCitizenx] [varchar] (2) NULL ,
	[dBirthDte] [datetime] NULL ,
	[sBirthPlc] [varchar] (30) NULL ,
	[sHouseNox] [varchar] (5) NULL ,
	[sAddressx] [varchar] (50) NULL ,
	[sTownIDxx] [varchar] (5) NULL ,
	[sMobileNo] [varchar] (30) NULL ,
	[sPhoneNox] [varchar] (30) NULL ,
	[sEmailAdd] [varchar] (30) NULL ,
	[sTaxIDNox] [varchar] (15) NULL ,
	[sAddlInfo] [varchar] (50) NULL ,
	[sCompnyNm] [varchar] (50) NULL ,
	[sClientNo] [varchar] (15) NULL ,
	[cRecdStat] [char] (1) NULL ,
	[sSpouseID] [varchar] (10) NOT NULL ,
	[sModified] [varchar] (8) NULL ,
	[dModified] [datetime] NULL ,
	[vTimeStmp] [timestamp] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Color] (
	[sColorIDx] [varchar] (5) NOT NULL ,
	[sColorNme] [varchar] (15) NULL ,
	[cRecdStat] [char] (1) NULL ,
	[sModified] [varchar] (8) NULL ,
	[dModified] [datetime] NULL ,
	[vTimeStmp] [binary] (8) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Company] (
	[sCompnyID] [varchar] (2) NOT NULL ,
	[sCompnyNm] [varchar] (30) NULL ,
	[sCompnyCd] [varchar] (5) NULL ,
	[cRecdStat] [char] (1) NULL ,
	[sModified] [varchar] (8) NULL ,
	[dModified] [datetime] NULL ,
	[vTimeStmp] [binary] (8) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Country] (
	[sCntryCde] [varchar] (2) NOT NULL ,
	[sCntryNme] [varchar] (25) NULL ,
	[sNational] [varchar] (25) NULL ,
	[cRecdStat] [char] (1) NULL ,
	[sModified] [varchar] (8) NULL ,
	[dModified] [datetime] NULL ,
	[vTimeStmp] [binary] (8) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CP_Expense_Detail] (
	[sTransNox] [varchar] (10) NOT NULL ,
	[nEntryNox] [tinyint] NOT NULL ,
	[sDescript] [varchar] (50) NULL ,
	[nAmountxx] [numeric](9, 2) NULL ,
	[sModified] [varchar] (8) NULL ,
	[dModified] [datetime] NULL ,
	[vTimeStmp] [timestamp] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CP_Expense_Master] (
	[sTransNox] [varchar] (10) NOT NULL ,
	[sBranchCd] [varchar] (2) NOT NULL ,
	[dTranDate] [datetime] NOT NULL ,
	[nTotalExp] [numeric](8, 2) NOT NULL ,
	[sModified] [varchar] (8) NULL ,
	[dModified] [datetime] NULL ,
	[vTimestmp] [timestamp] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CP_Inventory] (
	[sStockIDx] [varchar] (10) NOT NULL ,
	[sBarrcode] [varchar] (20) NOT NULL ,
	[sDescript] [varchar] (50) NOT NULL ,
	[sSupplier] [varchar] (10) NULL ,
	[sBrandIDx] [varchar] (5) NOT NULL ,
	[sModelIDx] [varchar] (7) NOT NULL ,
	[sMadeIDxx] [varchar] (3) NULL ,
	[sColorIDx] [varchar] (5) NULL ,
	[sCategIDx] [varchar] (5) NULL ,
	[sCardIDxx] [varchar] (5) NULL ,
	[cCellPhon] [char] (1) NULL ,
	[cCellCard] [char] (1) NULL ,
	[cCellLoad] [char] (1) NULL ,
	[cWalletxx] [char] (1) NULL ,
	[nPurPrice] [numeric](10, 2) NULL ,
	[nLastPrce] [numeric](10, 2) NULL ,
	[nSelPrice] [numeric](10, 2) NULL ,
	[dLastDate] [datetime] NULL ,
	[cRecdStat] [char] (1) NULL ,
	[sModified] [varchar] (8) NULL ,
	[dModified] [datetime] NULL ,
	[vTimeStmp] [timestamp] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CP_Inventory_Ledger] (
	[sStockIDx] [varchar] (10) NULL ,
	[sBranchCd] [varchar] (2) NULL ,
	[dTransact] [datetime] NULL ,
	[sSourceCd] [varchar] (4) NULL ,
	[sSourceNo] [varchar] (10) NULL ,
	[nQtyInxxx] [smallint] NULL ,
	[nQtyOutxx] [smallint] NULL ,
	[nEntryNox] [tinyint] NULL ,
	[nQtyOnHnd] [smallint] NULL ,
	[dModified] [datetime] NULL ,
	[vTimeStmp] [timestamp] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CP_Inventory_Master] (
	[sStockIDx] [varchar] (10) NOT NULL ,
	[sBranchCd] [varchar] (2) NULL ,
	[nBegQtyxx] [numeric](7, 2) NULL ,
	[nQtyOnHnd] [decimal](7, 2) NULL ,
	[nReorderx] [smallint] NULL ,
	[nMinLevel] [smallint] NULL ,
	[nMaxLevel] [smallint] NULL ,
	[dBegInvxx] [datetime] NULL ,
	[cRecdStat] [char] (1) NULL ,
	[sModified] [varchar] (8) NULL ,
	[dModified] [datetime] NULL ,
	[vTimeStmp] [timestamp] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CP_JobOrder_Detail] (
	[sTransNox] [varchar] (10) NOT NULL ,
	[nEntryNox] [tinyint] NOT NULL ,
	[sDescript] [varchar] (50) NOT NULL ,
	[nPartsAmt] [numeric](8, 2) NULL ,
	[nLaborAmt] [numeric](8, 2) NULL ,
	[nQuantity] [smallint] NULL ,
	[nDiscount] [smallint] NULL ,
	[dModified] [datetime] NULL ,
	[vTimeStmp] [timestamp] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CP_JobOrder_Master] (
	[sTransNox] [varchar] (10) NOT NULL ,
	[dTransact] [datetime] NOT NULL ,
	[sSupplier] [varchar] (10) NULL ,
	[sJobOrdNo] [varchar] (10) NOT NULL ,
	[sClientID] [varchar] (10) NULL ,
	[sReferNox] [varchar] (10) NULL ,
	[sTelNoxxx] [varchar] (30) NULL ,
	[sBrandIDx] [varchar] (5) NULL ,
	[sModelIDx] [varchar] (7) NULL ,
	[sIMEINoxx] [varchar] (20) NULL ,
	[dPurchase] [datetime] NULL ,
	[cWarranty] [char] (1) NULL ,
	[sBckJobNo] [varchar] (10) NULL ,
	[cCategory] [char] (1) NULL ,
	[sCategory] [varchar] (20) NULL ,
	[sComplent] [varchar] (50) NULL ,
	[nLaborTot] [numeric](8, 2) NULL ,
	[nPartsTot] [numeric](8, 2) NULL ,
	[nMiscChrg] [numeric](8, 2) NULL ,
	[nTranTotl] [numeric](8, 2) NULL ,
	[nAmtPaidx] [numeric](8, 2) NULL ,
	[dPaymentx] [datetime] NULL ,
	[sPaymRecv] [varchar] (30) NULL ,
	[sModified] [varchar] (8) NULL ,
	[dModified] [datetime] NULL ,
	[vTimeStmp] [timestamp] NULL ,
	[cTranStat] [char] (1) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CP_Serial_Dummy] (
	[sStockIDx] [varchar] (12) NULL ,
	[sIMEINoxx] [varchar] (20) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CP_Serial_Ledger] (
	[sSerialID] [varchar] (10) NOT NULL ,
	[sBranchcd] [varchar] (2) NOT NULL ,
	[dTransact] [datetime] NOT NULL ,
	[nEntryNox] [tinyint] NOT NULL ,
	[sSourceCd] [varchar] (4) NOT NULL ,
	[sSourceNo] [varchar] (10) NOT NULL ,
	[cSoldStat] [char] (1) NULL ,
	[cLocation] [char] (1) NULL ,
	[dModified] [datetime] NULL ,
	[vTimeStmp] [timestamp] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CP_Serial_Master] (
	[sSerialID] [varchar] (10) NOT NULL ,
	[sBranchCd] [varchar] (20) NOT NULL ,
	[sIMEINoxx] [varchar] (20) NOT NULL ,
	[sStockIDx] [varchar] (10) NOT NULL ,
	[cSoldStat] [char] (1) NULL ,
	[cLocation] [char] (1) NULL ,
	[sClientID] [varchar] (20) NULL ,
	[sModified] [varchar] (8) NULL ,
	[dModified] [datetime] NULL ,
	[vTimeStmp] [binary] (10) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CP_Serial_Transfer_Detail] (
	[sTransNox] [varchar] (10) NOT NULL ,
	[nEntryNox] [tinyint] NOT NULL ,
	[sSerialID] [varchar] (10) NOT NULL ,
	[nUnitPrce] [numeric](8, 2) NULL ,
	[dModified] [datetime] NULL ,
	[vTimeStmp] [timestamp] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CP_Serial_Transfer_Master] (
	[sTransNox] [varchar] (10) NOT NULL ,
	[dTransact] [datetime] NOT NULL ,
	[sDestinat] [varchar] (2) NOT NULL ,
	[sOriginxx] [varchar] (2) NOT NULL ,
	[sRemarksx] [varchar] (100) NULL ,
	[sReferNox] [varchar] (10) NULL ,
	[cReceived] [char] (1) NULL ,
	[dReceived] [datetime] NULL ,
	[cTranStat] [char] (1) NULL ,
	[nEntryNox] [tinyint] NULL ,
	[sRequestx] [varchar] (15) NULL ,
	[sApproved] [varchar] (15) NULL ,
	[sModified] [varchar] (8) NULL ,
	[dModified] [datetime] NULL ,
	[vTimeStmp] [timestamp] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CP_SO_Cheque] (
	[sTransNox] [varchar] (10) NOT NULL ,
	[dTransact] [datetime] NULL ,
	[sClientID] [varchar] (10) NULL ,
	[sBankIDxx] [varchar] (7) NULL ,
	[nTranTotl] [numeric](10, 2) NULL ,
	[nCashAmnt] [numeric](10, 2) NULL ,
	[nCheqAmnt] [numeric](10, 2) NULL ,
	[sAccntNum] [varchar] (25) NULL ,
	[sSalesInv] [varchar] (6) NULL ,
	[dModified] [datetime] NULL ,
	[vTimeStmp] [timestamp] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CP_SO_Credit] (
	[sTransNox] [varchar] (10) NOT NULL ,
	[dTransact] [datetime] NULL ,
	[sClientID] [varchar] (10) NULL ,
	[sCreditID] [varchar] (5) NULL ,
	[nTranTotl] [numeric](10, 2) NULL ,
	[sAcctNmbr] [varchar] (25) NULL ,
	[nPercentx] [numeric](5, 2) NULL ,
	[nCashAmnt] [numeric](10, 2) NULL ,
	[nCardAmnt] [numeric](10, 2) NULL ,
	[sSalesInv] [varchar] (6) NULL ,
	[dModified] [datetime] NULL ,
	[vTimeStmp] [timestamp] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CP_SO_Detail] (
	[sTransNox] [varchar] (10) NOT NULL ,
	[nEntryNox] [tinyint] NULL ,
	[sStockIDx] [varchar] (10) NULL ,
	[nQuantity] [smallint] NULL ,
	[nPurPrice] [numeric](9, 2) NULL ,
	[nUnitPrce] [numeric](9, 2) NULL ,
	[nDiscount] [numeric](4, 2) NULL ,
	[nDiscAmnt] [numeric](6, 2) NULL ,
	[nSubTotal] [numeric](9, 2) NULL ,
	[dModified] [datetime] NULL ,
	[vTimeStmp] [timestamp] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CP_SO_Installment] (
	[sTransNox] [varchar] (10) NULL ,
	[dTransact] [datetime] NULL ,
	[sClientID] [varchar] (10) NULL ,
	[nTranTotl] [decimal](7, 2) NULL ,
	[nDownPaym] [decimal](7, 2) NULL ,
	[nBalancex] [decimal](7, 2) NULL ,
	[nPaymTerm] [smallint] NULL ,
	[nMonthlyP] [decimal](7, 2) NULL ,
	[sSalesInv] [varchar] (6) NULL ,
	[dModified] [datetime] NULL ,
	[vTimeStmp] [timestamp] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CP_SO_Master] (
	[sTransNox] [varchar] (10) NOT NULL ,
	[dTransact] [datetime] NULL ,
	[sClientID] [varchar] (10) NULL ,
	[sSalesInv] [varchar] (15) NULL ,
	[nTranTotl] [numeric](10, 2) NULL ,
	[nAmtPaidx] [numeric](10, 2) NULL ,
	[nGiftCpnx] [varchar] (10) NULL ,
	[cTranStat] [char] (1) NULL ,
	[sCashierx] [varchar] (15) NULL ,
	[sRemarksx] [varchar] (30) NULL ,
	[sModified] [varchar] (8) NULL ,
	[dModified] [datetime] NULL ,
	[vTimeStmp] [timestamp] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CP_SO_Serial] (
	[sTransNox] [varchar] (10) NOT NULL ,
	[nEntryNox] [tinyint] NOT NULL ,
	[sSerialID] [varchar] (10) NOT NULL ,
	[dModified] [datetime] NULL ,
	[vTimeStmp] [timestamp] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CP_Transfer_Detail] (
	[sTransNox] [varchar] (10) NOT NULL ,
	[nEntryNox] [tinyint] NOT NULL ,
	[sStockIDx] [varchar] (10) NULL ,
	[nQuantity] [smallint] NULL ,
	[nUnitPrce] [numeric](9, 2) NULL ,
	[dModified] [datetime] NULL ,
	[vTimeStmp] [timestamp] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CP_Transfer_Master] (
	[sTransNox] [varchar] (10) NOT NULL ,
	[dTransact] [datetime] NULL ,
	[sDestinat] [varchar] (2) NULL ,
	[sOriginxx] [varchar] (2) NULL ,
	[sRemarksx] [varchar] (100) NULL ,
	[sReferNox] [varchar] (10) NULL ,
	[cReceived] [char] (1) NULL ,
	[dReceived] [datetime] NULL ,
	[cTranStat] [char] (1) NULL ,
	[nEntryNox] [tinyint] NULL ,
	[sRequestx] [varchar] (15) NULL ,
	[sApproved] [varchar] (15) NULL ,
	[sModified] [varchar] (8) NULL ,
	[dModified] [datetime] NULL ,
	[vTimeStmp] [timestamp] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Credit_Card] (
	[sCreditID] [varchar] (5) NOT NULL ,
	[sCreditNm] [varchar] (25) NULL ,
	[nPercentx] [smallint] NULL ,
	[cRecdStat] [char] (1) NULL ,
	[sModified] [varchar] (20) NULL ,
	[dModified] [datetime] NULL ,
	[vTimeStmp] [timestamp] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ELoad_Ledger] (
	[sStockIDx] [varchar] (10) NOT NULL ,
	[sBranchCd] [varchar] (2) NOT NULL ,
	[dTransact] [datetime] NULL ,
	[sReferNox] [varchar] (25) NULL ,
	[sPhoneNum] [varchar] (25) NULL ,
	[sSourceCd] [varchar] (4) NULL ,
	[sSourceNo] [varchar] (10) NULL ,
	[sTransNox] [tinyint] NULL ,
	[nQtyInxxx] [decimal](9, 2) NULL ,
	[nQtyOutxx] [decimal](9, 2) NULL ,
	[nEntryNox] [tinyint] NULL ,
	[nQtyOnHnd] [decimal](9, 2) NULL ,
	[sModified] [varchar] (10) NULL ,
	[dModified] [datetime] NULL ,
	[vTimeStmp] [timestamp] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ELoad_Matrix] (
	[sMatrixID] [varchar] (5) NULL ,
	[sStockIDx] [varchar] (10) NULL ,
	[sMatrixNm] [varchar] (25) NULL ,
	[nAmountxx] [numeric](8, 2) NULL ,
	[nSelPrice] [numeric](8, 2) NULL ,
	[sModified] [varchar] (20) NULL ,
	[dModified] [datetime] NULL ,
	[vTimeStmp] [timestamp] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Employee_Log(2005)] (
	[sEmployID] [varchar] (10) NULL ,
	[sBarrcode] [varchar] (20) NULL ,
	[dTranDate] [datetime] NULL ,
	[sAMInxxxx] [varchar] (10) NULL ,
	[sAMOutxxx] [varchar] (10) NULL ,
	[sPMInxxxx] [varchar] (10) NULL ,
	[sPMOutxxx] [varchar] (10) NULL ,
	[sOTimeInx] [varchar] (10) NULL ,
	[sOTimeOut] [varchar] (10) NULL ,
	[nTranLine] [tinyint] NULL ,
	[nTardyxxx] [decimal](18, 0) NULL ,
	[nOverTime] [decimal](18, 0) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Made] (
	[sMadeIDxx] [varchar] (3) NOT NULL ,
	[sMadeName] [varchar] (25) NULL ,
	[cRecdStat] [char] (1) NULL ,
	[sModified] [varchar] (20) NULL ,
	[dModified] [datetime] NULL ,
	[vTimeStmp] [timestamp] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Model] (
	[sModelIDx] [varchar] (7) NOT NULL ,
	[sModelNme] [varchar] (25) NULL ,
	[sBrandIDx] [varchar] (5) NULL ,
	[cRecdStat] [char] (1) NULL ,
	[sModified] [varchar] (20) NULL ,
	[dModified] [datetime] NULL ,
	[vTimeStmp] [binary] (8) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[PO_Receiving_Detail] (
	[sTransNox] [varchar] (10) NOT NULL ,
	[nEntryNox] [tinyint] NOT NULL ,
	[sStockIDx] [varchar] (10) NOT NULL ,
	[nQuantity] [smallint] NULL ,
	[nUnitPrce] [numeric](8, 2) NULL ,
	[nDiscount] [numeric](8, 2) NULL ,
	[nIncentiv] [numeric](8, 2) NULL ,
	[dModified] [datetime] NULL ,
	[vTimeStmp] [timestamp] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[PO_Receiving_Master] (
	[sTransNox] [varchar] (10) NOT NULL ,
	[dTransact] [datetime] NULL ,
	[sClientID] [varchar] (10) NULL ,
	[sReferNox] [varchar] (10) NULL ,
	[sRemarksx] [varchar] (128) NULL ,
	[cTranStat] [char] (1) NULL ,
	[sModified] [varchar] (10) NULL ,
	[dModified] [datetime] NULL ,
	[vTimeStmp] [timestamp] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[PO_Receiving_Serial] (
	[sTransNox] [varchar] (10) NOT NULL ,
	[nEntryNox] [tinyint] NOT NULL ,
	[sSerialID] [varchar] (10) NOT NULL ,
	[dModified] [datetime] NULL ,
	[vTimeStmp] [timestamp] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[PO_Return_Detail] (
	[sTransNox] [varchar] (10) NOT NULL ,
	[nEntryNox] [tinyint] NOT NULL ,
	[sStockIDx] [varchar] (10) NULL ,
	[nQuantity] [smallint] NULL ,
	[nUnitPrce] [decimal](9, 2) NULL ,
	[dModified] [datetime] NULL ,
	[vTimeStmp] [binary] (8) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[PO_Return_Master] (
	[sTransNox] [varchar] (10) NULL ,
	[dTransact] [datetime] NULL ,
	[sClientID] [varchar] (10) NULL ,
	[cTranStat] [char] (1) NULL ,
	[sRemarksx] [varchar] (128) NULL ,
	[sModified] [varchar] (8) NULL ,
	[dModified] [datetime] NULL ,
	[vTimeStmp] [binary] (8) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[PO_Return_Serial] (
	[sTransNox] [varchar] (10) NOT NULL ,
	[nEntryNox] [tinyint] NOT NULL ,
	[sSerialID] [varchar] (10) NULL ,
	[dModified] [datetime] NULL ,
	[vTimeStmp] [binary] (8) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Province] (
	[sProvIDxx] [varchar] (2) NOT NULL ,
	[sProvName] [varchar] (25) NULL ,
	[cRecdStat] [char] (1) NULL ,
	[sModified] [varchar] (8) NULL ,
	[dModified] [datetime] NULL ,
	[vTimeStmp] [binary] (8) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Results] (
	[sSerialID] [varchar] (10) NOT NULL ,
	[sBranchCd] [varchar] (2) NULL ,
	[sEngineNo] [varchar] (20) NULL ,
	[sFrameNox] [varchar] (20) NULL ,
	[sModelIDx] [varchar] (7) NULL ,
	[sColorIDx] [varchar] (5) NULL ,
	[sMCInvIDx] [varchar] (7) NULL ,
	[cSoldStat] [char] (1) NULL ,
	[cLocation] [char] (1) NULL ,
	[cDeliverd] [char] (1) NULL ,
	[cRegister] [char] (1) NULL ,
	[sSalesInv] [varchar] (10) NULL ,
	[cCSRValid] [char] (1) NULL ,
	[cPNPClear] [char] (1) NULL ,
	[sWarrntNo] [varchar] (10) NULL ,
	[sCompnyID] [varchar] (10) NULL ,
	[sClientID] [varchar] (20) NULL ,
	[sFileNoxx] [varchar] (20) NULL ,
	[sRegORNox] [varchar] (15) NULL ,
	[sCRENoxxx] [varchar] (10) NULL ,
	[sCRNoxxxx] [varchar] (10) NULL ,
	[sPlateNoP] [varchar] (8) NULL ,
	[sPlateNoH] [varchar] (8) NULL ,
	[sStickrNo] [varchar] (8) NULL ,
	[sModified] [varchar] (8) NULL ,
	[dModified] [datetime] NULL ,
	[vTimeStmp] [binary] (10) NULL ,
	[cOK] [char] (1) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Sales_Person] (
	[sEmployID] [varchar] (8) NOT NULL ,
	[sLastName] [varchar] (20) NULL ,
	[sFrstName] [varchar] (20) NULL ,
	[sMiddName] [varchar] (20) NULL ,
	[cRecdStat] [char] (1) NULL ,
	[sModified] [varchar] (8) NULL ,
	[dModified] [datetime] NULL ,
	[vTimeStmp] [timestamp] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Supplier] (
	[sSupplyID] [varchar] (5) NOT NULL ,
	[sSupplyNm] [varchar] (50) NOT NULL ,
	[sCPersonx] [varchar] (50) NULL ,
	[sTelNoxxx] [varchar] (30) NULL ,
	[sFaxNoxxx] [varchar] (30) NULL ,
	[sAddressx] [varchar] (150) NULL ,
	[sTownIDxx] [varchar] (4) NULL ,
	[cRecdStat] [char] (1) NULL ,
	[sModified] [varchar] (8) NULL ,
	[dModified] [datetime] NULL ,
	[vTimeStmp] [timestamp] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Supplier2] (
	[sClientID] [varchar] (10) NOT NULL ,
	[sBranchCd] [varchar] (2) NOT NULL ,
	[sCPerson1] [varchar] (30) NULL ,
	[sCPPosit1] [varchar] (15) NULL ,
	[sTelNoxxx] [varchar] (30) NULL ,
	[sFaxNoxxx] [varchar] (30) NULL ,
	[sTermIDxx] [varchar] (5) NULL ,
	[nDiscount] [decimal](4, 2) NULL ,
	[nCredLimt] [decimal](10, 2) NULL ,
	[nABalance] [decimal](10, 2) NULL ,
	[dCltSince] [datetime] NULL ,
	[cRecdStat] [char] (1) NULL ,
	[sModified] [varchar] (8) NULL ,
	[dModified] [datetime] NULL ,
	[vTimeStmp] [binary] (8) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Term] (
	[sTermIDxx] [varchar] (5) NOT NULL ,
	[sTermName] [varchar] (25) NULL ,
	[nTermDays] [smallint] NULL ,
	[nDiscDays] [smallint] NULL ,
	[nDiscount] [decimal](4, 2) NULL ,
	[cRecdStat] [char] (1) NULL ,
	[sModified] [varchar] (8) NULL ,
	[dModified] [datetime] NULL ,
	[vTimeStmp] [binary] (8) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TownCity] (
	[sTownIDxx] [varchar] (4) NOT NULL ,
	[sTownName] [varchar] (30) NULL ,
	[sZippCode] [varchar] (4) NULL ,
	[sProvIDxx] [varchar] (2) NULL ,
	[cHasRoute] [char] (1) NULL ,
	[cRecdStat] [char] (1) NULL ,
	[sModified] [varchar] (8) NULL ,
	[dModified] [datetime] NULL ,
	[vTimeStmp] [binary] (8) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[xxxAppObject] (
	[sProdctId] [varchar] (10) NOT NULL ,
	[sApplName] [varchar] (30) NULL ,
	[sApplPath] [varchar] (50) NULL ,
	[sDriveTrn] [varchar] (2) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[xxxComputer] (
	[sCompName] [varchar] (30) NOT NULL ,
	[sDepartID] [varchar] (2) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[xxxDeletedRec] (
	[sTransNox] [varchar] (10) NOT NULL ,
	[sBranchCd] [varchar] (2) NOT NULL ,
	[sStatemnt] [varchar] (1024) NULL ,
	[dModified] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[xxxDepartment] (
	[sDepartID] [varchar] (2) NOT NULL ,
	[sDepartNm] [varchar] (15) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[xxxImportTable] (
	[sReferNox] [varchar] (8) NOT NULL ,
	[sTableNme] [varchar] (50) NOT NULL ,
	[sBranchCd] [varchar] (2) NULL ,
	[cTableTyp] [varchar] (1) NULL ,
	[sRefFld01] [varchar] (9) NULL ,
	[sRefFld02] [varchar] (9) NULL ,
	[sRefFld03] [varchar] (9) NULL ,
	[sRefFld04] [varchar] (9) NULL ,
	[sRefFld05] [varchar] (9) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[xxxMenuObject] (
	[sMenuIDxx] [varchar] (10) NOT NULL ,
	[sProdctID] [varchar] (10) NOT NULL ,
	[sMenuName] [varchar] (62) NULL ,
	[sMenuDesc] [varchar] (100) NULL ,
	[sRemarksx] [varchar] (50) NULL ,
	[nUserRght] [tinyint] NULL ,
	[nAddRight] [tinyint] NULL ,
	[nUpdRight] [tinyint] NULL ,
	[nCanRight] [tinyint] NULL ,
	[nDelRight] [tinyint] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[xxxReport] (
	[sReportID] [varchar] (6) NOT NULL ,
	[sReportNm] [varchar] (50) NULL ,
	[sFileName] [varchar] (50) NULL ,
	[sReportHd] [varchar] (50) NULL ,
	[sProdctID] [varchar] (8) NULL ,
	[nUserRght] [tinyint] NULL ,
	[cSaveRepx] [char] (1) NULL ,
	[cLogRepxx] [char] (1) NULL ,
	[sProductID] [varchar] (50) NULL ,
	[sModified] [varchar] (8) NULL ,
	[dModified] [datetime] NULL ,
	[vTimeStmp] [binary] (8) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[xxxReportDetail] (
	[sReportID] [varchar] (6) NOT NULL ,
	[nEntryNox] [tinyint] NOT NULL ,
	[sFileName] [varchar] (50) NULL ,
	[sReportHd] [varchar] (50) NULL ,
	[sModified] [varchar] (1) NULL ,
	[dModified] [datetime] NULL ,
	[vTimeStmp] [binary] (8) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[xxxReportMaster] (
	[sReportID] [varchar] (6) NOT NULL ,
	[sReportNm] [varchar] (128) NULL ,
	[sProdctID] [varchar] (54) NULL ,
	[nUserRght] [tinyint] NULL ,
	[cSaveRepx] [char] (1) NULL ,
	[cLogRepxx] [char] (1) NULL ,
	[sRepLibxx] [varchar] (32) NULL ,
	[sRepClass] [varchar] (64) NULL ,
	[sModified] [varchar] (1) NULL ,
	[dModified] [datetime] NULL ,
	[vTimeStmp] [binary] (8) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[xxxReportsLog] (
	[sRepFName] [varchar] (10) NOT NULL ,
	[sReportID] [varchar] (6) NULL ,
	[dGenerate] [datetime] NULL ,
	[sUserIDxx] [varchar] (8) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[xxxSkin] (
	[sSkinCode] [varchar] (2) NOT NULL ,
	[sSkinName] [varchar] (20) NULL ,
	[nColorWT0] [decimal](8, 0) NULL ,
	[nColorWB0] [decimal](8, 0) NULL ,
	[nColorFT0] [decimal](8, 0) NULL ,
	[nColorFB0] [decimal](8, 0) NULL ,
	[nColorET0] [decimal](8, 0) NULL ,
	[nColorEB0] [decimal](8, 0) NULL ,
	[nColorHT0] [decimal](8, 0) NULL ,
	[nColorHB0] [decimal](8, 0) NULL ,
	[nColorHT1] [decimal](8, 0) NULL ,
	[nColorHB1] [decimal](8, 0) NULL ,
	[nColorHT2] [decimal](8, 0) NULL ,
	[nColorHB2] [decimal](8, 0) NULL ,
	[nColorHT3] [decimal](8, 0) NULL ,
	[nColorHB3] [decimal](8, 0) NULL ,
	[nColorHT4] [decimal](8, 0) NULL ,
	[nColorHB4] [decimal](8, 0) NULL ,
	[nColorBC0] [decimal](8, 0) NULL ,
	[nColorBC1] [decimal](8, 0) NULL ,
	[nColorBC2] [decimal](8, 0) NULL ,
	[nColorTC0] [decimal](8, 0) NULL ,
	[sTitleBar] [varchar] (20) NULL ,
	[sQSImagex] [varchar] (20) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[xxxSysMonitor] (
	[sProdctID] [varchar] (10) NOT NULL ,
	[sBranchCd] [varchar] (2) NULL ,
	[dTransact] [datetime] NULL ,
	[cInitProc] [char] (1) NULL ,
	[cInitOKxx] [char] (1) NULL ,
	[sUserIDxx] [varchar] (8) NULL ,
	[sComptrNm] [varchar] (31) NULL ,
	[cCriticl1] [char] (1) NULL ,
	[cCriticl2] [char] (1) NULL ,
	[cCriticl3] [char] (1) NULL ,
	[cCriticl4] [char] (1) NULL ,
	[cCriticl5] [char] (1) NULL ,
	[cIgnorex1] [char] (1) NULL ,
	[cIgnorex2] [char] (1) NULL ,
	[cIgnorex3] [char] (1) NULL ,
	[cIgnorex4] [char] (1) NULL ,
	[cIgnorex5] [char] (1) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[xxxSysObject] (
	[sClientID] [varchar] (10) NOT NULL ,
	[sClientNm] [varchar] (50) NULL ,
	[sAddressx] [varchar] (40) NULL ,
	[sTownName] [varchar] (25) NULL ,
	[sZippCode] [varchar] (4) NULL ,
	[sProvName] [varchar] (25) NULL ,
	[sTelNoxxx] [varchar] (40) NULL ,
	[sFaxNoxxx] [varchar] (12) NULL ,
	[sApproved] [varchar] (30) NULL ,
	[sSysAdmin] [varchar] (15) NULL ,
	[sProdctID] [varchar] (10) NOT NULL ,
	[sProdctNm] [varchar] (50) NULL ,
	[sNetWarex] [varchar] (8) NULL ,
	[sMachinex] [varchar] (8) NULL ,
	[dSysDatex] [datetime] NULL ,
	[dLicencex] [datetime] NULL ,
	[nNetError] [int] NULL ,
	[sBranchCd] [varchar] (2) NULL ,
	[sSkinCode] [varchar] (2) NULL ,
	[dCapturex] [datetime] NULL ,
	[vTimeStmp] [binary] (1) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[xxxSysUser] (
	[sUserIDxx] [varchar] (8) NOT NULL ,
	[sLogNamex] [varchar] (15) NULL ,
	[sPassword] [varchar] (15) NULL ,
	[sUserName] [varchar] (35) NULL ,
	[sEmployNo] [varchar] (8) NULL ,
	[nUserLevl] [int] NULL ,
	[cUserType] [char] (1) NULL ,
	[sProdctID] [varchar] (10) NOT NULL ,
	[cUserStat] [char] (1) NULL ,
	[nSysError] [int] NULL ,
	[cLogStatx] [char] (1) NULL ,
	[cLockStat] [char] (1) NULL ,
	[cAllwLock] [char] (1) NULL ,
	[cAllwView] [char] (1) NULL ,
	[sCompName] [varchar] (30) NULL ,
	[sSkinCode] [varchar] (2) NULL ,
	[sModified] [varchar] (8) NULL ,
	[dModified] [datetime] NULL ,
	[vTimeStmp] [binary] (8) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[xxxSysUserLog] (
	[sLogNoxxx] [varchar] (10) NOT NULL ,
	[sUserIDxx] [varchar] (8) NOT NULL ,
	[dLogInxxx] [datetime] NULL ,
	[dLogOutxx] [datetime] NULL ,
	[sProdctID] [varchar] (10) NULL ,
	[sCompName] [varchar] (30) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[xxxTransactionSource] (
	[sSourceID] [varchar] (4) NOT NULL ,
	[sSourceNm] [varchar] (30) NULL ,
	[sSystemCd] [varchar] (2) NULL ,
	[sTableNme] [varchar] (50) NULL ,
	[sClientTp] [varchar] (15) NULL ,
	[cTranType] [char] (1) NULL ,
	[cRecdStat] [char] (1) NULL ,
	[vTimeStmp] [binary] (8) NULL 
) ON [PRIMARY]
GO

 CREATE  INDEX [IX_Branch] ON [dbo].[Branch]([sBranchCd]) ON [PRIMARY]
GO

 CREATE  UNIQUE  INDEX [IX_Brand] ON [dbo].[Brand]([sBrandIDx]) ON [PRIMARY]
GO

 CREATE  UNIQUE  INDEX [IX_Card] ON [dbo].[Card]([sCardIDxx]) ON [PRIMARY]
GO

 CREATE  UNIQUE  INDEX [IX_Category] ON [dbo].[Category]([sCategIDx]) ON [PRIMARY]
GO

 CREATE  UNIQUE  INDEX [IX_Client_Master] ON [dbo].[Client_Master]([sClientID]) ON [PRIMARY]
GO

 CREATE  UNIQUE  INDEX [IX_Color] ON [dbo].[Color]([sColorIDx]) ON [PRIMARY]
GO

 CREATE  UNIQUE  INDEX [IX_CP_Expense_Detail] ON [dbo].[CP_Expense_Detail]([sTransNox], [nEntryNox]) ON [PRIMARY]
GO

 CREATE  UNIQUE  INDEX [IX_CP_Expense_Master] ON [dbo].[CP_Expense_Master]([sTransNox], [sBranchCd]) ON [PRIMARY]
GO

 CREATE  UNIQUE  INDEX [IX_CP_Inventory] ON [dbo].[CP_Inventory]([sBarrcode]) ON [PRIMARY]
GO

 CREATE  UNIQUE  INDEX [IX_CP_Inventory_Ledger] ON [dbo].[CP_Inventory_Ledger]([sStockIDx], [nEntryNox]) ON [PRIMARY]
GO

 CREATE  UNIQUE  INDEX [IX_CP_JobOrder_Detail] ON [dbo].[CP_JobOrder_Detail]([sTransNox], [nEntryNox]) ON [PRIMARY]
GO

 CREATE  UNIQUE  INDEX [IX_CP_JobOrder_Master] ON [dbo].[CP_JobOrder_Master]([sTransNox]) ON [PRIMARY]
GO

 CREATE  UNIQUE  INDEX [IX_CP_Serial_Dummy] ON [dbo].[CP_Serial_Dummy]([sIMEINoxx]) ON [PRIMARY]
GO

 CREATE  UNIQUE  INDEX [IX_CP_Serial_Ledger] ON [dbo].[CP_Serial_Ledger]([sSerialID], [nEntryNox]) ON [PRIMARY]
GO

 CREATE  UNIQUE  INDEX [IX_CP_Serial_Master] ON [dbo].[CP_Serial_Master]([sIMEINoxx]) ON [PRIMARY]
GO

 CREATE  UNIQUE  INDEX [IX_CP_SO_Credit] ON [dbo].[CP_SO_Credit]([sTransNox]) ON [PRIMARY]
GO

 CREATE  UNIQUE  INDEX [IX_CP_SO_Detail] ON [dbo].[CP_SO_Detail]([sTransNox], [nEntryNox]) ON [PRIMARY]
GO

 CREATE  UNIQUE  INDEX [IX_CP_SO_Installment] ON [dbo].[CP_SO_Installment]([sTransNox]) ON [PRIMARY]
GO

 CREATE  UNIQUE  INDEX [IX_CP_SO_Master] ON [dbo].[CP_SO_Master]([sTransNox]) ON [PRIMARY]
GO

 CREATE  UNIQUE  INDEX [IX_ELoad_Ledger] ON [dbo].[ELoad_Ledger]([sStockIDx], [sReferNox]) ON [PRIMARY]
GO

 CREATE  UNIQUE  INDEX [IX_ELoad_Matrix] ON [dbo].[ELoad_Matrix]([sMatrixID], [sStockIDx]) ON [PRIMARY]
GO

 CREATE  UNIQUE  INDEX [IX_Model] ON [dbo].[Model]([sModelIDx]) ON [PRIMARY]
GO

 CREATE  UNIQUE  INDEX [IX_PO_Receivig_Detail] ON [dbo].[PO_Receiving_Detail]([sTransNox], [nEntryNox]) ON [PRIMARY]
GO

 CREATE  UNIQUE  INDEX [IX_PO_Receiving_Serial] ON [dbo].[PO_Receiving_Serial]([sTransNox], [nEntryNox], [sSerialID]) ON [PRIMARY]
GO

 CREATE  UNIQUE  INDEX [IX_Province] ON [dbo].[Province]([sProvIDxx]) ON [PRIMARY]
GO

 CREATE  UNIQUE  INDEX [IX_Supplier] ON [dbo].[Supplier2]([sClientID], [sBranchCd]) ON [PRIMARY]
GO

 CREATE  UNIQUE  INDEX [IX_TownCity] ON [dbo].[TownCity]([sTownIDxx]) ON [PRIMARY]
GO

 CREATE  UNIQUE  INDEX [IX_xxxAppObject] ON [dbo].[xxxAppObject]([sProdctId]) ON [PRIMARY]
GO

 CREATE  UNIQUE  INDEX [IX_xxxDeletedRec] ON [dbo].[xxxDeletedRec]([sTransNox]) ON [PRIMARY]
GO

 CREATE  UNIQUE  INDEX [IX_xxxMenuObject] ON [dbo].[xxxMenuObject]([sMenuIDxx]) ON [PRIMARY]
GO

 CREATE  UNIQUE  INDEX [IX_xxxReport] ON [dbo].[xxxReport]([sReportID]) ON [PRIMARY]
GO

 CREATE  UNIQUE  INDEX [IX_xxxReportDetail] ON [dbo].[xxxReportDetail]([sReportID], [nEntryNox]) ON [PRIMARY]
GO

 CREATE  UNIQUE  INDEX [IX_xxxReportMaster] ON [dbo].[xxxReportMaster]([sReportID]) ON [PRIMARY]
GO

 CREATE  UNIQUE  INDEX [IX_xxxReportsLog] ON [dbo].[xxxReportsLog]([sRepFName]) ON [PRIMARY]
GO

 CREATE  UNIQUE  INDEX [IX_xxxSkin] ON [dbo].[xxxSkin]([sSkinCode]) ON [PRIMARY]
GO

 CREATE  UNIQUE  INDEX [IX_xxxSysObject] ON [dbo].[xxxSysObject]([sClientID], [sProdctID]) ON [PRIMARY]
GO

 CREATE  UNIQUE  INDEX [IX_xxxSysUser] ON [dbo].[xxxSysUser]([sUserIDxx]) ON [PRIMARY]
GO

 CREATE  UNIQUE  INDEX [IX_xxxTransactionSource] ON [dbo].[xxxTransactionSource]([sSourceID]) ON [PRIMARY]
GO

