VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function PrintTrans() As Boolean
   Dim CRXSubreport As Report
   Dim CRXSections As Sections
   Dim CRXSection As Section
   Dim CRXSubreportObj As SubreportObject
   Dim CRXReportObjects As ReportObjects
   Dim CRXReportObject As Object

   Dim lrs As ADODB.Recordset
   Dim lors As ADODB.Recordset
   Dim lnCtr As Integer
   Dim lsOldProc As String
   Dim loSubReport As Report
   
   lsOldProc = "PrintTrans"
   On Error GoTo errProc

   PrintTrans = True
   
   Set lrs = New ADODB.Recordset

   lrs.Fields.Append "nField01", adInteger, 3
   lrs.Fields.Append "sField01", adVarChar, 10
   lrs.Fields.Append "sField02", adVarChar, 25
   lrs.Fields.Append "sField03", adVarChar, 100
   lrs.Open
   
   With oTrans
      For lnCtr = 0 To .ItemCount - 1
         lrs.AddNew
         lrs("sField02").Value = .Detail(lnCtr, "sBarrCode")
         lrs("sField03").Value = .Detail(lnCtr, "sDescript")
         lrs("sField04").Value = .Detail(lnCtr, "sBrandNme")
         lrs("sField05").Value = .Detail(lnCtr, "sSerialNo")
         lrs("nField01").Value = .Detail(lnCtr, "nQuantity")
      Next
   End With
   
   
   Set CRXSections = p_oReport.Sections
   For Each CRXSection In CRXSections
      Set CRXReportObjects = CRXSection.ReportObjects
      For Each CRXReportObject In CRXReportObjects
         If CRXReportObject.Kind = crSubreportObject Then
            Set CRXSubreportObj = CRXReportObject
            Set CRXSubreport = CRXSubreportObj.OpenSubreport
            Select Case CRXSubreportObj.Name
            Case "SubReceipt"
               p_oReport.Sections("Da").Suppress = Not prcReceiptPayment
               If p_oReport.Sections("Da").Suppress Then openSource
               CRXSubreport.Database.SetDataSource p_oRepSource
            Case "SubSpareparts"
               p_oReport.Sections("Db").Suppress = Not prcSpareparts
               If p_oReport.Sections("Db").Suppress Then openSource
               CRXSubreport.Database.SetDataSource p_oRepSource
            End Select
         End If
      Next CRXReportObject
   Next CRXSection
   
   'assign important info to the report
   Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\CPPurchaseReturnForm.rpt")
   oReport.DiscardSavedData
   oReport.FieldMappingType = crAutoFieldMapping
   oReport.Database.SetDataSource lrs

   Set lors = New ADODB.Recordset
   If lors.State = adStateOpen Then lors.Close
   
   lors.Open "SELECT" _
               & "  CONCAT(a.sAddressx, ', ', b.sTownName, ', ', c.sProvName, ' ', b.sZippCode) as Address" _
               & ", a.sCompnyNm" _
            & " From Client_Master a" _
               & ", TownCity b" _
               & ", Province c" _
            & " WHERE a.sClientID = " & strParm(oTrans.Master("sSupplier")) _
               & " AND a.sTownIDxx = b.sTownIDxx" _
               & " AND b.sProvIDxx = c.sProvIDxx" _
            , oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   
   oReport.Sections("RH").ReportObjects("txtRefNo").SetText "MC" & "-" & Right(oTrans.Master("sTransNox"), 8)
   oReport.Sections("RH").ReportObjects("txtDate").SetText txtField(1).Text
   oReport.Sections("PH").ReportObjects("txtTo").SetText lors("sCompnyNm")
   oReport.Sections("PH").ReportObjects("txtToAddress").SetText lors("Address")
   oReport.Sections("PH").ReportObjects("txtFrom").SetText txtField(2)
   oReport.Sections("PH").ReportObjects("txtFromAddress").SetText oApp.Address & ", " & oApp.TownCity & ", " & oApp.Province & " " & oApp.ZippCode
   oReport.Sections("PF").ReportObjects("txtPrepared").SetText oApp.UserName
   oReport.Sections("PF").ReportObjects("txtRemarks").SetText txtField(6).Text
   
   oReport.PrintOutEx False, 1
   lors.Close

endPoc:
   oTrans.CloseTransaction (oTrans.Master(0))
   Set oReport = Nothing
   Set lrs = Nothing
   Set lors = Nothing
   Exit Function
errProc:
   PrintTrans = False
   ShowError lsOldProc & "( " & " )"
End Function


