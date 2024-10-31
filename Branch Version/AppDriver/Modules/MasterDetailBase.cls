VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MasterDetailBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Rex S. Adversalo
' XerSys Computing
' Canaoalan, Binmaley, Pangasinan
'
' Copyright 2003-2004 and beyond
' All Rights Reserved
'
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
' €  All  rights reserved. No part of this  software  €€  This Software is Owned by        €
' €  may be reproduced or transmitted in any form or  €€                                   €
' €  by   any   means,  electronic   or  mechanical,  €€    GUANZON MERCHANDISING CORP.    €
' €  including recording, or by information  storage  €€     Guanzon Bldg. Perez Blvd.     €
' €  and  retrieval  systems, without  prior written  €€           Dagupan City            €
' €  from the author.                                 €€  Tel No. 522-1085 ; 522-0863      €
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
'
' ==========================================================================================
'  XerSys [ 07/27/2004 02:31 am ]
'     Start creating this object.
'     This object is will be the base class of all transaction objects. This will handle
'        the following:
'           Retrieving transaction
'           Saving transaction
'           Posting transaction
'           Canceling transaction
'           Deleting transaction
'           Searching transaction
' ------------------------------------------------------------------------------------------
'  XerSys [ 08/13/2004 05:05 pm ]
'     Oh boy, its over 2 weeks since i update this object, nevertheless, let's begin
'  XerSys [ 09/02/2004 10:28 am ]
'     Oh boy, its almost a month since i update this object, nevertheless, let's get the
'        job done.
'  XerSys [ 02/07/2006 11:05 am ]
'     Override the aggreed modification policy...
'        Modification Policy
'           Encoder     = Add Only
'           Supervisor  = Add, modify and cancel transaction keyed a day before up to now
'           Manager     = Add, modify and cancel transaction keyed a month befor up to now
'           Audit       = view transaction
'           SysAdmin    = approved modification of transaction
'           Sys Owner   = Add, modify, cancel, delete transaction
'           Engineer    = same as owner
'     Modification policy would be manually set depending on the needs of a certain object.
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Explicit

Private p_oAppDrivr As AppDriver
Private p_oMaster As Recordset
Private p_oDetail As Recordset
Private p_oMod As New MainModules

Private p_sMaster As String
Private p_sDetail As String
Private p_sMasSQL As String
Private p_sLgrSQL As String
Private p_sRefFld As String

Private p_asLstVal() As String
Private p_asLstSQL() As String
Private p_asLstCol() As String
Private p_asLstCri() As String
Private p_asLstTit() As String
Private p_asLstPic() As String

Private p_asLgrLstSQL() As String
Private p_asLgrLstArg() As String
Private p_asLgrLstCol() As String
Private p_asLgrLstCri() As String
Private p_asLgrLstTit() As String
Private p_asLgrLstPic() As String

Private p_asMasRefFld() As String
Private p_asLgrRefFld() As String

Private pnEditMode As xeEditMode
Private pbRecExist As Boolean
Private pbInitTran As Boolean
Private pnCtr As Long

Public Event WillDelete(Cancel As Boolean)
Public Event Delete(Deleted As Boolean)
Public Event DeleteComplete()

Public Event WillSave(Cancel As Boolean)
Public Event Save(Saved As Boolean)
Public Event SaveComplete()

Public Event WillCancel(Cancel As Boolean)
Public Event Cancel(Cancelled As Boolean)
Public Event CancelComplete()

Public Event LoadOthers()
Public Event InitValue()

Public Event DetailRetrieved(ByVal Index As Integer)
Public Event MasterRetrieved(ByVal Index As Integer)

Property Set AppDriver(oAppDriver As AppDriver)
10       Set p_oAppDrivr = oAppDriver
End Property

Property Get MasterTable() As String
10       MasterTable = p_sMaster
End Property

Property Let MasterTable(Value As String)
10       p_sMaster = Value
End Property

Property Get MasterQuery() As String
10       MasterQuery = p_sMasSQL
End Property

Property Let MasterQuery(Value As String)
10       p_sMasSQL = Value
End Property

Property Get DetailTable() As String
10       DetailTable = p_sDetail
End Property

Property Let DetailTable(Value As String)
10       p_sDetail = Value
End Property

Property Get DetailQuery() As String
10       DetailQuery = p_sLgrSQL
End Property

Property Let DetailQuery(Value As String)
10       p_sLgrSQL = Value
End Property

Property Get BrowseDetailQuery(ByVal Index As Integer) As String
10       BrowseDetailQuery = p_asLgrLstSQL(Index)
End Property

Property Let BrowseDetailQuery(ByVal Index As Integer, ByVal Value As String)
10       p_asLgrLstSQL(Index) = Value
End Property

Property Get BrowseDetailColumn(ByVal Index As Integer) As String
10       BrowseDetailColumn = p_asLgrLstCol(Index)
End Property

Property Let BrowseDetailColumn(ByVal Index As Integer, ByVal Value As String)
10       p_asLgrLstCol(Index) = Value
End Property

Property Get BrowseDetailCriteria(ByVal Index As Integer) As String
10       BrowseDetailCriteria = p_asLgrLstCri(Index)
End Property

Property Let BrowseDetailCriteria(ByVal Index As Integer, ByVal Value As String)
10       p_asLgrLstCri(Index) = Value
End Property

Property Get BrowseDetailTitle(ByVal Index As Integer) As String
10       BrowseDetailTitle = p_asLgrLstTit(Index)
End Property

Property Let BrowseDetailTitle(ByVal Index As Integer, ByVal Value As String)
10       p_asLgrLstTit(Index) = Value
End Property

Property Get BrowseDetailFormat(ByVal Index As Integer) As String
10       BrowseDetailFormat = p_asLgrLstPic(Index)
End Property

Property Let BrowseDetailFormat(ByVal Index As Integer, ByVal Value As String)
10       p_asLgrLstPic(Index) = Value
End Property

Property Get BrowseQuery(ByVal Index As Integer) As String
10       BrowseQuery = p_asLstSQL(Index)
End Property

Property Let BrowseQuery(ByVal Index As Integer, ByVal Value As String)
10       p_asLstSQL(Index) = Value
End Property

Property Get BrowseColumn(ByVal Index As Integer) As String
10       BrowseColumn = p_asLstCol(Index)
End Property

Property Let BrowseColumn(ByVal Index As Integer, ByVal Value As String)
10       p_asLstCol(Index) = Value
End Property

Property Get BrowseCriteria(ByVal Index As Integer) As String
10       BrowseCriteria = p_asLstCri(Index)
End Property

Property Let BrowseCriteria(ByVal Index As Integer, ByVal Value As String)
10       p_asLstCri(Index) = Value
End Property

Property Get BrowseTitle(ByVal Index As Integer) As String
10       BrowseTitle = p_asLstTit(Index)
End Property

Property Let BrowseTitle(ByVal Index As Integer, ByVal Value As String)
10       p_asLstTit(Index) = Value
End Property

Property Get BrowseFormat(ByVal Index As Integer) As String
10       BrowseFormat = p_asLstPic(Index)
End Property

Property Let BrowseFormat(ByVal Index As Integer, ByVal Value As String)
10       p_asLstPic(Index) = Value
End Property

Property Let Reference(Value As String)
10       p_sRefFld = Value
End Property

Property Get Detail(ByVal Row As Long, ByVal Index As Variant) As Variant
10       On Error Resume Next
   
20       If pbInitTran = False Then Exit Property
30       If Row > (p_oDetail.RecordCount - 1) Then Exit Property
   
40       p_oDetail.Move Row, adBookmarkFirst
50       If IsNumeric(Index) Then
60          If Index = 0 Or p_asLgrLstSQL(Index) = Empty Then
70             Detail = p_oDetail(Index).Value
80          Else
90             Detail = getDetDesc(Row, Index, p_oDetail(Index))
100         End If
110      Else
120         Detail = p_oDetail(Index).Value
130      End If
End Property

Property Let Detail(ByVal Row As Long, ByVal Index As Variant, Value As Variant)
10       On Error Resume Next
   
20       If pbInitTran = False Then Exit Property
30       If Row > (p_oDetail.RecordCount - 1) Then Exit Property
   
40       p_oDetail.Move Row, adBookmarkFirst
50       If IsNumeric(Index) Then
60          If Index = 0 Or p_asLgrLstSQL(Index) = Empty Then
70             p_oDetail(Index) = Value
80          Else
90             getDetCode Row, Index, Value, False
100         End If
110      Else
120         p_oDetail(Index) = Value
130      End If
End Property

Property Get Master(Index As Variant) As Variant
10       If pbInitTran = False Then Exit Property
20       If p_oMaster.EOF Then Exit Property
   
30       If IsNumeric(Index) Then
40          If Index = 0 Or p_asLstSQL(Index) = Empty Then
50             Master = p_oMaster(Index)
60          Else
70             Master = p_asLstVal(Index)
80          End If
90       Else
100         Master = p_oMaster(Index)
110      End If
End Property

Property Let Master(Index As Variant, Value As Variant)
10       If pbInitTran = False Then Exit Property
20       If p_oMaster.EOF Then Exit Property
   
30       If IsNumeric(Index) Then
40          If Index = 0 Or p_asLstSQL(Index) = Empty Then
50             p_oMaster(Index) = Value
60          Else
70             getMasCode Index, Value, False
80          End If
90       Else
100          p_oMaster(Index) = Value
110      End If
End Property

Property Get MasterValue(Index As Variant) As Variant
10       If pbInitTran = False Then Exit Property
20       If p_oMaster.EOF Then Exit Property
   
30       MasterValue = p_oMaster(Index)
End Property

Property Get MasterOriginal(Index As Variant) As Variant
10       If pbInitTran = False Then Exit Property
20       If p_oMaster.EOF Then Exit Property
   
30       MasterOriginal = p_oMaster(Index).OriginalValue
End Property

Property Get ItemCount() As Long
10       If pbInitTran = False Then Exit Property
   
20       ItemCount = p_oDetail.RecordCount
End Property

Property Let EditMode(Value As xeEditMode)
10       pnEditMode = Value
End Property

Property Get EditMode() As xeEditMode
10       EditMode = pnEditMode
End Property

Function InitTransaction() As Boolean
10       Dim lasFieldRef() As String
20       Dim lsOldProc As String
   
30       lsOldProc = p_oAppDrivr.ProcName("InitTransaction")
40       On Error GoTo errProc
50       InitTransaction = False

   ' first check the application driver
60       If p_oMod.isAppDriverOK(p_oAppDrivr) = False Then GoTo endProc
70       If p_sMaster = Empty Or p_sMasSQL = Empty Or _
         p_sDetail = Empty Or p_sLgrSQL = Empty Or _
         p_sRefFld = Empty Then GoTo endProc
      
   ' then check the reference field
80       lasFieldRef = Split(p_sRefFld, "|")
90       If UBound(lasFieldRef) <> 1 Then GoTo endProc
   
100      p_asMasRefFld = Split(lasFieldRef(0), "»")
110      p_asLgrRefFld = Split(lasFieldRef(1), "»")

120      If UBound(p_asMasRefFld) <> UBound(p_asLgrRefFld) Then GoTo endProc
   
   ' now initialize the receiving recordset
130      clearMaster
   
   ' after initializing the master, set master properties
140      ReDim p_asLstVal(p_oMaster.Fields.Count - 1) As String
150      ReDim p_asLstSQL(p_oMaster.Fields.Count - 1) As String
160      ReDim p_asLstCol(p_oMaster.Fields.Count - 1) As String
170      ReDim p_asLstTit(p_oMaster.Fields.Count - 1) As String
180      ReDim p_asLstPic(p_oMaster.Fields.Count - 1) As String
190      ReDim p_asLstCri(p_oMaster.Fields.Count - 1) As String
   
   ' after initializing the detail, set detail properties
200      clearDetail
210      ReDim p_asLgrLstSQL(p_oDetail.Fields.Count - 1) As String
220      ReDim p_asLgrLstCol(p_oDetail.Fields.Count - 1) As String
230      ReDim p_asLgrLstTit(p_oDetail.Fields.Count - 1) As String
240      ReDim p_asLgrLstPic(p_oDetail.Fields.Count - 1) As String
250      ReDim p_asLgrLstCri(p_oDetail.Fields.Count - 1) As String
   
260      InitTransaction = True
270      pbInitTran = True
280      pnEditMode = xeModeReady
   
endProc:
290      p_oAppDrivr.ProcName lsOldProc
300      Exit Function
errProc:
310      ShowError lsOldProc
End Function

Function NewTransaction() As Boolean
10       Dim lsOldProc As String
   
20       lsOldProc = p_oAppDrivr.ProcName("NewTransaction")
30       On Error GoTo errProc
40       NewTransaction = False

50       If pbInitTran = False Then GoTo endProc
   
60       clearMaster
70       For pnCtr = 0 To UBound(p_asLstVal)
80          p_asLstVal(pnCtr) = ""
90       Next
   
100      clearDetail
    
110      pnEditMode = xeModeAddNew
120      NewTransaction = True
130      RaiseEvent InitValue
'110      RaiseEvent InitDetail(0)
   
endProc:
140      p_oAppDrivr.ProcName lsOldProc
150      Exit Function
errProc:
160      ShowError lsOldProc
End Function

Function OpenTransaction(ByVal sReference As String) As Boolean
10       Dim lsReference() As String
20       Dim lsOldProc As String
30       Dim lsSQL As String
40       Dim Cancel As Boolean
   
50       lsOldProc = p_oAppDrivr.ProcName("OpenTransaction")
60       On Error GoTo errProc
70       OpenTransaction = False

80       If pbInitTran = False Then GoTo endProc
   
90       lsReference = Split(sReference, "»")
100      If UBound(lsReference) <> UBound(p_asMasRefFld) Then GoTo endProc
   
110      With p_oAppDrivr
120         If p_oMaster.State = adStateOpen Then
130            If Not p_oMaster.EOF Then
140               If p_oMaster.EditMode <> adEditNone Then p_oMaster.CancelUpdate
150            End If
160            p_oMaster.Close
170         End If
      
      ' set the select statement for the master
180         lsSQL = p_sMasSQL
190         For pnCtr = 0 To UBound(lsReference)
200            If lsReference(pnCtr) <> Empty Then
210               lsSQL = p_oMod.AddCondition(lsSQL, p_asMasRefFld(pnCtr) & _
                     " = " & p_oMod.strParm(lsReference(pnCtr)))
220            End If
230         Next
240         p_oMaster.Open lsSQL, .Connection, , adLockOptimistic, adCmdText
   
'      If p_oDetail.State = adStateOpen Then
'         If Not p_oDetail.EOF Then
'            If p_oDetail.EditMode <> adEditNone Then p_oDetail.CancelUpdate
'         End If
'         p_oDetail.Close
'      End If
250         Set p_oDetail = New Recordset
      ' set the select statement for the ledger
260         lsSQL = p_sLgrSQL
270         For pnCtr = 0 To UBound(lsReference)
280            lsSQL = p_oMod.AddCondition(lsSQL, p_asLgrRefFld(pnCtr) & _
                  " = " & p_oMod.strParm(lsReference(pnCtr)))
290         Next
300         p_oDetail.Open lsSQL, .Connection, , adLockOptimistic, adCmdText
      
      ' joined table carry their individual table constraint, thus, if
      '  new record is added all constraints must be satisfied. With regards to
      '  this, a constraint violation occur, and update to the record will occur
      '  immediately. Thus an error will occur. Now deactive the connection in
      '  order to set the recordset to batch update form.
310         Set p_oDetail.ActiveConnection = Nothing
320         If p_oDetail.EOF Then
330            p_oDetail.AddNew
'340            RaiseEvent InitDetail(0)
340         End If
350      End With
   
   ' now verify if there is a record retrieved
360      If p_oMaster.EOF = True And p_oMaster.BOF = True Then GoTo endProc
   
370      LoadDescription
380      OpenTransaction = True
390      pbRecExist = True
400      pnEditMode = xeModeReady
   
410      RaiseEvent LoadOthers
endProc:
420      p_oAppDrivr.ProcName lsOldProc
430      Exit Function
errProc:
440      ShowError lsOldProc
End Function

Function DeleteTransaction() As Boolean
10       Dim lsMasSQL As String
20       Dim lsDetSQL As String
30       Dim lsField As String
40       Dim lsOldProc As String
50       Dim lnCol As Long
60       Dim Cancel As Boolean
   
70       lsOldProc = p_oAppDrivr.ProcName("DeleteTransaction")
80       On Error GoTo errProc
90       DeleteTransaction = False

100      If pbInitTran = False Then GoTo endProc
   
110      If MsgBox("Are you sure you want to Delete this Transaction?", _
                  vbQuestion + vbYesNo, "Confirm Cancel") <> vbYes Then GoTo endProc
   
120      If HasRights2Delete = False Then GoTo endProc

130      If pnEditMode <> xeModeReady Then
140         MsgBox "Deletion of unsave transaction is not allowed!!!" & vbCrLf & vbCrLf & _
               "Request can not be granted!!!", vbCritical, "Warning"
150         GoTo endProc
160      End If
    
170      If pbRecExist = False Then
180         MsgBox "No transaction is available to delete!!!", vbCritical, "Warning"
190         GoTo endProc
200      End If
   
210      If isPosted Then
220         MsgBox "Deletion of posted or cancelled transaction is not allowed!!!" & vbCrLf & vbCrLf & _
               "Request not granted!!!", vbCritical, "Warning"
230         GoTo endProc
240      End If
   
250      p_oAppDrivr.Connection.BeginTrans
   
   ' raise willdelete for deletion validation and related record deletion
260      Cancel = False
270      RaiseEvent WillDelete(Cancel)
280      If Cancel Then
290         p_oAppDrivr.Connection.RollbackTrans
300         GoTo endProc
310      End If
   
320      RaiseEvent Delete(Cancel)
330      If Cancel Then
340         p_oAppDrivr.Connection.CommitTrans
350         NewTransaction
360         pbRecExist = False
370         pnEditMode = xeModeReady
380         DeleteTransaction = True
      
390         RaiseEvent DeleteComplete
400         GoTo endProc
410      End If
   
420      lsMasSQL = "DELETE " & p_sMaster
430      lsDetSQL = "DELETE " & p_sDetail
440      For pnCtr = 0 To UBound(p_asMasRefFld)
450         lsField = p_asMasRefFld(pnCtr)
460         lnCol = InStr(lsField, ".")
470         If lnCol > 0 Then lsField = Mid(lsField, lnCol + 1)
480         lsMasSQL = p_oMod.AddCondition(lsMasSQL, lsField & " = " & _
                        p_oMod.strParm(p_oMaster(lsField).OriginalValue))
                        
490         lsField = p_asLgrRefFld(pnCtr)
500         lnCol = InStr(lsField, ".")
510         If lnCol > 0 Then lsField = Mid(lsField, lnCol + 1)
520         lsDetSQL = p_oMod.AddCondition(lsDetSQL, lsField & " = " & _
                        p_oMod.strParm(p_oMaster(lsField).OriginalValue))
530      Next
   
540      With p_oAppDrivr
550         If p_oAppDrivr.RegisDelete(lsMasSQL) = False Then GoTo endProc
      
560         .Connection.Execute lsMasSQL, pnCtr, adCmdText
   
570         If pnCtr <= 0 Then
580            Err.Raise 32010, "Recordset", "Unable to delete record"
590            GoTo endProc
600         End If
      
610         If p_oAppDrivr.RegisDelete(lsDetSQL) = False Then GoTo endProc
620         .Connection.Execute lsDetSQL, pnCtr, adCmdText
      
630         If pnCtr < 0 Then
640            Err.Raise 32010, "Recordset", "Unable to delete record"
650            GoTo endProc
660         End If
      
670         .Connection.CommitTrans
680      End With
   
   ' after deleting transaction add new transaction
690      NewTransaction
700      DeleteTransaction = True
710      pbRecExist = False
720      pnEditMode = xeModeReady
   
730      RaiseEvent DeleteComplete
endProc:
740      p_oAppDrivr.ProcName lsOldProc
750      Exit Function
errProc:
760      ShowError lsOldProc
End Function

Function CancelTransaction() As Boolean
10       Dim lsOldProc As String
20       Dim lsField As String
30       Dim lsSQL As String
40       Dim lnCol As Long
50       Dim Cancel As Boolean
   
60       lsOldProc = p_oAppDrivr.ProcName("CancelTransaction")
70       On Error GoTo errProc
80       CancelTransaction = False

90       If pbInitTran = False Then GoTo endProc
   
100      If MsgBox("Are you sure you want to Cancel this Transaction?", _
                  vbQuestion + vbYesNo, "Confirm Cancel") <> vbYes Then GoTo endProc
   
110      If HasRights2Cancel = False Then GoTo endProc
   
120      If pnEditMode <> xeModeReady Then
130         MsgBox "Cancellation of unsave transaction is not allowed!!!" & vbCrLf & vbCrLf & _
               "Request can not be granted!!!", vbCritical, "Warning"
140         GoTo endProc
150      End If
    
160      If pbRecExist = False Then
170         MsgBox "No transaction is available to cancel!!!", vbCritical, "Warning"
180         GoTo endProc
190      End If
   
200      With p_oAppDrivr
210         .Connection.BeginTrans
   
      ' raise willdelete for deletion validation and related record deletion
220         Cancel = False
230         RaiseEvent WillCancel(Cancel)
240         If Cancel Then
250            .Connection.RollbackTrans
260            GoTo endProc
270         End If
      
280         RaiseEvent Cancel(Cancel)
290         If Cancel Then
300            .Connection.CommitTrans
310            pnEditMode = xeModeReady
320            CancelTransaction = True
         
330            RaiseEvent CancelComplete
340            GoTo endProc
350         End If
      
360         lsSQL = "UPDATE " & p_sMaster & _
               " SET cTranStat = " & xeStateCancelled & _
                  ", dModified = GETDATE()"
      
370         For pnCtr = 0 To UBound(p_asMasRefFld)
380            lsField = p_asMasRefFld(pnCtr)
390            lnCol = InStr(lsField, ".")
400            If lnCol > 0 Then lsField = Mid(lsField, lnCol + 1)
410            lsSQL = p_oMod.AddCondition(lsSQL, lsField & " = " & _
                        p_oMod.strParm(p_oMaster(lsField)))
420         Next

430         .Connection.Execute lsSQL, pnCtr, adCmdText
      
440         If pnCtr <= 0 Then
450            Err.Raise 32011, "Recordset", "Unable to update changes"
460            GoTo endProc
470         End If
   
480         p_oMaster("cTranStat") = xeStateCancelled
490         .Connection.CommitTrans
500      End With
   
   ' after deleting transaction add new transaction
510      CancelTransaction = True
520      pnEditMode = xeModeReady
   
530      RaiseEvent CancelComplete
endProc:
540      p_oAppDrivr.ProcName lsOldProc
550      Exit Function
errProc:
560      ShowError lsOldProc
End Function

Function SaveTransaction() As Boolean
10       Dim loMaster As Recordset
20       Dim loDetail As Recordset
30       Dim lsOldProc As String
40       Dim lsMasSQL As String
50       Dim lsDetSQL As String
60       Dim lsField As String
70       Dim lnCol As Integer
80       Dim dModified As Date
90       Dim Cancel As Boolean
   
100      lsOldProc = p_oAppDrivr.ProcName("SaveTransaction")
110      SaveTransaction = False
120      On Error GoTo errProc

130      If pbInitTran = False Then GoTo endProc
140      If pnEditMode = xeModeReady Then GoTo endProc

150      If MsgBox("Are you sure you want to Save this Transaction?", _
         vbYesNo + vbInformation, "Confirm") <> vbYes Then GoTo endProc
         
160      If HasRights2Save = False Then GoTo endProc
   
170      If isRefEmpty Then
180         MsgBox "Reference Fields can not Contain Empty Value!!!" & vbCrLf & _
                     " Verify your Entry then Try Again!!!", vbCritical, "Warning"
190         GoTo endProc
200      End If

210      dModified = p_oAppDrivr.ServerDate
220      p_oAppDrivr.Connection.BeginTrans
   
230      Cancel = False
   ' record validation must be in this event
240      RaiseEvent WillSave(Cancel)
250      If Cancel = True Then
260         p_oAppDrivr.Connection.RollbackTrans
270         GoTo endProc
280      End If

   ' Comment out the following code to lessen processing time in saving
   '     record. Allow the user of the class to check the modification of
   '     the transaction.
'190      If isDetailModify = False And isMasterModify = False Then
'200         SaveTransaction = True
'210         pnEditMode = xeModeReady
'
'      ' after saving the main table, trigger the savecomplete event
'220         RaiseEvent SaveComplete
'230         GoTo endProc
'240      End If
         
290      Cancel = False
300      RaiseEvent Save(Cancel)
   
   ' if record is saved by the client then endproc
310      If Cancel Then
320         p_oAppDrivr.Connection.CommitTrans
      
330         RaiseEvent SaveComplete
340         SaveTransaction = Cancel
350         pbRecExist = True
360         pnEditMode = xeModeReady
370         GoTo endProc
380      End If
   
390      p_oMaster("sModified") = p_oMod.Encrypt(p_oAppDrivr.UserID)
400      p_oMaster("dModified") = dModified
   
   ' detail table must be re-opened to include the actual fields only
410      Set loMaster = New Recordset
420      Set loDetail = New Recordset
   
430      lsMasSQL = "SELECT * FROM " & p_sMaster
440      lsDetSQL = "SELECT * FROM " & p_sDetail & " ORDER BY nEntryNox"
   
450      If pnEditMode = xeModeAddNew Then
460         lsMasSQL = p_oMod.AddCondition(lsMasSQL, "0 = 1")
470         lsDetSQL = p_oMod.AddCondition(lsDetSQL, "0 = 1")
      
480         loMaster.Open lsMasSQL, p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
490         loMaster.AddNew
500         loDetail.Open lsDetSQL, p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
510      Else
      ' set the retrieval arguments of the recordset
520         For pnCtr = 0 To UBound(p_asLgrRefFld)
530            lsField = p_asLgrRefFld(pnCtr)
540            lnCol = InStr(lsField, ".")
550            If lnCol > 0 Then lsField = Mid(lsField, lnCol + 1)
560            lsDetSQL = p_oMod.AddCondition(lsDetSQL, lsField & " = " & _
                     p_oMod.strParm(p_oMaster(lsField).OriginalValue))
      
570            lsField = p_asMasRefFld(pnCtr)
580            lnCol = InStr(lsField, ".")
590            If lnCol > 0 Then lsField = Mid(lsField, lnCol + 1)
600            lsMasSQL = p_oMod.AddCondition(lsMasSQL, lsField & " = " & _
                     p_oMod.strParm(p_oMaster(lsField).OriginalValue))
610         Next
620         loMaster.Open lsMasSQL, p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
630         loDetail.Open lsDetSQL, p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
640      End If
   
650      lsField = p_asLgrRefFld(0)
660      lnCol = InStr(lsField, ".")
670      If lnCol > 0 Then lsField = Mid(lsField, lnCol + 1)
   
   ' now move the content of the detail to the receiving recordset
680      If p_oDetail.RecordCount > 0 Then p_oDetail.Move 0, adBookmarkFirst
690      For pnCtr = 0 To p_oDetail.RecordCount - 1
700         If p_oDetail(lsField) <> Empty Then
710            If loDetail.EOF Then loDetail.AddNew
      
         ' subtract 3 columns to exclude the TimeStamp column
720            For lnCol = 0 To loDetail.Fields.Count - 3
730               lsMasSQL = loDetail(lnCol).Name
740               loDetail(lsMasSQL) = p_oDetail(lsMasSQL)
750            Next
760            p_oDetail("dModified") = dModified
770            loDetail("dModified") = dModified
780            loDetail.MoveNext
790         End If
800         p_oDetail.MoveNext
810      Next
   
820      Do While Not loDetail.EOF
830         lsDetSQL = "DELETE " & DetailTable
840         For pnCtr = 0 To loDetail.Fields.Count - 3
850            lsMasSQL = loDetail(lnCol).Name
860            lsDetSQL = p_oMod.AddCondition(lsDetSQL, lsMasSQL & " = " & _
                        p_oMod.strParm(loDetail(lsMasSQL)))
870         Next
      
880         If p_oAppDrivr.RegisDelete(lsDetSQL) = False Then GoTo endProc
      
890         loDetail.Delete adAffectCurrent
900         loDetail.MoveNext
910      Loop
   
   ' move the content of the master to the receiving recordset
920      For lnCol = 0 To loMaster.Fields.Count - 2
930        lsMasSQL = loMaster(lnCol).Name
940        loMaster(lsMasSQL) = p_oMaster(lsMasSQL)
950      Next
   
960      loMaster.UpdateBatch adAffectAllChapters
970      loDetail.UpdateBatch adAffectAllChapters
   
980      p_oAppDrivr.Connection.CommitTrans
   
990      SaveTransaction = True
1000     pnEditMode = xeModeReady
1010     pbRecExist = True

   ' after saving the main table, trigger the savecomplete event
1020     RaiseEvent SaveComplete
   
endProc:
1030     Set loDetail = Nothing
1040     Set loMaster = Nothing
1050     p_oAppDrivr.ProcName lsOldProc
1060     Exit Function
errProc:
1070     p_oAppDrivr.Connection.RollbackTrans
1080     ShowError lsOldProc
End Function

Function PostTransaction(ByVal sReference As String) As Boolean
10       Dim lsOldProc As String
20       Dim lasRef() As String
30       Dim lsField As String
40       Dim lnCol As Long
50       Dim lsSQL As String
   
60       lsOldProc = p_oAppDrivr.ProcName("PostTransaction")
70       PostTransaction = False
80       On Error GoTo errProc

90       If pbInitTran = False Then GoTo endProc
   
100      lasRef = Split(sReference, "»")
110      If UBound(lasRef) <> UBound(p_asMasRefFld) Then GoTo endProc
   
120      lsSQL = "UPDATE " & p_sMaster & _
            " SET cTranStat = " & xeStatePosted & _
               ", dModified = GETDATE()"
   
130      For pnCtr = 0 To UBound(p_asMasRefFld)
140         lsField = p_asMasRefFld(pnCtr)
150         lnCol = InStr(lsField, ".")
160         If lnCol > 0 Then lsField = Mid(lsField, lnCol + 1)
      
170         lsSQL = p_oMod.AddCondition(lsSQL, lsField & " = " & p_oMod.strParm(lasRef(pnCtr)))
180      Next
   
190      p_oAppDrivr.Connection.Execute lsSQL, pnCtr, adCmdText
   
200      If pnCtr <= 0 Then
210         Err.Raise 32011, "Recordset", "Unable to update changes"
220         GoTo endProc
230      End If
   
240      p_oMaster("cTranStat") = xeStatePosted
250      PostTransaction = True
   
endProc:
260      p_oAppDrivr.ProcName lsOldProc
270      Exit Function
errProc:
280      ShowError lsOldProc
End Function

Function CloseTransaction(ByVal sReference As String) As Boolean
10       Dim lasRef() As String
20       Dim lsOldProc As String
30       Dim lsSQL As String
40       Dim lsField As String
50       Dim lnCol As Long
   
60       lsOldProc = p_oAppDrivr.ProcName("CloseTransaction")
70       CloseTransaction = False
80       On Error GoTo errProc

90       If pbInitTran = False Then GoTo endProc
   
100      lasRef = Split(sReference, "»")
110      If UBound(lasRef) <> UBound(p_asMasRefFld) Then GoTo endProc
   
120      lsSQL = "UPDATE " & p_sMaster & _
            " SET cTranStat = " & xeStateClosed & _
               ", dModified = GETDATE()"
   
130      For pnCtr = 0 To UBound(p_asMasRefFld)
140         lsField = p_asMasRefFld(pnCtr)
150         lnCol = InStr(lsField, ".")
160         If lnCol > 0 Then lsField = Mid(lsField, lnCol + 1)
      
170         lsSQL = p_oMod.AddCondition(lsSQL, lsField & " = " & p_oMod.strParm(lasRef(pnCtr)))
180      Next
      
190      p_oAppDrivr.Connection.Execute lsSQL, pnCtr, adCmdText
      
200      If pnCtr = 0 Then
210         Err.Raise 32011, "Recordset", "Unable to update changes"
220         GoTo endProc
230      End If
240      p_oMaster("cTranStat") = xeStateClosed
   
250      CloseTransaction = True
   
endProc:
260      p_oAppDrivr.ProcName lsOldProc
270      Exit Function
errProc:
280      ShowError lsOldProc
End Function

Function SearchMaster(ByVal lnIndex As Integer, _
                        Optional lsValue As Variant = "") As Boolean
10       Dim lsOldProc As String
   
20       lsOldProc = p_oAppDrivr.ProcName("SearchTransaction")
30       SearchMaster = False
40       On Error GoTo errProc

50       If pbInitTran = False Then GoTo endProc
60       If p_oMaster.EOF Then GoTo endProc
   
70       If p_asLstSQL(lnIndex) <> Empty Then
80          getMasCode lnIndex, lsValue, True
90       End If
100      SearchMaster = True
   
endProc:
110      p_oAppDrivr.ProcName lsOldProc
120      Exit Function
errProc:
130      ShowError lsOldProc
End Function

Function SearchDetail(ByVal nRow As Integer, _
                        ByVal nIndex As Integer, _
                        Optional lsValue As Variant = "") As Boolean
10       Dim lsOldProc As String
   
20       lsOldProc = p_oAppDrivr.ProcName("SearchDetail")
30       SearchDetail = False
40       On Error GoTo errProc

50       If pbInitTran = False Then GoTo endProc
60       If p_oMaster.EOF Then GoTo endProc
   
70       If p_asLgrLstSQL(nIndex) <> Empty Then
80          p_oDetail.Move nRow, adBookmarkFirst
90          getDetCode nRow, nIndex, lsValue, True
100      End If
110      SearchDetail = True
   
endProc:
120      p_oAppDrivr.ProcName lsOldProc
130      Exit Function
errProc:
140      ShowError lsOldProc
End Function

Function SearchTransaction(ByVal sRefValue As String, ByVal sRefField As String, sDescript As String) As Boolean
10       Dim loLookUp As frmLookUp
20       Dim lors As Recordset
30       Dim lsOldProc As String
40       Dim lasValue() As String
50       Dim lasField() As String
60       Dim lsRef As String
70       Dim lsSQL As String

80       lsOldProc = p_oAppDrivr.ProcName("SearchTransaction")
90       SearchTransaction = False
100      On Error GoTo errProc

110      If pbInitTran = False Then GoTo endProc
   
120      If sRefValue = Empty Then
130         Set loLookUp = New frmLookUp
140         Set loLookUp.AppDriver = p_oAppDrivr
      
150         loLookUp.SQLSource = p_asLstSQL(0)
160         loLookUp.FldName = p_asLstCol(0)
170         loLookUp.FldTitle = p_asLstTit(0)
180         loLookUp.FldFormat = p_asLstPic(0)
190         loLookUp.FldCriteria = p_asLstCri(0)
200         loLookUp.showSearch = True
      
210         loLookUp.Show 1
220         lsRef = loLookUp.SelectedItem
   
230         If lsRef <> Empty Then
240            lasValue = Split(lsRef, "»")
         
250            lsRef = ""
260            For pnCtr = 0 To UBound(p_asMasRefFld)
270               lsRef = lsRef & lasValue(pnCtr) & "»"
280            Next
290            lsRef = Left(lsRef, Len(lsRef) - 1)
300         End If
310      Else
320         lasValue = Split(sRefValue, "»")
330         lasField = Split(sRefField, "»")
340         If UBound(lasValue) <> UBound(lasField) Then GoTo endProc
      
350         lsSQL = p_asLstSQL(0)
360         For pnCtr = 0 To UBound(lasField)
370            If lasValue(pnCtr) <> Empty Then
380               If lasField(pnCtr) = sDescript Then
390                  lsSQL = p_oMod.AddCondition(lsSQL, lasField(pnCtr) & " LIKE " & _
                        p_oMod.strParm(lasValue(pnCtr) & "%"))
400               Else
410                  lsSQL = p_oMod.AddCondition(lsSQL, lasField(pnCtr) & " = " & _
                        p_oMod.strParm(lasValue(pnCtr)))
420               End If
430            End If
440         Next
            
450         Set lors = New Recordset
460         lors.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockReadOnly, adCmdText

470         If lors.EOF Then
480            GoTo endProc
490         End If
      
500         If lors.RecordCount > 1 Then
510            Set loLookUp = New frmLookUp
520            Set loLookUp.AppDriver = p_oAppDrivr
      
530            Set loLookUp.RowSource = lors
540            loLookUp.FldName = p_asLstCol(0)
550            loLookUp.FldTitle = p_asLstTit(0)
560            loLookUp.FldFormat = p_asLstPic(0)
570            loLookUp.showSearch = False
         
580            loLookUp.Show 1
590            lsRef = loLookUp.SelectedItem
         
600            If lsRef <> Empty Then
610               lasValue = Split(lsRef, "»")
            
620               lsRef = ""
630               For pnCtr = 0 To UBound(p_asMasRefFld)
640                  lsRef = lsRef & lasValue(pnCtr) & "»"
650               Next
660               lsRef = Left(lsRef, Len(lsRef) - 1)
670            End If
680         Else
690            lsRef = ""
700            For pnCtr = 0 To UBound(p_asMasRefFld)
710               lsRef = lsRef & lors.Fields(pnCtr) & "»"
720            Next
730            lsRef = Left(lsRef, Len(lsRef) - 1)
740         End If
750      End If
   
760      If lsRef = Empty Then
770         GoTo endProc
780      End If
790      SearchTransaction = OpenTransaction(lsRef)
      
endProc:
800      Set loLookUp = Nothing
810      Set lors = Nothing

820      p_oAppDrivr.ProcName lsOldProc
830      Exit Function
errProc:
840      ShowError lsOldProc
End Function

Function ClearTransaction() As Boolean
10       Dim lsOldProc As String
   
20       lsOldProc = p_oAppDrivr.ProcName("ClearTransaction")
30       On Error GoTo errProc
40       ClearTransaction = False

50       If pbInitTran = False Then GoTo endProc
   
60       clearMaster
70       clearDetail
   
80       ClearTransaction = True
   
endProc:
90       p_oAppDrivr.ProcName lsOldProc
100      Exit Function
errProc:
110      ShowError lsOldProc
End Function

Function AddDetail() As Boolean
10       Dim lsOldProc As String
   
20       lsOldProc = p_oAppDrivr.ProcName("AddDetail")
30       On Error GoTo errProc
40       AddDetail = False

50       If pbInitTran = False Then GoTo endProc
60       p_oDetail.AddNew
70       AddDetail = True

endProc:
80       p_oAppDrivr.ProcName lsOldProc
90       Exit Function
errProc:
100      ShowError lsOldProc
End Function

Function DeleteDetail(Index As Long) As Boolean
10       Dim lsOldProc As String
   
20       lsOldProc = p_oAppDrivr.ProcName("DeleteDetail")
30       On Error GoTo errProc
40       DeleteDetail = False

50       If pbInitTran = False Then GoTo endProc
60       If p_oDetail.RecordCount - 1 < Index Then GoTo endProc
   
70       p_oDetail.Move Index, adBookmarkFirst
80       p_oDetail.Delete adAffectCurrent
90       DeleteDetail = True

endProc:
100      p_oAppDrivr.ProcName lsOldProc
110      Exit Function
errProc:
120      ShowError lsOldProc
End Function

Function GetMaster(ByVal Index As Integer, Condition As String, Optional Search As Variant = False) As String
10       Dim lors As Recordset
20       Dim lsOldProc As String
30       Dim lsSQL As String
40       Dim lasSelect() As String
   
50       lsOldProc = p_oAppDrivr.ProcName("GetMaster")
60       On Error GoTo errProc
70       GetMaster = ""

80       If p_oMaster.Fields.Count - 1 < Index Then GoTo endProc
90       lsSQL = p_oMod.AddCondition(p_asLstSQL(Index), Condition)
100      Set lors = New Recordset
         
110      lors.Open lsSQL, p_oAppDrivr.Connection, adOpenKeyset, , adCmdText
   
120      If lors.EOF Then GoTo endProc
130      If lors.RecordCount = 1 Then
140         If Search = True Then
150            GetMaster = p_oMod.KwikBrowse(p_oAppDrivr, lors, p_asLstCol(Index), p_asLstTit(Index), _
                     p_asLstPic(Index), , p_asLstCri(Index))
160         Else
170            For pnCtr = 0 To lors.Fields.Count - 1
180               GetMaster = GetMaster & IIf(IsNull(lors(pnCtr)), "", lors(pnCtr)) & "»"
190            Next
200            GetMaster = Left(GetMaster, Len(GetMaster) - 1)
210         End If
220      Else
230         GetMaster = p_oMod.KwikBrowse(p_oAppDrivr, lors, p_asLstCol(Index), p_asLstTit(Index), _
                     p_asLstPic(Index), , p_asLstCri(Index))
240      End If
   
endProc:
250      Set lors = Nothing
260      p_oAppDrivr.ProcName lsOldProc
270      Exit Function
errProc:
280      ShowError lsOldProc
End Function

Function GetDetail(ByVal Row As Integer, _
                     Condition As String, _
                     Optional Search As Variant = False) As String
10       Dim lors As Recordset
20       Dim lsOldProc As String
30       Dim lsSQL As String
40       Dim lasSelect() As String
   
50       lsOldProc = p_oAppDrivr.ProcName("GetDetail")
60       On Error GoTo errProc
70       GetDetail = ""

80       If p_oDetail.RecordCount - 1 < Row Then GoTo endProc
90       lsSQL = p_oMod.AddCondition(p_asLgrLstSQL(0), Condition)

100      Set lors = New Recordset
110      lors.Open lsSQL, p_oAppDrivr.Connection, adOpenKeyset, , adCmdText
   
120      If lors.EOF Then GoTo endProc
130      If lors.RecordCount = 1 Then
'      If Search = True Then
'         GetDetail = p_oMod.KwikBrowse(p_oAppDrivr, lors, p_asLgrLstCol(0), p_asLgrLstTit(0), p_asLgrLstPic(0), , p_asLgrLstCri(0))
'      Else
140            For pnCtr = 0 To lors.Fields.Count - 1
150               GetDetail = GetDetail & IIf(IsNull(lors(pnCtr)), "", lors(pnCtr)) & "»"
160            Next
170            GetDetail = Left(GetDetail, Len(GetDetail) - 1)
'      End If
180      Else
190         GetDetail = p_oMod.KwikBrowse(p_oAppDrivr, lors, p_asLgrLstCol(0), p_asLgrLstTit(0), p_asLgrLstPic(0), , p_asLgrLstCri(0))
200      End If
   
endProc:
210      Set lors = Nothing
220      p_oAppDrivr.ProcName lsOldProc
230      Exit Function
errProc:
240      ShowError lsOldProc
End Function

Function GetNextTrans(ByVal Field As String, ByVal Condition As String) As String
10       Dim lors As Recordset
20       Dim lsSQL As String
30       Dim lnCode As Long
40       Dim lnLen As Long

50       lsSQL = "SELECT TOP 1 " & Field & _
      " FROM " & p_sMaster & _
      " WHERE " & Condition & _
      " ORDER BY " & Field & " DESC"
      
60       Set lors = New Recordset
70       lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText

80       lsSQL = IIf(lors.EOF, Empty, lors(0))
90       lnLen = lors(0).DefinedSize
100      lnCode = 1
110      If lsSQL <> Empty Then
120         If Left(lsSQL, 2) = Format(Date, "YY") Then
130            lnCode = CLng(Mid(lsSQL, 3)) + 1
140         End If
150      End If
160      Set lors = Nothing
170      GetNextTrans = Format(Date, "yy") & Format(lnCode, String(lnLen - 2, "0"))
End Function

Sub clearMaster()
10       If p_oMaster.State = adStateOpen Then
20          If p_oMaster.EditMode <> adEditNone Then p_oMaster.CancelUpdate
30          p_oMaster.Close
40       End If
   
50       p_oMaster.Open p_oMod.AddCondition(p_sMasSQL, "0 = 1"), p_oAppDrivr.Connection, , _
                       adLockOptimistic, adCmdText
60       p_oMaster.AddNew
End Sub

Sub clearDetail()
10       If p_oDetail.State = adStateOpen Then
20          If p_oDetail.EOF Then p_oDetail.Move 0, adBookmarkFirst
30          If p_oDetail.EditMode <> adEditNone Then p_oDetail.CancelUpdate
40          p_oDetail.Close
50       End If
   
60       p_oDetail.Open p_oMod.AddCondition(p_sLgrSQL, "0 = 1"), p_oAppDrivr.Connection, , _
                      adLockOptimistic, adCmdText
70       Set p_oDetail.ActiveConnection = Nothing
80       p_oDetail.AddNew
End Sub

Function isDetailModify() As Boolean
10       Dim lors As Recordset
20       Dim lsOldProc As String
30       Dim lsSQL As String
40       Dim lsField As String
50       Dim lsColumn As String
60       Dim lnCol As Integer
70       Dim lnRow As Long
   
80       lsOldProc = p_oAppDrivr.ProcName("isDetailModify")
90       On Error GoTo errProc
100      isDetailModify = True
   
110      If pnEditMode = xeModeAddNew Then GoTo endProc
   
   ' set the retrieval arguments of the recordset
120      lsSQL = p_sLgrSQL
130      For pnCtr = 0 To UBound(p_asLgrRefFld)
140         lsField = p_asLgrRefFld(pnCtr)
150         lsColumn = lsField
160         lnCol = InStr(lsColumn, ".")
170         If lnCol > 0 Then lsColumn = Mid(lsColumn, lnCol + 1)
180         lsSQL = p_oMod.AddCondition(lsSQL, lsField & " = " & _
                  p_oMod.strParm(p_oMaster(lsColumn).Value))
190      Next
200      Set lors = New Recordset
210      lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
   
220      lnRow = 0
230      If p_oDetail.RecordCount > 0 Then p_oDetail.MoveFirst
240      lsField = p_asLgrRefFld(0)
250      lnCol = InStr(lsField, ".")
260      If lnCol > 0 Then lsField = Mid(lsField, lnCol + 1)
270      For pnCtr = 0 To p_oDetail.RecordCount - 1
280         If IsNull(p_oDetail(lsField)) Then Exit For
290         If p_oDetail(lsField) = Empty Then Exit For
300         lnRow = pnCtr + 1
310         p_oDetail.MoveNext
320      Next
   
330      If lors.RecordCount <> lnRow Then GoTo endProc
   
340      p_oDetail.MoveFirst
350      For pnCtr = 0 To lnRow - 1
360         For lnCol = 0 To lors.Fields.Count - 1
370            If p_oDetail(lnCol) <> lors(lnCol) Then GoTo endProc
380         Next
390         lors.MoveNext
400         p_oDetail.MoveNext
410      Next
   
420      isDetailModify = False
   
endProc:
430      Set lors = Nothing
440      p_oAppDrivr.ProcName lsOldProc
450      Exit Function
errProc:
460      ShowError lsOldProc
End Function

Function isMasterModify() As Boolean
10       Dim lors As Recordset
20       Dim lsOldProc As String
30       Dim lsSQL As String
40       Dim lsField As String
50       Dim lnCol As Integer
60       Dim lnRow As Long
   
70       lsOldProc = p_oAppDrivr.ProcName("isMasterModify")
80       On Error GoTo errProc
90       isMasterModify = True
   
100      If pnEditMode = xeModeAddNew Then GoTo endProc
   
   ' set the retrieval arguments of the recordset
110      lsSQL = p_sMasSQL
120      For pnCtr = 0 To UBound(p_asMasRefFld)
130         lsField = p_asMasRefFld(pnCtr)
140         lnCol = InStr(lsField, ".")
150         If lnCol > 0 Then lsField = Mid(lsField, lnCol + 1)
160         lsSQL = p_oMod.AddCondition(lsSQL, p_asMasRefFld(pnCtr) & " = " & _
                  p_oMod.strParm(p_oMaster(lsField).Value))
170      Next
180      Set lors = New Recordset
190      lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
   
200      For lnCol = 0 To lors.Fields.Count - 1
210         If p_oMaster(lnCol) <> lors(lnCol) Then GoTo endProc
220      Next
   
230      isMasterModify = False
   
endProc:
240      Set lors = Nothing
250      p_oAppDrivr.ProcName lsOldProc
260      Exit Function
errProc:
270      ShowError lsOldProc
End Function

Function isPosted() As Boolean
10       isPosted = True
20       If Not (p_oMaster("cTranStat") = xeStatePosted Or _
         p_oMaster("cTranStat") = xeStateCancelled) Then isPosted = False
End Function

Sub ViewUserModify()
10       Dim lors As Recordset
20       Dim lsOldProc As String
30       Dim lsSQL As String
   
40       lsOldProc = p_oAppDrivr.ProcName("isMasterModify")
50       On Error GoTo errProc
   
60       lsSQL = "SELECT sUserName" & _
            " FROM xxxSysUser" & _
            " WHERE sUserIDxx = " & p_oMod.strParm(p_oMod.Decrypt(Master("sModified")))

70       Set lors = New Recordset
80       lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
   
90       If lors.EOF Then
100         MsgBox "User does not Exist in the User Table!!!", vbCritical, "Warning"
110         GoTo endProc
120      End If
   
130      MsgBox "    ««« LAST MODIFICATION »»»" & vbCrLf & vbCrLf & _
            "User Name      : " & p_oMod.Decrypt(lors("sUserName"), p_oAppDrivr.Machinex) & vbCrLf & _
            "Date Modified : " & Format(Master("dModified"), "MMMM DD, YYYY HH:NN:SS") _
            , vbInformation, "Notice"
         
endProc:
140      Set lors = Nothing
150      p_oAppDrivr.ProcName lsOldProc
160      Exit Sub
errProc:
170      ShowError lsOldProc
End Sub

Private Function isRefEmpty() As Boolean
10       Dim lsField As String
20       Dim lnCol As String
   
30       isRefEmpty = True
40       For pnCtr = 0 To UBound(p_asMasRefFld)
50          lsField = p_asLgrRefFld(pnCtr)
60          lnCol = InStr(lsField, ".")
70          If lnCol > 0 Then lsField = Mid(lsField, lnCol + 1)
80          If p_oMaster(lsField) = Empty Then Exit Function
90       Next
100      isRefEmpty = False
End Function

Private Sub ShowError(ByVal lsProcName As String)
10       With p_oAppDrivr
20          .ShowError "MasterDetailBase", .ProcName(lsProcName), Err.Number, Err.Description, Erl
30       End With
40       With Err
50          .Raise .Number, .Source, .Description
60       End With
End Sub

Private Sub Class_Initialize()
10       Set p_oMaster = New Recordset
20       p_oMaster.CursorLocation = adUseClient
30       Set p_oDetail = New Recordset
40       p_oMaster.CursorLocation = adUseClient
End Sub

Private Sub Class_Terminate()
10       Set p_oMaster = Nothing
20       Set p_oDetail = Nothing
30       Set p_oMod = Nothing
End Sub

Private Sub getMasCode(ByVal lnIndex As Integer, ByVal lsValue As String, ByVal lbSearch As Boolean)
10       Dim lors As Recordset
20       Dim lsSQL As String
30       Dim lsOldProc As String
40       Dim lasRef() As String

50       lsOldProc = p_oAppDrivr.ProcName("getMasCode")
60       On Error GoTo errProc

70       If lsValue = p_asLstVal(lnIndex) And lbSearch = False Then GoTo endProc
   
80       If p_asLstCri(lnIndex) = Empty Then
90          If p_asLstCol(lnIndex) = Empty Then
100            GoTo endProc
110         Else
120            p_asLstCri(lnIndex) = p_asLstCol(lnIndex)
130         End If
140      End If
      
150      lasRef = Split(p_asLstCri(lnIndex), "»")
                  
160      lsSQL = p_asLstSQL(lnIndex)
   
170      If lsValue <> "" Then
180         lsSQL = p_oMod.AddCondition(lsSQL, lasRef(1) & _
            IIf(lbSearch, " Like " & p_oMod.strParm(lsValue & "%"), _
               " = " & p_oMod.strParm(lsValue)))
190      ElseIf lbSearch = False Then
200         GoTo endWithClear
210      End If
   
220      Set lors = New Recordset
230      lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
                     
240      If lors.EOF Then GoTo endWithClear
250      If lors.RecordCount = 1 Then
260         p_oMaster(lnIndex) = lors(0)
270         p_asLstVal(lnIndex) = lors(1)
280         RaiseEvent MasterRetrieved(lnIndex)
290         GoTo endProc
300      End If
   
310      lsSQL = p_oMod.KwikBrowse(p_oAppDrivr, lors, p_asLstCol(lnIndex), p_asLstTit(lnIndex), p_asLstPic(lnIndex), , p_asLstCri(lnIndex))
   
320      If lsSQL = "" Then GoTo endProc
330      lasRef = Split(lsSQL, "»")
340      p_oMaster(lnIndex) = lasRef(0)
350      p_asLstVal(lnIndex) = lasRef(1)
360      RaiseEvent MasterRetrieved(lnIndex)
   
endProc:
370      Set lors = Nothing
380      p_oAppDrivr.ProcName lsOldProc
390      Exit Sub
endWithClear:
400      p_oMaster(lnIndex) = ""
410      p_asLstVal(lnIndex) = ""
420      RaiseEvent MasterRetrieved(lnIndex)
430      GoTo endProc
errProc:
440      ShowError lsOldProc
End Sub

Private Sub getMasDesc(ByVal lnIndex As Integer, ByVal lsValue As String)
10       Dim lors As Recordset
20       Dim lsSQL As String
30       Dim lsOldProc As String
40       Dim lasRef() As String
50       Dim lnCtr As Integer

60       lsOldProc = p_oAppDrivr.ProcName("getMasDesc")
70       On Error GoTo errProc
      
80       If lsValue = "" Then GoTo endProc
   
90       If p_asLstCri(lnIndex) = Empty Then
100         If p_asLstCol(lnIndex) = Empty Then
110            GoTo endProc
120         Else
130            p_asLstCri(lnIndex) = p_asLstCol(lnIndex)
140         End If
150      End If
      
160      p_asLstVal(lnIndex) = ""
      
170      lasRef = Split(p_asLstCri(lnIndex), "»")
                  
180      lsSQL = p_asLstSQL(lnIndex)
190      lsSQL = p_oMod.AddCondition(lsSQL, lasRef(0) & " = " & p_oMod.strParm(lsValue))
   
200      Set lors = New Recordset
210      lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
                     
220      If lors.EOF Then GoTo endProc
230      p_asLstVal(lnIndex) = lors(1)
   
endProc:
240      Set lors = Nothing
250      p_oAppDrivr.ProcName lsOldProc
260      Exit Sub
errProc:
270      ShowError lsOldProc
End Sub

Private Function getDetDesc(ByVal lnRow As Integer, ByVal lnIndex As Variant, ByVal lsValue As String) As String
10       Dim lors As Recordset
20       Dim lsSQL As String
30       Dim lsOldProc As String
40       Dim lasRef() As String
50       Dim lnCtr As Integer

60       lsOldProc = p_oAppDrivr.ProcName("getDetDesc")
70       On Error GoTo errProc
80       getDetDesc = ""
   
90       If lsValue = "" Then GoTo endProc
   
100      If p_asLgrLstCri(lnIndex) = Empty Then
110         If p_asLgrLstCol(lnIndex) = Empty Then
120            GoTo endProc
130         Else
140            p_asLgrLstCri(lnIndex) = p_asLgrLstCol(lnIndex)
150         End If
160      End If
      
170      lasRef = Split(p_asLgrLstCri(lnIndex), "»")
                  
180      lsSQL = p_asLgrLstSQL(lnIndex)
190      lsSQL = p_oMod.AddCondition(lsSQL, lasRef(0) & " = " & p_oMod.strParm(lsValue))
   
200      Set lors = New Recordset
210      lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText

220      If lors.EOF Then GoTo endProc
230      getDetDesc = lors(1)
   
endProc:
240      Set lors = Nothing
250      p_oAppDrivr.ProcName lsOldProc
260      Exit Function
errProc:
270      ShowError lsOldProc
End Function

Private Sub getDetCode(ByVal lnRow As Integer, _
                              ByVal lnIndex As Integer, _
                              ByVal lsValue As String, _
                              ByVal lbSearch As Boolean)
10       Dim lors As Recordset
20       Dim lsSQL As String
30       Dim lsOldProc As String
40       Dim lasRef() As String

50       lsOldProc = p_oAppDrivr.ProcName("getDetCode")
60       On Error GoTo errProc
   
70       If p_asLgrLstCri(lnIndex) = Empty Then
80          If p_asLgrLstCol(lnIndex) = Empty Then
90             GoTo endProc
100         Else
110            p_asLgrLstCri(lnIndex) = p_asLgrLstCol(lnIndex)
120         End If
130      End If
      
140      p_oDetail.Move lnRow, adBookmarkFirst
150      lasRef = Split(p_asLgrLstCri(lnIndex), "»")
                  
160      lsSQL = p_asLgrLstSQL(lnIndex)
   
170      If lsValue <> "" Then
180         lsSQL = p_oMod.AddCondition(lsSQL, lasRef(1) & _
            IIf(lbSearch, " Like " & p_oMod.strParm(lsValue & "%"), _
               " = " & p_oMod.strParm(lsValue)))
190      ElseIf lbSearch = False Then
200         GoTo endWithClear
210      End If
   
220      Set lors = New Recordset
230      lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
                     
240      If lors.EOF Then GoTo endWithClear
250      If lors.RecordCount = 1 Then
260         p_oDetail(lnIndex) = lors(0)
270         RaiseEvent DetailRetrieved(lnIndex)
280         GoTo endProc
290      End If
   
300      lsSQL = p_oMod.KwikBrowse(p_oAppDrivr, lors, p_asLgrLstCol(lnIndex), p_asLgrLstTit(lnIndex), p_asLgrLstPic(lnIndex))
   
310      If lsSQL = "" Then GoTo endProc
320      lasRef = Split(lsSQL, "»")
330      p_oDetail(lnIndex) = lasRef(0)
340      RaiseEvent DetailRetrieved(lnIndex)
   
endProc:
350      Set lors = Nothing
360      p_oAppDrivr.ProcName lsOldProc
370      Exit Sub
endWithClear:
380      p_oDetail(lnIndex) = ""
390      RaiseEvent DetailRetrieved(lnIndex)
400      GoTo endProc
errProc:
410      ShowError lsOldProc
End Sub

Private Sub LoadDescription()
10       For pnCtr = 1 To p_oMaster.Fields.Count - 1
20          If p_asLstSQL(pnCtr) <> Empty Then
30             If IsNull(p_oMaster(pnCtr)) Then Exit Sub
40             getMasDesc pnCtr, p_oMaster(pnCtr)
50          End If
60       Next
End Sub

Private Function HasRights2Save() As Boolean
10       Dim lsOldProc As String
20       Dim lnUserRght As Integer, lsUserName As String, lsUserIDxx As String
30       Dim lasRights() As String, lnRights As Integer
40       Dim lnSeek As Integer, lbSeek As Boolean
   
50       lsOldProc = p_oAppDrivr.ProcName("HasRights2Save")
60       On Error GoTo errProc
70       HasRights2Save = False
   
80       If isPosted Then
90          If p_oAppDrivr.UserLevel = xeSysMaster Then
100            HasRights2Save = True
110         Else
120            MsgBox "Modification of posted or cancelled transaction is not allowed!!!" & vbCrLf & vbCrLf & _
               "Request not granted!!!", vbCritical, "Warning"
130         End If
140         GoTo endProc
150      End If

160      With p_oAppDrivr
170         If .MDIMain.Controls(.MenuName).Tag = "" Then
180            ReDim lasRights(4) As String
   
190            lasRights(1) = 192
200            lasRights(2) = 192
210            lasRights(3) = 192
220         Else
230            lasRights = Split(.MDIMain.Controls(.MenuName).Tag, "»")
240         End If
      
250         lbSeek = False
260         If pnEditMode = xeModeAddNew Then
         ' Add Rights
270            lnRights = CInt(lasRights(1))
280            If lnRights And .UserLevel = 0 Then
290               lnSeek = MsgBox("User Doesn't Have the Right to Create New Transaction!!!" & _
                              vbCrLf & "Seek Authorized User's Approval?", _
                              vbQuestion + vbYesNo + vbDefaultButton2, "Confirm")
300               If lnSeek = vbNo Then GoTo endProc
310               lbSeek = True
320            End If
330         ElseIf pnEditMode = xeModeUpdate Then
         ' update Rights
340            lnRights = CInt(lasRights(2))
         
350            If (lnRights And .UserLevel) = 0 Then
360               Select Case .UserLevel
            Case xeEncoder
370                  lbSeek = True
380               Case xeSupervisor
390                  If DateDiff("d", p_oMaster("dModified"), p_oAppDrivr.ServerDate) > 2 Then lbSeek = True
400               Case xeManager
410                  If DateDiff("d", p_oMaster("dModified"), p_oAppDrivr.ServerDate) > 7 Then lbSeek = True
420               Case xeAudit
430                  lbSeek = True
440               Case xeSysAdmin
450                  lbSeek = True
460               End Select
470            End If
480         End If
      
490         If lbSeek = False Then
500            HasRights2Save = True
510            GoTo endProc
520         End If
      
530         lnSeek = MsgBox("User Doesn't Have the Right to Update Transaction!!!" & _
                        vbCrLf & "Seek Authorized User's Approval?", _
                        vbQuestion + vbYesNo + vbDefaultButton2, "Confirm")
540         If lnSeek = vbNo Then GoTo endProc
      
550         If p_oMod.GetApproval(p_oAppDrivr, lnUserRght, lsUserIDxx, lsUserName) = False Then GoTo endProc
      
560         If pnEditMode = xeModeAddNew Then
570            If lnUserRght And .UserLevel = 0 Then
580               MsgBox "Approving User Doesn't Have the Right to Create New Transaction!!!" & _
                  "Seek Other User's Assistance then Try Again", vbCritical, "Warning"
590               GoTo endProc
600            End If
610         ElseIf pnEditMode = xeModeUpdate Then
620            If lnUserRght And .UserLevel = 0 Then
630               Select Case lnUserRght
            Case xeSupervisor
640                  If DateDiff("d", p_oMaster("dModified"), p_oAppDrivr.ServerDate) <= 2 Then lbSeek = False
650               Case xeManager
660                  If DateDiff("d", p_oMaster("dModified"), p_oAppDrivr.ServerDate) <= 31 Then lbSeek = False
670               Case xeSysAdmin
680                  If .UserLevel <> xeSysAdmin Then lbSeek = False
690               End Select
            
700               If lbSeek Then
710                  MsgBox "Approving User Doesn't Have the Right to Update Transaction!!!" & _
                     "Seek Other User's Assistance then Try Again", vbCritical, "Warning"
720                  GoTo endProc
730               End If
740            End If
750         End If
760      End With
770      HasRights2Save = True

endProc:
780      p_oAppDrivr.ProcName lsOldProc
790      Exit Function
errProc:
800      ShowError lsOldProc
End Function

Private Function HasRights2Cancel() As Boolean
10       Dim lsOldProc As String
20       Dim lnUserRght As Integer, lsUserName As String, lsUserIDxx As String
30       Dim lasRights() As String, lnRights As Integer
40       Dim lnSeek As Integer, lbSeek As Boolean
   
50       lsOldProc = p_oAppDrivr.ProcName("HasRights2Cancel")
60       On Error GoTo errProc
70       HasRights2Cancel = False
   
80       If isPosted Then
90          If p_oAppDrivr.UserLevel = xeSysMaster Then
100            HasRights2Cancel = True
110         Else
120            MsgBox "Cancellation of posted or cancelled transaction is not allowed!!!" & vbCrLf & vbCrLf & _
               "Request not granted!!!", vbCritical, "Warning"
130         End If
140         GoTo endProc
150      End If
   
160      With p_oAppDrivr
170         If .MDIMain.Controls(.MenuName).Tag = "" Then
180            ReDim lasRights(4) As String
   
190            lasRights(1) = 192
200            lasRights(2) = 192
210            lasRights(3) = 192
220         Else
230            lasRights = Split(.MDIMain.Controls(.MenuName).Tag, "»")
240         End If
      
      ' Cancel Rights
250         lnRights = CInt(lasRights(3))
260         If (lnRights And .UserLevel) > 0 Then
270            HasRights2Cancel = True
280            GoTo endProc
290         End If
      
300         lbSeek = False
310         Select Case .UserLevel
      Case xeEncoder
320            lbSeek = True
330         Case xeSupervisor
340            If DateDiff("d", p_oMaster("dModified"), .ServerDate) > 2 Then lbSeek = True
350         Case xeManager
360            If DateDiff("d", p_oMaster("dModified"), .ServerDate) > 7 Then lbSeek = True
370         Case xeSysAdmin
380            lbSeek = True
390         Case xeAudit
400            lbSeek = True
410         End Select

420         If lbSeek = False Then
430            HasRights2Cancel = True
440            GoTo endProc
450         End If
   
460         lnSeek = MsgBox("User Doesn't Have the Right to Cancel Transaction!!!" & _
                        vbCrLf & "Seek Authorized User's Approval?", _
                        vbQuestion + vbYesNo + vbDefaultButton2, "Confirm")
470         If lnSeek = vbNo Then GoTo endProc
      
480         If p_oMod.GetApproval(p_oAppDrivr, lnUserRght, lsUserIDxx, lsUserName) = False Then GoTo endProc
   
490         If lnUserRght And .UserLevel = 0 Then
500            Select Case lnUserRght
         Case xeSupervisor
510               If DateDiff("d", p_oMaster("dModified"), .ServerDate) <= 2 Then lbSeek = False
520            Case xeManager
530               If DateDiff("d", p_oMaster("dModified"), .ServerDate) <= 30 Then lbSeek = False
540            Case xeSysAdmin
550               If .UserLevel <> xeSysAdmin Then lbSeek = False
560            End Select
         
570            If lbSeek Then
580               MsgBox "Approving User Doesn't Have the Right to Cancel Transaction!!!" & _
                  "Seek Other User's Assistance then Try Again", vbCritical, "Warning"
590               GoTo endProc
600            End If
610         End If
620      End With
630      HasRights2Cancel = True

endProc:
640      p_oAppDrivr.ProcName lsOldProc
650      Exit Function
errProc:
660      ShowError lsOldProc
End Function

Private Function HasRights2Delete() As Boolean
10       Dim lsOldProc As String
20       Dim lnUserRght As Integer, lsUserName As String, lsUserIDxx As String
30       Dim lasRights() As String, lnRights As Integer
40       Dim lnSeek As Integer
   
50       lsOldProc = p_oAppDrivr.ProcName("HasRights2Delete")
60       On Error GoTo errProc
70       HasRights2Delete = False
   
80       If isPosted Then
90          If p_oAppDrivr.UserLevel = xeSysMaster Then
100            HasRights2Delete = True
110         Else
120            MsgBox "Deletion of posted or cancelled transaction is not allowed!!!" & vbCrLf & vbCrLf & _
               "Request not granted!!!", vbCritical, "Warning"
130         End If
140         GoTo endProc
150      End If

160      With p_oAppDrivr
170         If .MDIMain.Controls(.MenuName).Tag = "" Then
180            ReDim lasRights(4) As String
   
190            lasRights(1) = 192
200            lasRights(2) = 192
210            lasRights(3) = 192
220         Else
230            lasRights = Split(.MDIMain.Controls(.MenuName).Tag, "»")
240         End If
      
250         lnRights = CInt(lasRights(4))
260         If (lnRights And .UserLevel) > 0 Then
270            HasRights2Delete = True
280            GoTo endProc
290         End If
300         lnSeek = MsgBox("User Doesn't Have the Right to Delete Transaction!!!" & _
                        vbCrLf & "Seek Authorized User's Approval?", _
                        vbQuestion + vbYesNo + vbDefaultButton2, "Confirm")
310         If lnSeek = vbNo Then GoTo endProc
      
320         If p_oMod.GetApproval(p_oAppDrivr, lnUserRght, lsUserIDxx, lsUserName) = False Then GoTo endProc
   
330         If lnRights And .UserLevel > 0 Then
340            HasRights2Delete = True
350            GoTo endProc
360         End If
      
370         MsgBox "Approving User Doesn't Have the Right to Delete Transaction!!!" & _
            "Seek Other User's Assistance then Try Again", vbCritical, "Warning"
380      End With
   
endProc:
390      p_oAppDrivr.ProcName lsOldProc
400      Exit Function
errProc:
410      ShowError lsOldProc
End Function
