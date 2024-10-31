VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SystemUser"
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
' SystemUser v1.0
'     An object for system user.
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
'  XerSys [ 05/20/2004 03:10 pm ]
'     Start creating this object.
'  XerSys [ 05/22/2004 09:20 am ]
'     Continue coding this object.
'  XerSys [ 06/02/2004 11:04 am ]
'     Continue coding this object.
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€

Private p_oAppDrivr As AppDriver
Private p_oRecordst As Recordset
Private p_oMod As New MainModules

Private p_sUserIDxx As String
Private p_sLogNamex As String
Private p_sPassword As String
Private p_sUserName As String
Private p_sEmployNo As String
Private p_nUserLevl As Integer
Private p_cUserType As Byte
Private p_sProdctID As String
Private p_cUserStat As Byte
Private p_nSysError As Integer
Private p_cLogStatx As Byte
Private p_cLockStat As Byte
Private p_cAllwLock As Byte
Private p_cAllwView As Byte
Private p_sSkinCode As String
Private p_sCompName As String

Private psModlName As String
Private pnEditMode As xeEditMode
Private psSQL As String
Private pnCtr As Long

Property Set AppDriver(oAppDriver As AppDriver)
10       Set p_oAppDrivr = oAppDriver
End Property

Property Get UserID() As String
10       UserID = p_sUserIDxx
End Property

'Property Let UserID(ByVal Value As String)
'   If Ambient.UserMode Then Err.Raise 382
'   p_sUserIDxx = Value
'End Property

Property Get LogName() As String
10       LogName = p_sLogNamex
End Property

Property Let LogName(ByVal Value As String)
10       p_sLogNamex = Value
End Property

Property Get Password() As String
10       Password = p_sPassword
End Property

Property Let Password(ByVal Value As String)
10       p_sPassword = Value
End Property

Property Get UserName() As String
10       UserName = p_sUserName
End Property

Property Let UserName(ByVal Value As String)
10       p_sUserName = Value
End Property

Property Get EmployeeNo() As String
10       EmployNo = p_sEmployNo
End Property

Property Let EmployeeNo(ByVal Value As String)
10       p_sEmployNo = Value
End Property

Property Get UserLevel() As Integer
10       UserLevel = p_nUserLevl
End Property

Property Let UserLevel(ByVal Value As Integer)
10       p_nUserLevl = Value
End Property

Property Get UserType() As String
10       UserType = p_cUserType
End Property

Property Let UserType(ByVal Value As String)
10       p_cUserType = Value
End Property

Property Get ProductID() As String
10       ProductID = p_sProdctID
End Property

Property Let ProductID(ByVal Value As String)
10       p_sProdctID = Value
End Property

Property Get UserStatus() As xeUserStatus
10       UserStatus = p_cUserStat
End Property

Property Let UserStatus(ByVal Value As xeUserStatus)
10       p_cUserStat = Value
End Property

Property Get SysError() As Integer
10       SysError = p_nSysError
End Property

Property Let SysError(ByVal Value As Integer)
10       p_nSysError = Value
End Property

Property Get LogStatus() As xeLogStatus
10       LogStatus = p_cLogStatx
End Property

Property Let LogStatus(ByVal Value As xeLogStatus)
10       p_cLogStatx = Value
End Property

Property Get LockStatus() As xeLockStatus
10       LockStatus = p_cLockStat
End Property

Property Let LockStatus(ByVal Value As xeLockStatus)
10       p_cLockStat = Value
End Property

Property Let AllowLock(ByVal Value As Boolean)
10       p_cAllwLock = IIf(Value, 1, 0)
End Property

Property Get AllowLock() As Boolean
10       AllowLock = p_cAllwLock = 1
End Property

Property Let AllowView(ByVal Value As Boolean)
10       p_cAllwView = IIf(Value, 1, 0)
End Property

Property Get AllowView() As Boolean
10       AllowView = p_cAllwView = 1
End Property

Property Get Skin() As String
10       Skin = p_sSkinCode
End Property

Property Let Skin(ByVal Value As String)
10       p_sSkinCode = Value
End Property

Sub AddNew()
10       pnEditMode = xeModeAddNew
20       initRecord
End Sub

Function Update() As Boolean
10       Dim lsOldProc As String
   
20       If isAppDriverOK = False Then GoTo endProc
   
30       lsOldProc = p_oAppDrivr.ProcName("Update")
40       Update = False
50       ''On Error GoTo errProc
   
   ' always check the connection
60       If isAppDriverOK Then GoTo endProc
   
   ' first verify the entries then notify user for confirmation
70       If Not isRecordOK Then GoTo endProc
   
80       With p_oMod
90          If pnEditMode = xeModeAddNew Then
100            psSQL = "INSERT INTO xxxSysUser" & _
                     "  sUserIDxx" & _
                     ", sLogNamex" & _
                     ", sPassword" & _
                     ", sUserName" & _
                     ", sEmployNo" & _
                     ", nUserLevl" & _
                     ", cUserType" & _
                     ", sProdctID" & _
                     ", cUserStat" & _
                     ", nSysError" & _
                     ", cLogStatx" & _
                     ", sSkinCode" & _
                     ", cLockStat" & _
                     ", cAllwLock" & _
                     ", cAllwView" & ")" & _
                  " VALUES ("
110            psSQL = psSQL & _
                     .strParm(p_sUserIDxx) & _
                     ", " & .strParm(p_sLogNamex) & _
                     ", " & .strParm(p_sPassword) & _
                     ", " & .strParm(p_sUserName) & _
                     ", " & .strParm(p_sEmployNo) & _
                     ", " & p_nUserLevl & _
                     ", " & .strParm(p_cUserType) & _
                     ", " & .strParm(p_sProdctID) & _
                     ", " & .strParm(p_cUserStat) & _
                     ", " & p_nSysError & _
                     ", " & .strParm(p_cLogStatx) & _
                     ", " & .strParm(p_sSkinCode) & _
                     ", " & .strParm(p_cLockStat) & _
                     ", " & .strParm(p_cAllwView) & ")"
120         Else
130            psSQL = "UPDATE xxxSysUser SET" & _
                     ", sLogNamex = " & .strParm(p_sLogNamex) & _
                     ", sPassword = " & .strParm(p_sPassword) & _
                     ", sUserName = " & .strParm(p_sUserName) & _
                     ", sEmployNo = " & .strParm(p_sEmployNo) & _
                     ", nUserLevl = " & p_nUserLevl & _
                     ", cUserType = " & .strParm(p_cUserType) & _
                     ", sProdctID = " & .strParm(p_sProdctID) & _
                     ", cUserStat = " & .strParm(p_cUserStat) & _
                     ", nSysError = " & p_nSysError & _
                     ", cLogStatx = " & .strParm(p_cLogStatx) & _
                     ", sSkinCode = " & .strParm(p_sSkinCode) & _
                     ", cLockStat = " & .strParm(p_cLockStat) & _
                     ", cAllwLock = " & .strParm(p_cAllwLock) & _
                     ", cAllwView = " & .strParm(p_cAllwView) & _
                  " WHERE sUserIDxx = " & .strParm(p_sUserIDxx)
140         End If
150      End With
   
160      ExecuteSQL
170      Update = True
   
endProc:
180      p_oAppDrivr.ProcName lsOldProc
190      Exit Function
errProc:
200      ShowError lsOldProc
End Function
   
Function Suspend() As Boolean
10       Dim lsOldProc As String
20       Dim lnSuspend As VbMsgBoxResult
   
30       If isAppDriverOK = False Then GoTo endProc
   
40       lsOldProc = p_oAppDrivr.ProcName("Suspend")
50       Suspend = False
60       ''On Error GoTo errProc
   
70       If pnEditMode <> xeModeReady Then GoTo endProc
   
80       lnSuspend = MsgBox("Suspend User?", vbQuestion + vbYesNo, "Confirm")
   
90       If lnSuspend <> vbYes Then GoTo endProc
   
100      With p_oMod
110         psSQL = "UPDATE xxxSysUser" & _
                  " SET cUserStat = " & .strParm(xeUserStateSuspended) & _
               " WHERE sUserIDxx = " & .strParm(UserID)
120      End With
      
130      ExecuteSQL
140      Suspend = True
   
endProc:
150      p_oAppDrivr.ProcName lsOldProc
160      Exit Function
errProc:
170      ShowError lsOldProc
End Function

Sub LogOutUser()
10       Dim lsOldProc As String
   
20       If isAppDriverOK = False Then GoTo endProc
   
30       lsOldProc = p_oAppDrivr.ProcName("LogOutUser")
40       ''On Error GoTo errProc
   
   ' logout user
50       With p_oMod
60          psSQL = "UPDATE xxxSysUser SET" & _
                  "  cLogStatx = " & .strParm(xeLogStateOut) & _
                  ", cLockStat = " & .strParm(p_cLockStat) & _
                  ", nSysError = " & p_nSysError & _
                  ", sCompName = ''" & _
               " WHERE sLogName = " & .strParm(p_sUserIDxx)
70          ExecuteSQL
         
80          psSQL = "UPDATE xxxSystemLog SET" & _
                  " dLogOut = GETDATE()" & _
               " WHERE sLogNo = " & .strParm(sLogNo)
               
90          ExecuteSQL
100      End With
   
endProc:
110      p_oAppDrivr.ProcName lsOldProc
120      Exit Sub
errProc:
130      ShowError lsOldProc
End Sub

Private Sub initRecord()
10       p_sUserIDxx = p_oMod.GetNextCode("xxxSysUser", "sUserIDxx", True, p_oAppDrivr.Connection)
20       p_sLogNamex = ""
30       p_sPassword = ""
40       p_sUserName = ""
50       p_sEmployNo = ""
60       p_nUserLevl = 0
70       p_cUserType = xeUserTypeLocal
80       p_sProdctID = ""
90       p_cUserStat = xeUserStateActive
100      p_nSysError = 0
110      p_cLogStatx = xeLogStateOut
120      p_cLockStat = xeLockStateLock
End Sub

Function LogUser(ByVal sUserID As String) As Boolean
10       Dim lors As Recordset
20       Dim loCrypt As Crypto
30       Dim lsUser As String
40       Dim lsPassword As String
50       Dim lbLogIn As Boolean
60       Dim lsOldProc As String
70       Dim lsLogNo As String
   
80       If isAppDriverOK = False Then GoTo endProc
   
90       lsOldProc = p_oAppDrivr.ProcName("LogUser")
100      LogUser = False
110      ''On Error GoTo errProc
   
120      psSQL = "SELECT *" & _
            " FROM xxxSysUser" & _
            " WHERE sUserIDxx = " & p_oMod.strParm(sUserID)

130      Set lors = New Recordset
140      lors.Open psSQL, p_oAppDrivr.Connection, , , adCmdText
   
150      If lors.EOF Then GoTo endProc
   
160      p_sUserIDxx = lors("sUserIDxx")
170      p_sLogNamex = lors("sLogNamex")
180      p_sPassword = lors("sPassword")
190      p_sUserName = lors("sUserName")
200      p_sEmployNo = lors("sEmployNo")
210      p_nUserLevl = lors("nUserLevl")
220      p_cUserType = lors("cUserType")
230      p_sProdctID = lors("sProdctID")
240      p_cUserStat = lors("cUserStat")
250      p_nSysError = lors("nSysError")
260      p_cLogStatx = lors("cLogStatx")
270      p_cLockStat = lors("cLockStat")
280      p_cAllwLock = lors("cAllwLock")
290      p_cAllwView = lors("cAllwView")
300      p_sSkinCode = lors("sSkinCode")
310      p_sCompName = p_oAppDrivr.ComputerName

320      Set loCrypt = New Crypto
330      loCrypt.Signature = p_oAppDrivr.Machinex

340      loCrypt.InBuffer = p_sUserName
350      loCrypt.Decrypt
360      p_sUserName = loCrypt.OutBuffer

370      loCrypt.InBuffer = p_sLogNamex
380      loCrypt.Decrypt
390      p_sLogNamex = loCrypt.OutBuffer

400      lors.Close
   
   ' user with administrator level up don't follow standard login procedure
410      If p_nUserLevl < xeSysAdmin Then
      ' check first if user is currently login to another computer
420         If cLogStat = "1" Then
430            MsgBox "You are currently logged to another workstation!!!" & vbCrLf & _
                  "Loging to two station simultaneously is not permitted!!!" & vbCrLf & vbCrLf & _
                  App.EXEName & " " & App.Major & "." & App.Minor & " will now quit!!!", _
                  vbInformation, "Warning"
440            GoTo endProc
450         End If

      ' check if user right is lock or not
460         If p_cLockStat = xeLockStateLock Then
470            If Not UnlockUser Then
480               MsgBox "Unlocking of user rights failed!!!" & vbCrLf & _
                  App.EXEName & " " & App.Major & "." & App.Minor & " will now quit!!!", _
                  vbInformation, "Warning"
490               GoTo endProc
500            End If
510         End If
520      End If

530      p_oAppDrivr.Connection.BeginTrans
540      With p_oMod
550         lsLogNo = .GetNextCode("xxxSysUserLog", "sLogNoxxx", True, p_oAppDrivr.Connection)
   
      ' insert the user log
560         psSQL = "INSERT INTO xxxSysUserLog (" & _
                  "  sLogNoxxx" & _
                  ", sUserIDxx" & _
                  ", dLogInxxx" & _
                  ", sProdctID" & _
                  ", sCompName" & _
               ") VALUES (" & _
                  .strParm(lsLogNo) & _
                  ", " & .strParm(p_sUserIDxx) & _
                  ", GETDATE()" & _
                  ", " & .strParm(p_oAppDrivr.ProductID) & _
                  ", " & .strParm(p_oAppDrivr.ComputerName) & ")"
570         ExecuteSQL
      
      ' update the user info
580         psSQL = "UPDATE xxxSysUser SET" & _
                  "  cLogStatx = " & .strParm(xeLogStateIn) & _
                  ", sCompName = " & .strParm(p_oAppDrivr.ComputerName) & _
               " WHERE sUserIdxx = " & .strParm(p_sUserIDxx)
590         ExecuteSQL
600      End With
610      p_oAppDrivr.Connection.CommitTrans
   
620      LogUser = True

endProc:
630      Set loCrypt = Nothing
640      p_oAppDrivr.ProcName lsOldProc
650      Exit Function
errProc:
660      ShowError lsOldProc
End Function

Function LockUser() As Boolean
10       Dim lsOldProc As String
   
20       If isAppDriverOK = False Then GoTo endProc
   
30       lsOldProc = p_oAppDrivr.ProcName("LockUser")
40       LockUser = False
50       ''On Error GoTo errProc
   
60       p_cLockStat = xeLockStateLock
70       p_nSysError = p_nSysError + 1
80       LockUser = True

endProc:
90       p_oAppDrivr.ProcName lsOldProc
100      Exit Function
errProc:
110      ShowError lsOldProc
End Function

Private Function UnlockUser() As Boolean
10       Dim lors As Recordset
20       Dim loForm As frmUnlock
30       Dim loCrypt As Crypto
40       Dim loMod As MainModules
50       Dim lsUser As String
60       Dim lsPassword As String
70       Dim lsOldProc As String
80       Dim lsSQL As String
90       Dim lnCtr As Integer
100      Dim lbLogIn As Boolean
   
110      If isAppDriverOK = False Then GoTo endProc
   
120      lsOldProc = oAppDriver.ProcName("UnlockUser")
130      ''On Error GoTo errProc
140      UnlockUser = False
      
150      Set lors = New Recordset
160      Set loCrypt = New Crypto
170      loCrypt.Signature = p_oAppDrivr.Machinex
180      Set loForm = New frmUnlock
   
190      lnCtr = 0
200      lbLogIn = False
210      Do
220         loForm.AppPath = p_oAppDrivr.AppPath
230         loForm.Show 1
240         If loForm.Cancel = True Then
         ' user either clicked the cancel button or the close key
250            lnCtr = 2
260         Else
         ' Encrypt the username and password
270            loCrypt.InBuffer = loForm.txtPassword
280            loCrypt.Encrypt
290            lsPassword = loCrypt.OutBuffer
        
300            loCrypt.InBuffer = LCase(loForm.txtUserName)
310            loCrypt.Encrypt
320            lsUser = loCrypt.OutBuffer
         
330            With p_oMod
            ' search for the user
340               lsSQL = "SELECT *" & _
                     " FROM xxxSysUser" & _
                     " WHERE sLogNamex = " & .strParm(lsUser) & _
                        " AND sPassword = " & .strParm(lsPassword)
350            End With
         
360            If lors.State = adStateOpen Then lors.Close
         
370            lors.Open lsSQL, oAppDriver.Connection, , , adCmdText
380            If lors.EOF Then
390               MsgBox "Verify your log name and/or password", vbCritical, "Login Error"
400            Else
410               lbLogIn = True
420            End If
430         End If
440         lnCtr = lnCtr + 1
450      Loop Until (lbLogIn = True) Or (lnCtr = 3)
   
460      If lbLogIn = False Then GoTo endProc
      
470      UnlockUser = True
   
endProc:
480      Unload loForm
490      Set lors = Nothing
500      Set loCrypt = Nothing
510      Set loForm = Nothing
   
520      oAppDriver.ProcName lsOldProc
530      Exit Function
errProc:
540      With Err
550         oAppDriver.ShowError "MainModules", _
                           oAppDriver.ProcName(lsOldProc), .Number, .Description, Erl
560         .Raise .Number, .Source, .Description, .HelpFile, .HelpContext
570         .Clear
580      End With
End Function

Private Sub ShowError(ByVal lsProcName As String)
10       With p_oAppDrivr
20          .ShowError "SystemUser", .ProcName(lsProcName), Err.Number, Err.Description, Erl
30       End With
40       With Err
50          .Raise .Number, .Source, .Description, .HelpFile, .HelpContext
60          .Clear
70       End With
End Sub

Private Sub ExecuteSQL()
10       p_oAppDrivr.Connection.Execute psSQL, pnCtr, adCmdText
   
20       If pnCtr = 0 Then
30          Err.Raise 1001, "ADODB Recordset", "Unable to update changes"
40          Err.Clear
50       End If
End Sub

Private Function isAppDriverOK() As Boolean
10       isAppDriverOK = True
20       If p_oAppDrivr Is Nothing Then
30          MsgBox "Application is not yet set!!!" & vbCrLf & _
               "Can not further proceed!!!", vbCritical, "Warning"
40          isAppDriverOK = False
50       End If
End Function

