VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AppDriver"
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
'  XerSys [ 05/20/2004 10:28 am ]
'     Start revising this object. i adopt the idea of the system environment
'  variables, definition and loading from the TBoxPlus v97.01 of the late Mr. Sison and
'  combine it with the idea of my dear friend Marlon Sayson (with his consent).
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€

Option Explicit

Private Const xeDriverID As String = "GRider"
Private Const xeSignature As String = "08220326"

' system environment variables
Private p_sClientId As String
Private p_sClientNm As String
Private p_sAddressx As String
Private p_sTownName As String
Private p_sZippCode As String
Private p_sProvName As String
Private p_sTelNoxxx As String
Private p_sFaxNoxxx As String
Private p_sApproved As String
Private p_sSysAdmin As String
Private p_sProdctID As String
Private p_sProdctNm As String
Private p_sNetWarex As String
Private p_sMachinex As String
Private p_sApplPath As String
Private p_dSysDatex As Date
Private p_nNetError As Long
Private p_sBranchCd As String
Private p_dCapturex As Date
Private p_sComptrID As String
Private p_sMenuName As String

' system user variables
Private p_oSysUser As SystemUser
Private p_ocnMain As Connection
Private p_oMod As New MainModules
Private p_oMDIForm As MDIForm

' skin variables
Private p_sSkinCode As String
Private p_sFormPict As String
Private p_sQSImage As String
Private p_sFldName As String
Private p_anColor() As Long

Private p_sCompName As String
Private p_sProcName As String

Private pbErrorLog As Boolean
Private psModlName As String

Property Get Connection() As Connection
10       Set Connection = p_ocnMain
End Property

Property Get ClientId() As String
10       ClientId = p_sClientId
End Property

Property Get ClientName() As String
10       ClientName = p_sClientNm
End Property

Property Get Address() As String
10       Address = p_sAddressx
End Property

Property Get TownCity() As String
10       TownCity = p_sTownName
End Property

Property Get ZippCode() As String
10       ZippCode = p_sZippCode
End Property

Property Get Province() As String
10       Province = p_sProvName
End Property

Property Get TelNo() As String
10       TelNo = p_sTelNoxxx
End Property

Property Get FaxNo() As String
10       FaxNo = p_sFaxNoxxx
End Property

Property Get Approved() As String
10       Approved = p_sApproved
End Property

Property Get SysAdmin() As String
10       SysAdmin = p_sSysAdmin
End Property

Property Get ProductID() As String
10       ProductID = p_sProdctID
End Property

Property Get ProductName() As String
10       ProductName = p_sProdctNm
End Property

Property Get NetWarex() As String
10       NetWarex = p_sNetWarex
End Property

Property Get Machinex() As String
10       Machinex = p_sMachinex
End Property

Property Get AppPath() As String
10       AppPath = p_sApplPath
End Property

Property Get SysDate() As Date
10       SysDate = p_dSysDatex
End Property

Property Get NetError() As Long
10       NetError = p_nNetError
End Property

Property Get BranchCode() As String
10       BranchCode = p_sBranchCd
End Property

Property Get UserID() As String
10       UserID = p_oSysUser.UserID
End Property

Property Get LogName() As String
10       LogName = p_oSysUser.LogName
End Property

Property Get UserName() As String
10       UserName = p_oSysUser.UserName
End Property

Property Get EmployeeNo() As String
10       EmployeeNo = p_oSysUser.EmployeeNo
End Property

Property Get UserLevel() As Integer
10       UserLevel = p_oSysUser.UserLevel
End Property

Property Get ComputerName() As String
10       ComputerName = p_sCompName
End Property

Property Get ComputerID() As String
10       ComputerID = p_sComptrID
End Property

Property Get MenuName() As String
10       MenuName = p_sMenuName
End Property

Property Let MenuName(Value As String)
10       p_sMenuName = Value
End Property

Property Set MDIMain(Value As Object)
10       Set p_oMDIForm = Value
   
20       SetMDIMain
End Property

Property Get MDIMain() As Object
10       Set MDIMain = p_oMDIForm
End Property

Property Get CaptureDate() As Date
10       CaptureDate = p_dCapturex
End Property

Function LogIn(sProduct As String) As Boolean
   ' sProduct will be the controlling variable. If sProduct contains a valid value
   '  then it will not allow the user to select an application to be executed; if it
   '  contains an empty string then selection is available and the selected application
   '  will be selected.
10       Dim loProduct As Recordset
20       Dim lors As Recordset
30       Dim loForm As frmLogin
40       Dim loCrypt As Crypto
50       Dim lsUser As String
60       Dim lsPassword As String
70       Dim lsOldProc As String
80       Dim lsSQL As String
90       Dim lnCtr As Long
100      Dim lbLogIn As Boolean
   
110      lsOldProc = ProcName("LogIn")
120      lbLogIn = False
130      'On Error GoTo errProc
   
140      Set loProduct = New Recordset
150      Set lors = New Recordset
160      Set loForm = New frmLogin
170      Set loCrypt = New Crypto
180      loCrypt.Signature = xeSignature

190      frmLogin.lblCompany = p_sClientNm
200      frmLogin.lblAddress = p_sAddressx
210      If Not (IsNull(p_sTelNoxxx) Or p_sTelNoxxx = Empty) Then
220         frmLogin.lblTelNo = "Tel. No:" & p_sTelNoxxx
230      End If
   
240      If sProduct = Empty Then
250         lsSQL = "SELECT" & _
            "  a.sProdctID" & _
            ", a.sProdctNm" & _
            ", b.sApplName" & _
            ", b.sApplPath" & _
         " FROM xxxSysObject a" & _
            ", xxxAppObject b" & _
         " WHERE a.sProdctID = b.sProdctID" & _
            " AND a.sClientID = " & p_oMod.strParm(p_sClientId) & _
            " AND a.sProdctID <> " & p_oMod.strParm(xeDriverID)
260      Else
270         lsSQL = "SELECT" & _
            "  a.sProdctID" & _
            ", a.sProdctNm" & _
            ", b.sApplName" & _
            ", b.sApplPath" & _
         " FROM xxxSysObject a" & _
            ", xxxAppObject b" & _
         " WHERE a.sProdctID = b.sProdctID" & _
            " AND sClientID = " & p_oMod.strParm(p_sClientId) & _
            " AND a.sProdctID = " & p_oMod.strParm(sProduct)
280      End If
290      loProduct.Open lsSQL, p_ocnMain, adOpenStatic, , adCmdText
         
300      If loProduct.EOF Then GoTo endProc
   
310      With loForm
320         loForm.lblCompany = p_sClientNm
330         loForm.lblAddress = Trim(p_sAddressx) & ", " & Trim(p_sTownName) & ", " & _
      Trim(p_sProvName) & " " & Trim(p_sZippCode) & ", "
340         loForm.lblTelNo = "Tel No: " & p_sTelNoxxx
350         Do
360            .Combo1.AddItem (loProduct("sProdctNm"))
370            loProduct.MoveNext
380         Loop While Not loProduct.EOF
390         If sProduct <> Empty Then .Combo1.ListIndex = 0
400      End With
   
410      lnCtr = 0
420      Do
430         loForm.Show 1
440         If loForm.Cancel = True Then
         ' user either clicked the cancel button or the close key
450            lnCtr = 2
460         Else
         ' Encrypt the username and password
470            loCrypt.InBuffer = loForm.txtPassword
480            loCrypt.Encrypt
490            lsPassword = loCrypt.OutBuffer
        
500            loCrypt.InBuffer = LCase(loForm.txtUserName)
510            loCrypt.Encrypt
520            lsUser = loCrypt.OutBuffer
         
530            With p_oMod
            ' search for the user
540               lsSQL = "SELECT *" & _
                     " FROM xxxSysUser" & _
                     " WHERE sLogNamex = " & .strParm(lsUser) & _
                        " AND sPassword = " & .strParm(lsPassword)
550            End With
         
560            If lors.State = adStateOpen Then lors.Close
570            lors.Open lsSQL, p_ocnMain, , , adCmdText
580            If lors.EOF Then
590               MsgBox "Verify your log name and/or password", vbCritical, "Login Error"
600            Else
610               lbLogIn = True
620            End If
630         End If
640         lnCtr = lnCtr + 1
650      Loop Until (lbLogIn = True) Or (lnCtr = 3)
   
660      If lbLogIn = False Then GoTo endProc
   
670      loProduct.MoveFirst
680      loProduct.Find "sProdctNm = " & p_oMod.strParm(loForm.Combo1.Text), 0, adSearchForward, 0
   
690      If loProduct.EOF Then GoTo endProc

   ' check first if user is allowed to use the application
700      If lors("cUserStat") = xeUserStateSuspended Then
710         MsgBox "User is currently suspended!!!" & vbCrLf & _
               "Application used is not allowed!!!", vbCritical, "Warning"
720         GoTo endProc
730      End If
   
740      If lors("cUserType") = xeUserTypeLocal Then
750         If lors("sProdctID") <> loProduct("sProdctID") Then
760            lors.Find "sProdctID = " & _
                  p_oMod.strParm(loProduct("sProdctID")), 1, adSearchForward
770            If lors.EOF Then
780               MsgBox "User is not a member of this application!!!" & vbCrLf & _
                     "Application used is not allowed!!!", vbCritical, "Warning"
790               GoTo endProc
800            End If
810         End If
820      End If
   
830      If sProduct <> Empty Then
840         If p_oSysUser.LogUser(lors("sUserIDxx")) = False Then GoTo endProc
850         getSkin p_oSysUser.Skin
860         LogIn = True
870         GoTo endProc
880      End If

890      If Shell(loProduct("sApplPath") & "\" & _
            loProduct("sApplName") & " " & loProduct("sProdctID") & " " & _
            lors("sUserIDxx"), vbMaximizedFocus) = 0 Then GoTo endProc
900      LogIn = True
   
endProc:
910      Unload loForm
920      Set loCrypt = Nothing
930      Set lors = Nothing
940      Set loForm = Nothing
950      Set loProduct = Nothing
   
960      ProcName lsOldProc
970      Exit Function
errProc:
980      With Err
990         ShowError psModlName, ProcName(lsOldProc), .Number, .Description, Erl
1000        .Raise .Number, .Source, .Description, .HelpFile, .HelpContext
1010        .Clear
1020     End With
End Function

Function LoadEnv(sProduct As String, Optional sUserID As Variant) As Boolean
10       Dim lors As New Recordset
20       Dim loCrypt As Crypto
30       Dim loINI As INIFile
40       Dim lsServerNm As String
50       Dim lsDatabase As String
60       Dim lsPassword As String
70       Dim lsUserName As String
80       Dim lsClientCd As String
90       Dim lsOldProc As String
100      Dim lsSQL As String
   
110      'On Error GoTo errProc
120      LoadEnv = False
130      p_sProcName = "Load Environment"
   
140      Set loINI = New INIFile
150      Set lors = New Recordset
160      Set loCrypt = New Crypto
170      loCrypt.Signature = xeSignature
   
   'prepare the ini file that we will used in comparing the records...
180      loINI.FileName = Environ("windir") & "\GhostRider.ini"
190      If Not loINI.IsFileExist Then
200         Err.Raise 53
210         Err.Clear
220      End If
   
230      lsServerNm = loINI.getTextValue(sProduct, "ServerName")
240      lsDatabase = loINI.getTextValue(sProduct, "Database")
250      lsUserName = loINI.getTextValue(sProduct, "UserName")
260      lsPassword = loINI.getTextValue(sProduct, "Password")
270      lsClientCd = loINI.getTextValue(sProduct, "ClientID")
   
280      p_sComptrID = loINI.getTextValue(sProduct, "ComputerID")
   ' check first variables contain valid data
               
290      If lsServerNm = Empty Or _
         lsDatabase = Empty Or _
         lsUserName = Empty Or _
         lsClientCd = Empty Then
300         Err.Raise 32001, , "INI file is corrupt"
310      End If

   ' decrypt the user and the password
320      loCrypt.InBuffer = lsUserName
330      loCrypt.Decrypt
340      lsUserName = loCrypt.OutBuffer

350      If lsPassword <> Empty Then
360         loCrypt.InBuffer = lsPassword
370         loCrypt.Decrypt
380         lsPassword = loCrypt.OutBuffer

390         lsSQL = "Provider=SQLOLEDB.1;" & _
                  "Persist Security Info=False;" & _
                  "User ID=" & lsUserName & ";" & _
                  "Password=" & lsPassword & ";" & _
                  "Initial Catalog=" & lsDatabase & ";" & _
                  "Data source=" & lsServerNm
400      Else
410         lsSQL = "Provider=SQLOLEDB.1;" & _
                  "Persist Security Info=False;" & _
                  "User ID=" & lsUserName & ";" & _
                  "Initial Catalog=" & lsDatabase & ";" & _
                  "Data source=" & lsServerNm
420      End If
   
   'Set the connection and prepare the record to be extracted...
430      p_ocnMain.CursorLocation = adUseClient
440      p_ocnMain.CommandTimeout = 0
450      p_ocnMain.Open lsSQL
   
   ' after setting the connection, set the system user object
460      Set p_oSysUser.AppDriver = Me
      
470      With p_oMod
   ' get the computer name first
480         p_sCompName = .getComputerName()

490         lsSQL = "SELECT" & _
                  "  a.sClientID" & _
                  ", a.sClientNm" & _
                  ", a.sAddressx" & _
                  ", a.sTownName" & _
                  ", a.sZippCode" & _
                  ", a.sProvName" & _
                  ", a.sTelNoxxx" & _
                  ", a.sFaxNoxxx" & _
                  ", a.sApproved" & _
                  ", a.sSysAdmin" & _
                  ", a.sProdctID" & _
                  ", a.sProdctNm" & _
                  ", a.sNetWarex" & _
                  ", a.sMachinex" & _
                  ", a.dSysDatex" & _
                  ", a.dLicencex" & _
                  ", a.nNetError" & _
                  ", a.sBranchCd" & _
                  ", a.sSkinCode" & _
                  ", a.dCapturex" & _
                  ", b.sApplPath"
500         lsSQL = lsSQL & _
               " FROM xxxSysObject a" & _
                  ", xxxAppObject b" & _
               " WHERE a.sProdctID = b.sProdctID" & _
                  " AND a.sClientID = " & .strParm(lsClientCd) & _
                  " AND a.sProdctID = " & .strParm(sProduct)
510      End With
   
520      lors.Open lsSQL, p_ocnMain, adOpenStatic, adLockOptimistic, adCmdText
   
530      If lors.EOF Then
540         Err.Raise 32001, , "System file is corrupt"
550         Err.Clear
560      End If
   
570      p_nNetError = lors("nNetError").Value
   
   ' check error count
580      If p_nNetError > 200 Then
590         Err.Raise 32000, , "Application has reached 200 application error limit."
600         Err.Clear
610      End If
   
620      p_sClientId = lors("sClientId")
630      p_sClientNm = lors("sClientNm")
640      p_sAddressx = lors("sAddressx")
650      p_sTownName = lors("sTownName")
660      p_sZippCode = lors("sZippCode")
670      p_sProvName = lors("sProvName")
680      p_sTelNoxxx = lors("sTelNoxxx")
690      p_sFaxNoxxx = lors("sFaxNoxxx")
700      p_sApproved = lors("sApproved")
710      p_sSysAdmin = lors("sSysAdmin")
720      p_sProdctID = lors("sProdctID")
730      p_sProdctNm = lors("sProdctNm")
740      p_sNetWarex = lors("sNetWarex")
750      p_sMachinex = lors("sMachinex")
760      p_sApplPath = lors("sApplPath")
770      p_dSysDatex = lors("dSysDatex")
780      p_nNetError = lors("nNetError")
790      p_sBranchCd = lors("sBranchCd")
800      p_dCapturex = lors("dCapturex")
   
   ' decrypt the encrypted fields
810      loCrypt.InBuffer = p_sNetWarex
820      loCrypt.Decrypt
830      p_sNetWarex = loCrypt.OutBuffer
840      loCrypt.InBuffer = p_sMachinex
850      loCrypt.Decrypt
860      p_sMachinex = loCrypt.OutBuffer
   
   ' check the authors signature
870      If isSignatureOK = False Then
880         Err.Raise 32002, , "System File is corrupted."
890         Err.Clear
900      End If
   
   ' check system date
910      If lors("dSysDatex").Value > ServerDate Then
920         lors("nNetError").Value = p_nNetError + 1
930         lors.UpdateBatch adAffectCurrent
940         MsgBox "Back Date is not allowed!!!" & vbCrLf & _
               App.EXEName & " " & App.Major & "." & App.Minor & " will now quit!!!", _
               vbInformation, "Warning"
950         GoTo endProc
'960      Else
'970         lors("dSysDatex").Value = ServerDate
'980         lors.UpdateBatch adAffectAllChapters
990      End If
1000     lors.Close
   
1010     If Not IsMissing(sUserID) Then
1020        LoadEnv = p_oSysUser.LogUser(sUserID)
1030        If LoadEnv = False Then GoTo endProc
1040        getSkin p_oSysUser.Skin
1050     End If
1060     LoadEnv = True
   
endProc:
1070     Set lors = Nothing
1080     Set loCrypt = Nothing
1090     Set loINI = Nothing
   
1100     ProcName lsOldProc
1110     Exit Function
errProc:
1120     With Err
1130        ShowError psModlName, ProcName(lsOldProc), .Number, .Description, Erl
1140     End With
End Function

Sub ShowError(sModule As String, sProcedure As String, _
   nErrorNumber As Long, sErrorDescription As String, Optional xerl As Integer = 0)
   ' Purpose  :  Provide a central error-handling mechanism.
   ' Accepts  :  sModule - the module in which the error was
   '                encountered (form, class, standard, and so on.)
   '             sProcedure - the name of the procedure in which
   '                the error was encountered.
   '             nErrorNumber - the numeric identifier of the
   '                error.
   '             sErrorDescription - the text description of
   '                the error.
10       Dim lsMessage As String
20       Dim lsCaption As String
30       Dim lnLogFile As Integer

   ' Obtain a free file handle.
40       lnLogFile = FreeFile
   
   ' Open the error log text file in Append mode.
   ' If the file doesn't exist, the Open statement
   ' creates it.
50       Open Environ("windir") & "\" & "GMCSys.log" For Append As #lnLogFile
   
   ' Write the header.
60       Print #lnLogFile, "*** Error Encountered " & VBA.Now & " ***"

   ' Write the pertinent error information to the log file.
70       Print #lnLogFile, "Error: " & nErrorNumber
80       Print #lnLogFile, "Description: " & sErrorDescription
90       Print #lnLogFile, "Procedure: " & sProcedure
100      Print #lnLogFile, "Module: " & sModule
110      Print #lnLogFile, "User: " & p_oSysUser.LogName
120      Print #lnLogFile, "Line Number: " & xerl
   
   ' Write a blank line to the log file.
130      Print #lnLogFile, ""
   
   ' Close the error log text file.
140      Close #lnLogFile

   ' Build the error message for display to the user.
150      lsMessage = "Error: " & sErrorDescription & vbCrLf & vbCrLf & _
               "Module: " & sModule & vbCrLf & _
               "Procedure: " & sProcedure & vbCrLf & _
               "Line Number: " & xerl & vbCrLf & vbCrLf & _
               "Please notify the GMC SEG or Rex Adversalo for assistance." & vbCrLf & vbCrLf & _
               "Be sure to note the information shown in this dialog box " & _
               "as well as an explanation of what you were doing when this error occurred."

   ' Build the caption for the message box. The caption shows
   ' the version number of the program.
160      lsCaption = App.EXEName & " " & App.Major & _
                "." & App.Minor & " - Unexpected Error!"

170      MsgBox lsMessage, vbCritical, lsCaption

   ' after displaying error, add one to user error
180      If p_oSysUser.AllowLock Then p_oSysUser.LockUser
   
endProc:
190      Close #lnLogFile
200      Exit Sub
End Sub

'Function Authorize() As Boolean
'   Dim lfrmAuthorize As frmAuthorization
'   Dim loCrypt As clsCryptoFilterBox
'   Dim lsPassword As String
'   Dim lsLogName As String
'   Dim psSQL As String
'
'   Authorize = False
'
'   Set lfrmAuthorize = New frmAuthorization
'   Set loCrypt = New clsCryptoFilterBox
'   lfrmAuthorize.Show 1
'
'   If lfrmAuthorize.Cancelled Then Exit Function
'
'   loCrypt.InBuffer = lfrmAuthorize.txtPassword.Text
'   loCrypt.Encrypt
'   lsPassword = loCrypt.OutBuffer
'
'   loCrypt.InBuffer = lfrmAuthorize.txtUserName.Text
'   loCrypt.Encrypt
'   lsLogName = loCrypt.OutBuffer
'
'   Set lfrmAuthorize = Nothing
'
'   If lsPassword <> sNetWarex Then Exit Function
''   If lsLogName <> sNetWareX Then Exit Function
'
'   Authorize = True
'End Function
'
Function ProcName(Optional Value As Variant) As String
10       ProcName = p_sProcName
20       p_sProcName = IIf(IsMissing(Value), p_sProcName, Value)
End Function

Function getUserName(ByVal sUserID As String) As String
10       Dim lors As Recordset
20       Dim lsOldProc As String
   
30       lsOldProc = ProcName("getUserName")
40       getUserName = Empty
50       'On Error GoTo errProc
   
60       If sUserID = Empty Then GoTo endProc
   
70       Set lors = New Recordset
80       lors.Open "SELECT sUserName" & _
            " FROM xxxSysUser" & _
            " WHERE sUserIDxx = " & p_oMod.strParm(sUserID), p_ocnMain, , , adCmdText
   
90       If Not lors.EOF Then getUserName = p_oMod.Decrypt(lors("sUserName"), Machinex)
   
endProc:
100      ProcName lsOldProc
110      Exit Function
errProc:
120      With Err
130         ShowError psModlName, ProcName(lsOldProc), .Number, .Description
140         .Raise .Number, .Source, .Description, .HelpFile, .HelpContext
150         .Clear
160      End With
End Function

Function getLogName(ByVal sUserID As String) As String
10       Dim lors As Recordset
20       Dim lsOldProc As String
   
30       lsOldProc = ProcName("getLogName")
40       getLogName = Empty
50       'On Error GoTo errProc
   
60       If sUserID = Empty Then GoTo endProc
   
70       Set lors = New Recordset
80       lors.Open "SELECT sLogNamex" & _
            " FROM xxxSysUser" & _
            " WHERE sUserIDxx = " & p_oMod.strParm(sUserID), p_ocnMain, , , adCmdText
   
90       If Not lors.EOF Then getLogName = p_oMod.Decrypt(lors("sLogNamex"), Machinex)
   
endProc:
100      ProcName lsOldProc
110      Exit Function
errProc:
120      With Err
130         ShowError psModlName, ProcName(lsOldProc), .Number, .Description
140         .Raise .Number, .Source, .Description, .HelpFile, .HelpContext
150         .Clear
160      End With
End Function

Function getColor(ColorCode As String) As Long
10       Dim lnCode As Integer
20       Dim lnTemp As Double
   
30       lnCode = InStr(1, p_sFldName, ColorCode, vbTextCompare)
40       lnTemp = lnCode / 4
50       lnCode = Int(lnTemp)
60       getColor = p_anColor(lnCode)
End Function

Function getPicture(FormType As Integer) As String
10       Select Case FormType
   Case xeFormMaintenance
20          getPicture = p_sFormPict
30       Case xeFormTransaction
40          getPicture = p_sFormPict
50       Case xeFormQuickSearch
60          getPicture = p_sQSImage
70       End Select
End Function
   
Sub getSkin(lsSkin As String)
10       Dim loSkin As Recordset
20       Dim lsOldProc As String
30       Dim lsSQL As String
40       Dim lnCtr As Integer

50       lsOldProc = ProcName("getSkin")
60       'On Error GoTo errProc
   
70       Set loSkin = New Recordset

80       lsSQL = "SELECT *" & _
            " FROM xxxSkin" & _
            " WHERE sSkinCode = " & p_oMod.strParm(lsSkin)
90       loSkin.Open lsSQL, p_ocnMain, , , adCmdText
   
100      ReDim p_anColor(loSkin.Fields.Count - 1) As Long
110      p_sFldName = ""
120      If loSkin.EOF Then
130         For lnCtr = 2 To loSkin.Fields.Count - 3
140            p_sFldName = p_sFldName & Right(loSkin(lnCtr).Name, 3) & ";"
150            p_anColor(lnCtr - 2) = 0
160         Next
170         p_sQSImage = ""
180         p_sFormPict = ""
190      Else
200         For lnCtr = 2 To loSkin.Fields.Count - 3
210            p_sFldName = p_sFldName & Right(loSkin(lnCtr).Name, 3) & ";"
220            p_anColor(lnCtr - 2) = loSkin(lnCtr)
230         Next
240         p_sQSImage = loSkin("sQSImagex")
250         p_sFormPict = loSkin("sTitleBar")
260      End If
   
endProc:
270      ProcName (lsOldProc)
280      Set loSkin = Nothing
290      Exit Sub
errProc:
300      With Err
310         ShowError psModlName, ProcName(lsOldProc), .Number, .Description, Erl
320         .Raise .Number, .Source, .Description, .HelpFile, .HelpContext
330         .Clear
340      End With
End Sub

Function RegisDelete(ByVal sQuery As String) As Boolean
10       Dim lsSQL As String
20       Dim lnRow As Long
   
30       If sQuery = Empty Then Exit Function
40       lsSQL = "INSERT INTO xxxDeletedRec " & _
               "( sTransNox" & _
               ", sBranchCd" & _
               ", sStatemnt" & _
               ", sModified" & _
               ", dModified" & _
            " ) VALUES ( " & _
               p_oMod.strParm(p_oMod.GetNextCode("xxxDeletedRec", "sTransNox", True, _
                     p_ocnMain, True, p_sBranchCd)) & _
               ", " & p_oMod.strParm(p_sBranchCd) & _
               ", " & p_oMod.strParm(sQuery) & _
               ", " & p_oMod.strParm(p_oSysUser.UserID) & _
               ", GETDATE()" & " )"

50       p_ocnMain.Execute lsSQL, lnRow, adCmdText
60       If lnRow = 0 Then
70          Err.Raise "30013", "AppDriver", "Unable to Register SQL Query Statement"
80          Exit Function
90       End If
100      RegisDelete = True
End Function

Private Sub SetMDIMain()
10       Dim lors As New Recordset
20       Dim loCtrl As Control
30       Dim lsOldProc As String
40       Dim lsSQL As String
50       Dim lsName As String
60       Dim lnCtr As Integer
   
70       'On Error GoTo errProc
80       p_sProcName = "Set MDIMain"
   
90       If UserLevel = 128 Then GoTo endProc
100      Set lors = New Recordset
110      For Each loCtrl In p_oMDIForm.Controls
120         If TypeName(loCtrl) = "Menu" Then
130            lsName = loCtrl.Name
140            lsSQL = "SELECT * FROM xxxMenuObject" & _
                  " WHERE sProdctID = " & p_oMod.strParm(ProductID) & _
                     " AND sMenuName = " & p_oMod.strParm(lsName)
150            lors.Open lsSQL, Connection, , , adCmdText
         
160            If Not lors.EOF Then
170               loCtrl.Tag = lors("sRemarksx") & _
                           "»" & lors("nAddRight") & _
                           "»" & lors("nUpdRight") & _
                           "»" & lors("nCanRight") & _
                           "»" & lors("nDelRight")
            
180               If (lors("nUserRght") And UserLevel) = 0 Then
190                  On Error Resume Next
200                  If loCtrl.Visible Then loCtrl.Visible = False
210                  'On Error GoTo errProc
220               End If
230            End If
240            lors.Close
250         End If
260      Next
   
endProc:
270      ProcName lsOldProc
   
280      Set lors = Nothing
290      Exit Sub
errProc:
300      With Err
310         ShowError psModlName, ProcName(lsOldProc), .Number, .Description, Erl
320         .Raise .Number, .Source, .Description, .HelpFile, .HelpContext
330         .Clear
340      End With
End Sub

Private Sub Class_Initialize()
10       pbErrorLog = False
20       psModlName = "AppDriver"
   
30       Set p_ocnMain = New Connection
40       Set p_oSysUser = New SystemUser
End Sub

Public Function ServerDate() As Variant
10       Dim lors As Recordset
   
20       Set lors = p_ocnMain.Execute("SELECT GETDATE()")
30       ServerDate = lors(0)
   
40       Set lors = Nothing
End Function

Private Sub Class_Terminate()
10       p_ocnMain.Close
   
20       Set p_ocnMain = Nothing
30       Set p_oSysUser = Nothing
40       Set p_oMod = Nothing
End Sub

Private Function isSignatureOK() As Boolean
10       Dim lors As Recordset
20       Dim lsSQL As String
30       Dim lsOldProc As String
   
40       lsOldProc = ProcName("isSignatureOK")
50       isSignatureOK = False
60       'On Error GoTo errProc
   
70       If p_sMachinex <> xeSignature Then GoTo endProc
   
   ' check the existence of the author and the system admin
80       lsSQL = "SELECT sUserIDxx" & _
               ", sUserName" & _
               ", sLogNamex" & _
            " FROM xxxSysUser" & _
            " WHERE sUserIDxx in (" & _
               p_oMod.strParm(p_sMachinex) & ", " & p_oMod.strParm(p_sNetWarex) & ")"
   
90       Set lors = New Recordset
100      lors.Open lsSQL, p_ocnMain, , , adCmdText
   
110      With lors
120         If .RecordCount <> 2 Then GoTo endProc
      
130         Do
140            If .Fields("sUserIDxx") = p_sNetWarex Then
150               If Trim(.Fields("sLogNamex")) <> Trim(p_sSysAdmin) Then GoTo endProc
160            End If
170            .MoveNext
180         Loop While Not .EOF
190      End With
200      isSignatureOK = True
   
endProc:
210      ProcName lsOldProc
220      Set lors = Nothing
230      Exit Function
errProc:
240      With Err
250         ShowError psModlName, ProcName(lsOldProc), .Number, .Description, Erl
260         .Raise .Number, .Source, .Description, .HelpFile, .HelpContext
270         .Clear
280      End With
End Function


