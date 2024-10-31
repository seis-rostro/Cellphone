VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MainModules"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Const MAX_COMPUTERNAME_LENGTH As Long = 31
Private Declare Function GetCompName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare Function GetFocus Lib "user32.dll" () As Long

Public Type FieldProperty
   Name As String
   Value As Variant
   Type As String
   Size As Integer
   Help As String
   Format As String
   Reference As Boolean
   Initialize As Boolean
   Empty As Boolean
   Query As String
   QColumn As String
   QTitle As String
   QFormat As String
   QType As String
   QReference As String
End Type

Public Enum xeLogical
   xeUnknown = -1
   xeNo = 0
   xeYes = 1
End Enum

Public Enum xeTransactionStatus
   xeStateOpen = 0
   xeStateClosed = 1
   xeStatePosted = 2
   xeStateCancelled = 3
   xeStateUnknown = 4
End Enum

Public Enum xeRecordStatus
   xeRecStateInactive = 0
   xeRecStateActive = 1
   xeRecStateUnknown = 2
End Enum

Public Enum xeUserStatus
   xeUserStateSuspended = 0
   xeUserStateActive = 1
End Enum

Public Enum xeUserType
   xeUserTypeLocal = 0
   xeUserTypeGlobal = 1
End Enum

Public Enum xeLogStatus
   xeLogStateOut = 0
   xeLogStateIn = 1
End Enum

Public Enum xeLockStatus
   xeLockStateUnLock = 0
   xeLockStateLock = 1
End Enum

Public Enum xeEditMode
   xeModeReady = 0
   xeModeAddNew = 1
   xeModeUpdate = 2
   xeModeDelete = 3
End Enum

Public Enum xeUserRights
   xeEncoder = 1
   xeSupervisor = 2
   xeManager = 4
   xeAudit = 8
   xeSysAdmin = 16
   xeSysOwner = 32
   xeEngineer = 64
   xeSysMaster = 128
End Enum

Public Enum xeFormType
   xeFormMaintenance = 0
   xeFormTransaction = 1
   xeFormQuickSearch = 2
   xeFormMessageBox = 3
   xeFormLedger = 4
   xeFormTransDetail = 5
   xeFormTransMaintenance = 6
   xeFormUnknown = 7
End Enum

Public Enum MCLocation
   xeLocWarehouse = 0
   xeLocBranch = 1
   xeLocSupplier = 2
   xeLocCustomer = 3
   xeLocUnknown = 4
End Enum

Public Enum ClientType
   xeClientCustomer = 0
   xeClientSupplier = 1
   xeClientTrucking = 2
End Enum

Public Enum AccoutStat
   xeActStatActive = 0
   xeActStatClosed = 1
   xeActStatDead = 2
   xeActStatImpounded = 3
   xeActStatDiscarded = 4
End Enum

Public Function strParm(ByVal sStrParm As String) As String
10       strParm = "'" & Replace(sStrParm, "'", "''") & "'"
End Function

Public Function dateParm(ByVal dDateParm As Date) As String
10       dateParm = "'" & Format(dDateParm, "mm/dd/yyyy hh:mm:ss") & "'"
End Function

Public Function KwikSearch(oAppDriver As AppDriver, _
                           ByVal sSource As String, _
                           Optional sFieldName As Variant, _
                           Optional sFieldDesc As Variant, _
                           Optional sFieldPict As Variant, _
                           Optional bSearch As Variant, _
                           Optional sFieldCrit As Variant) As Variant
10       Dim loLookUp As frmLookUp
20       Dim lsOldProc As String
25       Dim lnRecCount As Integer
   
30       lsOldProc = oAppDriver.ProcName("KwikSearch")
40       'On Error GoTo errProc
   
50       Set loLookUp = New frmLookUp
60       Set loLookUp.AppDriver = oAppDriver
   
70       loLookUp.SQLSource = sSource
80       If Not IsMissing(sFieldName) Then
90          If sFieldName <> "" Then loLookUp.FldName = sFieldName
100      End If
110      If Not IsMissing(sFieldDesc) Then
120         If sFieldDesc <> "" Then loLookUp.FldTitle = sFieldDesc
130      End If
140      If Not IsMissing(sFieldPict) Then
150         If sFieldPict <> "" Then loLookUp.FldFormat = sFieldPict
160      End If
170      If Not IsMissing(bSearch) Then loLookUp.showSearch = bSearch
180      If Not IsMissing(sFieldCrit) Then
190         If sFieldCrit <> "" Then loLookUp.FldCriteria = sFieldCrit
200      End If

210      loLookUp.Show 1
220      KwikSearch = loLookUp.SelectedItem
   
endProc:
230      Set loLookUp = Nothing
   
240      oAppDriver.ProcName lsOldProc
250      Exit Function
errProc:
260      With Err
270         oAppDriver.ShowError "MainModules", _
                           oAppDriver.ProcName(lsOldProc), .Number, .Description, Erl
280         .Raise .Number, .Source, .Description, .HelpFile, .HelpContext
290         .Clear
300      End With
End Function

Public Function KwikBrowse(oAppDriver As AppDriver, _
                           oSource As Recordset, _
                           Optional sFieldName As Variant, _
                           Optional sFieldDesc As Variant, _
                           Optional sFieldPict As Variant, _
                           Optional bSearch As Variant, _
                           Optional sFieldCrit As Variant) As Variant
10       Dim loLookUp As frmLookUp
20       Dim lsOldProc As String
   
30       lsOldProc = oAppDriver.ProcName("KwikBrowse")
40       'On Error GoTo errProc
   
50       Set loLookUp = New frmLookUp
60       Set loLookUp.AppDriver = oAppDriver
   
70       Set loLookUp.RowSource = oSource
80       If Not IsMissing(sFieldName) Then
90          If sFieldName <> "" Then loLookUp.FldName = sFieldName
100      End If
110      If Not IsMissing(sFieldDesc) Then
120         If sFieldDesc <> "" Then loLookUp.FldTitle = sFieldDesc
130      End If
140      If Not IsMissing(sFieldPict) Then
150         If sFieldPict <> "" Then loLookUp.FldFormat = sFieldPict
160      End If
170      If Not IsMissing(bSearch) Then loLookUp.showSearch = bSearch
180      If Not IsMissing(sFieldCrit) Then
190         If sFieldCrit <> "" Then loLookUp.FldCriteria = sFieldCrit
200      End If
   
'201      If loLookUp.LoadList = False Then GoTo endProc
'      MsgBox "No Record is Availble for Browsing!!!", vbInformation, "Notice"
'      GoTo endProc
'   End If

210      loLookUp.Show 1
220      KwikBrowse = loLookUp.SelectedItem
   
endProc:
230      Set loLookUp = Nothing
240      oAppDriver.ProcName lsOldProc
250      Exit Function
errProc:
260      With Err
270         oAppDriver.ShowError "MainModules", _
                           oAppDriver.ProcName(lsOldProc), .Number, .Description, Erl
280         .Raise .Number, .Source, .Description, .HelpFile, .HelpContext
290         .Clear
300      End With
End Function

Public Function isAppDriverOK(oAppDriver) As Boolean
10       isAppDriverOK = True
20       If oAppDriver Is Nothing Then
30          MsgBox "Application is not yet set!!!" & vbCrLf & _
               "Can not further proceed!!!", vbCritical, "Warning"
40          isAppDriverOK = False
50       End If
End Function

Public Function GetApproval(oAppDriver As AppDriver, _
                              nApproveRights As Integer, _
                              sApproveID As String, _
                              sApproveName As String) As Boolean
10       Dim lors As Recordset
20       Dim loForm As frmApproval
30       Dim loCrypt As Crypto
40       Dim loMod As MainModules
50       Dim lsUser As String
60       Dim lsPassword As String
70       Dim lsOldProc As String
80       Dim lsSQL As String
90       Dim lnCtr As Integer
100      Dim lbLogIn As Boolean
   
110      lsOldProc = oAppDriver.ProcName("GetApproval")
120      'On Error GoTo errProc
130      GetApproval = False
      
140      Set lors = New Recordset
150      Set loCrypt = New Crypto
160      loCrypt.Signature = oAppDriver.Machinex
   
170      Set loForm = New frmApproval
180      Set loMod = New MainModules
   
190      lnCtr = 0
200      lbLogIn = False
210      Do
220         loForm.AppPath = oAppDriver.AppPath
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
         
330            With loMod
            ' search for the user
340               lsSQL = "SELECT *" & _
                     " FROM xxxSysUser" & _
                     " WHERE sLogNamex = " & .strParm(lsUser) & _
                        " AND sPassword = " & .strParm(lsPassword) & _
                        " AND ( cUserType = '1' OR sProdctID = " & .strParm(oAppDriver.ProductID) & " )"
350            End With
         
360            If lors.State = adStateOpen Then lors.Close
         
370            lors.Open lsSQL, oAppDriver.Connection, , , adCmdText
380            If lors.EOF Then
390               MsgBox "Verify your log name and/or password", vbCritical, "Login Error"
400            Else
            ' check first if user has the right to approved procedure entry
410               If lors("nUserLevl") < xeSupervisor Then
420                  MsgBox "User has no right for this procedure!!!" & vbCrLf & vbCrLf & _
                        "Seek other user's assistance then try again!!!", _
                        vbCritical, "Warning"
430               Else
440                  lbLogIn = True
450               End If
460            End If
470         End If
480         lnCtr = lnCtr + 1
490      Loop Until (lbLogIn = True) Or (lnCtr = 3)
   
500      If lbLogIn = False Then GoTo endProc
      
   ' check first if user is not suspended
510      If lors("cUserStat") = xeUserStateSuspended Then
520         MsgBox "User is currently Suspended!!!" & vbCrLf & _
               "User has no Rights for Procedure Approval!!!", vbCritical, "Warning"
530         GoTo endProc
540      End If
   
550      If lors("cUserType") = xeUserTypeLocal Then
560         If Trim(lors("sProdctID")) <> Trim(oAppDriver.ProductID) Then
570            MsgBox "User is not a Member of this Application!!!" & vbCrLf & _
                  "User has no Rights for Procedure Approval!!!", vbCritical, "Warning"
580            GoTo endProc
590         End If
600      End If
   
610      nApproveRights = lors("nUserLevl")
620      sApproveID = lors("sUserIDxx")
630      sApproveName = lors("sUserName")
   
   ' decrypt User Name
640      loCrypt.InBuffer = sApproveName
650      loCrypt.Decrypt
660      sApproveName = loCrypt.OutBuffer
   
670      GetApproval = True
   
endProc:
680      Unload loForm
690      Set lors = Nothing
700      Set loCrypt = Nothing
710      Set loForm = Nothing
720      Set loMod = Nothing
   
730      oAppDriver.ProcName lsOldProc
740      Exit Function
errProc:
750      With Err
760         oAppDriver.ShowError "MainModules", _
                           oAppDriver.ProcName(lsOldProc), .Number, .Description, Erl
770         .Raise .Number, .Source, .Description, .HelpFile, .HelpContext
780         .Clear
790      End With
End Function

Public Function GetNextCode(ByVal Table As String, _
                              ByVal Field As String, _
                              ByVal YearFormat As Boolean, _
                              Connection As Connection, _
                              Optional ByBranch As Boolean = False, _
                              Optional Branch As String = "") As String
10       Dim lors As Recordset
20       Dim lsField As String
30       Dim lsSQL As String
40       Dim lnCode As Long
50       Dim lnLen As Long
   
60       If ByBranch = True And Branch = "" Then Exit Function
   
70       lsField = IIf(ByBranch, Branch, "")
80       lsField = lsField & IIf(YearFormat, Format(Date, "YY"), "")
   
90       lsSQL = "SELECT TOP 1 " & Field & _
            " FROM " & Table
100      If lsField <> "" Then
110         lsSQL = lsSQL & " WHERE " & Field & " LIKE '" & lsField & "%'"
120      End If
130      lsSQL = lsSQL & " ORDER BY " & Field & " DESC"

140      Set lors = New Recordset
150      lors.Open lsSQL, Connection, , , adCmdText
   
160      lsSQL = IIf(lors.EOF, Empty, lors(0))
170      lnLen = lors(0).DefinedSize
180      lnCode = 1
190      If lsSQL <> Empty Then lnCode = CLng(Mid(lsSQL, Len(lsField) + 1)) + 1

200      GetNextCode = lsField & Format(lnCode, String(lnLen - Len(lsField), "0"))
End Function

Public Function GetNextReference(ByVal Table As String, _
                              ByVal Field As String, _
                              ByVal Order As String, _
                              ByVal BranchField As String, _
                              ByVal Branch As String, _
                              Connection As Connection) As String
10       Dim lors As Recordset
20       Dim lsSQL As String
30       Dim lnCode As Long
40       Dim lnLen As Long
   
50       lsSQL = "SELECT TOP 1 " & Field & _
            " FROM " & Table & _
            " WHERE " & BranchField & " = " & strParm(Branch)
60       lsSQL = lsSQL & " ORDER BY " & Order & " DESC" & _
         ", " & Field & " DESC"
   
70       Set lors = New Recordset
80       lors.Open lsSQL, Connection, , , adCmdText

90       If lors.EOF Then
100         lsSQL = 0
110         lnLen = lors(0).DefinedSize
120      Else
130         lsSQL = IIf(lors(0) = "", 0, lors(0))
140         lnLen = lors(0).ActualSize
150      End If
160      lnCode = CLng(lsSQL) + 1

170      GetNextReference = Format(lnCode, String(lnLen, "0"))
End Function

' transalate the query to retrieve an empty record
Public Function getEmptyRec(ByVal sSQL As String) As String
10       Dim lnPosition As Integer
   
   ' parse first the Order by clause
20       lnPosition = InStr(1, sSQL, "ORDER BY", vbTextCompare)
30       If lnPosition > 0 Then sSQL = Trim(Left(sSQL, lnPosition - 1))
      
   ' parse the where clause too
40       lnPosition = InStr(1, sSQL, "WHERE", vbTextCompare)
50       If lnPosition > 0 Then sSQL = Trim(Left(sSQL, lnPosition - 1))
   
   ' the goal is to retrieve an empty record, so just create the simpliest query to do it
60       getEmptyRec = sSQL & " WHERE 0 = 1"
End Function

' adds the condtion passed to a query statement
Public Function AddCondition(ByVal sSQL As String, ByVal sCondition As String) As String
10       Dim lnPos As Integer
20       Dim lsOrder As String
   
   ' check first if sCondition contains a valid value
30       If sCondition = Empty Then
40          AddCondition = sSQL
50          Exit Function
60       End If

   ' parse first the Order by clause
70       lnPos = InStr(1, sSQL, "ORDER BY", vbTextCompare)
80       If lnPos > 0 Then
90          lsOrder = Trim(Mid(sSQL, lnPos))
100         sSQL = Trim(Left(sSQL, lnPos - 1))
110      End If
   
   ' then check if there is where clause
120      lnPos = InStr(1, sSQL, "WHERE", vbTextCompare)
130      If lnPos > 0 Then
140         sSQL = Trim(sSQL) & " AND " & sCondition
150      Else
160         sSQL = Trim(sSQL) & " WHERE " & sCondition
170      End If
   
180      AddCondition = Trim(sSQL) & " " & lsOrder
End Function

Public Sub SetNextFocus()
10       keybd_event &H9, 0, 0, 0
20       keybd_event &H9, 0, &H2, 0
End Sub

Public Sub SetPreviousFocus()
10       keybd_event &H10, 0, 0, 0
20       keybd_event &H9, 0, 0, 0
30       keybd_event &H10, 0, &H2, 0
End Sub

Public Sub CenterChildForm(oMDIForm As Object, ByRef oChild As Object)
10       Dim lnX, lnY As Long
   
20       lnX = oMDIForm.ScaleWidth
30       lnY = oMDIForm.ScaleHeight
   
40       oChild.Left = CLng((lnX - oChild.Width) / 2)
50       oChild.Top = CLng((lnY - oChild.Height) / 2)
   
60       If oChild.Left < 0 Then oChild.Left = 0
70       If oChild.Top < 0 Then oChild.Top = 0
End Sub

Public Function Encrypt(ByVal Code As String, Optional Signature As Variant) As String
10       Dim loCrypt As Crypto
   
20       Set loCrypt = New Crypto
   
30       If Not IsMissing(Signature) Then loCrypt.Signature = Signature
40       loCrypt.InBuffer = Code
50       loCrypt.Encrypt
60       Encrypt = loCrypt.OutBuffer
70       Set loCrypt = Nothing
End Function

Public Function Decrypt(ByVal Code As String, Optional Signature As Variant) As String
10       Dim loCrypt As Crypto
   
20       Set loCrypt = New Crypto
   
30       If Not IsMissing(Signature) Then loCrypt.Signature = Signature
40       loCrypt.InBuffer = Code
50       loCrypt.Decrypt
60       Decrypt = loCrypt.OutBuffer
70       Set loCrypt = Nothing
End Function

Public Function NumToText(ByVal dblVal As Double) As String
10        Static sOnes(0 To 9) As String
20        Static sTeens(0 To 9) As String
30        Static sTens(0 To 9) As String
40        Static sThousands(0 To 4) As String
50        Static bInit As Boolean
    
60        Dim i As Integer
70        Dim bAllZeros As Boolean
80        Dim bShowsThousands As Boolean
90        Dim sValue As String
100       Dim sBuffer As String
110       Dim sTemp As String
120       Dim iCol As Integer
130       Dim iChar As Integer
'
' Warning - this routine only handles positive values
'
140       Debug.Assert dblVal > 0
    
150       If bInit = False Then
        'Initialize array
160           bInit = True
170           sOnes(0) = "zero"
180           sOnes(1) = "one"
190           sOnes(2) = "two"
200           sOnes(3) = "three"
210           sOnes(4) = "four"
220           sOnes(5) = "five"
230           sOnes(6) = "six"
240           sOnes(7) = "seven"
250           sOnes(8) = "eight"
260           sOnes(9) = "nine"
270           sTeens(0) = "ten"
280           sTeens(1) = "eleven"
290           sTeens(2) = "twelve"
300           sTeens(3) = "thirteen"
310           sTeens(4) = "fourteen"
320           sTeens(5) = "fifteen"
330           sTeens(6) = "sixteen"
340           sTeens(7) = "seventeen"
350           sTeens(8) = "eighteen"
360           sTeens(9) = "nineteen"
370           sTens(0) = ""
380           sTens(1) = "ten"
390           sTens(2) = "twenty"
400           sTens(3) = "thirty"
410           sTens(4) = "forty"
420           sTens(5) = "fifty"
430           sTens(6) = "sixty"
440           sTens(7) = "seventy"
450           sTens(8) = "eighty"
460           sTens(9) = "ninety"
470           sThousands(0) = ""
480           sThousands(1) = "thousand"   'US numbering
490           sThousands(2) = "million"
500           sThousands(3) = "billion"
510           sThousands(4) = "trillion"
520       End If
    '
    ' Setup Error Handler
    '
530       On Error GoTo vbErrorHandler
    '
    ' Get fractional part of value (if any)
    '
540       sBuffer = "& " & Format$((dblVal - Int(dblVal)) * 100, "00") & "/100"
    '
    ' Convert main part to string
    '
550       sValue = CStr(Int(dblVal))
    
560       bAllZeros = True
    
570       For i = Len(sValue) To 1 Step -1
580           iChar = Val(Mid$(sValue, i, 1))
590           iCol = (Len(sValue) - i) + 1
        '
        'Action depends on 1's, 10's or 100's column
        '
600           Select Case (iCol Mod 3)
            Case 1  '1's position
610                   bShowsThousands = True
620                   If i = 1 Then
630                       sTemp = sOnes(iChar) & " "
640                   ElseIf Mid$(sValue, i - 1, 1) = "1" Then
650                       sTemp = sTeens(iChar) & " "
660                       i = i - 1
670                   ElseIf iChar > 0 Then
680                       sTemp = sOnes(iChar) & " "
690                   Else
700                       bShowsThousands = False
710                       If Mid$(sValue, i - 1, 1) <> "0" Then
720                           bShowsThousands = True
730                       ElseIf i > 2 Then
740                           If Mid$(sValue, i - 2, 1) <> "0" Then
750                               bShowsThousands = True
760                           End If
770                       End If
780                       sTemp = ""
790                   End If
800                   If bShowsThousands Then
810                       If iCol > 1 Then
820                           sTemp = sTemp & sThousands(iCol \ 3)
830                           If bAllZeros Then
840                               sTemp = sTemp & " "
850                           Else
860                               sTemp = sTemp & ", "
870                           End If
880                       End If
890                       bAllZeros = False
900                   End If
910                   sBuffer = sTemp & sBuffer
920               Case 2
930                   If iChar > 0 Then
940                       If Mid$(sValue, i + 1, 1) <> "0" Then
950                           sBuffer = sTens(iChar) & "-" & sBuffer
960                       Else
970                           sBuffer = sTens(iChar) & " " & sBuffer
980                       End If
990                   End If
1000              Case 0
1010                  If iChar > 0 Then
1020                      sBuffer = sOnes(iChar) & " hundred " & sBuffer
1030                  End If
1040          End Select
1050      Next i

1060      sBuffer = UCase$(Left$(sBuffer, 1)) & Mid$(sBuffer, 2)

EndNumToText:
1070      NumToText = sBuffer
1080      Exit Function

vbErrorHandler:
1090      sBuffer = "#Error#"
1100      Resume EndNumToText
End Function

Public Function getComputerName() As String
10       Dim lnLen As Long
20       Dim lsCompName As String
   
   'Create a buffer
30       lnLen = MAX_COMPUTERNAME_LENGTH + 1
40       lsCompName = String(lnLen, "X")
   
   'Get the computer name
50       GetCompName lsCompName, lnLen
   
   'get only the actual data
60       lsCompName = Left(lsCompName, lnLen)
   
70       getComputerName = lsCompName
End Function

Public Function GotFocus() As Long
10       GotFocus = GetFocus
End Function

Public Function AllowRecSave(oAppDrivr As AppDriver, _
                              cRecdStat As xeRecordStatus, _
                              nEditMode As xeEditMode, _
                              sRemarks As String, _
                              dModified As Date, _
                              sApprovID As String) As Boolean
10       Dim lnUserRght As Integer, lsUserName As String, lsUserIDxx As String
20       Dim lasRights() As String, lnRights As Integer
30       Dim lnSeek As Integer, lbSeek As Boolean
   
40       AllowRecSave = False
   
50       With oAppDrivr
60          lasRights = Split(sRemarks, "»")
      
70          lbSeek = False
80          If nEditMode = xeModeAddNew Then
         ' Add Rights
90             lnRights = CInt(lasRights(1))
100            If lnRights And .UserLevel = 0 Then
110               lnSeek = MsgBox("User Doesn't Have the Right to Create New Record!!!" & _
                              vbCrLf & "Seek Authorized User's Approval?", _
                              vbQuestion + vbYesNo + vbDefaultButton2, "Confirm")
120               If lnSeek = vbNo Then GoTo endProc
130               lbSeek = True
140            End If
150         ElseIf nEditMode = xeModeUpdate Then
         ' update Rights
160            lnRights = CInt(lasRights(2))
170            If lnRights And .UserLevel = 0 Then
180               Select Case .UserLevel
            Case xeEncoder
190                  lbSeek = True
200               Case xeSupervisor
210                  If DateDiff("d", dModified, .ServerDate) > 2 Then lbSeek = True
220               Case xeManager
230                  If DateDiff("d", dModified, .ServerDate) > 31 Then lbSeek = True
240               Case xeAudit
250                  lbSeek = True
260               Case xeSysAdmin
270                  lbSeek = True
280               End Select
290            End If
300         End If
      
310         If lbSeek = False Then
315            sApprovID = .UserID
320            AllowRecSave = True
330            GoTo endProc
340         End If
      
350         lnSeek = MsgBox("User Doesn't Have the Right to Update Record!!!" & _
                        vbCrLf & "Seek Authorized User's Approval?", _
                        vbQuestion + vbYesNo + vbDefaultButton2, "Confirm")
360         If lnSeek = vbNo Then GoTo endProc
      
370         If GetApproval(oAppDrivr, lnUserRght, lsUserIDxx, lsUserName) = False Then GoTo endProc
      
380         If nEditMode = xeModeAddNew Then
390            If lnUserRght And .UserLevel = 0 Then
400               MsgBox "Approving User Doesn't Have the Right to Create New Record!!!" & _
                  "Seek Other User's Assistance then Try Again", vbCritical, "Warning"
410               GoTo endProc
420            End If
430         ElseIf nEditMode = xeModeUpdate Then
440            If lnUserRght And .UserLevel = 0 Then
450               Select Case lnUserRght
            Case xeSupervisor
460                  If DateDiff("d", dModified, p_oAppDrivr.ServerDate) <= 2 Then lbSeek = False
470               Case xeManager
480                  If DateDiff("d", dModified, p_oAppDrivr.ServerDate) <= 31 Then lbSeek = False
490               Case xeSysAdmin
500                  If .UserLevel = xeSysAdmin Then lbSeek = False
510               End Select
            
520               If lbSeek Then
530                  MsgBox "Approving User Doesn't Have the Right to Update Record!!!" & _
                     "Seek Other User's Assistance then Try Again", vbCritical, "Warning"
540                  GoTo endProc
550               End If
560            End If
570         End If
580      End With

585      sApprovID = lsUserIDxx
590      AllowRecSave = True

endProc:
600      Exit Function
End Function

Public Function AllowTransSave(oAppDrivr As AppDriver, _
                              cTranStat As xeTransactionStatus, _
                              nEditMode As xeEditMode, _
                              sRemarks As String, _
                              dModified As Date, _
                              sApprovID As String) As Boolean
10       Dim lsOldProc As String
20       Dim lnUserRght As Integer, lsUserName As String, lsUserIDxx As String
30       Dim lasRights() As String, lnRights As Integer
40       Dim lnSeek As Integer, lbSeek As Boolean
   
50       AllowTransSave = False
   
60       If cTranStat = xeStatePosted Then
70          If oAppDrivr.UserLevel = xeSysMaster Then
80             AllowTransSave = True
90          Else
100            MsgBox "Modification of posted or cancelled transaction is not allowed!!!" & vbCrLf & vbCrLf & _
               "Request not granted!!!", vbCritical, "Warning"
110         End If
120         GoTo endProc
130      End If

140      With oAppDrivr
150         lasRights = Split(sRemarks, "»")
      
160         lbSeek = False
170         If nEditMode = xeModeAddNew Then
         ' Add Rights
180            lnRights = CInt(lasRights(1))
190            If lnRights And .UserLevel = 0 Then
200               lnSeek = MsgBox("User Doesn't Have the Right to Create New Transaction!!!" & _
                              vbCrLf & "Seek Authorized User's Approval?", _
                              vbQuestion + vbYesNo + vbDefaultButton2, "Confirm")
210               If lnSeek = vbNo Then GoTo endProc
220               lbSeek = True
230            End If
240         ElseIf nEditMode = xeModeUpdate Then
         ' update Rights
250            lnRights = CInt(lasRights(2))
         
260            If lnRights And .UserLevel = 0 Then
270               Select Case .UserLevel
            Case xeEncoder
280                  lbSeek = True
290               Case xeSupervisor
300                  If DateDiff("d", dModified, .ServerDate) > 2 Then lbSeek = True
310               Case xeManager
320                  If DateDiff("d", dModified, .ServerDate) > 30 Then lbSeek = True
330               Case xeAudit
340                  lbSeek = True
350               Case xeSysAdmin
360                  lbSeek = True
370               End Select
380            End If
390         End If
      
400         If lbSeek = False Then
405            sApprovID = .UserID
410            AllowTransSave = True
420            GoTo endProc
430         End If
      
440         lnSeek = MsgBox("User Doesn't Have the Right to Update Transaction!!!" & _
                        vbCrLf & "Seek Authorized User's Approval?", _
                        vbQuestion + vbYesNo + vbDefaultButton2, "Confirm")
450         If lnSeek = vbNo Then GoTo endProc
      
460         If GetApproval(oAppDrivr, lnUserRght, lsUserIDxx, lsUserName) = False Then GoTo endProc
      
470         If nEditMode = xeModeAddNew Then
480            If lnUserRght And .UserLevel = 0 Then
490               MsgBox "Approving User Doesn't Have the Right to Create New Transaction!!!" & _
                  "Seek Other User's Assistance then Try Again", vbCritical, "Warning"
500               GoTo endProc
510            End If
520         ElseIf nEditMode = xeModeUpdate Then
530            If lnUserRght And .UserLevel = 0 Then
540               Select Case lnUserRght
            Case xeSupervisor
550                  If DateDiff("d", dModified, oAppDrivr.ServerDate) <= 2 Then lbSeek = False
560               Case xeManager
570                  If DateDiff("d", dModified, oAppDrivr.ServerDate) <= 31 Then lbSeek = False
580               Case xeSysAdmin
590                  If .UserLevel <> xeSysAdmin Then lbSeek = False
600               End Select
            
610               If lbSeek Then
620                  MsgBox "Approving User Doesn't Have the Right to Update Transaction!!!" & _
                     "Seek Other User's Assistance then Try Again", vbCritical, "Warning"
630                  GoTo endProc
640               End If
650            End If
660         End If
670      End With

675      sApprovID = lsUserIDxx
680      AllowTransSave = True

endProc:
690      Exit Function
End Function


