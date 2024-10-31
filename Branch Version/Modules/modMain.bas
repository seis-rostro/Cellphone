Attribute VB_Name = "modMain"

Option Explicit

Private Const pxeMODULENAME As String = "modMain"
Private Const pxeCPMainID As String = "01"

Public oApp As clsAppDriver
Public oReport As CRAXDRT.Report
Public oRepApp As New CRAXDRT.Application
Public Declare Function GetFocus Lib "USER32" () As Long
Public Declare Function SetSysColors Lib "USER32" (ByVal nChanges As Long, lpSysColor As Long, lpColorValues As Long) As Long
Public Declare Function GetSysColor Lib "USER32" (ByVal nIndex As Long) As Long

Private Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Enum adReport
   ViewReport = 0
   PrintReport = 1
End Enum

Dim psCriteria(3) As String

Private Sub Main()
'   Dim lsCommand As String
'   Dim lasParam() As String
'   Dim loSysMonitor As clsSysMonitor
'
'   On Error GoTo errProc
'
'   lsCommand = Command()
'   lasParam = Split(lsCommand)
'
'   Set oApp = New clsAppDriver
'
'   If oApp.LoadEnv(lasParam(0), lasParam(1)) = False Then Exit Sub
'
'   Set oApp.mdiMain = mdiMain
'   mdiMain.Caption = oApp.ProductName
'   mdiMain.Show
'
'   If LCase(Mid(oApp.Config("sDBHostNm"), 1, 9)) <> LCase(Mid(oApp.ComputerName, 1, 9)) Or oApp.UserLevel = xeAudit Then Exit Sub
'
'   If oApp.isMainOffice = False Or oApp.IsWarehouse = False Then
'      Set loSysMonitor = New clsSysMonitor
'      Set loSysMonitor.AppDriver = oApp
'      loSysMonitor.ProductID = oApp.ProductID
'      loSysMonitor.InitMonitor
'      If loSysMonitor.StartMonitor = False Then
'         Unload mdiMain
'         Exit Sub
'      End If
'   End If
'
'endProc:
'   Exit Sub
'errProc:
'   MsgBox "Line No:" & Erl & vbCrLf & Err.Description, vbCritical, "Error"
'   End

   Set oApp = New clsAppDriver
   If oApp.LoadEnv("Telecom") = False Then
      Exit Sub
   End If

   If oApp.LogIn("Telecom") = False Then
      Exit Sub
   End If

   Set oApp.mdiMain = mdiMain
   mdiMain.Caption = oApp.ProductName
   mdiMain.Show

   If oApp.UserLevel = xeAudit Then
      Exit Sub
   Else
      If oApp.BranchCode <> "M001" Then
         If LCase(Mid(oApp.Config("sDBHostNm"), 1, 9)) <> LCase(Mid(oApp.ComputerName, 1, 9)) Then Exit Sub
      End If
   End If

'   Set loSysMonitor = New clsSysMonitor
'   Set loSysMonitor.AppDriver = oApp
'   loSysMonitor.ProductID = oApp.ProductID
'   loSysMonitor.InitMonitor
'   If loSysMonitor.StartMonitor = False Then
'      Unload mdiMain
'      Exit Sub
'   End If
End Sub

Public Sub SetNextFocus()
   keybd_event &H9, 0, 0, 0
   keybd_event &H9, 0, &H2, 0
End Sub

Public Sub SetPreviousFocus()
   keybd_event &H10, 0, 0, 0
   keybd_event &H9, 0, 0, 0
   keybd_event &H10, 0, &H2, 0
End Sub

Public Sub SetEnterKey()
   keybd_event &H13, 0, 0, 0
   keybd_event &H13, 0, &H2, 0
End Sub

Public Sub CenterChildForm(frmMDIForm As MDIForm, frmChild As Form)
   Dim lbX As Long, lbY As Long

   lbX = frmMDIForm.ScaleWidth
   lbY = frmMDIForm.ScaleHeight

   frmChild.Left = CLng((lbX - frmChild.Width) / 2)
   frmChild.Top = CLng((lbY - frmChild.Height) / 2)

   If frmChild.Left < 0 Then frmChild.Left = 0
   If frmChild.Top < 0 Then frmChild.Top = 0
End Sub

Public Function CodeFormat(sBranch As String, sCode As String) As String
   Dim lnLoc As Integer

   lnLoc = InStr(sCode, "-")
   If Len(sCode) < 6 Then
      CodeFormat = sBranch & Format(oApp.ServerDate, "YY") & String(6 - Len(sCode), "0") & sCode
   ElseIf lnLoc > 0 Then
      CodeFormat = Left(sCode, lnLoc - 1) & Right(sCode, Len(sCode) - lnLoc)
   Else
      CodeFormat = sCode
   End If
End Function

Public Sub setGrayText(ByVal lnColor As Long)
   SetSysColors 1, 17, lnColor
End Sub

Public Function TransStat(nStat As Integer) As String
   Select Case nStat
   Case 0
      TransStat = "OPEN"
   Case 1
      TransStat = "CLOSED"
   Case 2
      TransStat = "POSTED"
   Case 3
      TransStat = "CANCELLED"
   Case 4
      TransStat = "UNKNOWN"
   End Select
End Function

Public Function JobOrderStatus(ByVal nStatus As xeJobOrderStatus) As String
   Select Case nStatus
   Case 0
      JobOrderStatus = "OPEN"
   Case 1
      JobOrderStatus = "JOB ORDER"
   Case 2
      JobOrderStatus = "FOR REPAIR"
   Case 3
      JobOrderStatus = "RELEASED"
   Case 4
      JobOrderStatus = "CANCELLED"
   Case 5
      JobOrderStatus = "FORWARDED"
   Case 6
      JobOrderStatus = "REPAIRED"
   End Select
End Function

Public Function BranchStatus(ByVal sBranchCd As String, Optional sSQL As Variant) As Boolean
   Dim lrs As Recordset
   Dim lsSQL As String

   lsSQL = "SELECT * FROM Branch" & _
               " WHERE sBranchCd = " & strParm(sBranchCd)

   If Not IsMissing(sSQL) Then
      If Trim(sSQL) <> "" Then
         lsSQL = AddCondition(lsSQL, sSQL)
      End If
   End If

   Set lrs = New Recordset
   lrs.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText

   If Not lrs.EOF Then BranchStatus = True
   Set lrs = Nothing
End Function

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

Public Function getCTime(ByVal sTime) As String
   Const sALLOWEDCHAR As String = "0123456789"

   Dim sAllowed As String
   Dim sChar As String
   Dim retVal As String
   Dim sTTemp As String
   Dim sHH As String
   Dim sMM As String
   Dim sExt As String

   Dim lnLen As Integer
   Dim lnExt As Integer
   Dim lnCtr As Integer

   If sTime = "" Then GoTo endProc
   'add brakets to string for using LIKE
   sAllowed = "[" & sALLOWEDCHAR & "]"
   'get time
   sTime = LCase(Replace(sTime, " ", ""))
   'get length
   lnLen = Len(sTime)
   'check the length; maximum 7
   If lnLen < 4 Or lnLen > 7 Then GoTo endProc

   'get extension position
   If InStr(sTime, "a") > 4 Then
      lnExt = InStr(sTime, "a") - 1
   ElseIf InStr(sTime, "p") > 4 Then
      lnExt = InStr(sTime, "p") - 1
   Else
      lnExt = lnLen
   End If

   'set time to temp
   sTTemp = Left(sTime, lnExt)
   'get extension
   sExt = Right(sTime, lnLen - lnExt)

  ' get the minutes
   sMM = Right(sTTemp, 2)
   If Not IsNumeric(sMM) Then Exit Function
   sMM = Mid(sTTemp, 4, 2)

   'set the hour to temp
   sTTemp = Left(sTTemp, Len(sTTemp) - 2)
   sTTemp = Left(sTTemp, 2)
   'Now loop through all characters in the string removing all unwanted charaters
   For lnCtr = 1 To Len(sTTemp)
       sChar = Mid$(sTTemp, lnCtr, 1)
       If sChar Like sAllowed Then
           retVal = retVal & sChar
       End If
   Next
   'set hour
   sHH = retVal
   If Not IsNumeric(sHH) Then Exit Function

   getCTime = Format(sHH & ":" & sMM & sExt, "HH:MM AM/PM")

endProc:
   Exit Function
End Function

'Mac PH(08.07.12)
Public Sub HighlightOn(loTextbox As TextBox)
   With loTextbox
      .BackColor = oApp.getColor("HT1")
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
End Sub

Public Sub HighlightOff(loTextbox As TextBox)
   With loTextbox
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Public Sub EmphasizeField(loTextbox As TextBox)
   With loTextbox
      .BackColor = &HFF00&
      .ForeColor = &HFF&
      .FontBold = True
   End With
End Sub

Function strLongDate(ByVal Value As String) As String
   strLongDate = Format(Value, "Mmm dd, yyyy")
End Function

Function strCurrency(ByVal Value As String) As String
   If Value = "" Then Exit Function
   strCurrency = Format(CDbl(Value), "##,##0.00")
End Function

Sub SetGridRowColor(ByVal loGrid As MSFlexGrid, _
                     ByVal lnMode As Integer, _
                     ByVal lnCol As Integer, _
                     Optional ByVal lnRow As Integer = 0)
   Dim lnCtr As Integer

   Select Case lnMode
      Case 0 ' full
         With loGrid
            For lnCtr = 1 To .Rows - 1
               If lnCtr Mod 2 = 0 Then
                  .FillStyle = flexFillRepeat
                  .Row = lnCtr
                  .RowSel = lnCtr
                  .Col = lnCol
                  .ColSel = .Cols - 1
                  .CellBackColor = &HFFC0C0
                  .CellBackColor = &HFFC0FF
               End If
            Next
            .Row = .Rows - 1
         End With
      Case 1 ' single
         With loGrid
            If IsMissing(lnRow) Then Exit Sub
            If lnRow = 0 Or lnRow Mod 2 = 1 Then Exit Sub

            .FillStyle = flexFillRepeat
            .Row = lnRow
            .RowSel = lnRow
            .Col = lnCol
            .ColSel = .Cols - 1
            .CellBackColor = &HFFC0C0
            .CellBackColor = &HFFC0FF
            .Row = .Rows - 1
         End With
   End Select
End Sub

Function strShortDate(ByVal Value As String) As String
   strShortDate = Format(Value, "MM-DD-YYYY")
End Function

Function WhoIs(ByVal fsID As String, Optional ByVal fbCypher As Boolean = False) As String
   Dim lsSQL As String
   Dim loRS As Recordset

   If fbCypher Then
      fsID = Decrypt(fsID)
   End If

   lsSQL = "SELECT sUserName" & _
          " FROM xxxSysUser" & _
          " WHERE sUserIDxx = " & strParm(fsID)
   Set loRS = oApp.Connection.Execute(lsSQL, , adCmdText)

   If loRS.EOF Then
      WhoIs = "N-O-N-E"
   Else
      WhoIs = Decrypt(loRS("sUserName"), oApp.Machinex)
   End If

   Set loRS = Nothing
End Function

Public Function SaveOthers(ByVal loRS As Recordset _
                           , ByVal lsTable As String _
                           , ByVal lbEditMode As xeEditMode _
                           , Optional lsFldReference As String) As Boolean
   Dim lsOldProc As String
   Dim lsSQL As String
   Dim lnRow As Integer
   Dim lnCtr As Integer
   Dim lnField() As String
   Dim lnRef As Integer
   Dim lnCol As Integer

   lsOldProc = "SaveOthers"
   On Error GoTo errProc

   If Not (lbEditMode <> xeModeAddNew Or _
      lbEditMode <> xeModeUpdate) Then Exit Function

   For lnCtr = 0 To loRS.Fields.Count - 1
      If IsNull(loRS(lnCtr).OriginalValue) Or _
         loRS(lnCtr).OriginalValue <> _
         loRS(lnCtr) Then
         If IsNull(loRS(lnCtr)) = False Then
            lsSQL = lsSQL & ", " & _
                     loRS(lnCtr).Name & " = " & _
                     FieldParam(loRS(lnCtr).Type, _
                     loRS(lnCtr))
         End If
      End If
   Next

   If lsSQL = "" Then
      SaveOthers = True
      GoTo endProc
   End If

   If lbEditMode = xeModeAddNew Then
      lsSQL = "INSERT INTO " & lsTable & " SET" & _
                  Mid(lsSQL, 2) & _
                  ", sModified = " & strParm(Encrypt(oApp.UserID)) & _
                  ", dModified = " & dateParm(oApp.ServerDate)
   Else
      If IsMissing(lsFldReference) Then Exit Function
      lnField = Split(lsFldReference, "»")
      lsSQL = "UPDATE " & lsTable & " SET" & _
                  Mid(lsSQL, 2) & _
                  ", sModified = " & strParm(Encrypt(oApp.UserID)) & _
                  ", dModified = " & dateParm(oApp.ServerDate) & _
               " WHERE "

      For lnRef = 0 To loRS.Fields.Count - 1
         For lnCol = 0 To UBound(lnField)
            If loRS(lnField(lnCol)).Name = loRS(lnRef).Name Then
               lsSQL = lsSQL & _
                     loRS(lnRef).Name & " = " & _
                     strParm(loRS(lnRef)) & _
                     " AND "
            End If
         Next
      Next

      If Right(Trim(lsSQL), 3) <> "AND" Then GoTo endProc
      lsSQL = Left(lsSQL, Len(Trim(lsSQL)) - 3)
   End If

   Debug.Print lsSQL

   lnRow = oApp.Execute(lsSQL, lsTable)
   If lnRow <= 0 Then
      MsgBox "Unable to update " & lsTable & vbCrLf & lsSQL _
               , vbCritical, "Warning"
      GoTo endProc
   End If

   SaveOthers = True

endProc:
   Set loRS = Nothing
   Exit Function
errProc:
   ShowError lsOldProc & "( " & SaveOthers & " )", True
End Function

Public Function padLeft(ByVal StrToPad As String, _
      ByVal Length As Integer, _
      Optional Pad As Variant = " ") As String
   Dim lsPaded As String
   Dim lnLenStr As Integer
   
   lsPaded = Trim(StrToPad)
   lnLenStr = Len(StrToPad)
   
   If Length <= lnLenStr Then
      padLeft = StrToPad
      Exit Function
   End If
   
   padLeft = String(Length - lnLenStr, Pad) & lsPaded
End Function

Public Function padRight(ByVal StrToPad As String, _
      ByVal Length As Integer, _
      Optional Pad As Variant = " ") As String
   Dim lsPaded As String
   Dim lnLenStr As Integer
   
   lsPaded = Trim(StrToPad)
   lnLenStr = Len(StrToPad)
   
   If Length <= lnLenStr Then
      padRight = StrToPad
      Exit Function
   End If
   
   padRight = lsPaded & String(Length - lnLenStr, Pad)
End Function

Public Function MinusLeftChar(ByVal sGiven As String) As String
   On Error Resume Next
   
   If Len(sGiven) = 0 Then
      MinusLeftChar = ""
   Else
      MinusLeftChar = Mid$(sGiven, 2)
   End If
End Function

Public Function MinusRightChar(ByVal sGiven As String) As String
   On Error Resume Next
   
   If Len(sGiven) = 0 Then
      MinusRightChar = ""
   Else
      MinusRightChar = Left$(sGiven, Len(sGiven) - 1)
   End If
End Function

Public Function AccountStat(lnStat As Integer) As String
   Select Case lnStat
   Case 0
      AccountStat = "Active"
   Case 1
      AccountStat = "Closed"
   Case 2
      AccountStat = "Dead"
   Case 3
      AccountStat = "Impounded"
   Case 4
      AccountStat = "Discarded"
   End Select
End Function

Public Function ApplStat(nStat As Integer) As String
   Select Case nStat
   Case 0
      ApplStat = "Open"
   Case 1
      ApplStat = "Closed"
   Case 2
      ApplStat = "Approved"
   Case 3
      ApplStat = "Disapproved"
   Case 4
      ApplStat = "Selected"
   End Select
End Function

Public Function RatingStat(lsStat As String) As String
   lsStat = Trim(LCase(lsStat))

   Select Case lsStat
   Case "x"
      RatingStat = "Excellent"
   Case "g"
      RatingStat = "Good"
   Case "f"
      RatingStat = "Fair"
   Case "p"
      RatingStat = "Poor"
   Case "b"
      RatingStat = "Bad"
   Case "r"
      RatingStat = "Rejected"
   Case "n"
      RatingStat = "No Basis"
   Case "l"
      RatingStat = "Blacklist"
   End Select
End Function

Public Function isDatePosted(ByVal fdTranDate As Date) As Boolean
   Dim loRS As Recordset
   Dim lsSQL As String
   
   Set loRS = New Recordset
   loRS.Open "SELECT dUnEncode FROM Branch_Others WHERE sBranchCd = " & strParm(oApp.BranchCode) _
   , oApp.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
   
   If loRS.EOF Then
      isDatePosted = False
      Exit Function
   Else
      If IsNull(loRS("dUnEncode")) Then
         isDatePosted = False
         Exit Function
      Else
         If Format(loRS("dUnEncode"), "YYYYMMDD") > Format(fdTranDate, "YYYYMMDD") Then
            isDatePosted = False
            Exit Function
         End If
      End If
   End If

   lsSQL = "SELECT" & _
               " sTranDate" & _
            " FROM DTR_Summary" & _
            " WHERE sBranchCd = " & strParm(oApp.BranchCode) & _
               " AND cPostedxx = " & strParm(xeStatePosted) & _
            " ORDER BY sTranDate DESC" & _
            " LIMIT 1"

   Set loRS = New Recordset
   loRS.Open lsSQL, oApp.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText

   If loRS.EOF Then
      isDatePosted = False
      Exit Function
   End If

   If CDate(loRS("sTranDate")) <= fdTranDate Then
      MsgBox "Trasaction Date is not valid!!!" & vbCrLf & _
               "Please verify your entry then try again!!!", vbCritical, "WARNING"
      isDatePosted = False
      Exit Function
   End If

   isDatePosted = True
End Function

Public Function isTransValid(ByVal fdTranDate As Date, _
                                 ByVal fsTranType As String, _
                                 ByVal fsReferNox As String, ByVal fsAmountxx As Double) As Boolean
   Dim loRS As Recordset
   Dim lsSQL As String
   
   isTransValid = True
   
   Set loRS = New Recordset
   loRS.Open "SELECT dUnEncode FROM Branch_Others WHERE sBranchCd = " & strParm(oApp.BranchCode) _
   , oApp.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
   
   If loRS.EOF Then Exit Function
   
   If IsNull(loRS("dUnEncode")) Then
      Exit Function
   Else
      'she 2019-12-12
      'recode the alidation of unencoded transaction
      If DateDiff("d", loRS("dUnEncode"), fdTranDate) >= 0 Then
         'check the DTR_Summary here here
         lsSQL = "SELECT cPostedxx FROM DTR_Summary WHERE sBranchCd = " & strParm(oApp.BranchCode) & _
                  " AND sTranDate = " & strParm(Format(fdTranDate, "YYYYMMDD"))
         Debug.Print lsSQL
         Set loRS = New Recordset
         loRS.Open lsSQL, oApp.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
      
         If loRS.EOF Then
            isTransValid = True
         Else
            'if cPosted = 2, do not allow any transaction to encode
            If loRS("cPostedxx") = xeStatePosted Then
               MsgBox "DTR Date was already posted!!!" & vbCrLf & _
                     "Please verify your entry then try again!!!", vbCritical, "WARNING"
               isTransValid = False
            'cposted = 1 then check referno to DTR_Summary_Detail
            ElseIf loRS("cPostedxx") = xeStateClosed Then
               lsSQL = "SELECT b.cHasEntry, a.cPostedxx, b.nTranAmtx" & _
                  " FROM DTR_Summary a" & _
                  ", DTR_Summary_Detail b" & _
                  " WHERE a.sBranchCd = b.sBranchCd" & _
                  " AND a.sTranDate = b.sTranDate" & _
                  " AND a.sBranchCd = " & strParm(oApp.BranchCode) & _
                  " AND a.sTranDate = " & strParm(Format(fdTranDate, "YYYYMMDD")) & _
                  " AND b.sTranType = " & strParm(fsTranType) & _
                  " AND b.sReferNox = " & strParm(fsReferNox) & _
                  " AND b.nTranAmtx = " & fsAmountxx & _
                  " AND b.cHasEntry = " & strParm(xeNo)
               Debug.Print lsSQL
               Set loRS = New Recordset
               loRS.Open lsSQL, oApp.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
               
               If loRS.EOF Then
                  MsgBox "No Reference no found from unencoded transaction!!" & vbCrLf & _
                         "OR Transaction Amount is not equal to the unposted amount!!" & vbCrLf & _
                         " Pls check your entry then try again!!!"
                  isTransValid = False
               ElseIf loRS("cHasEntry") = xeStateClosed Then
                   MsgBox "Reference No was already posted!!!" & vbCrLf & _
                           " Pls check your entry then try again!!!"
                  isTransValid = False
               Else
                  isTransValid = True
               End If
            ElseIf loRS("cPostedxx") = xeStateOpen Then
               isTransValid = True
            Else
               isTransValid = False
            End If
         End If
      Else
         isTransValid = False
         MsgBox "Unable to encode previous Transaction!!!" & vbCrLf & _
                  " Pls inform MIS/COMPLIANCE DEPT!!!", vbInformation, "WARNING"
      End If
   End If

'   If Format(fdTranDate, "YYYYMMDD") = Format(oApp.ServerDate, "YYYYMMDD") Then Exit Function
'
'   lsSQL = "SELECT" & _
'               "  a.cPostedxx" & _
'            " FROM DTR_Summary a" & _
'               ", DTR_Summary_Detail b" & _
'            " WHERE a.sBranchCd = b.sBranchCd" & _
'               " AND a.sTranDate = b.sTranDate" & _
'               " AND a.sBranchCd = " & strParm(oApp.BranchCode) & _
'               " AND a.sTranDate = " & strParm(Format(fdTranDate, "YYYYMMDD")) & _
'               " AND b.sTranType = " & strParm(fsTranType) & _
'               " AND b.sReferNox = " & strParm(fsReferNox) & _
'               " AND b.cHasEntry = " & strParm(xeNo)
'
'   Set lors = New Recordset
'   lors.Open lsSQL, oApp.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'   If Not lors.EOF Then
'      If lors.RecordCount > 1 Then
'         MsgBox "Invalid Transaction detected!!!" & vbCrLf & _
'                  "Multiple record found!!!", vbCritical, "WARNING"
'         isTransValid = False
'      Else
'         If lors("cPostedxx") = xeStatePosted Then
'            MsgBox "Transaction date already posted!!!" & vbCrLf & _
'                     "Please verify your entry then try again!!!", vbCritical, "WARNING"
'            isTransValid = False
'         End If
'      End If
'   Else
'      lsSQL = "SELECT" & _
'               " cPostedxx" & _
'            " FROM DTR_Summary" & _
'            " WHERE sBranchCd = " & strParm(oApp.BranchCode) & _
'               " AND sTranDate = " & strParm(Format(fdTranDate, "YYYYMMDD"))
'      Debug.Print lsSQL
'      Set lors = New Recordset
'      lors.Open lsSQL, oApp.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
'      If Not lors.EOF Then
'         If lors("cPostedxx") = xeStatePosted Then
'            MsgBox "Transaction date already posted!!!" & vbCrLf & _
'                     "Please verify your entry then try again!!!", vbCritical, "WARNING"
'            isTransValid = False
'         Else
'            MsgBox "Transaction is not yet encoded!!!" & vbCrLf & _
'                        "Please verify your entry then try again!!!", vbCritical, "WARNING"
'            isTransValid = False
'         End If
'      End If
'   End If
End Function

Private Function chkUnencodedTrans(ByVal fsTranDate As String _
                                    , ByVal fsTranType As String _
                                    , ByVal fsReferNox As String) As Boolean
   Dim loRS As Recordset
   Dim loSrc As Recordset
   Dim lsProcName As String
   Dim lnRow As Long
   Dim lsSQL As String
   
   lsProcName = "chkUnencodedTrans"
   On Error GoTo errProc

   Set loRS = New Recordset
   With oApp
      lsSQL = "SELECT *" & _
                  " FROM DTR_Summary_Detail" & _
               " WHERE sBranchCd = " & strParm(oApp.BranchCode) & _
                  " AND DATE_FORMAT(dTransact,'%Y%m%d') = " & strParm(Format(fsTranDate, "YYYYMMDD")) & _
                  " AND sTranType = " & strParm(fsTranType) & _
                  " AND sReferNox = " & strParm(fsReferNox) & _
                  " AND cHasEntry = '0'"
                  
      loRS.Open lsSQL, .Connection, , , adCmdText
      If loRS.EOF = False Then
         Do
            Select Case loRS("sTranType")
            Case "CPSl"
               lsSQL = "SELECT *" & _
                        " FROM CP_SO_Master" & _
                        " WHERE sTransNox LIKE " & strParm(oApp.BranchCode & "%") & _
                           " AND DATE_FORMAT(dTransact,'%Y%m%d') = " & strParm(loRS("sTranDate")) & _
                           " AND sSalesInv = " & strParm(loRS("sReferNox"))
            Case "MCSc"
               lsSQL = "SELECT *" & _
                        " FROM Receipt_Master" & _
                        " WHERE sTransNox LIKE " & strParm(oApp.BranchCode & "%") & _
                           " AND cTranType = '9'" & _
                           " AND DATE_FORMAT(dTransact,'%Y%m%d') = " & strParm(loRS("sTranDate")) & _
                           " AND sORNoxxxx = " & strParm(loRS("sReferNox"))
            Case "CPLd"
               lsSQL = "SELECT *" & _
                        " FROM CP_SO_Eload" & _
                        " WHERE sTransNox LIKE " & strParm(oApp.BranchCode & "%") & _
                           " AND cTranType = '9'" & _
                           " AND DATE_FORMAT(dTransact,'%Y%m%d') = " & strParm(loRS("sTranDate")) & _
                           " AND sORNoxxxx = " & strParm(loRS("sReferNox"))
            End Select
      
            Set loSrc = New Recordset
            loSrc.Open lsSQL, oApp.Connection, , , adCmdText
      
            If Not loSrc.EOF Then
               lsSQL = "UPDATE DTR_Summary_Detail SET" & _
                           " cHasEntry = '1'" & _
                        " WHERE sBranchCd = " & strParm(oApp.BranchCode) & _
                           " AND sTranDate = " & strParm(loRS("sTranDate")) & _
                           " AND sReferNox = " & strParm(loRS("sReferNox")) & _
                           " AND sTranType = " & strParm(loRS("sTranType"))
               If oApp.Execute(lsSQL, "DTR_Summary_Detail") <= 0 Then
                  GoTo endProc
               End If
            End If
            
            loRS.MoveNext
         Loop Until loRS.EOF
      
      End If
      loRS.Close
   End With
   
   chkUnencodedTrans = True
   
endProc:
   Set loRS = Nothing
   Exit Function
errProc:
   ShowError lsProcName & "( " & " ) "
End Function


