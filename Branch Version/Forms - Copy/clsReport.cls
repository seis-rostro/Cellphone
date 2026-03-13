VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
 '€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Rex S. Adversalo
' XerSys Computing
' Canaoalan, Binmaley, Pangasinan
'
' Copyright 2003-2004 and beyond
' All Rights Reserved
'
'     General Report Viewer Object
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
'  XerSys [ 12/03/2004 11:48 am ]
'     Start creating this object.
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Explicit

Private p_oRepApp As Application
Private p_oReport As Report
Private p_oRepSource As Recordset
Private p_oProgress As Speedometer

Private WithEvents p_oRepViewer As frmViewer
Attribute p_oRepViewer.VB_VarHelpID = -1

Private p_sReportID As String
Private p_sRepHead1 As String
Private p_sRepHead2 As String
Private p_sPgeRange As String
Private p_nNoCopies As Integer
Private p_nOrientxx As Integer
Private p_bCollated As Boolean
Private p_bPreviewx As Boolean

Dim pbSave As Boolean

Property Let ReportHeading1(ByVal Value As String)
   p_sRepHead1 = Value
End Property

Property Let ReportHeading2(ByVal Value As String)
   p_sRepHead2 = Value
End Property

Property Let ReportID(ByVal Value As String)
   p_sReportID = Value
End Property

Property Set ReportObject(Value As Report)
   Set p_oReport = Value
End Property

Property Set ReportSource(Value As Recordset)
   Set p_oRepSource = Value
End Property

Property Get Collate() As Boolean
   Collate = p_bCollated
End Property

Property Let Collate(ByVal Value As Boolean)
   p_bCollated = Value
End Property

Property Get Copies() As Integer
   Copies = p_nNoCopies
End Property

Property Let Copies(ByVal Value As Integer)
   p_nNoCopies = Value
End Property

Property Get PageRange() As String
   PageRange = p_sPgeRange
End Property

Property Let PageRange(ByVal Value As String)
   p_sPgeRange = Value
End Property

Property Get Orientation() As Integer
   Orientation = p_nOrientxx
End Property

Property Let Orientation(ByVal Value As Integer)
   Select Case Value
   Case 0, 1
      p_nOrientxx = Value
   End Select
End Property

Sub InitReport()
   Set p_oReport = Nothing
   Set p_oRepSource = Nothing
   pbSave = False
End Sub

'Sub ShowReport()
'   Dim lsOldProc As String
'
  ' lsOldProc = oApp.ProcName("SaveReport")
 '  On Error GoTo errProc

   'Set p_oProgress = New Speedometer

   'With p_oProgress
    '  .InitProgress "Processing"
     ' .PrimaryRemarks = "Loading Report..."
      'DoEvents
   'End With

   'Set p_oRepViewer = New frmRepPreview
   'If p_oRepSource Is Nothing And p_oReport Is Nothing Then
   '   p_oRepViewer.AllowBrowse = True
   '   p_oRepViewer.Show
   '   Exit Sub
   'End If

   ' before showing report, be sure to save it first
   '  its the responsibility of the system administrator to
   '  purge unwanted reports
   'If p_sReportID = Empty Then Exit Sub

   'SaveReport

   'With p_oRepViewer
    '  .CRViewer91.ReportSource = p_oReport
     ' .Show
      '.CRViewer91.ViewReport

      'While .CRViewer91.IsBusy
       '  DoEvents
      'Wend
   'End With
   'p_oProgress.CloseProgress

'endProc:
 '  oApp.ProcName lsOldProc
  ' Exit Sub
'errProc:
 '  Set p_oProgress = Nothing
   'ShowError lsOldProc
  ' MsgBox Erl & " " & Err.Description, vbCritical, "Warning"
'End Sub

Sub PrintReport()
   Set p_oProgress = New Speedometer

   With p_oProgress
      .InitProgress "Processing..."
      .PrimaryRemarks = "Loading Report..."
   End With

   If p_oRepSource Is Nothing And p_oReport Is Nothing Then GoTo endProc

   ' before showing report, be sure to save it first
   '  its the responsibility of the system administrator to
   '  purge unwanted reports
   If p_sReportID = Empty Then Exit Sub

   SaveReport

   p_oReport.PrintOutEx False, 1

endProc:
   p_oProgress.CloseProgress
   Set p_oProgress = Nothing
   Exit Sub
End Sub

Private Function SaveReport() As Boolean
   Dim loRS As Recordset
   Dim lsOldProc As String
   Dim lnRepName As String
   Dim lsSQL As String
   Dim lnRow As Long

   lsOldProc = oApp.ProcName("SaveReport")
   On Error GoTo errProc
   SaveReport = False

   lsSQL = "SELECT" & _
               "  sReportNm" & _
               ", sFileName" & _
               ", sReportHd" & _
               ", cSaveRepx" & _
               ", cLogRepxx" & _
            " FROM xxxReport" & _
            " WHERE sReportID = " & strParm(p_sReportID)

   Set loRS = New Recordset
   loRS.Open lsSQL, oApp.Connection, , , adCmdText

   If loRS.EOF Then GoTo endProc

   If p_oReport Is Nothing Then
      If p_sRepHead1 = Empty Then p_sRepHead1 = loRS("sReportHd")
      If p_sRepHead2 = Empty Then p_sRepHead2 = "As of " & Format(Date, "MMMM DD, YYYY")

      Set p_oRepApp = New Application

      Set p_oReport = p_oRepApp.OpenReport(oApp.AppPath & "\Reports\" & loRS("sFileName") & ".rpt")
      Set p_oRepApp = Nothing

      p_oReport.DiscardSavedData
      p_oReport.FieldMappingType = crAutoFieldMapping
      p_oReport.Database.SetDataSource p_oRepSource
      With oApp

         p_oReport.Sections("RHa").ReportObjects("txtCompany").SetText "GMC Anolid Warehouse"
         p_oReport.Sections("RHa").ReportObjects("txtAddress").SetText "Anolid, Mangaldan, Pangasinan"

'         p_oReport.Sections("RHa").ReportObjects("txtCompany").SetText .ClientName
'         p_oReport.Sections("RHa").ReportObjects("txtAddress").SetText .Address & _
'                     ", " & .TownCity & ", " & .Province & " " & .ZippCode

         p_oReport.Sections("PH").ReportObjects("txtHeading1").SetText p_sRepHead1
         p_oReport.Sections("PH").ReportObjects("txtHeading2").SetText p_sRepHead2

         p_oReport.Sections("PF").ReportObjects("txtRptUser").SetText .UserName
      End With
   End If

   If pbSave = True Then
      SaveReport = True
      GoTo endProc
   End If

   If loRS("cSaveRepx") = 1 Then
      lnRepName = GetNextReport
      p_oReport.SaveAs oApp.AppPath & "\Reports\" & lnRepName, crDefaultFileFormat
      ' App.Path & "\Reports\" & lnRepName
   End If

   If loRS("cLogRepxx") = 1 Then
      With oApp
         lsSQL = "INSERT INTO xxxReportsLog (" & _
                     "  sReportID" & _
                     ", dGenerate" & _
                     ", sUserIDxx" & _
                     ", sRepFName )" & _
                  " VALUES (" & _
                     strParm(p_sReportID) & _
                     ", " & dateParm(.ServerDate) & _
                     ", " & strParm(.UserID) & _
                     ", " & strParm(lnRepName) & " )"

         .Connection.Execute lsSQL, lnRow, adCmdText
         If lnRow = 0 Then
            MsgBox "Unable to Register Report Generation!!!", vbCritical, "Warning"
            GoTo endProc
         End If
      End With
   End If

   pbSave = True
   SaveReport = True

endProc:
   Set loRS = Nothing
   oApp.ProcName lsOldProc
   Exit Function
errProc:
   ShowError lsOldProc
End Function

Private Function GetNextReport() As String
   Dim loRS As ADODB.Recordset
   Dim lsSQL As String
   Dim lnCode As Long

   lsSQL = "SELECT TOP 1 sRepFName" & _
            " FROM xxxReportsLog" & _
            " WHERE sReportID = " & strParm(p_sReportID) & _
               " AND sRepFName LIKE " & _
               strParm(oApp.BranchCode & Format(Date, "YY") & "%") & _
            " ORDER BY sRepFName DESC"

   Set loRS = New ADODB.Recordset
   loRS.Open lsSQL, oApp.Connection, , , adCmdText

   lsSQL = IIf(loRS.EOF, Empty, loRS(0))
   lnCode = 1
   If lsSQL <> Empty Then lnCode = CLng(Mid(lsSQL, 5)) + 1

   Set loRS = Nothing
   GetNextReport = oApp.BranchCode & Format(Date, "yy") & Format(lnCode, String(6, "0"))
End Function

Private Sub ShowError(ByVal lsProcName As String)
   With oApp
      .ShowError "ReportViewer", .ProcName(lsProcName), Err.Number, Err.Description, Erl
   End With
   With Err
      .Raise .Number, .Source, .Description
   End With
End Sub

'Private Sub Class_Initialize()
 '  Set p_oRepViewer = New frmRepPreview

  ' p_nNoCopies = 1
  ' p_bCollated = True
  ' p_sPgeRange = "xxx"
  ' p_nOrientxx = 1

   'pbSave = False
'End Sub

Private Sub Class_Terminate()
   Set p_oRepApp = Nothing
   Set p_oReport = Nothing
   Set p_oRepSource = Nothing
   Set p_oRepViewer = Nothing
End Sub

Private Sub p_oRepViewer_BrowseReport()
   Dim loRS As Recordset
   Dim lsOldProc As String
   Dim lsSQL As String
   Dim lasRepInfo() As String

   lsOldProc = oApp.ProcName("SaveReport")
   On Error GoTo errProc

   Set loRS = New ADODB.Recordset

   With oApp
      lsSQL = "SELECT" & _
                  "  a.sReportID" & _
                  ", a.sReportNm" & _
                  ", b.dGenerate" & _
                  ", b.sRepFName" & _
               " FROM xxxReport a" & _
                  ", xxxReportLog b" & _
               " WHERE a.sReportID = b.sReportID" & _
                  " AND b.sUserIDxx = " & strParm(.UserID)

      loRS.Open lsSQL, .Connection, adOpenStatic, adLockReadOnly, adCmdText
   End With

   If loRS.EOF Then
      MsgBox "No Report is Available for Preview!!!", vbInformation, "Warning"
      Exit Sub
   End If

   lsSQL = KwikBrowse(oApp, loRS, "sReportNm»dGenerate", _
                            "Report Name»Date Generated", "@»MM/DD/YYYY")

   If lsSQL = Empty Then Exit Sub

   lasRepInfo = Split(lsSQL, "»")
   Set p_oRepApp = New Application
   Set p_oReport = p_oRepApp.OpenReport(App.Path & "\" & lasRepInfo(3) & ".rpt")
   Set p_oRepApp = Nothing

  ' ShowReport

endProc:
   Set loRS = Nothing
   oApp.ProcName lsOldProc
   Exit Sub
errProc:
   ShowError lsOldProc
End Sub

Private Sub p_oRepViewer_PrintReport()
   If p_sPgeRange = "xxx" Then
      p_oReport.PrintOutEx False, p_nNoCopies, p_bCollated
   Else
      PrintRangeReport
   End If
End Sub

Private Sub PrintRangeReport()
   Dim lasRange1() As String
   Dim lasRange2() As String
   Dim lnctr As Integer

   lasRange1 = Split(p_sPgeRange, ",")
   For lnctr = 0 To UBound(lasRange1)
      lasRange2 = Split(lasRange1(lnctr), "-")
      Select Case UBound(lasRange2)
      Case 0
         p_oReport.PrintOutEx False, p_nNoCopies, p_bCollated, CLng(lasRange2(0)), CLng(lasRange2(0))
      Case 1
         p_oReport.PrintOutEx False, p_nNoCopies, p_bCollated, CLng(lasRange2(0)), CLng(lasRange2(1))
      Case Else
         Exit Sub
      End Select
   Next
End Sub

Private Sub p_oRepViewer_PrintSetup()
   Dim loSetup As frmPrintSetup

   Set loSetup = New frmPrintSetup
   Set loSetup.AppDriver = oApp
   Set loSetup.Report = p_oReport

   With loSetup
      .Copies = p_nNoCopies
      .Collate = p_bCollated
      .PageRange = p_sPgeRange
      .Orientation = p_nOrientxx
      .Show 1

      If .Cancelled Then GoTo endProc
      p_nNoCopies = .Copies
      p_bCollated = .Collate
      p_sPgeRange = .PageRange
      p_nOrientxx = .Orientation
   End With

endProc:
   Set loSetup = Nothing
   Exit Sub
End Sub
