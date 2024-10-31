VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmRepViewer 
   BorderStyle     =   0  'None
   Caption         =   "Print Preview"
   ClientHeight    =   12795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   21930
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   12795
   ScaleWidth      =   21930
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4650
      Top             =   420
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   420
      Index           =   3
      Left            =   3495
      TabIndex        =   0
      Top             =   480
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   741
      Caption         =   "&Close"
      AccessKey       =   "C"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmRepViewer.frx":0000
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   11760
      Index           =   0
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   915
      Width           =   21690
      _ExtentX        =   38259
      _ExtentY        =   20743
      BackColor       =   12632256
      BorderStyle     =   1
      Begin CRVIEWER9LibCtl.CRViewer9 CRViewer91 
         Height          =   11625
         Left            =   45
         TabIndex        =   2
         Top             =   45
         Width           =   21555
         lastProp        =   500
         _cx             =   38021
         _cy             =   20505
         DisplayGroupTree=   0   'False
         DisplayToolbar  =   -1  'True
         EnableGroupTree =   0   'False
         EnableNavigationControls=   -1  'True
         EnableStopButton=   0   'False
         EnablePrintButton=   -1  'True
         EnableZoomControl=   -1  'True
         EnableCloseButton=   0   'False
         EnableProgressControl=   0   'False
         EnableSearchControl=   0   'False
         EnableRefreshButton=   -1  'True
         EnableDrillDown =   0   'False
         EnableAnimationControl=   0   'False
         EnableSelectExpertButton=   0   'False
         EnableToolbar   =   -1  'True
         DisplayBorder   =   0   'False
         DisplayTabs     =   0   'False
         DisplayBackgroundEdge=   0   'False
         SelectionFormula=   ""
         EnablePopupMenu =   0   'False
         EnableExportButton=   -1  'True
         EnableSearchExpertButton=   0   'False
         EnableHelpButton=   0   'False
         LaunchHTTPHyperlinksInNewBrowser=   0   'False
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   420
      Index           =   0
      Left            =   75
      TabIndex        =   1
      Top             =   480
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   741
      Caption         =   "&Print"
      AccessKey       =   "P"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmRepViewer.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   420
      Index           =   2
      Left            =   2355
      TabIndex        =   4
      Top             =   480
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   741
      Caption         =   "&Browse"
      AccessKey       =   "B"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmRepViewer.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   420
      Index           =   1
      Left            =   1215
      TabIndex        =   3
      Top             =   480
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   741
      Caption         =   "Set&up"
      AccessKey       =   "u"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmRepViewer.frx":166E
   End
End
Attribute VB_Name = "frmRepViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmRepViewer"

Private p_oSkin As clsFormSkin
Private p_oSource As Report

Private p_bPreviewx As Boolean

Property Set ReportSource(ByVal Source As Report)
   Set p_oSource = Source
End Property

Property Let AllowBrowse(ByVal Value As Boolean)
   p_bPreviewx = True
End Property

Private Sub Form_Activate()
   If p_bPreviewx = True Then
      cmdButton(2).Visible = True
      cmdButton(3).Visible = True
      cmdButton(3).Left = 3495
   Else
      cmdButton(2).Visible = False
      cmdButton(3).Left = 2355
   End If
   CRViewer91.Refresh
End Sub

Private Sub Form_Initialize()
   Set p_oSource = New Report
End Sub

Private Sub Form_Load()
   CenterChildForm mdiMain, Me

   Set p_oSkin = New clsFormSkin
   Set p_oSkin.AppDriver = oApp
   Set p_oSkin.Form = Me
   p_oSkin.DisableClose = True
   p_oSkin.ApplySkin xeFormLedger
   
   CRViewer91.ReportSource = p_oSource
   CRViewer91.ViewReport
   DoEvents
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set p_oSkin = Nothing
   Set p_oSource = Nothing
End Sub

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   
   lsOldProc = "cmdButton_Click"
   ''On Error GoTo errProc

   Select Case Index
   Case 0
      p_oSource.PrintOutEx False, 1
   Case 1
      PrintSetup
   Case 2
      If cmdButton(Index).Caption = "&Browse" Then
         BrowseReport
      End If
   Case 3
      Unload Me
   End Select
  
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Sub PrintSetup()
   Dim loSetup As frmPrintSetup
   Dim lasRange1() As String
   Dim lasRange2() As String
   Dim lsPgeRange As String
   Dim lnCtr As Integer
   Dim lsOldProc As String
   
   lsOldProc = "PrintSetup"
   ''On Error GoTo errProc
   
   Set loSetup = New frmPrintSetup
   Set loSetup.AppDriver = oApp
   Set loSetup.Report = p_oSource
   
   With loSetup
      .Copies = 1
      .Collate = True
      .PageRange = "xxx"
      .Orientation = p_oSource.PaperOrientation - 1
      .Show 1

      If .Cancelled Then GoTo endProc

      If .PageRange = "xxx" Then
         p_oSource.PrintOutEx False, .Copies, .Collate
      Else
         lsPgeRange = .PageRange
         lasRange1 = Split(lsPgeRange, ",")
         For lnCtr = 0 To UBound(lasRange1)
            lasRange2 = Split(lasRange1(lnCtr), "-")
            Select Case UBound(lasRange2)
            Case 0
               p_oSource.PrintOutEx False, .Copies, .Collate, CLng(lasRange2(0)), CLng(lasRange2(0))
            Case 1
               p_oSource.PrintOutEx False, .Copies, .Collate, CLng(lasRange2(0)), CLng(lasRange2(1))
            Case Else
               Exit Sub
            End Select
         Next
      End If
   End With

endProc:
   Set loSetup = Nothing
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )"
End Sub

Private Sub BrowseReport()
   Dim loRepApp As Application
   Dim loRS As Recordset
   Dim lsOldProc As String
   Dim lsSQL As String
   Dim lasRepInfo() As String

   lsOldProc = "BrowseReport"
   ''On Error GoTo errProc
   
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
   Set loRepApp = New Application
   Set p_oSource = loRepApp.OpenReport(App.Path & "\" & lasRepInfo(3) & ".rpt")
   Set loRepApp = Nothing

   With CRViewer91
      .ReportSource = p_oSource
      .ViewReport
   End With
   
endProc:
   Set loRS = Nothing
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )"
End Sub

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


