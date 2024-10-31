VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmViewer 
   BorderStyle     =   0  'None
   Caption         =   "Report Viewer"
   ClientHeight    =   10230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10230
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   2
      Left            =   13995
      TabIndex        =   0
      Top             =   1395
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
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
      Picture         =   "frmViewer.frx":0000
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   9585
      Index           =   0
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   16907
      BackColor       =   12632256
      Begin CRVIEWER9LibCtl.CRViewer9 CRViewer91 
         Height          =   9420
         Left            =   60
         TabIndex        =   1
         Top             =   60
         Width           =   13545
         lastProp        =   500
         _cx             =   23892
         _cy             =   16616
         DisplayGroupTree=   0   'False
         DisplayToolbar  =   -1  'True
         EnableGroupTree =   0   'False
         EnableNavigationControls=   -1  'True
         EnableStopButton=   0   'False
         EnablePrintButton=   -1  'True
         EnableZoomControl=   -1  'True
         EnableCloseButton=   0   'False
         EnableProgressControl=   0   'False
         EnableSearchControl=   -1  'True
         EnableRefreshButton=   -1  'True
         EnableDrillDown =   0   'False
         EnableAnimationControl=   0   'False
         EnableSelectExpertButton=   0   'False
         EnableToolbar   =   -1  'True
         DisplayBorder   =   -1  'True
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
      Height          =   405
      Index           =   1
      Left            =   13995
      TabIndex        =   3
      Top             =   975
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
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
      Picture         =   "frmViewer.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   0
      Left            =   13995
      TabIndex        =   2
      Top             =   555
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
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
      Picture         =   "frmViewer.frx":0EF4
   End
End
Attribute VB_Name = "frmViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private p_oSkin As FormSkin
Private p_bRepPreview As Boolean

Public Event PrintReport()
Public Event PrintSetup()

Property Let AllowBrowse(ByVal Value As Boolean)
   p_bRepPreview = True
End Property

Private Sub Form_Activate()
Dim temp As String
   If p_bRepPreview = True Then
      cmdButton(2).Visible = True
   End If
End Sub

Private Sub Form_Load()

   CenterChildForm mdiMain, Me

   Set p_oSkin = New FormSkin
   Set p_oSkin.AppDriver = oApp
   Set p_oSkin.Form = Me
   p_oSkin.DisableClose = True
   p_oSkin.ApplySkin xeFormTransMaintenance
      
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set p_oSkin = Nothing
End Sub

Private Sub cmdButton_Click(Index As Integer)
Dim loSetup As frmPrintSetup
Dim lasRange1() As String
Dim lasRange2() As String
Dim lsPgeRange As String
Dim lnctr As Integer
Select Case Index
   
   Case 0
      CRViewer91.PrintReport
   Case 1
         Set loSetup = New frmPrintSetup
         Set loSetup.AppDriver = oApp
         Set loSetup.Report = oReport
         
         With loSetup
            .Copies = 1
            .Collate = True
            .PageRange = "xxx"
            .Orientation = 1
            .Show 1
      
            If .Cancelled Then GoTo endProc
      
            If .PageRange = "xxx" Then
               oReport.PrintOutEx False, .Copies, .Collate
            Else
               lsPgeRange = .PageRange
               lasRange1 = Split(lsPgeRange, ",")
               For lnctr = 0 To UBound(lasRange1)
                  lasRange2 = Split(lasRange1(lnctr), "-")
                  Select Case UBound(lasRange2)
                  Case 0
                     oReport.PrintOutEx False, .Copies, .Collate, CLng(lasRange2(0)), CLng(lasRange2(0))
                  Case 1
                     oReport.PrintOutEx False, .Copies, .Collate, CLng(lasRange2(0)), CLng(lasRange2(1))
                  Case Else
                     Exit Sub
                  End Select
               Next
            End If
         End With
   Case 2
      Unload Me
   End Select

endProc:
   Set loSetup = Nothing
   Exit Sub
End Sub

