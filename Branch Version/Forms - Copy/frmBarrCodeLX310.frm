VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrcontrol.ocx"
Begin VB.Form frmBarrCodeLX310 
   BorderStyle     =   0  'None
   Caption         =   "Bar Code Creator"
   ClientHeight    =   7425
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10140
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7425
   ScaleWidth      =   10140
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   735
      Index           =   1
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   8445
      _ExtentX        =   14896
      _ExtentY        =   1296
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   4875
         MaxLength       =   50
         TabIndex        =   3
         Top             =   90
         Width           =   1440
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   7215
         MaxLength       =   50
         TabIndex        =   9
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   2
         Left            =   7215
         MaxLength       =   50
         TabIndex        =   5
         Top             =   90
         Width           =   1095
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   1050
         TabIndex        =   7
         Top             =   360
         Width           =   5265
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   0
         Left            =   1050
         MaxLength       =   50
         TabIndex        =   1
         Top             =   90
         Width           =   2940
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Pur. Date"
         Height          =   285
         Index           =   3
         Left            =   4125
         TabIndex        =   2
         Top             =   105
         Width           =   795
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
         Height          =   285
         Index           =   2
         Left            =   6435
         TabIndex        =   8
         Top             =   375
         Width           =   750
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Price"
         Height          =   285
         Index           =   1
         Left            =   6435
         TabIndex        =   4
         Top             =   105
         Width           =   885
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   285
         Index           =   9
         Left            =   120
         TabIndex        =   6
         Top             =   375
         Width           =   840
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bar Code"
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   105
         Width           =   885
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   3
      Left            =   8790
      TabIndex        =   14
      Top             =   2040
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
      Caption         =   "   &Close"
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
      Picture         =   "frmBarrCodeLX310.frx":0000
      CaptionAlign    =   0
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   5970
      Index           =   0
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   1335
      Width           =   8445
      _ExtentX        =   14896
      _ExtentY        =   10530
      BackColor       =   12632256
      Begin CRVIEWER9LibCtl.CRViewer9 CRViewer91 
         Height          =   5805
         Left            =   60
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   60
         Width           =   8280
         lastProp        =   600
         _cx             =   14605
         _cy             =   10239
         DisplayGroupTree=   0   'False
         DisplayToolbar  =   -1  'True
         EnableGroupTree =   0   'False
         EnableNavigationControls=   -1  'True
         EnableStopButton=   0   'False
         EnablePrintButton=   -1  'True
         EnableZoomControl=   -1  'True
         EnableCloseButton=   0   'False
         EnableProgressControl=   -1  'True
         EnableSearchControl=   0   'False
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
         EnableLogonPrompts=   -1  'True
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   0
      Left            =   8790
      TabIndex        =   11
      Top             =   780
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
      Caption         =   "   &Print"
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
      Picture         =   "frmBarrCodeLX310.frx":077A
      CaptionAlign    =   0
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   2
      Left            =   8790
      TabIndex        =   13
      Top             =   1620
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
      Caption         =   "   &Browse"
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
      Picture         =   "frmBarrCodeLX310.frx":0EF4
      CaptionAlign    =   0
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   1
      Left            =   8790
      TabIndex        =   12
      Top             =   1200
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
      Caption         =   "   Set&up"
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
      Picture         =   "frmBarrCodeLX310.frx":166E
      CaptionAlign    =   0
   End
End
Attribute VB_Name = "frmBarrCodeLX310"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents oDriver As clsFormDriver
Attribute oDriver.VB_VarHelpID = -1
Private oSkin As clsFormSkin
Private bLoaded As Boolean

Dim psSelected() As String

Private Sub cmdButton_Click(Index As Integer)
   Dim lnCtr As Integer

   Select Case Index
      Case 0 'OK
         oReport.PrintOutEx False, 1
'         CRViewer91.PrintReport
      Case 1 'Setup
         PrintSetup
      Case 2 'Browse BarrCode
         SearchBarCode False
      Case 3 'Cancel
         Unload Me
'         .txtothers(0).SetFocus
   End Select
End Sub

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0

   If bLoaded = False Then
      bLoaded = True
      oDriver.DisableTextbox 1
      oDriver.DisableTextbox 2
   End If
End Sub

Private Sub Form_Load()
   Dim lsSQL As String

   CenterChildForm mdiMain, Me
   bLoaded = False
   
   Set oDriver = New clsFormDriver
   Set oDriver.AppDriver = oApp
   Set oDriver.MainForm = Me
   
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransDetail
      
   clearFields
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oDriver = Nothing
   Set oSkin = Nothing
   Set oReport = Nothing
End Sub

Private Sub oDriver_DisableOtherControl()
   oDriver.DisableTextbox 1
   oDriver.DisableTextbox 2
End Sub

Private Sub oDriver_EnableOtherControl()
   oDriver.DisableTextbox 1
   oDriver.DisableTextbox 2
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   oDriver.ColumnIndex = Index
   txtfield(Index).BackColor = &HE1FEFF
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyReturn, vbKeyUp, vbKeyDown
      Select Case KeyCode
      Case vbKeyReturn, vbKeyDown
         SetNextFocus
      Case vbKeyUp
         SetPreviousFocus
      End Select
   End Select
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lsSQL As String

   If KeyCode = vbKeyF3 Or KeyCode = 13 Then
      If Index = 0 Then
         SearchBarCode False
      End If
      If txtfield(Index).Text <> "" Then SetNextFocus
      KeyCode = 0
   End If
End Sub

Private Sub SearchBarCode(ByVal SearchValue As Boolean)
   Dim lsSearch As String
   Dim lrs As ADODB.Recordset
   Dim lsSQL As String
   
   Set lrs = New ADODB.Recordset
   lsSQL = "SELECT" _
                & " b.sBarrcode, " _
                & " b.nSelPrice, " _
                & " c.sBrandNme, " _
                & " d.sModelNme, " _
                & " b.sDescript, " _
                & " e.sColorNme  " _
                
   lsSQL = lsSQL _
         & " FROM CP_Inventory b" _
               & " LEFT JOIN CP_Brand c " _
                  & " ON b.sBrandIDx = c.sBrandIDx " _
               & " LEFT JOIN CP_Model d " _
                  & " ON b.sModelIDx = d.sModelIDx " _
               & " LEFT JOIN Color e " _
                  & " ON b.sColorIDx = e.sColorIDx " _
         & " WHERE b.cRecdStat = '" & xeRecStateActive & "' "

   If SearchValue Then
      lsSQL = lsSQL & " AND sBarrCode = '" & txtfield(0).Text & "'"
      lsSQL = lsSQL & " ORDER BY sBarrCode"
   Else
      lsSQL = lsSQL & " AND sBarrCode LIKE '" & txtfield(0).Text & "%' "
      lsSQL = lsSQL & " ORDER BY sBarrCode"
   End If
   lrs.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
            
      If lrs.RecordCount = 1 Then
         txtfield(0).Text = lrs("sBarrCode")
         txtfield(1).Text = Trim(IIf(IsNull(lrs("sBrandNme")), "", IIf(lrs("sBrandNme") = "N-O-N-E", "", lrs("sBrandNme"))) _
                              & " " & IIf(IsNull(lrs("sModelNme")), "", IIf(lrs("sModelNme") = "N-O-N-E", "", lrs("sModelNme"))) _
                              & " " & IIf(IsNull(lrs("sDescript")), "", lrs("sDescript")) _
                              & " " & IIf(IsNull(lrs("sColorNme")), "", lrs("sColorNme")))
         txtfield(2).Text = Format(lrs("nSelPrice"), "#,##0.00")
      ElseIf lrs.RecordCount > 1 Then
         lsSearch = KwikBrowse(oApp, lrs, _
                           "sBarrcode»sBrandNme»sModelNme»sDescript»sColorNme", _
                           "Bar Code»Brand»Model»Descript»Color")
                         
         If lsSearch <> "" Then
            psSelected = Split(lsSearch, "»")
            txtfield(0).Text = psSelected(0)
            txtfield(1).Text = Trim(IIf(IsNull(psSelected(2)), "", IIf(psSelected(2) = "N-O-N-E", "", psSelected(2))) _
                              & " " & IIf(IsNull(psSelected(3)), "", IIf(psSelected(3) = "N-O-N-E", "", psSelected(3))) _
                              & " " & IIf(IsNull(psSelected(4)), "", psSelected(4)) _
                              & " " & IIf(IsNull(psSelected(5)), "", psSelected(5)))
            txtfield(2).Text = Format(psSelected(1), "#,##0.00")
         End If
   Else
      clearFields
      txtfield(0).SetFocus
   End If
   
   Set lrs = Nothing
End Sub

Private Sub PrintSetup()
   Dim loSetup As frmPrintSetup
   Dim lasRange1() As String
   Dim lasRange2() As String
   Dim lsPgeRange As String
   Dim lnCtr As Integer

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
         For lnCtr = 0 To UBound(lasRange1)
            lasRange2 = Split(lasRange1(lnCtr), "-")
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

endProc:
   Set loSetup = Nothing
   Exit Sub
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   
   Select Case Index
   Case 3
      If txtfield(Index).Text = 0 Then
         MsgBox "Invalid Quantity!!!" & vbCrLf & _
         "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
         Exit Sub
      Else
         DisplayReport
      End If
   End Select
   txtfield(Index).BackColor = &HFFFFFF
End Sub

Private Sub clearFields()
   Dim lnCtr As Integer
   
   For lnCtr = 0 To 2
      txtfield(lnCtr).Text = ""
   Next
   txtfield(3).Text = 0
   txtfield(4).Text = ""
End Sub

Private Sub DisplayReport()
   Dim lnCtr As Integer
   Dim lrs As ADODB.Recordset

   Set lrs = New ADODB.Recordset
   lrs.Fields.Append "sField01", adVarChar, 150
   lrs.Fields.Append "sField02", adVarChar, 150
   lrs.Fields.Append "sField03", adVarChar, 150
   lrs.Fields.Append "lField01", adCurrency, 10
   lrs.Fields.Append "sField04", adVarChar, 6
   lrs.Open

   For lnCtr = 0 To (txtfield(3).Text - 1)
      lrs.AddNew
      lrs("sField01").Value = txtfield(1).Text
      lrs("sField02").Value = "*" + txtfield(0).Text + "*"
      lrs("sField03").Value = "*" + txtfield(0).Text + "*"
      lrs("lField01").Value = Format(txtfield(2).Text, "#,##0.00")
      lrs("sField04").Value = Format(txtfield(4).Text, "ddyymm")
   Next
   
   Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\CP_BarrCode_LX310.rpt")
   oReport.DiscardSavedData
   oReport.FieldMappingType = crAutoFieldMapping
   oReport.Database.SetDataSource lrs
   
   CRViewer91.ReportSource = oReport
   CRViewer91.ViewReport
   CRViewer91.Zoom 1
   Set lrs = Nothing
End Sub

Property Set ReportSource(ByVal Source As Report)
   Set oReport = Source
End Property

Private Sub Form_Initialize()
   Set oReport = New Report
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
   Case 3
      If Not IsNumeric(txtfield(Index).Text) Then txtfield(Index).Text = 0#
   Case 4
      If Not IsDate(txtfield(Index).Text) Then
         txtfield(Index).Text = ""
      Else
         txtfield(Index).Text = Format(txtfield(Index).Text, "MMM dd, yyyy")
      End If
   End Select
End Sub
