VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmStatus 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7155
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   7035
      Left            =   45
      TabIndex        =   0
      Top             =   60
      Width           =   10125
      _ExtentX        =   17859
      _ExtentY        =   12409
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   794
      TabMaxWidth     =   3175
      BackColor       =   8421504
      TabCaption(0)   =   "Spareparts"
      TabPicture(0)   =   "frmStatus.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "flxSP"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Inquiries"
      TabPicture(1)   =   "frmStatus.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "flxSMS"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   5295
         Top             =   -435
      End
      Begin MSFlexGridLib.MSFlexGrid flxSP 
         Height          =   6525
         Left            =   -74955
         TabIndex        =   1
         Top             =   480
         Width           =   7800
         _ExtentX        =   13758
         _ExtentY        =   11509
         _Version        =   393216
         FocusRect       =   0
         SelectionMode   =   1
      End
      Begin MSFlexGridLib.MSFlexGrid flxSMS 
         Height          =   6525
         Left            =   45
         TabIndex        =   2
         Top             =   480
         Width           =   7800
         _ExtentX        =   13758
         _ExtentY        =   11509
         _Version        =   393216
         FocusRect       =   0
         SelectionMode   =   1
      End
   End
   Begin VB.Image imgRed 
      Height          =   1500
      Left            =   6345
      Picture         =   "frmStatus.frx":0038
      Top             =   75
      Width           =   1500
   End
   Begin VB.Image imgOrange 
      Height          =   1500
      Left            =   4260
      Picture         =   "frmStatus.frx":3004
      Top             =   75
      Width           =   1500
   End
   Begin VB.Image imgGreen 
      Height          =   1500
      Left            =   2175
      Picture         =   "frmStatus.frx":5FDF
      Top             =   75
      Width           =   1500
   End
   Begin VB.Image imgBlue 
      Height          =   1500
      Left            =   90
      Picture         =   "frmStatus.frx":8FBE
      Top             =   75
      Width           =   1500
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pxeMODULENAME = "frmStatus"
'Private oFormSPTransferTrn As frmSP_Transfer_Posting
'Private oFormSPWarrantyTrn As frmSP_Warranty_Transfer_Posting
'Private oFormSPOrderCnfirm As frmSP_BranchOrder_Reg
Private oFormHotlineSMS As frmSMS

Private oTrans As clsTransMonitor

'Private p_oRS_SPTransfer As Recordset  'Transferred Spareparts to this branch
'Private p_oRS_SPOrderCon As Recordset  'Confirmation of Ordered Spareparts

Private p_oRS_HotlineSMS As Recordset

Dim pnCtr As Integer
Dim pdTimeStart As Date
Dim pbLoaded As Boolean
Dim pbActivated As Boolean

Property Get Loaded() As Boolean
   Loaded = pbLoaded
End Property

Property Let Activate(ByVal bActivated As Boolean)
   pbActivated = bActivated
End Property

Private Sub flxSMS_DblClick()
   With flxSMS
      If .TextMatrix(.Row, 1) <> "" Then Call showMessage(.TextMatrix(.Row, 1))
   End With
End Sub

Private Sub Form_Load()

'   Set p_oRS_SPTransfer = New Recordset  'Transferred Spareparts to this branch
'   Set p_oRS_SPOrderCon = New Recordset  'Confirmation of Ordered Spareparts
'   Set p_oRS_ChkClearng = New Recordset  'List of Checks for Clearing

   Set oTrans = New clsTransMonitor
   Set oTrans.AppDriver = oApp
   oTrans.InitMonitor

   Me.Left = mdiMain.ScaleWidth - Me.Width
   Me.Top = mdiMain.ScaleHeight - Me.Height

   Me.BackColor = oApp.getColor("wb0")
   SSTab1.BackColor = oApp.getColor("wb0")

   Call InitSPGrid
   Call ClearSPGrid
   Call InitSMSGrid
   Call ClearSMSGrid

'   Set oFormSPTransferTrn = New frmSP_Transfer_Posting
'   Set oFormSPWarrantyTrn = New frmSP_Warranty_Transfer_Posting
'   Set oFormSPOrderCnfirm = New frmSP_BranchOrder_Reg
   Set oFormHotlineSMS = New frmSMS

   If Not pbLoaded Then
      If oTrans.StartMonitor Then
'         Set p_oRS_SPTransfer = oTrans.oSPTransfer
'         Call loadSPTransfer
'
'         Set p_oRS_SPOrderCon = oTrans.oSPOrderCon
'         If Not TypeName(p_oRS_SPOrderCon) = "Nothing" Then
'            Call loadSPOrderCon
'         End If
         Set p_oRS_HotlineSMS = oTrans.oHotlineSMS
         Call loadHotlineSMS
               
         pbActivated = False
         pbLoaded = True
      End If
   End If
End Sub

Private Sub InitSPGrid()
   With flxSP
      .Rows = 2
      .Cols = 5

      .RowHeight(0) = 350
      .Font.Size = 10
      .FontWidth = 6
      .Font = "MS Sans Serif"

      'column title
      .TextMatrix(0, 1) = "TRANS. NO"
      .TextMatrix(0, 2) = "BRANCH"
      .TextMatrix(0, 3) = "DATE"
      .TextMatrix(0, 4) = "WRTY"

      .Row = 0
      'column alignment
      For pnCtr = 0 To .Cols - 1
         .Col = pnCtr
         .CellFontBold = True
         .CellAlignment = 3
      Next

      'column width
      .ColWidth(0) = 330
      .ColWidth(1) = 1400
      .ColWidth(2) = 4200
      .ColWidth(3) = 1800
      .ColWidth(4) = 0

      .ColAlignment(0) = 1
      .ColAlignment(1) = 3
      .ColAlignment(2) = 1
      .ColAlignment(3) = 3
      .ColAlignment(4) = 3

      .ScrollBars = flexScrollBarVertical
   End With
End Sub

Private Sub InitSMSGrid()
   With flxSMS
      .Rows = 2
      .Cols = 5

      .RowHeight(0) = 350
      .Font.Size = 10
      .FontWidth = 6
      .Font = "MS Sans Serif"

      'column title
            .TextMatrix(0, 1) = "TRANS. NO."
      .TextMatrix(0, 2) = "DATE RECV."
      .TextMatrix(0, 3) = "DIVISION"
      .TextMatrix(0, 4) = "CONTACT NO."

      .Row = 0
      'column alignment
      For pnCtr = 0 To .Cols - 1
         .Col = pnCtr
         .CellFontBold = True
         .CellAlignment = 3
      Next

      'column width
      .ColWidth(0) = 330
      .ColWidth(1) = 1500
      .ColWidth(2) = 1500
      .ColWidth(3) = 2600
      .ColWidth(4) = 1800

      .ColAlignment(0) = 1
            .ColAlignment(1) = 1
      .ColAlignment(2) = 3
      .ColAlignment(3) = 1
      .ColAlignment(4) = 3

      .ScrollBars = flexScrollBarVertical
   End With
End Sub


Private Sub ClearSPGrid()
   With flxSP
      .Rows = 2

      .TextMatrix(1, 0) = 1
      .TextMatrix(1, 1) = ""
      .TextMatrix(1, 2) = ""
      .TextMatrix(1, 3) = ""
      .TextMatrix(1, 4) = ""

      .Row = 1
      .Col = 1

      .ColSel = .Cols - 1
   End With
End Sub

Private Sub ClearSMSGrid()
   With flxSMS
      .Rows = 2

      .TextMatrix(1, 0) = 1
      .TextMatrix(1, 1) = ""
      .TextMatrix(1, 2) = ""
      .TextMatrix(1, 3) = ""

      .Row = 1
      .Col = 1

      .ColSel = .Cols - 1
   End With
End Sub

Sub loadSPTransfer()
'   Dim lsOldProc As String
'   Dim lnCtr As Integer
'
'   lsOldProc = "loadSPTransfer"
'   'On Error GoTo errProc
'
'   If p_oRS_SPTransfer.EOF Then
'      Call ClearSPGrid
'      GoTo endProc
'   End If
'
'   With flxSP
'      .Rows = p_oRS_SPTransfer.RecordCount + 1
'      .ColWidth(2) = 4200
''      If .Rows > 26 Then .ColWidth(2) = 4000
'
'      For pnCtr = 0 To p_oRS_SPTransfer.RecordCount - 1
'         .TextMatrix(pnCtr + 1, 0) = pnCtr + 1
'         .TextMatrix(pnCtr + 1, 1) = Format(Right(p_oRS_SPTransfer("sTransNox"), 10), "@@@@-@@@@@@")
'         .TextMatrix(pnCtr + 1, 2) = p_oRS_SPTransfer("sBranchNm")
'         .TextMatrix(pnCtr + 1, 3) = Format(p_oRS_SPTransfer("dTransact"), "MMM-DD-YYYY")
'         .TextMatrix(pnCtr + 1, 4) = p_oRS_SPTransfer("cIsWarnty")
'
'         .Row = pnCtr + 1
'         If (pnCtr + 1) Mod 2 = 0 Then
'            For lnCtr = 1 To .Cols - 1
'               .Col = lnCtr
'               .CellBackColor = oApp.getColor("fb0")
'            Next
'         End If
'
'         p_oRS_SPTransfer.MoveNext
'      Next
'
'      .Row = 1
'      .Col = 1
'
'      .ColSel = .Cols - 1
'   End With
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc
End Sub

Sub loadHotlineSMS()
   Dim lsOldProc As String
   Dim lnCtr As Integer

   lsOldProc = "loadHotlineSMS"
   'On Error GoTo errProc

   If p_oRS_HotlineSMS.EOF Then
      Call ClearSMSGrid
      GoTo endProc
   End If
            
         SSTab1.TabCaption(1) = "Inquiry" & " " & "(" & (p_oRS_HotlineSMS.RecordCount) & ")"
   With flxSMS
      .Rows = p_oRS_HotlineSMS.RecordCount + 1
      If .Rows > 26 Then .ColWidth(3) = 2450

      For pnCtr = 0 To p_oRS_HotlineSMS.RecordCount - 1
         .TextMatrix(pnCtr + 1, 0) = pnCtr + 1
         .TextMatrix(pnCtr + 1, 1) = p_oRS_HotlineSMS("sTransnox")
         .TextMatrix(pnCtr + 1, 2) = strLongDate(p_oRS_HotlineSMS("dTransact"))
         .TextMatrix(pnCtr + 1, 3) = p_oRS_HotlineSMS("sDivision")
         .TextMatrix(pnCtr + 1, 4) = p_oRS_HotlineSMS("sMobileNo")

         .Row = pnCtr + 1
         If (pnCtr + 1) Mod 2 = 0 Then
            For lnCtr = 1 To .Cols - 1
               .Col = lnCtr
               .CellBackColor = oApp.getColor("fb0")
            Next
         End If

         p_oRS_HotlineSMS.MoveNext
      Next

      .Row = 1
      .Col = 1

      .ColSel = .Cols - 1
   End With

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc
End Sub

Sub loadSPOrderCon()
'   Dim lsOldProc As String
'   Dim lnCtr As Integer
'
'   lsOldProc = "loadSPOrderCon"
'   'On Error GoTo errProc
'
'   If p_oRS_SPOrderCon.RecordCount = 0 Then
'      SSTab1.TabCaption(1) = "Spareparts" & " " & "(" & (p_oRS_SPTransfer.RecordCount + p_oRS_SPOrderCon.RecordCount) & ")"
'      If p_oRS_SPTransfer.RecordCount = 0 Then
'         Call ClearSPGrid
'         mdiMain.StatusBar1.Panels.Item(4).Picture = Nothing
'      End If
'      GoTo endProc
'   End If
'
'   SSTab1.TabCaption(1) = "Spareparts" & " " & "(" & (p_oRS_SPTransfer.RecordCount + p_oRS_SPOrderCon.RecordCount) & ")"
'   mdiMain.StatusBar1.Panels.Item(4).Picture = imgGreen
'   With flxSP
'      .Rows = p_oRS_SPTransfer.RecordCount + p_oRS_SPOrderCon.RecordCount + 1
'      .ColWidth(2) = 4200
''      If .Rows > 26 Then .ColWidth(2) = 4000
'      p_oRS_SPOrderCon.MoveFirst
'      For pnCtr = 0 To p_oRS_SPOrderCon.RecordCount - 1
'         .TextMatrix(pnCtr + p_oRS_SPTransfer.RecordCount + 1, 0) = pnCtr + 1
'         .TextMatrix(pnCtr + p_oRS_SPTransfer.RecordCount + 1, 1) = Format(Right(p_oRS_SPOrderCon("sTransNox"), 10), "@@@@-@@@@@@")
''         .TextMatrix(pnCtr + p_oRS_SPTransfer.RecordCount + 1, 2) = p_oRS_SPTransfer("sRemarksx")
''         .TextMatrix(pnCtr + p_oRS_SPTransfer.RecordCount + 1, 3) = Format(p_oRS_SPTransfer("dTransact"), "MMM-DD-YYYY")
'         .TextMatrix(pnCtr + p_oRS_SPTransfer.RecordCount + 1, 4) = 2
'
'         .Row = pnCtr + p_oRS_SPTransfer.RecordCount + 1
'         If (pnCtr + p_oRS_SPTransfer.RecordCount + 1) Mod 2 = 0 Then
'            For lnCtr = 1 To .Cols - 1
'               .Col = lnCtr
'               .CellBackColor = oApp.getColor("fb0")
'            Next
'         End If
'
'         p_oRS_SPOrderCon.MoveNext
'      Next
'
'      .Row = 1
'      .Col = 1
'
'      .ColSel = .Cols - 1
'   End With
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc
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

Private Sub Form_Unload(Cancel As Integer)
'   Set oFormSPTransferTrn = Nothing
'   Set oFormSPWarrantyTrn = Nothing
'   Set oFormSPOrderCnfirm = Nothing

'   Set p_oRS_SPTransfer = Nothing  'Transferred Spareparts to this branch
'   Set p_oRS_SPOrderCon = Nothing  'Confirmation of Ordered Spareparts
   
   Set p_oRS_HotlineSMS = Nothing
End Sub

Private Sub showLedger(ByVal nRow As Integer)
'   If p_oRS_3CWthAppnt.RecordCount = 0 Then Exit Sub
'   With oForm3CAppointment
'      p_oRS_3CWthAppnt.Move nRow, adBookmarkFirst
'      .AcctNmbr = p_oRS_3CWthAppnt("sAcctNmbr")
'      .AcctName = p_oRS_3CWthAppnt("xFullName")
'      .DateCreated = Format(p_oRS_3CWthAppnt("dTransact"), "MMMM DD, YYYY")
''      .DateProcess = Format(p_oRS_3CWthAppnt("dEndProcx"), "MMMM DD, YYYY")
'      .Remarks = p_oRS_3CWthAppnt("sRemarksx")
''      .AssAgent = p_oRS_3CWthAppnt("xAsgAgent")
''      .Agent = p_oRS_3CWthAppnt("xAgentxxx")
'      .PayingBranch = p_oRS_3CWthAppnt("xAppBrnch")
'      .CollectingBranch = p_oRS_3CWthAppnt("xColBrnch")
'
'      .Show 1
'   End With
End Sub

Private Sub showSPTransfer(ByVal sTransNox As String)
'   If p_oRS_SPTransfer.RecordCount = 0 Then Exit Sub
'   With oFormSPTransferTrn
'      .TransNox = sTransNox
'      .Show
'   End With
End Sub

Private Sub showMessage(ByVal sTransNox As String)
   If p_oRS_HotlineSMS.RecordCount = 0 Then Exit Sub
   With oFormHotlineSMS
      .TransNox = sTransNox
      .Show
      If Not .MessageLoaded Then Unload oFormHotlineSMS
   End With
End Sub

Private Sub showSPWarranty(ByVal sTransNox As String)
'   If p_oRS_SPTransfer.RecordCount = 0 Then Exit Sub
'   With oFormSPWarrantyTrn
'      .TransactionNo = sTransNox
'      .Show
'   End With
End Sub

Private Sub showSPOrderCon(ByVal sTransNox As String)
'   If p_oRS_SPOrderCon.RecordCount = 0 Then Exit Sub
'   With oFormSPOrderCnfirm
'      .TransactionNo = sTransNox
'      .Show
'   End With
End Sub

Private Sub Timer1_Timer()
   If pbLoaded Then
      If pbActivated Then
         If oTrans.StartMonitor Then
'            Select Case oTrans.LastList
'            Case -1
'            Case 1
'               Set p_oRS_3CWthAppnt = oTrans.o3CAppointment
'               Call load3CAppointment
'               pbActivated = False
'            Case 2
'               Set p_oRS_SPTransfer = oTrans.oSPTransfer
'               Call loadSPTransfer
'               pbActivated = False
'            Case 3
'               Set p_oRS_SPOrderCon = oTrans.oSPOrderCon
'               If Not TypeName(p_oRS_SPOrderCon) = "Nothing" Then
'                  Call loadSPOrderCon
'               End If
'               pbActivated = False
'            Case 4
'               Set p_oRS_MCTransfer = oTrans.oMCTransfer
'               Call loadMCTransfer
'               pbActivated = False
'            Case 5
'               Set p_oRS_ChkClearng = oTrans.oChkClearng
'               Call loadChkClearng
'               pbActivated = False
'            End Select
            Set p_oRS_HotlineSMS = oTrans.oHotlineSMS
            Call loadHotlineSMS
            pbActivated = False
         End If
      End If
   End If
End Sub
