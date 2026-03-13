VERSION 5.00
Object = "{0A7B56A6-35D0-4533-91DA-1715D3A0DD3E}#1.1#0"; "xrGridControl.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmTransfer_Acceptance_NoSerial 
   BorderStyle     =   0  'None
   Caption         =   "Transfer Acceptance (No IMEI No.)"
   ClientHeight    =   8055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11610
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8055
   ScaleWidth      =   11610
   ShowInTaskbar   =   0   'False
   Begin xrGridEditor.GridEditor GridEditor1 
      Height          =   5010
      Left            =   1635
      TabIndex        =   15
      Tag             =   "et0;eb0;et0;bc2"
      Top             =   2850
      Width           =   9780
      _ExtentX        =   17251
      _ExtentY        =   8837
      AllowBigSelection=   -1  'True
      AutoAdd         =   -1  'True
      AutoNumber      =   -1  'True
      BACKCOLOR       =   -2147483643
      BACKCOLORBKG    =   8421504
      BACKCOLORFIXED  =   -2147483633
      BACKCOLORSEL    =   -2147483635
      BORDERSTYLE     =   1
      COLS            =   2
      FILLSTYLE       =   0
      FIXEDCOLS       =   1
      FIXEDROWS       =   1
      FOCUSRECT       =   1
      EDITORBACKCOLOR =   -2147483643
      EDITORFORECOLOR =   -2147483640
      FORECOLOR       =   -2147483640
      FORECOLORFIXED  =   -2147483630
      FORECOLORSEL    =   -2147483634
      FORMATSTRING    =   ""
      Object.HEIGHT          =   5010
      GRIDCOLOR       =   12632256
      GRIDCOLORFIXED  =   0
      BeginProperty GRIDFONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GRIDLINES       =   1
      GRIDLINESFIXED  =   2
      GRIDLINEWIDTH   =   1
      MOUSEICON       =   "frmTransfer_Acceptance_NoSerial.frx":0000
      MOUSEPOINTER    =   0
      REDRAW          =   -1  'True
      RIGHTTOLEFT     =   0   'False
      ROWS            =   2
      SCROLLBARS      =   3
      SCROLLTRACK     =   0   'False
      SELECTIONMODE   =   0
      Object.TOOLTIPTEXT     =   ""
      WORDWRAP        =   0   'False
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   495
      Index           =   1
      Left            =   1575
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   873
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   1395
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   105
         Width           =   2220
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   2
         Left            =   4980
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   105
         Width           =   4785
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transmittal No."
         Height          =   285
         Index           =   3
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Branch Origin"
         Height          =   285
         Index           =   19
         Left            =   3750
         TabIndex        =   2
         Top             =   105
         Width           =   975
      End
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   5160
      Left            =   1575
      Tag             =   "wt0;fb0"
      Top             =   2775
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   9102
      BackColor       =   12632256
      ClipControls    =   0   'False
      BorderStyle     =   1
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1695
      Index           =   0
      Left            =   1575
      Tag             =   "wt0;fb0"
      Top             =   1050
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   2990
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   4
         Left            =   4980
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   150
         Width           =   2475
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   6
         Left            =   4980
         MaxLength       =   128
         MultiLine       =   -1  'True
         TabIndex        =   11
         Text            =   "frmTransfer_Acceptance_NoSerial.frx":001C
         Top             =   405
         Width           =   2475
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   0
         Left            =   1395
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   150
         Width           =   2235
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   900
         Index           =   7
         Left            =   1395
         MaxLength       =   128
         MultiLine       =   -1  'True
         TabIndex        =   13
         Text            =   "frmTransfer_Acceptance_NoSerial.frx":0022
         Top             =   660
         Width           =   8370
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   3
         Left            =   1395
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   405
         Width           =   2235
      End
      Begin VB.Label label 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "UNKNOWN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   7800
         TabIndex        =   14
         Tag             =   "eb0;wb0"
         Top             =   150
         Width           =   1965
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction No"
         Height          =   285
         Index           =   9
         Left            =   120
         TabIndex        =   4
         Top             =   165
         Width           =   1185
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date Received"
         Height          =   285
         Index           =   4
         Left            =   3750
         TabIndex        =   8
         Top             =   165
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date Transferred"
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   405
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   285
         Index           =   7
         Left            =   135
         TabIndex        =   12
         Top             =   660
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Approved By "
         Height          =   285
         Index           =   2
         Left            =   3750
         TabIndex        =   10
         Top             =   405
         Width           =   1200
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   1
      Left            =   90
      TabIndex        =   16
      Top             =   4770
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
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
      Picture         =   "frmTransfer_Acceptance_NoSerial.frx":0028
      CaptionAlign    =   0
      BackColor       =   14286077
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   3
      Left            =   90
      TabIndex        =   17
      Top             =   6030
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
      Picture         =   "frmTransfer_Acceptance_NoSerial.frx":07A2
      CaptionAlign    =   0
      BackColor       =   14286077
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   2
      Left            =   90
      TabIndex        =   18
      Top             =   5190
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
      Caption         =   "&New"
      AccessKey       =   "N"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmTransfer_Acceptance_NoSerial.frx":0F1C
      CaptionAlign    =   0
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   0
      Left            =   90
      TabIndex        =   19
      Top             =   5610
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
      Caption         =   "P&ost"
      AccessKey       =   "o"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmTransfer_Acceptance_NoSerial.frx":1696
      CaptionAlign    =   0
      BackColor       =   14286077
   End
End
Attribute VB_Name = "frmTransfer_Acceptance_NoSerial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  Programmed By  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  Rosalyn Lazo Descallar  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  Date Started  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  February 23, 2007  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'

Option Explicit

Private WithEvents oDriver As FormDriver
Attribute oDriver.VB_VarHelpID = -1
Private oSkin As FormSkin
Private bLoaded As Boolean
Private oRS As New ADODB.Recordset
Private poFileSys As FileSystemObject
Dim oForm As frmTransfer_Export

Dim txtfieldGotfocus As Boolean
Dim pbnewitem As Boolean
Dim psSelected() As String

Dim pnindex As Integer
Dim pnCtr As Integer
Dim temp As String
Dim RefDate As Date
Dim Time As String

Private Sub cmdButton_Click(Index As Integer)
Dim lsSearch As String
Dim lsRep As Integer
Dim lsSQL As String
Dim lnrow As Long
Dim response As String

With GridEditor1
   Select Case Index
      Case 0 'Post
         If label.Caption = "UNKNOWN" Then
            oDriver.RecordSave
         Else
            MsgBox "Transaction Already Posted!!!", vbCritical, "Warning"
         End If
      Case 1 'browse
         Search_Transmittal
         ShowGrid
      Case 2 'New
         oDriver.RecordNew
         ShowButton
         EmptyGrid
      Case 3 'close
         Unload Me
   End Select
End With
End Sub
Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0
      
   If bLoaded = False Then
      oDriver.RecordNew
      bLoaded = True
      label.Caption = "UNKNOWN"
      ShowButton
   End If
   GridEditor1.Refresh
End Sub

Private Sub ClearFields()
Dim Index As Integer

For Index = 0 To 7
   Select Case Index
      Case 0 To 2, 6, 7
         txtfield(Index).Text = ""
      Case 3, 4
         txtfield(Index).Text = Format(Date, "MMMM dd, yyyy")
   End Select
Next
txtfield(1).Enabled = True
txtfield(2).Enabled = True

End Sub

Private Sub ShowButton()
   If xrFrame1(1).Enabled = False Then xrFrame1(1).Enabled = True
   cmdButton(0).Visible = True
   cmdButton(1).Visible = True
   cmdButton(2).Visible = True
   cmdButton(3).Visible = True
End Sub

Private Sub oDriver_DisableOtherControl()
   oDriver.DisableTextbox 0
   oDriver.EnableTextbox 1
   oDriver.EnableTextbox 2
End Sub

Private Sub oDriver_EnableOtherControl()
   oDriver.DisableTextbox 0
   oDriver.EnableTextbox 1
   oDriver.EnableTextbox 2
End Sub

Private Sub oDriver_InitValue()
   oDriver.FieldReference(0) = True
   If Not oDriver.SetValue(0, getNextCode("CP_Transfer_Master", "sTransNox", True, oApp.Connection, True, oApp.BranchCode)) Then Exit Sub
   oDriver.FieldValue(3) = Date
   oDriver.FieldValue(4) = Format(Date, "MMMM dd, yyyy")
   label.Caption = "UNKNOWN"
   pbnewitem = True
End Sub

Private Sub Form_Load()
Dim lsSQL As String

   CenterChildForm mdiMain, Me
   bLoaded = False
    
   Set oDriver = New FormDriver
   Set oDriver.AppDriver = oApp
   Set oDriver.MainForm = Me
   
   InitGrid
   
   Set oSkin = New FormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransaction
   Set oForm = New frmTransfer_Export
     
   oDriver.RecQuery = "SELECT" _
                        & " sTransNox, " _
                        & " sReferNox, " _
                        & " sOriginxx, " _
                        & " dTransact, " _
                        & " dReceived, " _
                        & " sRequestx, " _
                        & " sApproved, " _
                        & " sRemarksx, " _
                        & " cTranStat, " _
                        & " sDestinat, " _
                        & " cReceived, " _
                        & " nEntryNox, " _
                        & " sModified, " _
                        & " dModified, " _
                        & " vTimeStmp  " _
                  & " FROM CP_Transfer_Master " _

   oDriver.InitRecForm

   oDriver.BrowseQuery = "SELECT" _
                  & " Distinct " _
                  & " a.sTransNox, " _
                  & " a.sReferNox, " _
                  & " a.dTransact, " _
                  & " b.sBranchNm  " _
            & " FROM CP_Transfer_Master a " _
               & " LEFT JOIN Branch b " _
                  & " ON a.sOriginxx = b.sBranchCd " _
            & " WHERE a.cTranStat = 0 " _
               & " AND a.sOriginxx <> '" & oApp.BranchCode & "'" _
               & " AND a.sDestinat = '" & oApp.BranchCode & "'" _
            & " ORDER BY a.dTransact Desc "

   oDriver.InitRecForm

   oDriver.BrowseFTitle(0) = "Transaction No"
   oDriver.BrowseFTitle(1) = "Transmittal No"
   oDriver.BrowseFTitle(2) = "Date Transferred"
   oDriver.BrowseFTitle(3) = "Branch Origin"
   
   oDriver.BrowseFFormat(2) = "MMMM dd, yyyy"
   
   'Origin
   oDriver.LookupQuery(2) = "SELECT" _
                           & " a.sBranchCd, " _
                           & " a.sBranchNm, " _
                           & " a.sAddressx + ' ' + b.sTownName xAddressx " _
                     & " FROM Branch a " _
                        & "LEFT JOIN TownCity b " _
                           & " ON a.sTownIdxx = b.sTownIDxx " _
                     & " WHERE a.cRecdStat = '" & xeRecStateActive & "'" _
                           & "AND a.sBranchCd <> '" & oApp.BranchCode & "'" _
                     & " ORDER BY sBranchNm "

   oDriver.LookupReference(2) = "sBranchCd»sBranchNm»xAddressx"
   oDriver.LookupColumn(2) = "sBranchNm»xAddressx"
   oDriver.LookupTitle(2) = "Branch»Address"

       
   oDriver.FieldFormat(0) = "@@-@@@@@@@@"
   oDriver.FieldSize(0) = Len(oDriver.FieldFormat(0))

   oDriver.FieldStart = 1
   oDriver.FieldFormat(3) = "MMMM DD, YYYY"
   oDriver.FieldFormat(4) = "MMMM DD, YYYY"
   EmptyGrid

End Sub

Private Sub InitGrid()
Dim Index As Integer

    With GridEditor1
      .Rows = 2
      .Cols = 6
      .Font = "MS Sans Serif"
              
      'column title
      .TextMatrix(0, 1) = "Bar Code"
      .TextMatrix(0, 2) = "Particulars"
      .TextMatrix(0, 3) = "SRP"
      .TextMatrix(0, 4) = "Stock ID"
      .TextMatrix(0, 5) = "Qty"
      .Row = 0
      
      'column width
      .ColWidth(0) = 400
      .ColWidth(1) = 2200
      .ColWidth(2) = 5350
      .ColWidth(3) = 1000
      .ColWidth(4) = 0
      .ColWidth(5) = 670
            
      For Index = 0 To 5
         .ColAlignment(1) = 1
      Next
      .ColAlignment(5) = 6
      
      .ColEnabled(2) = False
      .ColEnabled(3) = False
      .ColEnabled(4) = False
      
      .Row = 1
    End With
End Sub

Private Sub Search_Transmittal()
Dim orig As String
Dim lsSQL As String
Dim lsCondition As String

   orig = oDriver.BrowseQuery
   If pnindex = 1 Then
      lsCondition = " a.sReferNox like '%" & txtfield(1).Text & "%'"
      lsSQL = AddCondition(oDriver.BrowseQuery, lsCondition)
      oDriver.BrowseQuery = lsSQL
   ElseIf pnindex = 2 Then
      lsCondition = "a.sOriginxx = '" & oDriver.FieldValue(2) & "'"
      lsSQL = AddCondition(oDriver.BrowseQuery, lsCondition)
      oDriver.BrowseQuery = lsSQL
   End If
   
   oDriver.BrowseRecord
   oDriver.BrowseQuery = orig

End Sub

Private Sub ShowGrid()
Dim lsSQL As String
Dim lnCtr As Integer
Dim lnrow As Long

   lsSQL = "SELECT" _
               & " Distinct " _
               & " a.sTransNox, " _
               & " a.nEntryNox, " _
               & " a.sStockIdx, " _
               & " a.nQuantity, " _
               & " a.nUnitPrce, " _
               & " b.sBarrCode, " _
               & " b.sDescript, " _
               & " b.sCategIDx, " _
               & " c.sBrandNme, " _
               & " d.sModelNme, " _
               & " e.sColorNme, " _
               & " b.nSelPrice  " _
         & " FROM CP_Transfer_Detail a " _
               & " LEFT JOIN CP_Inventory b " _
                  & " ON a.sStockIDx = b.sStockIDx " _
               & " LEFT JOIN Brand c " _
                  & " ON b.sBrandIDx = c.sBrandIDx " _
               & " LEFT JOIN Model d " _
                  & " ON b.sModelIDx = d.sModelIDx " _
               & " LEFT JOIN Color e " _
                  & " ON b.sColorIDx = e.sColorIDx " _
         & " WHERE a.sTransNox = '" & oDriver.FieldValue(0) & " '" _
         & " ORDER BY a.nEntryNox "
   If oRS.State = adStateOpen Then oRS.Close
   oRS.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText

   If oRS.RecordCount <> 0 Then
      With GridEditor1
         .Rows = oRS.RecordCount + 1
         For lnCtr = 0 To oRS.RecordCount - 1
            .TextMatrix(lnCtr + 1, 0) = oRS("nEntryNox")
            .TextMatrix(lnCtr + 1, 1) = oRS("sBarrCode")
            .TextMatrix(lnCtr + 1, 2) = Trim(IIf(IsNull(oRS("sBrandNme")), "", oRS("sBrandNme")) & " " & _
                                          IIf(IsNull(oRS("sModelNme")), "", oRS("sModelNme")) & " " & _
                                          IIf(IsNull(oRS("sDescript")), "", oRS("sDescript")) & " " & _
                                          IIf(IsNull(oRS("sColorNme")), "", oRS("sColorNme")))
            .TextMatrix(lnCtr + 1, 3) = Format(oRS("nSelPrice"), "#,##0.00")
            .TextMatrix(lnCtr + 1, 4) = oRS("sStockIDx")
            .TextMatrix(lnCtr + 1, 5) = oRS("nQuantity")
            oRS.MoveNext
         Next
         If .Rows > 18 Then
            .ColWidth(2) = 5200
         Else
            .ColWidth(2) = 5350
         End If
         For pnCtr = 1 To .Cols - 1
            .ColEnabled(pnCtr) = False
         Next
      End With
      txtfield(4).Enabled = True
      txtfield(4).Locked = False
      txtfield(4).SetFocus
   Else
      Exit Sub
   End If
   Set oRS = Nothing
End Sub

Private Sub EmptyGrid()
   With GridEditor1
      .Rows = 2
      For pnCtr = 1 To .Cols - 1
         .TextMatrix(1, pnCtr) = ""
      Next
      .ColEnabled(1) = True
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oDriver = Nothing
   Set oSkin = Nothing
End Sub

Private Sub GridEditor1_GotFocus()
   GridEditor1.Col = 1
End Sub

Private Sub oDriver_LoadOtherData()
   txtfield(4).Text = Format(Date, "MMMM dd,yyyy")
   label.Caption = "UNKNOWN"
End Sub

Private Sub oDriver_Save(Saved As Boolean)
   Saved = False
End Sub

Private Sub oDriver_SaveComplete()
   MsgBox "Transaction Successfully Posted!!!", vbInformation, "Information"
   label.Caption = "POSTED"
   ShowButton
End Sub

Private Sub oDriver_WillSave(Cancel As Boolean)
   If label.Caption = "POSTED" Then
      MsgBox "Transaction Already Posted!!!", vbCritical, "Warning"
      Exit Sub
   End If
   With GridEditor1
      If oDriver.FieldValue(2) = "" Then
         MsgBox "Invalid Destination Detected!!!", vbCritical, "Warning"
         txtfield(2).SetFocus
         Cancel = True
      ElseIf oDriver.FieldValue(3) = "" Then
         MsgBox "Invalid Transaction Date Detected!!!", vbCritical, "Warning"
         txtfield(3).SetFocus
         Cancel = True
      Else
         Time = Format(Now, "hh:nn:ss AM/PM")
         Cancel = Not Save_CP_Inventory_Ledger
            If Cancel Then Exit Sub
         oDriver.FieldValue(4) = CDate(txtfield(4).Text) & " " & Time
         oDriver.FieldValue(8) = 1   'cTranStat
         oDriver.FieldValue(9) = oApp.BranchCode
         oDriver.FieldValue(10) = 1  'cReceived
         oDriver.FieldValue(11) = .Rows - 1
      End If
   End With
End Sub

Private Function Save_CP_Inventory_Ledger() As Boolean
Dim lsSQL As String
Dim lnrow As Long
Dim lnEntry As Integer
Dim QOH As Integer
Dim lrs As New ADODB.Recordset
Dim oRS As New ADODB.Recordset

Save_CP_Inventory_Ledger = True
On Error GoTo errProc
   
   With GridEditor1

      For pnCtr = 1 To .Rows - 1
         Set oRS = New ADODB.Recordset
         lsSQL = "SELECT * " _
               & " FROM CP_Inventory_Ledger " _
               & " WHERE sStockIDx = '" & .TextMatrix(pnCtr, 4) & "'" _
                  & " AND sBranchCd = '" & oApp.BranchCode & "'" _
                  & " AND sSourceCd = 'CPDl'" _
                  & " AND sSourceNo = '" & oDriver.FieldValue(0) & "'"
         If oRS.State = adStateOpen Then oRS.Close
         oRS.Open lsSQL, oApp.Connection, adOpenStatic, adLockOptimistic, adCmdText
         
         If oRS.RecordCount <> 0 Then
            MsgBox "Transaction Already Posted!!!", vbCritical, "Warning"
            Save_CP_Inventory_Ledger = False
            GoTo endProc
         End If

         'Get last Entry No.
         lnEntry = getEntryNo("CP_Inventory_Ledger", "'" & .TextMatrix(pnCtr, 4) & "'", _
                     "'" & oApp.BranchCode & "'")
           
         'Get QOH
         QOH = getQuantity("'" & .TextMatrix(pnCtr, 4) & "'", "'" & oApp.BranchCode & "'") _
                  + .TextMatrix(pnCtr, 5)
         
            'Add Record, CP_Inventory_Ledger
            lsSQL = "INSERT INTO CP_Inventory_Ledger " _
                  & "( sStockIDx, " _
                  & "  sBranchCd, " _
                  & "  sLocation, " _
                  & "  sSourceCd, " _
                  & "  sSourceNo, " _
                  & "  nQtyInxxx, " _
                  & "  nQtyOutxx, " _
                  & "  nQtyOnHnd, " _
                  & "  nEntryNox, " _
                  & "  dTransact, " _
                  & "  dModified) " _
            & "VALUES " _
                  & "('" & .TextMatrix(pnCtr, 4) & "', " _
                  & "'" & oApp.BranchCode & "', " _
                  & "'" & oDriver.FieldValue(2) & "'," _
                  & "'CPDl' , " _
                  & "'" & oDriver.FieldValue(0) & "', " _
                  & "'" & .TextMatrix(pnCtr, 5) & "', " _
                  & "'0', " _
                  & "'" & CLng(QOH) & "', " _
                  & "'" & lnEntry & "', " _
                  & "'" & CDate(oDriver.FieldValue(4)) & " " & Time & "', " _
                  & " getdate())"
            oApp.Connection.Execute lsSQL, lnrow, adCmdText
            
            Set lrs = New ADODB.Recordset
            lsSQL = "SELECT * " _
                  & " FROM CP_Inventory_Master " _
                  & " WHERE sStockIDx = '" & .TextMatrix(pnCtr, 4) & "'" _
                     & " AND sBranchCd = '" & oApp.BranchCode & "'"
            If lrs.State = adStateOpen Then lrs.Close
            lrs.Open lsSQL, oApp.Connection, adOpenStatic, adLockOptimistic, adCmdText

            If lrs.RecordCount <> 0 Then
               'Update QOH, CP_Inventory_Master
               lsSQL = "UPDATE CP_Inventory_Master SET" _
                     & " nQtyOnHnd = '" & CLng(QOH) & "', " _
                     & " sModified = '" & Encrypt(oApp.UserID) & "'," _
                     & " dModified = getdate() " _
               & " WHERE sStockIDx = '" & .TextMatrix(pnCtr, 4) & "'" _
                     & " AND sBranchCd = '" & oApp.BranchCode & "'"
               oApp.Connection.Execute lsSQL, lnrow, adCmdText
            Else
               lsSQL = "INSERT INTO CP_Inventory_Master " _
                  & "( sStockIDx, " _
                  & "  sBranchCd, " _
                  & "  nBegQtyxx, " _
                  & "  nQtyOnHnd, " _
                  & "  nReorderx, " _
                  & "  nMinLevel, " _
                  & "  nMaxLevel, " _
                  & "  dBegInvxx, " _
                  & "  cRecdStat, " _
                  & "  sModified, " _
                  & "  dModified) " _
                      & "VALUES " _
                      & "('" & .TextMatrix(pnCtr, 4) & "', " _
                      & "'" & oApp.BranchCode & "', " _
                      & "'" & 0 & "', " _
                      & "'" & CLng(QOH) & "', " _
                      & "'" & 1 & "', " _
                      & "'" & 1 & "', " _
                      & "'" & 1 & "', " _
                      & "'" & oApp.ServerDate & "', " _
                      & " '" & xeRecStateActive & "', " _
                      & " '" & Encrypt(oApp.UserID) & "', " _
                      & " getdate())"
               oApp.Connection.Execute lsSQL, lnrow, adCmdText
            End If
         
      Next

      If lnrow <= 0 Then
         MsgBox "Unable to Update Inventory Ledger!!!" & vbCrLf & vbCrLf & _
         "Notify ROSALYN LAZO DESCALLAR for Assistance", vbCritical, "Warning"
         Save_CP_Inventory_Ledger = False
         GoTo endProc
      End If
   End With
   Set lrs = Nothing
   Set oRS = Nothing

endProc:
   Exit Function
errProc:
   Save_CP_Inventory_Ledger = False
   MsgBox Err.Description, vbCritical, "Warning"
End Function

Private Sub txtField_GotFocus(Index As Integer)
   oDriver.ColumnIndex = Index
   txtfieldGotfocus = True
   pnindex = Index
   txtfield(Index).BackColor = &HE1FEFF
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Or KeyCode = 13 Then
      If Index = 2 Then
         oDriver.RecordSearch txtfield(Index).Text
         If txtfield(Index).Text <> "" Then SetNextFocus
      End If
      KeyCode = 0
   End If
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   txtfield(Index).BackColor = &HFFFFFF
End Sub

Private Sub txtfield_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
   Case 3, 4
      If Not IsDate(txtfield(Index).Text) Then
         txtfield(Index).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
      Else
         txtfield(Index).Text = Format(txtfield(Index).Text, "MMMM DD, YYYY")
      End If
   End Select
   Cancel = Not oDriver.ValidateField(Index)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyReturn, vbKeyUp, vbKeyDown
      Select Case KeyCode
      Case vbKeyReturn, vbKeyDown
         If GetFocus = GridEditor1.hWnd Then Exit Sub
         SetNextFocus
      Case vbKeyUp
         SetPreviousFocus
      End Select
   End Select
End Sub

'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  Date Tested  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  February 23, 2007  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'

'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤    Version 1    ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  Date Finished  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  February 23, 2007  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'



