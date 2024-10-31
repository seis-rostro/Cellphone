VERSION 5.00
Object = "{0A7B56A6-35D0-4533-91DA-1715D3A0DD3E}#1.1#0"; "xrGridControl.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTransfer_NoSerial 
   BorderStyle     =   0  'None
   Caption         =   "Branch Transfer Transaction  (NO IMIE No.) "
   ClientHeight    =   8010
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11610
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8010
   ScaleWidth      =   11610
   ShowInTaskbar   =   0   'False
   Begin xrGridEditor.GridEditor GridEditor1 
      Height          =   4980
      Left            =   1635
      TabIndex        =   14
      Tag             =   "et0;eb0;et0;bc2"
      Top             =   2850
      Width           =   9810
      _ExtentX        =   17304
      _ExtentY        =   8784
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
      Object.HEIGHT          =   4980
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
      MOUSEICON       =   "frmTransfer_NoSerial.frx":0000
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
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   873
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   2
         Left            =   1335
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   105
         Width           =   4365
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   0
         Left            =   7815
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   105
         Width           =   1980
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction No"
         Height          =   285
         Index           =   9
         Left            =   6570
         TabIndex        =   2
         Top             =   120
         Width           =   1185
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Destination"
         Height          =   285
         Index           =   19
         Left            =   120
         TabIndex        =   0
         Top             =   105
         Width           =   975
      End
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   5130
      Left            =   1575
      Tag             =   "wt0;fb0"
      Top             =   2775
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   9049
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
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   2990
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   5
         Left            =   5250
         MaxLength       =   128
         MultiLine       =   -1  'True
         TabIndex        =   11
         Text            =   "frmTransfer_NoSerial.frx":001C
         Top             =   405
         Width           =   2460
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   3
         Left            =   1335
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   405
         Width           =   2460
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   900
         Index           =   4
         Left            =   1335
         MaxLength       =   300
         MultiLine       =   -1  'True
         TabIndex        =   13
         Text            =   "frmTransfer_NoSerial.frx":0022
         Top             =   660
         Width           =   8445
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   1335
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   150
         Width           =   2460
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   6
         Left            =   5250
         MaxLength       =   128
         MultiLine       =   -1  'True
         TabIndex        =   9
         Text            =   "frmTransfer_NoSerial.frx":0028
         Top             =   150
         Width           =   2460
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
         Left            =   7815
         TabIndex        =   15
         Tag             =   "eb0;wb0"
         Top             =   150
         Width           =   1965
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Requested By "
         Height          =   285
         Index           =   0
         Left            =   4035
         TabIndex        =   10
         Top             =   420
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   285
         Index           =   7
         Left            =   135
         TabIndex        =   12
         Top             =   675
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   285
         Index           =   1
         Left            =   135
         TabIndex        =   6
         Top             =   420
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transmittal No."
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   150
         Width           =   1185
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Approved By "
         Height          =   285
         Index           =   3
         Left            =   4035
         TabIndex        =   8
         Top             =   150
         Width           =   1200
      End
   End
   Begin MSComCtl2.Animation Progress 
      Height          =   720
      Left            =   300
      TabIndex        =   23
      TabStop         =   0   'False
      Tag             =   "wt0;wb0"
      Top             =   720
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   1270
      _Version        =   393216
      BackColor       =   1842204
      FullWidth       =   61
      FullHeight      =   48
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   3
      Left            =   90
      TabIndex        =   21
      Top             =   6000
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
      Caption         =   "&Cancel"
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
      Picture         =   "frmTransfer_NoSerial.frx":002E
      CaptionAlign    =   0
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   0
      Left            =   90
      TabIndex        =   18
      Top             =   4740
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
      Caption         =   "&Save"
      AccessKey       =   "S"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmTransfer_NoSerial.frx":07A8
      CaptionAlign    =   0
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   4
      Left            =   90
      TabIndex        =   19
      Top             =   5580
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
      Picture         =   "frmTransfer_NoSerial.frx":0F22
      CaptionAlign    =   0
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   5
      Left            =   90
      TabIndex        =   22
      Top             =   6000
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
      Picture         =   "frmTransfer_NoSerial.frx":169C
      CaptionAlign    =   0
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   405
      Index           =   1
      Left            =   90
      TabIndex        =   16
      Top             =   4320
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
      Caption         =   "Searc&h"
      AccessKey       =   "h"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmTransfer_NoSerial.frx":1E16
      CaptionAlign    =   0
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   2
      Left            =   90
      TabIndex        =   20
      Top             =   5580
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
      Caption         =   "&Delete"
      AccessKey       =   "D"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmTransfer_NoSerial.frx":2590
      CaptionAlign    =   0
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   6
      Left            =   90
      TabIndex        =   17
      Top             =   5160
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
      Caption         =   "Pre&view"
      AccessKey       =   "v"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmTransfer_NoSerial.frx":2D0A
      CaptionAlign    =   0
   End
End
Attribute VB_Name = "frmTransfer_NoSerial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  Programmed By  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  Rosalyn Lazo Descallar  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  Date Started  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  February 19, 2007  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'

Option Explicit

Private WithEvents oDriver As FormDriver
Attribute oDriver.VB_VarHelpID = -1
Private oSkin As FormSkin
Private bLoaded As Boolean
Private oRS As New ADODB.Recordset
Private poFileSys As FileSystemObject

Dim txtfieldGotfocus As Boolean
Dim pbnewitem As Boolean
Dim psSelected() As String
Dim lrsTarget As New ADODB.Recordset
Dim Drive As String
Dim temp As String
Dim Reference As String

Dim pnindex As Integer
Dim pnCtr As Integer

Dim Time As String
Dim Branch As String
Dim Code As String
Dim Address As String

Private Sub cmdButton_Click(Index As Integer)
Dim lsSearch As String
Dim lsCancel As Integer
Dim lsSQL As String
Dim lnrow As Long

   Select Case Index
      Case 0   'save
         oDriver.RecordSave
      Case 1   'search
         If txtfieldGotfocus Then
            If pnindex = 2 Then oDriver.RecordSearch txtField(pnindex).Text
         End If
      Case 2   'delete
         With GridEditor1
            If .Rows <> 2 Then
               .DeleteRow
            End If
         End With
      Case 3   'cancel
         If pbnewitem = True Then
            If GridEditor1.TextMatrix(1, 6) <> "" Then
               lsCancel = MsgBox("Are you sure you want to Cancel this Transaction?" & vbCrLf & _
               "This Entry will be Erased!!!", vbQuestion + vbYesNo, "Confirm")
               If lsCancel <> vbYes Then Exit Sub
            End If
         End If
         InitButton xeModeReady
         
         EmptyGrid
      Case 4   'new
         oDriver.RecordNew
         InitButton xeModeAddNew
         EmptyGrid
      Case 5   'close
         Unload Me
      Case 6   'Print
         Print_Transaction
   End Select
End Sub

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0
   
   If bLoaded = False Then
      oDriver.RecordNew
      oDriver.DisableTextbox 0
      bLoaded = True
      oDriver.ShowButton 2
   End If
   GridEditor1.Refresh
End Sub

Private Sub Form_Deactivate()
   Progress.Stop
   Progress.Close
End Sub

Private Sub GridEditor1_AddingRow(Cancel As Boolean)
   With GridEditor1
      If .TextMatrix(.Row, 1) = "" Then
         MsgBox "Invalid Bar Code!!!", vbCritical, "Warning"
         Cancel = True
      ElseIf .TextMatrix(.Row, 2) = "" Then
         .Col = 1
         Cancel = True
      ElseIf .TextMatrix(.Row, 5) = "0" Or .TextMatrix(.Row, 5) = "" Then
         MsgBox "Invalid Quantity!!!", vbCritical, "Warning"
         Cancel = True
      ElseIf CLng(.TextMatrix(.Row, 5)) > CLng(.TextMatrix(.Row, 7)) Then
         Cancel = True
      End If
   End With
End Sub

Private Sub GridEditor1_EditorValidate(Cancel As Boolean)
   With GridEditor1
      If .Col = 5 Then
         If .TextMatrix(.Row, 5) <> "" And .TextMatrix(.Row, 7) <> "" Then
            If CLng(.TextMatrix(.Row, 5)) > CLng(.TextMatrix(.Row, 7)) Then
               MsgBox "Invalid Quantity!!!" & vbCrLf & _
                      "Quantity Greater than Quantity On Hand!!!", vbCritical, "Warning"
               .Col = 5
            ElseIf .TextMatrix(.Row, 5) = 0# Then
               MsgBox "Invalid Quantity!!!", vbCritical, "Warning"
               .Col = 5
            End If
         End If
      End If
   End With
End Sub

Private Sub GridEditor1_RowColChange()
With GridEditor1
   If .TextMatrix(.Row, 1) <> "" And .TextMatrix(.Row, 4) = "" Then
      .Col = 1
   End If
End With
End Sub

Private Sub oDriver_DisableOtherControl()
   oDriver.DisableTextbox 0
End Sub

Private Sub oDriver_EnableOtherControl()
   oDriver.DisableTextbox 0
End Sub

Private Sub oDriver_InitValue()
   oDriver.FieldReference(0) = True
   If Not oDriver.SetValue(0, getNextCode("CP_Transfer_Master", "sTransNox", True, oApp.Connection, True, oApp.BranchCode)) Then Exit Sub
   
   pbnewitem = True
   label.Caption = "UNKNOWN"
   oDriver.FieldValue(1) = Reference_No
   txtField(1).Text = oDriver.FieldValue(1)
   txtField(3).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
   oDriver.FieldValue(3) = Date

End Sub

Function Reference_No() As String
   Dim lrs As Recordset
   Dim lsSQL As String
   Dim lnCtr As Long
   
   lsSQL = "SELECT TOP 1" & _
            " sReferNox" & _
            " FROM CP_Transfer_Master " & _
            " WHERE sReferNox LIKE " & strParm(oApp.BranchCode & "-%") & _
            " ORDER BY sReferNox DESC"
   
   Set lrs = New Recordset
   lrs.Open lsSQL, oApp.Connection, , , adCmdText
   
   If lrs.EOF Then
      lnCtr = 1
   Else
      If Left(lrs("sReferNox"), 2) = oApp.BranchCode Then
         lnCtr = CLng(Right(lrs("sReferNox"), 5)) + 1
      Else
         lnCtr = 1
      End If
   End If
   
   Reference_No = oApp.BranchCode & "-" & Format(Date, "yy") & Format(lnCtr, "00000")
   Set lrs = Nothing
End Function


Private Sub Form_Load()
Dim lsSQL As String

   CenterChildForm mdiMain, Me
   bLoaded = False
    
   Set oDriver = New FormDriver
   Set oDriver.AppDriver = oApp
   Set oDriver.MainForm = Me
   
   InitButton xeModeAddNew
   InitGrid
   
   Set oSkin = New FormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransaction
    
   oDriver.RecQuery = "SELECT" _
                        & " sTransNox, " _
                        & " sReferNox, " _
                        & " sDestinat, " _
                        & " dTransact, " _
                        & " sRemarksx, " _
                        & " sRequestx, " _
                        & " sApproved, " _
                        & " sOriginxx, " _
                        & " nEntryNox, " _
                        & " cTranStat, " _
                        & " cReceived, " _
                        & " sModified, " _
                        & " dModified, " _
                        & " vTimeStmp  " _
                  & " FROM CP_Transfer_Master " _

   oDriver.BrowseQuery = "SELECT" _
                  & " a.sTransNox, " _
                  & " a.sReferNox, " _
                  & " a.dTransact, " _
                  & " b.sBranchNm  " _
            & " FROM CP_Transfer_Master a " _
               & " LEFT JOIN Branch b " _
                  & " ON a.sDestinat = b.sBranchCd " _
            & " WHERE cTranStat = 0 " _
               & " AND sOriginxx = '" & oApp.BranchCode & "' " _
            & " ORDER BY dTransact Desc "
   
   oDriver.InitRecForm

   oDriver.BrowseFTitle(0) = "Transaction No"
   oDriver.BrowseFTitle(1) = "Transmittal No"
   oDriver.BrowseFTitle(2) = "Date"
   oDriver.BrowseFTitle(3) = "Destination"
   
   oDriver.BrowseFFormat(2) = "MMMM dd, yyyy"


   'Destination
   oDriver.LookupQuery(2) = "SELECT" _
                           & " a.sBranchCd, " _
                           & " a.sBranchNm, " _
                           & " a.sAddressx + ' ' + b.sTownName xAddressx " _
                     & " FROM Branch a " _
                        & "LEFT JOIN TownCity b " _
                           & " ON a.sTownIdxx = b.sTownIDxx " _
                     & " WHERE a.cRecdStat = '" & xeRecStateActive & "'" _
                           & " AND a.sBranchCd <> '" & oApp.BranchCode & "'" _
                     & " ORDER BY sBranchNm "

   oDriver.LookupReference(2) = "sBranchCd»sBranchNm»xAddressx"
   oDriver.LookupColumn(2) = "sBranchNm»xAddressx"
   oDriver.LookupTitle(2) = "Branch»Address"
       
   oDriver.FieldFormat(0) = "@@-@@@@@@@@"
   oDriver.FieldSize(0) = Len(oDriver.FieldFormat(0))

   oDriver.FieldStart = 2
   oDriver.FieldFormat(3) = "MMMM DD, YYYY"
   EmptyGrid

End Sub

Private Sub InitGrid()
    With GridEditor1
      .Rows = 2
      .Cols = 8
      .Font = "MS Sans Serif"
              
      'column title
      .TextMatrix(0, 1) = "Bar Code"
      .TextMatrix(0, 2) = "Particulars"
      .TextMatrix(0, 3) = "SRP"
      .TextMatrix(0, 4) = "Stock ID"
      .TextMatrix(0, 5) = "Qty"
      .TextMatrix(0, 6) = "Pur.Price"
      .TextMatrix(0, 7) = "QOH"
      .Row = 0
      
      'column width
      .ColWidth(0) = 400
      .ColWidth(1) = 2200
      .ColWidth(2) = 4750
      .ColWidth(3) = 0
      .ColWidth(4) = 0
      .ColWidth(5) = 670
      .ColWidth(6) = 1000
      .ColWidth(7) = 670

      .ColFormat(1) = ">"
      .ColFormat(6) = "#,##0.00"
      
      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      .ColAlignment(4) = 1
      .ColAlignment(5) = 6
      .ColAlignment(6) = 6
      .ColAlignment(7) = 6
      
      .ColDefault(5) = 0
      .ColNumberOnly(5) = True
      .ColNumberOnly(6) = True
      
      .ColEnabled(2) = False
      .ColEnabled(3) = False
      .ColEnabled(4) = False
      .ColEnabled(7) = False
      
      .Row = 1
    End With
End Sub

Private Sub EmptyGrid()
   With GridEditor1
      .Rows = 2
      For pnCtr = 1 To .Cols - 1
         .TextMatrix(1, pnCtr) = ""
      Next
      .ColEnabled(1) = True
      .ColEnabled(5) = True
      .ColEnabled(6) = True
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oDriver = Nothing
   Set oSkin = Nothing
End Sub

Private Sub GridEditor1_GotFocus()
   GridEditor1.Col = 1
End Sub

Private Sub GridEditor1_KeyDown(KeyCode As Integer, Shift As Integer)
   With GridEditor1
      If KeyCode = vbKeyF3 Then
         If .Col = 1 Then SearchBarCode
      End If
   End With
End Sub

Private Sub Search_Transmittal()
Dim orig As String
Dim lsSQL As String
Dim lsCondition As String

   orig = oDriver.BrowseQuery
   Select Case pnindex
      Case 1
         lsCondition = " a.sReferNox like '%" & txtField(1).Text & "'"
         lsSQL = AddCondition(oDriver.BrowseQuery, lsCondition)
         oDriver.BrowseQuery = lsSQL
      Case 2
         lsCondition = " b.sBranchNm like '" & txtField(2).Text & "%'"
         lsSQL = AddCondition(oDriver.BrowseQuery, lsCondition)
         oDriver.BrowseQuery = lsSQL
   End Select
   If pnindex = 1 Then
   ElseIf pnindex = 2 Then
   End If
   oDriver.BrowseRecord
   oDriver.BrowseQuery = orig

End Sub

Private Sub Print_Transaction()
Dim lnCtr As Integer
Dim lrs As New ADODB.Recordset
Dim lsSQL As String
Dim lrsDetail As New ADODB.Recordset

   Set lrs = New ADODB.Recordset
   lrs.Fields.Append "sField01", adVarChar, 150
   lrs.Fields.Append "sField02", adVarChar, 20
   lrs.Fields.Append "nField01", adInteger, 5
   lrs.Fields.Append "nField02", adInteger, 5
   lrs.Open

   'CP_Transfer_Master
   lsSQL = "SELECT" _
               & " a.sTransNox, " _
               & " a.sReferNox, " _
               & " a.sApproved, " _
               & " a.sDestinat, " _
               & " b.sBranchNm, " _
               & " c.sTownIDxx, " _
               & " b.sAddressx + ', ' + c.sTownName xAddressx " _
         & " FROM CP_Transfer_Master a " _
            & " LEFT JOIN Branch b " _
               & " ON a.sOriginxx = b.sBranchCd " _
            & " LEFT JOIN TownCity c " _
               & " ON b.sTownIDxx = c.sTownIDxx " _
         & " WHERE a.sTransNox = '" & oDriver.FieldValue(0) & "' " _

   If oRS.State = adStateOpen Then oRS.Close
   oRS.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText

   If oRS.EOF Then
      MsgBox "No Record Found!!!" & vbCrLf & _
             "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
      Exit Sub
   End If
      Progress.Open App.Path & "\images\FINDCOMP.AVI"
      Progress.Play
   
      For lnCtr = 0 To oRS.RecordCount - 1
         lrs.AddNew
         
            lsSQL = "SELECT" _
                     & " a.nEntryNox, " _
                     & " a.nQuantity, " _
                     & " b.sDescript, " _
                     & " c.sBrandNme, " _
                     & " d.sModelNme, " _
                     & " e.sColorNme, " _
                     & " b.sBarrcode  " _
                  & " FROM CP_Transfer_Detail a " _
                     & " LEFT JOIN CP_Inventory b " _
                        & " ON a.sStockIDx = b.sStockIDx " _
                     & " LEFT JOIN Brand c " _
                        & " ON b.sBrandIdx = c.sBrandIDx " _
                     & " LEFT JOIN Model d  " _
                        & " ON b.sModelIDx = d.sModelIDx " _
                     & " LEFT JOIN Color e  " _
                        & " ON b.sColorIDx = e.sColorIDx " _
                  & " WHERE a.sTransNox = '" & oRS("sTransNox") & "' " _
                  & " ORDER BY a.nEntryNox "
            
            If lrsDetail.State = adStateOpen Then lrsDetail.Close
            lrsDetail.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText

            Do While Not lrsDetail.EOF
               lrs.AddNew
               lrs("sField01").Value = Trim(IIf(IsNull(lrsDetail("sBrandNme")), "", lrsDetail("sBrandNme")) _
                                       & " " & IIf(IsNull(lrsDetail("sModelNme")), "", lrsDetail("sModelNme")) _
                                       & " " & IIf(IsNull(lrsDetail("sDescript")), "", lrsDetail("sDescript")) _
                                       & " " & IIf(IsNull(lrsDetail("sColorNme")), "", lrsDetail("sColorNme")))
               lrs("sField02").Value = lrsDetail("sBarrcode")
               lrs("nField01").Value = lrsDetail("nEntryNox")
               lrs("nField02").Value = lrsDetail("nQuantity")
               lrsDetail.MoveNext
            Loop
            
         oRS.MoveNext
      Next
      
      Branch = txtField(2)
      getBranch Code, Branch, Address
         
      Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\CP_Transmittal_NoSerial.rpt")
      oReport.DiscardSavedData
      oReport.FieldMappingType = crAutoFieldMapping
      oReport.Database.SetDataSource lrs
      
      oRS.MoveFirst
      
      With oReport
         .Sections("PH").ReportObjects("txtReportDate").SetText txtField(3).Text
         .Sections("PH").ReportObjects("txtTransmittal").SetText txtField(1).Text
         
         .Sections("PH").ReportObjects("txtToBranch").SetText Branch
         .Sections("PH").ReportObjects("txtToAddress").SetText Address
         .Sections("PH").ReportObjects("txtFromBranch").SetText oRS("sBranchNm")
         .Sections("PH").ReportObjects("txtFromAddress").SetText oRS("xAddressx")
      
         .Sections("PF").ReportObjects("txtApproved").SetText txtField(6).Text
         .Sections("PF").ReportObjects("txtPrepared").SetText oApp.UserName
         .Sections("RF").ReportObjects("txtRemarks").SetText txtField(4).Text
      End With
      
      Set lrs = Nothing
      Set oRS = Nothing
      Set lrsDetail = Nothing
      
      frmViewer.CRViewer91.ReportSource = oReport
      frmViewer.CRViewer91.ViewReport
      frmViewer.Show

End Sub

Private Sub SearchBarCode()
Dim lsSQL As String
Dim lsCondition As String
Dim lsSearch As String
Dim lnCtr As Integer
   
   With GridEditor1
      lsSQL = "SELECT" _
            & " a.sBarrcode, " _
            & " a.sStockIDx, " _
            & " b.sBrandNme, " _
            & " c.sModelNme, " _
            & " a.sDescript, " _
            & " d.sColorNme, " _
            & " a.nSelPrice, " _
            & " e.nQtyOnHnd, " _
            & " a.cWdSerial, " _
            & " a.nPurPrice  " _
         & " FROM CP_Inventory a " _
            & " LEFT JOIN CP_Inventory_Master e " _
               & " ON a.sStockIdx = e.sStockIDx " _
            & " LEFT JOIN Brand b " _
               & " ON a.sBrandIdx = b.sBrandIdx " _
            & " LEFT JOIN Model c " _
               & " ON a.sModelIdx = c.sModelIdx " _
            & " LEFT JOIN Color d " _
               & " ON a.sColorIDx = d.sColorIDx " _
         & " WHERE a.sBarrcode like  '%" & .TextMatrix(.Row, 1) & "%' " _
            & " AND a.cWdSerial = 0 and a.cWalletxx = 0 and a.cCellLoad = 0  " _
            & " AND e.sBranchCd = '" & oApp.BranchCode & "'"
      If oRS.State = adStateOpen Then oRS.Close
      oRS.Open lsSQL, oApp.Connection, adOpenKeyset, adLockOptimistic, adCmdText

      If Not oRS.EOF Then
         If oRS.RecordCount = 1 Then
            .TextMatrix(.Row, 1) = IIf(IsNull(oRS(0)), "", oRS(0))
            .TextMatrix(.Row, 2) = Trim(IIf(IsNull(oRS(2)), "", oRS(2)) _
                                 & " " & IIf(IsNull(oRS(3)), "", oRS(3)) _
                                 & " " & IIf(IsNull(oRS(4)), "", oRS(4)) _
                                 & " " & IIf(IsNull(oRS(5)), "", oRS(5)))
            .TextMatrix(.Row, 3) = IIf(IsNull(oRS(6)), 0, Format(oRS(6), "#,##0.00"))
            .TextMatrix(.Row, 4) = IIf(IsNull(oRS(1)), "", oRS(1))
            .TextMatrix(.Row, 6) = IIf(IsNull(oRS(9)), 0, Format(oRS(9), "#,##0.00"))
            .TextMatrix(.Row, 7) = IIf(IsNull(oRS(7)), 0, oRS(7))
            .Col = 5
         Else
            lsSearch = KwikSearch(oApp, lsSQL, _
                       "sBarrcode»sBrandNme»sModelNme»sDescript»sColorNme", _
                       "Bar Code»Brand»Model»Description»Color")
            If lsSearch <> "" Then
               psSelected = Split(lsSearch, "»")
               oDriver.LookupValue(0) = psSelected(0)
               .TextMatrix(.Row, 1) = IIf(IsNull(psSelected(0)), "", psSelected(0))
               .TextMatrix(.Row, 2) = Trim(IIf(IsNull(psSelected(2)), "", psSelected(2)) _
                                    & " " & IIf(IsNull(psSelected(3)), "", psSelected(3)) _
                                    & " " & IIf(IsNull(psSelected(4)), "", psSelected(4)) _
                                    & " " & IIf(IsNull(psSelected(5)), "", psSelected(5)))
               .TextMatrix(.Row, 3) = IIf(IsNull(psSelected(6)), 0, Format(psSelected(6), "#,##0.00"))
               .TextMatrix(.Row, 4) = IIf(IsNull(psSelected(1)), "", psSelected(1))
               .TextMatrix(.Row, 6) = IIf(IsNull(psSelected(9)), 0, Format(psSelected(9), "#,##0.00"))
               .TextMatrix(.Row, 7) = IIf(IsNull(psSelected(7)), 0, psSelected(7))
               .Col = 5
            End If
         End If
         .SetFocus
         .Refresh
      Else
         For pnCtr = 1 To .Cols
            .TextMatrix(.Row, pnCtr) = ""
         Next
         .Col = 1
         MsgBox "Bar Code Not Existing!!!", vbInformation, "Information"
      End If
         
      For lnCtr = 1 To .Rows - 1
         If lnCtr <> .Row And .TextMatrix(.Row, 3) <> "" Then
            If .TextMatrix(lnCtr, 3) = .TextMatrix(.Row, 3) Then
               MsgBox "Duplicate Entry!!!" & vbCrLf & vbCrLf & _
               "Update Quantity of Row" & " " & lnCtr, vbCritical, "Warning"
               For pnCtr = 1 To .Cols
                  .TextMatrix(.Row, pnCtr) = ""
               Next
               .SetFocus
            End If
         End If
      Next
      .Refresh
      Set oRS = Nothing
   End With
   
End Sub

Private Sub InitButton(lnStat As Integer)
   Dim lbShow As Boolean

   lbShow = IIf(lnStat = 0, False, True)
   cmdButton(0).Visible = lbShow
   cmdButton(1).Visible = lbShow
   cmdButton(2).Visible = lbShow
   cmdButton(3).Visible = lbShow
   
   cmdButton(4).Visible = Not lbShow
   cmdButton(5).Visible = Not lbShow
   
   With GridEditor1
      .ColEnabled(1) = lbShow
   End With

   xrFrame1(0).Enabled = lbShow
   If Not lbShow Then cmdButton(4).SetFocus
   
End Sub


Private Sub oDriver_LoadOtherData()
   oDriver.FieldValue(3) = Format(oDriver.FieldValue(3), "m/d/yyyy")
   Select Case oDriver.FieldValue(9)
      Case 0
         label.Caption = "UNKNOWN"
      Case 1
         label.Caption = "POSTED"
   End Select
End Sub

Private Sub oDriver_Save(Saved As Boolean)
   Saved = False
End Sub

Private Sub oDriver_SaveComplete()
Dim orig As String
Dim lsSQL As String
Dim lsCondition As String
   orig = oDriver.BrowseQuery
   lsCondition = " a.sTransNox = '" & Reference & "' "
   lsSQL = AddCondition(oDriver.BrowseQuery, lsCondition)
   oDriver.BrowseQuery = lsSQL
   oDriver.BrowseRecord
   oDriver.BrowseQuery = orig
   InitButton xeModeReady
End Sub

Private Sub oDriver_WillSave(Cancel As Boolean)
   With GridEditor1
      If txtField(1).Text = "" Then
         MsgBox "Invalid Transmittal No. Detected!!!", vbCritical, "Warning"
         txtField(1).SetFocus
         Cancel = True
      ElseIf oDriver.FieldValue(2) = "" Then
         MsgBox "Invalid Destination Detected!!!", vbCritical, "Warning"
         txtField(2).SetFocus
         Cancel = True
      ElseIf oDriver.FieldValue(3) = "" Then
         MsgBox "Invalid Transaction Date Detected!!!", vbCritical, "Warning"
         txtField(3).SetFocus
         Cancel = True
      Else
         Time = Format(Now, "hh:nn:ss AM/PM")
         If pbnewitem Then
            Cancel = Not Save_CP_Transfer_Detail
               If Cancel Then Exit Sub
            Cancel = Not Update_CP_Inventory
               If Cancel Then Exit Sub
         End If
         oDriver.FieldValue(3) = CDate(txtField(3).Text) & " " & Time
         oDriver.FieldValue(7) = oApp.BranchCode
         oDriver.FieldValue(8) = .TextMatrix(.Rows - 1, 0)
         oDriver.FieldValue(9) = 0  'cTranStat
         oDriver.FieldValue(10) = 0  'cReceived
         Reference = oDriver.FieldValue(0)
         Branch = txtField(2)
      End If
   End With
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   oDriver.ColumnIndex = Index
   txtfieldGotfocus = True
   pnindex = Index
   txtField(Index).BackColor = &HE1FEFF
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Or KeyCode = 13 Then
      If Index = 2 Then
         oDriver.RecordSearch txtField(Index).Text
         If txtField(Index).Text <> "" Then SetNextFocus
      End If
      KeyCode = 0
   End If
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   If Index = 3 Then
      If Not IsDate(txtField(Index).Text) Then
         txtField(Index).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
      Else
         txtField(Index).Text = Format(txtField(Index).Text, "MMMM DD, YYYY")
      End If
   End If
   txtField(Index).BackColor = &HFFFFFF
End Sub

Private Sub txtfield_Validate(Index As Integer, Cancel As Boolean)
   If Index = 3 Then
      If Not IsDate(txtField(Index).Text) Then
         txtField(Index).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
      Else
         txtField(Index).Text = Format(txtField(Index).Text, "MMMM DD, YYYY")
      End If
   End If
   txtField(Index).Text = TitleCase(txtField(Index).Text)
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
      Case 27
         Call Modified("CP_Transfer_Master", "sTransNox = '" & oDriver.FieldValue(0) & "' ")
   End Select
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
               & " f.sColorNme, " _
               & " e.nQtyOnHnd, " _
               & " b.nSelPrice  "
      lsSQL = lsSQL _
         & " FROM CP_Transfer_Detail a " _
               & " LEFT JOIN CP_Inventory b " _
                  & " ON a.sStockIDx = b.sStockIDx " _
               & " LEFT JOIN CP_Inventory_Master e " _
                  & " ON a.sStockIDx = e.sStockIDx " _
               & " LEFT JOIN Brand c " _
                  & " ON b.sBrandIDx = c.sBrandIDx " _
               & " LEFT JOIN Model d " _
                  & " ON b.sModelIDx = d.sModelIDx " _
               & " LEFT JOIN Color f " _
                  & " ON b.sColorIDx = f.sColorIDx " _
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
            .TextMatrix(lnCtr + 1, 6) = Format(oRS("nUnitPrce"), "#,##0.00")
            .TextMatrix(lnCtr + 1, 7) = oRS("nQtyOnHnd")
            oRS.MoveNext
         Next
         .ColEnabled(1) = False
         .ColEnabled(5) = False
         If .Rows > 20 Then
            .ColWidth(2) = 4500
         Else
            .ColWidth(2) = 4750
         End If
      End With
      oDriver.HideButton 2
      oDriver.ShowButton 7
      oDriver.ShowButton 8
   Else
      Exit Sub
   End If

   Set oRS = Nothing

End Sub

Private Function Save_CP_Transfer_Detail() As Boolean
Dim lsSQL As String
Dim lnrow As Long
   
Save_CP_Transfer_Detail = True
On Error GoTo errProc
   
   With GridEditor1
      'Insert Record
      For pnCtr = 1 To .Rows - 1
         If .TextMatrix(pnCtr, 1) = "" Then Exit For
         lsSQL = "INSERT INTO CP_Transfer_Detail " _
               & "( sTransNox, " _
               & "  nEntryNox, " _
               & "  sStockIDx, " _
               & "  nQuantity, " _
               & "  nUnitPrce, " _
               & "  dModified) " _
         & "VALUES " _
               & "('" & oDriver.FieldValue(0) & "', " _
               & "'" & .TextMatrix(pnCtr, 0) & "', " _
               & "'" & .TextMatrix(pnCtr, 4) & "', " _
               & "'" & CLng(.TextMatrix(pnCtr, 5)) & "', " _
               & "'" & CDbl(.TextMatrix(pnCtr, 6)) & "', " _
               & " getdate())"
         oApp.Connection.Execute lsSQL, lnrow, adCmdText
      Next
      
      If lnrow <= 0 Then
         MsgBox "Unable to Save Transfer Detail!!!" & vbCrLf & vbCrLf & _
         "Notify ROSALYN LAZO DESCALLAR for Assistance", vbCritical, "Warning"
         Save_CP_Transfer_Detail = False
         GoTo endProc
      End If
      
   End With

endProc:
   Exit Function
errProc:
   Save_CP_Transfer_Detail = False
   MsgBox Err.Description, vbCritical, "Warning"
End Function

Private Function Update_CP_Inventory() As Boolean
Dim lsSQL As String
Dim lnrow As Long
Dim lnEntry As Integer
Dim QOH As Integer

Update_CP_Inventory = True
On Error GoTo errProc
   
   With GridEditor1
         For pnCtr = 1 To .Rows - 1
            If .TextMatrix(pnCtr, 1) = "" Then Exit For
            
            'Get last Entry No.
            lnEntry = getEntryNo("CP_Inventory_Ledger", "'" & .TextMatrix(pnCtr, 4) & "'", "'" & oApp.BranchCode & "'")
     
            'Get QOH
            QOH = getQuantity("'" & .TextMatrix(pnCtr, 4) & "'", "'" & oApp.BranchCode & "'") _
                     - .TextMatrix(pnCtr, 5)
            
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
                     & "'" & oApp.BranchCode & "'," _
                     & "'" & oDriver.FieldValue(2) & "', " _
                     & "'CPDv' , " _
                     & "'" & oDriver.FieldValue(0) & "', " _
                     & " 0, " _
                     & "'" & CLng(.TextMatrix(pnCtr, 5)) & "', " _
                     & "'" & CLng(QOH) & "', " _
                     & "'" & lnEntry & "', " _
                     & "'" & CDate(oDriver.FieldValue(3)) & " " & Time & "', " _
                     & " getdate())"
               oApp.Connection.Execute lsSQL, lnrow, adCmdText
            
               'Update QOH, CP_Inventory_Master
               lsSQL = "UPDATE CP_Inventory_Master SET" _
                     & " nQtyOnHnd = '" & CLng(QOH) & "', " _
                     & " sModified = '" & Encrypt(oApp.UserID) & "', " _
                     & " dModified = getdate() " _
               & " WHERE sStockIDx = '" & .TextMatrix(pnCtr, 4) & "' " _
                     & " And sBranchCd = '" & oApp.BranchCode & "' "
               oApp.Connection.Execute lsSQL, lnrow, adCmdText
         
               'Update QOH, CP_Inventory
               lsSQL = "UPDATE CP_Inventory SET" _
                     & " sModified = '" & Encrypt(oApp.UserID) & "', " _
                     & " dModified = getdate() " _
               & " WHERE sStockIDx = '" & .TextMatrix(pnCtr, 4) & "' "
               oApp.Connection.Execute lsSQL, lnrow, adCmdText

         Next
   
         If lnrow <= 0 Then
            MsgBox "Unable to Update Inventory Ledger!!!" & vbCrLf & vbCrLf & _
            "Notify ROSALYN LAZO DESCALLAR for Assistance", vbCritical, "Warning"
            Update_CP_Inventory = False
            GoTo endProc
         End If
         
   End With

endProc:
   Exit Function
errProc:
   Update_CP_Inventory = False
   MsgBox Err.Description, vbCritical, "Warning"
End Function

'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  Date Tested  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  February 19, 2007  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'

'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤    Version 1    ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  Date Finished  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  February 19, 2007  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'

'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤    Add Export   ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  Date Finished  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  February 21, 2007  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'

'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  March 24, 2008  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'Include cWdSerial in Table CP_Inventory



