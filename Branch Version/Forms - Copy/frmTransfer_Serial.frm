VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTransfer_Serial 
   BorderStyle     =   0  'None
   Caption         =   "Branch Transfer Transaction  (w/ IMIE No.) "
   ClientHeight    =   8010
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11610
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8010
   ScaleWidth      =   11610
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox GridEditor1 
      BackColor       =   &H80000005&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   4980
      Left            =   1650
      ScaleHeight     =   4920
      ScaleWidth      =   9735
      TabIndex        =   14
      Tag             =   "et0;eb0;et0;bc2"
      Top             =   2850
      Width           =   9795
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
         Index           =   0
         Left            =   7815
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   105
         Width           =   1980
      End
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
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction No"
         Height          =   285
         Index           =   9
         Left            =   6585
         TabIndex        =   2
         Top             =   120
         Width           =   1185
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   3
      Left            =   90
      TabIndex        =   20
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
      Picture         =   "frmTransfer_Serial.frx":0000
      CaptionAlign    =   0
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   0
      Left            =   90
      TabIndex        =   16
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
      Picture         =   "frmTransfer_Serial.frx":077A
      CaptionAlign    =   0
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   5
      Left            =   90
      TabIndex        =   19
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
      Picture         =   "frmTransfer_Serial.frx":0EF4
      CaptionAlign    =   0
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   405
      Index           =   1
      Left            =   90
      TabIndex        =   15
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
      Picture         =   "frmTransfer_Serial.frx":166E
      CaptionAlign    =   0
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
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   2
      Left            =   90
      TabIndex        =   17
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
      Picture         =   "frmTransfer_Serial.frx":1DE8
      CaptionAlign    =   0
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
         Index           =   6
         Left            =   5250
         MaxLength       =   128
         MultiLine       =   -1  'True
         TabIndex        =   9
         Text            =   "frmTransfer_Serial.frx":2562
         Top             =   150
         Width           =   2460
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
         Height          =   900
         Index           =   4
         Left            =   1335
         MaxLength       =   300
         MultiLine       =   -1  'True
         TabIndex        =   13
         Text            =   "frmTransfer_Serial.frx":2568
         Top             =   660
         Width           =   8445
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
         Height          =   240
         Index           =   5
         Left            =   5250
         MaxLength       =   128
         MultiLine       =   -1  'True
         TabIndex        =   11
         Text            =   "frmTransfer_Serial.frx":256E
         Top             =   405
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
         TabIndex        =   22
         Tag             =   "eb0;wb0"
         Top             =   150
         Width           =   1965
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
         Caption         =   "Date"
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   420
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   285
         Index           =   7
         Left            =   120
         TabIndex        =   12
         Top             =   675
         Width           =   1200
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
   End
   Begin MSComCtl2.Animation Progress 
      Height          =   720
      Left            =   315
      TabIndex        =   21
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
      Index           =   4
      Left            =   90
      TabIndex        =   18
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
      Picture         =   "frmTransfer_Serial.frx":2574
      CaptionAlign    =   0
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   6
      Left            =   90
      TabIndex        =   23
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
      Picture         =   "frmTransfer_Serial.frx":2CEE
      CaptionAlign    =   0
   End
End
Attribute VB_Name = "frmTransfer_Serial"
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
Dim oSerial As frmCP_Serial_Transfer

Dim txtfieldGotfocus As Boolean
Dim pbnewitem As Boolean
Dim psSelected() As String
Dim lrsTarget As New ADODB.Recordset
Dim Drive As String
Dim Reference As String
Dim pbExisting As Boolean

Dim Branch As String
Dim Code As String
Dim Address As String

Dim pnIndex As Integer
Dim pnCtr As Integer
Dim Time As String

Private Sub cmdButton_Click(Index As Integer)
Dim lsSearch As String
Dim lsCancel As Integer
Dim lsSQL As String
Dim lnRow As Long


   Select Case Index
      Case 0   'save
         oDriver.RecordSave
      Case 1   'search
         If txtfieldGotfocus Then
            If pnIndex = 2 Then oDriver.RecordSearch txtField(pnIndex).Text
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
         Cancel = True
      ElseIf .TextMatrix(.Row, 5) = "" Then
         Cancel = True
      ElseIf pbExisting = True Then
         Cancel = True
      End If
   End With
End Sub

Private Sub GridEditor1_EditorValidate(Cancel As Boolean)
   With GridEditor1
      If .TextMatrix(.Row, 1) = "" Then
         MsgBox "Invalid IMEI No.!!!", vbCritical, "Warning"
         .SetFocus
      End If
   End With
End Sub

Private Sub GridEditor1_RowColChange()
   With GridEditor1
      If .Row <> 0 And Trim(.TextMatrix(.Row, 1)) <> "" And .ColEnabled(1) = True Then
         Search_Serial
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
   oDriver.FieldValue(0) = Transaction_No
   txtField(0).Text = oDriver.FieldValue(0)
   oDriver.FieldValue(1) = Reference_No
   txtField(1).Text = oDriver.FieldValue(1)
   
   Label.Caption = "UNKNOWN"
   txtField(3).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
   oDriver.FieldValue(3) = Date
   pbnewitem = True
   pbExisting = False
   
End Sub

Private Sub Form_Load()
Dim lsSQL As String

   CenterChildForm mdiMain, Me
   bLoaded = False
    
   Set oDriver = New FormDriver
   Set oDriver.AppDriver = oApp
   Set oDriver.MainForm = Me
   Set oSerial = New frmCP_Serial_Transfer
   
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
                  & " FROM CP_Serial_Transfer_Master " _

   oDriver.BrowseQuery = "SELECT" _
                  & " a.sTransNox, " _
                  & " a.sReferNox, " _
                  & " a.dTransact, " _
                  & " b.sBranchNm  " _
            & " FROM CP_Serial_Transfer_Master a " _
               & " LEFT JOIN Branch b " _
                  & " ON a.sDestinat = b.sBranchCd " _
            & " WHERE cTranStat = 0 " _
               & " AND sOriginxx = '" & oApp.BranchCode & "'" _
            & " ORDER BY dTransact Desc "
   
   oDriver.InitRecForm

   oDriver.BrowseColumn(0) = "sTransNox"
   oDriver.BrowseColumn(1) = "sReferNox"
   oDriver.BrowseColumn(2) = "dTransact"
   oDriver.BrowseColumn(3) = "sBranchNm"
   
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
       
   oDriver.FieldStart = 2
   oDriver.FieldFormat(3) = "MMMM DD, YYYY"
   EmptyGrid

End Sub

Function Reference_No() As String
   Dim lrs As Recordset
   Dim lsSQL As String
   Dim lnCtr As Long
   
   lsSQL = "SELECT TOP 1" & _
            " sReferNox" & _
            " FROM CP_Serial_Transfer_Master " & _
            " WHERE sReferNox LIKE " & strParm(oApp.BranchCode & "W-%") & _
            " ORDER BY sReferNox DESC"
   
   Set lrs = New Recordset
   lrs.Open lsSQL, oApp.Connection, , , adCmdText
   
   If lrs.EOF Then
      lnCtr = 1
   Else
      If Left(lrs("sReferNox"), 2) = oApp.BranchCode Then
         lnCtr = CLng(Right(lrs("sReferNox"), 4)) + 1
      Else
         lnCtr = 1
      End If
   End If
   
   Reference_No = oApp.BranchCode & "W-" & Format(Date, "yy") & Format(lnCtr, "0000")
   Set lrs = Nothing
End Function

Function Transaction_No() As String
   Dim lrs As Recordset
   Dim lsSQL As String
   Dim lnCtr As Long
   
   lsSQL = "SELECT TOP 1" & _
            " sTransNox" & _
            " FROM CP_Serial_Transfer_Master " & _
            " WHERE sTransNox LIKE " & strParm(oApp.BranchCode & "W-%") & _
            " ORDER BY sTransNox DESC"
   
   Set lrs = New Recordset
   lrs.Open lsSQL, oApp.Connection, , , adCmdText
   
   If lrs.EOF Then
      lnCtr = 1
   Else
      If Left(lrs("sTransNox"), 2) = oApp.BranchCode Then
         lnCtr = CLng(Right(lrs("sTransNox"), 4)) + 1
      Else
         lnCtr = 1
      End If
   End If
   
   Transaction_No = oApp.BranchCode & "W-" & Format(Date, "yy") & Format(lnCtr, "0000")
   Set lrs = Nothing
End Function
Private Sub InitGrid()
    With GridEditor1
      .Rows = 2
      .Cols = 8
      .Font = "MS Sans Serif"
              
      'column title
      .TextMatrix(0, 1) = "IMEI No."
      .TextMatrix(0, 2) = "Bar Code"
      .TextMatrix(0, 3) = "Particulars"
      .TextMatrix(0, 4) = "SRP"
      .TextMatrix(0, 5) = "Stock ID"
      .TextMatrix(0, 6) = "Serial ID"
      .TextMatrix(0, 7) = "Pur. Price"
      .Row = 0
      
      'column width
      .ColWidth(0) = 400
      .ColWidth(1) = 1600
      .ColWidth(2) = 1800
      .ColWidth(3) = 4850
      .ColWidth(4) = 1000
      .ColWidth(5) = 0
      .ColWidth(6) = 0
      .ColWidth(7) = 0
              
      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      .ColAlignment(4) = 6
      .ColAlignment(5) = 1
      .ColAlignment(6) = 1
      .ColAlignment(7) = 6
            
      .ColEnabled(2) = False
      .ColEnabled(3) = False
      .ColEnabled(4) = False
      .ColEnabled(5) = False
      .ColEnabled(6) = False
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
      If .ColEnabled(1) = False Then .ColEnabled(1) = True
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
         If .Col = 1 Then
            Search_Serial
         End If
      End If
   End With
End Sub

Private Sub Print_Transaction()
Dim lnCtr As Integer
Dim lrs As New ADODB.Recordset
Dim lsSQL As String
Dim lrsDetail As New ADODB.Recordset

   Set lrs = New ADODB.Recordset
   lrs.Fields.Append "sField06", adVarChar, 150
   lrs.Fields.Append "sField09", adVarChar, 10
   lrs.Fields.Append "sField10", adVarChar, 20
   lrs.Open

   'CP_Serial_Transfer_Master
   lsSQL = "SELECT" _
               & " a.sTransNox, " _
               & " a.sReferNox, " _
               & " a.sApproved, " _
               & " a.sDestinat, " _
               & " b.sBranchNm, " _
               & " b.sAddressx + ', ' + c.sTownName xAddressx, " _
               & " a.dTransact, " _
               & " a.sRemarksx, " _
               & " a.sApproved  " _
         & " FROM CP_Serial_Transfer_Master a " _
            & " LEFT JOIN Branch b " _
               & " ON a.sOriginxx = b.sBranchCd " _
            & " LEFT JOIN TownCity c " _
               & " ON b.sTownIDxx = c.sTownIDxx " _
         & " WHERE a.sTransNox = '" & Reference & "' "
   
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
                  & " b.sIMEINoxx, " _
                  & " b.sStockIDx, " _
                  & " c.sDescript, " _
                  & " d.sBrandNme, " _
                  & " e.sModelNme, " _
                  & " f.sColorNme  " _
               & " FROM CP_Serial_Transfer_Detail a " _
                  & " LEFT JOIN CP_Serial_Master b " _
                     & " ON a.sSerialID = b.sSerialID " _
                  & " LEFT JOIN CP_Inventory c " _
                     & " ON b.sStockIDx = c.sStockIDx " _
                  & " LEFT JOIN Brand d " _
                     & " ON c.sBrandIdx = d.sBrandIDx " _
                  & " LEFT JOIN Model e  " _
                     & " ON c.sModelIDx = e.sModelIDx " _
                  & " LEFT JOIN Color f  " _
                     & " ON c.sColorIDx = f.sColorIdx " _
               & " WHERE a.sTransNox = '" & oRS("sTransNox") & "' " _
               & " ORDER BY a.nEntryNox "

            If lrsDetail.State = adStateOpen Then lrsDetail.Close
            lrsDetail.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
            Do While Not lrsDetail.EOF
               lrs.AddNew
               lrs("sField06").Value = Trim(IIf(IsNull(lrsDetail("sBrandNme")), "", lrsDetail("sBrandNme")) _
                                       & " " & IIf(IsNull(lrsDetail("sModelNme")), "", lrsDetail("sModelNme")) _
                                       & " " & IIf(IsNull(lrsDetail("sDescript")), "", lrsDetail("sDescript")) _
                                       & " " & IIf(IsNull(lrsDetail("sColorNme")), "", lrsDetail("sColorNme")))
               lrs("sField09").Value = lrsDetail("sStockIDx")
               lrs("sField10").Value = lrsDetail("sIMEINoxx")
               lrsDetail.MoveNext
            Loop

         oRS.MoveNext
      Next
      
      Branch = txtField(2)
      getBranch Code, Branch, Address
         
      Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\CP_Transmittal_Serial.rpt")
      oReport.DiscardSavedData
      oReport.FieldMappingType = crAutoFieldMapping
      oReport.Database.SetDataSource lrs
      
      oRS.MoveFirst
      
      With oReport
         .Sections("PH").ReportObjects("txtReportDate").SetText Format(oRS("dTransact"), "MMMM dd, yyyy")
         .Sections("PH").ReportObjects("txtTransmittal").SetText oRS("sReferNox")
         .Sections("PH").ReportObjects("txtToBranch").SetText Branch
         .Sections("PH").ReportObjects("txtToAddress").SetText Address
         .Sections("PH").ReportObjects("txtFromBranch").SetText oRS("sBranchNm")
         .Sections("PH").ReportObjects("txtFromAddress").SetText oRS("xAddressx")
      
         .Sections("PF").ReportObjects("txtApproved").SetText oRS("sApproved")
         .Sections("PF").ReportObjects("txtPrepared").SetText oApp.UserName
         .Sections("RF").ReportObjects("txtRemarks").SetText oRS("sRemarksx")
      End With

      Set lrs = Nothing
      Set oRS = Nothing
      Set lrsDetail = Nothing
      
      frmViewer.CRViewer91.ReportSource = oReport
      frmViewer.CRViewer91.ViewReport
      frmViewer.Show

End Sub

Private Sub Search_Serial()
Dim lsSQL As String
Dim lsCondition As String
Dim lsSearch As String
Dim lnCtr As Integer
   
   With GridEditor1
      lsSQL = "SELECT" _
            & " Distinct " _
            & " f.sIMEINoxx, " _
            & " a.sBarrcode, " _
            & " a.sStockIDx, " _
            & " b.sBrandNme, " _
            & " c.sModelNme, " _
            & " a.sDescript, " _
            & " d.sColorNme, " _
            & " a.nSelPrice, " _
            & " f.sSerialID, " _
            & " a.nPurPrice  "
      lsSQL = lsSQL _
         & " FROM CP_Serial_Master f " _
            & " LEFT JOIN CP_Inventory a " _
               & " ON f.sStockIDx = a.sStockIDx " _
            & " LEFT JOIN CP_Inventory_Master e " _
               & " ON a.sStockIdx = e.sStockIDx " _
            & " LEFT JOIN Brand b " _
               & " ON a.sBrandIdx = b.sBrandIdx " _
            & " LEFT JOIN Model c " _
               & " ON a.sModelIdx = c.sModelIdx " _
            & " LEFT JOIN Color d " _
               & " ON a.sColorIDx = d.sColorIDx " _
         & " WHERE f.sIMEINoxx like  '%" & .TextMatrix(.Row, 1) & "%' " _
            & " AND (sCategIDx = '01001' or sCategIDx = '01002' or sCategIDx = '01003') " _
            & " AND f.cSoldStat = 0 " _
            & " AND f.sBranchCd = '" & oApp.BranchCode & "' " _
            & " AND f.cLocation = 1 "
      
      If oRS.State = adStateOpen Then oRS.Close
      oRS.Open lsSQL, oApp.Connection, adOpenKeyset, adLockOptimistic, adCmdText
      
         If Not oRS.EOF Then
            If oRS.RecordCount = 1 Then
               .TextMatrix(.Row, 1) = IIf(IsNull(oRS(0)), "", oRS(0))
               .TextMatrix(.Row, 2) = IIf(IsNull(oRS(1)), "", oRS(1))
               .TextMatrix(.Row, 3) = Trim(IIf(IsNull(oRS(3)), "", oRS(3)) & " " & _
                                       IIf(IsNull(oRS(4)), "", oRS(4)) & " " & _
                                       IIf(IsNull(oRS(5)), "", oRS(5)) & " " & _
                                       IIf(IsNull(oRS(6)), "", oRS(6)))
               .TextMatrix(.Row, 4) = IIf(IsNull(oRS(7)), "", Format(oRS(7), "#,##0.00"))
               .TextMatrix(.Row, 5) = IIf(IsNull(oRS(2)), "", oRS(2))
               .TextMatrix(.Row, 6) = IIf(IsNull(oRS(8)), "", oRS(8))
               .TextMatrix(.Row, 7) = IIf(IsNull(oRS(9)), "", Format(oRS(9), "#,##0.00"))
            Else
               lsSearch = KwikSearch(oApp, lsSQL, _
                          "sIMEINoxx»sBarrcode»sBrandNme»sModelNme»sDescript»sColorNme", _
                          "IMEI No.»Bar Code»Brand»Model»Description»Color")
               If lsSearch <> "" Then
                  psSelected = Split(lsSearch, "»")
                  .TextMatrix(.Row, 1) = IIf(IsNull(psSelected(0)), "", psSelected(0))
                  .TextMatrix(.Row, 2) = IIf(IsNull(psSelected(1)), "", psSelected(1))
                  .TextMatrix(.Row, 3) = Trim(IIf(IsNull(psSelected(3)), "", psSelected(3)) & " " & _
                                          IIf(IsNull(psSelected(4)), "", psSelected(4)) & " " & _
                                          IIf(IsNull(psSelected(5)), "", psSelected(5)) & " " & _
                                          IIf(IsNull(psSelected(6)), "", psSelected(6)))
                  .TextMatrix(.Row, 4) = IIf(IsNull(psSelected(7)), "", Format(psSelected(7), "#,##0.00"))
                  .TextMatrix(.Row, 5) = IIf(IsNull(psSelected(2)), "", psSelected(2))
                  .TextMatrix(.Row, 6) = IIf(IsNull(psSelected(8)), "", psSelected(8))
                  .TextMatrix(.Row, 7) = IIf(IsNull(psSelected(9)), "", Format(psSelected(9), "#,##0.00"))
               Else
                  pbExisting = True
                  Exit Sub
               End If
            End If
            
            For pnCtr = 1 To .Rows - 1
               If .Row <> pnCtr Then
                  If .TextMatrix(.Row, 6) = .TextMatrix(pnCtr, 6) Then
                     MsgBox "Duplicate IMEI Entry!!!" & vbCrLf & _
                     "Verify your Entry", vbCritical, "Warning"
                     For lnCtr = 2 To 7
                        .TextMatrix(.Row, lnCtr) = ""
                     Next
                     .SetFocus
                     pbExisting = True
                     Exit Sub
                  End If
               End If
            Next
            .Refresh
            If .TextMatrix(.Row, 1) <> "" Then
               .Rows = .Rows + 1
               .Row = .Row + 1
            End If
            .SetFocus
         Else
            MsgBox "IMEI NO. Not Existing!!!", vbCritical, "Warning"
            For pnCtr = 1 To .Cols
               .TextMatrix(.Row, pnCtr) = ""
            Next
            .Col = 1
            .SetFocus
            .Refresh
         End If
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
   Select Case oDriver.FieldValue(9)
      Case 0
         Label.Caption = "UNKNOWN"
      Case 1
         Label.Caption = "POSTED"
   End Select
   oDriver.FieldValue(3) = Format(oDriver.FieldValue(3), "m/d/yyyy")
End Sub

Private Sub oDriver_Save(Saved As Boolean)
   Saved = False
End Sub

Private Sub oDriver_SaveComplete()
Dim orig As String
Dim lsSQL As String
Dim lsCondition As String
   oDriver.ShowButton 4
   MsgBox "Transaction Successfully Saved!!!", vbInformation, "Information"
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
            Cancel = Not Save_CP_Serial
               If Cancel Then Exit Sub
            Cancel = Not Update_CP_Inventory
               If Cancel Then Exit Sub
            Cancel = Not Save_CP_Inventory_Ledger
               If Cancel Then Exit Sub
            oDriver.FieldValue(3) = CDate(txtField(3).Text) & " " & Time
            oDriver.FieldValue(7) = oApp.BranchCode
            oDriver.FieldValue(8) = .TextMatrix(.Rows - 1, 0)
            oDriver.FieldValue(9) = 0  'cTranStat
            oDriver.FieldValue(10) = 0  'cReceived
            Reference = oDriver.FieldValue(0)
            Branch = txtField(2)
         End If
      End If
   End With
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   oDriver.ColumnIndex = Index
   txtfieldGotfocus = True
   pnIndex = Index
   txtField(Index).BackColor = &HE1FEFF
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Or KeyCode = 13 Then
      If Index = 2 Then
         oDriver.RecordSearch txtField(Index).Text
         If txtField(Index).Text <> "" Then SetNextFocus
      End If
   End If
   KeyCode = 0
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

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
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
         Call Modified("CP_Serial_Transfer_Master", "sTransNox = '" & oDriver.FieldValue(0) & "' ")
   End Select
End Sub

Private Sub ShowGrid()
Dim lsSQL As String
Dim lnCtr As Integer
Dim lnRow As Long

   lsSQL = "SELECT" _
               & " Distinct " _
               & " a.sSerialID, " _
               & " a.sTransNox, " _
               & " a.nEntryNox, " _
               & " a.nUnitPrce, " _
               & " b.sIMEINoxx, " _
               & " b.sStockIDx, " _
               & " c.sBarrCode, " _
               & " c.sDescript, " _
               & " e.sBrandNme, " _
               & " f.sModelNme, " _
               & " g.sColorNme, " _
               & " c.nSelPrice  "
   lsSQL = lsSQL _
         & " FROM CP_Serial_Transfer_Detail a " _
               & " LEFT JOIN CP_Serial_Master b " _
                  & " ON a.sSerialID = b.sSerialID " _
               & " LEFT JOIN CP_Inventory c " _
                  & " ON b.sStockIDx = c.sStockIDx " _
               & " LEFT JOIN CP_Inventory_Master d " _
                  & " ON b.sStockIDx = d.sStockIDx " _
               & " LEFT JOIN Brand e " _
                  & " ON c.sBrandIDx = e.sBrandIDx " _
               & " LEFT JOIN Model f " _
                  & " ON c.sModelIDx = f.sModelIDx " _
               & " LEFT JOIN Color g " _
                  & " ON c.sColorIDx = g.sColorIDx " _
         & " WHERE a.sTransNox = '" & oDriver.FieldValue(0) & " '" _
         & " ORDER BY a.nEntryNox "
   If oRS.State = adStateOpen Then oRS.Close
   oRS.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText

   If oRS.RecordCount <> 0 Then
      With GridEditor1
         .Rows = oRS.RecordCount + 1
         For lnCtr = 0 To oRS.RecordCount - 1
            .TextMatrix(lnCtr + 1, 0) = oRS("nEntryNox")
            .TextMatrix(lnCtr + 1, 1) = oRS("sIMEINoxx")
            .TextMatrix(lnCtr + 1, 2) = oRS("sBarrCode")
            .TextMatrix(lnCtr + 1, 3) = Trim(IIf(IsNull(oRS("sBrandNme")), "", oRS("sBrandNme")) & " " & _
                                          IIf(IsNull(oRS("sModelNme")), "", oRS("sModelNme")) & " " & _
                                          IIf(IsNull(oRS("sDescript")), "", oRS("sDescript")) & " " & _
                                          IIf(IsNull(oRS("sColorNme")), "", oRS("sColorNme")))
            .TextMatrix(lnCtr + 1, 4) = Format(oRS("nSelPrice"), "#,##0.00")
            .TextMatrix(lnCtr + 1, 5) = oRS("sStockIDx")
            .TextMatrix(lnCtr + 1, 6) = oRS("sSerialID")
            .TextMatrix(lnCtr + 1, 7) = oRS("nUnitPrce")
            oRS.MoveNext
         Next
         If .Rows > 20 Then
            .ColWidth(3) = 4600
         Else
            .ColWidth(3) = 4850
         End If
         .ColEnabled(1) = False
      End With
   Else
      Exit Sub
   End If

   Set oRS = Nothing

End Sub

Private Function Save_CP_Inventory_Ledger() As Boolean
Dim lsSQL As String
Dim lnRow As Long
Dim lnEntry As Integer
Dim QOH As Integer

Save_CP_Inventory_Ledger = True
On Error GoTo errProc
   
   With GridEditor1
      For pnCtr = 1 To .Rows - 1
         If .TextMatrix(pnCtr, 1) = "" Then Exit For
         'Search sSourceNo
         lsSQL = "SELECT" _
                  & " sStockIDx, " _
                  & " sSourceNo  " _
               & " FROM CP_Inventory_Ledger " _
               & " WHERE sStockIdx = '" & .TextMatrix(pnCtr, 5) & "'" _
                  & " AND sSourceNo = '" & oDriver.FieldValue(0) & "'" _
                  & " AND sSourceCd = 'CPDv' " _
                  & " AND sBranchCd = '" & oApp.BranchCode & "'"
         If oRS.State = adStateOpen Then oRS.Close
         oRS.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
         
            'Get QOH
            QOH = getQuantity("'" & .TextMatrix(pnCtr, 5) & "'", "'" & oApp.BranchCode & "'")
            
            If oRS.EOF = False Then
               'Update Record, CP_Inventory_Ledger
               lsSQL = "UPDATE CP_Inventory_Ledger SET" _
                        & " nQtyOutxx = nQtyOutxx + 1 , " _
                        & " nQtyOnHnd = '" & CLng(QOH) & "'," _
                        & " dModified = getdate() " _
                  & " WHERE sStockIdx = '" & .TextMatrix(pnCtr, 5) & "'" _
                     & " AND sSourceNo = '" & oDriver.FieldValue(0) & "'" _
                     & " AND sSourceCd = 'CPDv' " _
                     & " AND sBranchCd = '" & oApp.BranchCode & "'"
               oApp.Connection.Execute lsSQL, lnRow, adCmdText
            
            Else
               'Get last Entry No.
               lnEntry = getEntryNo("CP_Inventory_Ledger", "'" & .TextMatrix(pnCtr, 5) & "'", _
                           "'" & oApp.BranchCode & "'")
               
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
                     & "('" & .TextMatrix(pnCtr, 5) & "', " _
                     & "'" & oApp.BranchCode & "', " _
                     & "'" & oDriver.FieldValue(2) & "', " _
                     & "'CPDv' , " _
                     & "'" & oDriver.FieldValue(0) & "', " _
                     & " 0, " _
                     & "'1', " _
                     & "'" & CLng(QOH) & "', " _
                     & "'" & lnEntry & "', " _
                     & "'" & CDate(oDriver.FieldValue(3)) & " " & Time & "', " _
                     & " getdate())"
               oApp.Connection.Execute lsSQL, lnRow, adCmdText
               
            End If
         Set oRS = Nothing
      Next

      If lnRow <= 0 Then
         MsgBox "Unable to Update Inventory Ledger!!!" & vbCrLf & vbCrLf & _
         "Notify ROSALYN LAZO DESCALLAR for Assistance", vbCritical, "Warning"
         Save_CP_Inventory_Ledger = False
         GoTo endProc
      End If
   End With

endProc:
   Exit Function
errProc:
   Save_CP_Inventory_Ledger = False
   MsgBox Err.Description, vbCritical, "Warning"
End Function

Private Function Update_CP_Inventory() As Boolean
Dim lsSQL As String
Dim lnRow As Long

Update_CP_Inventory = True
On Error GoTo errProc
   
   With GridEditor1
         For pnCtr = 1 To .Rows - 1
            If .TextMatrix(pnCtr, 1) = "" Then Exit For
            'Update QOH, CP_Inventory_Master
            lsSQL = "UPDATE CP_Inventory_Master SET" _
                  & " nQtyOnHnd = nQtyOnHnd - 1, " _
                  & " sModified = '" & Encrypt(oApp.UserID) & "'," _
                  & " dModified = getdate() " _
            & " WHERE sStockIDx = '" & .TextMatrix(pnCtr, 5) & "'" _
                  & " AND sBranchCd = '" & oApp.BranchCode & "'"
            oApp.Connection.Execute lsSQL, lnRow, adCmdText
         Next
   
         If lnRow <= 0 Then
            MsgBox "Unable to Update Inventory Master!!!" & vbCrLf & vbCrLf & _
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

Private Function Save_CP_Serial() As Boolean
Dim lsSQL As String
Dim lnRow As Long
Dim lnEntry As Integer
   
Save_CP_Serial = True
On Error GoTo errProc
   
   With GridEditor1
      For pnCtr = 1 To .Rows - 1
         If .TextMatrix(pnCtr, 1) = "" Then Exit For
         
         'Get last Entry No.
         lnEntry = getIMEIEntry("'" & .TextMatrix(pnCtr, 6) & "'")
         
         'Save_CP_Serial_Transfer_Detail
         lsSQL = "INSERT INTO CP_Serial_Transfer_Detail " _
               & "( sTransNox, " _
               & "  nEntryNox, " _
               & "  sSerialID, " _
               & "  nUnitPrce, " _
               & "  dModified) " _
         & "VALUES " _
               & "('" & oDriver.FieldValue(0) & "', " _
               & "'" & .TextMatrix(pnCtr, 0) & "', " _
               & "'" & .TextMatrix(pnCtr, 6) & "', " _
               & "'" & CDbl(.TextMatrix(pnCtr, 7)) & "', " _
               & " getdate())"
         oApp.Connection.Execute lsSQL, lnRow, adCmdText

         'Update CP_Serial_Master
         lsSQL = "UPDATE CP_Serial_Master SET" _
               & " sBranchCd = '" & oDriver.FieldValue(2) & "', " _
               & " cLocation = '0', " _
               & " sModified = '" & Encrypt(oApp.UserID) & "', " _
               & " dModified = getdate() " _
         & " WHERE sSerialID = '" & .TextMatrix(pnCtr, 6) & "' "
         oApp.Connection.Execute lsSQL, lnRow, adCmdText
                        
         'CP_Serial_Ledger
         lsSQL = "INSERT INTO CP_Serial_Ledger " _
               & "( sSerialID, " _
               & "  sBranchCd, " _
               & "  dTransact, " _
               & "  nEntryNox, " _
               & "  sSourceCd, " _
               & "  sSourceNo, " _
               & "  cSoldStat, " _
               & "  cLocation, " _
               & "  dModified) " _
         & "VALUES " _
               & "('" & .TextMatrix(pnCtr, 6) & "', " _
               & "'" & oApp.BranchCode & "', " _
               & "'" & CDate(oDriver.FieldValue(3)) & " " & Time & "', " _
               & "'" & lnEntry & "', " _
               & "'CPDv', " _
               & "'" & oDriver.FieldValue(0) & "', " _
               & "'0'," _
               & "'1', " _
               & " getdate())"
         oApp.Connection.Execute lsSQL, lnRow, adCmdText
                        
      Next
            
      If lnRow <= 0 Then
         MsgBox "Unable to Save CP Serial!!!" & vbCrLf & vbCrLf & _
         "Notify ROSALYN LAZO DESCALLAR for Assistance", vbCritical, "Warning"
         Save_CP_Serial = False
         GoTo endProc
      End If
   End With

endProc:
   Exit Function
errProc:
   Save_CP_Serial = False
   MsgBox Err.Description, vbCritical, "Warning"
End Function

'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  Date Tested  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  February 25, 2007  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'

'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤    Version 1    ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  Date Finished  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  February 26, 2007  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'




