VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmLoad_Transfer 
   BorderStyle     =   0  'None
   Caption         =   "Load Wallet Transfer"
   ClientHeight    =   4125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7500
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   4125
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1260
      Index           =   2
      Left            =   1590
      Tag             =   "wt0;fb0"
      Top             =   1260
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   2223
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   0
         Left            =   1305
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   165
         Width           =   1980
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   2
         Left            =   1305
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   810
         Width           =   4365
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   1305
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   555
         Width           =   2025
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   240
         Left            =   1350
         Tag             =   "et0;ht2"
         Top             =   210
         Width           =   1980
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Reference No."
         Height          =   285
         Index           =   4
         Left            =   180
         TabIndex        =   22
         Top             =   180
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Destination"
         Height          =   285
         Index           =   19
         Left            =   180
         TabIndex        =   2
         Top             =   810
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date Transact"
         Height          =   285
         Index           =   6
         Left            =   180
         TabIndex        =   0
         Top             =   555
         Width           =   1110
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1485
      Index           =   0
      Left            =   1590
      Tag             =   "wt0;fb0"
      Top             =   2535
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   2619
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   5
         Left            =   1290
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   495
         Width           =   2460
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   4
         Left            =   1290
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   1005
         Width           =   2025
      End
      Begin VB.TextBox txtothers 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   0
         Left            =   1290
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   240
         Width           =   2460
      End
      Begin VB.TextBox txtothers 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   1290
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   750
         Width           =   4365
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
         Left            =   4080
         MaxLength       =   128
         MultiLine       =   -1  'True
         TabIndex        =   13
         Tag             =   "ht0;ft0"
         Text            =   "frmLoad_Transfer.frx":0000
         Top             =   1005
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Balance"
         Height          =   255
         Index           =   5
         Left            =   165
         TabIndex        =   6
         Top             =   495
         Width           =   1185
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cell Number"
         Height          =   285
         Index           =   3
         Left            =   165
         TabIndex        =   10
         Top             =   1005
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bar Code"
         Height          =   255
         Index           =   2
         Left            =   165
         TabIndex        =   4
         Top             =   240
         Width           =   1185
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   285
         Index           =   1
         Left            =   165
         TabIndex        =   8
         Top             =   750
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount "
         Height          =   240
         Index           =   0
         Left            =   3435
         TabIndex        =   12
         Top             =   1020
         Width           =   585
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   3
      Left            =   90
      TabIndex        =   20
      Top             =   2100
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
      Picture         =   "frmLoad_Transfer.frx":0006
      CaptionAlign    =   0
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   0
      Left            =   90
      TabIndex        =   15
      Top             =   840
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
      Picture         =   "frmLoad_Transfer.frx":0780
      CaptionAlign    =   0
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   2
      Left            =   90
      TabIndex        =   17
      Top             =   1260
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
      Picture         =   "frmLoad_Transfer.frx":0EFA
      CaptionAlign    =   0
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   5
      Left            =   90
      TabIndex        =   19
      Top             =   2100
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
      Picture         =   "frmLoad_Transfer.frx":1674
      CaptionAlign    =   0
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   405
      Index           =   1
      Left            =   90
      TabIndex        =   16
      Top             =   1260
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
      Picture         =   "frmLoad_Transfer.frx":1DEE
      CaptionAlign    =   0
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   4
      Left            =   90
      TabIndex        =   18
      Top             =   1680
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
      Picture         =   "frmLoad_Transfer.frx":2568
      CaptionAlign    =   0
      BackColor       =   14286077
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   6
      Left            =   90
      TabIndex        =   14
      Top             =   840
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
      Caption         =   "&Update"
      AccessKey       =   "U"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmLoad_Transfer.frx":2CE2
      CaptionAlign    =   0
      BackColor       =   14286077
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   615
      Index           =   1
      Left            =   1590
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   1085
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   8
         Left            =   1305
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   180
         Width           =   2685
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   240
         Left            =   1350
         Tag             =   "et0;ht2"
         Top             =   225
         Width           =   2700
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction No."
         Height          =   285
         Index           =   7
         Left            =   150
         TabIndex        =   24
         Top             =   195
         Width           =   1200
      End
   End
End
Attribute VB_Name = "frmLoad_Transfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents oDriver As FormDriver
Attribute oDriver.VB_VarHelpID = -1
Private oSkin As FormSkin
Private bLoaded As Boolean
Dim pbnewitem As Boolean

Dim oRS As New ADODB.Recordset
Dim pbLoading As Boolean
Dim pnindex As Integer
Dim psSelected() As String
Dim pnUserRights As Integer
Dim psUserID As String
Dim psUserName As String
Dim Time As String


Private Sub cmdButton_Click(Index As Integer)
Dim lnCtr As Integer
Dim lsRep As Integer

   Select Case Index
   Case 0
      oDriver.RecordSave
   Case 1
      oDriver.RecordSearch
      txtfield(pnindex).SetFocus
   Case 2
      oDriver.RecordNew
   Case 3
      pbLoading = True
      oDriver.RecordCancelUpdate
   Case 4
      oDriver.BrowseRecord
   Case 5
      Unload Me
   Case 6
      If oRS.State = adStateOpen Then oRS.Close
      oRS.Open "SELECT * From ELoad_Ledger " _
               & "WHERE sStockIDx = '" & oDriver.FieldValue(6) & "' " _
               & "ORDER by nEntryNox Desc" _
               , oApp.Connection, adOpenDynamic, adLockReadOnly, adCmdText

      oRS.MoveFirst
      If oRS("sSourceNo") <> oDriver.FieldValue(8) Then
         MsgBox "Item Has other Transactions!!!" & vbCrLf & _
               "Update Not Permitted!!!" & vbCrLf & _
               "" & vbCrLf & _
               "Notify ROSALYN LAZO DESCALLAR for Assistance!!!", vbInformation, "Notice"
         Exit Sub
      End If

      If DateDiff("d", oDriver.FieldValue(1), Date) > 1 Then
         lsRep = MsgBox("Update Not Permitted!!!" & vbCrLf & _
                  "Seek for Approval!", vbQuestion + vbYesNo, "Confirm")
         If lsRep <> vbYes Then Exit Sub
         If Not GetApproval(oApp, pnUserRights, psUserID, psUserName) Then Exit Sub
            If pnUserRights < xeSupervisor Then
               MsgBox "Approving User is Not Authorized!!!", vbCritical, "Warning"
               Exit Sub
            End If
      Else
         If oApp.UserLevel = xeEncoder Then
            lsRep = MsgBox("Seek for Approval?", vbQuestion + vbYesNo, "Confirm")
            If lsRep <> vbYes Then Exit Sub
            If Not GetApproval(oApp, pnUserRights, psUserID, psUserName) Then Exit Sub
            If pnUserRights < xeSupervisor Then
               MsgBox "Approving User is Not Authorized!!!", vbCritical, "Warning"
               Exit Sub
            End If
         End If
      End If
      pbLoading = False
      oDriver.RecordUpdate
   End Select

End Sub

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0

   If bLoaded = False Then
      oDriver.RecordNew
      bLoaded = True
   End If
End Sub

Private Sub Form_Load()
CenterChildForm mdiMain, Me
       
   bLoaded = False
   
   Set oDriver = New FormDriver
   Set oDriver.AppDriver = oApp
   Set oDriver.MainForm = Me
   
   Set oSkin = New FormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransaction
   
   oDriver.RecQuery = "SELECT " _
                        & "  sReferNox, " _
                        & "  dTransact, " _
                        & "  sBranchCd, " _
                        & "  nQtyOutxx, " _
                        & "  sPhoneNum, " _
                        & "  nQtyOnHnd, " _
                        & "  sStockIDx, " _
                        & "  sSourceCd, " _
                        & "  sSourceNo, " _
                        & "  sTransNox, " _
                        & "  nQtyInxxx, " _
                        & "  nEntryNox, " _
                        & "  sModified, " _
                        & "  dModified  " _
                     & " FROM ELoad_Ledger " _

                    
   oDriver.BrowseQuery = "SELECT" _
                        & " Top 1000 " _
                        & " a.sReferNox, " _
                        & " a.sPhoneNum, " _
                        & " a.nQtyOutxx, " _
                        & " a.dTransact, " _
                        & " b.sBranchNm, " _
                        & " c.sDescript  " _
                     & " FROM ELoad_Ledger a " _
                        & " LEFT JOIN Branch b " _
                           & " ON a.sBranchCd = b.sBranchCd " _
                        & " LEFT JOIN CP_Inventory c " _
                           & " ON a.sStockIDx = c.sStockIDx " _
                     & " WHERE a.sSourceCd = 'CPDv' " _
                     & " ORDER BY a.dTransact desc"
   
   oDriver.InitRecForm

   oDriver.BrowseColumn(0) = "sReferNox"
   oDriver.BrowseColumn(1) = "sPhoneNum"
   oDriver.BrowseColumn(2) = "dTransact"
   oDriver.BrowseColumn(3) = "nQtyOutxx"
   oDriver.BrowseColumn(4) = "sDescript"
   oDriver.BrowseColumn(5) = "sBranchNm"
   
   oDriver.BrowseFTitle(0) = "Ref. No."
   oDriver.BrowseFTitle(1) = "Cell No."
   oDriver.BrowseFTitle(2) = "Date"
   oDriver.BrowseFTitle(3) = "Amount"
   oDriver.BrowseFTitle(4) = "Description"
   oDriver.BrowseFTitle(5) = "Branch"
   
   oDriver.BrowseFFormat(2) = "MMMM dd, yyyy"
   oDriver.BrowseFFormat(3) = "#,##0.00"
   
   'Branch
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
   
   oDriver.FieldStart = 0
   oDriver.FieldFormat(1) = "MMMM dd, yyyy"
   oDriver.FieldFormat(3) = "#,##0.00"
   oDriver.FieldFormat(5) = "#,##0.00"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oDriver = Nothing
   Set oSkin = Nothing
End Sub

Private Sub oDriver_DisableOtherControl()
   oDriver.DisableTextbox 5
   oDriver.DisableTextbox 8
End Sub

Private Sub oDriver_EnableOtherControl()
   oDriver.DisableTextbox 5
   oDriver.DisableTextbox 8
End Sub

Private Sub oDriver_InitValue()
   oDriver.FieldReference(8) = True
   oDriver.FieldValue(8) = Transaction_No
   txtfield(8).Text = oDriver.FieldValue(8)
      
   txtothers(0).Text = ""
   txtothers(1).Text = ""
   txtothers(1).Enabled = False
   txtfield(1).Text = Format(oApp.ServerDate, "MMMM dd, yyyy")
   txtfield(3).Text = "0.00"
   txtfield(5).Text = "0.00"
   oDriver.DisableTextbox 5
   
   oDriver.FieldValue(1) = Date
   oDriver.FieldValue(7) = "CPDv"
   oDriver.FieldValue(8) = ""
   oDriver.FieldValue(9) = 1
   oDriver.FieldValue(10) = 0
   pbLoading = False
   pbnewitem = True
End Sub

Function Transaction_No() As String
   Dim lrs As Recordset
   Dim lsSQL As String
   Dim lnCtr As Long
   
   lsSQL = "SELECT TOP 1" & _
            " sTransNox" & _
            " FROM ELoad_Ledger " & _
            " WHERE sTransNox LIKE " & strParm(oApp.BranchCode & "L-%") & _
            " ORDER BY sTransNox DESC"
   
   Set lrs = New Recordset
   lrs.Open lsSQL, oApp.Connection, , , adCmdText
   
   If lrs.EOF Then
      lnCtr = 1
   Else
      If Left(lrs("sTransNox"), 2) = oApp.BranchCode Then
         lnCtr = CLng(Right(lrs("sTransNox"), 8)) + 1
      Else
         lnCtr = 1
      End If
   End If
   
   Transaction_No = oApp.BranchCode & "LT-" & Format(Date, "yy") & Format(lnCtr, "00000000")
   Set lrs = Nothing
End Function


Private Sub oDriver_LoadOtherData()
Dim lsSQL As String
   
   pbnewitem = False
   
   lsSQL = "SELECT" _
                & " sStockIDx, " _
                & " sBarrCode, " _
                & " sDescript " _
        & " FROM CP_Inventory  " _
        & " Where sStockIDx = '" & oDriver.FieldValue(6) & "' "
   If oRS.State = adStateOpen Then oRS.Close
   oRS.Open lsSQL, oApp.Connection, adOpenKeyset, adLockOptimistic, adCmdText
   
   If Not oRS.EOF Then
      oDriver.FieldValue(6) = oRS(0)
      txtothers(0).Text = oRS(1)
      txtothers(1).Text = oRS(2)
   End If
   oDriver.FieldValue(1) = Format(oDriver.FieldValue(1), "m/d/yyyy")
   
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   If pbLoading Then Exit Sub
   oDriver.ColumnIndex = Index
   pnindex = Index
   txtfield(Index).BackColor = &HE1FEFF
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Or KeyCode = 13 Then
      oDriver.RecordSearch txtfield(2).Text
      If txtfield(Index).Text <> "" Then SetNextFocus
      KeyCode = 0
   End If
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
   Case 27
      Call Modified("ELoad_Ledger", "sReferNox = '" & oDriver.FieldValue(0) & "' ")
   End Select
End Sub

Private Sub oDriver_WillSave(Cancel As Boolean)
Dim lsSQL As String

   If oDriver.FieldValue(0) = "" Then
      MsgBox "Invalid Reference No. Detected!!!", vbCritical, "Warning"
      txtfield(0).SetFocus
      Cancel = True
   ElseIf oDriver.FieldValue(1) = "" Then
      MsgBox "Invalid Date Detected!!!", vbCritical, "Warning"
      txtfield(1).SetFocus
      Cancel = True
   ElseIf oDriver.FieldValue(2) = "" Then
      MsgBox "Invalid Destination Detected!!!", vbCritical, "Warning"
      txtfield(2).SetFocus
      Cancel = True
   ElseIf oDriver.FieldValue(3) = "" Or oDriver.FieldValue(3) = 0# Then
      MsgBox "Invalid Amount Detected!!!", vbCritical, "Warning"
      txtfield(3).SetFocus
      Cancel = True
   ElseIf oDriver.FieldValue(4) = "" Then
      MsgBox "Invalid Cell No. Detected!!!", vbCritical, "Warning"
      txtfield(4).SetFocus
      Cancel = True
   Else
      Time = Format(Now, "hh:nn:ss AM/PM")
      If pbnewitem Then
         Cancel = Not Update_CP_Inventory
            If Cancel Then Exit Sub
         oDriver.FieldValue(5) = CDbl(txtfield(5).Text) - CDbl(txtfield(3).Text)
      Else
         Cancel = Not Delete_Transaction
            If Cancel Then Exit Sub
         Cancel = Not Update_CP_Inventory
            If Cancel Then Exit Sub
      End If
      oDriver.FieldValue(1) = CDate(txtfield(1).Text) & " " & Time
      oDriver.FieldValue(3) = CDbl(txtfield(3).Text)
      oDriver.FieldValue(12) = Encrypt(oApp.UserID)
      
      'Get last Entry No.
      lsSQL = "SELECT" _
               & " sStockIDx ," _
               & " nEntryNox  " _
            & " FROM ELoad_Ledger " _
            & " WHERE sStockIdx = '" & oDriver.FieldValue(6) & "' " _
            & " ORDER by nEntryNox desc "
      If oRS.State = adStateOpen Then oRS.Close
      oRS.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
         If Not oRS.EOF Then
            oDriver.FieldValue(11) = oRS("nEntryNox") + 1
         Else
            oDriver.FieldValue(11) = 1
         End If
      Set oRS = Nothing
      
   End If
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   txtfield(Index).BackColor = &H80000005
End Sub

Private Sub txtfield_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 1
         If Not IsDate(txtfield(Index).Text) Then
            txtfield(Index).Text = Format(oApp.ServerDate, "MMMM dd,yyyy")
         Else
            txtfield(Index).Text = Format(txtfield(Index).Text, "MMMM dd,yyyy")
         End If
      Case 3, 5
         If Not IsNumeric(txtfield(Index).Text) Then
            txtfield(Index).Text = 0#
            txtfield(Index).SetFocus
         Else
            txtfield(Index).Text = Format(CDbl(txtfield(Index).Text), "#,##0.00")
         End If
   End Select
   Cancel = Not oDriver.ValidateField(Index)
End Sub

Private Sub SearchBarCode()
Dim lsSQL As String
Dim lsSearch As String
   
   lsSQL = "SELECT" _
         & " a.sBarrcode, " _
         & " a.sStockIDx, " _
         & " a.sDescript, " _
         & " b.nQtyOnHnd  " _
      & " FROM CP_Inventory a " _
         & " LEFT JOIN CP_Inventory_Master b " _
            & " ON a.sStockIDx = b.sStockIDx " _
      & " WHERE a.sBarrcode like  '" & txtothers(0).Text & "%' " _
         & " AND cWalletxx = 1 "
      If oRS.State = adStateOpen Then oRS.Close
      oRS.Open lsSQL, oApp.Connection, adOpenKeyset, adLockOptimistic, adCmdText

   If Not oRS.EOF Then
      If oRS.RecordCount = 1 Then
         txtothers(0).Text = oRS(0)
         txtothers(1).Text = oRS(2)
         txtfield(5).Text = Format(oRS(3), "#,##0.00")
         oDriver.FieldValue(6) = oRS(1)
      Else
         lsSearch = KwikSearch(oApp, lsSQL, _
                    "sBarrcode»sDescript»nQtyOnHnd", _
                    "Bar Code»Description»Qty. On Hand", _
                    "@»@»#,##0.00")
                    
         If lsSearch <> "" Then
            psSelected = Split(lsSearch, "»")
            txtothers(0).Text = psSelected(0)
            txtothers(1).Text = psSelected(2)
            txtfield(5).Text = Format(psSelected(3), "#,##0.00")
            oDriver.FieldValue(6) = psSelected(1)
         End If
      End If
   End If
   Set oRS = Nothing
      
End Sub

Private Function Delete_Transaction() As Boolean
Dim lsSQL As String
Dim lnCtr As Integer
Dim lnrow As Long
   
Delete_Transaction = True
On Error GoTo errProc

   'Roll Back QOH in CP_Inventory_Master
   lsSQL = "SELECT" _
            & " sReferNox, " _
            & " nQtyOutxx, " _
            & " sStockIDx  " _
         & " FROM ELoad_Ledger " _
         & " WHERE sReferNox = '" & oDriver.FieldValue(0) & "' "
   If oRS.State = adStateOpen Then oRS.Close
   oRS.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
      If oRS.RecordCount <> 0 Then
         oDriver.FieldValue(5) = (CDbl(oDriver.FieldValue(5)) + CDbl(oRS(1))) - CDbl(txtfield(3).Text)
         Do While Not oRS.EOF
            lsSQL = "UPDATE CP_Inventory_Master SET" _
                  & " nQtyOnHnd = nQtyOnHnd + '" & CDbl(oRS("nQtyOutxx")) & "' " _
                  & " sModified = '" & Encrypt(oApp.UserID) & "' ," _
                  & " dModified = getdate() " _
            & " WHERE sStockIDx = '" & oRS("sStockIDx") & "' " _
                  & " And sBranchCd = '" & oApp.BranchCode & "' "
            oApp.Connection.Execute lsSQL, lnrow, adCmdText
            oRS.MoveNext
         Loop
      End If
   Set oRS = Nothing
                     
endProc:
   Exit Function
errProc:
   Delete_Transaction = False
   MsgBox Err.Description, vbCritical, "Warning"
End Function

Private Function Update_CP_Inventory() As Boolean
Dim lsSQL As String
Dim lnrow As Long

Update_CP_Inventory = True
On Error GoTo errProc
           
      'Update QOH, CP_Inventory_Master
      lsSQL = "UPDATE CP_Inventory_Master SET" _
            & " nQtyOnHnd = nQtyOnHnd - '" & CDbl(txtfield(3).Text) & "', " _
            & " sModified = '" & Encrypt(oApp.UserID) & "', " _
            & " dModified = getdate() " _
      & " WHERE sStockIDx = '" & oDriver.FieldValue(6) & "' " _
            & " And sBranchCd = '" & oApp.BranchCode & "' "
      oApp.Connection.Execute lsSQL, lnrow, adCmdText

      If lnrow <= 0 Then
         MsgBox "Unable to Update Inventory Master!!!" & vbCrLf & vbCrLf & _
         "Notify ROSALYN LAZO DESCALLAR for Assistance", vbCritical, "Warning"
         Update_CP_Inventory = False
         GoTo endProc
      End If

endProc:
   Exit Function
errProc:
   Update_CP_Inventory = False
   MsgBox Err.Description, vbCritical, "Warning"
End Function

Private Sub txtothers_GotFocus(Index As Integer)
   txtothers(Index).BackColor = &HE1FEFF
End Sub

Private Sub txtothers_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Or KeyCode = 13 Then
      SearchBarCode
      If oDriver.FieldValue(6) <> "" Then
         SetNextFocus
      Else
         MsgBox "Bar Code Not Existing!!!", vbCritical, "Warning"
         txtothers(0).SetFocus
      End If
      KeyCode = 0
   End If

End Sub

Private Sub txtothers_LostFocus(Index As Integer)
   txtothers(Index).BackColor = &H80000005
End Sub
