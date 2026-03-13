VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmAdjustment 
   BorderStyle     =   0  'None
   Caption         =   "Adjustment"
   ClientHeight    =   4320
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9030
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4320
   ScaleWidth      =   9030
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame2 
      Height          =   495
      Left            =   1620
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   873
      BackColor       =   12632256
      ClipControls    =   0   'False
      BorderStyle     =   1
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   0
         Left            =   1545
         TabIndex        =   1
         Top             =   90
         Width           =   2310
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   4830
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   90
         Width           =   2310
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   0
         Top             =   105
         Width           =   1530
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   195
         Index           =   5
         Left            =   4095
         TabIndex        =   2
         Top             =   120
         Width           =   525
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   3075
      Left            =   1620
      Tag             =   "wt0;fb0"
      Top             =   1140
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   5424
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   5
         Left            =   1560
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   2610
         Width           =   5550
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   660
         Index           =   4
         Left            =   1560
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   1935
         Width           =   5550
      End
      Begin VB.TextBox txtothers 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   0
         Left            =   1560
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   180
         Width           =   2280
      End
      Begin VB.TextBox txtothers 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   1560
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   435
         Width           =   2280
      End
      Begin VB.TextBox txtothers 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   2
         Left            =   1560
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   690
         Width           =   2280
      End
      Begin VB.TextBox txtothers 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   3
         Left            =   1560
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   945
         Width           =   5550
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmAdjustment.frx":0000
         Left            =   1545
         List            =   "frmAdjustment.frx":000A
         TabIndex        =   15
         Text            =   "Combo1"
         Top             =   1590
         Width           =   2310
      End
      Begin VB.TextBox txtothers 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   4
         Left            =   4830
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   690
         Width           =   2280
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   4845
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   1620
         Width           =   2265
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Approved By"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   2625
         Width           =   1080
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   225
         Index           =   2
         Left            =   135
         TabIndex        =   18
         Top             =   1950
         Width           =   840
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Type"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   14
         Top             =   1650
         Width           =   1500
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bar Code"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   4
         Top             =   195
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Brand"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   6
         Top             =   465
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   825
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Color"
         Height          =   195
         Index           =   10
         Left            =   4095
         TabIndex        =   10
         Top             =   705
         Width           =   585
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
         Height          =   195
         Index           =   14
         Left            =   4095
         TabIndex        =   16
         Top             =   1665
         Width           =   645
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   3
      Left            =   105
      TabIndex        =   24
      Top             =   2295
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
      Picture         =   "frmAdjustment.frx":0017
      CaptionAlign    =   0
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   0
      Left            =   105
      TabIndex        =   22
      Top             =   1455
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
      Picture         =   "frmAdjustment.frx":0791
      CaptionAlign    =   0
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   2
      Left            =   105
      TabIndex        =   23
      Top             =   1875
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
      Picture         =   "frmAdjustment.frx":0F0B
      CaptionAlign    =   0
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   4
      Left            =   105
      TabIndex        =   25
      Top             =   2295
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
      Picture         =   "frmAdjustment.frx":1685
      CaptionAlign    =   0
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   405
      Index           =   1
      Left            =   105
      TabIndex        =   26
      Top             =   1875
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
      Picture         =   "frmAdjustment.frx":1DFF
      CaptionAlign    =   0
   End
End
Attribute VB_Name = "frmAdjustment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  Programmed By  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  Rosalyn Lazo Descallar  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  Date Started  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  August 28, 2007  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'

Option Explicit

Private WithEvents oDriver As FormDriver
Attribute oDriver.VB_VarHelpID = -1
Private oSkin As FormSkin
Private bLoaded As Boolean
Private oRS As New ADODB.Recordset

Dim txtfieldGotfocus As Boolean
Dim pbnewitem As Boolean
Dim psSelected() As String

Dim pnindex As Integer
Dim pnCtr As Integer
Dim Time As String

Private Sub cmdButton_Click(Index As Integer)
Dim lsSearch As String
Dim lsCancel As Integer
Dim lsSQL As String
Dim lnrow As Long

   Select Case Index
      Case 0 'save
         oDriver.RecordSave
      Case 1 'search
         SearchBarCode
      Case 2 'new
         oDriver.RecordNew
         InitButton xeModeAddNew
      Case 3 'cancel
         oDriver.RecordCancelUpdate
         InitButton xeModeReady
      Case 4 'close
         Unload Me
      End Select
End Sub

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0
   
   If bLoaded = False Then
      oDriver.RecordNew
      oDriver.DisableTextbox 0
      bLoaded = True
   End If
   
End Sub

Private Sub oDriver_DisableOtherControl()
   oDriver.DisableTextbox 0
End Sub

Private Sub oDriver_EnableOtherControl()
   oDriver.DisableTextbox 0
End Sub

Private Sub oDriver_InitValue()
   oDriver.FieldReference(0) = True
   If Not oDriver.SetValue(0, getNextCode("CP_Inventory_Adjustment", "sTransNox", True, oApp.Connection, True, oApp.BranchCode)) Then Exit Sub
   pbnewitem = True
   
   txtfield(1).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
   
   For pnCtr = 0 To 4
      txtothers(pnCtr).Text = ""
   Next
   oDriver.FieldValue(1) = Date
   Combo1.Text = "In"
   txtothers(0).SetFocus
   
End Sub

Private Sub Form_Load()
Dim lsSQL As String

   CenterChildForm mdiMain, Me
   bLoaded = False
    
   Set oDriver = New FormDriver
   Set oDriver.AppDriver = oApp
   Set oDriver.MainForm = Me
   
   InitButton xeModeReady
   
   Set oSkin = New FormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransaction

   oDriver.RecQuery = "SELECT" _
                           & " sTransNox ," _
                           & " dTransact ," _
                           & " sStockIDx ," _
                           & " nQuantity ," _
                           & " sRemarksx ," _
                           & " sApproved ," _
                           & " cTranStat ," _
                           & " sModified ," _
                           & " dModified ," _
                           & " vTimeStmp  " _
                    & " FROM CP_Inventory_Adjustment " _
   
   oDriver.BrowseQuery = "SELECT" _
                     & " Distinct " _
                     & " a.sTransNox, " _
                     & " a.dTransact, " _
                     & " b.sBarrcode, " _
                     & " b.sDescript  " _
               & " FROM CP_Inventory_Adjustment a " _
                  & " LEFT JOIN CP_Inventory b " _
                     & " ON a.sStockIDx = b.sStockIDx " _
                  & " LEFT JOIN Brand c " _
                     & " ON b.sBrandIDx = c.sBrandIDx " _
                  & " LEFT JOIN Model d " _
                     & " ON b.sModelIDx = d.sModelIDx " _
                  & " LEFT JOIN Color e " _
                     & " ON b.sColorIdx = e.sColorIDx " _
                     
   oDriver.InitRecForm
   
   oDriver.BrowseFTitle(0) = "Transaction No"
   oDriver.BrowseFTitle(1) = "Date"
   oDriver.BrowseFTitle(2) = "Bar Code"
   oDriver.BrowseFTitle(3) = "Description"
   
   oDriver.BrowseFFormat(1) = "MMMM dd, yyyy"
       
   oDriver.FieldFormat(0) = "@@-@@@@@@@@"
   oDriver.FieldSize(0) = Len(oDriver.FieldFormat(0))
   
   oDriver.FieldFormat(1) = "MMMM DD, YYYY"
   oDriver.FieldStart = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oDriver = Nothing
   Set oSkin = Nothing
End Sub

Private Sub SearchBarCode()
Dim lsSQL As String
Dim lsCondition As String
Dim lsSearch As String
Dim lnCtr As Integer
   
   lsSQL = "SELECT" _
          & " a.sBarrcode, " _
          & " b.sBrandNme, " _
          & " c.sModelNme, " _
          & " a.sDescript, " _
          & " d.sColorNme, " _
          & " a.sStockIDx  " _
      & " FROM CP_Inventory a " _
          & " LEFT JOIN Brand b " _
            & " ON a.sBrandIdx = b.sBrandIdx " _
          & " LEFT JOIN Model c " _
            & " ON a.sModelIdx = c.sModelIdx " _
          & " LEFT JOIN Color d " _
            & " ON a.sColorIDx = d.sColorIDx " _
      & " WHERE a.sBarrcode like  '%" & txtothers(0).Text & "%' " _
         & " AND a.scategidx in('01005','01006') "
   
   If oRS.State = adStateOpen Then oRS.Close
   oRS.Open lsSQL, oApp.Connection, adOpenKeyset, adLockOptimistic, adCmdText
   
   If Not oRS.EOF Then
      If oRS.RecordCount = 1 Then
         oDriver.FieldValue(2) = oRS(5)
         txtothers(0).Text = oRS(0)
         For lnCtr = 1 To 4
            txtothers(lnCtr).Text = oRS(lnCtr)
            txtothers(lnCtr).Locked = True
         Next
      Else
         lsSearch = KwikSearch(oApp, lsSQL, _
                    "sBarrcode»sBrandNme»sModelNme»sDescript»sColorNme", _
                    "Bar Code»Brand»Model»Description»Color")
         If lsSearch <> "" Then
            psSelected = Split(lsSearch, "»")
            oDriver.FieldValue(2) = psSelected(5)
            txtothers(0).Text = psSelected(0)
            For lnCtr = 1 To 4
               txtothers(lnCtr).Text = IIf(IsNull(psSelected(lnCtr)), "", psSelected(lnCtr))
               txtothers(lnCtr).Locked = True
            Next
         End If
      End If
      Combo1.SetFocus
   Else
      txtothers(0).Tag = ""
      For lnCtr = 0 To 4
         txtothers(lnCtr).Text = ""
      Next
      txtothers(0).SetFocus
   End If
   Set oRS = Nothing
   
End Sub

Private Sub InitButton(lnStat As Integer)
   Dim lbShow As Boolean

   lbShow = IIf(lnStat = 0, False, True)
   cmdButton(0).Visible = lbShow
   cmdButton(1).Visible = lbShow
   cmdButton(3).Visible = lbShow
   
   cmdButton(2).Visible = Not lbShow
   cmdButton(4).Visible = Not lbShow
   
End Sub

Private Sub oDriver_LoadOtherData()
   txtothers(1).Locked = True
End Sub

Private Sub oDriver_Save(Saved As Boolean)
   Saved = False
End Sub

Private Sub oDriver_SaveComplete()
   MsgBox "Record Successfully Saved!!!", vbInformation, "Information"
   InitButton xeModeReady
End Sub

Private Sub oDriver_WillSave(Cancel As Boolean)
   If oDriver.FieldValue(1) = "" Then
      MsgBox "Invalid Date Detected!!!", vbCritical, "Warning"
      txtfield(1).SetFocus
      Cancel = True
   ElseIf oDriver.FieldValue(2) = "" Then
      MsgBox "Invalid Bar Code Detected!!!", vbCritical, "Warning"
      txtothers(0).SetFocus
      Cancel = True
   ElseIf oDriver.FieldValue(3) = "" Or oDriver.FieldValue(3) = 0# Then
      MsgBox "Invalid Quantity Detected!!!", vbCritical, "Warning"
      txtfield(3).SetFocus
      Cancel = True
   ElseIf oDriver.FieldValue(4) = "" Then
      MsgBox "Invalid Remarks Detected!!!", vbCritical, "Warning"
      txtfield(4).SetFocus
      Cancel = True
   ElseIf oDriver.FieldValue(5) = "" Then
      MsgBox "Invalid Approving Officer Detected!!!", vbCritical, "Warning"
      txtfield(5).SetFocus
      Cancel = True
   Else
      Time = Format(Now, "hh:nn:ss AM/PM")
      Cancel = Not Update_CP_Inventory
         If Cancel Then Exit Sub
      oDriver.FieldValue(1) = CDate(txtfield(1).Text) & " " & Time
      Select Case Combo1.Text
         Case "In"
            oDriver.FieldValue(6) = 0 'TranStat
         Case "Out"
            oDriver.FieldValue(6) = 1 'TranStat
      End Select
   End If
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   oDriver.ColumnIndex = Index
   txtfieldGotfocus = True
   pnindex = Index
   txtfield(Index).BackColor = &HE1FEFF
End Sub

Private Sub txtothers_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Or KeyCode = 13 Then
      If Index = 0 Then
         SearchBarCode
      End If
      KeyCode = 0
   End If
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   txtfield(Index).BackColor = &HFFFFFF
End Sub

Private Sub txtfield_Validate(Index As Integer, Cancel As Boolean)

   Select Case Index
      Case 1
         If Not IsDate(txtfield(Index).Text) Then
            txtfield(Index).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
         Else
            txtfield(Index).Text = Format(txtfield(Index).Text, "MMMM DD, YYYY")
         End If
      Case 3
         If Not IsNumeric(txtfield(Index).Text) Then
            txtfield(Index).Text = 0#
         End If
   End Select
   Cancel = Not oDriver.ValidateField(Index)
   txtfield(Index).Text = TitleCase(txtfield(Index).Text)
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

Private Function Update_CP_Inventory() As Boolean
Dim lsSQL As String
Dim lnrow As Long
Dim lnEntry As Integer
Dim QOH As Integer
Dim QIn As Integer
Dim QOut As Integer
Dim lrs As ADODB.Recordset

Update_CP_Inventory = True
On Error GoTo errProc
               
      'Get last Entry No.
      lnEntry = getEntryNo("CP_Inventory_Ledger", "'" & oDriver.FieldValue(2) & "'", "'" & oApp.BranchCode & "'")
           
      'Get QOH
      Select Case Combo1.Text
         Case "In"
            QOH = getQuantity("'" & oDriver.FieldValue(2) & "'", "'" & oApp.BranchCode & "'") + txtfield(3).Text
            QIn = txtfield(3).Text
            QOut = 0
         Case "Out"
            QOH = getQuantity("'" & oDriver.FieldValue(2) & "'", "'" & oApp.BranchCode & "'") - txtfield(3).Text
            QIn = 0
            QOut = txtfield(3).Text
      End Select
      
      'Add Record, CP_Inventroy_Ledger
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
            & "('" & oDriver.FieldValue(2) & "', " _
            & "'" & oApp.BranchCode & " ', " _
            & "'" & oApp.BranchCode & " ', " _
            & "'CPAj' , " _
            & "'" & oDriver.FieldValue(0) & "', " _
            & "'" & CLng(QIn) & "', " _
            & "'" & CLng(QOut) & "', " _
            & "'" & CLng(QOH) & "', " _
            & "'" & lnEntry & "', " _
            & "'" & CDate(oDriver.FieldValue(1)) & " " & Time & "', " _
            & " getdate())"
      oApp.Connection.Execute lsSQL, lnrow, adCmdText
   
      'Update QOH, CP_Inventory_Master
      lsSQL = "UPDATE CP_Inventory_Master SET" _
            & " nQtyOnHnd = '" & CLng(QOH) & "', " _
            & " sModified = '" & Encrypt(oApp.UserID) & "', " _
            & " dModified = getdate() " _
      & " WHERE sStockIDx = '" & oDriver.FieldValue(2) & "' " _
            & " And sBranchCd = '" & oApp.BranchCode & "' "
      oApp.Connection.Execute lsSQL, lnrow, adCmdText
      
      Set lrs = Nothing

endProc:
   Exit Function
errProc:
   Update_CP_Inventory = False
   MsgBox Err.Description, vbCritical, "Warning"
End Function

Private Sub txtothers_GotFocus(Index As Integer)
   txtothers(Index).BackColor = &HE1FEFF
   If txtothers(Index).Text <> "" Then
      txtothers(Index).SelStart = 0
      txtothers(Index).SelLength = Len(txtothers(Index).Text)
   End If
End Sub

Private Sub txtothers_LostFocus(Index As Integer)
   txtothers(Index).BackColor = &HFFFFFF
End Sub

Private Sub txtothers_Validate(Index As Integer, Cancel As Boolean)
   txtothers(Index).Text = TitleCase(txtothers(Index).Text)
   If Index = 0 Then
      If oDriver.FieldValue(2) = "" Then
         MsgBox "Invalid Barrcode!!!", vbInformation, "Notice"
         txtothers(Index).Text = ""
         txtothers(Index).SetFocus
      End If
   End If
End Sub

'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  Date Tested  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  August 28, 2007  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'

'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤    Version 1    ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  Date Finished  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  August 29, 2007  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'



