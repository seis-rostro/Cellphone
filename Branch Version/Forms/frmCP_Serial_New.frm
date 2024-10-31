VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmCP_Serial_New 
   BorderStyle     =   0  'None
   Caption         =   "Unit IMEI No."
   ClientHeight    =   4545
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4545
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame2 
      Height          =   555
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   979
      BackColor       =   12632256
      ClipControls    =   0   'False
      BorderStyle     =   1
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   0
         Left            =   1485
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   105
         Width           =   1920
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Serial ID"
         Height          =   195
         Index           =   1
         Left            =   225
         TabIndex        =   0
         Top             =   150
         Width           =   1140
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   240
         Left            =   1530
         Tag             =   "et0;ht2"
         Top             =   150
         Width           =   1920
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   2385
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   1125
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   4207
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   1485
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   435
         Width           =   2835
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   3
         Left            =   1485
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   1980
         Width           =   2835
      End
      Begin VB.TextBox txtothers 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   0
         Left            =   1485
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   180
         Width           =   2835
      End
      Begin VB.TextBox txtothers 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   2
         Left            =   1485
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   1470
         Width           =   2835
      End
      Begin VB.TextBox txtothers 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   1485
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   1215
         Width           =   2835
      End
      Begin VB.TextBox txtothers 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   3
         Left            =   1485
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   1725
         Width           =   2835
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Info"
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
         Index           =   16
         Left            =   210
         TabIndex        =   6
         Top             =   840
         Width           =   750
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   210
         X2              =   4310
         Y1              =   1065
         Y2              =   1065
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "IMEI No."
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   4
         Top             =   465
         Width           =   780
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Branch"
         Height          =   195
         Index           =   5
         Left            =   210
         TabIndex        =   13
         Top             =   2010
         Width           =   1140
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bar Code"
         Height          =   195
         Index           =   2
         Left            =   210
         TabIndex        =   2
         Top             =   210
         Width           =   780
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
         Height          =   195
         Index           =   4
         Left            =   210
         TabIndex        =   9
         Top             =   1500
         Width           =   1140
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Brand"
         Height          =   195
         Index           =   10
         Left            =   210
         TabIndex        =   7
         Top             =   1245
         Width           =   780
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   3
         Left            =   210
         TabIndex        =   11
         Top             =   1755
         Width           =   1140
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   4
      Left            =   105
      TabIndex        =   15
      Top             =   3750
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmCP_Serial_New.frx":0000
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   2
      Left            =   870
      TabIndex        =   16
      Top             =   3750
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmCP_Serial_New.frx":077A
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   3
      Left            =   870
      TabIndex        =   17
      Top             =   3750
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmCP_Serial_New.frx":0EF4
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   705
      Index           =   7
      Left            =   1635
      TabIndex        =   18
      Top             =   3750
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmCP_Serial_New.frx":166E
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   6
      Left            =   1635
      TabIndex        =   19
      Top             =   3750
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmCP_Serial_New.frx":1DE8
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   1
      Left            =   3165
      TabIndex        =   21
      Top             =   3750
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmCP_Serial_New.frx":2562
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   8
      Left            =   2400
      TabIndex        =   20
      Top             =   3750
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
      Caption         =   "&Ledger"
      AccessKey       =   "L"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCP_Serial_New.frx":2CDC
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   0
      Left            =   3930
      TabIndex        =   22
      Top             =   3750
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmCP_Serial_New.frx":3456
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   5
      Left            =   3930
      TabIndex        =   23
      Top             =   3750
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmCP_Serial_New.frx":3BD0
      PicturePos      =   1
   End
End
Attribute VB_Name = "frmCP_Serial_New"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents oDriver As FormDriver
Attribute oDriver.VB_VarHelpID = -1
Private oSkin As FormSkin
Private bLoaded As Boolean
Private oRS As New ADODB.Recordset

Dim pbnewitem As Boolean
Dim psSelected() As String
Dim pnctr As Integer
Dim pnindex As Integer
Dim txtfieldGotfocus As Boolean
Dim txtOthersGotfocus As Boolean

Dim pnUserRights As Integer
Dim psUserID As String
Dim psUserName As String

Private Sub cmdButton_Click(Index As Integer)
   Dim lsApproval As Integer
   Dim lsSQL As String
   
   Select Case Index
   Case 0   'cancel
      oDriver.RecordCancelUpdate
   Case 1   'browse
      oDriver.BrowseRecord
   Case 2   'save
      If oApp.UserLevel = xeEncoder Or oApp.UserLevel = xeSupervisor Then
         lsApproval = MsgBox("Seek for Approval?", vbQuestion + vbYesNo, "Confirm")
         If lsApproval = vbYes Then
            If Not GetApproval(oApp, pnUserRights, psUserID, psUserName) Then Exit Sub
            If pnUserRights < xeManager Then
               MsgBox "Approving User is Not Authorized!!!", vbCritical, "Warning"
               Exit Sub
            Else
                oDriver.RecordSave
            End If
         End If
      Else
          oDriver.RecordSave
      End If
     
   Case 3   'update
      lsSQL = "SELECT" _
             & " sSerialID " _
         & " FROM CP_Serial_Ledger " _
      & " Where sSerialID = '" & oDriver.FieldValue(0) & "' "
      If oRS.State = adStateOpen Then oRS.Close
      oRS.Open lsSQL, oApp.Connection, adOpenKeyset, adLockOptimistic, adCmdText

      If oRS.EOF = False Then
         MsgBox "IMEI No. has Other Transactions!!!" & vbCrLf & _
            "Update Not Permitted!!!", vbCritical, "Warning"
         Exit Sub
      End If
      
      If oApp.UserLevel = xeEncoder Or oApp.UserLevel = xeSupervisor Then
         lsApproval = MsgBox("Seek for Approval?", vbQuestion + vbYesNo, "Confirm")
         If lsApproval = vbYes Then
            If Not GetApproval(oApp, pnUserRights, psUserID, psUserName) Then Exit Sub
            If pnUserRights < xeManager Then
               MsgBox "Approving User is Not Authorized!!!", vbCritical, "Warning"
               Exit Sub
            Else
               oDriver.RecordUpdate
            End If
         End If
      Else
         oDriver.RecordUpdate
      End If
   Case 4   'new
      oDriver.RecordNew
   Case 5   'close
      Unload Me
   Case 6   'delete
      MsgBox "Delete Not Permitted!!!" & vbCrLf & vbCrLf & _
      "Please Notify ROSALYN LAZO DESCALLAR" & vbCrLf & _
      "for Assistance!!!", vbCritical, "Warning"
'      oDriver.RecordDelete
   Case 7   'search
      If txtOthersGotfocus And pnindex = 0 Then SearchBarCode False
 
   Case 8   'ledger
      If Not pbnewitem Then
         frmCP_Serial_Ledger.BarrCode = oDriver.FieldValue(0)
         frmCP_Serial_Ledger.Show 1
      Else
         MsgBox "Unable to Load IMEI Ledger!!!" & vbCrLf & _
                "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
      End If
   End Select
End Sub

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0
   
   If bLoaded = False Then
      oDriver.RecordNew
      oDriver.DisableTextbox 0
      ClearFields
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
   oSkin.ApplySkin

   oDriver.RecQuery = "SELECT" _
                        & " sSerialID, " _
                        & " sIMEINoxx, " _
                        & " sStockIDx, " _
                        & " sBranchCd, " _
                        & " cSoldStat, " _
                        & " cLocation, " _
                        & " sClientID, " _
                        & " sModified, " _
                        & " dModified, " _
                        & " vTimeStmp  " _
                    & " FROM CP_Serial_Master "
                                     
   oDriver.BrowseQuery = "SELECT" _
                  & " a.sIMEINoxx, " _
                  & " b.sBarrCode, " _
                  & " b.sDescript, " _
                  & " c.sBrandNme, " _
                  & " d.sModelNme, " _
                  & " e.sBranchNm, " _
                  & " f.sColorNme  " _
            & " FROM CP_Serial_Master a " _
               & " LEFT JOIN CP_Inventory b " _
                  & " ON a.sStockIDx = b.sStockIDx " _
               & " LEFT JOIN Brand c " _
                  & " ON b.sBrandIDx = c.sBrandIDx " _
               & " LEFT JOIN Model d " _
                  & " ON b.sModelIDx = d.sModelIDx " _
               & " LEFT JOIN Branch e " _
                  & " ON a.sBranchCd = e.sBranchCd " _
               & " LEFT JOIN Color f " _
                  & " ON b.sColorIdx = f.sColorIDx " _
            & "ORDER BY a.sIMEINoxx "

   oDriver.InitRecForm

   oDriver.BrowseColumn(0) = "sIMEINoxx"
   oDriver.BrowseColumn(1) = "sBarrCode"
   oDriver.BrowseColumn(2) = "sBrandNme"
   oDriver.BrowseColumn(3) = "sModelNme"
   oDriver.BrowseColumn(4) = "sDescript"
   oDriver.BrowseColumn(5) = "sColorNme"
   oDriver.BrowseColumn(6) = "sBranchNm"
   
   oDriver.BrowseFTitle(0) = "IMEI No."
   oDriver.BrowseFTitle(1) = "Bar Code"
   oDriver.BrowseFTitle(2) = "Brand"
   oDriver.BrowseFTitle(3) = "Model"
   oDriver.BrowseFTitle(4) = "Description"
   oDriver.BrowseFTitle(5) = "Color"
   oDriver.BrowseFTitle(6) = "Branch"
   
   'Branch
   oDriver.LookupQuery(3) = "SELECT" _
                     & " a.sBranchCd, " _
                     & " a.sBranchNm, " _
                     & " a.sAddressx + ' ' + b.sTownName xAddressx " _
               & " FROM Branch a " _
                  & "LEFT JOIN TownCity b " _
                     & " ON a.sTownIdxx = b.sTownIDxx " _
               & " WHERE a.cRecdStat = '" & xeRecStateActive & "'" _
               & " ORDER BY sBranchNm "

   oDriver.LookupReference(3) = "sBranchCd»sBranchNm»xAddressx"
   oDriver.LookupColumn(3) = "sBranchNm»xAddressx"
   oDriver.LookupTitle(3) = "Branch»Address"
       
   oDriver.FieldFormat(0) = "@@-@@@@@@@@"
   oDriver.FieldSize(0) = Len(oDriver.FieldFormat(0))
   
   oDriver.FieldStart = 1

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oDriver = Nothing
   Set oSkin = Nothing
End Sub

Private Sub oDriver_DisableOtherControl()
   oDriver.DisableTextbox 0
End Sub

Private Sub oDriver_EnableOtherControl()
   oDriver.DisableTextbox 0
End Sub

Private Sub oDriver_InitValue()

   oDriver.FieldReference(0) = True
   If Not oDriver.SetValue(0, getNextCode("CP_Serial_Master", "sSerialID", True, oApp.Connection, True, oApp.BranchCode)) Then Exit Sub
   pbnewitem = True
  
   oDriver.FieldValue(3) = oApp.BranchCode
   txtfield(3).Text = "Main Office"
   ClearFields

End Sub

Private Sub ClearFields()
   For pnctr = 0 To txtothers.Count - 1
     txtothers(pnctr).Text = ""
   Next
End Sub

Private Sub oDriver_LoadOtherData()
   Dim lsSQL As String
   Dim lnctr As Integer
   
   pbnewitem = False
   
   lsSQL = "SELECT" _
             & " a.sBarrCode, " _
             & " b.sBrandNme, " _
             & " c.sModelNme,  " _
             & " a.sDescript  " _
       & " FROM CP_Inventory a " _
          & " LEFT JOIN Brand b " _
             & " ON a.sBrandIDx = b.sBrandIDx " _
          & " LEFT JOIN Model c " _
             & " ON a.sModelIDx = c.sModelIDx " _
      & " Where a.sStockIDx = '" & oDriver.FieldValue(2) & "' "
   If oRS.State = adStateOpen Then oRS.Close
   oRS.Open lsSQL, oApp.Connection, adOpenKeyset, adLockOptimistic, adCmdText
   
   For lnctr = 0 To 3
      txtothers(lnctr).Text = IIf(IsNull(oRS(lnctr)), "", oRS(lnctr))
      txtothers(lnctr).Enabled = False
   Next
   txtothers(0).Enabled = True
   Set oRS = Nothing
   
End Sub

Private Sub oDriver_Save(Saved As Boolean)
   Saved = False
End Sub

Private Sub oDriver_SaveComplete()
   ClearFields
End Sub

Private Sub oDriver_WillSave(Cancel As Boolean)
   If oDriver.FieldValue(1) = "" Then
      MsgBox "Invalid IMEI No. Detected!!!", vbCritical, "Warning"
      txtfield(1).SetFocus
      Cancel = True
   ElseIf oDriver.FieldValue(2) = "" Then
      MsgBox "Invalid Bar Code Detected!!!", vbCritical, "Warning"
      txtfield(2).SetFocus
      Cancel = True
   Else
      oDriver.FieldValue(4) = 0
      oDriver.FieldValue(5) = 1
   End If
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   oDriver.ColumnIndex = Index
   txtfieldGotfocus = True
   pnindex = Index
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Or KeyCode = 13 Then
      If Index = 3 Then
         oDriver.RecordSearch txtfield(Index).Text
         If txtfield(Index).Text <> "" Then SetNextFocus
      End If
      KeyCode = 0
   End If
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   txtfieldGotfocus = False
End Sub

Private Sub txtfield_Validate(Index As Integer, Cancel As Boolean)

If txtfield(1).Text <> "" Then oDriver.FieldValue(1) = txtfield(1).Text
Cancel = Not oDriver.ValidateField(Index)
End Sub

Private Sub txtothers_GotFocus(Index As Integer)
   txtOthersGotfocus = True
   If txtothers(Index).Text <> "" Then
      txtothers(Index).SelStart = 0
      txtothers(Index).SelLength = Len(txtothers(Index).Text)
   End If
   txtfieldGotfocus = False
   pnindex = Index
End Sub

Private Sub txtothers_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Or KeyCode = 13 Then
      If Index = 0 Then
         SearchBarCode False
         If txtothers(Index).Text <> "" Then SetNextFocus
      End If
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
         Call Modified("CP_Serial_Master", "sSerialID = '" & oDriver.FieldValue(0) & "' ")
   End Select
End Sub

Private Sub oDriver_Delete(Deleted As Boolean)
   Deleted = True
End Sub

Private Sub txtothers_LostFocus(Index As Integer)
   txtOthersGotfocus = False
End Sub

Private Sub SearchBarCode(ByVal SearchValue As Boolean)
Dim lsSQL As String
Dim lsSearch As String
Dim lnctr As Integer
   
   lsSQL = "SELECT" _
          & " a.sBarrcode, " _
          & " b.sBrandNme, " _
          & " c.sModelNme, " _
          & " a.sDescript, " _
          & " a.sStockIDx, " _
          & " d.sColorNme, " _
          & " a.cWdSerial  " _
      & " FROM CP_Inventory a " _
          & " LEFT JOIN Brand b " _
            & " ON a.sBrandIdx = b.sBrandIdx " _
          & " LEFT JOIN Model c " _
            & " ON a.sModelIdx = c.sModelIdx " _
          & " LEFT JOIN Color d " _
            & " ON a.sColorIDx = d.sColorIDx " _
      & " WHERE cWdSerial = 1 " _
         & " AND a.sBarrCode LIKE '" & txtothers(0).Text & "%' " _
      & " ORDER BY sBarrCode"
   If oRS.State = adStateOpen Then oRS.Close
   oRS.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   If oRS.RecordCount = 1 Then
      For lnctr = 0 To 3
         txtothers(lnctr).Text = oRS(lnctr)
      Next
      oDriver.FieldValue(2) = oRS(4)
      
   ElseIf oRS.RecordCount > 1 Then
        lsSearch = KwikBrowse(oApp, oRS, _
               "sBarrcode»sBrandNme»sModelNme»sDescript»sColorNme", _
               "Bar Code»Brand»Model»Description»Color")
                        
        If lsSearch <> "" Then
            psSelected = Split(lsSearch, "»")
            For lnctr = 0 To 3
               txtothers(lnctr).Text = psSelected(lnctr)
            Next
            oDriver.FieldValue(2) = psSelected(4)
        End If
   Else
      For lnctr = 0 To 3
         txtothers(lnctr).Text = ""
      Next
      MsgBox "Record Not Existing!!!", vbInformation, "Information"
   End If
   
   For pnctr = 1 To txtothers.Count - 1
     txtothers(pnctr).Enabled = False
   Next

   Set oRS = Nothing

End Sub



