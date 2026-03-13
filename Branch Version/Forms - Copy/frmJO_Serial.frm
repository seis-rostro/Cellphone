VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmJO_Serial 
   BorderStyle     =   0  'None
   Caption         =   "New Serial"
   ClientHeight    =   4365
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6945
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   ScaleHeight     =   4365
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1770
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   3122
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   1035
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   585
         Width           =   2280
      End
      Begin VB.TextBox txtothers 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   0
         Left            =   1035
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   840
         Width           =   2280
      End
      Begin VB.TextBox txtothers 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   1035
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   1095
         Width           =   2280
      End
      Begin VB.TextBox txtothers 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   2
         Left            =   4155
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   1095
         Width           =   2430
      End
      Begin VB.TextBox txtothers 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   3
         Left            =   1035
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   1350
         Width           =   5550
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   0
         Left            =   1035
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   150
         Width           =   2280
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "IMEI NO."
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   608
         Width           =   1140
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   240
         Left            =   1080
         Tag             =   "et0;ht2"
         Top             =   180
         Width           =   2280
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bar Code"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   4
         Top             =   870
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Brand"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   6
         Top             =   1125
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
         Height          =   195
         Index           =   7
         Left            =   3360
         TabIndex        =   8
         Top             =   1125
         Width           =   1005
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   10
         Top             =   1365
         Width           =   825
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Serial ID"
         Height          =   195
         Index           =   14
         Left            =   120
         TabIndex        =   0
         Top             =   180
         Width           =   645
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   0
      Left            =   4560
      TabIndex        =   20
      Top             =   3585
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
      Picture         =   "frmJO_Serial.frx":0000
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   2
      Left            =   6090
      TabIndex        =   22
      Top             =   3585
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
      Picture         =   "frmJO_Serial.frx":077A
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   705
      Index           =   1
      Left            =   5325
      TabIndex        =   21
      Top             =   3585
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
      Picture         =   "frmJO_Serial.frx":0EF4
      PicturePos      =   1
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   1035
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   2340
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   1826
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtothers 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   6
         Left            =   4140
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   630
         Width           =   2445
      End
      Begin VB.TextBox txtothers 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   5
         Left            =   1035
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   630
         Width           =   2310
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   2
         Left            =   1035
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   120
         Width           =   5550
      End
      Begin VB.TextBox txtothers 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   4
         Left            =   1035
         MultiLine       =   -1  'True
         TabIndex        =   15
         Text            =   "frmJO_Serial.frx":166E
         Top             =   375
         Width           =   5550
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contact #"
         Height          =   195
         Index           =   2
         Left            =   3360
         TabIndex        =   18
         Top             =   660
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Town"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   16
         Top             =   645
         Width           =   960
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
         Height          =   225
         Index           =   11
         Left            =   120
         TabIndex        =   12
         Top             =   135
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   225
         Index           =   12
         Left            =   120
         TabIndex        =   14
         Top             =   375
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmJO_Serial"
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
Dim pnCtr As Integer
Dim pnindex As Integer
Dim txtfieldGotfocus As Boolean
Dim txtOthersGotfocus As Boolean

Private Sub cmdButton_Click(Index As Integer)
   Dim lsSearch As String
   Dim lsRep As Integer
   
   Select Case Index
   Case 0   'Save
      oDriver.RecordSave
   Case 1   'Search
      If pnindex = 2 Then oDriver.RecordSearch txtField(2).Text
   Case 2   'cancel
      Unload Me
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
      txtothers(0).SetFocus
      txtField(1).Text = frmJobOrder.txtField(2).Text
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
                        & " sClientID, " _
                        & " sStockIDx, " _
                        & " sBranchCd, " _
                        & " cSoldStat, " _
                        & " cLocation, " _
                        & " sModified, " _
                        & " dModified, " _
                        & " vTimeStmp  " _
                    & " FROM CP_Serial_Master "
                                     
   oDriver.InitRecForm
   
   'Customer
   oDriver.LookupQuery(2) = "SELECT" _
                  & " a.sClientID, " _
                  & " a.sLastName + ', ' + a.sFrstName + ' ' + a.sMiddName as xFullName, " _
                  & " a.sAddressx + ', ' + b.sTownName as xAddressx " _
               & " FROM Client_Master a " _
                  & " LEFT JOIN TownCity b " _
                     & " ON a.sTownIDxx = b.sTownIDxx " _
               & " ORDER BY slastname, sfrstname, smiddname "

   oDriver.LookupReference(2) = "sClientID»xFullName»xAddressx"
   oDriver.LookupColumn(2) = "xFullName»xAddressx"
   oDriver.LookupTitle(2) = "Customer Name»Address"
       
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
   ClearFields

End Sub

Private Sub ClearFields()
   For pnCtr = 0 To txtothers.Count - 1
     txtothers(pnCtr).Text = ""
   Next
End Sub

Private Sub oDriver_LoadOtherData()
   Dim lsSQL As String
   Dim lnCtr As Integer
   
   pbnewitem = False
   
   If oDriver.FieldValue(3) <> "" Then
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
         & " Where a.sStockIDx = '" & oDriver.FieldValue(3) & "' "
      If oRS.State = adStateOpen Then oRS.Close
      oRS.Open lsSQL, oApp.Connection, adOpenKeyset, adLockOptimistic, adCmdText
   
      For lnCtr = 0 To 3
         txtothers(lnCtr).Text = oRS(lnCtr)
         txtothers(lnCtr).Enabled = False
      Next
      txtothers(0).Enabled = True
      Set oRS = Nothing
   End If
      
End Sub
Private Sub Search_Client(ByVal SearchValue As Boolean)
Dim lsSearch As String
Dim lrs As ADODB.Recordset
Dim lsSQL As String

   Set lrs = New ADODB.Recordset
   lsSQL = "SELECT" _
               & " a.sClientID, " _
               & " a.sLastName + ', ' + a.sFrstName + ' ' + a.sMiddName xFullName, " _
               & " a.sAddressx, " _
               & " b.sTownName, " _
               & " a.sPhoneNox, " _
               & " a.sAddressx + ', ' + b.sTownName as xAddressx " _
            & " FROM Client_Master a " _
               & " LEFT JOIN TownCity b " _
                  & " ON a.sTownIDxx = b.sTownIDxx " _

   If SearchValue Then
      lsSQL = lsSQL & " WHERE sLastName + ', ' + sFrstName + ' ' + sMiddName = '" & txtField(2).Text & "'"
   Else
      lsSQL = lsSQL & " WHERE sLastName + ', ' + sFrstName + ' ' + sMiddName LIKE '" & txtField(2).Text & "%' "
   End If
   lsSQL = lsSQL & " ORDER BY sLastName + ', ' + sFrstName + ' ' + sMiddName"
   lrs.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   
   If lrs.RecordCount = 1 Then
      oDriver.FieldValue(2) = lrs("sClientID")
      txtField(2).Text = lrs("xFullName")
      txtothers(4).Text = IIf(IsNull(lrs("sAddressx")), "", lrs("sAddressx"))
      txtothers(5).Text = IIf(IsNull(lrs("sTownName")), "", lrs("sTownName"))
      txtothers(6).Text = IIf(IsNull(lrs("sPhoneNox")), "", lrs("sPhoneNox"))
   ElseIf lrs.RecordCount > 1 Then
        lsSearch = KwikBrowse(oApp, lrs, _
                          "sClientID»" _
                        & "xFullName»" _
                     & "xAddressx", _
                          "Client ID»" _
                        & "Name»" _
                        & "Address")

        If lsSearch <> "" Then
            psSelected = Split(lsSearch, "»")
            oDriver.FieldValue(2) = psSelected(0)
            txtField(2).Text = lrs("xFullName")
            txtothers(4).Text = IIf(IsNull(psSelected(2)), "", psSelected(2))
            txtothers(5).Text = IIf(IsNull(psSelected(3)), "", psSelected(3))
            txtothers(6).Text = IIf(IsNull(psSelected(4)), "", psSelected(4))
        End If
   Else
      frmCustomer.Show 1
   End If
   Set lrs = Nothing

End Sub
Private Sub oDriver_Save(Saved As Boolean)
   Saved = False
End Sub

Private Sub oDriver_SaveComplete()
   MsgBox "IMEI No Added!!!", vbInformation, "Information"
   ClearFields
   Unload Me
End Sub

Private Sub oDriver_WillSave(Cancel As Boolean)
   If oDriver.FieldValue(1) = "" Then
      MsgBox "Invalid IMEI No. Detected!!!", vbCritical, "Warning"
      txtField(1).SetFocus
      Cancel = True
   ElseIf oDriver.FieldValue(3) = "" Then
      MsgBox "Invalid Unit Detected!!!", vbCritical, "Warning"
      txtothers(0).SetFocus
      Cancel = True
   Else
      If pbnewitem Then
         Cancel = Not Save_CP_Serial_Ledger
            If Cancel Then Exit Sub
      End If
      oDriver.FieldValue(4) = oApp.BranchCode
      oDriver.FieldValue(5) = 1
      oDriver.FieldValue(6) = 2
   End If
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   oDriver.ColumnIndex = Index
   txtfieldGotfocus = True
   pnindex = Index
   txtField(Index).BackColor = &HE1FEFF
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Or KeyCode = vbKeyF3 Then
      If Index = 2 Then Search_Client False
      If txtField(Index).Text <> "" Then SetNextFocus
   End If
   KeyCode = 0
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   txtfieldGotfocus = False
   txtField(Index).BackColor = &HFFFFFF
End Sub

Private Sub txtfield_Validate(Index As Integer, Cancel As Boolean)
'If Index = 2 Then
'   If oDriver.FieldValue(2) = "" Then frmCustomer.Show 1
'End If
oDriver.FieldValue(1) = txtField(1).Text
End Sub

Private Sub txtothers_GotFocus(Index As Integer)
   txtOthersGotfocus = True
   pnindex = Index
   txtOthersGotfocus = False
   txtothers(Index).BackColor = &HE1FEFF
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
   End Select
End Sub

Private Sub oDriver_Delete(Deleted As Boolean)
   Deleted = True
End Sub

Private Sub txtothers_LostFocus(Index As Integer)
   txtOthersGotfocus = False
   txtothers(Index).BackColor = &HFFFFFF
End Sub

Private Function Save_CP_Serial_Ledger() As Boolean
Dim lsSQL As String
Dim lnrow As Long
   
Save_CP_Serial_Ledger = True
On Error GoTo errProc
      
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
         & "('" & oDriver.FieldValue(0) & "', " _
         & "'" & oApp.BranchCode & "', " _
         & "'" & oApp.ServerDate & "', " _
         & "'1', " _
         & "'CPJO', " _
         & "'99000001', " _
         & "'1'," _
         & "'2', " _
         & " getdate())"
   oApp.Connection.Execute lsSQL, lnrow, adCmdText
         
   If lnrow <= 0 Then
      MsgBox "Unable to Insert CP_Serial_Ledger!!!", vbCritical, "Warning"
      Save_CP_Serial_Ledger = False
      GoTo endProc
   End If

endProc:
   Exit Function
errProc:
   Save_CP_Serial_Ledger = False
   MsgBox Err.Description, vbCritical, "Warning"
End Function

Private Sub SearchBarCode(ByVal SearchValue As Boolean)
Dim lsSQL As String
Dim lsSearch As String
Dim lnCtr As Integer
   
   lsSQL = "SELECT" _
          & " a.sBarrcode, " _
          & " b.sBrandNme, " _
          & " c.sModelNme, " _
          & " a.sDescript, " _
          & " a.sStockIDx, " _
          & " d.sColorNme, " _
          & " a.sCategIDx  " _
      & " FROM CP_Inventory a " _
          & " LEFT JOIN Brand b " _
            & " ON a.sBrandIdx = b.sBrandIdx " _
          & " LEFT JOIN Model c " _
            & " ON a.sModelIdx = c.sModelIdx " _
          & " LEFT JOIN Color d " _
            & " ON a.sColorIDx = d.sColorIDx " _
      & " WHERE (sCategIDx = '01001' or sCategIDx = '01002' or sCategIDx = '01002') " _
         & " AND a.sBarrCode LIKE '" & txtothers(0).Text & "%' " _
      & " ORDER BY sBarrCode"
   If oRS.State = adStateOpen Then oRS.Close
   oRS.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   If oRS.RecordCount = 1 Then
      For lnCtr = 0 To 3
         txtothers(lnCtr).Text = oRS(lnCtr)
      Next
      oDriver.FieldValue(3) = oRS(4)

   ElseIf oRS.RecordCount > 1 Then
        lsSearch = KwikBrowse(oApp, oRS, _
               "sBarrcode»sBrandNme»sModelNme»sDescript»sColorNme", _
               "Bar Code»Brand»Model»Description»Color")
                        
        If lsSearch <> "" Then
            psSelected = Split(lsSearch, "»")
            For lnCtr = 0 To 3
               txtothers(lnCtr).Text = psSelected(lnCtr)
            Next
            oDriver.FieldValue(3) = psSelected(4)
        End If
   Else
      frmJO_Inventory.Show 1
   End If
   
   For pnCtr = 1 To txtothers.Count - 1
     txtothers(pnCtr).Enabled = False
   Next

   Set oRS = Nothing

End Sub

