VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmCP_Serial 
   BorderStyle     =   0  'None
   Caption         =   "CP Serial Status"
   ClientHeight    =   4905
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6615
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4905
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   3330
      Index           =   0
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   5874
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "frmCP_Serial.frx":0000
         Left            =   4530
         List            =   "frmCP_Serial.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1020
         Width           =   1725
      End
      Begin VB.CheckBox chkServUnit 
         Caption         =   "Service Unit"
         Height          =   195
         Left            =   4530
         TabIndex        =   23
         Tag             =   "et0;fb0"
         Top             =   1800
         Width           =   1725
      End
      Begin VB.CheckBox chkSoldStat 
         Caption         =   "Sold"
         Height          =   195
         Left            =   4530
         TabIndex        =   22
         Tag             =   "et0;fb0"
         Top             =   1575
         Width           =   1725
      End
      Begin VB.TextBox txtOther 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   1050
         TabIndex        =   13
         Top             =   2190
         Width           =   2595
      End
      Begin VB.TextBox txtOther 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   1050
         TabIndex        =   11
         Top             =   1890
         Width           =   2595
      End
      Begin VB.TextBox txtOther 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   1050
         TabIndex        =   9
         Top             =   1590
         Width           =   2595
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   1050
         TabIndex        =   15
         Top             =   2490
         Width           =   5205
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmCP_Serial.frx":0004
         Left            =   4530
         List            =   "frmCP_Serial.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   690
         Width           =   1725
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   1050
         TabIndex        =   17
         Top             =   2790
         Width           =   5205
      End
      Begin VB.TextBox txtOther 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1050
         TabIndex        =   7
         Top             =   1290
         Width           =   2595
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   1050
         TabIndex        =   5
         Top             =   990
         Width           =   2595
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   1050
         TabIndex        =   3
         Top             =   690
         Width           =   2595
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   1050
         TabIndex        =   1
         Top             =   225
         Width           =   1920
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Type"
         Height          =   195
         Index           =   8
         Left            =   3705
         TabIndex        =   20
         Top             =   1050
         Width           =   690
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Color"
         Height          =   195
         Index           =   7
         Left            =   135
         TabIndex        =   12
         Top             =   2220
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
         Height          =   195
         Index           =   6
         Left            =   135
         TabIndex        =   10
         Top             =   1920
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Brand"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   8
         Top             =   1620
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   14
         Top             =   2520
         Width           =   645
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Location"
         Height          =   195
         Index           =   22
         Left            =   3705
         TabIndex        =   18
         Top             =   720
         Width           =   645
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Branch"
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   16
         Top             =   2820
         Width           =   660
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Barrcode"
         Height          =   195
         Index           =   10
         Left            =   135
         TabIndex        =   4
         Top             =   1005
         Width           =   780
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Serial No"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   2
         Top             =   720
         Width           =   780
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   6
         Top             =   1320
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Serial ID"
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
         Index           =   0
         Left            =   135
         TabIndex        =   0
         Top             =   270
         Width           =   765
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   285
         Index           =   0
         Left            =   1155
         Tag             =   "et0;ht2"
         Top             =   330
         Width           =   1920
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   0
      Left            =   5775
      TabIndex        =   31
      Top             =   4110
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
      Picture         =   "frmCP_Serial.frx":0008
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   1
      Left            =   4995
      TabIndex        =   30
      Top             =   4110
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
      Picture         =   "frmCP_Serial.frx":0782
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   2
      Left            =   2655
      TabIndex        =   25
      Top             =   4110
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
      Picture         =   "frmCP_Serial.frx":0EFC
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   4
      Left            =   1875
      TabIndex        =   24
      Top             =   4110
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
      Picture         =   "frmCP_Serial.frx":1676
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   5
      Left            =   5775
      TabIndex        =   32
      Top             =   4110
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
      Picture         =   "frmCP_Serial.frx":1DF0
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   705
      Index           =   7
      Left            =   3435
      TabIndex        =   26
      Top             =   4110
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
      Picture         =   "frmCP_Serial.frx":256A
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   3
      Left            =   2655
      TabIndex        =   27
      Top             =   4110
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
      Picture         =   "frmCP_Serial.frx":2CE4
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   6
      Left            =   3435
      TabIndex        =   28
      Top             =   4110
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
      Picture         =   "frmCP_Serial.frx":345E
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   8
      Left            =   4215
      TabIndex        =   29
      Top             =   4110
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
      Picture         =   "frmCP_Serial.frx":3BD8
      PicturePos      =   1
   End
End
Attribute VB_Name = "frmCP_Serial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmCPSerial"

Private WithEvents oDriver As clsFormDriver
Attribute oDriver.VB_VarHelpID = -1
Private oForm As frmCPSerialLedger
Private oSkin As clsFormSkin
Private bLoaded As Boolean
Private oFormReg As Object

Dim psStockIDx As String
Dim psBranchCd As String
Dim psSupplier As String

Dim pbLoadRecord As Boolean
Dim pnCtr As Integer, pnIndex As Integer

Private Sub cmdButton_Click(Index As Integer)
   Dim lsSearch As String
   Dim lsSelected() As String
   Dim lsSQL As String
   Dim lsOldProc As String
   
   lsOldProc = "cmdButton_Click"
   '''On Error GoTo errProc

   Select Case Index
   Case 0
      oDriver.RecordCancelUpdate
   Case 1
      oDriver.BrowseRecord
   Case 2
      oDriver.RecordSave
      If oDriver.FieldValue(0) <> "" Then
         Call SaveSerialCost
      End If
   Case 3
      If Trim(oDriver.FieldValue(0)) <> "" Then oDriver.RecordUpdate
   Case 5
      Unload Me
   Case 6
      MsgBox "Unable to Delete Record" & vbCrLf & _
             "Deleting Record is prohibited!!!", vbCritical, "Warning"
   Case 7
      oDriver.RecordSearch
      txtField(pnIndex).SetFocus
   Case 8
      If pbLoadRecord Then
         oForm.SerialID = oDriver.FieldValue(0)
         oForm.Caption = "CP Serial Ledger"
         oForm.txtField(0).Text = txtField(0).Text
         oForm.txtField(1).Text = txtField(3).Text
         oForm.txtField(2).Text = txtField(4).Text
         oForm.txtField(3).Text = txtField(2).Text
         oForm.Show 1
         
'         If oForm.SystemID <> "" And oForm.TransactionNo <> "" Then
'            Select Case oForm.SystemID
'            Case "CPDv"
'               'Set oFormReg = New frmMCDeliveryMaintenance
'            Case "CPDA"
'               'Set oFormReg = New frmDAMaintenance
'            Case "CPPR"
'               'Set oFormReg = New frmMCPOReturnReg
'            Case "CPBT"
'               'Set oFormReg = New frmBDMaintenance
'            Case Else
'               Exit Sub
'            End Select
'
'            oFormReg.TransactionNo = oForm.TransactionNo
'            oFormReg.Show
'         End If
      Else
         MsgBox "Unable to Load Serial Ledger!!!" & vbCrLf & _
                "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
      End If
   End Select

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Sub Combo1_GotFocus()
   With Combo1
      .BackColor = oApp.getColor("HT1")
   End With
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
      SetNextFocus
      KeyCode = 0
   End If
End Sub

Private Sub Combo1_LostFocus()
   With Combo1
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub Form_Activate()
   Dim lsOldProc As String
   
   lsOldProc = "Form_Activate"
   '''On Error GoTo errProc
   
   oApp.MenuName = Me.Tag
   Me.ZOrder 0
   
   If bLoaded Then
      oDriver.RecordNew
      bLoaded = False
   End If
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String
   
   lsOldProc = "Form_Load"
   '''On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oDriver = New clsFormDriver
   Set oDriver.AppDriver = oApp
   Set oDriver.MainForm = Me

   Set oForm = New frmCPSerialLedger
      
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin

   oDriver.RecQuery = "SELECT" _
                        & "  sSerialID" _
                        & ", sSerialNo" _
                        & ", sSupplier" _
                        & ", sStockIDx" _
                        & ", sBranchCd" _
                        & ", cLocation" _
                        & ", cSoldStat" _
                        & ", cUnitType" _
                        & ", cUnitClas" _
                        & ", sModified" _
                        & ", dModified" _
                     & " FROM CP_Inventory_Serial"
   
   oDriver.BrowseQuery = "SELECT Distinct" _
                        & "  a.sSerialID" _
                        & ", a.sSerialNo" _
                        & ", c.sBrandNme" _
                        & ", d.sModelNme" _
                        & ", e.sColorNme" _
                        & ", f.sCompnyNm" _
                     & " FROM CP_Inventory_Serial a" _
                        & " LEFT JOIN Client_Master f" _
                           & " ON a.sSupplier = f.sClientID" _
                        & ", CP_Inventory b" _
                           & " LEFT JOIN CP_Brand c" _
                              & " ON b.sBrandIDx = c.sBrandIDx" _
                           & " LEFT JOIN CP_Model d" _
                              & " ON b.sModelIDx = d.sModelIDx" _
                           & " LEFT JOIN Color e" _
                              & " ON b.sColorIDx = e.sColorIDx" _
                     & " WHERE a.sStockIDx = b.sStockIDx" _
                     & " ORDER BY a.sSerialID"

   oDriver.InitRecForm
   
   oDriver.BrowseFReference(0) = True
   oDriver.BrowseFTitle(0) = "Serial ID"
   oDriver.BrowseFTitle(1) = "Serial No"
   oDriver.BrowseFTitle(2) = "Brand Name"
   oDriver.BrowseFTitle(3) = "Model Name"
   oDriver.BrowseFTitle(4) = "Color Name"
   oDriver.BrowseFTitle(5) = "Supplier"
   
   oDriver.LookupQuery(2) = "SELECT" _
                              & "  a.sClientID" _
                              & ", a.sCompnyNm" _
                           & " FROM Client_Master a" _
                              & ", CP_Supplier b" _
                           & " WHERE a.sClientID = b.sClientID" _
                              & " AND b.sBranchCd = " & strParm(oApp.BranchCode) _
                              & " AND b.cRecdStat = " & strParm(xeRecStateActive) _
                           & " ORDER BY a.sCompnyNm"
   
   oDriver.LookupReference(2) = "a.sClientID蒼.sCompnyNm"
   oDriver.LookupColumn(2) = "sCompnyNm"
   oDriver.LookupTitle(2) = "Supplier"
   
   oDriver.LookupQuery(3) = "SELECT" _
                              & "  a.sStockIDx" _
                              & ", a.sBarrCode" _
                              & ", b.sBrandNme" _
                              & ", c.sModelNme" _
                           & " FROM CP_Inventory a" _
                              & " LEFT JOIN CP_Brand b" _
                                 & " ON a.sBrandIDx = b.sBrandIDx" _
                              & " LEFT JOIN CP_Model c" _
                                 & " ON a.sModelIDx = c.sModelIDx" _
                           & " ORDER BY a.sBarrCode"
                        
   oDriver.LookupReference(3) = "a.sStockIDx製BarrCode蓑.sBrandNme蓊.sModelNme"
   oDriver.LookupColumn(3) = "sBarrCode製BrandNme製ModelNme"
   oDriver.LookupTitle(3) = "BarrCode翡rand Name膂odel Name"
   
   oDriver.LookupQuery(4) = "SELECT" _
                              & "  sBranchCd" _
                              & ", sBranchNm" _
                           & " FROM Branch" _
                           & " ORDER BY sBranchNm"
   
   oDriver.LookupReference(4) = "sBranchCd製BranchNm"
   oDriver.LookupColumn(4) = "sBranchNm"
   oDriver.LookupTitle(4) = "Branch Name"
   
   oDriver.FieldFormat(0) = "@@@@-@@@@@@"
   oDriver.FieldSize(0) = Len(oDriver.FieldFormat(0))
   
   oDriver.FieldStart = 1

   Combo1.ListIndex = -1
   Combo1.List(0) = "Warehouse"
   Combo1.List(1) = "Branch"
   Combo1.List(2) = "Supplier"
   Combo1.List(3) = "Customer"
   Combo1.List(4) = "Unknown"
   Combo1.List(5) = "Service Center"
   Combo1.List(6) = "Service Unit"
   
   Combo2.ListIndex = -1
   Combo2.List(0) = "LDU"
   Combo2.List(1) = "Regular"
   Combo2.List(2) = "Free"
   Combo2.List(3) = "Live"
   Combo2.List(4) = "Service"
   Combo2.List(5) = "RDU"
   Combo2.List(6) = "Others"

   
   psStockIDx = ""
   psBranchCd = ""
   psSupplier = ""
   
   bLoaded = True

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oDriver = Nothing
   Set oSkin = Nothing
   Set oForm = Nothing
End Sub

Private Sub oDriver_DisableOtherControl()
   Combo1.Enabled = False
End Sub

Private Sub oDriver_EnableOtherControl()
   oDriver.DisableTextbox 0
   Combo1.Enabled = True
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

Private Sub oDriver_InitValue()
   If oDriver.SetValue(0, GetNextCode("CP_Inventory_Serial", "sSerialID", True, oApp.Connection, True, oApp.BranchCode)) = False Then Exit Sub
   oDriver.FieldValue(5) = xeLocBranch
   oDriver.FieldValue(6) = xeNo
   oDriver.FieldValue(7) = 1
   oDriver.FieldValue(8) = 0
   
   Combo1.ListIndex = oDriver.FieldValue(5)
   Combo2.ListIndex = oDriver.FieldValue(7)
   
   chkSoldStat.Value = oDriver.FieldValue(6)
   chkServUnit.Value = oDriver.FieldValue(8)
   
   txtOther(0).Text = txtOther(0).Tag
   txtOther(1).Text = txtOther(1).Tag
   txtOther(2).Text = txtOther(2).Tag
   txtOther(3).Text = txtOther(3).Tag
   
   oDriver.FieldValue(2) = psSupplier
   oDriver.FieldValue(3) = psStockIDx
   oDriver.FieldValue(4) = psBranchCd
   
   txtField(2).Text = txtField(2).Tag
   txtField(3).Text = txtField(3).Tag
   txtField(4).Text = txtField(4).Tag
   
   pbLoadRecord = False
End Sub

Private Sub oDriver_LoadOtherData()
   Dim lsOldProc As String
   
   lsOldProc = "oDriver_LoadOtherData"
   '''On Error GoTo errProc

   Dim lrs As ADODB.Recordset
   
   For pnCtr = 0 To txtOther.Count - 1
      txtOther(pnCtr).Text = ""
      txtOther(pnCtr).Tag = ""
   Next
   
   Combo1.ListIndex = IIf(IsNull(oDriver.FieldValue(5)), -1, IIf(Trim(oDriver.FieldValue(5)) = "", -1, oDriver.FieldValue(5)))
   chkSoldStat.Value = oDriver.FieldValue(6)
   Combo2.ListIndex = oDriver.FieldValue(7)
   chkServUnit.Value = oDriver.FieldValue(8)
   
   Set lrs = New ADODB.Recordset
   lrs.Open "SELECT" _
               & "  a.sDescript" _
               & ", b.sBrandNme" _
               & ", c.sModelNme" _
               & ", d.sColorNme" _
            & " FROM CP_Inventory a" _
               & " LEFT JOIN CP_Brand b" _
                  & " ON a.sBrandIDx = b.sBrandIDx" _
               & " LEFT JOIN CP_Model c" _
                  & " ON a.sModelIDx = c.sModelIDx" _
               & " LEFT JOIN Color d" _
                  & " ON a.sColorIDx = d.sColorIDx" _
            & " WHERE a.sStockIDx = " & strParm(oDriver.FieldValue(3)) _
            , oApp.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
   
   If Not lrs.EOF Then
      txtOther(0).Text = lrs("sDescript")
      txtOther(1).Text = IFNull(lrs("sBrandNme"), "")
      txtOther(2).Text = IFNull(lrs("sModelNme"), "")
      txtOther(3).Text = IFNull(lrs("sColorNme"), "")
      
      txtOther(0).Tag = txtOther(0).Text
      txtOther(1).Tag = txtOther(1).Text
      txtOther(2).Tag = txtOther(2).Text
      txtOther(3).Tag = txtOther(3).Text
   End If
   pbLoadRecord = True

endProc:
   Set lrs = Nothing
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )"
End Sub

Private Sub oDriver_WillSave(Cancel As Boolean)
   Dim lnCtr As Integer
   Dim lsOldProc As String
   
   lsOldProc = "oDriver_WillSave"
   '''On Error GoTo errProc
   
   If Combo1.ListIndex <> -1 Then oDriver.FieldValue(5) = CStr(Combo1.ListIndex)
   oDriver.FieldValue(6) = chkSoldStat.Value
   If Combo2.ListIndex <> -1 Then oDriver.FieldValue(7) = CStr(Combo2.ListIndex)
   oDriver.FieldValue(8) = chkServUnit.Value
   psSupplier = IFNull(oDriver.FieldValue(2), "")
   psStockIDx = oDriver.FieldValue(3)
   psBranchCd = oDriver.FieldValue(4)
   
   txtField(2).Tag = txtField(2).Text
   txtField(3).Tag = txtField(3).Text
   txtField(4).Tag = txtField(4).Text
   
   txtOther(0).Tag = txtOther(0).Text
   txtOther(1).Tag = txtOther(1).Text
   txtOther(2).Tag = txtOther(2).Text
   txtOther(3).Tag = txtOther(3).Text
   
endProc:
   Exit Sub
errProc:
   Cancel = True
   ShowError lsOldProc & "( " & Cancel & " )"
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = oApp.getColor("HT1")
   End With
   oDriver.ColumnIndex = Index
   pnIndex = Index
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String
   
   lsOldProc = "txtField_KeyDown"
   '''On Error GoTo errProc
   
   If KeyCode = vbKeyReturn Or KeyCode = vbKeyF3 Then
      With txtField(Index)
         If KeyCode = vbKeyF3 Then
            oDriver.RecordSearch .Text
            If .Text <> "" Then SetNextFocus
         Else
            If .Text <> "" Then oDriver.RecordSearch .Text
         End If
      End With
      KeyCode = 0
   End If
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " _
                       & "  " & Index _
                       & ", " & KeyCode _
                       & ", " & Shift _
                       & " )", True
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   Dim lsOldProc As String
   Dim lors As Recordset
   
   lsOldProc = "txtField_Validate"
   '''On Error GoTo errProc
   
   With txtField(Index)
      If Trim(.Text) <> "" Then .Text = UCase(Left(.Text, 1)) & Right(.Text, Len(.Text) - 1)
      
      Select Case Index
      Case 1
         .Text = UCase(.Text)
         Cancel = Not oDriver.ValidateField(Index)
      Case Else
         Cancel = Not oDriver.ValidateField(Index)
         If Index = 3 Then
            Set lors = New Recordset
            lors.Open "SELECT" _
                        & "  a.sDescript" _
                        & ", b.sBrandNme" _
                        & ", c.sModelNme" _
                        & ", d.sColorNme" _
                     & " FROM CP_Inventory a" _
                        & " LEFT JOIN CP_Brand b" _
                           & " ON a.sBrandIDx = b.sBrandIDx" _
                        & " LEFT JOIN CP_Model c" _
                           & " ON a.sModelIDx = c.sModelIDx" _
                        & " LEFT JOIN Color d" _
                           & " ON a.sColorIDx = d.sColorIDx" _
                     & " WHERE sStockIDx = " & strParm(IFNull(oDriver.FieldValue(3), "")) _
            , oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
            
            txtOther(0).Text = ""
            txtOther(1).Text = ""
            txtOther(2).Text = ""
            txtOther(3).Text = ""
            If Not lors.EOF Then
               txtOther(0).Text = lors("sDescript")
               txtOther(1).Text = IFNull(lors("sBrandNme"), "")
               txtOther(2).Text = IFNull(lors("sModelNme"), "")
               txtOther(3).Text = IFNull(lors("sColorNme"), "")
            End If
            Set lors = Nothing
         End If
      End Select
   End With
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " _
                       & "  " & Index _
                       & ", " & Cancel _
                       & " )", True
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

Private Sub SaveSerialCost()
   Dim lsSQL As String
   Dim loData As Recordset
   
   Set loData = New Recordset
   loData.Open "SELECT sSerialID " & _
               " FROM CP_Serial_Cost" & _
               " WHERE sSerialID = " & strParm(oDriver.FieldValue(0)) _
   , oApp.Connection, , , adCmdText
   
   If loData.EOF Then
      lsSQL = "INSERT INTO CP_Serial_Cost" & _
               " SET sSerialID = " & strParm(oDriver.FieldValue(0)) & _
               ", nPurPrice = 0 " & _
               ", dPricexxx = '2016-01-17' "
      
      oApp.Execute lsSQL, "CP_Serial_Cost"
   End If
End Sub
