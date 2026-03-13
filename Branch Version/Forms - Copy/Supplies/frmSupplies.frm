VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmSupplies 
   BorderStyle     =   0  'None
   Caption         =   "Supplies"
   ClientHeight    =   6375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9555
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6375
   ScaleWidth      =   9555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin xrControl.xrFrame xrFrame2 
      Height          =   4350
      Left            =   60
      Tag             =   "wt0;fb0"
      Top             =   1065
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   7673
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   13
         Left            =   1380
         TabIndex        =   40
         Top             =   1935
         Width           =   2010
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   12
         Left            =   5760
         TabIndex        =   11
         Top             =   2625
         Width           =   2010
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   7
         Left            =   5760
         TabIndex        =   9
         Top             =   1920
         Width           =   2910
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   11
         Left            =   1335
         TabIndex        =   14
         Top             =   3570
         Width           =   2115
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   10
         Left            =   1335
         TabIndex        =   13
         Text            =   "Text 1"
         Top             =   3255
         Width           =   2115
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   9
         Left            =   1335
         TabIndex        =   12
         Top             =   2940
         Width           =   2115
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   8
         Left            =   1335
         TabIndex        =   10
         Top             =   2625
         Width           =   3240
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   6
         Left            =   5760
         TabIndex        =   8
         Top             =   1605
         Width           =   2910
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   5760
         TabIndex        =   7
         Top             =   1290
         Width           =   2910
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   5760
         TabIndex        =   6
         Top             =   975
         Width           =   2910
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   1380
         TabIndex        =   5
         Top             =   1605
         Width           =   2010
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   1380
         TabIndex        =   4
         Top             =   1290
         Width           =   3150
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   1380
         TabIndex        =   3
         Top             =   975
         Width           =   3165
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   1395
         TabIndex        =   16
         Top             =   375
         Width           =   2010
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Branch Code"
         Height          =   195
         Index           =   13
         Left            =   420
         TabIndex        =   41
         Top             =   1980
         Width           =   930
      End
      Begin VB.Shape Shape2 
         Height          =   4095
         Index           =   3
         Left            =   120
         Top             =   75
         Width           =   9150
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qty on Hnd"
         Height          =   195
         Index           =   12
         Left            =   4785
         TabIndex        =   29
         Top             =   2670
         Width           =   810
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total "
         Height          =   195
         Index           =   11
         Left            =   390
         TabIndex        =   28
         Top             =   3615
         Width           =   405
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "End Qty"
         Height          =   195
         Index           =   10
         Left            =   390
         TabIndex        =   27
         Top             =   3300
         Width           =   570
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Beg. Qty"
         Height          =   195
         Index           =   9
         Left            =   375
         TabIndex        =   26
         Top             =   2985
         Width           =   615
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Beg. Inv"
         Height          =   195
         Index           =   8
         Left            =   375
         TabIndex        =   25
         Top             =   2670
         Width           =   600
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cost Date"
         Height          =   195
         Index           =   7
         Left            =   4785
         TabIndex        =   24
         Top             =   1965
         Width           =   705
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ave. Purc"
         Height          =   195
         Index           =   6
         Left            =   4785
         TabIndex        =   23
         Top             =   1650
         Width           =   705
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Purc"
         Height          =   195
         Index           =   5
         Left            =   4785
         TabIndex        =   22
         Top             =   1335
         Width           =   675
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Purc. Price"
         Height          =   195
         Index           =   4
         Left            =   4785
         TabIndex        =   21
         Top             =   1020
         Width           =   780
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Measure ID"
         Height          =   195
         Index           =   3
         Left            =   420
         TabIndex        =   20
         Top             =   1650
         Width           =   825
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier"
         Height          =   195
         Index           =   2
         Left            =   420
         TabIndex        =   19
         Top             =   1335
         Width           =   570
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   1
         Left            =   405
         TabIndex        =   18
         Top             =   1020
         Width           =   795
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stock ID"
         Height          =   195
         Index           =   0
         Left            =   420
         TabIndex        =   17
         Top             =   405
         Width           =   630
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   285
         Left            =   1500
         Tag             =   "et0;ht2"
         Top             =   495
         Width           =   2070
      End
      Begin VB.Line Line1 
         X1              =   4710
         X2              =   4710
         Y1              =   270
         Y2              =   2325
      End
      Begin VB.Shape Shape2 
         Height          =   1695
         Index           =   2
         Left            =   255
         Top             =   2385
         Width           =   8865
      End
      Begin VB.Shape Shape2 
         Height          =   2085
         Index           =   1
         Left            =   255
         Top             =   240
         Width           =   8865
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   480
      Index           =   1
      Left            =   60
      Tag             =   "wt0;fb0"
      Top             =   525
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   847
      Begin VB.TextBox txtOthers 
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
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   0
         Top             =   90
         Width           =   1680
      End
      Begin VB.TextBox txtOthers 
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
         Height          =   285
         Index           =   1
         Left            =   4350
         TabIndex        =   2
         Top             =   90
         Width           =   4905
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Stock ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   195
         TabIndex        =   15
         Top             =   135
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Branch Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   9
         Left            =   3150
         TabIndex        =   1
         Top             =   105
         Width           =   1590
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   0
      Left            =   8730
      TabIndex        =   30
      Top             =   5595
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
      Picture         =   "frmSupplies.frx":0000
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   1
      Left            =   7950
      TabIndex        =   31
      Top             =   5595
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
      Picture         =   "frmSupplies.frx":077A
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   2
      Left            =   5610
      TabIndex        =   32
      Top             =   5595
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
      Picture         =   "frmSupplies.frx":0EF4
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   4
      Left            =   5610
      TabIndex        =   33
      Top             =   5595
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
      Picture         =   "frmSupplies.frx":166E
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   5
      Left            =   8730
      TabIndex        =   34
      Top             =   5595
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
      Picture         =   "frmSupplies.frx":1DE8
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   705
      Index           =   7
      Left            =   6390
      TabIndex        =   35
      Top             =   5595
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
      Picture         =   "frmSupplies.frx":2562
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   3
      Left            =   6390
      TabIndex        =   36
      Top             =   5595
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
      Picture         =   "frmSupplies.frx":2CDC
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   6
      Left            =   5610
      TabIndex        =   37
      Top             =   5595
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
      Picture         =   "frmSupplies.frx":3456
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   8
      Left            =   7170
      TabIndex        =   38
      Top             =   5595
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
      Picture         =   "frmSupplies.frx":3BD0
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   9
      Left            =   5610
      TabIndex        =   39
      Top             =   5595
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
      Caption         =   "&Replace"
      AccessKey       =   "R"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmSupplies.frx":434A
      PicturePos      =   1
   End
End
Attribute VB_Name = "frmSupplies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Private Const pxeMODULENAME = "frmSupplies"
'
'Private WithEvents oTrans As clsSupplies
'Private oSkin As clsFormSkin
'Private bLoaded As Boolean
'Private oRS As New ADODB.Recordset
'
'Dim pbtxtOthers As Boolean
'Dim pnCtr As Integer, pnIndex As Integer
'
'Dim psCPInventory As String
'Dim pbEnblButtons As Boolean
'Dim pbNewInvntory As Boolean
'
'Private Sub cmdButton_Click(Index As Integer)
'   Dim lsSearch As String
'   Dim lnRep As Integer
'   Dim lsOldProc As String
'
'   lsOldProc = "cmdButton_Click"
'   'On Error GoTo errProc
'
'   Select Case Index
'   Case 0 'cancel
'      oTrans.RecordCancelUpdate
'      pbEnblButtons = False
'   Case 1 'browse
'      oTrans.OpenRecord
'   Case 2 'save
'      oTrans.SaveRecord
'   Case 3 'update
'      oTrans.RecordUpdate
'   Case 4 'new
'      oTrans.NewRecord
'   Case 5 'close
'      Unload Me
'   Case 6 'delete
'      oTrans.RecordDelete
'   Case 7 'search
'      If pbtxtOthers Then
'         oTrans.SearchRecord
'         txtField(pnIndex).SetFocus
'      End If
'   Case 8 'ledger
'      If Not pbNewInvntory Then
'        With frmCP_MatrixLedger
'            .txtDateFrom = Format(DateAdd("m", -1, oApp.ServerDate), "MMMM DD, YYYY")
'            .txtDateThru = Format(oApp.ServerDate, "MMMM DD, YYYY")
'            .txtField(0) = txtField(0)
'            .txtField(1) = txtField(1)
'            .txtField(2) = txtField(2)
'            .txtField(3) = txtField(3)
'
'            .StockID = oTrans.FieldValue(5)
'            .Show 1
'         End With
'      Else
'         MsgBox "Unable to Load Inventory Ledger!!!" & vbCrLf & _
'                "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
'      End If
'   End Select
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " & Index & " )", True
'End Sub
'
'Private Sub Form_Activate()
'   Dim lsOldProc As String
'
'   lsOldProc = "Form_Activate"
'   'On Error GoTo errProc
'
'   oApp.MenuName = Me.Tag
'   Me.ZOrder 0
'
''   If bLoaded = False Then
''      oTrans.RecordCancelUpdate
''      oTrans_InitValue
''      bLoaded = True
''      txtOthers(0).SetFocus
''   End If
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " & " )", True
'End Sub
'
'Private Sub Form_Load()
'   Dim lsSQL As String
'   Dim lsOldProc As String
'
'   lsOldProc = "Form_Load"
'   'On Error GoTo errProc
'
'   CenterChildForm mdiMain, Me
'
'   bLoaded = False
'   Set oRS = New ADODB.Recordset
'
'   Set oTrans = New clsSupplies
'   Set oTrans.AppDriver = oApp
'
'
'   Set oSkin = New clsFormSkin
'   Set oSkin.AppDriver = oApp
'   Set oSkin.Form = Me
'   oSkin.ApplySkin
'
'   ClearFields
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " & " )", True
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'   Set oTrans = Nothing
'   Set oSkin = Nothing
'   Set oRS = Nothing
'End Sub
'
'Private Sub oTrans_DisableOtherControl()
'   For pnCtr = 0 To txtOthers.Count - 1
'      txtOthers(pnCtr).Enabled = False
'   Next
'
'   txtOthers(0).Enabled = True
'   txtOthers(1).Enabled = True
'
'   oTrans.hideButton 6
'   oTrans.hideButton 9
'End Sub
'
'Private Sub oTrans_EnableOtherControl()
'   For pnCtr = 0 To txtOthers.Count - 1
'      Select Case pnCtr
'      Case 8, 9
'         txtOthers(pnCtr).Enabled = False
'      Case Else
'         txtOthers(pnCtr).Enabled = True
'      End Select
'   Next
'
'   If oTrans.EditMode = xeModeUpdate Then
'      txtOthers(0).Enabled = pbEnblButtons
'      txtOthers(1).Enabled = pbEnblButtons
'
'   End If
'End Sub
'
'Private Sub InitOthers()
'   For pnCtr = 0 To txtOthers.Count - 3
'      Select Case pnCtr
'      Case 0 To 3, 5, 6, 7
'         oRS(pnCtr) = 0
'         txtOthers(pnCtr).Text = Format(oRS(pnCtr), "#,##0.00")
'      Case 4
'         oRS(pnCtr) = oApp.ServerDate
'         txtOthers(pnCtr).Text = Format(oRS(pnCtr), "MMM DD, YYYY")
'      End Select
'   Next
'
'   oRS("cRecdStat") = xeRecStateActive
'   oRS("sStockIDx") = oTrans.FieldValue(5)
'   oRS("sBranchCd") = oApp.BranchCode
'   oRS("nFloatAmt") = 0
'   oRS("nLedgerNo") = 0
'End Sub
'
'Private Sub oTrans_InitValue()
'   Dim lsOldProc As String
'
'   lsOldProc = "oTrans_InitValue"
'   'On Error GoTo errProc
'
'   oTrans.FieldReference(0) = True
'   oTrans.FieldValue(0) = NewBarrCode
'   txtField(0).Text = oTrans.FieldValue(0)
'   oTrans.FieldValue(5) = GetNextCode("CP_Load_Matrix", "sStockIDx", True, oApp.Connection, True, oApp.BranchCode)
'   oTrans.FieldValue(6) = xeRecStateActive
'
'   For pnCtr = 0 To txtOthers.Count - 1
'      Select Case pnCtr
'      Case 0 To 3, 5, 6, 7
'         txtOthers(pnCtr).Text = 0
'      Case 4
'         txtOthers(pnCtr).Text = Format(oApp.ServerDate, "MMM DD, YYYY")
'      End Select
'      txtOthers(pnCtr).Tag = ""
'   Next
'
'   txtOthers(0).Locked = False
'   txtOthers(1).Locked = False
'
'   oRS.AddNew
'   InitOthers
'   pbEnblButtons = True
'   pbNewInvntory = True
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " & " )"
'End Sub
'
''Private Sub oTrans_LoadOtherData()
''   Dim lsOldProc As String
''   Dim lsSQL As String
''
''   lsOldProc = "oTrans_LoadOtherData"
''   'On Error GoTo errProc
''
''   Set oRS = New ADODB.Recordset
''   lsSQL = AddCondition(psCPInventory, "sStockIDx = " & strParm(oTrans.FieldValue(5)) _
''                                    & " AND sBranchCd = " & strParm(oApp.BranchCode)) _
''                                    & " AND cRecdStat = " & strParm(xeRecStateActive)
''   oRS.Open lsSQL, oApp.Connection, adOpenStatic, adLockOptimistic, adCmdText
''   Set oRS.ActiveConnection = Nothing
''
''   If oRS.EOF Then
''      oRS.AddNew
''      InitOthers
''   Else
''      For pnCtr = 0 To txtOthers.Count - 1
''         Select Case pnCtr
''         Case 0 To 3, 5F, 6, 7
''            txtOthers(pnCtr).Text = Format(IIf(IsNull(oRS(pnCtr)), 0, oRS(pnCtr)), "#,##0.00")
''         Case 4
''            txtOthers(pnCtr).Text = IIf(IsNull(oRS(pnCtr)), "", Format(oRS(pnCtr), "MMM DD, YYYY"))
''         End Select
''      Next
''      pbNewInvntory = False
''   End If
''
''   txtOthers(8).Text = oTrans.FieldValue(0)
''   txtOthers(8).Tag = txtOthers(8).Text
''
''   txtOthers(9).Text = oTrans.FieldValue(1)
''   txtOthers(9).Tag = txtOthers(9).Text
''
''endProc:
''   Exit Sub
''errProc:
''   ShowError lsOldProc & "( " & " )"
''End Sub
'
'Private Sub oTrans_Save(Saved As Boolean)
'   Saved = False
'End Sub
'
'Private Sub oTrans_WillSave(Cancel As Boolean)
'   Dim lsOldProc As String
'   Dim lnCtr As Integer
'
'   lsOldProc = "oTrans_WillSave"
'   'On Error GoTo errProc
'
'   If oTrans.FieldValue(0) = "" Then
'      MsgBox "Invalid BarrCode detected!!!", vbCritical, "Warning"
'      txtField(0).SetFocus
'      Cancel = True
'   ElseIf oTrans.FieldValue(1) = "" Then
'      MsgBox "Invalid Description detected!!!", vbCritical, "Warning"
'      txtField(1).SetFocus
'      Cancel = True
'   ElseIf oTrans.FieldValue(5) = "" Then
'      MsgBox "Invalid Stock ID Detected!!!" & vbCrLf & _
'               "Please contact GMC_SEG for assistant!!!", vbCritical, "Warning"
'      Cancel = True
'   Else
'      Cancel = Not UpdateCPInventory
'      If pbNewInvntory Then Cancel = Not SaveCPInventoryLedger
'   End If
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " & Cancel & " )"
'End Sub
''
''Private Sub txtField_GotFocus(Index As Integer)
''   With txtField(Index)
''      .SelStart = 0
''      .SelLength = Len(.Text)
''      .BackColor = oApp.getColor("HT1")
''   End With
''
''   oTrans.ColumnIndex = Index
''   pbtxtOthers = False
''   pnIndex = Index
''End Sub
'
'Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'   Dim lsOldProc As String
'
'   lsOldProc = "txtField_KeyDown"
'   'On Error GoTo errProc
'
'   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
'      With txtField(Index)
'         If KeyCode = vbKeyF3 Then
'            oTrans.SearchRecord.Text
'            If .Text <> "" Then SetNextFocus
'         Else
'            If .Text <> "" Then oTrans.SearchRecord.Text
'         End If
'      End With
'      KeyCode = 0
'   End If
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " & Index _
'                       & ", " & KeyCode _
'                       & ", " & Shift & " )", True
'End Sub
'
'
'Private Sub txtField_LostFocus(Index As Integer)
'   With txtField(Index)
'      .BackColor = oApp.getColor("EB")
'   End With
'End Sub
'
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'   Select Case KeyCode
'   Case vbKeyReturn, vbKeyUp, vbKeyDown
'      Select Case KeyCode
'      Case vbKeyReturn, vbKeyDown
'         SetNextFocus
'      Case vbKeyUp
'         SetPreviousFocus
'      End Select
'   End Select
'End Sub
'
'Private Sub oTrans_Delete(Deleted As Boolean)
'   Deleted = True
'End Sub
'
'Private Sub oTrans_DeleteComplete()
'   For pnCtr = 0 To txtOthers.Count - 1
'      txtOthers(pnCtr).Text = ""
'   Next
'End Sub
'
'Private Sub txtOthers_Validate(Index As Integer, Cancel As Boolean)
'   Dim lsOldProc As String
'   Dim lnCtr As Integer
'
'   lsOldProc = "txtOthers_Validate"
'   'On Error GoTo errProc
'
''   With txtOthers(Index)
''      .Text = TitleCase(.Text)
''
''      Select Case Index
''      Case 0 To 3, 5, 6, 7
''         If Not IsNumeric(.Text) Then .Text = 0
''         .Text = Format(.Text, "#,##0.00")
''         oRS(Index) = CDbl(.Text)
''      Case 8, 9
''         If Trim(.Text) = "" Then
''            If oRS.EOF Then Exit Sub
''            InitOthers
''            txtOthers(IIf(Index = 8, 9, 8)).Text = ""
''            txtOthers(IIf(Index = 8, 9, 8)).Tag = ""
''
''            For lnCtr = 0 To txtField.Count - 1
''               txtField(lnCtr).Text = ""
''               txtField(lnCtr).Tag = txtField(lnCtr).Text
''            Next
''            Exit Sub
''         End If
''
''         If Trim(LCase(.Tag)) <> Trim(LCase(.Text)) Then SearchBarCode .Text, IIf(Index = 8, True, False)
''         .Tag = .Text
''      End Select
''   End With
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " & Cancel & " )", True
'End Sub
'Private Sub ShowError(ByVal lsProcName As String, Optional bEnd As Boolean = False)
'   With oApp
'      .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
'      If bEnd Then
'         .xShowError
'         End
'      Else
'         With Err
'            .Raise .Number, .Source, .Description
'         End With
'      End If
'   End With
'End Sub
'
'Private Sub ClearFields()
'   For pnCtr = 0 To txtField.Count - 1
'      Select Case pnCtr
'      Case 0
''        txtField(pnCtr).Text = Format(oTrans.Master(pnCtr), "@@@@-@@@@@@")
'      Case 7
'         txtField(pnCtr).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
'      Case Else
'         txtField(pnCtr).Text = Empty
'      End Select
'   Next
'End Sub
'
