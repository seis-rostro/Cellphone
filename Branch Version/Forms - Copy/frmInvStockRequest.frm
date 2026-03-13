VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmInvStockRequest 
   BorderStyle     =   0  'None
   Caption         =   "Inventory Stock Request w/ ROQ Computation"
   ClientHeight    =   9630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15255
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9630
   ScaleWidth      =   15255
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame3 
      Height          =   3105
      Left            =   6825
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   5477
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   3015
         Left            =   30
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   30
         Width           =   8205
         _ExtentX        =   14473
         _ExtentY        =   5318
         _Version        =   393216
         FocusRect       =   0
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   3105
      Left            =   1590
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   5220
      _ExtentX        =   9208
      _ExtentY        =   5477
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   4485
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Text            =   "A"
         Top             =   1972
         Width           =   600
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   9
         Left            =   2310
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   "00"
         Top             =   2362
         Width           =   600
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   8
         Left            =   1095
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Text            =   "00"
         Top             =   2362
         Width           =   600
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   2310
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   "00"
         Top             =   1972
         Width           =   600
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   1095
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "00"
         Top             =   1972
         Width           =   600
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   1095
         TabIndex        =   5
         Text            =   "United Excelsior Marketing Inc"
         Top             =   750
         Width           =   3990
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   1095
         TabIndex        =   7
         Text            =   "United Excelsior Marketing Inc"
         Top             =   1140
         Width           =   3990
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   1095
         TabIndex        =   9
         Text            =   "United Excelsior Marketing Inc"
         Top             =   1530
         Width           =   3990
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   10
         Left            =   3945
         MaxLength       =   3
         TabIndex        =   21
         Text            =   "22"
         Top             =   2362
         Width           =   1140
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   1095
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   "M001-12-000001"
         Top             =   180
         Width           =   1860
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   11
         Left            =   3660
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   "Mmm dd, yyyy"
         Top             =   180
         Width           =   1425
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Classification"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   6
         Left            =   3255
         TabIndex        =   14
         Top             =   2040
         Width           =   1125
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ROQ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   5
         Left            =   1860
         TabIndex        =   18
         Top             =   2430
         Width           =   405
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AMC"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   585
         TabIndex        =   16
         Top             =   2430
         Width           =   375
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QOH"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   1860
         TabIndex        =   12
         Top             =   2040
         Width           =   405
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "On Trnsit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   210
         TabIndex        =   10
         Top             =   2040
         Width           =   750
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Brand"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   13
         Left            =   465
         TabIndex        =   4
         Top             =   825
         Width           =   495
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   14
         Left            =   465
         TabIndex        =   6
         Top             =   1215
         Width           =   495
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Color"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   15
         Left            =   510
         TabIndex        =   8
         Top             =   1605
         Width           =   450
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Order"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   16
         Left            =   3045
         TabIndex        =   20
         Top             =   2377
         Width           =   810
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   3210
         TabIndex        =   2
         Top             =   255
         Width           =   390
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   360
         Left            =   1185
         Tag             =   "et0;ht2"
         Top             =   270
         Width           =   1860
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trans. No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   105
         TabIndex        =   0
         Top             =   255
         Width           =   855
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5775
      Left            =   1575
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   3675
      Width           =   13560
      _ExtentX        =   23918
      _ExtentY        =   10186
      _Version        =   393216
      FillStyle       =   1
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   90
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
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
      Picture         =   "frmInvStockRequest.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   90
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   1170
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&History"
      AccessKey       =   "H"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmInvStockRequest.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   90
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   540
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&OK"
      AccessKey       =   "O"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmInvStockRequest.frx":180C
   End
End
Attribute VB_Name = "frmInvStockRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
'Private Const pxeMODULENAME = "frmInvStockRequest"
'Private Const pxeVisibleRow = 20
'
'Private oSkin As clsFormSkin
'Private WithEvents oTrans As clsMCStockOrder
'Private poRSRecOrder As Recordset
'Private poRS As Recordset
'
'Private pnActiveRow As Integer
'Private pbControl As Boolean
'Private pnIndex As Integer
'Private pbByModel As Boolean
'Private pbLoaded As Boolean
'
'Private Sub cmdButton_Click(Index As Integer)
'   Dim lsOldProc As String
'
'   lsOldProc = "cmdButton_Click"
'   On Error GoTo errProc
'
'   Select Case Index
'      Case 0 'OK/save
'         If SaveTransaction Then
'            MsgBox "Transaction Saved Successfully.", vbInformation, "Notice"
'            InitTransaction
'         Else
'            MsgBox "Unable to Save Transaction. Please verify your entry.", vbInformation, "Notice"
'         End If
'      Case 1 'history
'         With frmMCOrderHistory
'            .Brand = poRSRecOrder.Fields("sBrandNme")
'            .Model = poRSRecOrder.Fields("sModelNme")
'
'            If pbByModel Then
'               .History = oTrans.GetHistoryModel(poRSRecOrder.Fields("sModelIDx"))
'            Else
'               .History = oTrans.getHistory(poRSRecOrder.Fields("sMCInvIDx"))
'            End If
'
'            .Show vbModal
'         End With
'      Case 2 'cancel
'      Case 3 'post
'      Case 4 'close
'         Unload Me
'   End Select
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " & " )", True
'End Sub
'Private Sub Form_Activate()
'   Dim lsOldProc As String
'
'   lsOldProc = "Form_Activate"
'   On Error GoTo errProc
'
'   oApp.MenuName = Me.Tag
'   Me.ZOrder 0
'
'   If Not pbLoaded Then pbLoaded = True
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " & " )", True
'End Sub
'
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'   With MSFlexGrid1
'      Select Case KeyCode
'         Case vbKeyReturn, vbKeyDown
'            If GetFocus = MSFlexGrid2.hwnd Then
'
'            Else
'               If pnIndex = 10 And pnActiveRow < .Rows - 1 Then
'                  ' this does not trigger lost focus or validate
'                  If pnIndex = 10 Then
'                     Call txtField_Validate(pnIndex, False)
'                  End If
'                  .Row = pnActiveRow + 1
'
'                  Call HiglightRow(Me.MSFlexGrid1, .Row, 1)
'                  pnActiveRow = .Row
'                  If .Row > 19 And Not .RowIsVisible(.Row) Then _
'                     .TopRow = .Row - (pxeVisibleRow - 3)
'
'                  Call setFieldInfo
'                  With txtField(10)
'                     .SelStart = 0
'                     .SelLength = Len(.Text)
'                     .SetFocus
'                  End With
'                  Exit Sub
'               Else
'                  SetNextFocus
'               End If
'            End If
'         Case vbKeyUp
'            If pbControl Then
'               If .Row > 2 Then
'                  ' this does not trigger lost focus or validate
'                  If pnIndex = 10 Then
'                     Call txtField_Validate(pnIndex, False)
'                  End If
'
'                  If .Row = .TopRow Then .TopRow = .TopRow - 1
'
'                  .Row = pnActiveRow - 1
'                  Call HiglightRow(Me.MSFlexGrid1, .Row, 1)
'                  pnActiveRow = .Row
'
'                  Call setFieldInfo
'                  With txtField(10)
'                     .SelStart = 0
'                     .SelLength = Len(.Text)
'                     .SetFocus
'                  End With
'               End If
'            Else
'               SetPreviousFocus
'            End If
'         Case vbKeyControl
'            pbControl = True
'            KeyCode = 0
'      End Select
'   End With
'End Sub
'
'Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'   If pbControl Then
'      If KeyCode = pbControl Then pbControl = False
'   End If
'End Sub
'
'Private Sub Form_Load()
'   Dim lsOldProc As String
'   lsOldProc = "Form_Load"
'
'   On Error GoTo errProc
'
'   Set oSkin = New clsFormSkin
'   Set oSkin.AppDriver = oApp
'   Set oSkin.Form = Me
'   oSkin.ApplySkin xeFormTransEqualLeft
'   CenterChildForm mdiMain, Me
'
'   Set oTrans = New clsMCStockOrder
'   Set oTrans.AppDriver = oApp
'   oTrans.Branch = oApp.BranchCode
'
'   InitTransaction
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " & " )", True
'End Sub
'
'Private Sub AddDetail()
'   Dim lsOldProc As String
'
'   lsOldProc = "AddDetail"
'   On Error GoTo errProc
'
'   With poRSRecOrder
'      poRS.AddNew
'      poRS("sBrandIDx") = .Fields("sBrandIDx")
'      poRS("sModelIDx") = .Fields("sModelIDx")
'      poRS("sColorIDx") = .Fields("sColorIDx")
'      poRS("nQuantity") = .Fields("nQuantity")
'      poRS("sMCInvIDx") = .Fields("sMCInvIDx")
'      poRS("sBrandNme") = .Fields("sBrandNme")
'      poRS("sModelNme") = .Fields("sModelNme")
'      poRS("sColorNme") = .Fields("sColorNme")
'      poRS("cCategory") = IIf(pbByModel, "Mod", "Inv")
'
'      If pbByModel = False Then
'         poRS("cClassify") = .Fields("cClassify")
'         poRS("nAveMonsl") = .Fields("nAveMonsl")
'         poRS("nRecOrder") = .Fields("nRecOrder")
'         poRS("nQtyOnHnd") = .Fields("nQtyOnHnd")
'         poRS("nOnTranst") = .Fields("nOnTranst")
'      Else
'         poRS("cClassify") = .Fields("cClassMdl")
'         poRS("nAveMonsl") = .Fields("nAveMonMd")
'         poRS("nRecOrder") = .Fields("nRecOrdMd")
'         poRS("nQtyOnHnd") = .Fields("nQtyOnHMd")
'         poRS("nOnTranst") = .Fields("nOnTrnsMd")
'      End If
'   End With
'
'   Call LoadDetail
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " & " )", True
'End Sub
'
'Private Sub LoadDetail()
'   Dim lsOldProc As String
'   Dim lnCtr As Integer
'   Dim lnRow As Integer
'
'   lsOldProc = "LoadDetail"
'   On Error GoTo errProc
'
'   With MSFlexGrid2
'      InitGrid2
'      lnRow = poRS.RecordCount
'
'      .Rows = lnRow + 2
'
'      If .Rows > 11 Then
'         .ColWidth(1) = 2490
'      Else
'         .ColWidth(1) = 2750
'      End If
'
'      lnCtr = 0
'      poRS.Sort = "sModelNme ASC"
'      If poRS.RecordCount <> 0 Then poRS.MoveFirst
'      Do Until poRS.EOF
'         DoEvents
'         .TextMatrix(lnCtr + 1, 0) = lnCtr + 1
'         .TextMatrix(lnCtr + 1, 1) = poRS("sModelNme")
'         .TextMatrix(lnCtr + 1, 2) = poRS("sColorNme")
'         .TextMatrix(lnCtr + 1, 3) = poRS("nQtyOnHnd")
'         .TextMatrix(lnCtr + 1, 4) = poRS("nAveMonsl")
'         .TextMatrix(lnCtr + 1, 5) = poRS("nRecOrder")
'         .TextMatrix(lnCtr + 1, 6) = poRS("nQuantity")
'
'         lnCtr = lnCtr + 1
'         poRS.MoveNext
'      Loop
'      .TextMatrix(lnCtr + 1, 0) = lnCtr + 1
'   End With
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " & " )", True
'End Sub
'
'Private Sub UpdateDetail()
'   Dim lsOldProc As String
'
'   lsOldProc = "UpdateDetail"
'   On Error GoTo errProc
'
'   With MSFlexGrid2
'      poRS.Move .Row - 1, adBookmarkFirst
'      poRS("nQuantity") = poRSRecOrder("nQuantity")
'      .TextMatrix(.Row, 6) = poRS("nQuantity")
'   End With
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " & " )", True
'End Sub
'
'Private Sub DeleteDetail()
'   Dim lsOldProc As String
'
'   lsOldProc = "DeleteDetail"
'   On Error GoTo errProc
'
'   poRS.Move MSFlexGrid2.Row - 1, adBookmarkFirst
'   poRS.Delete
'
'   Call LoadDetail
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " & " )", True
'End Sub
'
'Private Sub InitGrid()
'   Dim lnCtr As Integer
'
'   With MSFlexGrid1
'      .Cols = 10
'      .Rows = 3
'      .FixedRows = 2
'      .Clear
'
'      pnActiveRow = 0
'      .Row = 0
'      .TextMatrix(0, 0) = " "
'      .TextMatrix(0, 1) = " "
'      .TextMatrix(0, 2) = " "
'      .TextMatrix(0, 3) = " "
'      .TextMatrix(0, 4) = "AMC"
'      .TextMatrix(0, 5) = "AMC"
'      .TextMatrix(0, 6) = "ROQ"
'      .TextMatrix(0, 7) = "ROQ"
'      .TextMatrix(0, 8) = "QOH"
'      .TextMatrix(0, 9) = "QOH"
'      'Row 1
'      .TextMatrix(1, 0) = "No"
'      .TextMatrix(1, 1) = "Brand"
'      .TextMatrix(1, 2) = "Model"
'      .TextMatrix(1, 3) = "Color"
'      .TextMatrix(1, 4) = "Model"
'      .TextMatrix(1, 5) = "Inv."
'      .TextMatrix(1, 6) = "Model"
'      .TextMatrix(1, 7) = "Inv."
'      .TextMatrix(1, 8) = "Model"
'      .TextMatrix(1, 9) = "Inv."
'
'      .MergeCells = flexMergeFree 'disables colsel procedure
'      .MergeRow(0) = True
'
'      .Row = 0
'      'Column Width
'      .ColWidth(0) = 530
'      .ColWidth(1) = 2975 '2725
'      .ColWidth(2) = 3380
'      .ColWidth(3) = 1800
'      .ColWidth(4) = 800
'      .ColWidth(5) = 800
'      .ColWidth(6) = 800
'      .ColWidth(7) = 800
'      .ColWidth(8) = 800
'      .ColWidth(9) = 800
'
'      'Column Alignment
'      For lnCtr = 0 To .Cols - 1
'         .Col = lnCtr
'         .CellFontBold = True
'         .CellAlignment = 3
'      Next
'      .Row = 1
'      'Column Alignment
'      For lnCtr = 0 To .Cols - 1
'         .Col = lnCtr
'         .CellFontBold = True
'         .CellAlignment = 3
'      Next
'
'      .Row = 2
'      .ColAlignment(0) = flexAlignLeftCenter
'      .ColAlignment(1) = flexAlignLeftCenter
'      .ColAlignment(2) = flexAlignLeftCenter
'      .ColAlignment(3) = flexAlignLeftCenter
'      .ColAlignment(4) = flexAlignRightCenter
'      .ColAlignment(5) = flexAlignRightCenter
'      .ColAlignment(6) = flexAlignRightCenter
'      .ColAlignment(7) = flexAlignRightCenter
'      .ColAlignment(8) = flexAlignRightCenter
'      .ColAlignment(9) = flexAlignRightCenter
'
'      Call HiglightRow(Me.MSFlexGrid1, .Row, 1)
'      pnActiveRow = .Row
'   End With
'End Sub
'
'Private Sub InitGrid2()
'   Dim lnCtr As Integer
'
'   With MSFlexGrid2
'      .Cols = 7
'      .Rows = 2
'
'      .Clear
'
'      .TextMatrix(0, 0) = "No."
'      .TextMatrix(0, 1) = "Model"
'      .TextMatrix(0, 2) = "Color"
'      .TextMatrix(0, 3) = "QOH"
'      .TextMatrix(0, 4) = "AMC"
'      .TextMatrix(0, 5) = "ROQ"
'      .TextMatrix(0, 6) = "QTY"
'      .TextMatrix(1, 0) = "1"
'
'      .Row = 0
'      'Column Width
'      .ColWidth(0) = 442
'      .ColWidth(1) = 2750 '2490
'      .ColWidth(2) = 1800
'      .ColWidth(3) = 800
'      .ColWidth(4) = 800
'      .ColWidth(5) = 800
'      .ColWidth(6) = 800
'
'      'Column Alignment
'      For lnCtr = 0 To .Cols - 1
'         .Col = lnCtr
'         .CellFontBold = True
'         .CellAlignment = 3
'      Next
'
'      .Row = 1
'      .ColAlignment(0) = flexAlignLeftCenter
'      .ColAlignment(1) = flexAlignLeftCenter
'      .ColAlignment(2) = flexAlignLeftCenter
'      .ColAlignment(3) = flexAlignRightCenter
'      .ColAlignment(4) = flexAlignRightCenter
'      .ColAlignment(5) = flexAlignRightCenter
'      .ColAlignment(6) = flexAlignRightCenter
'
'      If Not pbLoaded Then
'         .Col = 1
'         .ColSel = .Cols - 1
'      End If
'   End With
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'   Set oSkin = Nothing
'   pnActiveRow = 0
'End Sub
'
'Private Sub MSFlexGrid1_Click()
'   Call setFieldInfo(True)
'   With txtField(10)
'      .SelStart = 0
'      .SelLength = Len(.Text)
'      .SetFocus
'   End With
'End Sub
'
'Private Sub MSFlexGrid1_DblClick()
'   Call cmdButton_Click(1)
'End Sub
'
'Private Sub MSFlexGrid1_SelChange()
'   With MSFlexGrid1
'      Call HiglightRow(Me.MSFlexGrid1, .Row, 1)
'      pnActiveRow = .Row
'
'      Call setFieldInfo
'   End With
'End Sub
'
'Private Sub MSFlexGrid2_Click()
'
'   With MSFlexGrid2
'      .Col = 1
'      .ColSel = .Cols - 1
'      If .MouseRow = .Rows - 1 Then GoTo endWithFocus
'   End With
'
'   Call setDetailInfo
'
'endWithFocus:
'   With txtField(10)
'      .SelStart = 0
'      .SelLength = Len(.Text)
'      .SetFocus
'   End With
'End Sub
'
'Private Sub HiglightRow(ByVal grid As MSFlexGrid, _
'                           ByVal Row As Integer, _
'                           ByVal colstart As Integer)
'   Dim lnCtr As Integer
'
'   With grid
'      If Row < 2 Then Exit Sub
'
'      If Row <> pnActiveRow Then
'         .Row = pnActiveRow
'         For lnCtr = colstart To .Cols - 1
'            .Col = lnCtr
'            .CellBackColor = &HFFFFFF
'         Next
'
'         .Row = Row
'         For lnCtr = colstart To .Cols - 1
'            .Col = lnCtr
'            .CellBackColor = &H8000000D
'         Next
'      End If
'   End With
'End Sub
'
'Private Sub MSFlexGrid2_DblClick()
'
'   With MSFlexGrid2
'      If .MouseRow = .Rows - 1 Then Exit Sub
'   End With
'
'   With frmMCOrderHistory
'      .Brand = poRS.Fields("sBrandNme")
'      .Model = poRS.Fields("sModelNme")
'
'      If poRS.Fields("cCategory") = "Mod" Then
'         .History = oTrans.GetHistoryModel(poRSRecOrder.Fields("sModelIDx"))
'      Else
'         .History = oTrans.getHistory(poRSRecOrder.Fields("sMCInvIDx"))
'      End If
'
'      .Show vbModal
'   End With
'End Sub
'
'Private Sub oTrans_OthersRetrieved(ByVal Index As Integer, ByVal Value As Variant)
'   Dim lsCondition As String
'
'   txtField(Index) = Value
'
'   Select Case Index
'      Case 1
'         txtField(2) = ""
'         txtField(3) = ""
'      Case 2
'         txtField(3) = ""
'   End Select
'
'   Call findMotorcycle("sBrandNme = " & strParm(txtField(1)), _
'                        IIf(txtField(2) <> "", "sModelNme = " & strParm(txtField(2)), ""), _
'                        IIf(txtField(3) <> "", "sColorNme = " & strParm(txtField(3)), ""))
'End Sub
'
'Private Sub txtField_GotFocus(Index As Integer)
'   With txtField(Index)
'      Select Case Index
'         Case 1, 2, 3, 10
'            .SelStart = 0
'            .SelLength = Len(.Text)
'            .BackColor = oApp.getColor("HT1")
'      End Select
'   End With
'
'   pnIndex = Index
'End Sub
'
'Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'   If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
'      If pbControl Then KeyCode = 0
'   End If
'
'   With oTrans
'      Select Case KeyCode
'         Case vbKeyF3
'            Call .SearchOthers(Index, txtField(Index))
'      End Select
'   End With
'End Sub
'
'Private Sub txtField_KeyPress(Index As Integer, KeyAscii As Integer)
'   If Index <> 10 Then Exit Sub
'   Select Case KeyAscii
'      Case vbKey0 To vbKey9
'      Case vbKeyBack, vbKeyClear, vbKeyDelete
'      Case vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown, vbKeyTab, vbKeyReturn
'      Case Else
'         KeyAscii = 0
'         Beep
'   End Select
'End Sub
'
'Private Sub txtField_LostFocus(Index As Integer)
'   With txtField(Index)
'      Select Case Index
'         Case 1, 2, 3, 10
'            .BackColor = oApp.getColor("EB")
'
'            If Index = 10 Then Exit Sub
'      End Select
'   End With
'   pnIndex = 0
'End Sub
'
'Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
'   Select Case Index
'      Case 1, 2, 3
'         oTrans.Others(Index) = txtField(Index)
'      Case 10
'         With poRSRecOrder
'            .Fields("nQuantity") = IIf(IsNumeric(txtField(10)), txtField(10), 0)
'            If .Fields("nQuantity") <> 0 And _
'                  findMCOnOrder(.Fields("sModelNme"), .Fields("sColorNme")) = False Then
'               Call AddDetail
'            ElseIf findMCOnOrder(.Fields("sModelNme"), .Fields("sColorNme")) = True Then
'               If .Fields("nQuantity") = 0 Then
'                  Call DeleteDetail
'               Else
'                  Call UpdateDetail
'               End If
'            End If
'         End With
'   End Select
'End Sub
'
'Private Sub setFieldInfo(Optional ByVal lbClick As Boolean = False)
'   Dim lsOldProc As String
'
'   lsOldProc = "setFieldInfo"
'   On Error GoTo errProc
'
'   If Not lbClick Then
'      If txtField(1) <> "" And txtField(2) <> "" And txtField(3) <> "" Then
'         pbByModel = False
'      ElseIf (txtField(1) <> "" And txtField(2) <> "" And txtField(3) = "") Or _
'         (txtField(1) <> "" And txtField(2) = "" And txtField(3) = "") Then
'         pbByModel = True
'      End If
'   Else
'      pbByModel = False
'   End If
'
'   With poRSRecOrder
'      .Move MSFlexGrid1.Row - 2, adAffectCurrent
'
'      txtField(10) = .Fields("nQuantity")
'
'      If pbByModel = False Then
'         txtField(1) = .Fields("sBrandNme")
'         txtField(2) = .Fields("sModelNme")
'         txtField(3) = .Fields("sColorNme")
'         txtField(5) = .Fields("cClassify")
'         txtField(6) = .Fields("nOnTranst")
'         txtField(7) = .Fields("nQtyOnHnd")
'         txtField(8) = .Fields("nAveMonSl")
'         txtField(9) = .Fields("nRecOrder")
'      Else
'         txtField(5) = .Fields("cClassMdl")
'         txtField(6) = .Fields("nOnTrnsMd")
'         txtField(7) = .Fields("nQtyOnHMd")
'         txtField(8) = .Fields("nAveMonMd")
'         txtField(9) = .Fields("nRecOrdMd")
'      End If
'
'      pnActiveRow = MSFlexGrid1.Row
'
'      Call findMCOnOrder(.Fields("sModelNme"), .Fields("sColorNme"))
'   End With
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " & " )", True
'End Sub
'
'Private Sub setDetailInfo()
'   Dim lsOldProc As String
'
'   lsOldProc = "setDetailInfo"
'   On Error GoTo errProc
'
'   With poRS
'      .Move MSFlexGrid2.Row - 1, adAffectCurrent
'
'      txtField(1) = .Fields("sBrandNme")
'      txtField(2) = .Fields("sModelNme")
'      txtField(3) = .Fields("sColorNme")
'      txtField(10) = .Fields("nQuantity")
'      txtField(5) = .Fields("cClassify")
'      txtField(6) = .Fields("nOnTranst")
'      txtField(7) = .Fields("nQtyOnHnd")
'      txtField(8) = .Fields("nAveMonSl")
'      txtField(9) = .Fields("nRecOrder")
'
'      Call findMotorcycle("sBrandNme = " & strParm(poRS("sBrandNme")), _
'                     "sModelNme = " & strParm(poRS("sModelNme")), _
'                     "sColorNme = " & strParm(poRS("sColorNme")), False)
'      pnActiveRow = MSFlexGrid1.Row
'   End With
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " & " )", True
'End Sub
'
'Private Sub ClearFields()
'   Dim lotxt As TextBox
'
'   For Each lotxt In txtField
'      lotxt = ""
'   Next
'End Sub
'
'Private Sub LoadMaster()
'   With oTrans
'      txtField(0) = Format(.Master("sTransNox"), "@@@@-@@-@@@@@@")
'      txtField(11) = strLongDate(.Master("dTransact"))
'   End With
'End Sub
'Private Sub LoadRecOrder()
'   Dim lnCtr As Integer
'   Dim lsOldProc As String
'
'   lsOldProc = "LoadRecOrder"
'   On Error GoTo errProc
'
'   With MSFlexGrid1
'      .Rows = poRSRecOrder.RecordCount + 3
'
'      If .Rows > 20 Then
'         .ColWidth(1) = 2725
'      Else
'         .ColWidth(1) = 2975
'      End If
'   End With
'
'   With poRSRecOrder
'      .MoveFirst
'      lnCtr = 0
'      Do Until .EOF
'         DoEvents
'         MSFlexGrid1.TextMatrix(lnCtr + 2, 0) = lnCtr + 1
'         MSFlexGrid1.TextMatrix(lnCtr + 2, 1) = .Fields("sBrandNme")
'         MSFlexGrid1.TextMatrix(lnCtr + 2, 2) = .Fields("sModelNme")
'         MSFlexGrid1.TextMatrix(lnCtr + 2, 3) = .Fields("sColorNme")
'         MSFlexGrid1.TextMatrix(lnCtr + 2, 4) = .Fields("nAveMonMd")
'         MSFlexGrid1.TextMatrix(lnCtr + 2, 5) = .Fields("nAveMonsl")
'         MSFlexGrid1.TextMatrix(lnCtr + 2, 6) = .Fields("nRecOrdMd")
'         MSFlexGrid1.TextMatrix(lnCtr + 2, 7) = .Fields("nRecOrder")
'         MSFlexGrid1.TextMatrix(lnCtr + 2, 8) = .Fields("nQtyOnHMd")
'         MSFlexGrid1.TextMatrix(lnCtr + 2, 9) = .Fields("nQtyOnHnd")
'         lnCtr = lnCtr + 1
'         .MoveNext
'      Loop
'   End With
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " & " )", True
'End Sub
'Private Sub findMotorcycle(ByVal sBrandNme As String, _
'                              Optional ByVal sModelNme As String = "", _
'                              Optional ByVal sColorNme As String = "", _
'                              Optional ByVal bDisplayV As Boolean = True)
'   Dim lnPos As Integer
'
'   With poRSRecOrder
'      Call .Find(sBrandNme, 0, adSearchForward, 1)
'      If Not .EOF Then
'         lnPos = .AbsolutePosition
'      Else
'         lnPos = 1
'      End If
'
'      If sModelNme <> "" Then
'         Call .Find(sModelNme, lnPos - 1, adSearchForward, 1)
'         If Not .EOF Then
'            lnPos = .AbsolutePosition
'         Else
'            lnPos = lnPos
'         End If
'
'
'         If sColorNme <> "" Then
'            Call .Find(sColorNme, lnPos - 1, adSearchForward, 1)
'            If Not .EOF Then
'               lnPos = .AbsolutePosition
'            Else
'               lnPos = lnPos
'            End If
'         End If
'      End If
'
'      .Cancel
'   End With
'
'   With MSFlexGrid1
'      .Row = lnPos + 1
'      .TopRow = .Row
'   End With
'
'   Call HiglightRow(Me.MSFlexGrid1, MSFlexGrid1.Row, 1)
'   If bDisplayV Then Call setFieldInfo
'End Sub
'
'Private Function findMCOnOrder(ByVal sModelNme As String, _
'                              ByVal sColorNme As String) As Boolean
'
'   Dim lnPos As Integer
'
'   With poRS
'      Call .Find("sModelNme = " & strParm(sModelNme), 0, adSearchForward, adBookmarkFirst)
'      If Not .EOF Then
'         lnPos = .AbsolutePosition
'
'         Call .Find("sColorNme = " & strParm(sColorNme), lnPos - 1, adSearchForward, adBookmarkFirst)
'         If Not .EOF Then
'            If .Fields("sModelNme") = sModelNme Then
'               lnPos = .AbsolutePosition
'               findMCOnOrder = True
'            Else
'               lnPos = 0
'            End If
'         Else
'            lnPos = 0
'         End If
'      Else
'         lnPos = 0
'      End If
'      .Cancel
'   End With
'
'   With MSFlexGrid2
'      .Row = IIf(lnPos = 0, .Rows - 1, lnPos)
'      If .Row > 10 Then .TopRow = .Row - 9
'      .Col = 1
'      .ColSel = .Cols - 1
'   End With
'End Function
'
'Private Function SaveTransaction() As Boolean
'   Dim lnCtr As Integer
'   Dim lsOldProc As String
'
'   lsOldProc = "SaveTransaction"
'   On Error GoTo errProc
'
'   If poRS.RecordCount = 0 Then GoTo endProc
'
'   'pass detail to class
'   With oTrans
'      poRS.MoveFirst
'      For lnCtr = 0 To poRS.RecordCount - 1
'         If poRS.Fields("sBrandIDx") = "" Then GoTo move2Next
'
'         .Detail(lnCtr, "sBrandIDx") = poRS("sBrandIDx")
'         .Detail(lnCtr, "sModelIDx") = poRS("sModelIDx")
'         .Detail(lnCtr, "sColorIDx") = poRS("sColorIDx")
'         .Detail(lnCtr, "nQuantity") = poRS("nQuantity")
'         .Detail(lnCtr, "nRecOrder") = poRS("nRecOrder")
'         .Detail(lnCtr, "nQtyOnHnd") = poRS("nQtyOnHnd")
'         .Detail(lnCtr, "cClassify") = poRS("cClassify")
'         .Detail(lnCtr, "sMCInvIDx") = poRS("sMCInvIDx")
'         .AddDetail
'move2Next:
'         poRS.MoveNext
'      Next
'   'save
'      SaveTransaction = .SaveTransaction
'      Set poRS = Nothing
'      Set poRSRecOrder = Nothing
'   End With
'
'endProc:
'   Exit Function
'errProc:
'   ShowError lsOldProc & "( " & " )", True
'End Function
'
'Private Sub InitTransaction()
'   pbLoaded = False
'   ClearFields
'   InitGrid
'   InitGrid2
'
'   oTrans.InitTransaction
'   oTrans.NewTransaction
'   Set poRSRecOrder = oTrans.LoadRecOrder
'
'   LoadMaster
'   LoadRecOrder
'
'   Set poRS = New Recordset
'   With poRS.Fields
'      .Append "sBrandIDx", adVarChar, 7
'      .Append "sModelIDx", adVarChar, 9
'      .Append "sColorIDx", adVarChar, 7
'      .Append "nQuantity", adInteger
'      .Append "nRecOrder", adInteger
'      .Append "nQtyOnHnd", adInteger
'      .Append "cClassify", adChar, 1
'      .Append "sMCInvIDx", adVarChar, 9
'      .Append "nAveMonsl", adInteger
'      .Append "nOnTranst", adInteger
'      .Append "sBrandNme", adVarChar, 30
'      .Append "sModelNme", adVarChar, 30
'      .Append "sColorNme", adVarChar, 30
'      .Append "cCategory", adChar, 3
'      poRS.Open
'   End With
'
'   pbLoaded = True
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
