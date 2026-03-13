VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0#1.0#0"; "xrGridControl.ocx"
Begin VB.Form frmCP_CustomerOrder_Posting 
   BorderStyle     =   0  'None
   Caption         =   "Spareparts Customer Order Posting"
   ClientHeight    =   7800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11940
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7800
   ScaleWidth      =   11940
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   10590
      TabIndex        =   20
      Top             =   4335
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
      Picture         =   "frmCP_CustomerOrder_Posting.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   10590
      TabIndex        =   17
      Top             =   2445
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
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
      Picture         =   "frmCP_CustomerOrder_Posting.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   10590
      TabIndex        =   18
      Top             =   3075
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "Con&firm"
      AccessKey       =   "f"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCP_CustomerOrder_Posting.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   10590
      TabIndex        =   19
      Top             =   3705
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "Cance&l"
      AccessKey       =   "l"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCP_CustomerOrder_Posting.frx":166E
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   525
      Index           =   1
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   10245
      _ExtentX        =   18071
      _ExtentY        =   926
      Begin VB.TextBox txtField 
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
         Height          =   315
         Index           =   6
         Left            =   4725
         MaxLength       =   50
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   90
         Width           =   5415
      End
      Begin VB.TextBox txtField 
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
         Height          =   315
         Index           =   7
         Left            =   915
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   90
         Width           =   2250
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Company Name"
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
         Index           =   8
         Left            =   3300
         TabIndex        =   2
         Top             =   135
         Width           =   1410
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Trans.No."
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
         Left            =   45
         TabIndex        =   0
         Top             =   135
         Width           =   825
      End
   End
   Begin xrGridEditor.GridEditor GridEditor1 
      Height          =   4800
      Left            =   90
      TabIndex        =   16
      Tag             =   "et0;eb0;et0;bc2"
      Top             =   2880
      Width           =   10260
      _ExtentX        =   18098
      _ExtentY        =   8467
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
      Object.HEIGHT          =   4800
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
      MOUSEICON       =   "frmCP_CustomerOrder_Posting.frx":1DE8
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
      Height          =   1755
      Index           =   0
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   1080
      Width           =   10260
      _ExtentX        =   18098
      _ExtentY        =   3096
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   915
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   150
         Width           =   2265
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   7305
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   675
         Width           =   1905
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   915
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   675
         Width           =   5235
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   3
         Left            =   915
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   1005
         Width           =   5235
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   4
         Left            =   915
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   1335
         Width           =   5235
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   645
         Index           =   5
         Left            =   7305
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   1005
         Width           =   2835
      End
      Begin VB.Shape Shape4 
         Height          =   330
         Index           =   0
         Left            =   7335
         Top             =   180
         Width           =   2775
      End
      Begin VB.Shape Shape3 
         Height          =   390
         Index           =   0
         Left            =   7305
         Top             =   150
         Width           =   2835
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "unknown"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   7365
         TabIndex        =   21
         Tag             =   "eb0;et0"
         Top             =   225
         Width           =   2715
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   315
         Left            =   1005
         Tag             =   "et0;ht2"
         Top             =   255
         Width           =   2265
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trans. No."
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   4
         Top             =   210
         Width           =   750
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transact. Date"
         Height          =   195
         Index           =   1
         Left            =   6195
         TabIndex        =   12
         Top             =   735
         Width           =   1065
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   6
         Top             =   735
         Width           =   660
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   10
         Top             =   1380
         Width           =   570
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   195
         Index           =   5
         Left            =   6195
         TabIndex        =   14
         Top             =   1050
         Width           =   630
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Full Name"
         Height          =   195
         Index           =   6
         Left            =   90
         TabIndex        =   8
         Top             =   1065
         Width           =   705
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H8000000F&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   285
         Index           =   0
         Left            =   7365
         Tag             =   "et0;et0"
         Top             =   210
         Width           =   2730
      End
   End
End
Attribute VB_Name = "frmCP_CustomerOrder_Posting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Private Const pxeMODULENAME = "frmSP_CustomerOrder_Posting"
'
'Private WithEvents oTrans As clsCustomerOrder
'Private oSkin As clsFormSkin
'
'Dim pnCtr As Integer
'Dim pbLoaded As Boolean
'
'Private Sub cmdButton_Click(Index As Integer)
'   Dim lsOldProc As String
'   Dim lsRep As String
'
'   lsOldProc = "cmdButton_Click"
'   On Error GoTo errProc
'
'   With GridEditor1
'      Select Case Index
'      Case 0 'Browse
'         If oTrans.SearchTransaction() Then
'            LoadMaster
'            LoadDetail
'         End If
'         txtField(7).SetFocus
'      Case 1 'Post
'         If pbLoaded Then
'            lsRep = MsgBox("Confirm Transaction!!!", vbYesNo + vbQuestion, "Confrim")
'
'            If lsRep = vbYes Then
'               If Not oTrans.CloseTransaction(oTrans.Master("sTransNox")) Then
'                  MsgBox "Unable to Confirm Transaction!!!", vbCritical, "Warning"
'               Else
'                  MsgBox "Transaction Save Successfully!!!", vbInformation, "Confirm"
'                  Label2.Caption = Format(TransStat(oTrans.Master("cTranStat")), ">")
'               End If
'            End If
'         Else
'            MsgBox "Unable to Confirm Transaction!!!" & vbCrLf & _
'                   "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
'         End If
'      Case 2 'Cancel
'         If pbLoaded Then
'            If oTrans.CancelTransaction Then
'               MsgBox "Transaction Cancelled Successfully!!!", vbInformation, "Notice"
'               Label2.Caption = Format(TransStat(oTrans.Master("cTranStat")), ">")
'            Else
'               MsgBox "Unable to Cancel Transaction!!!", vbCritical, "Warning"
'            End If
'         Else
'            MsgBox "Unable to Cancel Transaction!!!" & vbCrLf & _
'                   "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
'         End If
'      Case 3 'Close
'         Unload Me
'      End Select
'   End With
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " & Index & " )", True
'End Sub
'
'Private Sub Form_Activate()
'   oApp.MenuName = Me.Tag
'   Me.ZOrder 0
'
'   GridEditor1.Refresh
'End Sub
'
'Private Sub Form_Load()
'   Dim lsOldProc As String
'
'   lsOldProc = "Form_Load"
'   On Error GoTo errProc
'
'   CenterChildForm mdiMain, Me
'
'   Set oTrans = New clsCustomerOrder
'   Set oTrans.AppDriver = oApp
'
'   oTrans.TransStatus = 10
'   oTrans.InitTransaction
'
'   InitGrid
'   InitEntry
'
'   Set oSkin = New clsFormSkin
'   Set oSkin.AppDriver = oApp
'   Set oSkin.Form = Me
'   oSkin.ApplySkin xeFormTransMaintenance
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
'End Sub
'
'Private Sub oTrans_DetailRetrieved(ByVal Index As Integer)
'   With GridEditor1
'      .TextMatrix(.Row, Index) = oTrans.Detail(.Row - 1, IIf(Index = 4, Index + 1, Index))
'   End With
'End Sub
'
'Private Sub oTrans_MasterRetrieved(ByVal Index As Integer)
'   txtField(Index).Text = oTrans.Master(Index)
'End Sub
'
'Private Sub InitEntry()
'   For pnCtr = 0 To txtField.Count - 1
'     txtField(pnCtr).Text = ""
'   Next
'
'   txtField(6).Tag = ""
'   txtField(7).Tag = ""
'   Label2.Caption = "UNKNOWN"
'
'   With GridEditor1
'      .Rows = 2
'      .Row = 1
'      .Col = 1
'
'      .ColWidth(2) = 4500
'      .ColWidth(3) = 2880
'
'      'empty row
'      .TextMatrix(1, 1) = ""
'      .TextMatrix(1, 2) = ""
'      .TextMatrix(1, 3) = ""
'      .TextMatrix(1, 4) = 0
'   End With
'
'   pbLoaded = False
' End Sub
'
'Private Sub InitGrid()
'   Dim lnCtr As Integer
'
'   With GridEditor1
'      .Cols = 5
'      .Rows = 2
'      .Font = "MS Sans Serif"
'
'      'Column Title
'      .TextMatrix(0, 1) = "Barcode"
'      .TextMatrix(0, 2) = "Description"
'      .TextMatrix(0, 3) = "Model"
'      .TextMatrix(0, 4) = "Qty."
'
'      .Row = 0
'
'      'Column Alignment
'      For lnCtr = 0 To .Cols - 1
'         .Col = lnCtr
'         .CellFontBold = True
'         .CellAlignment = 3
'      Next
'
'      'Column Width
'      .ColWidth(0) = 300
'      .ColWidth(1) = 1900
'      .ColWidth(4) = 600
'
'      .ColAlignment(1) = 1
'      .ColAlignment(2) = 1
'      .ColAlignment(3) = 1
'
'      .ColEnabled(3) = False
'
'      .ColNumberOnly(4) = True
'      .ColFormat(4) = "#,##0"
'      .ColDefault(4) = 0
'   End With
'End Sub
'
'Private Sub txtField_GotFocus(Index As Integer)
'   With txtField(Index)
'      .SelStart = 0
'      .SelLength = Len(txtField(Index).Text)
'      .BackColor = oApp.getColor("HT1")
'   End With
'End Sub
'
'Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'   Dim lsOldProc As String
'
'   lsOldProc = "txtField_KeyDown"
'   On Error GoTo errProc
'
'   Select Case Index
'   Case 6, 7
'      If KeyCode = vbKeyF3 Then
'         If oTrans.SearchTransaction() = True Then
'            LoadMaster
'            LoadDetail
'         Else
'            If txtField(0).Text <> "" Then InitEntry
'         End If
'         txtField(7).SetFocus
'         KeyCode = 0
'      End If
'   End Select
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " _
'                       & "  " & Index _
'                       & ", " & KeyCode _
'                       & ", " & Shift _
'                       & " )", True
'End Sub
'
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'   Select Case KeyCode
'   Case vbKeyReturn, vbKeyUp, vbKeyDown
'      Select Case KeyCode
'      Case vbKeyReturn, vbKeyDown
'         If GetFocus = GridEditor1.hwnd Then Exit Sub
'         SetNextFocus
'      Case vbKeyUp
'         SetPreviousFocus
'      End Select
'   End Select
'End Sub
'
'Private Sub LoadMaster()
'   For pnCtr = 2 To 5
'      txtField(pnCtr).Text = IIf(IsNull(oTrans.Master(pnCtr)), "", oTrans.Master(pnCtr))
'   Next
'
'   txtField(0).Text = Format(oTrans.Master(0), "@@-@@@@@@")
'   txtField(1).Text = Format(oTrans.Master(1), "MMMM DD, YYYY")
'   Label2.Caption = Format(TransStat(oTrans.Master("cTranStat")), ">")
'
'   txtField(7).Text = txtField(0).Text
'   txtField(6).Text = txtField(2).Text
'
'   txtField(7).Tag = txtField(0).Text
'   txtField(6).Tag = txtField(2).Text
'
'   pbLoaded = True
'End Sub
'
'Private Sub LoadDetail()
'   Dim lnCtr As Integer
'
'   With GridEditor1
'      .Rows = IIf(oTrans.ItemCount = 0, 2, oTrans.ItemCount + 1)
'
'      If .Rows > 22 Then
'         .ColWidth(2) = 4400
'         .ColWidth(3) = 2780
'      Else
'         .ColWidth(2) = 4500
'         .ColWidth(3) = 2880
'      End If
'
'      For pnCtr = 0 To oTrans.ItemCount - 1
'         For lnCtr = 1 To 4
'            .TextMatrix(pnCtr + 1, lnCtr) = oTrans.Detail(pnCtr, IIf(lnCtr = 4, lnCtr + 1, lnCtr))
'         Next
'      Next
'   End With
'End Sub
'
'Private Sub txtField_LostFocus(Index As Integer)
'   With txtField(Index)
'      .BackColor = oApp.getColor("EB")
'   End With
'End Sub
'
'Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
'   Dim lsOldProc As String
'
'   lsOldProc = "txtField_Validate"
'   On Error GoTo errProc
'
'   With txtField(Index)
'      Select Case Index
'      Case 6, 7
'         If .Text = "" Then
'            InitEntry
'            GoTo endProc
'         End If
'
'         If .Text <> .Tag Then
'            If oTrans.SearchTransaction(IIf(Index = 7, CodeFormat(oApp.BranchCode, .Text) _
'               , .Text) _
'               , IIf(Index = 7, True, False)) = True Then
'
'               LoadMaster
'               LoadDetail
'            Else
'               InitEntry
'               .SetFocus
'            End If
'         End If
'      End Select
'   End With
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " _
'                       & "  " & Index _
'                       & ", " & Cancel _
'                       & " )", True
'End Sub
'
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
