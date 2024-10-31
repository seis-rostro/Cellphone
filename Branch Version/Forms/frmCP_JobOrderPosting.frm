VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0"; "xrGridControl.ocx"
Begin VB.Form frmCP_JobOrderPosting 
   BorderStyle     =   0  'None
   Caption         =   "Warranty to Service Center Posting"
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10995
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7200
   ScaleWidth      =   10995
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame2 
      Height          =   3150
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   3990
      Width           =   9315
      _ExtentX        =   16431
      _ExtentY        =   5556
      BackColor       =   12632256
      Enabled         =   0   'False
      BorderStyle     =   1
      Begin xrGridEditor.GridEditor GridEditor1 
         Height          =   3000
         Left            =   45
         TabIndex        =   28
         Tag             =   "et0;eb0;et0;bc2"
         Top             =   60
         Width           =   9180
         _ExtentX        =   16193
         _ExtentY        =   5292
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
         Object.HEIGHT          =   3000
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
         MOUSEICON       =   "frmCP_JobOrderPosting.frx":0000
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
   End
   Begin xrControl.xrFrame xrFrame1 
      DragMode        =   1  'Automatic
      Height          =   2910
      Index           =   0
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   1065
      Width           =   9315
      _ExtentX        =   16431
      _ExtentY        =   5133
      BackColor       =   12632256
      Enabled         =   0   'False
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   6
         Left            =   6510
         TabIndex        =   21
         Top             =   705
         Width           =   2625
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   645
         Index           =   4
         Left            =   960
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   1365
         Width           =   4455
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   8
         Left            =   6510
         TabIndex        =   25
         Top             =   1365
         Width           =   1830
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   9
         Left            =   6510
         TabIndex        =   27
         Top             =   1695
         Width           =   1830
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   5
         Left            =   6510
         MaxLength       =   25
         TabIndex        =   19
         Top             =   375
         Width           =   2625
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   7
         Left            =   6510
         TabIndex        =   23
         Top             =   1035
         Width           =   1830
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   3
         Left            =   960
         TabIndex        =   13
         Top             =   1035
         Width           =   4455
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   960
         TabIndex        =   9
         Top             =   705
         Width           =   1830
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   3585
         MaxLength       =   10
         TabIndex        =   11
         Top             =   705
         Width           =   1830
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Index           =   0
         Left            =   975
         TabIndex        =   7
         Top             =   210
         Width           =   1620
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   645
         Index           =   10
         Left            =   960
         MaxLength       =   512
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   2025
         Width           =   8175
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrpt."
         Height          =   195
         Index           =   8
         Left            =   5835
         TabIndex        =   20
         Top             =   750
         Width           =   600
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   195
         Index           =   4
         Left            =   345
         TabIndex        =   14
         Top             =   1395
         Width           =   570
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Color"
         Height          =   195
         Index           =   3
         Left            =   6075
         TabIndex        =   26
         Top             =   1755
         Width           =   360
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IMEI/Barcode"
         Height          =   195
         Index           =   7
         Left            =   5430
         TabIndex        =   18
         Top             =   420
         Width           =   1005
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   16
         Top             =   2055
         Width           =   675
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "S. Center"
         Height          =   195
         Index           =   6
         Left            =   255
         TabIndex        =   12
         Top             =   1080
         Width           =   660
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "J.O. No."
         Height          =   195
         Index           =   18
         Left            =   2910
         TabIndex        =   10
         Top             =   735
         Width           =   585
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Brand"
         Height          =   195
         Index           =   5
         Left            =   6015
         TabIndex        =   22
         Top             =   1095
         Width           =   420
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
         Height          =   195
         Index           =   9
         Left            =   6000
         TabIndex        =   24
         Top             =   1425
         Width           =   435
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trans. Date"
         Height          =   195
         Index           =   1
         Left            =   75
         TabIndex        =   8
         Top             =   720
         Width           =   840
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trans. #"
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
         Left            =   90
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   315
         Left            =   1080
         Tag             =   "et0;ht2"
         Top             =   315
         Width           =   1620
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   9630
      TabIndex        =   33
      Top             =   3045
      Width           =   1275
      _ExtentX        =   2249
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
      Picture         =   "frmCP_JobOrderPosting.frx":001C
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   9630
      TabIndex        =   29
      Top             =   525
      Width           =   1275
      _ExtentX        =   2249
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
      Picture         =   "frmCP_JobOrderPosting.frx":0796
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   9630
      TabIndex        =   31
      Top             =   1785
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   1058
      Caption         =   "&Repaired"
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
      Picture         =   "frmCP_JobOrderPosting.frx":0F10
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   525
      Index           =   1
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   525
      Width           =   9315
      _ExtentX        =   16431
      _ExtentY        =   926
      BackColor       =   12632256
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
         Index           =   13
         Left            =   3600
         MaxLength       =   50
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   90
         Width           =   3210
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
         Index           =   12
         Left            =   975
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   90
         Width           =   1620
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
         Index           =   11
         Left            =   7350
         MaxLength       =   50
         TabIndex        =   5
         Text            =   "Text"
         Top             =   90
         Width           =   1800
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Company"
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
         Left            =   2760
         TabIndex        =   2
         Top             =   135
         Width           =   780
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&J.O. No."
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
         Left            =   135
         TabIndex        =   0
         Top             =   135
         Width           =   720
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
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
         Index           =   2
         Left            =   6870
         TabIndex        =   4
         Top             =   135
         Width           =   465
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   9630
      TabIndex        =   32
      Top             =   2415
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   1058
      Caption         =   "Replaced"
      AccessKey       =   "Replaced"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCP_JobOrderPosting.frx":168A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   9630
      TabIndex        =   30
      Top             =   1155
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   1058
      Caption         =   "Received"
      AccessKey       =   "Received"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCP_JobOrderPosting.frx":1E04
   End
End
Attribute VB_Name = "frmCP_JobOrderPosting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Private Const pxeMODULENAME = "frmCP_JobOrderPosting"
'
'Private WithEvents oTrans As clsJobOrder
'Private oSkin As clsFormSkin
'
'Dim pbLoad As Boolean
'Dim pnCtr As Integer, pnIndex As Integer
'
'Private Sub cmdButton_Click(Index As Integer)
'   Dim lsOldProc As String
'   Dim lnRep As String
'
'   lsOldProc = "cmdButton_Click"
'   ''On Error GoTo errProc
'
'   txtField_LostFocus pnIndex
'   Select Case Index
'   Case 0
'      If oTrans.SearchForwardedJO() Then
'         LoadMaster
'         LoadDetail
'         pbLoad = True
'         txtField(11).SetFocus
'      Else
'         pbLoad = False
'         If txtField(0).Text <> "" Then pbLoad = True
'         txtField(13).SetFocus
'      End If
'   Case 1
'      If pbLoad Then
'         lnRep = MsgBox("Post Transaction!!!", vbYesNo + vbQuestion, "Confrim")
'
'         If lnRep = vbYes Then
'            If oTrans.Master("dTransact") > CDate(txtField(11).Text) Then
'               MsgBox "Invalid Receiving Date!!!" & vbCrLf & _
'                        "Please verify your entry then try again!!!", vbCritical, "Warning"
'            End If
'
'            If lnRep = vbYes Then
'               If Not oTrans.Repaired(CDate(txtField(11).Text)) Then
'                  MsgBox "Unable to Post Transaction!!!", vbCritical, "Warning"
'               Else
'                  MsgBox "Transaction Post Successfully!!!", vbInformation, "Confirm"
'                  ClearFields
'               End If
'            End If
'         End If
'      Else
'         MsgBox "Unable to Post Transaction!!!" & vbCrLf & _
'                "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
'      End If
'   Case 2
'      Unload Me
'   Case 3
'      If pbLoad Then
'         If oTrans.Replaced(CDate(txtField(11).Text)) Then
'            MsgBox "Transaction Post Successfully!!!", vbInformation, "Confirm"
'            ClearFields
'         End If
'      End If
'   Case 4
'      If pbLoad Then
'         If oTrans.Received(CDate(txtField(11).Text)) Then
'            MsgBox "Transaction Save Successfully!!!", vbInformation, "Confirm"
'            ClearFields
'         End If
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
'   oApp.MenuName = Me.Tag
'   Me.ZOrder 0
'
'   GridEditor1.Refresh
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
'Private Sub Form_Load()
'   Dim lsOldProc As String
'
'   lsOldProc = "Form_Load"
'   ''On Error GoTo errProc
'
'   CenterChildForm mdiMain, Me
'
'   Set oTrans = New clsJobOrder
'   Set oTrans.AppDriver = oApp
'
'   oTrans.JOStatus = xeJOStateForwarded
'   oTrans.InitTransaction
'   oTrans.NewTransaction
'
'   Set oSkin = New clsFormSkin
'   Set oSkin.AppDriver = oApp
'   Set oSkin.Form = Me
'   oSkin.ApplySkin xeFormTransMaintenance
'
'   InitGrid
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
'End Sub
'
'Private Sub oTrans_DetailRetrieved(ByVal Index As Integer)
'   With GridEditor1
'      .TextMatrix(.Row, Index) = oTrans.Detail(.Row - 1, Index)
'   End With
'End Sub
'
'Private Sub oTrans_MasterRetrieved(ByVal Index As Integer)
'   With oTrans
'      txtField(Index).Text = IFNull(.Master(Index), "")
'   End With
'End Sub
'
'Private Sub ClearFields()
'   For pnCtr = 0 To 13
'      Select Case pnCtr
'      Case 1, 11
'         txtField(pnCtr).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
'      Case Else
'         txtField(pnCtr).Text = ""
'         txtField(pnCtr).Tag = ""
'      End Select
'   Next
'
'   With GridEditor1
'      .Rows = 2
'
'      'empty row
'      .TextMatrix(1, 1) = ""
'      .TextMatrix(1, 2) = ""
'      .TextMatrix(1, 3) = ""
'   End With
'End Sub
'
'Private Sub InitGrid()
'   With GridEditor1
'      .Rows = 2
'      .Cols = 4
'      .Font = "MS Sans Serif"
'
'      'column title
'      .TextMatrix(0, 1) = "Description"
'      .TextMatrix(0, 2) = "Refer No"
'      .TextMatrix(0, 3) = "New Refer No"
'      .Row = 0
'
'      'column alignment
'      For pnCtr = 0 To .Cols - 1
'         .Col = pnCtr
'         .CellFontBold = True
'         .CellAlignment = 3
'      Next
'
'      'column width
'      .ColWidth(0) = 330
'      .ColWidth(1) = 3700
'      .ColWidth(2) = 2500
'      .ColWidth(3) = 2500
'
'      .ColAlignment(1) = 1
'      .ColAlignment(2) = 1
'      .ColAlignment(3) = 1
'
'      .EditorBackColor = oApp.getColor("HT1")
'
'      .Row = 1
'      .Col = 1
'   End With
'End Sub
'
'Private Sub LoadMaster()
'   For pnCtr = 0 To 10
'      Select Case pnCtr
'      Case 0
'         txtField(pnCtr).Text = Format(oTrans.Master(pnCtr), "@@@@-@@@@@@")
'      Case 2
'         txtField(pnCtr).Text = oTrans.Master(pnCtr)
'         txtField(12).Text = txtField(pnCtr).Text
'         txtField(12).Tag = txtField(12).Text
'      Case 3
'         txtField(pnCtr).Text = oTrans.Master(pnCtr)
'         txtField(13).Text = txtField(pnCtr).Text
'         txtField(13).Tag = txtField(13).Text
'      Case 1
'         txtField(pnCtr).Text = Format(oTrans.Master(pnCtr), "MMMM DD, YYYY")
'      Case Else
'         txtField(pnCtr).Text = IIf(IsNull(oTrans.Master(pnCtr)), "", oTrans.Master(pnCtr))
'      End Select
'   Next
'   pbLoad = True
'End Sub
'
'Private Sub LoadDetail()
'   With GridEditor1
'      .Rows = IIf(oTrans.ItemCount = 0, 2, oTrans.ItemCount + 1)
'
'      For pnCtr = 1 To .Rows - 1
'         If Not IsNull(oTrans.Detail(0, 1)) Then
'            .TextMatrix(pnCtr, 1) = IIf(oTrans.Detail(pnCtr - 1, "sDescript") = "", "", oTrans.Detail(pnCtr - 1, "sDescript"))
'            .TextMatrix(pnCtr, 2) = IIf(oTrans.Detail(pnCtr - 1, "sReferNox") = "", "", oTrans.Detail(pnCtr - 1, "sReferNox"))
'         Else
'            .TextMatrix(pnCtr, 1) = ""
'            .TextMatrix(pnCtr, 2) = ""
'         End If
'      Next
'   End With
'End Sub
'
'Private Sub txtField_GotFocus(Index As Integer)
'   With txtField(Index)
'      If Index = 11 Then .Text = Format(.Text, "MM/DD/YYYY")
'      .SelStart = 0
'      .SelLength = Len(.Text)
'      .BackColor = oApp.getColor("HT1")
'   End With
'
'   pnIndex = Index
'End Sub
'
'Private Sub txtField_LostFocus(Index As Integer)
'   With txtField(Index)
'      .BackColor = oApp.getColor("EB")
'   End With
'End Sub
'
'Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
'   With txtField(Index)
'      .Text = TitleCase(.Text)
'      Select Case Index
'      Case 12, 13
'         If .Text = "" Then
'            ClearFields
'            Exit Sub
'         End If
'
'         If Trim(LCase(.Text)) <> Trim(LCase(.Tag)) Then
'            If oTrans.SearchForwardedJO(.Text, IIf(Index = 12, True, False)) Then
'               LoadMaster
'               LoadDetail
'            Else
'               ClearFields
'               .SetFocus
'            End If
'         End If
'      Case 11
'         If Not IsDate(.Text) Then .Text = oApp.ServerDate
'         .Text = Format(.Text, "MMMM DD, YYYY")
'      End Select
'   End With
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