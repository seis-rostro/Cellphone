VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0#1.0#0"; "xrGridControl.ocx"
Begin VB.Form frmCP_BranchOrder_Reg 
   BorderStyle     =   0  'None
   Caption         =   "Spareparts Branch Order"
   ClientHeight    =   6915
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11700
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6915
   ScaleWidth      =   11700
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   8
      Left            =   10365
      TabIndex        =   19
      Top             =   2790
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "Print"
      AccessKey       =   "Print"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCP_BranchOrder_Reg.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   7
      Left            =   10365
      TabIndex        =   20
      Top             =   3420
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
      Picture         =   "frmCP_BranchOrder_Reg.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   10365
      TabIndex        =   13
      Top             =   900
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
      Picture         =   "frmCP_BranchOrder_Reg.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   10365
      TabIndex        =   14
      Top             =   1530
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
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
      Picture         =   "frmCP_BranchOrder_Reg.frx":166E
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   6
      Left            =   10365
      TabIndex        =   15
      Top             =   2160
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
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
      Picture         =   "frmCP_BranchOrder_Reg.frx":1DE8
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   10365
      TabIndex        =   16
      Top             =   900
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
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
      Picture         =   "frmCP_BranchOrder_Reg.frx":2562
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   10365
      TabIndex        =   21
      Top             =   3420
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
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
      Picture         =   "frmCP_BranchOrder_Reg.frx":2CDC
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   10365
      TabIndex        =   17
      Top             =   1530
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
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
      Picture         =   "frmCP_BranchOrder_Reg.frx":3456
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   10365
      TabIndex        =   18
      Top             =   2160
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Del. Row"
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
      Picture         =   "frmCP_BranchOrder_Reg.frx":3BD0
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   525
      Index           =   1
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   10050
      _ExtentX        =   17727
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
         Index           =   7
         Left            =   5160
         MaxLength       =   50
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   90
         Width           =   4770
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
         Index           =   6
         Left            =   1455
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   90
         Width           =   2265
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Branch Name"
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
         Left            =   3960
         TabIndex        =   2
         Top             =   135
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction No"
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
         Left            =   90
         TabIndex        =   0
         Top             =   135
         Width           =   1320
      End
   End
   Begin xrGridEditor.GridEditor GridEditor1 
      Height          =   4140
      Left            =   90
      TabIndex        =   12
      Tag             =   "et0;eb0;et0;bc2"
      Top             =   2610
      Width           =   10065
      _ExtentX        =   17754
      _ExtentY        =   7303
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
      Object.HEIGHT          =   4140
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
      MOUSEICON       =   "frmCP_BranchOrder_Reg.frx":434A
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
      Height          =   1485
      Index           =   0
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   1080
      Width           =   10065
      _ExtentX        =   17754
      _ExtentY        =   2619
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   1260
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   150
         Width           =   2265
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   1260
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   660
         Width           =   2265
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   5
         Left            =   1260
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   990
         Width           =   5730
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   4
         Left            =   8310
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   990
         Width           =   1620
      End
      Begin VB.Shape Shape4 
         Height          =   330
         Index           =   0
         Left            =   7140
         Top             =   180
         Width           =   2775
      End
      Begin VB.Shape Shape3 
         Height          =   390
         Index           =   0
         Left            =   7110
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
         Left            =   7170
         TabIndex        =   22
         Tag             =   "eb0;et0"
         Top             =   225
         Width           =   2715
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   315
         Left            =   1350
         Tag             =   "et0;ht2"
         Top             =   255
         Width           =   2265
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction No"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   4
         Top             =   210
         Width           =   1095
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transact. Date"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   6
         Top             =   735
         Width           =   1065
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   195
         Index           =   5
         Left            =   90
         TabIndex        =   8
         Top             =   1065
         Width           =   630
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Requested By"
         Height          =   195
         Index           =   6
         Left            =   7140
         TabIndex        =   10
         Top             =   1065
         Width           =   1005
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H8000000F&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   285
         Index           =   0
         Left            =   7170
         Tag             =   "et0;et0"
         Top             =   210
         Width           =   2730
      End
   End
End
Attribute VB_Name = "frmCP_BranchOrder_Reg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Private Const pxeMODULENAME = "frmSP_BranchOrder_Reg"
'
'Private WithEvents oTrans As clsBranchOrder
'Private oSkin As clsFormSkin
'
'Dim pnIndex As Integer, pnCtr As Integer
'Dim pbGridFocus As Boolean, pbEditMode As Boolean
'Dim psTransNox As String
'
'Property Let TransactionNo(lsTransNox As String)
'   psTransNox = lsTransNox
'End Property
'
'Private Sub cmdButton_Click(Index As Integer)
'   Dim lsOldProc As String
'   Dim lnRep As Integer
'
'   lsOldProc = "cmdButton_Click"
'   On Error GoTo errProc
'
'   With GridEditor1
'      Select Case Index
'      Case 0 'Save
'         If .Rows > 2 Then
'            pnCtr = 0
'            Do While pnCtr < .Rows
'               If Trim(.TextMatrix(pnCtr, 1)) = "" Then
'                  .Row = pnCtr
'                  If oTrans.DeleteDetail(.Row - 1) Then .DeleteRow
'               Else
'                  pnCtr = pnCtr + 1
'               End If
'            Loop
'
'            If .Rows > 18 Then
'               .ColWidth(2) = 4400
'               .ColWidth(3) = 2580
'            Else
'               .ColWidth(2) = 4500
'               .ColWidth(3) = 2680
'            End If
'         End If
'
'         If isEntryOK Then
'            If oTrans.SaveTransaction = True Then
'               MsgBox "Transaction Saved Successfully!!!", vbInformation, "Notice"
'               InitButton xeModeReady
'               txtField(7).SetFocus
'               pbEditMode = False
'            Else
'               MsgBox "Unable to Save Transaction!!!", vbCritical, "Warning"
'            End If
'         End If
'      Case 1
'         If pbGridFocus Then
'            If oTrans.SearchDetail(.Row - 1, 1) Then .Col = 1
'            .Refresh
'            .SetFocus
'         End If
'      Case 2
'         If .Rows > 2 Then
'            If oTrans.DeleteDetail(.Row - 1) Then .DeleteRow
'
'            If .Rows > 16 Then
'               .ColWidth(2) = 4400
'               .ColWidth(3) = 2580
'            Else
'               .ColWidth(2) = 4500
'               .ColWidth(3) = 2680
'            End If
'         End If
'      Case 3
'         lnRep = MsgBox("Transaction is in Update Mode!!!" & vbCrLf & _
'                        "Do you want to Cancel Transaction!!!", vbYesNo + vbQuestion, "Confirm")
'
'         If lnRep = vbYes Then
'            If oTrans.SearchTransaction(oTrans.Master(0), True) = True Then
'               LoadMaster
'               LoadDetail
'            Else
'               ClearFields
'            End If
'
'            InitButton xeModeReady
'            txtField(7).SetFocus
'            pbEditMode = False
'         End If
'      Case 4
'         If oTrans.SearchTransaction() Then
'            LoadMaster
'            LoadDetail
'         Else
'            If txtField(0).Text = "" Then ClearFields
'         End If
'
'         txtField(pnIndex).SetFocus
'         .Refresh
'      Case 5
'         If txtField(0).Text <> "" Then
'            oTrans.UpdateTransaction
'            InitButton xeModeUpdate
'            pbEditMode = True
'            txtField(1).SetFocus
'         Else
'            MsgBox "No Transaction to Update!!!" & vbCrLf & _
'                   "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
'         End If
'      Case 6
'         If txtField(0).Text <> "" Then
'            If oTrans.CancelTransaction = True Then
'               MsgBox "Transaction Cancelled Successfully!!!", vbInformation, "Notice"
'               Label2.Caption = Format(TransStat(oTrans.Master("cTranStat")), ">")
'            Else
'               MsgBox "Unable to Cancel Transaction!!!", vbCritical, "Warning"
'            End If
'         Else
'            MsgBox "No Transaction to Cancel!!!", vbInformation, "Notice"
'         End If
'      Case 7
'         Unload Me
'      Case 8
'         If Not pbEditMode Then
'            If txtField(0).Text <> "" Then
'               lnRep = MsgBox("Print Transaction!!!", vbYesNo + vbQuestion, "Confirm")
'               If lnRep = vbYes Then
'                  If Not PrintTrans Then MsgBox "Unable to Print Transaction!!!", vbCritical, "Warning"
'               End If
'            End If
'         Else
'            MsgBox "Unable to Print Transaction!!!", vbCritical, "Warning"
'         End If
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
'   Set oTrans = New clsBranchOrder
'   Set oTrans.AppDriver = oApp
'   oTrans.InitTransaction
'
'   Set oSkin = New clsFormSkin
'   Set oSkin.AppDriver = oApp
'   Set oSkin.Form = Me
'   oSkin.ApplySkin xeFormTransMaintenance
'
'   InitForm
'   ClearFields
'   InitButton xeModeReady
'   pbEditMode = False
'
'   If Not Trim(psTransNox) = "" Then
'      If oTrans.SearchTransaction(Trim(psTransNox), True) Then
'         LoadMaster
'         LoadDetail
'      End If
'   End If
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
'Private Sub GridEditor1_AddingRow(Cancel As Boolean)
'   With GridEditor1
'      If .TextMatrix(.Row, 1) = "" Then
'         Cancel = True
'      ElseIf CDbl(.TextMatrix(.Row, 4)) = 0 Then
'         Cancel = True
'      End If
'      If Not Cancel Then oTrans.AddDetail
'
'      If .Rows > 16 Then
'         .ColWidth(2) = 4400
'         .ColWidth(3) = 2580
'      End If
'   End With
'End Sub
'
'Private Sub GridEditor1_EditorValidate(Cancel As Boolean)
'   With GridEditor1
'      oTrans.Detail(.Row - 1, IIf(.Col = 4, .Col + 1, .Col)) = .TextMatrix(.Row, .Col)
'   End With
'End Sub
'
'Private Sub GridEditor1_GotFocus()
'   With GridEditor1
'      .EditorBackColor = oApp.getColor("HT1")
'   End With
'   pbGridFocus = True
'End Sub
'
'Private Sub GridEditor1_KeyDown(KeyCode As Integer, Shift As Integer)
'   Dim lsOldProc As String
'
'   lsOldProc = "GridEditor1_KeyDown"
'   On Error GoTo errProc
'
'   If pbEditMode Then
'      If KeyCode = vbKeyF3 Then
'         With GridEditor1
'            If .Col = 1 Or .Col = 2 Then
'               If oTrans.SearchDetail(.Row - 1, .Col, .TextMatrix(.Row, .Col)) Then .Col = 5
'               .Refresh
'               .SetFocus
'            End If
'            KeyCode = 0
'         End With
'      End If
'   End If
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " _
'                       & "  " & KeyCode _
'                       & ", " & Shift _
'                       & " )", True
'End Sub
'
'Private Sub GridEditor1_LostFocus()
'   With GridEditor1
'      .EditorBackColor = oApp.getColor("EB")
'   End With
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
'Private Sub ClearFields()
'   For pnCtr = 0 To 7
'      Select Case pnCtr
'      Case 2, 3
'      Case Else
'         txtField(pnCtr).Text = ""
'         txtField(pnCtr).Tag = ""
'      End Select
'   Next
'
'   Label2.Caption = "UNKNOWN"
'
'   With GridEditor1
'      .Rows = 2
'      .Row = 1
'      .Col = 1
'
'      .ColWidth(2) = 3900
'      .ColWidth(3) = 2680
'
'      'empty row
'      .TextMatrix(1, 1) = ""
'      .TextMatrix(1, 2) = ""
'      .TextMatrix(1, 3) = ""
'      .TextMatrix(1, 4) = 0
'      .TextMatrix(1, 5) = 0
'   End With
' End Sub
'
'Private Sub InitForm()
'   Dim lnCtr As Integer
'
'   With GridEditor1
'      .Cols = 6
'      .Rows = 2
'      .Font = "MS Sans Serif"
'
'      'Column Title
'      .TextMatrix(0, 1) = "Barcode"
'      .TextMatrix(0, 2) = "Description"
'      .TextMatrix(0, 3) = "Model"
'      .TextMatrix(0, 4) = "Qty."
'      .TextMatrix(0, 5) = "App."
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
'      .ColWidth(5) = 600
'
'      .ColAlignment(1) = 1
'      .ColAlignment(2) = 1
'      .ColAlignment(3) = 1
'
'      .ColEnabled(3) = False
'      .ColEnabled(5) = False
'
'      .ColNumberOnly(4) = True
'      .ColFormat(4) = "#,##0"
'      .ColDefault(4) = 0
'      .ColDefault(5) = 0
'
'      .EditorBackColor = oApp.getColor("HT1")
'   End With
'End Sub
'
'Private Sub txtField_GotFocus(Index As Integer)
'   With txtField(Index)
'      If Index = 1 Then .Text = Format(.Text, "MM/DD/YY")
'      .SelStart = 0
'      .SelLength = Len(.Text)
'      .BackColor = oApp.getColor("HT1")
'   End With
'
'   pbGridFocus = False
'   pnIndex = Index
'End Sub
'
'Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'   Dim lsOldProc As String
'
'   lsOldProc = "txtField_KeyDown"
'   On Error GoTo errProc
'
'   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
'      With txtField(Index)
'         Select Case Index
'         Case 6, 7
'            If oTrans.SearchTransaction(IIf(Index = 6, CodeFormat(oApp.BranchCode, .Text), .Text), IIf(Index = 6, True, False)) = True Then
'               LoadMaster
'               LoadDetail
'            Else
'               ClearFields
'            End If
'
'            .SelStart = 0
'            .SelLength = Len(.Text)
'            .SetFocus
'         End Select
'         KeyCode = 0
'      End With
'   End If
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
'   Case vbKeyF8
''      If txtField(0).Text <> "" And oTrans.EditMode = xeModeReady Then
''         If oApp.UserLevel = xeEngineer Then
''            If oTrans.DeleteTransaction Then ClearFields
''         End If
''      End If
'   End Select
'End Sub
'
'Private Sub InitButton(lnStat As Integer)
'   Dim lbShow As Boolean
'
'   lbShow = IIf(lnStat = 0, False, True)
'   xrFrame1(1).Enabled = Not lbShow
'   cmdButton(4).Visible = Not lbShow
'   cmdButton(5).Visible = Not lbShow
'   cmdButton(6).Visible = Not lbShow
'   cmdButton(7).Visible = Not lbShow
'
'   cmdButton(0).Visible = lbShow
'   cmdButton(1).Visible = lbShow
'   cmdButton(2).Visible = lbShow
'   cmdButton(3).Visible = lbShow
'
'   With GridEditor1
'      .ColEnabled(1) = lbShow
'      .ColEnabled(2) = lbShow
'      .ColEnabled(4) = lbShow
'   End With
'
'   xrFrame1(0).Enabled = lbShow
'End Sub
'
'Private Sub LoadMaster()
'   For pnCtr = 0 To 7
'      Select Case pnCtr
'      Case 0, 6
'         txtField(pnCtr).Text = Format(oTrans.Master("sTransNox"), "@@@@-@@@@@@")
'         txtField(pnCtr).Tag = txtField(pnCtr).Text
'      Case 1
'         txtField(pnCtr).Text = Format(oTrans.Master(pnCtr), "MMMM DD, YYYY")
'      Case 2
'         txtField(pnCtr + 5).Text = oTrans.Master(pnCtr)
'         txtField(pnCtr + 5).Tag = txtField(pnCtr + 5).Text
'      Case 2, 3
'      Case 7
'         txtField(pnCtr).Text = oTrans.Master(10)
'         txtField(pnCtr).Tag = oTrans.Master(10)
'      Case Else
'         txtField(pnCtr).Text = IIf(IsNull(oTrans.Master(pnCtr)), "", oTrans.Master(pnCtr))
'      End Select
'   Next
'
'   Label2.Caption = Format(TransStat(oTrans.Master("cTranStat")), ">")
'End Sub
'
'Private Sub LoadDetail()
'   Dim lnRow As Integer
'   Dim lnCol As Integer
'
'   With GridEditor1
'      .Rows = IIf(oTrans.ItemCount = 0, 2, oTrans.ItemCount + 1)
'
'      If .Rows > 16 Then
'         .ColWidth(2) = 3800
'         .ColWidth(3) = 2580
'      Else
'         .ColWidth(2) = 3900
'         .ColWidth(3) = 2680
'      End If
'
'      For lnRow = 0 To oTrans.ItemCount - 1
'         For lnCol = 1 To 4
'            .TextMatrix(lnRow + 1, lnCol) = oTrans.Detail(lnRow, IIf(lnCol = 4, lnCol + 1, lnCol))
'         Next
'         .TextMatrix(lnRow + 1, 5) = 0
'         If oTrans.Master("cTranStat") = xeStatePosted Then .TextMatrix(lnRow + 1, 5) = oTrans.Detail(lnRow, "nApproved")
'      Next
'   End With
'End Sub
'
'Private Function isEntryOK() As Boolean
'   isEntryOK = False
'
'   If txtField(2).Text = "" Then
'      MsgBox "Invalid Branch Detected!!!" & vbCrLf & _
'             "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
'      txtField(2).SetFocus
'      GoTo endProc
'   End If
'
'   With GridEditor1
'      If .TextMatrix(1, 1) = "" Then
'         MsgBox "Detail is required!!!" & vbCrLf & _
'                "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
'         .SetFocus
'         .Row = 1
'         .Col = 1
'         GoTo endProc
'      End If
'   End With
'
'   isEntryOK = True
'
'endProc:
'   Exit Function
'End Function
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
'
'      Select Case Index
'      Case 1
'         If Not IsDate(.Text) Then .Text = oApp.ServerDate
'         .Text = Format(.Text, "MMMM DD, YYYY")
'      End Select
'
'      If Index < 6 Then oTrans.Master(Index) = .Text
'   End With
'End Sub
'
'Public Function PrintTrans() As Boolean
'   Dim loreport As frmRepViewer
'   Dim lnRep As Integer
'   Dim lsBarrCode As String
'
'   Dim lors As ADODB.Recordset
'   Dim lrs As ADODB.Recordset
'   Dim lnCtr As Integer
'   Dim lsOldProc As String
'
'   lsOldProc = "PrintTrans"
'   On Error GoTo errProc
'
'   PrintTrans = True
'
'   Set lors = New ADODB.Recordset
'
'   lors.Fields.Append "nQuantity", adInteger, 3
'   lors.Fields.Append "sModel", adVarChar, 50
''   lors.Fields.Append "sColor", adVarChar, 50
'   lors.Fields.Append "sDescription", adVarChar, 50
'   lors.Fields.Append "sBarrCode", adVarChar, 25
'   lors.Open
'
'   With GridEditor1
'      For lnCtr = 1 To .Rows - 1
'         lors.AddNew
'         lors("nQuantity").Value = .TextMatrix(lnCtr, 4)
'         lors("sModel").Value = .TextMatrix(lnCtr, 3)
''         lors("sColor").Value = .TextMatrix(lnCtr, 4)
'         lors("sDescription").Value = .TextMatrix(lnCtr, 2)
'         lors("sBarrCode").Value = .TextMatrix(lnCtr, 1)
'      Next
'   End With
'
'   ' assign important info to the report
'   Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\Purchase.rpt")
'   oReport.DiscardSavedData
'   oReport.FieldMappingType = crAutoFieldMapping
'   oReport.Database.SetDataSource lors
'
'   Set lrs = New ADODB.Recordset
'   lrs.Open "Select" _
'               & "  CONCAT(b.sTownName, ', ', c.sProvName, ' ', b.sZippCode) xAddressx" _
'            & " From Branch a" _
'               & ", TownCity b" _
'                  & " Left Join Province c" _
'                     & " On b.sProvIDxx = c.sProvIDxx" _
'            & " Where a.sTownIDxx = b.sTownIDxx" _
'               & " And a.sBranchCd = " & Left(oTrans.Master("sTransNox"), 2) _
'            , oApp.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
'
''   If Not lrs.EOF Then oReport.Sections("PH").ReportObjects("txtDeliver").SetText "           " & txtField(4).Text & vbCrLf & " " & lrs("xAddressx")
'    oReport.Sections("PH").ReportObjects("txtDeliver").SetText "UEMI Main"
'   oReport.Sections("PH").ReportObjects("txtSupplier").SetText "Suzuki Phils Inc."
''   oReport.Sections("PH").ReportObjects("txtTerm").SetText txtField(6).Text
''   oReport.Sections("PH").ReportObjects("txtDDate").SetText txtField(7).Text
'   oReport.Sections("PH").ReportObjects("txtDate").SetText txtField(1).Text
'   oReport.Sections("PF").ReportObjects("txtUserRpt").SetText oApp.UserName
'
'   lnRep = MsgBox("Do you want to create TXT file???", vbYesNo + vbQuestion, "Confirm")
'   If lnRep = vbYes Then
'      With lors
'         lnCtr = 1
'         lors.MoveFirst
'         Do
'            If Len(lors("sBarrCode")) = 15 Then
'               lsBarrCode = lors("sBarrCode")
'            ElseIf Len(lors("sBarrCode")) < 15 Then
'               lsBarrCode = lors("sBarrCode") & Replace(FormatNumber(0, 15 - Len(lors("sBarrCode")), vbFalse), ".", "")
'            Else
'               lsBarrCode = Left(lors("sBarrCode"), 15)
'            End If
'            LogPOEntry Format(lnCtr, "000") & lsBarrCode & Format(lors("nQuantity"), "000000")
'            lnCtr = lnCtr + 1
'            lors.MoveNext
'         Loop Until lors.EOF
'      End With
'   End If
'
'   Set loreport = New frmRepViewer
'   Set loreport.ReportSource = oReport
'   loreport.Show
'
'   lors.Close
'   lrs.Close
'
'endProc:
'   oTrans.CloseTransaction (oTrans.Master(0))
'   Set oReport = Nothing
'   Set lors = Nothing
'   Set lrs = Nothing
'   Exit Function
'errProc:
'   PrintTrans = False
'   ShowError lsOldProc & "( " & " )"
'End Function
'Private Sub LogPOEntry(ByVal lsDetail As String)
'   Open "C:\GGC_Systems\SPPurchased\" & oTrans.Master("sTransNox") & ".txt" For Append As #1
'   Write #1, lsDetail
'   Close #1
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