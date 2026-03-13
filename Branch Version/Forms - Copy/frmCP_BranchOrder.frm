VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0"; "xrGridControl.ocx"
Begin VB.Form frmCP_BranchOrder 
   BorderStyle     =   0  'None
   Caption         =   "Branch Order"
   ClientHeight    =   6885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11610
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6885
   ScaleMode       =   0  'User
   ScaleWidth      =   11610
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   2175
      Left            =   1575
      Tag             =   "wt0;fb0"
      Top             =   525
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   3836
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   3930
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   1005
         Width           =   5865
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   3930
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   675
         Width           =   5865
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   5
         Left            =   1305
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   1005
         Width           =   1920
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   660
         Index           =   4
         Left            =   1305
         MultiLine       =   -1  'True
         TabIndex        =   11
         Text            =   "frmCP_BranchOrder.frx":0000
         Top             =   1335
         Width           =   8490
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   1305
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   675
         Width           =   1920
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
         Left            =   1305
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   150
         Width           =   1920
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   195
         Index           =   3
         Left            =   3270
         TabIndex        =   8
         Top             =   1050
         Width           =   570
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transact. Date"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   2
         Top             =   735
         Width           =   1065
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Requested By"
         Height          =   195
         Index           =   6
         Left            =   210
         TabIndex        =   6
         Top             =   1050
         Width           =   1005
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Branch"
         Height          =   195
         Index           =   2
         Left            =   3270
         TabIndex        =   4
         Top             =   735
         Width           =   525
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   195
         Index           =   5
         Left            =   585
         TabIndex        =   10
         Top             =   1395
         Width           =   630
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transact. No."
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
         TabIndex        =   0
         Top             =   210
         Width           =   1185
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   315
         Left            =   1395
         Tag             =   "et0;ht2"
         Top             =   255
         Width           =   1920
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   90
      TabIndex        =   15
      Top             =   4050
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
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
      Picture         =   "frmCP_BranchOrder.frx":0016
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   90
      TabIndex        =   13
      Top             =   2790
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
      Picture         =   "frmCP_BranchOrder.frx":0790
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   90
      TabIndex        =   14
      Top             =   3420
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
      Picture         =   "frmCP_BranchOrder.frx":0F0A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   90
      TabIndex        =   16
      Top             =   4680
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
      Picture         =   "frmCP_BranchOrder.frx":1684
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   90
      TabIndex        =   17
      Top             =   3420
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
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
      Picture         =   "frmCP_BranchOrder.frx":1DFE
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   90
      TabIndex        =   19
      Top             =   4680
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
      Picture         =   "frmCP_BranchOrder.frx":2578
   End
   Begin xrGridEditor.GridEditor GridEditor1 
      Height          =   4020
      Left            =   1575
      TabIndex        =   12
      Tag             =   "et0;eb0;et0;bc2"
      Top             =   2745
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   7091
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
      Object.HEIGHT          =   4020
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
      MOUSEICON       =   "frmCP_BranchOrder.frx":2CF2
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
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   6
      Left            =   90
      TabIndex        =   18
      Top             =   4050
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
      Picture         =   "frmCP_BranchOrder.frx":2D0E
   End
End
Attribute VB_Name = "frmCP_BranchOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Private Const pxeMODULENAME = "frmCP_BranchOrder"
'
'Private WithEvents oTrans As clsBranchOrder
'Private oSkin As clsFormSkin
'
'Dim pnCtr As Integer, pnIndex As Integer
'Dim pbGridFocus As Boolean, pbSave As Boolean
'
'Private Sub cmdButton_Click(Index As Integer)
'   Dim lsOldProc As String
'   Dim lnRep As Integer
'
'   lsOldProc = "cmdButton_Click"
'   'On Error GoTo errProc
'
'   txtField_LostFocus pnIndex
'   With GridEditor1
'      GridEditor1_EditorValidate False
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
'            .ColWidth(3) = 3900
'            If .Rows > 20 Then .ColWidth(3) = 3800
'         End If
'
'         If isEntryOK Then
'            If oTrans.SaveTransaction = True Then
'               MsgBox "Transaction Saved Successfully!!!", vbInformation, "Notice"
'               InitButton xeModeReady
'
'               lnRep = MsgBox("Print Transaction!!!", vbYesNo + vbQuestion, "Confirm")
'               If lnRep = vbYes Then
'                  If Not PrintTrans Then MsgBox "Unable to Print Transaction!!!", vbCritical, "Warning"
'               End If
'               pbSave = True
'            Else
'               MsgBox "Unable to Save Transaction!!!", vbCritical, "Warning"
'            End If
'         End If
'      Case 1 'Search
'         If pbGridFocus Then
'            If oTrans.SearchDetail(.Row - 1, 1) Then .Col = 1
'            .Refresh
'            .SetFocus
'         End If
'      Case 2 'Delete row
'         If .Rows > 2 Then
'            If oTrans.DeleteDetail(.Row - 1) Then .DeleteRow
'
'            If .Rows > 20 Then
'               .ColWidth(2) = 4400
'               .ColWidth(3) = 2780
'            Else
'               .ColWidth(2) = 4500
'               .ColWidth(3) = 2880
'            End If
'         End If
'      Case 3 'Cancel
'         lnRep = MsgBox("Transaction is in Update Mode!!!" & vbCrLf & _
'                        "Do you want to Cancel Transaction!!!", vbYesNo + vbQuestion, "Confirm")
'
'         If lnRep = vbYes Then
'            oTrans.NewTransaction
'            InitButton xeModeReady
'            ClearFields
'         Else
'            txtField(pnIndex).SetFocus
'         End If
'         pbSave = False
'      Case 4 'News
'         oTrans.NewTransaction
'         InitButton xeModeAddNew
'         txtField(1).SetFocus
'         ClearFields
'      Case 5 'Close
'         Unload Me
'      Case 6 'Print
'         If pbSave Then
'            lnRep = MsgBox("Print Transaction!!!", vbYesNo + vbQuestion, "Confirm")
'            If lnRep = vbYes Then
'               If Not PrintTrans Then MsgBox "Unable to Print Transaction!!!", vbCritical, "Warning"
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
'   'On Error GoTo errProc
'
'   CenterChildForm mdiMain, Me
'
'   Set oTrans = New clsBranchOrder
'   Set oTrans.AppDriver = oApp
'
'   oTrans.InitTransaction
'   oTrans.NewTransaction
'
'   Set oSkin = New clsFormSkin
'   Set oSkin.AppDriver = oApp
'   Set oSkin.Form = Me
'   oSkin.ApplySkin xeFormTransaction
'
'   InitGrid
'   ClearFields
'   InitButton xeModeAddNew
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
'      If .Rows > 20 Then
'         .ColWidth(2) = 4400
'         .ColWidth(3) = 2780
'      End If
'   End With
'End Sub
'
'Private Sub GridEditor1_EditorValidate(Cancel As Boolean)
'   With GridEditor1
'      Select Case .Col
'      Case 4
'         oTrans.Detail(.Row - 1, "nQuantity") = .TextMatrix(.Row, .Col)
'      Case Else
'         oTrans.Detail(.Row - 1, .Col) = .TextMatrix(.Row, .Col)
'      End Select
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
'   'On Error GoTo errProc
'
'   With GridEditor1
'      If KeyCode = vbKeyF3 Then
'         If oTrans.SearchDetail(.Row - 1, .Col, .TextMatrix(.Row, .Col)) Then .Col = 4
'         KeyCode = 0
'      End If
'   End With
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
'      Select Case Index
'      Case 6
'         .TextMatrix(.Row, 4) = oTrans.Detail(.Row - 1, "nQuantity")
'      Case Else
'         .TextMatrix(.Row, Index) = oTrans.Detail(.Row - 1, Index)
'      End Select
'   End With
'End Sub
'
'Private Sub oTrans_MasterRetrieved(ByVal Index As Integer)
'   txtField(Index).Text = oTrans.Master(Index)
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
'   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
'      With txtField(Index)
'         If KeyCode = vbKeyF3 Then
'            oTrans.SearchMaster Index, .Text
'            If .Text <> "" Then SetNextFocus
'         Else
'            If .Text <> "" Then oTrans.SearchMaster Index, .Text
'         End If
'      End With
'      KeyCode = 0
'   End If
'End Sub
'
'Private Sub InitButton(lnStat As Integer)
'   Dim lbShow As Boolean
'
'   lbShow = IIf(lnStat = 0, False, True)
'   cmdButton(4).Visible = Not lbShow
'   cmdButton(5).Visible = Not lbShow
'
'   cmdButton(0).Visible = lbShow
'   cmdButton(1).Visible = lbShow
'   cmdButton(2).Visible = lbShow
'   cmdButton(3).Visible = lbShow
'   xrFrame1.Enabled = lbShow
'
'   With GridEditor1
'      .ColEnabled(1) = lbShow
'      .ColEnabled(2) = lbShow
'      .ColEnabled(4) = lbShow
'   End With
'
'   If Not lbShow Then cmdButton(4).SetFocus
'End Sub
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
'      .ColWidth(2) = 3150
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
'
'      .EditorBackColor = oApp.getColor("HT1")
'   End With
'End Sub
'
'Private Sub ClearFields()
'   For pnCtr = 0 To txtField.Count - 1
'      Select Case pnCtr
'      Case 0
'         txtField(pnCtr).Text = Format(oTrans.Master(pnCtr), "@@@@-@@@@@@")
'      Case 1
'         txtField(pnCtr).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
'      Case Else
'         txtField(pnCtr).Text = ""
'      End Select
'   Next
'
'   With GridEditor1
'      .Rows = 2
'      .Row = 1
'      .Col = 1
'
'      .TextMatrix(1, 1) = ""
'      .TextMatrix(1, 2) = ""
'      .TextMatrix(1, 3) = ""
'      .TextMatrix(1, 4) = 0
'
'      .ColWidth(3) = 3900
'      '2880
'   End With
'End Sub
'
'Private Function isEntryOK() As Boolean
'   isEntryOK = False
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
'      If Index = 1 Then
'         If Not IsDate(.Text) Then .Text = oApp.ServerDate
'         .Text = Format(.Text, "MMMM DD, YYYY")
'      End If
'
'      oTrans.Master(Index) = .Text
'   End With
'End Sub
'
'Public Function PrintTrans() As Boolean
'   Dim lrs As New ADODB.Recordset
'   Dim lnCtr As Integer
'   Dim lsOldProc As String
'
'   lsOldProc = "printTrans"
'   'On Error GoTo errProc
'
'   PrintTrans = False
'
'   Set lrs = New ADODB.Recordset
'
'   lrs.Fields.Append "nEntryNo", adInteger, 3
'   lrs.Fields.Append "sBarrCode", adVarChar, 23
'   lrs.Fields.Append "sDescription", adVarChar, 60
'   lrs.Fields.Append "sModel", adVarChar, 30
'   lrs.Fields.Append "nQuantity", adInteger, 5
'   lrs.Fields.Append "nUnitPrice", adDouble, 20
'   lrs.Fields.Append "nTotal", adDouble, 20
'   lrs.Open
'
'   With oTrans
'      For lnCtr = 0 To .ItemCount - 1
'         lrs.AddNew
'         lrs("nEntryNo").Value = .Detail(lnCtr, 0)
'         lrs("sBarrCode").Value = .Detail(lnCtr, 1)
'         lrs("sDescription").Value = .Detail(lnCtr, 2)
'         lrs("sModel").Value = .Detail(lnCtr, 3)
'         lrs("nQuantity").Value = .Detail(lnCtr, 5)
'      Next
'   End With
'
'   ' assign important info to the report
'   Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\StockOrder.rpt")
'   oReport.DiscardSavedData
'   oReport.FieldMappingType = crAutoFieldMapping
'   oReport.Database.SetDataSource lrs
'
'   oReport.Sections("RHa").ReportObjects("txtRefNo").SetText "SP" & "-" & Right(oTrans.Master("sTransNox"), 8)
'   oReport.Sections("RHa").ReportObjects("txtDate").SetText txtField(1).Text
'   oReport.Sections("RHb").ReportObjects("txtToBranch").SetText txtField(2).Text
'   oReport.Sections("RHb").ReportObjects("txtToAddress").SetText oTrans.Master("xAddressx")
'   oReport.Sections("RF").ReportObjects("txtRemarks").SetText txtField(5).Text
'   oReport.Sections("RF").ReportObjects("txtPrepared").SetText oApp.UserName
'   oReport.Sections("RF").ReportObjects("txtApproved").SetText txtField(4).Text
'   oReport.PrintOutEx False, 1
'   lrs.Close
'   PrintTrans = True
'
'endProc:
'   oTrans.CloseTransaction (oTrans.Master(0))
'   Set oReport = Nothing
'   Exit Function
'errProc:
'   ShowError lsOldProc & "( " & " )"
'End Function
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
