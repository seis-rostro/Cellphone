VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0"; "xrGridControl.ocx"
Begin VB.Form frmSuppliesRequestApproval 
   BorderStyle     =   0  'None
   Caption         =   "Supplies Request Approval"
   ClientHeight    =   7395
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11310
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7395
   ScaleWidth      =   11310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   75
      TabIndex        =   0
      Top             =   2460
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
      Picture         =   "frmSuppliesRequestApproval.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   75
      TabIndex        =   8
      Top             =   570
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
      Picture         =   "frmSuppliesRequestApproval.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   75
      TabIndex        =   9
      Top             =   1200
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Approve"
      AccessKey       =   "A"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmSuppliesRequestApproval.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   75
      TabIndex        =   10
      Top             =   1830
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&DisApprv"
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
      Picture         =   "frmSuppliesRequestApproval.frx":166E
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   6750
      Index           =   1
      Left            =   1560
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   9660
      _ExtentX        =   17039
      _ExtentY        =   11906
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   2
         Left            =   1200
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1470
         Width           =   3930
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   1200
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   375
         Width           =   1935
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   1
         Left            =   1200
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   1005
         Width           =   2010
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   1215
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   3255
         Width           =   2550
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   1215
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   2850
         Width           =   2550
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   1215
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   2430
         Width           =   3705
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   7665
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   2865
         Width           =   1815
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   4845
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   2865
         Width           =   1785
      End
      Begin xrGridEditor.GridEditor GridEditor1 
         Height          =   2820
         Left            =   90
         TabIndex        =   12
         Top             =   3795
         Width           =   9450
         _ExtentX        =   16669
         _ExtentY        =   4974
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
         Object.HEIGHT          =   2820
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
         MOUSEICON       =   "frmSuppliesRequestApproval.frx":1DE8
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
      Begin VB.Shape Shape2 
         Height          =   1545
         Index           =   1
         Left            =   105
         Top             =   2220
         Width           =   9435
      End
      Begin VB.Shape Shape2 
         Height          =   2085
         Index           =   0
         Left            =   105
         Top             =   105
         Width           =   9420
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   270
         Index           =   5
         Left            =   225
         TabIndex        =   20
         Top             =   1470
         Width           =   900
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Trans. No."
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
         Index           =   3
         Left            =   240
         TabIndex        =   19
         Top             =   405
         Width           =   1110
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date "
         Height          =   300
         Index           =   0
         Left            =   240
         TabIndex        =   18
         Top             =   1065
         Width           =   915
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
         Height          =   195
         Index           =   2
         Left            =   270
         TabIndex        =   17
         Top             =   3345
         Width           =   585
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Entry No"
         Height          =   195
         Index           =   3
         Left            =   270
         TabIndex        =   16
         Top             =   2940
         Width           =   615
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   15
         Top             =   2460
         Width           =   795
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   405
         Left            =   1335
         Tag             =   "et0;ht2"
         Top             =   480
         Width           =   1995
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qty on Hnd"
         Height          =   195
         Index           =   6
         Left            =   3945
         TabIndex        =   14
         Top             =   2940
         Width           =   810
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Act on Hnd"
         Height          =   195
         Index           =   1
         Left            =   6765
         TabIndex        =   13
         Top             =   2925
         Width           =   810
      End
   End
End
Attribute VB_Name = "frmSuppliesRequestApproval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
'Private Const pxeMODULENAME = "frmShiftSchedApproval"
'
'Private WithEvents oTrans As ggcSuppliesRequest
'Private oSkin As clsFormSkin
'Private bLoaded As Boolean
'
'Dim pnIndex As Integer
'Dim pnCtr As Integer
'Dim pbSearched As Boolean
'Private psTransNox As String
'Public Property Let TransNox(Value As String)
'   psTransNox = Value
'End Property
'
'Private Sub cmdButton_Click(Index As Integer)
'   Dim lsOldProc As String
'   Dim lnRep As Integer
'
'   lsOldProc = "cmdButton_Click"
'   'On Error GoTo errProc
'
'   Select Case Index
'   Case 0   'close
'      Unload Me
'   Case 1   'search
'      If pnIndex = 0 Or pnIndex = 1 Then
'         If pnIndex = 0 Then
'            If oTrans.SearchTransaction(txtSearch(pnIndex).Text, False) Then
'               ClearFields
'               LoadMaster
'               Call InitFields
'            End If
'         Else
'            If oTrans.SearchTransaction(txtSearch(pnIndex).Text) Then
'               ClearFields
'               LoadMaster
'               Call InitFields
'            End If
'         End If
'         pnIndex = 3
'      Else
'         If oTrans.SearchTransaction("") Then
'            ClearFields
'            LoadMaster
'            Call InitFields
'         End If
'      End If
'   Case 2   'approve
'      If oTrans.Master(0) <> "" Then
'         If oTrans.CloseTransaction(oTrans.Master(0)) Then
'            MsgBox "Transaction was closed successfuly!!!", vbInformation, "Notice"
'         Else
'            MsgBox "Closing/Posting transaction failed!!!", vbInformation, "Notice"
'         End If
'         Call ClearFields
'      End If
'      GoTo endWithFocus
'   Case 3   'disapprove
'      If oTrans.Master(0) <> "" And txtField(0) <> "" Then
'         If oTrans.CancelTransaction Then
'            MsgBox "Transaction was cancelled!!!", vbInformation, "Notice"
'         Else
'            MsgBox "Transaction cancellation failed!!!", vbInformation, "Notice"
'         End If
'         ClearFields
'      End If
'      GoTo endWithFocus
'   End Select
'
'endProc:
'   Exit Sub
'endWithFocus:
'   txtSearch(0) = ""
'   txtSearch(1) = ""
'   txtSearch(0).SetFocus
'   GoTo endProc
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
'   If bLoaded = False Then
'      bLoaded = True
'   End If
'
'   pbSearched = False
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " & " )", True
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
'Private Sub Form_Load()
'   Dim lsOldProc As String
'
'   lsOldProc = "Form_Load"
''   'On Error GoTo errProc
'
'   CenterChildForm mdiMain, Me
'
'   Set oTrans = New ggcSuppliesRequest
'   Set oTrans.AppDriver = oApp
'
'   If LCase(oApp.ProductID) = "petmgr" Then
'      oTrans.TransStatus = 10
'   Else
'      oTrans.TransStatus = 0
'   End If
'
'   oTrans.Branch = oApp.BranchCode
'   oTrans.InitTransaction
'   If psTransNox <> "" Then
'      '@@@ soft-monitor
'      Call oTrans.OpenTransaction(psTransNox)
'   End If
'
'   Set oSkin = New clsFormSkin
'   Set oSkin.AppDriver = oApp
'   Set oSkin.Form = Me
'   oSkin.ApplySkin xeFormTransaction
'
'   ClearFields
'   InitGrid
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
'Private Sub LoadMaster()
'   Dim loTxt As TextBox
'
'   For Each loTxt In txtField
'      Select Case loTxt.Index
'         Case 1, 3
'            loTxt.Text = strLongDate(oTrans.Master(loTxt.Index))
'         Case Else
'            loTxt.Text = oTrans.Master(loTxt.Index)
'      End Select
'   Next
'
'   txtSearch(0) = txtField(0)
'   txtSearch(1) = txtField(2)
'
'   If oTrans.Master("cTranStat") = "4" Then
'      Label2.Caption = "APPLIED"
'   Else
'      Label2.Caption = TransStat(CInt(oTrans.Master("cTranStat")))
'   End If
'
'   pbSearched = True
'End Sub
'
'
'
'Private Sub oTrans_MasterRetrieved(ByVal Index As Variant, ByVal Value As Variant)
'   Select Case Index
'      Case 1, 3
'      txtField(Index) = strLongDate(oTrans.Master(Index))
'      Case 9
'         Label2.Caption = TransStat(CInt(Value))
'   End Select
'End Sub
'
'Private Sub txtField_GotFocus(Index As Integer)
'
'   Select Case Index
'      Case 1, 3
'         txtField(Index) = strShortDate(oTrans.Master(Index))
'   End Select
'
'   With txtField(Index)
'      .BackColor = oApp.getColor("HT1")
'      .SelStart = 0
'      .SelLength = Len(.Text)
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
'Private Sub txtSearch_LostFocus(Index As Integer)
'   With txtSearch(Index)
'      .BackColor = oApp.getColor("EB")
'   End With
'
'   pnIndex = Index
'End Sub
'Private Sub txtSearch_GotFocus(Index As Integer)
'   With txtSearch(Index)
'      .BackColor = oApp.getColor("HT1")
'      .SelStart = 0
'      .SelLength = Len(.Text)
'   End With
'
'   pnIndex = Index
'End Sub
'Private Sub txtSearch_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
'      Select Case Index
'      Case 0
'         If oTrans.SearchTransaction(txtSearch(Index).Text, False) Then
'            ClearFields
'            LoadMaster
'            Call InitFields
'         End If
'      Case 1
'         If oTrans.SearchTransaction(txtSearch(Index).Text) Then
'            ClearFields
'            LoadMaster
'            Call InitFields
'         End If
'      End Select
'   End If
'End Sub
'
'
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
'Private Sub ClearFields()
'   For pnCtr = 0 To txtField.Count - 1
'      Select Case pnCtr
'      Case 0
'         txtField(pnCtr).Text = Format(oTrans.Master(pnCtr), "@@@@-@@@@@@")
'      Case 2
'         txtField(pnCtr).Text = oTrans.Master(pnCtr)
'      Case 1
'         txtField(pnCtr).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
'      Case Else
'        txtField(pnCtr).Text = Empty
'      End Select
'   Next
'
'   With GridEditor1
'      .Rows = 2
'      .ColWidth(3) = 3100
'
'      'empty row
'      .TextMatrix(1, 1) = ""
'      .TextMatrix(1, 2) = ""
'      .TextMatrix(1, 3) = ""
'      .TextMatrix(1, 4) = "0.00"
'      .TextMatrix(1, 5) = "0"
'      .TextMatrix(1, 6) = "0"
'   End With
'
''   chkField.Value = 0
''   pbSave = False
''   pbClosedTrans = False
'End Sub
'
'Private Sub InitGrid()
'   With GridEditor1
'      .Rows = 2
'      .Cols = 6
'      .Font = "MS Sans Serif"
'
'      'column title
'      .TextMatrix(0, 1) = "Entry No"
'      .TextMatrix(0, 2) = "Description"
'      .TextMatrix(0, 3) = "Quantity"
'      .TextMatrix(0, 4) = "Qty on Hnd"
'      .TextMatrix(0, 5) = "Act on Hnd"
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
'      .ColWidth(1) = 2600
'      .ColWidth(2) = 2500
'      .ColWidth(3) = 2500
'      .ColWidth(4) = 1020
'      .ColWidth(5) = 1020
'
'      .ColFormat(4) = "#,##0.00"
'      .ColFormat(5) = "#,##0"
'      .ColFormat(6) = "#,##0"
'      .ColDefault(4) = 0#
'      .ColDefault(5) = 0
'
'
'      .ColAlignment(1) = 1
'      .ColAlignment(2) = 1
'      .ColAlignment(3) = 1
'      .ColAlignment(4) = 6
'      .ColAlignment(5) = 6
'
'
'      .EditorBackColor = oApp.getColor("HT1")
'
'      .Row = 1
'      .Col = 1
'   End With
'End Sub
''
''
''
''Private Sub InitGrid()
''   With GridEditor1
''      .Rows = 2
''      .Cols = 6
''      .Font = "MS Sans Serif"
''
''
''      'Cols Title
''      .TextMatrix(0, 1) = "Entry No"
''      .TextMatrix(0, 2) = "Description"
''      .TextMatrix(0, 3) = "Quantity"
''      .TextMatrix(0, 4) = "Qty on Hnd"
''      .TextMatrix(0, 5) = "Qty on Hnd"
''      .Row = 0
''
''      For pnCtr = 0 To .Cols - 1
''               .Cols = pnCtr
''               .CellFontBold = True
''               .CellAlignment = 3
''      Next
''
''      'Cols Width
''      .ColWidth(0) = 330
''      .ColWidth(1) = 1330
''      .ColWidth(2) = 1330
''      .ColWidth(3) = 1330
''      .ColWidth(4) = 1000
''      .ColWidth(5) = 1000
''      'Cols Format
''      .ColFormat(3) = "#,##0"
''      .ColFormat(4) = "#,##0"
''      .ColFormat(5) = "#,##0"
''      .ColDefault(4) = 0
''      .ColDefault(6) = 0
''
''      .ColAlignment(1) = 1
''      .ColAlignment(2) = 1
''      .ColAlignment(3) = 1
''      .ColAlignment(4) = 6
''      .ColAlignment(5) = 6
''      .EditorBackColor = oApp.getColor("HT1")
''
''      .Row = 1
''      .Col = 1
''   End With
''
''
''End Sub
