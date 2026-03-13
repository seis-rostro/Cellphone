VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0"; "xrGridControl.ocx"
Begin VB.Form frmSuppliesRequestReg 
   BorderStyle     =   0  'None
   Caption         =   "Supplies Request Register"
   ClientHeight    =   7305
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11265
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7305
   ScaleWidth      =   11265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   9930
      TabIndex        =   0
      Top             =   1740
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
      Picture         =   "frmSuppliesRequestReg.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   9915
      TabIndex        =   5
      Top             =   480
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
      Picture         =   "frmSuppliesRequestReg.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   9930
      TabIndex        =   6
      Top             =   1110
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
      Picture         =   "frmSuppliesRequestReg.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   9915
      TabIndex        =   8
      Top             =   1110
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
      Picture         =   "frmSuppliesRequestReg.frx":166E
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   600
      Index           =   4
      Left            =   9915
      TabIndex        =   9
      Top             =   1740
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
      Picture         =   "frmSuppliesRequestReg.frx":1DE8
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   6750
      Index           =   1
      Left            =   75
      Tag             =   "wt0;fb0"
      Top             =   480
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
         TabIndex        =   12
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
         Enabled         =   0   'False
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
         TabIndex        =   11
         Text            =   "Text 1"
         Top             =   3255
         Width           =   2550
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Enabled         =   0   'False
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
         Enabled         =   0   'False
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
         Text            =   "Text 1"
         Top             =   2865
         Width           =   1815
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         TabIndex        =   10
         Text            =   "Text 1"
         Top             =   2865
         Width           =   1785
      End
      Begin xrGridEditor.GridEditor GridEditor1 
         Height          =   2820
         Left            =   90
         TabIndex        =   13
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
         MOUSEICON       =   "frmSuppliesRequestReg.frx":2562
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
         TabIndex        =   21
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
         TabIndex        =   20
         Top             =   405
         Width           =   1110
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date "
         Height          =   300
         Index           =   0
         Left            =   240
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
         Top             =   2925
         Width           =   810
      End
   End
End
Attribute VB_Name = "frmSuppliesRequestReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
'Private Const pxeMODULENAME = "frmSuppliesRequestReg"
'
'Private WithEvents oTrans As ggcSuppliesRequest
'Private oSkin As clsFormSkin
'Private bLoaded As Boolean
'
'Dim psSelected() As String
'Dim pnCtr As Integer
'Dim pnIndex As Integer
'
'Private Sub cmdButton_Click(Index As Integer)
'10       Dim lsOldProc As String
'20       Dim lnRep As Integer
'
'30       lsOldProc = "cmdButton_Click"
'40       'On Error GoTo errProc
'
'50       Select Case Index
'   Case 0   'save
'60          If oTrans.SaveTransaction(True) Then
'70             MsgBox "Transaction was successfuly updated.", vbInformation, "Notice"
'80             ClearFields
'90             InitForm 0
'100         End If
'110      Case 4   'cancel trans
'120         lnRep = MsgBox("Transaction is in Update Mode!!!" & vbCrLf & _
'                  "Do you want to Cancel Transaction!!!", vbYesNo + vbQuestion, "Confirm")
'
'130         If lnRep = vbYes Then
'140            InitForm 0
'150            ClearFields
'160         End If
'170      Case 2   'update
''      If txtField(0).Text <> "" Then
''         If oTrans.UpdateTransaction Then InitForm 1
''      End If
'
'180      Case 3   'browse
'190          If pnIndex = 0 Or pnIndex = 1 Then
'200            If pnIndex = 0 Then
'210               If oTrans.SearchTransaction(txtSearch(pnIndex).Text, False) Then
'220                  ClearFields
'230                  LoadMaster
'240               End If
'250            Else
'260               If oTrans.SearchTransaction(txtSearch(pnIndex).Text) Then
'270                  ClearFields
'280                  LoadMaster
'290               End If
'300            End If
'310            pnIndex = 3
'320         Else
'330            If oTrans.SearchTransaction("") Then
'340               ClearFields
'350               LoadMaster
'360            End If
'370         End If
'380      Case 5   'close
'390         Unload Me
'400      End Select
'
'endProc:
'410      Exit Sub
'errProc:
'420      ShowError lsOldProc & "( " & Index & " )", True
'End Sub
'
'Private Sub Form_Activate()
'10       Dim lsOldProc As String
'
'20       lsOldProc = "Form_Activate"
'30       'On Error GoTo errProc
'
'40       oApp.MenuName = Me.Tag
'50       Me.ZOrder 0
'
'60       If bLoaded = False Then
'70          bLoaded = True
'80       End If
'
'endProc:
'90       Exit Sub
'errProc:
'100      ShowError lsOldProc & "( " & " )", True
'End Sub
'
'Private Sub Form_Load()
'10       Dim lsOldProc As String
'
'20       lsOldProc = "Form_Load"
'30       'On Error GoTo errProc
'
'40       CenterChildForm mdiMain, Me
'
'50       Set oTrans = New ggcSuppliesRequest
'60       Set oTrans.AppDriver = oApp
'
'70       oTrans.Branch = oApp.BranchCode
'80       oTrans.InitTransaction
'
'90       Set oSkin = New clsFormSkin
'100      Set oSkin.AppDriver = oApp
'110      Set oSkin.Form = Me
'120      oSkin.ApplySkin xeFormTransMaintenance
'
'130      InitGrid
'140      ClearFields
'endProc:
'150      Exit Sub
'errProc:
'160      ShowError lsOldProc & "( " & " )", True
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'10       Set oTrans = Nothing
'20       Set oSkin = Nothing
'End Sub
'
'
'
'Private Sub InitForm(lnStat As Integer)
'10       Dim lbShow As Boolean
'
'20       lbShow = IIf(lnStat = 0, False, True)
'30       cmdButton(3).Visible = Not lbShow
'40       cmdButton(2).Visible = Not lbShow
'50       cmdButton(5).Visible = Not lbShow
'60       txtField(1).Enabled = Not lbShow
'
'
'80       cmdButton(0).Visible = lbShow
'90       cmdButton(4).Visible = lbShow
'
'100      If lbShow Then
'110         txtField(2).SetFocus
'120      End If
'End Sub
'
'Private Sub LoadMaster()
'10       Dim loTxt As TextBox
'
'20       For Each loTxt In txtField
'30          Select Case loTxt.Index
'         Case 1, 3
'40                loTxt.Text = strLongDate(oTrans.Master(loTxt.Index))
'50             Case Else
'60                loTxt.Text = oTrans.Master(loTxt.Index)
'70          End Select
'80       Next
'
'90       txtSearch(0) = txtField(0)
'100      txtSearch(1) = txtField(2)
'
'110   If oTrans.Master("cTranStat") = "4" Then
'111      Label2.Caption = "APPLIED"
'112   Else
'113      Label2.Caption = TransStat(CInt(oTrans.Master("cTranStat")))
'114   End If
'
'End Sub
'
'Private Sub oTrans_MasterRetrieved(ByVal Index As Variant, ByVal Value As Variant)
'10       Select Case Index
'      Case 1, 3
'20             txtField(Index).Text = strLongDate(oTrans.Master(Index))
'30          Case 9
'40             Label2.Caption = TransStat(CInt(Value))
'50          Case Else
'60             With txtField(Index)
'70                .Text = Value
'80             End With
'90       End Select
'End Sub
'
'Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
'10       With txtField(Index)
'20          oTrans.Master(Index) = .Text
'30       End With
'End Sub
'
'Private Sub txtField_GotFocus(Index As Integer)
'10       Select Case Index
'      Case 1, 3
'20             txtField(Index).Text = Format(oTrans.Master(Index), "MM-DD-YYYY")
'30       End Select
'
'40       With txtField(Index)
'50          .BackColor = oApp.getColor("HT1")
'60          .SelStart = 0
'70          .SelLength = Len(.Text)
'80       End With
'
'90       pnIndex = Index
'End Sub
'
'Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'10       Select Case KeyCode
'   Case vbKeyF3
'20          If Index = 2 Or Index = 5 Then
'30             If oTrans.SearchMaster(Index, txtField(Index).Text) Then
'40                SetNextFocus
'50             End If
'60          End If
'70       Case vbKeyReturn
'80          If Index = 2 Or Index = 5 Then
'90             If txtField(Index) <> "" Then
'100               Call oTrans.SearchMaster(Index, txtField(Index).Text)
'110            Else
'120               oTrans.Master(Index) = txtField(Index).Text
'130            End If
'140         End If
'150      End Select
'End Sub
'
'Private Sub txtField_LostFocus(Index As Integer)
'10       With txtField(Index)
'20          .BackColor = oApp.getColor("EB")
'30       End With
'End Sub
'
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'10       Select Case KeyCode
'   Case vbKeyReturn, vbKeyUp, vbKeyDown
'20          Select Case KeyCode
'      Case vbKeyReturn, vbKeyDown
'30             SetNextFocus
'40          Case vbKeyUp
'50             SetPreviousFocus
'60          End Select
'70       End Select
'End Sub
'
'Private Sub txtSearch_LostFocus(Index As Integer)
'10       With txtSearch(Index)
'20          .BackColor = oApp.getColor("EB")
'30       End With
'
'40       pnIndex = Index
'End Sub
'
'Private Sub txtSearch_GotFocus(Index As Integer)
'10       With txtSearch(Index)
'20          .BackColor = oApp.getColor("HT1")
'30          .SelStart = 0
'40          .SelLength = Len(.Text)
'50       End With
'
'60       pnIndex = Index
'End Sub
'
'Private Sub txtSearch_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'10       If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
'20          Select Case Index
'      Case 0
'30             If oTrans.SearchTransaction(txtSearch(Index).Text, False) Then
'40                ClearFields
'50                LoadMaster
'60             End If
'70          Case 1
'80             If oTrans.SearchTransaction(txtSearch(Index).Text) Then
'90                ClearFields
'100               LoadMaster
'110            End If
'120         End Select
'130      End If
'End Sub
'
'Private Sub ShowError(ByVal lsProcName As String, Optional bEnd As Boolean = False)
'10       With oApp
'20          .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
'30          If bEnd Then
'40             .xShowError
'50             End
'60          Else
'70             With Err
'80                .Raise .Number, .Source, .Description
'90             End With
'100         End If
'110      End With
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
'Private Sub InitGrid()
'   With GridEditor1
'      .Rows = 2
'      .Cols = 6
'      .Font = "MS Sans Serif"
'
'      'Column Title
'      .TextMatrix(0, 1) = "Entry No"
'      .TextMatrix(0, 2) = "Description"
'      .TextMatrix(0, 3) = "Quantity"
'      .TextMatrix(0, 4) = "Qty on Hnd"
'      .TextMatrix(0, 5) = "Act on Hnd"
'      .Row = 0
'
'      'Column Width & Alignment
'      For pnCtr = 0 To .Cols - 1
'            .Col = pnCtr
'            .CellFontBold = True
'            .CellAlignment = 3
'
'      Next
'      .ColWidth(0) = "330"
'      .ColWidth(1) = "1330"
'      .ColWidth(2) = "1500"
'      .ColWidth(3) = "1000"
'      .ColWidth(4) = "1000"
'      .ColWidth(5) = "1000"
'      'End for Width
'
'      'Column Format & Alignment
'      .ColFormat(4) = "#,##0.00"
'      .ColFormat(5) = "#,##0"
'      .ColFormat(6) = "#,##0"
'      .ColDefault(4) = 0
'      .ColDefault(5) = 0
'
'      .ColAlignment(1) = 1
'      .ColAlignment(2) = 1
'      .ColAlignment(3) = 1
'      .ColAlignment(4) = 6
'      .ColAlignment(5) = 6
'      .EditorBackColor = oApp.getColor("HT1")
'      .Row = 1
'      .Col = 1
'      'End for format
'   End With
'End Sub
'
'
