VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0"; "xrGridControl.ocx"
Begin VB.Form frmSuppliesAdj 
   BorderStyle     =   0  'None
   Caption         =   "Supplies Adjustment"
   ClientHeight    =   7680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11655
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7680
   ScaleWidth      =   11655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin xrControl.xrFrame xrFrame1 
      Height          =   7125
      Index           =   1
      Left            =   1575
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   10050
      _ExtentX        =   17727
      _ExtentY        =   12568
      BackColor       =   -2147483626
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   10
         Left            =   7965
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   3390
         Width           =   1830
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   9
         Left            =   7965
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   3000
         Width           =   1830
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   8
         Left            =   4275
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   3390
         Width           =   2010
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   7
         Left            =   4275
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   3000
         Width           =   2010
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   5
         Left            =   1140
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   3000
         Width           =   1950
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   6
         Left            =   1140
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   3390
         Width           =   1950
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   4
         Left            =   1140
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   2580
         Width           =   3705
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   510
         Index           =   3
         Left            =   1155
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1650
         Width           =   8640
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
         Left            =   1155
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   315
         Width           =   1905
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   1155
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1260
         Width           =   2805
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   1155
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   870
         Width           =   1905
      End
      Begin xrGridEditor.GridEditor GridEditor1 
         Height          =   3135
         Left            =   75
         TabIndex        =   15
         Top             =   3900
         Width           =   9885
         _ExtentX        =   17436
         _ExtentY        =   5530
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
         Object.HEIGHT          =   3135
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
         MOUSEICON       =   "frmSupplier_Adjustment.frx":0000
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
         Height          =   2295
         Index           =   0
         Left            =   75
         Top             =   90
         Width           =   9870
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Balance"
         Height          =   195
         Index           =   7
         Left            =   7020
         TabIndex        =   22
         Top             =   3450
         Width           =   585
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Price"
         Height          =   195
         Index           =   6
         Left            =   7020
         TabIndex        =   21
         Top             =   3060
         Width           =   690
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qty Out"
         Height          =   195
         Index           =   4
         Left            =   3570
         TabIndex        =   20
         Top             =   3450
         Width           =   555
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qty In"
         Height          =   195
         Index           =   2
         Left            =   3570
         TabIndex        =   19
         Top             =   3060
         Width           =   435
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Entry No"
         Height          =   195
         Index           =   3
         Left            =   195
         TabIndex        =   18
         Top             =   3060
         Width           =   615
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stock ID"
         Height          =   195
         Index           =   1
         Left            =   195
         TabIndex        =   17
         Top             =   3450
         Width           =   630
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   16
         Top             =   2640
         Width           =   795
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   195
         Index           =   5
         Left            =   195
         TabIndex        =   14
         Top             =   1755
         Width           =   630
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         Height          =   195
         Index           =   9
         Left            =   195
         TabIndex        =   13
         Top             =   1320
         Width           =   360
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
         Left            =   195
         TabIndex        =   12
         Top             =   345
         Width           =   1110
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   315
         Left            =   1245
         Tag             =   "et0;ht2"
         Top             =   420
         Width           =   1920
      End
      Begin VB.Shape Shape2 
         Height          =   1380
         Index           =   1
         Left            =   75
         Top             =   2430
         Width           =   9870
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date "
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   11
         Top             =   930
         Width           =   810
      End
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   600
      Index           =   3
      Left            =   105
      TabIndex        =   23
      Top             =   2445
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
      Picture         =   "frmSupplier_Adjustment.frx":001C
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   105
      TabIndex        =   24
      Top             =   555
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
      Picture         =   "frmSupplier_Adjustment.frx":0796
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   105
      TabIndex        =   25
      Top             =   2445
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
      Picture         =   "frmSupplier_Adjustment.frx":0F10
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   600
      Index           =   1
      Left            =   105
      TabIndex        =   26
      Top             =   1815
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
      Picture         =   "frmSupplier_Adjustment.frx":168A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   105
      TabIndex        =   27
      Top             =   1815
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
      Picture         =   "frmSupplier_Adjustment.frx":1E04
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   105
      TabIndex        =   28
      Top             =   1185
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Memo"
      AccessKey       =   "M"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmSupplier_Adjustment.frx":257E
   End
End
Attribute VB_Name = "frmSuppliesAdj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Private Const pxeMODULENAME = "frmSuppliesAdj"
'
'Private WithEvents oTrans As clsSuppliesAdj
'Private oSkin As clsFormSkin
'Private pnIndex As Integer
'Dim pnCtr As Integer
'
'Private Sub cmdButton_Click(Index As Integer)
'   Dim lsOldProc As String
'   Dim lnRep As Integer
'   Dim loObj As Object
'
'   lsOldProc = "cmdButton_Click"
'   'On Error GoTo errProc
'
'   Select Case Index
'   Case 0
'      If oTrans.SaveTransaction = True Then
'         MsgBox "Record Updated Successfully!!!", vbInformation, "Notice"
'         Call cmdButton_Click(4) 'new
'      Else
'         MsgBox "Unable to Update Record!!!", vbCritical, "Warning"
'         txtField(pnIndex).SetFocus
'      End If
'   Case 1
'      oTrans.SearchMaster pnIndex
'      txtField(pnIndex).SetFocus
'   Case 2
'      oTrans.GetMemo loObj
'   Case 3
'      lnRep = MsgBox("Transaction is in Update Mode!!!" & vbCrLf & _
'                     "Do you want to Cancel Transaction!!!", vbYesNo + vbQuestion, "Confirm")
'
'      If lnRep = vbYes Then
'         If oTrans.UnSaveTransaction = False Then
'            MsgBox "Unable to cancel update of record!", vbCritical, "Warning"
'         Else
'            InitButton xeModeReady
'         End If
'      Else
'         txtField(pnIndex).SetFocus
'      End If
'   Case 4
'      If oTrans.NewTransaction = False Then
'         MsgBox "Unable to create new record!", vbCritical, "Warning"
'      Else
'         InitButton xeModeAddNew
'         txtField(1).SetFocus
'      End If
'   Case 5
'      Unload Me
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
'   Set oTrans = New clsSuppliesAdj
'   Set oTrans.AppDriver = oApp
'
'   oTrans.Branch = oApp.BranchCode
'   oTrans.InitTransaction
''   oTrans.NewTransaction
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
'
'Private Sub Form_Unload(Cancel As Integer)
'   Set oTrans = Nothing
'   Set oSkin = Nothing
'End Sub
'
'Private Sub oTrans_Master()
'   Dim pnCtr As Integer
'   For pnCtr = 0 To 11
'      Select Case pnCtr
'      Case 0
'         txtField(pnCtr).Text = Format(oTrans.Master(pnCtr), "@@@@-@@@@@@")
'      Case 1
'         txtField(pnCtr).Text = Format(oTrans.Master(pnCtr), "MMMM DD, YYYY")
'      Case 2
'         txtField(pnCtr).Text = oTrans.Master(pnCtr)
'      Case 3, 12
'      Case 4, 5
'         txtField(pnCtr).Text = Format(oTrans.Master(pnCtr), "#,##0")
'      Case Else
''         txtField(pnCtr).Text = IIf(IsNull(oTrans.Master(pnCtr)), "", oTrans.Master(pnCtr))
'      End Select
'   Next
'
'   xrFrame1(1).Enabled = False
'End Sub
'
''Private Sub oTrans_MasterRetrieved(ByVal Index As Variant)
''   Select Case Index
''   Case 7 To 11, 13
''      txtField(Index) = IFNull(oTrans.Master(Index))
''   End Select
''End Sub
'
'
'Private Sub txtField_GotFocus(Index As Integer)
'   With txtField(Index)
'      If Index = 1 Then .Text = Format(.Text, "MM/DD/YYYY")
'      .SelStart = 0
'      .SelLength = Len(.Text)
'      .BackColor = oApp.getColor("HT1")
'   End With
'   pnIndex = Index
'End Sub
'
'Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'   Dim lsOldProc As String
'
'   lsOldProc = "txtField_KeyDown"
'   'On Error GoTo errProc
'
'   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
'      With txtField(Index)
'         Select Case Index
'         Case 7, 8
'            If KeyCode = vbKeyF3 Then
'               oTrans.SearchMaster Index, txtField(Index).Text
'            End If
'         End Select
'      End With
'      KeyCode = 0
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
'         SetNextFocus
'      Case vbKeyUp
'         SetPreviousFocus
'      End Select
'   End Select
'End Sub
'
'Private Sub InitButton(lnStat As Integer)
'   Dim lbShow As Boolean
'
'   lbShow = IIf(lnStat = 0, False, True)
'   cmdButton(0).Visible = lbShow
'   cmdButton(1).Visible = lbShow
'   cmdButton(3).Visible = lbShow
'
'   cmdButton(4).Visible = Not lbShow
'   cmdButton(5).Visible = Not lbShow
'
'   xrFrame1(1).Enabled = lbShow
'
'   If Not lbShow Then cmdButton(4).SetFocus
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
'   'On Error GoTo errProc
'
'   With txtField(Index)
'      .Text = TitleCase(.Text)
'
'      Select Case Index
'      Case 1
'         If Not IsDate(.Text) Then .Text = oApp.ServerDate
'         .Text = Format(.Text, "MMMM DD, YYYY")
'      Case 4, 5
'         If Not IsNumeric(.Text) Then txtField(Index).Text = 0
'         If oTrans.Master("cHsSerial") = xeYes Then
'            If Index = 4 Then
'               If oTrans.Master("cLocation") = xeLocBranch Then .Text = 0
'            Else
'               If oTrans.Master("cLocation") <> xeLocBranch Then .Text = 0
'            End If
'            If CDbl(.Text) > 1 Then .Text = 0
'         End If
'         .Text = Format(.Text, "#,##0")
'      End Select
'      oTrans.Master(Index) = .Text
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
'Private Sub InitGrid()
'   With GridEditor1
'      .Rows = 2
'      .Cols = 7
'      .Font = "MS Sans Serif"
'
'      'column title
'      .TextMatrix(0, 1) = "Entry No"
'      .TextMatrix(0, 2) = "Description"
'      .TextMatrix(0, 3) = "Qty In"
'      .TextMatrix(0, 4) = "Qty Out"
'      .TextMatrix(0, 5) = "Unit Price"
'      .TextMatrix(0, 6) = "Balance"
'
'
'      .Row = 0
'
'      'column alignment
'      For pnCtr = 0 To .Cols - 1
'            .Col = pnCtr
'            .CellFontBold = True
'            .CellAlignment = 3
'      Next
'
'      'column width
'      .ColWidth(0) = 330
'      .ColWidth(1) = 1000
'      .ColWidth(2) = 2600
'      .ColWidth(3) = 1000
'      .ColWidth(4) = 1000
'      .ColWidth(5) = 1000
'      .ColWidth(6) = 1000
'
'      .ColFormat(4) = "#,##0.00"
'      .ColNumberOnly(6) = True
'      .ColDefault(4) = 0#
'      .ColDefault(5) = 0
'      .ColDefault(6) = 0
'
'      .ColAlignment(1) = 1
'      .ColAlignment(2) = 1
'      .ColAlignment(3) = 1
'      .ColAlignment(4) = 6
'      .ColAlignment(5) = 6
'      .ColAlignment(6) = 6
'
'      .ColEnabled(3) = False
'      .ColEnabled(5) = False
'
'      .EditorBackColor = oApp.getColor("HT1")
'
'      .Row = 1
'      .Col = 1
'   End With
'End Sub
'Private Sub ClearFields()
'   For pnCtr = 0 To txtField.Count - 1
'      Select Case pnCtr
'      Case 0
'         txtField(pnCtr).Text = Format(oTrans.Master(0), "@@@@-@@@@@@@@")
'      Case 1
'         txtField(pnCtr).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
'
'      Case Else
'         txtField(pnCtr).Text = Empty
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
''   chkField(1).value = 0
'
''   oTrans.BackLoad = chkField(1).value
'End Sub
