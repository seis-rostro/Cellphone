VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0"; "xrGridControl.ocx"
Begin VB.Form frmSalesPackage 
   BorderStyle     =   0  'None
   Caption         =   "Sales Package Maintenance"
   ClientHeight    =   5685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6795
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5685
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   Begin xrGridEditor.GridEditor GridEditor1 
      Height          =   3405
      Left            =   165
      TabIndex        =   2
      Tag             =   "et0;eb0;et0;bc2"
      Top             =   1185
      Width           =   6480
      _ExtentX        =   11430
      _ExtentY        =   6006
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
      Object.HEIGHT          =   3405
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
      MOUSEICON       =   "frmSalesPackage.frx":0000
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
      Height          =   4110
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   6600
      _ExtentX        =   11642
      _ExtentY        =   7250
      BorderStyle     =   1
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
         Index           =   0
         Left            =   1050
         TabIndex        =   1
         Top             =   135
         Width           =   5355
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   315
         Left            =   1140
         Tag             =   "et0;ht2"
         Top             =   240
         Width           =   5355
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "CP Model"
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
         Index           =   2
         Left            =   105
         TabIndex        =   0
         Top             =   165
         Width           =   855
      End
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   705
      Index           =   0
      Left            =   4380
      TabIndex        =   3
      Top             =   4905
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
      Picture         =   "frmSalesPackage.frx":001C
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   705
      Index           =   2
      Left            =   5940
      TabIndex        =   5
      Top             =   4905
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
      Picture         =   "frmSalesPackage.frx":0796
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   705
      Index           =   1
      Left            =   5160
      TabIndex        =   4
      Top             =   4905
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
      Picture         =   "frmSalesPackage.frx":0F10
      PicturePos      =   1
   End
End
Attribute VB_Name = "frmSalesPackage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Private Const pxeMODULENAME = "frmSalesPackage"
'
'Private WithEvents oTrans As clsSetPackage
'Private oSkin As clsFormSkin
'
'Dim pbGridGotFocus As Boolean
'
'Private Sub cmdButton_Click(Index As Integer)
'   Dim lsOldProc As String
'   Dim lnCtr As Integer
'
'   lsOldProc = "cmdButton_Click"
'   'On Error GoTo errProc
'
'   Select Case Index
'   Case 0
'      If txtField(0).Text = "" Then
'         MsgBox "Invalid MC Model!!!" & vbCrLf & _
'                "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
'      Else
'         With GridEditor1
'            If .Rows > 2 Then
'               lnCtr = 2
'               Do While lnCtr < .Rows
'                  If CDbl(.TextMatrix(lnCtr, 3)) = 0 Then
'                     .Row = lnCtr
'                     If oTrans.DeleteDetail(.Row - 1) Then .DeleteRow
'                  Else
'                     lnCtr = lnCtr + 1
'                  End If
'               Loop
'            End If
'         End With
'
'         If oTrans.SaveTransaction Then
'            MsgBox "Record Save Successfully!!!", vbInformation, "Confirm"
'            InitValue
'         Else
'            MsgBox "Unable to Save Record!!!", vbCritical, "Warning"
'         End If
'      End If
'   Case 1
'      If Not pbGridGotFocus Then
'         If oTrans.SearchTransaction Then LoadDetail
'      End If
'   Case 2
'      Unload Me
'   Case 3
'      If pbGridGotFocus Then
'         With GridEditor1
'            If .Rows > 2 Then
'               If oTrans.DeleteDetail(.Row - 1) Then .DeleteRow
'            End If
'         End With
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
'   With GridEditor1
'      .Refresh
'   End With
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
'   Set oTrans = New clsSetPackage
'   Set oTrans.AppDriver = oApp
'   oTrans.InitTransaction
'
'   Set oSkin = New clsFormSkin
'   Set oSkin.AppDriver = oApp
'   Set oSkin.Form = Me
'   oSkin.ApplySkin
'
'   InitGrid
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " & " )", True
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'   Set oSkin = Nothing
'   Set oTrans = Nothing
'End Sub
'
'Private Sub InitGrid()
'   Dim lnCtr As Integer
'
'   With GridEditor1
'      .Rows = 2
'      .Cols = 4
'      .Font = "MS Sans Serif"
'
'      'Column Title
'      .TextMatrix(0, 1) = "BarrCode"
'      .TextMatrix(0, 2) = "Description"
'      .TextMatrix(0, 3) = "Qty"
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
'      .ColWidth(0) = 330
'      .ColWidth(1) = 1700
'      .ColWidth(2) = 3780
'      .ColWidth(3) = 600
'
'      .ColAlignment(1) = 1
'      .ColAlignment(2) = 1
'
'      .ColDefault(3) = 0
'
'      .ColNumberOnly(3) = True
'
'      .Row = 1
'      .Col = 1
'   End With
'End Sub
'
'Private Sub GridEditor1_EditorValidate(Cancel As Boolean)
'   With GridEditor1
'      Select Case .Col
'      Case 1, 2
'         oTrans.Detail(.Row - 1, .Col) = .TextMatrix(.Row, .Col)
'      Case 3
'         If Not IsNumeric(.TextMatrix(.Row, .Col)) Then .TextMatrix(.Row, .Col) = 0
'         oTrans.Detail(.Row - 1, .Col) = CDbl(.TextMatrix(.Row, .Col))
'      End Select
'   End With
'End Sub
'
'Private Sub GridEditor1_GotFocus()
'   pbGridGotFocus = True
'End Sub
'
'Private Sub oTrans_DetailRetrieved(ByVal Index As Integer)
'   With GridEditor1
'      .TextMatrix(.Row, Index) = oTrans.Detail(.Row - 1, Index)
'   End With
'End Sub
'
'Private Sub txtField_GotFocus(Index As Integer)
'   With txtField(Index)
'      .SelStart = 0
'      .SelLength = Len(.Text)
'   End With
'
'   pbGridGotFocus = False
'End Sub
'
'Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'   Dim lsOldProc As String
'
'   lsOldProc = "txtField_KeyDown"
'   'On Error GoTo errProc
'
'   If KeyCode = vbKeyF3 Then
'      If oTrans.SearchTransaction(txtField(Index).Text, False) Then
'         LoadDetail
'         txtField(Index).SelStart = 0
'         txtField(Index).SelLength = Len(txtField(Index).Text)
'      End If
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
'   Case vbKeyReturn, vbKeyDown
'      If GetFocus = GridEditor1.hwnd Then Exit Sub
'      SetNextFocus
'   Case vbKeyUp
'      SetPreviousFocus
'   End Select
'End Sub
'
'Private Sub InitValue()
'   With GridEditor1
'      .Rows = 2
'      .ColWidth(2) = 3780
'
'      .TextMatrix(1, 1) = ""
'      .TextMatrix(1, 2) = ""
'      .TextMatrix(1, 3) = 0
'   End With
'
'   txtField(0).Text = ""
'   txtField(0).Tag = ""
'End Sub
'
'Private Sub oTrans_MasterRetrieved(ByVal Index As Integer)
'   txtField(Index).Text = oTrans.Master(Index)
'End Sub
'
'Private Sub LoadDetail()
'   Dim lnCtr As Integer
'   Dim lsOldProc As String
'
'   lsOldProc = "LoadDetail"
'   'On Error GoTo errProc
'
'   txtField(0).Text = oTrans.Master("sModelNme")
'   txtField(0).Tag = txtField(0).Text
'
'   With GridEditor1
'      .Rows = oTrans.ItemCount + 1
'
'      For lnCtr = 0 To oTrans.ItemCount - 1
'         .TextMatrix(lnCtr + 1, 1) = oTrans.Detail(lnCtr, "sBarrCode")
'         .TextMatrix(lnCtr + 1, 2) = oTrans.Detail(lnCtr, "sDescript")
'         .TextMatrix(lnCtr + 1, 3) = IIf(IsNull(oTrans.Detail(lnCtr, "nQuantity")), 0, oTrans.Detail(lnCtr, "nQuantity"))
'      Next
'   End With
'
'   oTrans.UpdateTransaction
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " & " )"
'End Sub
'
'Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
'   Dim lsOldProc As String
'
'   lsOldProc = "txtField_Validate"
'   'On Error GoTo errProc
'
'   With txtField(Index)
'      If Trim(.Text) <> "" Then .Text = UCase(Left(.Text, 1)) & Right(.Text, Len(.Text) - 1)
'
'      If .Text = "" Then
'         InitValue
'      Else
'         If LCase(.Tag) <> LCase(.Text) Then
'            If oTrans.SearchTransaction(.Text, False) Then
'               LoadDetail
'            Else
'               InitValue
'            End If
'         End If
'      End If
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
'Private Sub GridEditor1_AddingRow(Cancel As Boolean)
'   With GridEditor1
'      If .TextMatrix(.Row, 1) = "" And CDbl(.TextMatrix(.Row, 3)) = 0 Then
'         Cancel = True
'      Else
'         If Not Cancel Then oTrans.AddDetail
'      End If
'   End With
'End Sub
'
'Private Sub GridEditor1_KeyDown(KeyCode As Integer, Shift As Integer)
'   Dim lsOldProc As String
'
'   lsOldProc = "GridEditor1_KeyDown"
'   'On Error GoTo errProc
'
'   If KeyCode = vbKeyF3 Then
'      With GridEditor1
'         Select Case .Col
'         Case 1, 2
'            oTrans.SearchDetail .Row - 1, .Col, .TextMatrix(.Row, .Col)
'
'            .SetFocus
'            .Refresh
'         End Select
'      End With
'      KeyCode = 0
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
'
