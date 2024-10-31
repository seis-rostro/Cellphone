VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0"; "xrGridControl.ocx"
Begin VB.Form frmSuppliesPurchase 
   BorderStyle     =   0  'None
   Caption         =   "Supplies Purchase"
   ClientHeight    =   6300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11040
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6300
   ScaleWidth      =   11040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   90
      TabIndex        =   0
      Top             =   1785
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
      Picture         =   "frmSupplies_Purchase.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   90
      TabIndex        =   8
      Top             =   1785
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
      Picture         =   "frmSupplies_Purchase.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   630
      Index           =   0
      Left            =   90
      TabIndex        =   9
      Top             =   495
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   1111
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
      Picture         =   "frmSupplies_Purchase.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   90
      TabIndex        =   10
      Top             =   1155
      Width           =   1260
      _ExtentX        =   2223
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
      Picture         =   "frmSupplies_Purchase.frx":166E
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   90
      TabIndex        =   11
      Top             =   2415
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
      Picture         =   "frmSupplies_Purchase.frx":1DE8
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   90
      TabIndex        =   12
      Top             =   1155
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
      Picture         =   "frmSupplies_Purchase.frx":2562
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   6
      Left            =   90
      TabIndex        =   13
      Top             =   2415
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
      Picture         =   "frmSupplies_Purchase.frx":2CDC
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   5745
      Index           =   1
      Left            =   1560
      Tag             =   "wt0;fb0"
      Top             =   495
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   10134
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   630
         Index           =   3
         Left            =   5460
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1005
         Width           =   3660
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
         Left            =   1200
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   375
         Width           =   1905
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   1200
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1320
         Width           =   4230
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   1200
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   930
         Width           =   1905
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   7
         Left            =   7095
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   2340
         Width           =   1995
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   6
         Left            =   4110
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   2340
         Width           =   2010
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   5
         Left            =   1215
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   2340
         Width           =   1950
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   4
         Left            =   1215
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   1920
         Width           =   3705
      End
      Begin xrGridEditor.GridEditor GridEditor1 
         Height          =   2775
         Left            =   90
         TabIndex        =   15
         Top             =   2850
         Width           =   9195
         _ExtentX        =   16219
         _ExtentY        =   4895
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
         Object.HEIGHT          =   2775
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
         MOUSEICON       =   "frmSupplies_Purchase.frx":3456
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
         Height          =   1035
         Index           =   1
         Left            =   105
         Top             =   1770
         Width           =   9165
      End
      Begin VB.Shape Shape2 
         Height          =   1620
         Index           =   0
         Left            =   105
         Top             =   135
         Width           =   9165
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   195
         Index           =   5
         Left            =   4680
         TabIndex        =   23
         Top             =   990
         Width           =   630
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier"
         Height          =   195
         Index           =   9
         Left            =   240
         TabIndex        =   22
         Top             =   1380
         Width           =   570
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
         TabIndex        =   21
         Top             =   405
         Width           =   1110
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date "
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   20
         Top             =   990
         Width           =   810
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Price"
         Height          =   195
         Index           =   6
         Left            =   6270
         TabIndex        =   19
         Top             =   2400
         Width           =   690
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
         Height          =   195
         Index           =   2
         Left            =   3300
         TabIndex        =   18
         Top             =   2400
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
         Top             =   2400
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
         Top             =   1980
         Width           =   795
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   315
         Left            =   1335
         Tag             =   "et0;ht2"
         Top             =   480
         Width           =   1920
      End
   End
End
Attribute VB_Name = "frmSuppliesPurchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Private Const pxeMODULENAME = "frmSuppliesPurchase"
'
'Private WithEvents oTrans  As ggcSuppliesPurchase
'Private oSkin As clsFormSkin
'
'Dim pnIndex As Integer
'Dim pbGridFocus As Boolean
'Dim pnCtr As Integer
'Dim pbSave As Boolean
'Dim pbGridValidate As Boolean
'Dim pbPosted As Boolean
'
'Private Sub cmdButton_Click(Index As Integer)
'   Dim lsOldProc As String
'   Dim lsRep As String
'
'   lsOldProc = "cmdButton_Click"
'   'On Error GoTo errProc
'
'   If Not pbGridFocus And Index = 0 Then Call txtField_Validate(pnIndex, False)
'   With GridEditor1
'      Select Case Index
'      Case 0
'         If .Rows > 2 Then
'            pnCtr = 0
'            Do While pnCtr < .Rows
'               If Trim(.TextMatrix(pnCtr, 1)) = "" Then
'                  .Row = pnCtr
'                  If oTrans.deleteDetail(.Row - 1) Then .DeleteRow
'               Else
'                  pnCtr = pnCtr + 1
'               End If
'            Loop
'
'            .ColWidth(3) = 2300
'            If .Rows > 16 Then .ColWidth(3) = 2100
'         End If
'
'         If isEntryOK Then
'            If oTrans.SaveTransaction Then
'               MsgBox "Record Updated Successfully!!!", vbInformation, "Confirm"
'               InitButton xeModeReady
'
'               lsRep = MsgBox("Print Transaction!!!", vbYesNo + vbQuestion, "Confirm")
'               If lsRep = vbYes Then
'                  If Not PrintTrans Then MsgBox "Unable to Print Transaction!!!", vbCritical, "Warning"
'               End If
'               pbSave = True
'            Else
'               MsgBox "Unable to Update Record!!!", vbCritical, "Warning"
'            End If
'         End If
'      Case 1
'         If pbGridFocus Then
'            If oTrans.searchDetail(.Row - 1, 1) Then .Col = 1
'            .Refresh
'            .SetFocus
'         Else
'            oTrans.SearchMaster pnIndex
'         End If
'      Case 2
'         If .Rows > 2 Then
'            If oTrans.deleteDetail(.Row - 1) Then .DeleteRow
'
'            For pnCtr = 1 To .Rows - 1
'               .TextMatrix(pnCtr, 0) = pnCtr
'            Next
'
'            .ColWidth(3) = 2300
'            If .Rows > 16 Then .ColWidth(3) = 2100
'         End If
'      Case 3
'         lsRep = MsgBox("Transaction is in Update Mode!!!" & vbCrLf & _
'                        "Do you want to Cancel Transaction!!!", vbYesNo + vbQuestion, "Confirm")
'
'         If lsRep = vbYes Then
'            oTrans.NewTransaction
'            ClearFields
'            InitButton xeModeReady
'         Else
'            txtField(pnIndex).SetFocus
'         End If
'         pbSave = False
'      Case 4
'         oTrans.NewTransaction
'         ClearFields
'         InitButton xeModeAddNew
'         txtField(1).SetFocus
'      Case 5
'         If pbSave Then
'            lsRep = MsgBox("Print Transaction!!!", vbYesNo + vbQuestion, "Confirm")
'            If lsRep = vbYes Then
'               If Not PrintTrans Then MsgBox "Unable to Print Transaction!!!", vbCritical, "Warning"
'            End If
'         Else
'            MsgBox "Unable to Print Transaction!!!", vbCritical, "Warning"
'         End If
'      Case 6
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
'   GridEditor1.Refresh
'
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
'   Set oTrans = New ggcSuppliesPurchase
'   Set oTrans.AppDriver = oApp
'
'   oTrans.InitTransaction
''   oTrans.NewTransaction
'
'   Set oSkin = New clsFormSkin
'   Set oSkin.AppDriver = oApp
'   Set oSkin.Form = Me
'   oSkin.ApplySkin xeFormTransEqualLeft
'
'   InitGrid
'   ClearFields
'   InitButton xeModeAddNew
'
'
'   pbGridValidate = False
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
'      ElseIf .TextMatrix(.Row, 7) = 0 Then
'         Cancel = True
'      End If
'
'      If Not Cancel Then
'         If .Row = .Rows - 1 Then oTrans.addDetail
'      End If
'
'      If .Rows > 16 Then .ColWidth(3) = 2100
'   End With
'End Sub
'
'Private Sub GridEditor1_EditorValidate(Cancel As Boolean)
'   Dim lsOldProc As String
'
'   lsOldProc = "GridEditor1_EditorValidate"
'   'On Error GoTo errProc
'
'   With GridEditor1
'      If pbGridValidate Then
'         pbGridValidate = False
'         Exit Sub
'      End If
'
'      If .Col = 1 Or .Col = 2 Then
'         .TextMatrix(.Row, .Col) = compareSerial(.TextMatrix(.Row, .Col), .Row)
'      End If
'
'      oTrans.Detail(.Row - 1, .Col) = .TextMatrix(.Row, .Col)
'      If oTrans.Detail(.Row - 1, "cHsSerial") = xeYes Then
'         Select Case .Col
'         Case 1, 2
'            If .TextMatrix(.Row, .Col) <> "" Then
'               If .Row = .Rows - 1 Then
'                  .Rows = .Rows + 1
'                  oTrans.addDetail
'                  .Col = 0
'               End If
'
'               .Row = .Rows - 1
'            End If
'         Case 7
'            If CDbl(.TextMatrix(.Row, .Col)) <> 1 Then .TextMatrix(.Row, .Col) = 1
'         End Select
'      End If
'
'      If .Rows > 16 Then
'         .TopRow = .Rows - 1
'         .ColWidth(3) = 2100
'      End If
'   End With
'   pbGridValidate = True
'
'endProc:
'   GridEditor1.Refresh
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " & Cancel & " )", True
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
'   If KeyCode = vbKeyF3 Then
'      With GridEditor1
'         If oTrans.searchDetail(.Row - 1, .Col, .TextMatrix(.Row, .Col)) Then
'            If oTrans.Detail(.Row - 1, "cHsSerial") = xeYes Then
'               If .Row = .Rows - 1 Then
'                  .Rows = .Rows + 1
'                  oTrans.addDetail
'               End If
'
'               .Row = .Rows - 1
'               .Col = 1
'            Else
'               .Col = 7
'            End If
'         Else
'            oTrans.Detail(.Row - 1, .Col) = .TextMatrix(.Row, .Col)
'            .Col = 1
'         End If
'
'         .Refresh
'         .SetFocus
'         If .Rows > 16 Then
'            .TopRow = .Rows - 1
'            .ColWidth(3) = 2100
'         End If
'         KeyCode = 0
'      End With
'   End If
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
'      If cmdButton(0).Visible Then oTrans.Detail(.Row - 1, .Col) = .TextMatrix(.Row, .Col)
'      If .Rows > 16 Then .TopRow = .Rows - 1
'   End With
'
'   pbGridValidate = False
'End Sub
'
''Private Sub oTrans_DetailRetrieved(ByVal Index As Integer)
''   With GridEditor1
''      .TextMatrix(.Row, Index) = oTrans.Detail(.Row - 1, Index)
''   End With
''End Sub
'
''Private Sub oTrans_MasterRetrieved(ByVal Index As Integer)
''   txtField(Index).Text = oTrans.Master(Index)
''End Sub
'
'Private Sub InitGrid()
'   With GridEditor1
'      .Rows = 2
'      .Cols = 5
'      .Font = "MS Sans Serif"
'
'      'column title
'      .TextMatrix(0, 1) = "Entry No"
'      .TextMatrix(0, 2) = "Description"
'      .TextMatrix(0, 3) = "Quantity"
'      .TextMatrix(0, 4) = "Unit Price"
'
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
'      .ColWidth(1) = 2000
'      .ColWidth(2) = 2000
'      .ColWidth(4) = 1020
'      .ColWidth(5) = 500
'
'      .ColFormat(4) = "#,##0.00"
'      .ColFormat(5) = "#,##0"
'      .ColDefault(4) = 0#
'      .ColDefault(5) = 0
'
'      .ColAlignment(1) = 1
'      .ColAlignment(2) = 1
'      .ColAlignment(3) = 1
'      .ColAlignment(4) = 6
'      .ColAlignment(5) = 6
'
'      .ColEnabled(3) = False
'      .ColEnabled(4) = False
'      .ColEnabled(5) = False
'
'      .EditorBackColor = oApp.getColor("HT1")
'
'      .Row = 1
'      .Col = 1
'   End With
'End Sub
'
'Private Sub txtField_GotFocus(Index As Integer)
'   With txtField(Index)
'      If Index = 1 Then .Text = Format(.Text, "MM/DD/YY")
'
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
'   'On Error GoTo errProc
'
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
'Private Sub InitButton(lnStat As Integer)
'   Dim lbShow As Boolean
'
'   lbShow = IIf(lnStat = 0, False, True)
'   cmdButton(4).Visible = Not lbShow
'   cmdButton(5).Visible = Not lbShow
'   cmdButton(6).Visible = Not lbShow
'
'   cmdButton(0).Visible = lbShow
'   cmdButton(1).Visible = lbShow
'   cmdButton(2).Visible = lbShow
'   cmdButton(3).Visible = lbShow
'
'   For pnCtr = 1 To txtField.Count - 3
'      txtField(pnCtr).Enabled = lbShow
'   Next
'
'   With GridEditor1
'      .ColEnabled(1) = lbShow
'      .ColEnabled(2) = lbShow
'      .ColEnabled(7) = lbShow
'   End With
'
'   If Not lbShow Then cmdButton(4).SetFocus
'End Sub
'
'Private Function PrintTrans() As Boolean
'   Dim loreport As frmRepViewer
'
'   Dim lrs As ADODB.Recordset
'   Dim lors As ADODB.Recordset
'   Dim lorsOthers As ADODB.Recordset
'   Dim lnCtr As Integer
'   Dim lsOldProc As String
'   Dim lsStockIDx As String
'
'   lsOldProc = "PrintTrans"
'   'On Error GoTo errProc
'
'   PrintTrans = True
'   Set lrs = New ADODB.Recordset
'   lrs.Fields.Append "nField01", adInteger, 3
'   lrs.Fields.Append "nField02", adChar, 1
'   lrs.Fields.Append "sField01", adVarChar, 20
'   lrs.Fields.Append "sField02", adVarChar, 128
'   lrs.Fields.Append "sField03", adVarChar, 20
'   lrs.Fields.Append "sField04", adVarChar, 128
'   lrs.Fields.Append "sField05", adVarChar, 100
'   lrs.Fields.Append "sField06", adVarChar, 25
'   lrs.Fields.Append "sField07", adVarChar, 25
'   lrs.Open
'
'   With oTrans
'      For lnCtr = 0 To .ItemCount - 1
'         lrs.AddNew
'         lrs.Fields("nField01") = oTrans.Detail(lnCtr, "nQuantity")
'         lrs.Fields("nField02") = oTrans.Detail(lnCtr, "cHsSerial")
'         lrs.Fields("sField01") = oTrans.Detail(lnCtr, "sTransNox")
'         lrs.Fields("sField02") = oTrans.Detail(lnCtr, "sStockIDx")
'         lrs.Fields("sField03") = oTrans.Detail(lnCtr, "sBarrCode")
'         lrs.Fields("sField04") = oTrans.Detail(lnCtr, "sDescript")
'         lrs.Fields("sField05") = oTrans.Detail(lnCtr, "sBrandNme")
'         lrs.Fields("sField06") = IFNull(oTrans.Detail(lnCtr, "sSerialNo"), "")
'         lrs.Fields("sField07") = IFNull(oTrans.Detail(lnCtr, "sReferNox"), "")
'
'         Set lorsOthers = New Recordset
'         lorsOthers.Open "SELECT" & _
'                              "  a.sReferNox" & _
'                              ", a.sSalesInv" & _
'                           " FROM CP_PO_Receiving_Master a" & _
'                              ", CP_PO_Receiving_Detail b" & _
'                           " WHERE a.sTransNox = b.sTransNox" & _
'                              " AND a.cTranStat <> " & strParm(xeStateCancelled) & _
'                              " AND b.sSerialID = " & strParm(IFNull(oTrans.Detail(lnCtr, "sSerialID"), "")) & _
'                           " ORDER BY a.sTransNox DESC" & _
'                           " LIMIT 1" _
'         , oApp.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'         If Not lorsOthers.EOF Then
'
'         End If
'         Set lorsOthers = Nothing
'      Next
'      lrs.Sort = "nField02,sField05,sField03,sField06"
'   End With
'
'   Set lors = New ADODB.Recordset
'   If lors.State = adStateOpen Then lors.Close
'
'   lors.Open "SELECT" _
'               & "  CONCAT(a.sAddressx, ', ', b.sTownName, ', ', c.sProvName, ' ', b.sZippCode) as xAddressx" _
'               & ", a.sBranchNm" _
'            & " From Branch a" _
'               & ", TownCity b" _
'               & ", Province c" _
'            & " WHERE a.sBranchCd = " & strParm(Left(oTrans.Master("sTransNox"), Len(oApp.BranchCode))) _
'               & " AND a.sTownIDxx = b.sTownIDxx" _
'               & " AND b.sProvIDxx = c.sProvIDxx" _
'            , oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
'
'   Set lors = New ADODB.Recordset
'   If lors.State = adStateOpen Then lors.Close
'
'   lors.Open "SELECT" _
'               & "  CONCAT(a.sAddressx, ', ', b.sTownName, ', ', c.sProvName, ' ', b.sZippCode) as xAddressx" _
'               & ", a.sBranchNm" _
'            & " From Branch a" _
'               & ", TownCity b" _
'               & ", Province c" _
'            & " WHERE a.sBranchCd = " & strParm(Left(oTrans.Master("sTransNox"), Len(oApp.BranchCode))) _
'               & " AND a.sTownIDxx = b.sTownIDxx" _
'               & " AND b.sProvIDxx = c.sProvIDxx" _
'            , oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
'
'   Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\CPPurchaseReturnFormOld.rpt")
'   'assign important info to the report
'   oReport.DiscardSavedData
'   oReport.FieldMappingType = crAutoFieldMapping
'   oReport.Database.SetDataSource lrs
'
'   oReport.Sections("PHa").ReportObjects("txtTransNox").SetText "CP-" & Right(oTrans.Master("sTransNox"), 10)
'   oReport.Sections("PHa").ReportObjects("txtDate").SetText txtField(1).Text
'   oReport.Sections("PHb").ReportObjects("txtTo").SetText txtField(2).Text
'   oReport.Sections("PHb").ReportObjects("txtToAddress").SetText txtField(3).Text
'   oReport.Sections("PHb").ReportObjects("txtFrom").SetText lors("sBranchNm")
'   oReport.Sections("PHb").ReportObjects("txtFromAddress").SetText lors("xAddressx")
'   oReport.Sections("RFb").ReportObjects("txtRemarks").SetText txtField(4).Text
'   oReport.Sections("PF").ReportObjects("txtRptUser").SetText oApp.UserName
'
'   Set loreport = New frmRepViewer
'   Set loreport.ReportSource = oReport
'   loreport.Show
'
'endPoc:
'   If Not pbPosted Then
'      oTrans.CloseTransaction (oTrans.Master(0))
'      pbPosted = True
'   End If
'   Set oReport = Nothing
'   Set lrs = Nothing
'   Set lors = Nothing
'   Set loreport = Nothing
'   Exit Function
'errProc:
'   PrintTrans = False
'   ShowError lsOldProc & "( " & " )"
'End Function
'
'Private Sub ClearFields()
'   For pnCtr = 0 To txtField.Count - 1
'      Select Case pnCtr
'      Case 0
'         txtField(pnCtr).Text = Format(oTrans.Master(pnCtr), "@@@@-@@@@@@")
'      Case 1
'         txtField(pnCtr).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
'      Case Else
'         txtField(pnCtr).Text = Empty
'      End Select
'   Next
'
'   With GridEditor1
'      .Rows = 2
'      .ColWidth(3) = 2300
'
'      'empty row
'      .TextMatrix(1, 1) = ""
'      .TextMatrix(1, 2) = ""
'      .TextMatrix(1, 3) = ""
'      .TextMatrix(1, 4) = "0.00"
'      .TextMatrix(1, 5) = "0"
'      .TextMatrix(1, 6) = ""
'      .TextMatrix(1, 7) = "0"
'   End With
'
'   pbSave = False
'   pbPosted = False
'
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
'      Case 1
'         If Not IsDate(.Text) Then .Text = oApp.ServerDate
'         .Text = Format(.Text, "MMMM DD, YYYY")
'      Case 3
'         .Text = Format(.Text, ">")
'      End Select
'
'      oTrans.Master(Index) = .Text
'   End With
'End Sub
'
'Private Function isEntryOK() As Boolean
'   If txtField(2).Text = "" Then
'      MsgBox "Supplier not found!!!" & vbCrLf & _
'             "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
'      txtField(2).SetFocus
'      GoTo EntryNotOK
'   End If
'
'   With GridEditor1
'      If Trim(.TextMatrix(1, 1)) = "" Then
'         MsgBox "Detail is required!!!" & vbCrLf & _
'                "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
'         .SetFocus
'         .Row = 1
'         .Col = 1
'         GoTo EntryNotOK
'      End If
'   End With
'
'EntryOK:
'   isEntryOK = True
'   Exit Function
'EntryNotOK:
'   isEntryOK = False
'End Function
'
'Private Function compareSerial(Value As String, Row As Integer) As String
'   Dim lnRep As Integer
'   Dim lnCtr As Integer
'   Dim lsValue As String
'   Dim lnValue As Integer
'
'   If Trim(Value) = "" Then
'      compareSerial = ""
'      Exit Function
'   End If
'
'   With GridEditor1
'      For lnCtr = 1 To .Rows - 1
'         If .TextMatrix(lnCtr, 1) = Value And lnCtr <> Row Then
'            If oTrans.Detail(lnCtr - 1, "cHsSerial") = xeYes Then
'               MsgBox "Duplicate Serial No!!!" & vbCrLf & _
'                        "Please Verify your entry then try again!!!", vbCritical, "Warning"
'            Else
'               lnRep = MsgBox("Duplicate Serial No!!!" & vbCrLf & _
'                                 "Item automatically add from existing serial!!!", vbYesNo + vbQuestion, "CONFIRMATION")
'               If lnRep = vbYes Then
'                  lsValue = InputBox("Please specify quantity for serial " & Value & vbCrLf & _
'                                       .TextMatrix(lnCtr, 2) & vbCrLf & _
'                                       .TextMatrix(lnCtr, 3), "Quantity", 0)
'                  lnValue = IIf(lsValue = "", 0, lsValue)
'
'                  .TextMatrix(lnCtr, 6) = .TextMatrix(lnCtr, 6) + lnValue
'                  oTrans.Detail(lnCtr - 1, "nQuantity") = CDbl(.TextMatrix(lnCtr, 6))
'               End If
'            End If
'            compareSerial = ""
'         Else
'            compareSerial = Value
'         End If
'      Next
'   End With
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
'
'
'
'
