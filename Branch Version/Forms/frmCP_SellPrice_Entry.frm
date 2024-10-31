VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0"; "xrGridControl.ocx"
Begin VB.Form frmCP_SellPrice_Entry 
   BorderStyle     =   0  'None
   Caption         =   "Delivery"
   ClientHeight    =   6900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11715
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   11715
   ShowInTaskbar   =   0   'False
   Begin xrGridEditor.GridEditor GridEditor1 
      Height          =   4980
      Left            =   1575
      TabIndex        =   6
      Tag             =   "et0;eb0;et0;bc2"
      Top             =   1800
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   8784
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
      Object.HEIGHT          =   4980
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
      MOUSEICON       =   "frmCP_SellPrice_Entry.frx":0000
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
      Index           =   2
      Left            =   105
      TabIndex        =   11
      Top             =   4050
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
      Picture         =   "frmCP_SellPrice_Entry.frx":001C
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1200
      Index           =   0
      Left            =   1575
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   10050
      _ExtentX        =   17727
      _ExtentY        =   2117
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   1665
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   165
         Width           =   2505
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   5055
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   675
         Width           =   4770
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   1665
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   675
         Width           =   2505
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   315
         Left            =   1725
         Tag             =   "et0;ht2"
         Top             =   240
         Width           =   2505
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Implementation Date"
         Height          =   285
         Index           =   1
         Left            =   165
         TabIndex        =   0
         Top             =   195
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
         Height          =   285
         Index           =   3
         Left            =   195
         TabIndex        =   2
         Top             =   690
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Brand"
         Height          =   285
         Index           =   4
         Left            =   4500
         TabIndex        =   4
         Top             =   690
         Width           =   540
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   105
      TabIndex        =   7
      Top             =   1530
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
      Picture         =   "frmCP_SellPrice_Entry.frx":0796
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   105
      TabIndex        =   10
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
      Picture         =   "frmCP_SellPrice_Entry.frx":0F10
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   105
      TabIndex        =   12
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
      Picture         =   "frmCP_SellPrice_Entry.frx":168A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   105
      TabIndex        =   13
      Top             =   4050
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
      Picture         =   "frmCP_SellPrice_Entry.frx":1E04
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   6
      Left            =   105
      TabIndex        =   14
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
      Picture         =   "frmCP_SellPrice_Entry.frx":257E
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   7
      Left            =   105
      TabIndex        =   9
      Top             =   2790
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Retrieve"
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
      Picture         =   "frmCP_SellPrice_Entry.frx":2CF8
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   8
      Left            =   105
      TabIndex        =   8
      Top             =   2160
      Visible         =   0   'False
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Basis"
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
      Picture         =   "frmCP_SellPrice_Entry.frx":3472
   End
End
Attribute VB_Name = "frmCP_SellPrice_Entry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Private Const pxeMODULENAME = "frmCP_SellPrice_Entry"
'
'Private WithEvents oTrans As ggcPriceUpdate.clsCPPriceUpdate
'Private oSkin As clsFormSkin
'
'Dim pnIndex As Integer
'Dim pbGridFocus As Boolean
'Dim pnCtr As Integer
'Dim pbSave As Boolean
'Dim pbGridValidate As Boolean
'
'Private Sub cmdButton_Click(Index As Integer)
'   Dim lsOldProc As String
'   Dim lnRep As String
'
'   lsOldProc = "cmdButton_Click"
'   ''On Error GoTo errProc
'
'   With GridEditor1
'      Select Case Index
'      Case 0 ' Save
'         If .Rows > 2 Then
'            pnCtr = 0
'            Do While pnCtr < .Rows
'               If Trim(.TextMatrix(pnCtr, 1)) = "" Then
'                  .Row = pnCtr
'                  If oTrans.deleteDetail(.Row - 1) Then .deleteRow
'               Else
'                  pnCtr = pnCtr + 1
'               End If
'            Loop
'
'            .ColWidth(2) = 3100
'            If .Rows > 20 Then .ColWidth(2) = 2850
'         End If
'
'         If isEntryOk Then
'            If oTrans.SaveTransaction Then
'               MsgBox "Record Updated Successfully!!!", vbInformation, "Confirm"
'
'               initButton xeModeReady
''               lnRep = MsgBox("Print Transaction!!!", vbYesNo + vbQuestion, "Confirm")
''               If lnRep = vbYes Then
''                  If Not PrintTrans Then MsgBox "Unable to Print Transaction!!!", vbCritical, "Warning"
''               End If
'               pbSave = True
'            Else
'               MsgBox "Unable to Update Record!!!", vbCritical, "Warning"
'            End If
'         End If
'      Case 1 ' Search
'         If pbGridFocus Then
'            If oTrans.searchDetail(.Row - 1, 1) Then .Col = 1
'            .Refresh
'            .SetFocus
'         Else
'            oTrans.SearchMaster pnIndex
'         End If
'      Case 2 ' Delete Row
'         If .Rows > 2 Then
'            If oTrans.deleteDetail(.Row - 1) Then .deleteRow
'
'            For pnCtr = 1 To .Rows - 1
'               .TextMatrix(pnCtr, 0) = pnCtr
'            Next
'
'            .ColWidth(2) = 3100
'            If .Rows > 20 Then .ColWidth(2) = 2850
'         End If
'      Case 3 ' Cancel
'         lnRep = MsgBox("Transaction is in Update Mode!!!" & vbCrLf & _
'                        "Do you want to Cancel Transaction!!!", vbYesNo + vbQuestion, "Confirm")
'
'         If lnRep = vbYes Then
'            oTrans.InitTransaction
'            Call ClearItem(-1)
'            initButton xeModeReady
'         Else
'            txtField(pnIndex).SetFocus
'         End If
'         pbSave = False
'      Case 4 ' New Entry
'         oTrans.InitTransaction
'         Call ClearItem(-1)
'         Call ClearFields
'         initButton xeModeAddNew
'
'         txtField(0).SetFocus
'      Case 5 ' Print Transaction
'         If pbSave Then
'            lnRep = MsgBox("Print Transaction!!!", vbYesNo + vbQuestion, "Confirm")
'            If lnRep = vbYes Then
'               If Not PrintTrans Then MsgBox "Unable to Print Transaction!!!", vbCritical, "Warning"
'            End If
'         Else
'            MsgBox "Unable to Print Transaction!!!", vbCritical, "Warning"
'         End If
'      Case 6 ' Close
'         Unload Me
'      Case 7 ' Retrieve
'         If oTrans.LoadDetai() Then
'            Call LoadItem(-1)
'         Else
'            Call ClearItem(-1)
'         End If
'      Case 8 ' Basis
'         If oTrans.SetPriceBasis() Then
'            Call LoadItem(-1)
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
'   ''On Error GoTo errProc
'
'   CenterChildForm mdiMain, Me
'
'   Set oTrans = New ggcPriceUpdate.clsCPPriceUpdate
'   Set oTrans.AppDriver = oApp
'   oTrans.InitTransaction
'
'   Set oSkin = New clsFormSkin
'   Set oSkin.AppDriver = oApp
'   Set oSkin.Form = Me
'   oSkin.ApplySkin xeFormTransaction
'
'   InitGrid
'   Call ClearItem(-1)
'   Call ClearFields
'   initButton xeModeAddNew
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
'      If Trim(.TextMatrix(.Row, 1)) = "" Then
'         Cancel = True
'      ElseIf CLng(.TextMatrix(.Row, 3)) = 0 Or CLng(.TextMatrix(.Row, 4)) = 0 Then
'         Cancel = True
'      End If
'
'      If Not Cancel Then
'         Debug.Print "adding rows"
'         If .Row = .Rows - 1 Then oTrans.addDetail
'      End If
'
'      If .Rows > 20 Then .ColWidth(2) = 2850
'   End With
'End Sub
'
'Private Sub GridEditor1_EditorValidate(Cancel As Boolean)
'   Dim lsOldProc As String
'
'   lsOldProc = "GridEditor1_EditorValidate"
'   ''On Error GoTo errProc
'
'   With GridEditor1
'      If pbGridValidate Then
'         pbGridValidate = False
'         Exit Sub
'      End If
'
'      Select Case .Col
'      Case 1, 2
'         If .TextMatrix(.Row, .Col) <> "" Then
'            If ItemExist(.TextMatrix(.Row, .Col), .Row, .Col) Then
'               .TextMatrix(.Row, .Col) = ""
'               Call ClearItem(.Row)
'               .SetFocus
'               GoTo endProc
'            End If
'
'            If Trim(.TextMatrix(.Row, .Col)) <> Trim(oTrans.Detail(.Row - 1, .Col)) Then
'               oTrans.Detail(.Row - 1, .Col) = .TextMatrix(.Row, .Col)
'               Call LoadItem(.Row)
'            End If
'         End If
'      Case 3, 4, 5, 6, 7
'         If IsNumeric(.TextMatrix(.Row, .Col)) Then
'            oTrans.Detail(.Row - 1, .Col) = CDbl(.TextMatrix(.Row, .Col))
'         End If
'         .TextMatrix(.Row, .Col) = Format(oTrans.Detail(.Row - 1, .Col), "#,##0.00")
'      End Select
'
'      If .Rows > 20 Then
'         '.TopRow = .Rows - 1
'         .ColWidth(2) = 2850
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
'   ''On Error GoTo errProc
'
'   If KeyCode = vbKeyF3 Then
'      With GridEditor1
'         If .Col = 1 Or .Col = 2 Then
'            Call oTrans.searchDetail(.Row - 1, .Col, .TextMatrix(.Row, .Col))
'         End If
'
'         .SetFocus
'         If .Rows > 20 Then
'            '.TopRow = .Rows - 1
'            .ColWidth(2) = 2850
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
''      If .Rows > 20 Then '.TopRow = .Rows - 1
'   End With
'
'   pbGridValidate = False
'End Sub
'
'Private Sub oTrans_DetailRetrieved()
'   Dim lnCtr As Integer
'
'   With GridEditor1
'      For lnCtr = 1 To .Cols - 1
'         If lnCtr < 3 Then
'            .TextMatrix(.Row, lnCtr) = oTrans.Detail(.Row - 1, lnCtr)
'         Else
'            .TextMatrix(.Row, lnCtr) = Format(oTrans.Detail(.Row - 1, lnCtr), "#,##0.00")
'         End If
'      Next
'      .SetFocus
'   End With
'End Sub
'
'Private Sub oTrans_MasterRetrieved(ByVal Index As Integer)
'   txtField(Index).Text = oTrans.Master(Index)
'End Sub
'
'Private Sub InitGrid()
'   With GridEditor1
'      .Rows = 2
'      .Cols = 8
'      .Font = "MS Sans Serif"
'
'      'column title
'      .TextMatrix(0, 1) = "BarrCode"
'      .TextMatrix(0, 2) = "Description"
'      .TextMatrix(0, 3) = "Last Price"
'      .TextMatrix(0, 4) = "SRP"
'      .TextMatrix(0, 5) = "3-mo 0%"
'      .TextMatrix(0, 6) = "6-mo 0%"
'      .TextMatrix(0, 7) = "12-mo 0%"
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
'      .ColWidth(1) = 1800
'      .ColWidth(2) = 2500
'      .ColWidth(3) = 950
'      .ColWidth(4) = 950
'      .ColWidth(5) = 950
'      .ColWidth(6) = 950
'      .ColWidth(7) = 950
'
'      .ColFormat(3) = "#,##0.00"
'      .ColFormat(4) = "#,##0.00"
'      .ColFormat(5) = "#,##0.00"
'      .ColFormat(6) = "#,##0.00"
'      .ColFormat(7) = "#,##0.00"
'
'      .ColNumberOnly(3) = True
'      .ColNumberOnly(4) = True
'      .ColNumberOnly(5) = True
'      .ColNumberOnly(6) = True
'      .ColNumberOnly(7) = True
'
'      .ColDefault(3) = 0#
'      .ColDefault(4) = 0#
'      .ColDefault(5) = 0#
'      .ColDefault(6) = 0#
'      .ColDefault(7) = 0#
'
'      .ColAlignment(1) = 1
'      .ColAlignment(2) = 1
'      .ColAlignment(3) = 6
'      .ColAlignment(4) = 6
'      .ColAlignment(5) = 6
'      .ColAlignment(6) = 6
'      .ColAlignment(7) = 6
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
'      If Index = 0 Then .Text = Format(.Text, "MM/DD/YY")
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
'   ''On Error GoTo errProc
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
'Private Sub initButton(lnStat As Integer)
'   Dim lbShow As Boolean
'
'   lbShow = IIf(lnStat = 0, False, True)
'   cmdButton(4).Visible = Not lbShow
'   cmdButton(6).Visible = Not lbShow
'
'   cmdButton(0).Visible = lbShow
'   cmdButton(1).Visible = lbShow
'   cmdButton(2).Visible = lbShow
'   cmdButton(3).Visible = lbShow
'   cmdButton(7).Visible = lbShow
'   cmdButton(8).Visible = lbShow
'
'   For pnCtr = 1 To txtField.Count - 1
'      txtField(pnCtr).Enabled = lbShow
'   Next
'
'   With GridEditor1
'      .ColEnabled(1) = lbShow
'      .ColEnabled(2) = lbShow
'      .ColEnabled(3) = lbShow
'      .ColEnabled(4) = lbShow
'      .ColEnabled(5) = lbShow
'      .ColEnabled(6) = lbShow
'      .ColEnabled(7) = lbShow
'   End With
'
'   If Not lbShow Then cmdButton(4).SetFocus
'End Sub
'
'Private Function PrintTrans() As Boolean
'   Dim loreport As frmRepViewer
'   Dim lrs As Recordset
'   Dim loRS As Recordset
'   Dim lnCtr As Integer
'   Dim lsOldProc As String
'   Dim lnTotlWSerial As Double
'   Dim lnTotlWOSerial As Double
'
'   lsOldProc = "PrinTrans"
'   ''On Error GoTo errProc
'
'   PrintTrans = True
'
'   Set lrs = New ADODB.Recordset
'
'   lrs.Fields.Append "nField01", adInteger, 3
'   lrs.Fields.Append "nField02", adChar, 1
'   lrs.Fields.Append "sField01", adVarChar, 20
'   lrs.Fields.Append "sField02", adVarChar, 128
'   lrs.Fields.Append "sField03", adVarChar, 20
'   lrs.Fields.Append "sField04", adVarChar, 12
'   lrs.Fields.Append "sField05", adVarChar, 100
'   lrs.Open
'
'   With oTrans
'      lnTotlWOSerial = 0
'      lnTotlWSerial = 0
'
'      For lnCtr = 0 To .ItemCount - 1
'         lrs.AddNew
'         lrs.Fields("nField01") = oTrans.Detail(lnCtr, "nQuantity")
'         lrs.Fields("nField02") = oTrans.Detail(lnCtr, "cHsSerial")
'         lrs.Fields("sField01") = oTrans.Detail(lnCtr, "sBarrCode")
'         lrs.Fields("sField02") = oTrans.Detail(lnCtr, "sDescript")
'         lrs.Fields("sField03") = oTrans.Detail(lnCtr, "sSerialNo")
'         lrs.Fields("sField04") = oTrans.Detail(lnCtr, "sStockIDx")
'         lrs.Fields("sField05") = oTrans.Detail(lnCtr, "sBrandNme")
'         If oTrans.Detail(lnCtr, "cHsSerial") = xeYes Then
'            lnTotlWSerial = lnTotlWSerial + 1
'         Else
'            lnTotlWOSerial = lnTotlWOSerial + CDbl(oTrans.Detail(lnCtr, "nQuantity"))
'         End If
'      Next
'      lrs.Sort = "nField02 DESC,sField05,sField05,sField03"
'   End With
'
'   'assign important info to the report
'   Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\CP_Transfer.rpt")
'   oReport.DiscardSavedData
'   oReport.FieldMappingType = crAutoFieldMapping
'   oReport.Database.SetDataSource lrs
'
'   Set loRS = New ADODB.Recordset
'   If loRS.State = adStateOpen Then loRS.Close
'
'   loRS.Open "SELECT" _
'               & "  a.sAddressx" _
'               & ", CONCAT(b.sTownName, ', ' , c.sProvName, ' ' , b.sZippCode) xTownName" _
'               & ", a.sBranchNm" _
'            & " FROM Branch a" _
'               & " LEFT JOIN TownCity b" _
'                  & " LEFT JOIN Province c" _
'                     & " ON b.sProvIDxx = c.sProvIDxx" _
'                  & " ON a.sTownIDxx = b.sTownIDxx" _
'            & " WHERE a.sBranchCd = " & strParm(oTrans.Master("sDestinat")) _
'            , oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
'
'   oReport.Sections("PHa").ReportObjects("txtRefNo").SetText "CP" & "-" & Right(oTrans.Master("sTransNox"), 10)
'   oReport.Sections("PHa").ReportObjects("txtDate").SetText txtField(1).Text
'   oReport.Sections("PHb").ReportObjects("txtTo").SetText loRS("sBranchNm")
'   oReport.Sections("PHb").ReportObjects("txtToAddress").SetText loRS("sAddressx") & IFNull(loRS("xTownName"), "")
'   oReport.Sections("PHb").ReportObjects("txtFrom").SetText oApp.ClientName
'   oReport.Sections("PHb").ReportObjects("txtFromAddress").SetText oApp.Address & ", " & oApp.TownCity & ", " & oApp.Province & " " & oApp.ZippCode
'   oReport.Sections("RFb").ReportObjects("txtRemarks").SetText txtField(4).Text
'   oReport.Sections("RFb").ReportObjects("txtWithSerial").SetText IIf(lnTotlWSerial = 0, "", Format(lnTotlWSerial, "#,##0"))
'   oReport.Sections("RFb").ReportObjects("txtWOutSerial").SetText IIf(lnTotlWOSerial = 0, "", Format(lnTotlWOSerial, "#,##0"))
'   oReport.Sections("PF").ReportObjects("txtRptUser").SetText oApp.UserName
'
'   Set loreport = New frmRepViewer
'   Set loreport.ReportSource = oReport
'   loreport.Show
'
'   PrintTrans = True
'
'endPoc:
'   Set loreport = Nothing
'   Set oReport = Nothing
'   Set lrs = Nothing
'   Set loRS = Nothing
'   Exit Function
'errProc:
'   PrintTrans = False
'   ShowError lsOldProc & "( " & " )"
'End Function
'
'Private Sub ClearFields()
'   txtField(0).Text = Format(Date, "MMM DD, YYYY")
'   txtField(1).Text = oTrans.Master("sCategrNm")
'   txtField(2).Text = oTrans.Master("sBrandNme")
'End Sub
'
'Private Sub ClearItem(ByVal lnRow As Integer)
'   With GridEditor1
'      If lnRow < 0 Then
'         .Rows = 2
'         .ColWidth(2) = 3100
'         lnRow = .Rows - 1
'      End If
'
'      'empty row
'      .TextMatrix(lnRow, 1) = ""
'      .TextMatrix(lnRow, 2) = ""
'      .TextMatrix(lnRow, 3) = "0.00"
'      .TextMatrix(lnRow, 4) = "0.00"
'      .TextMatrix(lnRow, 5) = "0.00"
'      .TextMatrix(lnRow, 6) = "0.00"
'      .TextMatrix(lnRow, 7) = "0.00"
'   End With
'
'   pbSave = False
'End Sub
'
'Private Sub LoadItem(ByVal lnRow As Integer)
'   Dim lnStart As Integer
'   Dim lnCtr As Integer
'
'   With GridEditor1
'      If lnRow < 0 Then
'         Debug.Print "loading All Items"
'         .Rows = oTrans.ItemCount + 1
'         If .Rows > 20 Then
'            .ColWidth(2) = 2850
'         Else
'            .ColWidth(2) = 3100
'         End If
'         lnStart = 1
'         lnRow = .Rows - 1
'      Else
'         lnStart = lnRow
'      End If
'
'      For lnCtr = lnStart To lnRow
'         'empty row
'         .TextMatrix(lnCtr, 1) = oTrans.Detail(lnCtr - 1, 1)
'         .TextMatrix(lnCtr, 2) = oTrans.Detail(lnCtr - 1, 2)
'         .TextMatrix(lnCtr, 3) = Format(oTrans.Detail(lnCtr - 1, 3), "#,##0.00")
'         .TextMatrix(lnCtr, 4) = Format(oTrans.Detail(lnCtr - 1, 4), "#,##0.00")
'         .TextMatrix(lnCtr, 5) = Format(oTrans.Detail(lnCtr - 1, 5), "#,##0.00")
'         .TextMatrix(lnCtr, 6) = Format(oTrans.Detail(lnCtr - 1, 6), "#,##0.00")
'         .TextMatrix(lnCtr, 7) = Format(oTrans.Detail(lnCtr - 1, 7), "#,##0.00")
'      Next
'   End With
'
'   pbSave = False
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
'      Select Case Index
'      Case 0
'         If Not IsDate(.Text) Then .Text = oApp.ServerDate
'         .Text = Format(.Text, "MMM DD, YYYY")
'      Case 1, 2
'         .Text = .Text
'      End Select
'
'      oTrans.Master(Index) = .Text
'   End With
'End Sub
'
'Private Function isEntryOk() As Boolean
'   If txtField(1).Text = "" Then
'      MsgBox "Invalid Category Detected!!!" & vbCrLf & _
'             "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
'      txtField(1).SetFocus
'      GoTo EntryNotOK
'   End If
'
'   With GridEditor1
'      If Trim(.TextMatrix(1, 1)) = "" Then
'         MsgBox "Detail is required!!!" & vbCrLf & _
'                "Please Verify your Entry then Try Again", vbCritical, "Warning"
'         .SetFocus
'         .Row = 1
'         .Col = 1
'         GoTo EntryNotOK
'      End If
'   End With
'
'EntryOK:
'   isEntryOk = True
'   Exit Function
'EntryNotOK:
'   isEntryOk = False
'End Function
'
'Private Function ItemExist(lsValue As String, lnRow As Integer, lnCol As Integer) As Boolean
'   Dim lnCtr As Integer
'
'   If Trim(lsValue) = "" Then Exit Function
'   If Trim(lsValue) = Trim(oTrans.Detail(lnRow - 1, lnCol)) Then Exit Function
'
'   With GridEditor1
'      For lnCtr = 1 To .Rows - 1
'         If .TextMatrix(lnCtr, 1) = lsValue And lnCtr <> lnRow Then
'            ItemExist = True
'            Debug.Print .Rows
'            Exit Function
'         End If
'      Next
'            Debug.Print .Rows
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
