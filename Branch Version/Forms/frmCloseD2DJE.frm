VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0"; "xrGridControl.ocx"
Begin VB.Form frmCloseD2DJE 
   BorderStyle     =   0  'None
   Caption         =   "Day-To-Day Transaction Adjustments"
   ClientHeight    =   5370
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13005
   LinkTopic       =   "Form1"
   ScaleHeight     =   5370
   ScaleWidth      =   13005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin xrControl.xrFrame xrFrame3 
      Height          =   840
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1482
      BackColor       =   12632256
      Enabled         =   0   'False
      BorderStyle     =   1
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   975
         TabIndex        =   1
         Top             =   75
         Width           =   4815
      End
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   975
         TabIndex        =   0
         Top             =   405
         Width           =   1830
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Branch"
         Height          =   195
         Index           =   14
         Left            =   405
         TabIndex        =   3
         Top             =   120
         Width           =   510
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   195
         Index           =   15
         Left            =   570
         TabIndex        =   2
         Top             =   420
         Width           =   345
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   3870
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   1410
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   6826
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   705
         Index           =   12
         Left            =   975
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1245
         Width           =   4815
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   11
         Left            =   975
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   75
         Width           =   4815
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   975
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   405
         Width           =   1830
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   975
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   750
         Width           =   1830
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "All Rights Reserved"
         Height          =   225
         Left            =   4020
         TabIndex        =   15
         Top             =   3555
         Width           =   1755
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "(c) 2014 and Beyond."
         Height          =   225
         Left            =   4020
         TabIndex        =   14
         Top             =   3060
         Width           =   1755
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Guanzon Group"
         Height          =   225
         Left            =   4020
         TabIndex        =   13
         Top             =   3300
         Width           =   1755
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   195
         Index           =   1
         Left            =   285
         TabIndex        =   11
         Top             =   1260
         Width           =   630
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account"
         Height          =   195
         Index           =   0
         Left            =   315
         TabIndex        =   9
         Top             =   120
         Width           =   600
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Debit"
         Height          =   195
         Index           =   2
         Left            =   540
         TabIndex        =   8
         Top             =   420
         Width           =   375
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Credit"
         Height          =   195
         Index           =   7
         Left            =   510
         TabIndex        =   7
         Top             =   750
         Width           =   405
      End
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   4725
      Left            =   6030
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   5370
      _ExtentX        =   9472
      _ExtentY        =   8334
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   4
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   3900
         Width           =   1830
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   3435
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   3900
         Width           =   1830
      End
      Begin xrGridEditor.GridEditor GridEditor1 
         Height          =   3705
         Left            =   45
         TabIndex        =   12
         Tag             =   "et0;eb0;et0;bc2"
         Top             =   45
         Width           =   5235
         _ExtentX        =   9234
         _ExtentY        =   6535
         AllowBigSelection=   -1  'True
         AutoAdd         =   0   'False
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
         Object.HEIGHT          =   3705
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
         MOUSEICON       =   "frmCloseD2DJE.frx":0000
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
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1950
         TabIndex        =   18
         Top             =   3960
         Width           =   1395
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   11625
      TabIndex        =   19
      Top             =   1845
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
      Picture         =   "frmCloseD2DJE.frx":001C
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   11625
      TabIndex        =   20
      Top             =   2475
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
      Picture         =   "frmCloseD2DJE.frx":0796
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   11625
      TabIndex        =   21
      Top             =   1215
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
      Picture         =   "frmCloseD2DJE.frx":0F10
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   11625
      TabIndex        =   22
      Top             =   600
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
      Picture         =   "frmCloseD2DJE.frx":168A
   End
End
Attribute VB_Name = "frmCloseD2DJE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmCloseD2DJE"

Private oTrans As clsDay2DayTrans
Private oSkin As clsFormSkin

Dim pbLoadRecord As Boolean
Dim pbGrd1Focus As Boolean
Dim pbGrd2Focus As Boolean
Dim pbSrchFocus As Boolean
Dim pbTextFocus As Boolean
Dim pnIndex As Integer
Dim pnCtr As Integer

Property Set Day2DayObject(foD2D As clsDay2DayTrans)
   Set oTrans = foD2D
End Property

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   Dim lnRep As Integer
   Dim ldDate As Date

   lsOldProc = "cmdButton_Click"
   ''On Error GoTo errProc

   txtField_LostFocus pnIndex
   With GridEditor1
      Select Case Index
      Case 0 ' Save
         If oTrans.SaveJE Then
            MsgBox "Save successfully!", vbOKOnly, "Confirmation"
         Else
            MsgBox "Unable to save!", vbOKOnly, "Confirmation"
         End If
      Case 1 ' Search
         Call oTrans.SearchAccount(11, "")
      Case 2 ' Journal
         oTrans.DeleteJE
      Case 3 ' Close
         Unload Me
      End Select
   End With
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyReturn, vbKeyUp, vbKeyDown
      Select Case KeyCode
      Case vbKeyReturn, vbKeyDown
         If GetFocus = GridEditor1.hwnd Then Exit Sub
         SetNextFocus
      Case vbKeyUp
         SetPreviousFocus
      End Select
   End Select
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String

   lsOldProc = "Form_Load"
   ''On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualRight

   txtSearch(0) = oTrans.Branch
   txtSearch(1) = Format(oTrans.Master("dTransact"), "YYYY/MM/DD")
   InitGrid
   ClearFields

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oTrans = Nothing
   Set oSkin = Nothing
End Sub

Private Sub GridEditor1_GotFocus()
   With GridEditor1
      .EditorBackColor = oApp.getColor("HT1")
   End With

   pbGrd1Focus = True
   pbSrchFocus = False
   pbTextFocus = False
End Sub

Private Sub GridEditor1_LostFocus()
   With GridEditor1
      .EditorBackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub ClearFields()
   Dim loTxt As TextBox

   For Each loTxt In txtField
      pnCtr = loTxt.Index
      Select Case pnCtr
      Case 11, 12
         txtField(pnCtr).Text = ""
      Case Else
         txtField(pnCtr).Text = "0.00"
      End Select
   Next

   With GridEditor1
      .Rows = 2
      .TextMatrix(1, 0) = "0.00"
      .TextMatrix(1, 2) = ""
      .TextMatrix(1, 3) = "0.00"
   End With
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = oApp.getColor("HT1")
   End With

   pnIndex = Index
   pbTextFocus = True
   pbGrd1Focus = False
   pbSrchFocus = False
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If Index = 11 Then
      Select Case KeyCode
      Case vbKeyF3, vbKeyReturn
         Call oTrans.SearchAccount(Index, txtField(Index))
         txtField(Index) = oTrans.AdditionalJE(GridEditor1.Row - 1, Index)
      End Select
   End If
End Sub

Private Sub txtSearch_GotFocus(Index As Integer)
   With txtSearch(Index)
      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = oApp.getColor("HT1")
   End With

   pnIndex = Index
   pbSrchFocus = True
   pbGrd1Focus = False
   pbTextFocus = False
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtSearch_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Or KeyCode = vbKeyF3 Then
      Select Case Index
      Case 0
         Call oTrans.getBranch(txtSearch(0), False, True)
      Case 1
         If oTrans.SearchTransaction(txtSearch(1), False) Then
'            LoadMaster
            LoadDetail
         End If
      End Select
   End If
End Sub

Private Sub txtSearch_LostFocus(Index As Integer)
   With txtSearch(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Public Sub LoadDetail()
   With GridEditor1
      .Rows = IIf(oTrans.TBItemCount = 0, 2, oTrans.TBItemCount + 1)
      .ColWidth(2) = 2640
      If .Rows > 12 Then .ColWidth(2) = 2440

      For pnCtr = 1 To .Rows - 1
         If Not IsNull(oTrans.TBalance(0, 1)) Then
            .TextMatrix(pnCtr, 1) = IFNull(oTrans.TBalance(pnCtr - 1, "nDebitAmt"), "0.00")
            .TextMatrix(pnCtr, 2) = IFNull(oTrans.TBalance(pnCtr - 1, "sDescript"), "")
            .TextMatrix(pnCtr, 3) = IFNull(oTrans.TBalance(pnCtr - 1, "nCredtAmt"), "0.00")
         Else
            .TextMatrix(pnCtr, 1) = "0.00"
            .TextMatrix(pnCtr, 2) = ""
            .TextMatrix(pnCtr, 3) = "0.00"
         End If
      Next
   End With
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   With txtField(Index)
      .Text = TitleCase(.Text)
      Select Case Index
      Case 1, 2
         If Not IsNumeric(.Text) Then .Text = 0#
         .Text = Format(.Text, "#,##0.00")
      End Select

      oTrans.AdditionalJE(GridEditor1.Row - 1, Index) = .Text
      .Text = oTrans.AdditionalJE(GridEditor1.Row - 1, Index)
   End With

   With GridEditor1
      Select Case Index
      Case 11
         .TextMatrix(.Row, 2) = .Text
      Case 1
         .TextMatrix(.Row, 1) = .Text
      Case 2
         .TextMatrix(.Row, 3) = .Text
      End Select
   End With
End Sub

Private Sub InitGrid()
   With GridEditor1
      .Rows = 2
      .Cols = 4
      .Font = "MS Sans Serif"

      'column title
      .TextMatrix(0, 1) = "Debit"
      .TextMatrix(0, 2) = "Account"
      .TextMatrix(0, 3) = "Credit"

      .Row = 0

      'column alignment
      For pnCtr = 0 To .Cols - 1
         .Col = pnCtr
         .CellFontBold = True
         .CellAlignment = 1
      Next

      'column width
      .ColWidth(0) = 330
      .ColWidth(1) = 1100
      .ColWidth(2) = 2640
      .ColWidth(3) = 1100

      .ColEnabled(1) = False
      .ColEnabled(2) = False
      .ColEnabled(3) = False

      .ColDefault(1) = "0.00"
      .ColDefault(3) = "0.00"

      .ColFormat(0) = "#,##0"
      .ColFormat(1) = "#,##0.00"
      .ColFormat(3) = "#,##0.00"

      .ColAlignment(1) = flexAlignRightCenter
      .ColAlignment(3) = flexAlignRightCenter

      .EditorBackColor = oApp.getColor("HT1")

      .Row = 1
      .Col = 1
   End With

End Sub

Private Sub ShowError(ByVal lsProcName As String, Optional bEnd As Boolean = False)
   With oApp
      .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
      If bEnd Then
         .xShowError
         End
      Else
         With Err
            .Raise .Number, .Source, .Description
         End With
      End If
   End With
End Sub

Private Sub txtSearch_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
   Case 0
      Call oTrans.getBranch(txtSearch(0), False, False)
   Case 1
      If oTrans.SearchTransaction(txtSearch(1), True) Then
'         LoadMaster
         LoadDetail
      End If
   End Select
End Sub

