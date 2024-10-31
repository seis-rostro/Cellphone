VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0"; "xrGridControl.ocx"
Begin VB.Form frmCloseDay2Day 
   BorderStyle     =   0  'None
   Caption         =   "Day-To-Day Transaction Closing"
   ClientHeight    =   7950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12975
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7950
   ScaleWidth      =   12975
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame2 
      Height          =   7260
      Left            =   7500
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   5370
      _ExtentX        =   9472
      _ExtentY        =   12806
      BackColor       =   12632256
      BorderStyle     =   1
      Begin xrGridEditor.GridEditor GridEditor1 
         Height          =   6405
         Left            =   45
         TabIndex        =   20
         Tag             =   "et0;eb0;et0;bc2"
         Top             =   45
         Width           =   5235
         _ExtentX        =   9234
         _ExtentY        =   11298
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
         Object.HEIGHT          =   6405
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
         MOUSEICON       =   "frmCloseDay2Day.frx":0000
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
      Height          =   2685
      Left            =   1590
      Tag             =   "wt0;fb0"
      Top             =   1410
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   4736
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   7
         Left            =   1005
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   2085
         Width           =   1830
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   6
         Left            =   3960
         TabIndex        =   17
         Top             =   1605
         Width           =   1830
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   5
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1275
         Width           =   1830
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   4
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1605
         Width           =   1830
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   3
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1275
         Width           =   1830
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   855
         Width           =   1830
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   75
         Width           =   4815
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   405
         Width           =   1830
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Balance"
         Height          =   195
         Index           =   7
         Left            =   360
         TabIndex        =   18
         Top             =   2100
         Width           =   585
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Deposit"
         Height          =   195
         Index           =   6
         Left            =   3360
         TabIndex        =   16
         Top             =   1620
         Width           =   540
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Def. Adv. Pmt"
         Height          =   195
         Index           =   5
         Left            =   2910
         TabIndex        =   14
         Top             =   1290
         Width           =   990
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Adv. Paymt"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   12
         Top             =   1620
         Width           =   810
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tot. Sales"
         Height          =   195
         Index           =   3
         Left            =   210
         TabIndex        =   10
         Top             =   1290
         Width           =   720
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Prev. Bal."
         Height          =   195
         Index           =   2
         Left            =   225
         TabIndex        =   8
         Top             =   870
         Width           =   690
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Branch"
         Height          =   195
         Index           =   0
         Left            =   405
         TabIndex        =   4
         Top             =   120
         Width           =   510
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   195
         Index           =   1
         Left            =   570
         TabIndex        =   6
         Top             =   420
         Width           =   345
      End
   End
   Begin xrControl.xrFrame xrFrame3 
      Height          =   840
      Left            =   1590
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1482
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   975
         TabIndex        =   3
         Top             =   405
         Width           =   1830
      End
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   975
         TabIndex        =   1
         Top             =   75
         Width           =   4815
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
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Branch"
         Height          =   195
         Index           =   14
         Left            =   405
         TabIndex        =   0
         Top             =   120
         Width           =   510
      End
   End
   Begin xrControl.xrFrame xrFrame4 
      Height          =   3705
      Left            =   1590
      Tag             =   "wt0;fb0"
      Top             =   4110
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   6535
      BackColor       =   12632256
      BorderStyle     =   1
      Begin xrGridEditor.GridEditor GridEditor2 
         Height          =   2835
         Left            =   30
         TabIndex        =   21
         Tag             =   "et0;eb0;et0;bc2"
         Top             =   60
         Width           =   5790
         _ExtentX        =   10213
         _ExtentY        =   5001
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
         Object.HEIGHT          =   2835
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
         MOUSEICON       =   "frmCloseDay2Day.frx":001C
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
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   105
      TabIndex        =   26
      Top             =   3975
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
      Picture         =   "frmCloseDay2Day.frx":0038
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   105
      TabIndex        =   22
      Top             =   885
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
      Picture         =   "frmCloseDay2Day.frx":07B2
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   105
      TabIndex        =   25
      Top             =   3360
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Post"
      AccessKey       =   "P"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCloseDay2Day.frx":0F2C
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   105
      TabIndex        =   24
      Top             =   2130
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Journal"
      AccessKey       =   "J"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCloseDay2Day.frx":16A6
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   105
      TabIndex        =   23
      Top             =   1500
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
      Picture         =   "frmCloseDay2Day.frx":1E20
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   7
      Left            =   105
      TabIndex        =   27
      Top             =   2745
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Process"
      AccessKey       =   "P"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCloseDay2Day.frx":259A
   End
End
Attribute VB_Name = "frmCloseDay2Day"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmCloseDay2Day"

Private oTrans As clsDay2DayTrans
Private oSkin As clsFormSkin

Dim pbLoadRecord As Boolean
Dim pbGrd1Focus As Boolean
Dim pbGrd2Focus As Boolean
Dim pbSrchFocus As Boolean
Dim pbTextFocus As Boolean
Dim pnIndex As Integer
Dim pnCtr As Integer

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   Dim lnRep As Integer
   Dim ldDate As Date
   Dim loForm As frmCloseD2DJE

   lsOldProc = "cmdButton_Click"
   ''On Error GoTo errProc

   txtField_LostFocus pnIndex
   With GridEditor1
      Select Case Index
      Case 0 ' Browse
         If oTrans.SearchTransaction("", False) Then
            LoadMaster
            LoadDetail
         End If
      Case 1 ' Post
         If Not pbLoadRecord Then Exit Sub
      Case 2 ' Journal
         If Not pbLoadRecord Then Exit Sub
         Set loForm = New frmCloseD2DJE
         oTrans.loadAdditionalJE
         Set loForm.Day2DayObject = oTrans
         loForm.Show 1
      Case 3 ' Close
         Unload Me
      Case 4 ' Search
         If pbSrchFocus Then
            Select Case pnIndex
            Case 0
               Call oTrans.getBranch(txtSearch(0), False, True)
            Case 1
               If oTrans.SearchTransaction(txtSearch(1), False) Then
                  LoadMaster
                  LoadDetail
               End If
            End Select
         End If
      Case 7
         If Not pbLoadRecord Then Exit Sub
         Call oTrans.ProcessDay2Day
         LoadDetail
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

   Set oTrans = New clsDay2DayTrans
   Set oTrans.AppDriver = oApp

   oTrans.TransStatus = 0
   oTrans.InitTransaction

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualLeft

   txtSearch(0) = oTrans.Branch

   InitGrid
   ClearFields
'   InitButton xeModeReady

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oTrans = Nothing
   Set oSkin = Nothing
End Sub

Private Sub GridEditor1_DblClick()
   With GridEditor1
      If .TextMatrix(.Row, 2) <> "" Then
         'Needs to reposition the record pointer before calling oTrans.LoadTAccount
         If oTrans.TBalance(.Row - 1, "sDescript") <> "" Then
            oTrans.LoadTAccount
            LoadTAccount
         End If
      End If
   End With
End Sub

Private Sub GridEditor1_GotFocus()
   With GridEditor1
      .EditorBackColor = oApp.getColor("HT1")
   End With

   pbGrd1Focus = True
   pbGrd2Focus = False
   pbSrchFocus = False
   pbTextFocus = False
End Sub

Private Sub GridEditor2_GotFocus()
   With GridEditor2
      .EditorBackColor = oApp.getColor("HT1")
   End With

   pbGrd2Focus = True
   pbGrd1Focus = False
   pbSrchFocus = False
   pbTextFocus = False
End Sub

Private Sub GridEditor1_LostFocus()
   With GridEditor1
      .EditorBackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub GridEditor2_LostFocus()
   With GridEditor2
      .EditorBackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub ClearFields()
   Dim loTxt As TextBox

   For Each loTxt In txtField
      pnCtr = loTxt.Index
      Select Case pnCtr
      Case 0
         txtField(pnCtr).Text = oTrans.Branch
      Case 1
         txtField(pnCtr).Text = Format(oApp.ServerDate, "YYYY/MM/DD")
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

   With GridEditor2
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
   pbGrd2Focus = False
   pbSrchFocus = False
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
   pbGrd2Focus = False
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

Private Sub LoadMaster()
   Dim loTxt As TextBox
   For Each loTxt In txtField
      pnCtr = loTxt.Index
      Select Case pnCtr
      Case 0
         txtField(pnCtr).Text = oTrans.Branch
      Case 1
         txtField(pnCtr).Text = Format(oTrans.Master(pnCtr), "YYYY/MM/DD")
      Case Else
         txtField(pnCtr).Text = Format(oTrans.Master(pnCtr), "#,##0.00")
      End Select
   Next

   pbLoadRecord = True
End Sub

Private Sub LoadDetail()
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

Private Sub LoadTAccount()
   With GridEditor2
      .Rows = IIf(oTrans.TAItemCount = 0, 2, oTrans.TAItemCount + 1)
      .ColWidth(2) = 3200
      If .Rows > 12 Then .ColWidth(2) = 3000

      For pnCtr = 1 To .Rows - 1
         .TextMatrix(pnCtr, 1) = IFNull(oTrans.TAccount(pnCtr - 1, "nDebitAmt"), "0.00")
         .TextMatrix(pnCtr, 2) = IFNull(oTrans.TAccount(pnCtr - 1, "sTransact"), "")
         .TextMatrix(pnCtr, 3) = IFNull(oTrans.TAccount(pnCtr - 1, "nCredtAmt"), "0.00")
      Next
   End With
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   With txtField(Index)
      .Text = TitleCase(.Text)
      Select Case Index
      Case 6
         If Not IsNumeric(.Text) Then .Text = 0#
         .Text = Format(.Text, "#,##0.00")
         oTrans.Master(Index) = .Text
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

   With GridEditor2
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
      .ColWidth(2) = 3200
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
         LoadMaster
         LoadDetail
      End If
   End Select
End Sub
