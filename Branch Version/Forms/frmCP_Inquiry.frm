VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0"; "xrGridControl.ocx"
Begin VB.Form frmCP_Inquiry 
   BorderStyle     =   0  'None
   Caption         =   "Parts Inquiry"
   ClientHeight    =   6900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9315
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6900
   ScaleWidth      =   9315
   ShowInTaskbar   =   0   'False
   Begin xrGridEditor.GridEditor GridEditor1 
      Height          =   2895
      Left            =   210
      TabIndex        =   18
      Top             =   3285
      Width           =   7380
      _ExtentX        =   13018
      _ExtentY        =   5106
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
      Object.HEIGHT          =   2895
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
      MOUSEICON       =   "frmCP_Inquiry.frx":0000
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
      Height          =   6240
      Index           =   0
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   7620
      _ExtentX        =   13441
      _ExtentY        =   11007
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   4650
         TabIndex        =   13
         Top             =   1500
         Width           =   2820
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   7
         Left            =   1035
         TabIndex        =   15
         Top             =   1800
         Width           =   6435
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   1035
         TabIndex        =   3
         Top             =   600
         Width           =   6435
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   4650
         TabIndex        =   11
         Top             =   1200
         Width           =   2820
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   1035
         TabIndex        =   9
         Top             =   1500
         Width           =   2820
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   1035
         TabIndex        =   7
         Top             =   1200
         Width           =   2820
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   8
         Left            =   5610
         TabIndex        =   17
         Tag             =   "ht0;ft0"
         Text            =   "00,000.00"
         Top             =   2100
         Width           =   1860
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   1035
         TabIndex        =   5
         Top             =   900
         Width           =   6435
      End
      Begin VB.TextBox txtField 
         Alignment       =   2  'Center
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
         Index           =   9
         Left            =   2070
         TabIndex        =   20
         Tag             =   "ht0"
         Top             =   5730
         Width           =   4515
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
         Height          =   285
         Index           =   0
         Left            =   1035
         TabIndex        =   1
         Top             =   165
         Width           =   1920
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
         Height          =   195
         Index           =   7
         Left            =   375
         TabIndex        =   8
         Top             =   1545
         Width           =   540
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Brand"
         Height          =   195
         Index           =   6
         Left            =   390
         TabIndex        =   6
         Top             =   1260
         Width           =   525
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   930
         Width           =   795
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Size"
         Height          =   195
         Index           =   9
         Left            =   4020
         TabIndex        =   12
         Top             =   1545
         Width           =   300
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Color"
         Height          =   195
         Index           =   8
         Left            =   4020
         TabIndex        =   10
         Top             =   1245
         Width           =   495
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   4
         Left            =   4650
         TabIndex        =   16
         Top             =   2175
         Width           =   930
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BarrCode"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   2
         Top             =   630
         Width           =   660
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
         Height          =   195
         Index           =   5
         Left            =   1020
         TabIndex        =   19
         Top             =   5760
         Width           =   795
      End
      Begin VB.Shape Shape2 
         Height          =   465
         Index           =   0
         Left            =   120
         Top             =   5640
         Width           =   7350
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FFFFFF&
         Height          =   465
         Index           =   1
         Left            =   120
         Top             =   5655
         Width           =   7350
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   195
         Index           =   1
         Left            =   285
         TabIndex        =   14
         Top             =   1815
         Width           =   630
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1110
         Tag             =   "et0;ht2"
         Top             =   255
         Width           =   1935
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
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
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   210
         Width           =   915
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   7965
      TabIndex        =   26
      Top             =   3420
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "Cl&ose"
      AccessKey       =   "o"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCP_Inquiry.frx":001C
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   7965
      TabIndex        =   23
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
      Picture         =   "frmCP_Inquiry.frx":0796
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   7965
      TabIndex        =   22
      Top             =   2790
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
      Picture         =   "frmCP_Inquiry.frx":0F10
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   7965
      TabIndex        =   24
      Top             =   2160
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
      Picture         =   "frmCP_Inquiry.frx":168A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   7965
      TabIndex        =   21
      Top             =   2160
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
      Picture         =   "frmCP_Inquiry.frx":1E04
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   7965
      TabIndex        =   25
      Top             =   2790
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
      Picture         =   "frmCP_Inquiry.frx":257E
   End
End
Attribute VB_Name = "frmCP_Inquiry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmCP_Inquiry"

Private WithEvents oTrans As clsStockInquiry
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin

Dim pbGridFocus As Boolean
Dim pnCtr As Integer, pnIndex As Integer

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   Dim lnRep As Integer

   lsOldProc = "cmdButton_Click"
   'On Error GoTo errProc

   txtField_LostFocus pnIndex
   With GridEditor1
      Select Case Index
      Case 0
         If oTrans.SearchTransaction() Then
            LoadMaster
            LoadDetail
         Else
            If txtField(0).Text = "" Then InitEntry
         End If
         .Refresh
      Case 1
         oTrans.NewTransaction
         InitButton xeModeAddNew

         InitEntry
         txtField(1).SetFocus
      Case 2
         Unload Me
      Case 3
         If oTrans.SaveTransaction Then
            MsgBox "Transaction Updated Successfully!!!", vbInformation, "Confirm"
            If oTrans.OpenTransaction(oTrans.Master("sTransNox")) Then
               LoadMaster
               LoadDetail
            End If
            InitButton xeModeReady
            .Refresh
         Else
            MsgBox "Unable to Update Transaction!!!", vbCritical, "Warning"
         End If
      Case 4
         If Not pbGridFocus Then
            oTrans.SearchMaster pnIndex
         End If
      Case 5
         lnRep = MsgBox("Transaction is in Update Mode!!!" & vbCrLf & _
                        "Do you want to Cancel Transaction!!!", vbYesNo + vbQuestion, "Confirm")

         If lnRep = vbYes Then
            If oTrans.OpenTransaction(oTrans.Master("sTransNox")) Then
               LoadMaster
               LoadDetail
            Else
               oTrans.InitTransaction
               InitEntry
            End If
            InitButton xeModeReady
         End If
         .Refresh
      End Select
   End With

endProc:

   Exit Sub
errProc:
   ShowError lsOldProc, True
End Sub

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0

   With GridEditor1
      .Refresh
   End With
End Sub

Private Sub Form_Load()
   CenterChildForm mdiMain, Me

   Set oTrans = New clsStockInquiry
   Set oTrans.AppDriver = oApp
   oTrans.InitTransaction
   oTrans.NewTransaction

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransMaintenance

   InitGrid
   InitEntry
   InitButton xeModeAddNew

   txtField(9).Text = oApp.UserName
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
   Set oTrans = Nothing
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
   Case vbKeyF12
      oTrans.ViewModify
   End Select
End Sub

Private Sub oTrans_DetailRetrieved(ByVal Index As Integer)
   With GridEditor1
      .TextMatrix(.Row, Index) = oTrans.Detail(.Row - 1, Index)
   End With
End Sub

Private Sub oTrans_MasterRetrieved(ByVal Index As Integer)
   If Index = 8 Then
      txtField(Index).Text = Format(oTrans.Master(Index), "#,##0.00")
   Else
      txtField(Index).Text = oTrans.Master(Index)
   End If
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = oApp.getColor("HT1")
   End With

   pbGridFocus = False
   pnIndex = Index
End Sub

Private Sub InitEntry()
   txtField(0).Text = Format(oTrans.Master("sTransNox"), "@@@@-@@@@@@")
   txtField(1).Text = ""
   txtField(2).Text = ""
   txtField(3).Text = ""
   txtField(4).Text = ""
   txtField(5).Text = ""
   txtField(6).Text = ""
   txtField(7).Text = ""
   txtField(8).Text = "0.00"

   With GridEditor1
      .Rows = 2

      .TextMatrix(1, 1) = ""
      .TextMatrix(1, 2) = ""
      .TextMatrix(1, 3) = ""
      .TextMatrix(1, 4) = 0
   End With
End Sub

Private Sub InitButton(lnStat As Integer)
   Dim lbShow As Boolean
   Dim lnCtr As Integer

   lbShow = IIf(lnStat = 0, False, True)
   cmdButton(0).Visible = Not lbShow
   cmdButton(1).Visible = Not lbShow
   cmdButton(2).Visible = Not lbShow

   cmdButton(3).Visible = lbShow
   cmdButton(4).Visible = lbShow
   cmdButton(5).Visible = lbShow
   xrFrame1(0).Enabled = lbShow
End Sub

Private Sub InitGrid()
   Dim lnCtr As Integer

   With GridEditor1
      .Cols = 4
      .Rows = 2
      .Font = "MS Sans Serif"

      'column title
      .TextMatrix(0, 1) = "Branch"
      .TextMatrix(0, 2) = "Last Received"
      .TextMatrix(0, 3) = "QOH"
      .Row = 0

      'column alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = 3
         .ColEnabled(lnCtr) = False
      Next

      'column width
      .ColWidth(0) = 350
      .ColWidth(1) = 4600
      .ColWidth(2) = 1500
      .ColWidth(3) = 850

      .ColAlignment(1) = 1
      .ColAlignment(2) = 3
      .ColAlignment(3) = 3

      .EditorBackColor = oApp.getColor("HT1")

      .WordWrap = True

      .Row = 1
      .Col = 1
   End With
End Sub

Private Sub LoadMaster()
   txtField(0).Text = Format(oTrans.Master("sTransNox"), "@@@@-@@@@@@")
   txtField(1).Text = oTrans.Master("sBarrCode")
   txtField(2).Text = oTrans.Master("sDescript")
   txtField(3).Text = IFNull(oTrans.Master("sModelNme"), "")
   txtField(4).Text = IFNull(oTrans.Master("sColorNme"), "")
   txtField(5).Text = IFNull(oTrans.Master("sBrandNme"), "")
   txtField(6).Text = IFNull(oTrans.Master("sSizeName"), "")
   txtField(7).Text = oTrans.Master("sRemarksx")
   txtField(8).Text = Format(oTrans.Master("nSelPrice"), "#,##0.00")
End Sub

Private Sub LoadDetail()
   Dim lnCtr As Integer

   With GridEditor1
      .Rows = IIf(oTrans.ItemCount = 0, 2, oTrans.ItemCount + 1)

      For pnCtr = 0 To oTrans.ItemCount - 1
         .TextMatrix(pnCtr + 1, 1) = oTrans.Detail(pnCtr, "sBranchNm")
         .TextMatrix(pnCtr + 1, 2) = IIf(oTrans.Detail(pnCtr, "dLastRcvd") = "1/1/1900", "", Format(oTrans.Detail(pnCtr, "dLastRcvd"), "MMM-DD-YYYY"))
         .TextMatrix(pnCtr + 1, 3) = oTrans.Detail(pnCtr, "nQtyOnHnd")
      Next
   End With
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String

   lsOldProc = "txtField_KeyDown"
   'On Error GoTo errProc

   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
      With txtField(Index)
         If Index = 1 Or Index = 2 Then
            If KeyCode = vbKeyF3 Then
               oTrans.SearchMaster Index, .Text
               If .Text <> "" Then SetNextFocus
            Else
               If .Text <> "" Then oTrans.SearchMaster Index, .Text
            End If
         End If
      End With
      KeyCode = 0
   End If

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index _
                       & ", " & KeyCode _
                       & ", " & Shift & " )", True
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   With txtField(Index)
      .Text = TitleCase(.Text)
      oTrans.Master(Index) = .Text
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
