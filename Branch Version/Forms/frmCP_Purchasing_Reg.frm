VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0"; "xrGridControl.ocx"
Begin VB.Form frmCP_Purchasing_Reg 
   BorderStyle     =   0  'None
   Caption         =   " Purchase Order"
   ClientHeight    =   7080
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11925
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7080
   ScaleWidth      =   11925
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   10575
      TabIndex        =   2
      Top             =   630
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
      Picture         =   "frmCP_Purchasing_Reg.frx":0000
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   2190
      Index           =   0
      Left            =   120
      Tag             =   "wt0;fb0"
      Top             =   1110
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   3863
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   9
         Left            =   8235
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   1320
         Width           =   1800
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   1335
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   75
         Width           =   2265
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   8220
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   615
         Width           =   1815
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   1320
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   570
         Width           =   5295
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   1320
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   915
         Width           =   5295
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   450
         Index           =   5
         Left            =   1320
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   1620
         Width           =   5310
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   6
         Left            =   8235
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   960
         Width           =   1800
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   4
         Left            =   1320
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   1260
         Width           =   5310
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Term"
         Height          =   195
         Index           =   8
         Left            =   6960
         TabIndex        =   25
         Top             =   1395
         Width           =   360
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   315
         Left            =   1350
         Tag             =   "et0;ht2"
         Top             =   165
         Width           =   2265
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction No"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   20
         Top             =   165
         Width           =   1095
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Date"
         Height          =   195
         Index           =   1
         Left            =   6960
         TabIndex        =   19
         Top             =   675
         Width           =   1230
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company Name"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   18
         Top             =   675
         Width           =   1125
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   17
         Top             =   1005
         Width           =   570
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   195
         Index           =   5
         Left            =   90
         TabIndex        =   16
         Top             =   1665
         Width           =   630
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Delivery"
         Height          =   195
         Index           =   6
         Left            =   6960
         TabIndex        =   15
         Top             =   1020
         Width           =   960
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Delivered To"
         Height          =   195
         Index           =   7
         Left            =   90
         TabIndex        =   14
         Top             =   1335
         Width           =   915
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "unknown"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   7575
         TabIndex        =   13
         Tag             =   "eb0;et0"
         Top             =   165
         Width           =   2385
      End
      Begin VB.Shape Shape3 
         Height          =   390
         Index           =   0
         Left            =   7515
         Top             =   105
         Width           =   2520
      End
      Begin VB.Shape Shape4 
         Height          =   330
         Index           =   0
         Left            =   7545
         Top             =   135
         Width           =   2460
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   8
      Left            =   10575
      TabIndex        =   4
      Top             =   1905
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
      Picture         =   "frmCP_Purchasing_Reg.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   7
      Left            =   10575
      TabIndex        =   5
      Top             =   2535
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
      Picture         =   "frmCP_Purchasing_Reg.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   6
      Left            =   10575
      TabIndex        =   3
      Top             =   1275
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
      Picture         =   "frmCP_Purchasing_Reg.frx":166E
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   540
      Index           =   1
      Left            =   135
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   10200
      _ExtentX        =   17992
      _ExtentY        =   953
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
         Index           =   7
         Left            =   1320
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   75
         Width           =   2145
      End
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
         Index           =   8
         Left            =   4995
         MaxLength       =   50
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   75
         Width           =   4965
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
         Height          =   285
         Index           =   9
         Left            =   75
         TabIndex        =   22
         Top             =   105
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Company Name"
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
         Index           =   8
         Left            =   3615
         TabIndex        =   21
         Top             =   105
         Width           =   1410
      End
   End
   Begin xrGridEditor.GridEditor GridEditor1 
      Height          =   3630
      Left            =   135
      TabIndex        =   23
      Tag             =   "et0;eb0;et0;bc2"
      Top             =   3315
      Width           =   10200
      _ExtentX        =   17992
      _ExtentY        =   6403
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
      Object.HEIGHT          =   3630
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
      MOUSEICON       =   "frmCP_Purchasing_Reg.frx":1DE8
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
Attribute VB_Name = "frmCP_Purchasing_Reg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmCP_Purchasing"

Private WithEvents oTrans As clsCPPurchasing
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin

Dim pnCtr As Integer
Dim pnIndex As Integer
Dim pbGridFocus As Boolean
Dim pbSave As Boolean
Dim pbEditMode As Boolean

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

Private Sub InitGrid()
   With GridEditor1
      .Rows = 2
      .Cols = 7
      .Font = "MS Sans Serif"

      'column title
      .TextMatrix(0, 1) = "BarrCode"
      .TextMatrix(0, 2) = "Description"
      .TextMatrix(0, 3) = "Brand"
      .TextMatrix(0, 4) = "Model"
      .TextMatrix(0, 5) = "Quantity"
      .TextMatrix(0, 6) = "Unit Prc"
      
      .Row = 0

      'column alignment
      For pnCtr = 0 To .Cols - 1
         .Col = pnCtr
         .CellFontBold = True
         .CellAlignment = 3
      Next

      'column width
      .ColWidth(0) = 400
      .ColWidth(1) = 1700
      .ColWidth(2) = 3000
      .ColWidth(3) = 1400
      .ColWidth(4) = 1500
      .ColWidth(5) = 850
      .ColWidth(6) = 1200
      
      .ColFormat(5) = 0#
      .ColFormat(6) = "0.00"
            
      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      .ColAlignment(4) = 6
      .ColAlignment(5) = 6
      
      .ColEnabled(3) = False
      .ColEnabled(4) = False
      
      .EditorBackColor = oApp.getColor("HT1")

      .Row = 1
      .Col = 1
   End With
End Sub

Public Function PrintTrans() As Boolean
   Dim lors As ADODB.Recordset
   Dim lrs As ADODB.Recordset
   Dim loData As ADODB.Recordset
   Dim lnCtr As Integer
   Dim lsOldProc As String

   lsOldProc = "PrintTrans"
   '''On Error GoTo errProc

   PrintTrans = True

   Set lors = New ADODB.Recordset

   lors.Fields.Append "nQuantity", adInteger, 5
   lors.Fields.Append "nUnitPrice", adCurrency
   lors.Fields.Append "sColor", adVarChar, 100
   lors.Fields.Append "sDescription", adVarChar, 100
   lors.Fields.Append "sBarrCode", adVarChar, 50
   lors.Open

   Set loData = New ADODB.Recordset
   loData.Open "SELECT " & _
                  " b.sBarrCode" & _
                  ", b.sDescript" & _
                  ", c.sModelNme" & _
                  ", d.sBrandNme" & _
                  ", e.sColorNme" & _
                  ", a.nQuantity" & _
                  ", a.nUnitPrce" & _
                  ", b.cHsSerial" & _
               " FROM CP_PO_Detail a" & _
               ", CP_Inventory b" & _
                     " LEFT JOIN CP_Model c" & _
                        " ON b.sModelIdx = c.sModelIDx" & _
                     " LEFT JOIN CP_Brand d" & _
                        " ON b.sBrandIDx = d.sBrandIDx" & _
                     " LEFT JOIN Color e" & _
                        " ON b.sColorIdx = e.sColorIDx" & _
               " WHERE a.sStockIDx = b.sStockIDx" & _
               " AND a.sTransNox = " & strParm(oTrans.Master("sTransNox")) _
               , oApp.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
   With GridEditor1
      For lnCtr = 1 To .Rows - 1
         lors.AddNew
            lors("nQuantity").Value = loData("nQuantity")
            lors("nUnitPrice").Value = loData("nUnitPrce")
            lors("sColor").Value = loData("sColorNme")
            lors("sDescription").Value = loData("sBrandNme") & "  " & loData("sModelNme")
            lors("sBarrcode").Value = loData("sBarrCode")
         loData.MoveNext
      Next
   End With

   ' assign important info to the report
   Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\Purchase.rpt")
   oReport.DiscardSavedData
   oReport.FieldMappingType = crAutoFieldMapping
   oReport.Database.SetDataSource lors

   Set lrs = New ADODB.Recordset
   lrs.Open "Select" _
               & "  CONCAT(b.sTownName, ', ', c.sProvName, ' ', b.sZippCode) xAddressx" _
            & " From Branch a" _
               & ", TownCity b" _
                  & " Left Join Province c" _
                     & " On b.sProvIDxx = c.sProvIDxx" _
            & " Where a.sTownIDxx = b.sTownIDxx" _
               & " And a.sBranchCd = " & strParm(oTrans.Master("sBranchCd")) _
            , oApp.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText

   If Not lrs.EOF Then oReport.Sections("PH").ReportObjects("txtDeliver").SetText txtField(4).Text
   oReport.Sections("PH").ReportObjects("txtSupplier").SetText txtField(2).Text
   oReport.Sections("PH").ReportObjects("txtDDate").SetText txtField(1).Text
   oReport.Sections("PH").ReportObjects("txtTransNo").SetText txtField(0).Text
   oReport.Sections("PH").ReportObjects("txtTerm").SetText txtField(9).Text
   oReport.Sections("PF").ReportObjects("txtUserRpt").SetText oApp.UserName

   oReport.PrintOutEx False, 1
   lors.Close
   lrs.Close

endProc:
   oTrans.CloseTransaction (oTrans.Master(0))
   Set oReport = Nothing
   Set lors = Nothing
   Set lrs = Nothing
   Exit Function
errProc:
   PrintTrans = False
   ShowError lsOldProc & "( " & " )"
End Function

Private Sub LoadMaster()

   For pnCtr = 0 To txtField.Count - 1
      Select Case pnCtr
      Case 0
         txtField(pnCtr).Text = Format(oTrans.Master(pnCtr), "@@@@-@@@@@@")
      Case 1
         txtField(pnCtr).Text = Format(oTrans.Master(pnCtr), "MMMM DD, YYYY")
      Case 2
         txtField(pnCtr).Text = oTrans.Master(2)
         txtField(pnCtr).Tag = txtField(pnCtr).Text
      Case 4
         txtField(pnCtr).Text = IFNull(oTrans.Master(8), "")
      Case 6
         txtField(pnCtr).Text = Format(oTrans.Master(11), "MMMM DD, YYYY")
      Case 7
         txtField(pnCtr).Text = Format(oTrans.Master(0), "@@@@-@@@@@@")
      Case 8
         txtField(pnCtr).Text = oTrans.Master(2)
      Case 9
        txtField(pnCtr).Text = IFNull(oTrans.Master(12), "")
      Case Else
         txtField(pnCtr).Text = IIf(IsNull(oTrans.Master(pnCtr)), "", oTrans.Master(pnCtr))
      End Select
   Next

   Label2.Caption = Format(TransStat(oTrans.Master("cTranStat")), ">")
   pbSave = True
End Sub

Private Sub LoadDetail()
Dim lnCtr As Integer

   With GridEditor1
      .Rows = IIf(oTrans.ItemCount = 0, 2, oTrans.ItemCount + 1)

      For pnCtr = 0 To oTrans.ItemCount - 1
         For lnCtr = 1 To .Cols - 1
            If lnCtr = 5 Then
               .TextMatrix(pnCtr + 1, lnCtr) = oTrans.Detail(pnCtr, 6)
            ElseIf lnCtr = 6 Then
               .TextMatrix(pnCtr + 1, lnCtr) = oTrans.Detail(pnCtr, 7)
            Else
               .TextMatrix(pnCtr + 1, lnCtr) = oTrans.Detail(pnCtr, lnCtr)
            End If
         Next
      Next
   End With
End Sub

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   Dim lsRep As String

   lsOldProc = "cmdButton_Click"
   '''On Error GoTo errProc
   With GridEditor1
      Select Case Index
      Case 4 ' Browse
         If oTrans.SearchTransaction() Then
            LoadMaster
            LoadDetail
         End If
      Case 6 'Cancel
         Unload Me
      Case 7 ' code for close
         Unload Me
      Case 8 'Print
         If pbSave Then
            lsRep = MsgBox("Print Transaction!!!", vbYesNo + vbQuestion, "Confirm")
            If lsRep = vbYes Then
               If Not PrintTrans Then MsgBox "Unable to Print Transaction!!!", vbCritical, "Warning"
            Else
               MsgBox "Unable to Print Transaction!!!", vbCritical, "Warning"
            End If
         End If
      End Select
   End With
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Sub Form_Activate()
   GridEditor1.Refresh

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
   '''On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oTrans = New clsCPPurchasing
   Set oTrans.AppDriver = oApp

   oTrans.InitTransaction
   oTrans.NewTransaction

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualRight

   InitGrid
   ClearFields

   pbEditMode = True

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set oTrans = Nothing
   Set oSkin = Nothing
End Sub

Private Sub GridEditor1_AddingRow(Cancel As Boolean)
   With GridEditor1
      If .TextMatrix(.Row, 1) = "" Then
         Cancel = True
      ElseIf .TextMatrix(.Row, 6) = 0 Then
         Cancel = True
      End If
      If Not Cancel Then oTrans.addDetail

   End With
End Sub
Private Sub GridEditor1_EditorValidate(Cancel As Boolean)
   With GridEditor1
      oTrans.Detail(.Row - 1, .Col) = .TextMatrix(.Row, .Col)
   End With
End Sub

Private Sub GridEditor1_GotFocus()
   With GridEditor1
      .EditorBackColor = oApp.getColor("HT1")
   End With
   pbGridFocus = True
End Sub
Private Sub GridEditor1_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String

   lsOldProc = "GridEditor1_KeyDown"
   '''On Error GoTo errProc

   With GridEditor1
      If KeyCode = vbKeyF3 Then
         If oTrans.searchDetail(.Row - 1, .Col, .TextMatrix(.Row, .Col)) Then .Col = 6
         KeyCode = 0
      End If
   End With

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " _
                       & "  " & KeyCode _
                       & ", " & Shift _
                       & " )", True
End Sub

Private Sub GridEditor1_LostFocus()
   With GridEditor1
      .EditorBackColor = oApp.getColor("EB")
      oTrans.Detail(.Row - 1, .Col) = .TextMatrix(.Row, .Col)
   End With
End Sub


Private Sub ClearFields()
   For pnCtr = 0 To txtField.Count - 1
      Select Case pnCtr
      Case 1, 6
         txtField(pnCtr).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
      Case 9, 10, 11
      Case Else
         txtField(pnCtr).Text = ""
         txtField(pnCtr).Tag = ""
      End Select
   Next

   With GridEditor1
      .Rows = 2

      'empty row
      .TextMatrix(1, 1) = ""
      .TextMatrix(1, 2) = ""
      .TextMatrix(1, 3) = ""
   End With

   pbSave = False
End Sub

Private Sub oTrans_DetailRetrieved(ByVal Index As Integer, ByVal Value As Variant)
      With GridEditor1
      If Index = 5 Then
         .TextMatrix(.Row, Index) = oTrans.Detail(.Row, 6)
      Else
         .TextMatrix(.Row, Index) = Value
      End If
   End With
End Sub

Private Sub oTrans_MasterRetrieved(ByVal Index As Variant, ByVal Value As Variant)
   txtField(Index).Text = Value
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      Select Case Index
      Case 1
         .Text = Format(.Text, "MM/DD/YY")
      End Select

      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = oApp.getColor("HT1")
   End With

   pbGridFocus = False
   pnIndex = Index
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim lsOldProc As String

   lsOldProc = "txtField_KeyDown"
   '''On Error GoTo errProc

   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
      With txtField(Index)
         If KeyCode = vbKeyF3 Then
            oTrans.SearchTransaction .Text
            LoadMaster
            LoadDetail
            If .Text <> "" Then SetNextFocus
         Else
            If .Text <> "" Then oTrans.SearchTransaction .Text, False
               LoadMaster
               LoadDetail
         End If
      End With
      KeyCode = 0
   End If

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " _
                       & "  " & Index _
                       & ", " & KeyCode _
                       & ", " & Shift _
                       & " )", True
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   With txtField(Index)
      .Text = TitleCase(.Text)
      Select Case Index
      Case 1
         If Not IsDate(.Text) Then .Text = oApp.ServerDate
         .Text = Format(.Text, "MMMM DD, YYYY")
      Case 5
         .Text = Format(.Text, ">")
      Case 10
         If .Text = "" Then
            ClearFields
            Exit Sub
         End If

         If Trim(LCase(.Text)) <> Trim(LCase(.Tag)) Then
            If oTrans.SearchTransaction(.Text, IIf(Index = 9, True, False)) Then
               LoadMaster
               LoadDetail
            Else
               ClearFields
               .SetFocus
            End If
         End If
      End Select

      If Index < 9 Then oTrans.Master(Index) = .Text
   End With
End Sub
'

