VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmCP_MatrixLedger 
   BorderStyle     =   0  'None
   Caption         =   "Inventory Ledger"
   ClientHeight    =   7680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9060
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7680
   ScaleWidth      =   9060
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4995
      Left            =   90
      TabIndex        =   15
      Tag             =   "et0;eb0;et0;bc2"
      Top             =   2535
      Width           =   8595
      _ExtentX        =   15161
      _ExtentY        =   8811
      _Version        =   393216
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1890
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   585
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   3334
      Enabled         =   0   'False
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   1125
         MaxLength       =   50
         TabIndex        =   7
         Top             =   1335
         Width           =   4170
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   1125
         MaxLength       =   30
         TabIndex        =   5
         Top             =   1035
         Width           =   4170
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
         Height          =   285
         Index           =   0
         Left            =   1125
         TabIndex        =   1
         Top             =   165
         Width           =   2730
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   1125
         MaxLength       =   50
         TabIndex        =   3
         Top             =   735
         Width           =   4170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Barcode"
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
         Left            =   90
         TabIndex        =   0
         Top             =   210
         Width           =   720
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   1215
         Tag             =   "et0;ht2"
         Top             =   255
         Width           =   2730
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   10
         Left            =   90
         TabIndex        =   2
         Top             =   795
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
         Height          =   195
         Index           =   11
         Left            =   90
         TabIndex        =   6
         Top             =   1395
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Brand"
         Height          =   195
         Index           =   12
         Left            =   90
         TabIndex        =   4
         Top             =   1095
         Width           =   420
      End
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   1890
      Left            =   5520
      Tag             =   "wt0;fb0"
      Top             =   585
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   3334
      Begin VB.TextBox txtDateThru 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1020
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   630
         Width           =   1890
      End
      Begin VB.TextBox txtDateFrom 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1020
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   300
         Width           =   1890
      End
      Begin xrControl.xrButton cmdButton 
         Height          =   480
         Index           =   0
         Left            =   1710
         TabIndex        =   14
         Top             =   1290
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   847
         Caption         =   "&Ok"
         AccessKey       =   "O"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmCP_MatrixLedger.frx":0000
      End
      Begin xrControl.xrButton cmdButton 
         Height          =   480
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   1290
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   847
         Caption         =   "&Load Ledger"
         AccessKey       =   "L"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmCP_MatrixLedger.frx":077A
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Date"
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
         Left            =   135
         TabIndex        =   8
         Tag             =   "et0;fb0"
         Top             =   90
         Width           =   510
      End
      Begin VB.Shape Shape1 
         Height          =   1035
         Index           =   1
         Left            =   105
         Top             =   180
         Width           =   2880
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date Thru"
         Height          =   195
         Index           =   1
         Left            =   195
         TabIndex        =   11
         Top             =   660
         Width           =   825
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date From"
         Height          =   195
         Index           =   2
         Left            =   195
         TabIndex        =   9
         Top             =   375
         Width           =   810
      End
   End
End
Attribute VB_Name = "frmCP_MatrixLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmCP_MatrixLedger"

Private oSkin As clsFormSkin

Dim psStockIDx As String

Property Let StockID(sStockID As String)
   psStockIDx = sStockID
End Property

Private Sub cmdButton_Click(Index As Integer)
   Select Case Index
   Case 0
      Unload Me
   Case 1
      Call loadLedger
   End Select
End Sub

Private Sub Form_Activate()
   Dim lrs As Recordset
      
   Set lrs = New Recordset
   lrs.Open "SELECT" & _
               " dTransact" & _
            " FROM CP_Load_Matrix_Ledger" & _
            " WHERE sStockIDx = " & strParm(psStockIDx) & _
               " AND sBranchCd = " & strParm(oApp.BranchCode) & _
            " ORDER BY dTransact DESC" & _
            " LIMIT 18" _
   , oApp.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
   
   If lrs.EOF Then
      MsgBox "No Ledger Found!!!", vbCritical, "Warning"
      GoTo endProc
   End If
      
   txtDateThru.Text = lrs("dTransact")
   lrs.MoveLast
   txtDateFrom.Text = lrs("dTransact")
   Call cmdButton_Click(1)
   
endProc:
   Set lrs = Nothing
   Exit Sub
errProc:
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyReturn, vbKeyUp, vbKeyDown
      Select Case KeyCode
      Case vbKeyReturn, vbKeyDown
         SetNextFocus
      Case vbKeyUp
         SetPreviousFocus
      End Select
   End Select
End Sub

Private Sub Form_Load()
   Dim lnCtr As Integer
   Dim lsOldProc As String
   
   lsOldProc = "Form_Load"
   ''On Error GoTo errProc

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.DisableClose = True
   oSkin.ApplySkin xeFormLedger

   With MSFlexGrid1
      .Cols = 7
      .Font = "MS Sans Serif"
      
      'column title
      .TextMatrix(0, 1) = "Date"
      .TextMatrix(0, 2) = "Source"
      .TextMatrix(0, 3) = "Source No."
      .TextMatrix(0, 4) = "In"
      .TextMatrix(0, 5) = "Out"
      .TextMatrix(0, 6) = "Branch/Supplier/Customer"
      
      'column alignment
      .Row = 0
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = 1
      Next

      'column width
      .ColWidth(0) = 330
      .ColWidth(1) = 1300
      .ColWidth(2) = 2000
      .ColWidth(3) = 1200
      .ColWidth(4) = 500
      .ColWidth(5) = 500
      .ColWidth(6) = 3000
      
      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      .ColAlignment(6) = 1
   End With
   
   Call clearGrid
   txtDateFrom = Format(DateAdd("m", -1, oApp.ServerDate), "MMMM DD, YYYY")
   txtDateThru = Format(oApp.ServerDate, "MMMM DD, YYYY")

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub loadLedger()
   Dim lorsSource As Recordset
   Dim lorsTable As Recordset
   Dim lsProcName As String
   Dim lsSourceNo As String
   Dim lsSQL As String
   Dim lnCtr As Integer

   lsProcName = "loadLedger"
   ''On Error GoTo errProc

   Set lorsSource = New Recordset
   lsSQL = "SELECT" _
               & "  a.sBarrcode" _
               & ", a.sDescript" _
               & ", b.dTransact" _
               & ", c.sSourceNm" _
               & ", b.sSourceNo" _
               & ", b.nAmtInxxx" _
               & ", b.nAmtOutxx" _
               & ", b.sSourceCd" _
            & " FROM CP_Load_Matrix a" _
               & ", CP_Load_Matrix_Ledger b" _
               & ", xxxTransactionSource c" _
            & " WHERE a.sStockIDx = " & strParm(psStockIDx) _
               & " AND b.sBranchCd = " & strParm(oApp.BranchCode) _
               & " AND a.sStockIDx = b.sStockIDx" _
               & " AND b.sSourceCd = c.sSourceID" _
            & " ORDER BY" _
               & "  b.dTransact" _
               & ", b.nLedgerNo"
   
   lorsSource.Open lsSQL, oApp.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
   If lorsSource.EOF Then
      Call clearGrid
      GoTo endProc
   End If
    
   With MSFlexGrid1
      .Rows = lorsSource.RecordCount + 1
      For lnCtr = 0 To lorsSource.RecordCount - 1
         .TextMatrix(lnCtr + 1, 0) = lnCtr + 1
         .TextMatrix(lnCtr + 1, 1) = Format(lorsSource("dTransact"), "MMM-DD-YYYY")
         .TextMatrix(lnCtr + 1, 2) = IFNull(lorsSource("sSourceNm"), "")
         .TextMatrix(lnCtr + 1, 3) = IFNull(Format(lorsSource("sSourceNo"), "@@@@-@@@@@@"), "")
         .TextMatrix(lnCtr + 1, 4) = Format(lorsSource("nAmtInxxx"), "#,##0.00")
         .TextMatrix(lnCtr + 1, 5) = Format(lorsSource("nAmtOutxx"), "#,##0.00")
         
         lsSQL = ""
         Select Case Right(LCase(lorsSource("sSourceCd")), 2)
         Case "bt" ' Backload Trucking
         Case "da" ' Delivery Acceptance
            lsSQL = "SELECT" & _
                        "  a.sTransNox" & _
                        ", a.sSalesInv xReferNox" & _
                        ", b.sCompnyNm xSourcexx" & _
                     " FROM CP_Load_Receiving_Master a" & _
                        ", Client_Master b" & _
                     " WHERE a.sSupplier= b.sClientID" & _
                        " AND a.sTransNox = " & strParm(lorsSource("sSourceNo"))
         Case "dl" ' Transfer Acceptance
            lsSQL = "SELECT" & _
                        "  a.sTransNox" & _
                        ", a.sReferNox xReferNox" & _
                        ", b.sBranchNm xSourcexx" & _
                     " FROM CP_Load_Transfer_Master a" & _
                        ", Branch b" & _
                     " WHERE LEFT(a.sTransNox, " & Len(oApp.BranchCode) & ") = b.sBranchCd" & _
                        " AND a.sTransNox = " & strParm(lorsSource("sSourceNo"))
         Case "dv", "bb", "ab" ' CP Transfer
            lsSQL = "SELECT" & _
                        "  a.sTransNox" & _
                        ", a.sReferNox xReferNox" & _
                        ", b.sBranchNm xSourcexx" & _
                     " FROM CP_Load_Transfer_Master a" & _
                        ", Branch b" & _
                     " WHERE a.sDestinat = b.sBranchCd" & _
                        " AND a.sTransNox = " & strParm(lorsSource("sSourceNo"))
         Case "sl" ' Sales
            lsSQL = "SELECT" & _
                        "  sTransNox " & _
                        ", sReferNox xReferNox" & _
                        ", sPhoneNum AS xSourcexx" & _
                     " FROM CP_SO_Eload" & _
                     " WHERE cTranStat <> " & strParm(xeStateCancelled) & _
                        " AND sTransNox = " & strParm(lorsSource("sSourceNo"))
         Case "dm", "cm" 'Adjustment
            lsSQL = "SELECT" & _
                        "  sTransNox " & _
                        ", sDocNmbrx xReferNox" & _
                        ", sRemarksx xSourcexx" & _
                     " FROM CP_Load_Adjustment" & _
                     " WHERE cTranStat <> " & strParm(xeStateCancelled) & _
                        " AND sTransNox = " & strParm(lorsSource("sSourceNo"))
         End Select
         
         Set lorsTable = New Recordset
         .TextMatrix(lnCtr + 1, 6) = ""
         lsSourceNo = "CP-" & Right(lorsSource("sSourceNo"), 10)
         If lsSQL <> "" Then
            lorsTable.Open lsSQL, oApp.Connection, , , adCmdText
            If lorsTable.EOF = False Then
               .TextMatrix(lnCtr + 1, 3) = IIf(lorsTable("xReferNox") = "", lsSourceNo, lorsTable("xReferNox"))
               .TextMatrix(lnCtr + 1, 6) = lorsTable("xSourcexx")
            End If
         End If
         lorsSource.MoveNext
      Next
      
      Set lorsTable = Nothing
   End With
   
endProc:
   Set lorsSource = Nothing
   Exit Sub
errProc:
   ShowError lsProcName
End Sub

Private Sub clearGrid()
   With MSFlexGrid1
      .Rows = 2
      
      .TextMatrix(1, 0) = 1
      .TextMatrix(1, 1) = ""
      .TextMatrix(1, 2) = ""
      .TextMatrix(1, 3) = ""
      .TextMatrix(1, 4) = 0
      .TextMatrix(1, 5) = 0
      .TextMatrix(1, 6) = ""
      
      .Row = 1
      .Col = 1
      .ColSel = .Cols - 1
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
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
