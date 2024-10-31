VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmCPSerialLedger 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Serial Ledger"
   ClientHeight    =   7680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9060
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   9060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5265
      Left            =   120
      TabIndex        =   15
      Tag             =   "et0;eb0;et0;bc2"
      Top             =   2295
      Width           =   8550
      _ExtentX        =   15081
      _ExtentY        =   9287
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
      Height          =   1665
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   5325
      _ExtentX        =   9393
      _ExtentY        =   2937
      Enabled         =   0   'False
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   1080
         TabIndex        =   3
         Top             =   630
         Width           =   4125
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   1080
         TabIndex        =   5
         Top             =   930
         Width           =   4125
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
         Left            =   1065
         TabIndex        =   1
         Top             =   105
         Width           =   2130
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   1080
         TabIndex        =   7
         Top             =   1230
         Width           =   4125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Brand-Model"
         Height          =   195
         Index           =   2
         Left            =   105
         TabIndex        =   4
         Top             =   975
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Serial No"
         Height          =   195
         Index           =   1
         Left            =   105
         TabIndex        =   2
         Top             =   705
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Serial ID"
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
         Top             =   150
         Width           =   750
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   1140
         Tag             =   "et0;ht2"
         Top             =   195
         Width           =   2130
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Color"
         Height          =   195
         Index           =   10
         Left            =   105
         TabIndex        =   6
         Top             =   1245
         Width           =   360
      End
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   1665
      Left            =   5475
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   2937
      Begin VB.TextBox txtDateThru 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1020
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   555
         Width           =   1890
      End
      Begin VB.TextBox txtDateFrom 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1020
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   225
         Width           =   1890
      End
      Begin xrControl.xrButton cmdButton 
         Height          =   480
         Index           =   0
         Left            =   1710
         TabIndex        =   14
         Top             =   1065
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
         Picture         =   "frmCPSerialLedger.frx":0000
      End
      Begin xrControl.xrButton cmdButton 
         Height          =   480
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   1065
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
         Picture         =   "frmCPSerialLedger.frx":077A
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
         Top             =   45
         Width           =   510
      End
      Begin VB.Shape Shape1 
         Height          =   825
         Index           =   1
         Left            =   105
         Top             =   135
         Width           =   2880
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date Thru"
         Height          =   195
         Index           =   5
         Left            =   195
         TabIndex        =   11
         Top             =   585
         Width           =   825
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date From"
         Height          =   195
         Index           =   4
         Left            =   195
         TabIndex        =   9
         Top             =   300
         Width           =   810
      End
   End
End
Attribute VB_Name = "frmCPSerialLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmCPSerialLedger"

Private oSkin As clsFormSkin

Dim psSerialID As String

Property Let SerialID(SerialID As String)
   psSerialID = SerialID
End Property

Private Sub cmdButton_Click(Index As Integer)
   Select Case Index
   Case 0
      Unload Me
   Case 1
      Call loadSerialLedger
   End Select
End Sub

Private Sub Form_Activate()
   Dim lrs As Recordset
   
   Set lrs = New Recordset
   lrs.Open "SELECT" & _
               " dTransact" & _
            " FROM CP_Inventory_Serial_Ledger" & _
            " WHERE sSerialID = " & strParm(psSerialID) & _
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
   'On Error GoTo errProc

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.DisableClose = True
   oSkin.ApplySkin xeFormLedger
   
   With MSFlexGrid1
      .Cols = 8
      .Font = "MS Sans Serif"
      
      'column title
      .TextMatrix(0, 1) = "Date"
      .TextMatrix(0, 2) = "Source"
      .TextMatrix(0, 3) = "Source No."
      .TextMatrix(0, 4) = "Branch"
      .TextMatrix(0, 5) = "Destination/Supplier/Customer"
      
      'column alignment
      .Row = 0
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = 1
      Next
      
      'column width
      .ColWidth(0) = 330
      .ColWidth(1) = 1200
      .ColWidth(2) = 2770
      .ColWidth(3) = 1140
      .ColWidth(4) = 3000
      .ColWidth(5) = 3000
      .ColWidth(6) = 0
      .ColWidth(7) = 0
      
      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      .ColAlignment(4) = 1
   End With
   
   Call clearGrid
   txtDateFrom = Format(DateAdd("m", -1, oApp.ServerDate), "MMMM DD, YYYY")
   txtDateThru = Format(oApp.ServerDate, "MMMM DD, YYYY")
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
End Sub

Private Sub MSFlexGrid1_GotFocus()
   MSFlexGrid1.BackColorSel = &HA4A36A
End Sub

Private Sub MSFlexGrid1_LostFocus()
   MSFlexGrid1.BackColorSel = &H8000000D
End Sub

Private Sub txtDateFrom_GotFocus()
   With txtDateFrom
      .Text = Format(.Text, "MMM/DD/YYYY")

      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = oApp.getColor("HT1")
   End With
End Sub

Private Sub txtDateFrom_LostFocus()
   With txtDateFrom
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtDateFrom_Validate(Cancel As Boolean)
   With txtDateFrom
      If Not IsDate(.Text) Then .Text = oApp.ServerDate
      .Text = Format(.Text, "MMMM DD, YYYY")
   End With
End Sub

Private Sub txtDateThru_GotFocus()
   With txtDateThru
      .Text = Format(.Text, "MM/DD/YYYY")

      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = oApp.getColor("HT1")
   End With
End Sub

Private Sub txtDateThru_LostFocus()
   With txtDateThru
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtDateThru_Validate(Cancel As Boolean)
   With txtDateThru
      If Not IsDate(.Text) Then .Text = oApp.ServerDate
      .Text = Format(.Text, "MMMM DD, YYYY")
   End With
End Sub

Private Sub loadSerialLedger()
   Dim lsProcName As String
   Dim lorsSource As Recordset
   Dim lorsTable As Recordset
   Dim lsSourceNo As String
   Dim lnCtr As Integer
   Dim lsSQL As String

   lsProcName = "loadSerialLedger"
   'On Error GoTo errProc
   
   Set lorsSource = New Recordset
   lorsSource.Open "SELECT" _
               & "  a.sSerialID" _
               & ", a.dTransact" _
               & ", b.sSourceNm" _
               & ", a.sSourceNo" _
               & ", c.sBranchNm" _
               & ", b.sSourceID" _
               & ", a.sSourceCd" _
            & " FROM CP_Inventory_Serial_Ledger a" _
               & ", xxxTransactionSource b" _
               & ", Branch c" _
            & " WHERE a.sSerialID = " & strParm(psSerialID) _
               & " AND a.sSourceCd = b.sSourceID" _
               & " AND a.sBranchCd = c.sBranchCd" _
               & " AND a.dTransact BETWEEN " & dateParm(txtDateFrom) _
                  & " AND " & dateParm(txtDateThru & " 23:59:59") _
            & " ORDER BY" _
               & "  a.dTransact" _
               & ", a.sSourceCd DESC" _
   , oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   
   If lorsSource.EOF Then
      Call clearGrid
      GoTo endProc
   End If
   
   With MSFlexGrid1
      .Rows = lorsSource.RecordCount + 1
      For lnCtr = 0 To lorsSource.RecordCount - 1
         .TextMatrix(lnCtr + 1, 0) = lnCtr + 1
         .TextMatrix(lnCtr + 1, 1) = IFNull(Format(lorsSource("dTransact"), "MMM-DD-YYYY"), "")
         .TextMatrix(lnCtr + 1, 2) = IFNull(lorsSource("sSourceNm"), "")
         .TextMatrix(lnCtr + 1, 3) = IFNull(Format(lorsSource("sSourceNo"), "@@@@-@@@@@@"), "")
         .TextMatrix(lnCtr + 1, 4) = IFNull(lorsSource("sBranchNm"), "")
         .TextMatrix(lnCtr + 1, 6) = IFNull(lorsSource("sSourceID"), "")
         .TextMatrix(lnCtr + 1, 7) = IFNull(lorsSource("sSourceNo"), "")
         
         lsSQL = ""
         Select Case Right(LCase(lorsSource("sSourceCd")), 2)
         Case "bt" ' Backload Trucking
         Case "da" ' Delivery Acceptance
            lsSQL = "SELECT" & _
                        "  a.sTransNox" & _
                        ", a.sReferNox xReferNox" & _
                        ", b.sCompnyNm xSourcexx" & _
                     " FROM CP_PO_Receiving_Master a" & _
                        ", Client_Master b" & _
                     " WHERE a.sSupplier= b.sClientID" & _
                        " AND a.sTransNox = " & strParm(lorsSource("sSourceNo"))
         Case "dl" ' Transfer Acceptance
            lsSQL = "SELECT" & _
                        "  a.sTransNox" & _
                        ", a.sReferNox xReferNox" & _
                        ", b.sBranchNm xSourcexx" & _
                     " FROM CP_Transfer_Master a" & _
                        ", Branch b" & _
                     " WHERE LEFT(a.sTransNox, " & Len(oApp.BranchCode) & ") = b.sBranchCd" & _
                        " AND a.sTransNox = " & strParm(lorsSource("sSourceNo"))
         Case "dv", "bb", "ab" ' CP Transfer
            lsSQL = "SELECT" & _
                        "  a.sTransNox" & _
                        ", a.sReferNox xReferNox" & _
                        ", b.sBranchNm xSourcexx" & _
                     " FROM CP_Transfer_Master a" & _
                        ", Branch b" & _
                     " WHERE a.sDestinat = b.sBranchCd" & _
                        " AND a.sTransNox = " & strParm(lorsSource("sSourceNo"))
         Case "pp" ' CP Purchase Replacement
            lsSQL = "SELECT" & _
                        "  RIGHT(a.sTransNox,8) AS xReferNox" & _
                        ", b.sCompnyNm xSourcexx" & _
                     " FROM CP_PO_Replacement_Master a" & _
                        ", Client_Master b" & _
                     " WHERE a.sClientID= b.sClientID" & _
                        " AND a.cTranStat <> " & strParm(xeStateCancelled) & _
                        " AND a.sTransNox = " & strParm(lorsSource("sSourceNo"))
         Case "pr" ' CP Purchase Return
            lsSQL = "SELECT" & _
                        "  a.sTransNox xRefernox" & _
                        ", b.sCompnyNm xSourcexx" & _
                     " FROM CP_PO_Return_Master a" & _
                        ", Client_Master b" & _
                     " WHERE a.sClientID = b.sClientID" & _
                        " AND a.cTranStat <> " & strParm(xeStateCancelled) & _
                        " AND a.sTransNox = " & strParm(lorsSource("sSourceNo"))
         Case "sc", "as" ' Service Center
            lsSQL = "SELECT" & _
                        "  a.sReferNox xRefernox" & _
                        ", b.sCompnyNm xSourcexx" & _
                     " FROM CP_JobOrder_Master a" & _
                        ", Client_Master b" & _
                     " WHERE a.sClientID = b.sClientID" & _
                        " AND a.cTranStat <> " & strParm(xeStateCancelled) & _
                        " AND a.sTransNox = " & strParm(lorsSource("sSourceNo"))
         Case "py" ' Payment to Supplier
            lsSQL = "SELECT" & _
                        "  RIGHT(a.sTransNox,8) AS xReferNox" & _
                        ", b.sCompnyNm xSourcexx" & _
                     " FROM Payment_Master a" & _
                        ", Client_Master b" & _
                     " WHERE a.sSupplier= b.sClientID" & _
                        " AND a.sTransNox = " & strParm(lorsSource("sSourceNo"))
         Case "sp" ' CP Sales Replacement
            lsSQL = "SELECT" & _
                        "  a.sTransNox xReferNox" & _
                        ", CONCAT(b.sLastName, ', ', b.sFrstName) AS xSourcexx" & _
                     " FROM CP_SO_Replacement_Master a" & _
                        ", Client_Master b" & _
                     " WHERE a.sClientID = b.sClientID" & _
                        " AND a.cTranStat <> " & strParm(xeStateCancelled) & _
                        " AND a.sTransNox = " & strParm(lorsSource("sSourceNo"))
         Case "jr" ' CP JobOrder Replacement
'            lsSQL = "SELECT" & _
'                        "  a.sReferNox xReferNox" & _
'                        ", CONCAT(b.sLastName, ', ', b.sFrstName) AS xSourcexx" & _
'                     " FROM CP_JobOrder_Master a" & _
'                        ", Client_Master b" & _
'                     " WHERE a.sClientID = b.sClientID" & _
'                        " AND a.cTranStat <> " & strParm(xeJOStateCancelled) & _
'                        " AND a.sTransNox = " & strParm(lorsSource("sSourceNo"))
         Case "sr" ' CP Sales Return
            lsSQL = "SELECT" & _
                        "  a.sTransNox xReferNox" & _
                        ", CONCAT(b.sLastName, ', ', b.sFrstName) AS xSourcexx" & _
                     " FROM CP_SO_Return_Master a" & _
                        ", Client_Master b" & _
                     " WHERE a.sClientID = b.sClientID" & _
                        " AND a.cTranStat <> " & strParm(xeStateCancelled) & _
                        " AND a.sTransNox = " & strParm(lorsSource("sSourceNo"))
         Case "sl", "ga" ' MC Sales
            lsSQL = "SELECT" & _
                        "  a.sTransNox " & _
                        ", a.sSalesInv xReferNox" & _
                        ", CONCAT(b.sLastName, ', ', b.sFrstName) AS xSourcexx" & _
                     " FROM CP_SO_Master a" & _
                        ", Client_Master b" & _
                     " WHERE a.sClientID = b.sClientID" & _
                        " AND a.cTranStat <> " & strParm(xeStateCancelled) & _
                        " AND a.sTransNox = " & strParm(lorsSource("sSourceNo"))
         Case "dm", "cm" 'Adjustment
            lsSQL = "SELECT" & _
                        "  sTransNox " & _
                        ", sDocNmbrx xReferNox" & _
                        ", sRemarksx xSourcexx" & _
                     " FROM CP_Inventory_Adjustment" & _
                     " WHERE cTranStat <> " & strParm(xeStateCancelled) & _
                        " AND sTransNox = " & strParm(lorsSource("sSourceNo"))
         Case "wl" ' Wholsale
            lsSQL = "SELECT" & _
                        "  a.sTransNox xRefernox" & _
                        ", CONCAT(b.sLastName, ', ', b.sFrstName) AS xSourcexx" & _
                     " FROM CP_WSO_Master a" & _
                        " Left Join Client_Master b" & _
                           " On a.sClientId = b.sClientId" & _
                     " WHERE a.cTranStat <> " & strParm(xeStateCancelled) & _
                        " AND a.sTransNox = " & strParm(lorsSource("sSourceNo"))
         Case "co" 'Charge Invoice
            lsSQL = "SELECT" & _
                        "  a.sTransNox " & _
                        ", a.sChrgeInv xReferNox" & _
                        ", CONCAT(b.sLastName, ', ', b.sFrstName) AS xSourcexx" & _
                     " FROM CP_CO_Master a" & _
                        ", Client_Master b" & _
                     " WHERE a.sClientID = b.sClientID" & _
                        " AND a.cTranStat <> " & strParm(xeStateCancelled) & _
                        " AND a.sTransNox = " & strParm(lorsSource("sSourceNo"))
         End Select
         
         Set lorsTable = New Recordset
         .TextMatrix(lnCtr + 1, 5) = ""
         lsSourceNo = "CP-" & Right(lorsSource("sSourceNo"), 10)
         Debug.Print lsSQL
         
         If lsSQL <> "" Then
            lorsTable.Open lsSQL, oApp.Connection, , , adCmdText
            
            If lorsTable.EOF = False Then
               .TextMatrix(lnCtr + 1, 3) = IIf(lorsTable("xReferNox") = "", lsSourceNo, lorsTable("xReferNox"))
               .TextMatrix(lnCtr + 1, 5) = lorsTable("xSourcexx")
            End If
         End If
         lorsSource.MoveNext
      Next
      
      Set lorsTable = Nothing
      .ColSel = .Cols - 1
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
      .TextMatrix(1, 4) = ""
      .TextMatrix(1, 5) = ""
      
      .Row = 1
      .Col = 1
      .ColSel = .Cols - 1
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
