VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmCP_Sales_Return_Maintenance 
   BorderStyle     =   0  'None
   Caption         =   "Sales Return"
   ClientHeight    =   7875
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11760
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7875
   ScaleWidth      =   11760
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame3 
      Height          =   510
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   10080
      _ExtentX        =   17780
      _ExtentY        =   900
      BackColor       =   14737632
      Begin VB.TextBox txtSearch 
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
         Index           =   2
         Left            =   720
         TabIndex        =   1
         Top             =   90
         Width           =   3330
      End
      Begin VB.TextBox txtSearch 
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
         Index           =   1
         Left            =   6660
         TabIndex        =   3
         Top             =   90
         Width           =   3330
      End
      Begin VB.TextBox txtSearch 
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
         Left            =   4800
         TabIndex        =   2
         Top             =   90
         Width           =   945
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Branch"
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
         Index           =   4
         Left            =   75
         TabIndex        =   26
         Top             =   135
         Width           =   720
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
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
         Index           =   1
         Left            =   5820
         TabIndex        =   24
         Top             =   150
         Width           =   810
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Refer#"
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
         Index           =   2
         Left            =   4125
         TabIndex        =   21
         Top             =   135
         Width           =   615
      End
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   3705
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   4020
      Width           =   10065
      _ExtentX        =   17754
      _ExtentY        =   6535
      BackColor       =   14737632
      Enabled         =   0   'False
      ClipControls    =   0   'False
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   3510
         Left            =   90
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   90
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   6191
         _Version        =   393216
         Appearance      =   0
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   2910
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   1080
      Width           =   10080
      _ExtentX        =   17780
      _ExtentY        =   5133
      BackColor       =   12632256
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   7830
         MaxLength       =   50
         TabIndex        =   22
         Top             =   1125
         Width           =   1965
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   1365
         TabIndex        =   12
         Top             =   660
         Width           =   4950
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
         Left            =   1365
         TabIndex        =   11
         Top             =   150
         Width           =   2310
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   7830
         MaxLength       =   50
         TabIndex        =   10
         Top             =   810
         Width           =   1965
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   570
         Index           =   3
         Left            =   1365
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   960
         Width           =   4950
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   870
         Index           =   5
         Left            =   1365
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   1845
         Width           =   4950
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   1365
         TabIndex        =   7
         Top             =   1545
         Width           =   4950
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6975
         TabIndex        =   25
         Top             =   255
         Width           =   2415
      End
      Begin VB.Shape Shape1 
         Height          =   330
         Left            =   6960
         Top             =   225
         Width           =   2460
      End
      Begin VB.Shape Shape3 
         Height          =   390
         Index           =   0
         Left            =   6930
         Top             =   195
         Width           =   2520
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Refer No"
         Height          =   285
         Index           =   0
         Left            =   6645
         TabIndex        =   23
         Top             =   1125
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   20
         Top             =   660
         Width           =   660
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
         Left            =   150
         TabIndex        =   19
         Top             =   195
         Width           =   1065
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transact. Date"
         Height          =   285
         Index           =   10
         Left            =   6630
         TabIndex        =   18
         Top             =   810
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Amount"
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
         Index           =   13
         Left            =   6585
         TabIndex        =   17
         Top             =   1800
         Width           =   2070
      End
      Begin VB.Label lblTrantotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Caption         =   "999,000.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   630
         Left            =   6585
         TabIndex        =   16
         Top             =   2085
         Width           =   3240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   195
         Index           =   11
         Left            =   -45
         TabIndex        =   15
         Top             =   1005
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   270
         Index           =   12
         Left            =   -45
         TabIndex        =   14
         Top             =   1875
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "*PIC"
         Height          =   285
         Index           =   5
         Left            =   0
         TabIndex        =   13
         Top             =   1545
         Width           =   1065
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   10425
      TabIndex        =   6
      Top             =   1785
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
      Picture         =   "frmCP_Sales_Return_Maintenance.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   10410
      TabIndex        =   4
      Top             =   525
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
      Picture         =   "frmCP_Sales_Return_Maintenance.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   10425
      TabIndex        =   5
      Top             =   1155
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
      Picture         =   "frmCP_Sales_Return_Maintenance.frx":0EF4
   End
End
Attribute VB_Name = "frmCP_Sales_Return_Maintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pxeMODULENAME = "frmCP_Sales_Return_Reg"
Private Const pxeAPPNAME = "Sales Return Reg"
Private WithEvents oTrans As clsCPSalesReturn
Attribute oTrans.VB_VarHelpID = -1

Private oSkin As clsFormSkin
Dim pbClosedTrans As Boolean
Dim pbLoad As Boolean
Dim psBranchCd As String
Dim psBranchNm As String


Private Sub cmdButton_Click(Index As Integer)
   Dim lnRep As Long
   
   Select Case Index
      Case 0 'browse
         If oTrans.SearchTransaction() Then
            LoadMaster
            LoadDetail
         End If
      Case 1 'close
         Unload Me
      Case 2 'Post
         If oTrans.Master("cTranstat") = xeStateClosed Then
            oTrans.postTransaction (oTrans.Master("sTransNox"))
            MsgBox "Transaction POST Successfully!", vbInformation, "INFO"
         ElseIf oTrans.Master("cTranstat") = xeStateOpen Then
            MsgBox "Unable to Post Transaction!!!" & vbCrLf & _
                "Please inform CP Dept to Print Transaction then Try Again!!!", vbCritical, "Warning"
         End If
         
   End Select
         
End Sub

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0
End Sub

Private Sub Form_Load()
Dim lsOldProc As String

   lsOldProc = "Form_Load"
    'On Error GoTo errProc

   CenterChildForm mdiMain, Me
   
   Set oTrans = New clsCPSalesReturn
   Set oTrans.AppDriver = oApp
      
   oTrans.InitTransaction
   
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransMaintenance
   oTrans.QueryMasterTable = "CP_CO_Master"
   oTrans.QueryDetailTable = "CP_CO_Detail"
    
   InitGrid
   initButton
   initForm
   clearFields
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
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

Private Sub LoadDetail()
   Dim lnRow As Integer
   Dim lnCol As Integer
   Dim lnSubTotal As Currency
   Dim loRS As Recordset
   Dim lsSQL As String
   
   lsSQL = "SELECT " & _
            " a.sTransNox " & _
            ", a.nQuantity " & _
            ", a.nUnitPrce " & _
            ", c.sSerialNo " & _
            ", b.sBarrCode" & _
            ", b.cHsSerial" & _
            ", b.sDescript " & _
            "FROM CP_SO_Return_Detail a" & _
               " LEFT JOIN CP_Inventory_Serial c" & _
                  " ON a.sSerialID = c.sSerialID" & _
            ", CP_Inventory b " & _
            " WHERE a.sTransNox = " & strParm(oTrans.Master(0)) & _
            " AND a.sStockIdx = b.sStockIdx"
      Set loRS = New Recordset
      loRS.Open lsSQL, oApp.Connection, adOpenDynamic, adLockReadOnly, adCmdText
      
    With MSFlexGrid1
      .Rows = loRS.RecordCount + 1
      For lnRow = 0 To loRS.RecordCount - 1
         For lnCol = 1 To .Cols - 1
            If lnCol = 1 Then 'imei
               .TextMatrix(lnRow + 1, lnCol) = IFNull(loRS("sSerialNo"), loRS("sBarrCode"))
            ElseIf lnCol = 2 Then 'desc
               .TextMatrix(lnRow + 1, lnCol) = loRS("sDescript")
            ElseIf lnCol = 3 Then 'qty
               .TextMatrix(lnRow + 1, lnCol) = loRS("nQuantity")
            ElseIf lnCol = 4 Then 'sel price
               .TextMatrix(lnRow + 1, lnCol) = Format(loRS("nUnitPrce"), "#,##0.00")
            ElseIf lnCol = 5 Then 'disc
               .TextMatrix(lnRow + 1, lnCol) = 0#
            ElseIf lnCol = 6 Then 'disc amt
               .TextMatrix(lnRow + 1, lnCol) = 0#
            ElseIf lnCol = 7 Then 'total
               .TextMatrix(lnRow + 1, lnCol) = Format(loRS("nQuantity") * loRS("nUnitPrce"), "#,##0.00")
            End If
         Next
      Next
   End With
End Sub

Private Sub InitGrid()
Dim lsOldProc As String
Dim lnCtr As Integer

lsOldProc = pxeMODULENAME & ".initGrid"
'On Error GoTo errProc
   
With MSFlexGrid1
   .Clear
   .Cols = 8
   .Rows = 2
      
   .TextMatrix(0, 0) = ""
   .TextMatrix(0, 1) = "IMEI/Barcode"
   .TextMatrix(0, 2) = "Description"
   .TextMatrix(0, 3) = "Qty."
   .TextMatrix(0, 4) = "Sel. Price"
   .TextMatrix(0, 5) = "Disc."
   .TextMatrix(0, 6) = "Dsc. Amt."
   .TextMatrix(0, 7) = "Total"
      
   .Row = 0
      
      'column alignment
   For lnCtr = 0 To .Cols - 1
      .Col = lnCtr
      .CellFontBold = True
      .CellAlignment = flexAlignCenterCenter
   Next
         
   .Row = 1
   .ColWidth(0) = "450"
   .ColWidth(1) = "1600"
   .ColWidth(2) = "2950"
   
   .Col = 0
   .ColSel = .Cols - 1
End With
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub LoadMaster()
   Dim pnCtr As Integer
   
   For pnCtr = 0 To txtField.Count - 1

      If pnCtr = 1 Then
         txtField(pnCtr) = Format(oTrans.Master(pnCtr), "MMMM DD, YYYY")
      ElseIf pnCtr = 5 Then
         txtField(pnCtr) = IFNull(oTrans.Master(5), "")
      ElseIf pnCtr = 6 Then
         txtField(pnCtr) = IFNull(oTrans.Master(15), "")
      Else
         txtField(pnCtr) = IFNull(oTrans.Master(pnCtr), "")
      End If
   Next
   
   txtSearch(0).text = oTrans.Master(6)
   txtSearch(1).text = oTrans.Master(2)
   
   lblTrantotal = Format(oTrans.Master("nTranTotl"), "#,##0.00")
   Label2.Caption = Format(TransStat(oTrans.Master("cTranStat")), ">")
     
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   
   With txtField(Index)
      If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
         Select Case Index
         Case 14
            If oTrans.SearchTransaction(.text) Then
               LoadMaster
               LoadDetail
            End If
         Case 15
            If oTrans.SearchTransaction(.text, False) Then
               LoadMaster
               LoadDetail
            End If
         End Select
      End If
   End With
End Sub

Private Sub initButton()
End Sub

Private Sub initForm()
Dim lsOldProc As String
   
lsOldProc = pxeMODULENAME & ".InitForm"
'On Error GoTo errProc


lblTrantotal = Format(0#, "#,##0.00")

   
Call initButton
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub clearFields()
   Dim pnCtr As Integer
   
   For pnCtr = 0 To txtField.Count - 1
      txtField(pnCtr).text = ""
   Next
End Sub

Private Sub SearchBranch(ByVal Search As Boolean, Optional Branch As Variant)
   Dim loRS As ADODB.Recordset
   Dim lsSelected() As String
   Dim lsSQL As String
   Dim lsBranchCd As String
   Dim lsBranchNm As String
   Dim lsOldProc As String

   lsOldProc = "SearchBranch"
   'On Error GoTo errProc

   lsSQL = "SELECT" _
               & "  sBranchCd" _
               & ", sBranchNm" _
            & " FROM Branch" _
            & " WHERE sBranchCd like 'C0%'" _
               & " AND cRecdStat = " & strParm(xeRecStateActive)

   If Not Search Then
      If Not IsMissing(Branch) Then lsSQL = lsSQL & " And sBranchNm LIKE " & strParm(Branch & "%")
   Else
      lsSQL = lsSQL & " And sBranchNm = " & strParm(Branch)
   End If

   Set loRS = New ADODB.Recordset

   loRS.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText

   If loRS.EOF Then
      psBranchCd = ""
      psBranchNm = ""
   ElseIf loRS.RecordCount = 1 Then
      psBranchCd = loRS("sBranchCd")
      psBranchNm = loRS("sBranchNm")
   Else
      lsSQL = KwikBrowse(oApp, loRS _
                           , "sBranchCd»sBranchNm" _
                           , "BranchCd»Branch Name" _
                          )

      If lsSQL <> "" Then
         lsSelected = Split(lsSQL, "»")
         psBranchCd = lsSelected(0)
         psBranchNm = lsSelected(1)
      End If
   End If

   Set loRS = Nothing

   clearFields
   txtSearch(2).text = psBranchNm
   txtSearch(2).Tag = txtField(0).text
   
   oTrans.Branch = Format(psBranchCd, "00")
   oTrans.InitTransaction
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " _
                       & "  " & IFNull(Search) _
                       & ", " & IFNull(Branch) _
                       & " )"
End Sub

Private Sub txtSearch_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Or KeyCode = vbKeyF3 Then
      With txtSearch(Index)
         Select Case Index
         Case 0, 1
            If oTrans.SearchTransaction(.text, IIf(Index = 0, True, False)) Then
               LoadMaster
               LoadDetail
               SetNextFocus
            Else
               If Index = 0 Then
                  clearFields
                  Exit Sub
               Else
               End If
            End If
         Case 2
            SearchBranch False, txtField(Index).text
            SetNextFocus
         End Select
      End With
   End If
End Sub
