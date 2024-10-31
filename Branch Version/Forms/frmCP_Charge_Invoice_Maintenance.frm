VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmCP_Charge_Invoice_Maintenance 
   BorderStyle     =   0  'None
   Caption         =   "Charge Invoice"
   ClientHeight    =   8295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11760
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8295
   ScaleWidth      =   11760
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame2 
      Height          =   3555
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   4620
      Width           =   10065
      _ExtentX        =   17754
      _ExtentY        =   6271
      BackColor       =   14737632
      ClipControls    =   0   'False
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   13
         Left            =   6705
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   2955
         Width           =   3240
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   2820
         Left            =   90
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   90
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   4974
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&AMOUNT PAID (F12)"
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
         Index           =   16
         Left            =   4515
         TabIndex        =   39
         Top             =   3075
         Width           =   2160
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   3480
      Index           =   0
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   1110
      Width           =   10080
      _ExtentX        =   17780
      _ExtentY        =   6138
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   7770
         TabIndex        =   25
         Top             =   1215
         Width           =   1995
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   7770
         TabIndex        =   27
         Top             =   1530
         Width           =   1995
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   9
         Left            =   7770
         TabIndex        =   31
         Top             =   2160
         Width           =   1995
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   12
         Left            =   7770
         TabIndex        =   37
         Top             =   3105
         Width           =   1995
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   10
         Left            =   7770
         TabIndex        =   33
         Top             =   2475
         Width           =   1995
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   11
         Left            =   7770
         TabIndex        =   35
         Top             =   2790
         Width           =   1995
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   7770
         TabIndex        =   29
         Top             =   1845
         Width           =   1995
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   1365
         MaxLength       =   50
         TabIndex        =   12
         Top             =   435
         Width           =   2310
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   915
         Index           =   7
         Left            =   1365
         MultiLine       =   -1  'True
         TabIndex        =   21
         Top             =   2265
         Width           =   4950
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   1365
         TabIndex        =   15
         Top             =   1005
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
         TabIndex        =   10
         Top             =   60
         Width           =   2310
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   600
         Index           =   4
         Left            =   1365
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1320
         Width           =   4950
      End
      Begin VB.CheckBox chkClientTp 
         Caption         =   "Company / Institution"
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
         Left            =   1365
         TabIndex        =   13
         Tag             =   "wt0;fb0"
         Top             =   780
         Width           =   2415
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   1365
         TabIndex        =   19
         Top             =   1950
         Width           =   4950
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice No."
         Height          =   195
         Index           =   2
         Left            =   6510
         TabIndex        =   24
         Top             =   1260
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Term"
         Height          =   195
         Index           =   14
         Left            =   6510
         TabIndex        =   26
         Top             =   1575
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Discount Rate"
         Height          =   195
         Index           =   0
         Left            =   6510
         TabIndex        =   30
         Top             =   2205
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Balance"
         Height          =   195
         Index           =   4
         Left            =   6510
         TabIndex        =   36
         Top             =   3150
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Discount Amt"
         Height          =   195
         Index           =   6
         Left            =   6510
         TabIndex        =   32
         Top             =   2520
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Limit"
         Height          =   195
         Index           =   7
         Left            =   6510
         TabIndex        =   34
         Top             =   2835
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Due Date"
         Height          =   195
         Index           =   10
         Left            =   6510
         TabIndex        =   28
         Top             =   1890
         Width           =   690
      End
      Begin VB.Shape Shape2 
         Height          =   330
         Left            =   7200
         Top             =   90
         Width           =   2535
      End
      Begin VB.Shape Shape3 
         Height          =   390
         Index           =   0
         Left            =   7170
         Top             =   60
         Width           =   2595
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
         Left            =   7290
         TabIndex        =   41
         Tag             =   "eb0;et0"
         Top             =   135
         Width           =   2385
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Name"
         Height          =   195
         Index           =   3
         Left            =   600
         TabIndex        =   14
         Top             =   1050
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transact. Date"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   11
         Top             =   480
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   195
         Index           =   12
         Left            =   435
         TabIndex        =   20
         Top             =   2625
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trans. No."
         Height          =   195
         Index           =   9
         Left            =   150
         TabIndex        =   9
         Top             =   105
         Width           =   750
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   285
         Left            =   1425
         Tag             =   "et0;ht2"
         Top             =   135
         Width           =   2325
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   195
         Index           =   11
         Left            =   495
         TabIndex        =   16
         Top             =   1365
         Width           =   570
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*PIC"
         Height          =   195
         Index           =   5
         Left            =   750
         TabIndex        =   18
         Top             =   1995
         Width           =   315
      End
      Begin VB.Label lblTrantotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Caption         =   "999,000.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   435
         Left            =   7170
         TabIndex        =   23
         Top             =   720
         Width           =   2595
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
         Left            =   7170
         TabIndex        =   22
         Top             =   465
         Width           =   2070
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   10410
      TabIndex        =   7
      Top             =   2505
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
      Picture         =   "frmCP_Charge_Invoice_Maintenance.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   10410
      TabIndex        =   6
      Top             =   615
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
      Picture         =   "frmCP_Charge_Invoice_Maintenance.frx":077A
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   525
      Index           =   1
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   10080
      _ExtentX        =   17780
      _ExtentY        =   926
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
         Index           =   16
         Left            =   930
         MaxLength       =   50
         TabIndex        =   1
         Top             =   105
         Width           =   3240
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
         Index           =   15
         Left            =   6720
         MaxLength       =   50
         TabIndex        =   3
         Top             =   105
         Width           =   3240
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
         Index           =   14
         Left            =   4740
         TabIndex        =   2
         Top             =   105
         Width           =   990
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
         Index           =   17
         Left            =   90
         TabIndex        =   42
         Top             =   135
         Width           =   780
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Custo&mer"
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
         Index           =   15
         Left            =   5850
         TabIndex        =   8
         Top             =   135
         Width           =   780
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Inv#"
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
         Left            =   4230
         TabIndex        =   0
         Top             =   135
         Width           =   540
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   10410
      TabIndex        =   5
      Top             =   1875
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "Pa&y"
      AccessKey       =   "y"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCP_Charge_Invoice_Maintenance.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   10410
      TabIndex        =   4
      Top             =   1245
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
      Picture         =   "frmCP_Charge_Invoice_Maintenance.frx":2326
   End
End
Attribute VB_Name = "frmCP_Charge_Invoice_Maintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pxeMODULENAME = "frmCP_Charge_Invoice_Reg"
Private Const pxeAPPNAME = "Charge Invoice History"
Private oTrans As ggcCPSales.clsCPChargeInvoice
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
   Case 1 'pay
      If oTrans.Master("cTranStat") = xeStatePosted Then
         If oTrans.PayTransaction(oTrans.Master("sTransNox")) Then
            MsgBox "Transaction Save Successfully!!!", vbInformation, "INFO"
            Call InitForm
         Else
            MsgBox "Unable to Pay Transaction!!!" & vbCrLf & _
                     "Please contact GGC SSG/SEG for assistance!!!", vbExclamation, "INFO"
         End If
      Else
         MsgBox "Unable to Pay Transaction!!!" & vbCrLf & _
                  "Transaction status must be either open or closed!!!", vbExclamation, "INFO"
      End If
   Case 2
      Unload Me
   Case 3
      If oTrans.Master("cTranstat") = xeStateClosed Then
         If oTrans.PostTransaction(oTrans.Master("sTransNox")) = True Then
            MsgBox "Transaction POST Successfully!", vbInformation, "INFO"
         Else
            MsgBox "Unable to Post Transaction!!!" & vbCrLf & _
             "Please inform CP Dept to Print Transaction then Try Again!!!", vbCritical, "Warning"
         End If
      ElseIf oTrans.Master("cTranstat") = xeStateOpen Then
         MsgBox "Unable to Post Transaction!!!" & vbCrLf & _
             "Please inform CP Dept to Print Transaction then Try Again!!!", vbCritical, "Warning"
      ElseIf oTrans.Master("cTranstat") = xeStateUnknown Then
         MsgBox "Transaction already paid!!!" & vbCrLf & _
             "Please inform CP Dept that the Transaction already paid!!!", vbCritical, "Warning"
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
    ''On Error GoTo errProc

   CenterChildForm mdiMain, Me
   
   Set oTrans = New ggcCPSales.clsCPChargeInvoice
   Set oTrans.AppDriver = oApp
      
   oTrans.InitTransaction
   
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualRight
   oTrans.HasSI = True
   
   InitGrid
   InitForm

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

Private Sub oTrans_MasterRetrieved(ByVal Index As Integer)
   oTrans.Master(Index) = txtField(Index)
End Sub

Private Sub LoadDetail()
   Dim lnRow As Integer
   Dim lnCol As Integer
   Dim lnSubTotal As Currency
   Dim lsSQL As String
   Dim lors As Recordset
   
   lsSQL = "SELECT " & _
            " a.sTransNox" & _
            ", b.sSerialNo" & _
            ", c.sBarrCode" & _
            ", c.sDescript" & _
            ", a.nQuantity" & _
            ", a.nUnitPrce" & _
            ", a.nDiscRate" & _
            ", a.nDiscAmtx" & _
            " FROM CP_CO_Detail a" & _
               " LEFT JOIN CP_Inventory_Serial b" & _
                  " ON a.sSerialID = b.sSerialID" & _
               " LEFT JOIN CP_Inventory c" & _
                  " ON a.sStockIDx = c.sStockIDx" & _
            " WHERE a.sTransNox = " & strParm(oTrans.Master("sTransNox"))
            
   Set lors = New Recordset
   lors.Open lsSQL, oApp.Connection, adOpenDynamic, adLockReadOnly, adCmdText
         
    With MSFlexGrid1
      .Rows = lors.RecordCount + 1
      For lnRow = 1 To lors.RecordCount
         For lnCol = 1 To .Cols - 1
            If lnCol = 1 Then 'imei
               .TextMatrix(lnRow, lnCol) = IFNull(lors("sSerialNo"), lors("sBarrCode"))
            ElseIf lnCol = 2 Then 'desc
               .TextMatrix(lnRow, lnCol) = lors("sDescript")
            ElseIf lnCol = 3 Then 'qty
               .TextMatrix(lnRow, lnCol) = lors("nQuantity")
            ElseIf lnCol = 4 Then 'sel price
               .TextMatrix(lnRow, lnCol) = Format(lors("nUnitPrce"), "#,##0.00")
            ElseIf lnCol = 5 Then 'disc
               .TextMatrix(lnRow, lnCol) = lors("nDiscRate")
            ElseIf lnCol = 6 Then 'disc amt
               .TextMatrix(lnRow, lnCol) = lors("nDiscAmtx")
            ElseIf lnCol = 7 Then 'total
               .TextMatrix(lnRow, lnCol) = Format(lors("nQuantity") * lors("nUnitPrce"), "#,##0.00")
            End If
         Next
         lors.MoveNext
      Next
   End With
End Sub

Private Sub InitGrid()
   Dim lsOldProc As String
   Dim lnCtr As Integer

   lsOldProc = pxeMODULENAME & ".initGrid"
   ''On Error GoTo errProc
   
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
      If pnCtr = 14 Then
         txtField(pnCtr) = IFNull(oTrans.Master(2), "")
      ElseIf pnCtr = 15 Then
         txtField(pnCtr) = IFNull(oTrans.Master(3), "")
      ElseIf pnCtr = 13 Then
         txtField(pnCtr) = Format(oTrans.Master(pnCtr), "#,##0.00")
      ElseIf pnCtr = 16 Then
         txtField(pnCtr) = txtField(16).Text
      ElseIf pnCtr = 9 Then
      ElseIf pnCtr = 10 Then
      ElseIf pnCtr = 11 Then
      Else
         txtField(pnCtr) = IFNull(oTrans.Master(pnCtr), "")
      End If
   Next
   
   lblTrantotal = Format(oTrans.Master("nTranTotl"), "#,##0.00")
   If oTrans.Master("cTranStat") = 4 Then
      Label2.Caption = "PAID"
   Else
      Label2.Caption = Format(TransStat(oTrans.Master("cTranStat")), ">")
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oTrans = Nothing
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   With txtField(Index)
     If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
         Select Case Index
         Case 14
            If oTrans.SearchTransaction(.Text, True) Then
               LoadMaster
               LoadDetail
               .SetFocus
            End If
         Case 15
            If oTrans.SearchTransaction(.Text, False) Then
               LoadMaster
               LoadDetail
               .SetFocus
            End If
         Case 16
            searchBranch False, .Text
            
            If oTrans.SearchTransaction(.Text, IIf(Index = 14, True, False)) Then
               LoadMaster
               LoadDetail
            Else
               If Index = 14 Then
                  Exit Sub
               Else
               End If
            End If
            txtField(14).SetFocus
         End Select
      End If
   End With
End Sub

Private Function PrinTrans()
   Dim lrs As Recordset
   Dim lors As Recordset
   Dim lnCtr As Integer
   Dim lsOldProc As String
   Dim lnTotalAmt As Currency
   Dim lsSQL As String

   lsOldProc = "PrinTrans"
   ''On Error GoTo errProc
   
   PrinTrans = True

      
   lsSQL = "SELECT " & _
            " a.sTransNox" & _
            ", b.sSerialNo" & _
            ", c.sBarrCode" & _
            ", c.sDescript" & _
            ", a.nQuantity" & _
            ", a.nUnitPrce" & _
            ", a.nDiscRate" & _
            ", a.nDiscAmtx" & _
            " FROM CP_CO_Detail a" & _
               " LEFT JOIN CP_Inventory_Serial b" & _
                  " ON a.sSerialID = b.sSerialID" & _
               " LEFT JOIN CP_Inventory c" & _
                  " ON a.sStockIDx = c.sStockIDx" & _
            " WHERE a.sTransNox = " & strParm(oTrans.Master("sTransNox"))
            
   Set lors = New Recordset
   lors.Open lsSQL, oApp.Connection, adOpenDynamic, adLockReadOnly, adCmdText
   
   Set lrs = New ADODB.Recordset

   lrs.Fields.Append "nField01", adInteger, 10
   lrs.Fields.Append "sField01", adVarChar, 200
   lrs.Fields.Append "sField02", adVarChar, 200
   lrs.Fields.Append "sField03", adVarChar, 200
   lrs.Fields.Append "sField04", adVarChar, 200
   lrs.Fields.Append "lField01", adCurrency
   lrs.Fields.Append "lField02", adCurrency
   lrs.Open

   With lors
      lnTotalAmt = 0
   
      For lnCtr = 0 To .RecordCount - 1
         lrs.AddNew
         lrs.Fields("nField01") = lors("nQuantity")
         lrs.Fields("sField01") = oTrans.Master(3)
         lrs.Fields("sField02") = oTrans.Master(4)
         lrs.Fields("sField03") = lors("sBarrCode") & " " & lors("sDescript") & " " & IFNull(lors("sSerialNo"), "")
         lrs.Fields("sField04") = oTrans.Master(2)
         lrs.Fields("lField01") = lors("nUnitPrce")
         lrs.Fields("lField02") = lors("nQuantity") * lors("nUnitPrce")

         lnTotalAmt = Format(lnTotalAmt + lors("nQuantity") * lors("nUnitPrce"), "#,##0.00")
         lors.MoveNext
      Next
      lrs.Sort = "sField01 DESC"
   End With

   'assign important info to the report
   
   Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\CP_CI.rpt")
   oReport.DiscardSavedData
   oReport.FieldMappingType = crAutoFieldMapping
   oReport.Database.SetDataSource lrs

   With oReport
      oReport.Sections("RF").ReportObjects("txtTotalAmt").SetText lnTotalAmt
   End With
   
   oReport.PrintOutEx False, 1
   lrs.Close
   PrinTrans = True

endProc:
   Set oReport = Nothing
   Set lrs = Nothing
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"

End Function

Private Sub InitForm()
   Dim lsOldProc As String
      
   lsOldProc = pxeMODULENAME & ".InitForm"
   ''On Error GoTo errProc
      
   txtField(0) = ""
   txtField(1) = ""
   txtField(2) = ""
   txtField(3) = ""
   txtField(4) = ""
   txtField(5) = ""
   txtField(6) = ""
   txtField(7) = ""
   txtField(8) = ""
   txtField(9) = Format(0#, "##0.00 %")
   txtField(10) = Format(0#, "#,##0.00")
   txtField(11) = Format(0#, "#,##0.00")
   txtField(12) = Format(0#, "#,##0.00")
   txtField(13) = Format(0#, "#,##0.00")
   
   lblTrantotal = Format(0#, "#,##0.00")
   chkClientTp.Value = 0
      
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub searchBranch(ByVal Search As Boolean, Optional Branch As Variant)
   Dim lors As ADODB.Recordset
   Dim lsSelected() As String
   Dim lsSQL As String
   Dim lsBranchCd As String
   Dim lsBranchNm As String
   Dim lsOldProc As String

   lsOldProc = "SearchBranch"
   ''On Error GoTo errProc

   lsSQL = "SELECT" _
               & "  sBranchCd" _
               & ", sBranchNm" _
            & " FROM Branch" _
            & " WHERE sBranchCd like '%'" _
               & " AND cRecdStat = " & strParm(xeRecStateActive)

   If Not Search Then
      If Not IsMissing(Branch) Then lsSQL = lsSQL & " And sBranchNm LIKE " & strParm(Branch & "%")
   Else
      lsSQL = lsSQL & " And sBranchNm = " & strParm(Branch)
   End If

   Set lors = New ADODB.Recordset

   lors.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText

   If lors.EOF Then
      psBranchCd = ""
      psBranchNm = ""
   ElseIf lors.RecordCount = 1 Then
      psBranchCd = lors("sBranchCd")
      psBranchNm = lors("sBranchNm")
   Else
      lsSQL = KwikBrowse(oApp, lors _
                           , "sBranchCd»sBranchNm" _
                           , "BranchCd»Branch Name" _
                          )

      If lsSQL <> "" Then
         lsSelected = Split(lsSQL, "»")
         psBranchCd = lsSelected(0)
         psBranchNm = lsSelected(1)
      End If
   End If

   Set lors = Nothing

   txtField(16).Text = psBranchNm
   txtField(16).Tag = txtField(0).Text
   
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
