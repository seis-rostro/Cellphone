VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmCPInvStockReqReg 
   BorderStyle     =   0  'None
   Caption         =   "CP Stock Request"
   ClientHeight    =   7365
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13200
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7365
   ScaleWidth      =   13200
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5565
      Left            =   90
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1695
      Width           =   11520
      _ExtentX        =   20320
      _ExtentY        =   9816
      _Version        =   393216
      Appearance      =   0
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   11865
      TabIndex        =   10
      Top             =   2430
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
      Picture         =   "frmCPInvStockReqReg.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   11865
      TabIndex        =   8
      Top             =   1170
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
      Picture         =   "frmCPInvStockReqReg.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   11865
      TabIndex        =   9
      Top             =   1800
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
      Picture         =   "frmCPInvStockReqReg.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   11865
      TabIndex        =   7
      Top             =   540
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Status"
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
      Picture         =   "frmCPInvStockReqReg.frx":166E
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   570
      Index           =   1
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   1110
      Width           =   11520
      _ExtentX        =   20320
      _ExtentY        =   1005
      BackColor       =   12632256
      ClipControls    =   0   'False
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
         Index           =   0
         Left            =   1740
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "M001-12-000001"
         Top             =   90
         Width           =   1980
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   9135
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "Mmm dd, yyyy"
         Top             =   90
         Width           =   1830
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   8490
         TabIndex        =   4
         Top             =   165
         Width           =   390
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   165
         Width           =   1335
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   525
      Index           =   0
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   11520
      _ExtentX        =   20320
      _ExtentY        =   926
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
         Left            =   1740
         TabIndex        =   1
         Top             =   90
         Width           =   1980
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Transaction No."
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
         Index           =   9
         Left            =   120
         TabIndex        =   0
         Top             =   150
         Width           =   1380
      End
      Begin VB.Shape Shape3 
         Height          =   390
         Index           =   0
         Left            =   8490
         Top             =   45
         Width           =   2505
      End
      Begin VB.Shape Shape4 
         Height          =   330
         Index           =   0
         Left            =   8520
         Top             =   75
         Width           =   2445
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
         Left            =   8550
         TabIndex        =   2
         Tag             =   "eb0;et0"
         Top             =   120
         Width           =   2385
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H8000000F&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   285
         Index           =   0
         Left            =   8550
         Tag             =   "et0;et0"
         Top             =   105
         Width           =   2400
      End
   End
End
Attribute VB_Name = "frmCPInvStockReqReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pxeMODULENAME = "frmCPInvStockReqReg"

Private oSkin As clsFormSkin
Private WithEvents oTrans As clsCPStockOrder
Attribute oTrans.VB_VarHelpID = -1

Private pbLoaded As Boolean

Private Sub cmdButton_Click(Index As Integer)
   Dim loForm As frmTranStat
   
   Select Case Index
   Case 0 'browse
      If oTrans.SearchTransaction(txtSearch, False) Then
         loadTransaction
      Else
         ClearFields
      End If
   Case 1 'cancel
      If oTrans.CancelTransaction Then
         MsgBox "Transaction Cancelled Succesfuly.", vbInformation, "Notice"
      End If
   Case 2 'close
      Unload Me
   Case 3 ' order Status
      Set loForm = New frmTranStat
         
      With frmTranStat
         Set .AppDriver = oApp
         .Show vbModal
         
         If Not .Cancelled Then
            oTrans.TransStatus = .TranStatus
            oTrans.InitTransaction
            ClearFields
         End If
      End With
   End Select
End Sub

Private Sub MSFlexGrid1_Click()
   With MSFlexGrid1
      .Col = 0
      .ColSel = .Cols - 1
   End With
End Sub

Private Sub Form_Activate()
   Dim lsOldProc As String

   If pbLoaded Then Exit Sub
   
   lsOldProc = "Form_Activate"
   ''On Error GoTo errProc

   oApp.MenuName = Me.Tag
   Me.ZOrder 0

   If Not pbLoaded Then pbLoaded = True
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String
   lsOldProc = "Form_Load"

   ''On Error GoTo errProc

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualRight
   CenterChildForm mdiMain, Me

   Set oTrans = New clsCPStockOrder
   Set oTrans.AppDriver = oApp
   oTrans.Branch = oApp.BranchCode
   oTrans.TransStatus = 1023
   oTrans.InitTransaction

   ClearFields
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub ClearFields()
   txtField(0) = ""
   txtField(1) = ""
   Label2.Caption = ""
   txtSearch = ""
   
   Call InitGrid
   If pbLoaded Then txtSearch.SetFocus
End Sub

Private Function loadTransaction() As Boolean
   Dim lnCtr As Integer
   
   ClearFields
      
   With oTrans
      txtField(0) = Format(.Master("sTransNox"), "@@@@@@-@@@@@@")
      txtField(1) = strLongDate(.Master("dTransact"))
      Label2.Caption = TransStat(.Master("cTranStat"))
      txtSearch = txtField(0)
      
      If .ItemCount > 20 Then
         MSFlexGrid1.ColWidth(1) = 2250
      Else
         MSFlexGrid1.ColWidth(1) = 2500
      End If
      
      MSFlexGrid1.Rows = .ItemCount + 1
      For lnCtr = 0 To .ItemCount - 1
         MSFlexGrid1.TextMatrix(lnCtr + 1, 0) = lnCtr + 1
         MSFlexGrid1.TextMatrix(lnCtr + 1, 1) = .Detail(lnCtr, "sBarrCode")
         MSFlexGrid1.TextMatrix(lnCtr + 1, 2) = .Detail(lnCtr, "sDescript")
         MSFlexGrid1.TextMatrix(lnCtr + 1, 3) = .Detail(lnCtr, "sBrandNme")
         MSFlexGrid1.TextMatrix(lnCtr + 1, 4) = IFNull(.Detail(lnCtr, "sModelNme"))
         MSFlexGrid1.TextMatrix(lnCtr + 1, 5) = IFNull(.Detail(lnCtr, "sColorNme"))
         MSFlexGrid1.TextMatrix(lnCtr + 1, 6) = .Detail(lnCtr, "nQuantity")
      Next
   End With
   
End Function

Private Sub InitGrid()
   Dim lnCtr As Integer

   With MSFlexGrid1
      .Clear
      .Cols = 7
      .Rows = 2
      
      .Row = 0
      .TextMatrix(0, 0) = ""
      .TextMatrix(0, 1) = "Bar Code"
      .TextMatrix(0, 2) = "Description"
      .TextMatrix(0, 3) = "Brand"
      .TextMatrix(0, 4) = "Model"
      .TextMatrix(0, 5) = "Color"
      .TextMatrix(0, 6) = "Qty"
      
      .ColWidth(0) = 450
      .ColWidth(1) = 2500
      .ColWidth(2) = 3000
      .ColWidth(3) = 1500
      .ColWidth(4) = 1500
      .ColWidth(5) = 1500
      .ColWidth(6) = 1000
      
      
      'Column Alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = 3
      Next
      .Row = 1
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

Private Sub oTrans_MasterRetrieved(ByVal Index As Variant, ByVal Value As Variant)
   Select Case Index
   Case 6
      Label2.Caption = TransStat(oTrans.Master("cTranStat"))
   End Select
End Sub

Private Sub txtSearch_GotFocus()
   txtSearch = Replace(txtSearch, "-", "")
   Call HighlightOn(txtSearch)
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyF3, vbKeyReturn
      Call cmdButton_Click(0)
   End Select
   KeyCode = 0
End Sub

Private Sub txtSearch_LostFocus()
   Call HighlightOff(txtSearch)
End Sub

