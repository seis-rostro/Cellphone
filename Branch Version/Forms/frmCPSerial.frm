VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmCPSerial 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Serial"
   ClientHeight    =   7680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9450
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   9450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   285
      Left            =   135
      TabIndex        =   7
      Top             =   7260
      Visible         =   0   'False
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5265
      Left            =   120
      TabIndex        =   6
      Tag             =   "et0;eb0;et0;bc2"
      Top             =   2295
      Width           =   9150
      _ExtentX        =   16140
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
      Width           =   7125
      _ExtentX        =   12568
      _ExtentY        =   2937
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   4695
         TabIndex        =   13
         Top             =   945
         Width           =   2025
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   1080
         TabIndex        =   9
         Top             =   945
         Width           =   2880
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   1080
         TabIndex        =   8
         Top             =   1245
         Width           =   2880
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   4695
         TabIndex        =   4
         Top             =   1245
         Width           =   2025
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
         TabIndex        =   2
         Top             =   120
         Width           =   2895
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   1065
         TabIndex        =   3
         Top             =   555
         Width           =   5670
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
         Height          =   195
         Index           =   4
         Left            =   4005
         TabIndex        =   14
         Top             =   1290
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descript."
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
         Left            =   120
         TabIndex        =   12
         Top             =   585
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model Name"
         Height          =   195
         Index           =   2
         Left            =   105
         TabIndex        =   11
         Top             =   1290
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Brand Name"
         Height          =   195
         Index           =   1
         Left            =   105
         TabIndex        =   10
         Top             =   1020
         Width           =   885
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
         Left            =   120
         TabIndex        =   1
         Top             =   165
         Width           =   720
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1140
         Tag             =   "et0;ht2"
         Top             =   210
         Width           =   2895
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Color"
         Height          =   195
         Index           =   11
         Left            =   4005
         TabIndex        =   5
         Top             =   990
         Width           =   360
      End
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   1665
      Left            =   7275
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   2937
      BackColor       =   12632256
      BorderStyle     =   1
      Begin xrControl.xrButton cmdButton 
         CausesValidation=   0   'False
         Height          =   450
         Left            =   120
         TabIndex        =   0
         Top             =   1125
         Width           =   1710
         _ExtentX        =   3016
         _ExtentY        =   794
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
         Picture         =   "frmCPSerial.frx":0000
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00E0E0E0&
         Height          =   945
         Index           =   1
         Left            =   150
         Top             =   135
         Width           =   1650
      End
      Begin VB.Shape Shape2 
         Height          =   1005
         Index           =   0
         Left            =   120
         Top             =   105
         Width           =   1710
      End
   End
End
Attribute VB_Name = "frmCPSerial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmCPSerial"

Private oSkin As clsFormSkin
Private oRS As ADODB.Recordset

Dim p_sStockIDx As String
Dim p_sBarrCode As String
Dim p_sDescript As String
Dim p_sBrandNme As String
Dim p_sModelNme As String
Dim p_sColorNme As String
Dim p_sCategrNm As String
Dim p_sBranchCd As String

Property Let StockID(lsStockID As String)
   p_sStockIDx = lsStockID
End Property

Property Let Barcode(lsBarcode As String)
   p_sBarrCode = lsBarcode
End Property

Property Let Description(lsDescript As String)
   p_sDescript = lsDescript
End Property

Property Let Brand(lsBrand As String)
   p_sBrandNme = lsBrand
End Property

Property Let Model(lsModel As String)
   p_sModelNme = lsModel
End Property

Property Let Color(lsColor As String)
   p_sColorNme = lsColor
End Property

Property Let Category(lsCategory As String)
   p_sCategrNm = lsCategory
End Property

Property Let Branch(lsBranch As String)
   p_sBranchCd = lsBranch
End Property

Private Sub cmdButton_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
   Dim lnCtr As Integer
   
   txtField(0).Text = p_sBarrCode
   txtField(1).Text = p_sDescript
   txtField(2).Text = p_sCategrNm
   txtField(3).Text = p_sModelNme
   txtField(4).Text = p_sBrandNme
   txtField(5).Text = p_sColorNme

   Set oRS = New ADODB.Recordset
   oRS.Open "SELECT" _
               & "  a.sSerialNo" _
               & ", b.sCompnyNm" _
               & ", d.sReferNox" _
               & ", d.sSalesInv" _
               & ", a.cLocation" _
               & ", a.cUnitType" _
               & ", a.cSoldStat" _
            & " FROM CP_Inventory_Serial a" _
               & " LEFT JOIN Client_Master b" _
                  & " ON a.sSupplier = b.sClientID" _
               & " LEFT JOIN CP_PO_Receiving_Serial c" _
                  & " LEFT JOIN CP_PO_Receiving_Master d" _
                     & " ON c.sTransNox = d.sTransNox" _
                  & " ON a.sSerialID = c.sSerialID" _
            & " WHERE a.sStockIDx = " & strParm(p_sStockIDx) _
               & " AND a.sBranchCd = " & strParm(p_sBranchCd) _
               & " AND (a.cLocation = " & xeLocBranch & " OR a.cLocation = " & xeLocServiceCenter & ")" _
            & " GROUP BY a.sSerialNo" _
            & " ORDER BY a.sSerialNo" _
   , oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
                  
    With MSFlexGrid1
      .Cols = 7
      .Font = "MS Sans Serif"
      
      'column title
      .TextMatrix(0, 1) = "SerialNo"
      .TextMatrix(0, 2) = "Supplier"
      .TextMatrix(0, 3) = "PO RefNo"
      .TextMatrix(0, 4) = "PO SInvN"
      .TextMatrix(0, 5) = "U-Type"
      .TextMatrix(0, 6) = "Status"
      
      
      'column alignment
      .Row = 0
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = 1
      Next
      
      'column width
      .ColWidth(0) = 330
      .ColWidth(1) = 2500
      .ColWidth(2) = 2100
      .ColWidth(3) = 1000
      .ColWidth(4) = 1000
      .ColWidth(5) = 1000
      .ColWidth(6) = 1000
      
      .Row = 1
      .Col = 1
      .ColSel = .Cols - 1
   End With
   
   If oRS.EOF Then Exit Sub
   
   With MSFlexGrid1
      .Rows = oRS.RecordCount + 1
      For lnCtr = 0 To oRS.RecordCount - 1
         .TextMatrix(lnCtr + 1, 0) = lnCtr + 1
         .TextMatrix(lnCtr + 1, 1) = oRS("sSerialNo")
         .TextMatrix(lnCtr + 1, 2) = IFNull(oRS("sCompnyNm"), "")
         .TextMatrix(lnCtr + 1, 3) = IFNull(oRS("sReferNox"), "")
         .TextMatrix(lnCtr + 1, 4) = IFNull(oRS("sSalesInv"), "")
         
         If oRS("cLocation") = xeLocServiceCenter Then
            .TextMatrix(lnCtr + 1, 5) = "SERVICE CENTER"
         Else
            Select Case oRS("cUnitType")
            Case 0: .TextMatrix(lnCtr + 1, 5) = "Demo Unit"
            Case 1: .TextMatrix(lnCtr + 1, 5) = "Regular"
            Case 2: .TextMatrix(lnCtr + 1, 5) = "Free Unit"
            Case 3: .TextMatrix(lnCtr + 1, 5) = "Live Unit"
            Case 4: .TextMatrix(lnCtr + 1, 5) = "Service Unit"
            End Select
         End If
         
        If IFNull(oRS("cSoldStat"), "0") = 1 Then
            .TextMatrix(lnCtr + 1, 6) = "Sales Return"
        Else
            .TextMatrix(lnCtr + 1, 6) = "Regular"
        End If
         
         oRS.MoveNext
      Next
      
      If .Rows > 22 Then
         .ColWidth(2) = 2670
         .ColWidth(4) = 2850
      End If
      
      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      .ColAlignment(4) = 1
      .ColAlignment(5) = 1
      .ColAlignment(6) = 1
      
   End With
End Sub

Private Sub Form_Load()
   Dim nCtr As Integer
   Dim lsOldProc As String
   
   lsOldProc = "Form_Load"
   ''On Error GoTo errProc

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.DisableClose = True
   oSkin.ApplySkin xeFormLedger
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
   Set oRS = Nothing
End Sub

Private Sub MSFlexGrid1_GotFocus()
   MSFlexGrid1.BackColorSel = &HA4A36A
End Sub

Private Sub MSFlexGrid1_LostFocus()
   MSFlexGrid1.BackColorSel = &H8000000D
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

