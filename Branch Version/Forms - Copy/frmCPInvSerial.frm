VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmCPInvSerial 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Serial"
   ClientHeight    =   5685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11055
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   11055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin xrControl.xrFrame xrFrame1 
      Height          =   4980
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   4260
      _ExtentX        =   7514
      _ExtentY        =   8784
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   825
         TabIndex        =   17
         Top             =   1965
         Width           =   2730
      End
      Begin VB.ComboBox cmbSerial 
         Height          =   315
         Index           =   1
         ItemData        =   "frmCPInvSerial.frx":0000
         Left            =   825
         List            =   "frmCPInvSerial.frx":000A
         TabIndex        =   13
         Text            =   "Combo1"
         Top             =   2610
         Width           =   2715
      End
      Begin VB.ComboBox cmbSerial 
         Height          =   315
         Index           =   0
         ItemData        =   "frmCPInvSerial.frx":001B
         Left            =   825
         List            =   "frmCPInvSerial.frx":002E
         TabIndex        =   12
         Text            =   "Combo1"
         Top             =   2265
         Width           =   2715
      End
      Begin VB.TextBox txtSerial 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   2955
         Width           =   3240
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   1065
         TabIndex        =   7
         Top             =   1455
         Width           =   2025
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   1065
         TabIndex        =   4
         Top             =   855
         Width           =   2880
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   1065
         TabIndex        =   3
         Top             =   1155
         Width           =   2880
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
         Top             =   120
         Width           =   2895
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   1065
         TabIndex        =   2
         Top             =   555
         Width           =   2880
      End
      Begin xrControl.xrButton cmdButton 
         CausesValidation=   0   'False
         Height          =   450
         Index           =   1
         Left            =   1350
         TabIndex        =   19
         Top             =   4335
         Width           =   1350
         _ExtentX        =   2381
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
         Picture         =   "frmCPInvSerial.frx":0069
      End
      Begin xrControl.xrButton cmdDetail 
         Height          =   450
         Index           =   2
         Left            =   2775
         TabIndex        =   20
         Top             =   4335
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   794
         SizeCW          =   0
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
         Picture         =   "frmCPInvSerial.frx":07E3
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IMEI"
         Height          =   195
         Index           =   5
         Left            =   90
         TabIndex        =   18
         Top             =   2010
         Width           =   330
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Branch"
         Height          =   195
         Index           =   10
         Left            =   75
         TabIndex        =   16
         Top             =   3000
         Width           =   510
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         Height          =   195
         Index           =   12
         Left            =   75
         TabIndex        =   15
         Top             =   2670
         Width           =   450
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Location"
         Height          =   195
         Index           =   13
         Left            =   75
         TabIndex        =   14
         Top             =   2340
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Color"
         Height          =   195
         Index           =   11
         Left            =   105
         TabIndex        =   9
         Top             =   1500
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model Name"
         Height          =   195
         Index           =   2
         Left            =   105
         TabIndex        =   8
         Top             =   1200
         Width           =   900
      End
      Begin VB.Line Line1 
         X1              =   75
         X2              =   4020
         Y1              =   1860
         Y2              =   1860
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
         TabIndex        =   6
         Top             =   585
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Brand Name"
         Height          =   195
         Index           =   1
         Left            =   105
         TabIndex        =   5
         Top             =   930
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
         TabIndex        =   0
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
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5025
      Left            =   4395
      TabIndex        =   10
      Tag             =   "et0;eb0;et0;bc2"
      Top             =   540
      Width           =   6570
      _ExtentX        =   11589
      _ExtentY        =   8864
      _Version        =   393216
      Rows            =   19
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
End
Attribute VB_Name = "frmCPInvSerial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Private Const pxeMODULENAME = "frmCPInvSerial"
'
'Private oSkin As clsFormSkin
'Private p_oRS As ADODB.Recordset
'Private p_oInvCount As clsCPInvCount
'
'Dim p_sStockIDx As String
'Dim p_sBarrCode As String
'Dim p_sDescript As String
'Dim p_sBrandNme As String
'Dim p_sModelNme As String
'Dim p_sColorNme As String
'Dim p_sBranchCd As String
'
'Property Let StockID(lsStockID As String)
'   p_sStockIDx = lsStockID
'End Property
'
'Property Let Barcode(lsBarcode As String)
'   p_sBarrCode = lsBarcode
'End Property
'
'Property Let Description(lsDescript As String)
'   p_sDescript = lsDescript
'End Property
'
'Property Let Brand(lsBrand As String)
'   p_sBrandNme = lsBrand
'End Property
'
'Property Let Model(lsModel As String)
'   p_sModelNme = lsModel
'End Property
'
'Property Let Color(lsColor As String)
'   p_sColorNme = lsColor
'End Property
'
'Property Let Branch(lsBranch As String)
'   p_sBranchCd = lsBranch
'End Property
'
'Property Set InvCount(loInvCount As clsCPInvCount)
'   Set p_oInvCount = loInvCount
'End Property
'
'Private Sub cmdButton_Click()
'   Unload Me
'End Sub
'
'Private Sub Form_Activate()
'   Dim lnCtr As Integer
'
'   txtField(0).Text = p_sBarrCode
'   txtField(1).Text = p_sDescript
'   txtField(2).Text = p_sModelNme
'   txtField(3).Text = p_sBrandNme
'   txtField(4).Text = p_sColorNme
'
'    With MSFlexGrid1
'      .Cols = 13
'      .Rows = 2
'      .Font = "MS San Serif"
'
'      'Column Title
'      .TextMatrix(0, 1) = "Serial ID"
'      .TextMatrix(0, 2) = "IMEI"
'      .TextMatrix(0, 3) = "Old-Loc"
'      .TextMatrix(0, 4) = "Old-Stat"
'      .TextMatrix(0, 5) = "Old-Branch"
'      .TextMatrix(0, 6) = "Old-Stock"
'      .TextMatrix(0, 7) = "Location"
'      .TextMatrix(0, 8) = "Status"
'      .TextMatrix(0, 9) = "Branch"
'      .TextMatrix(0, 10) = "Stock ID"
'      .TextMatrix(0, 11) = "Branch Cd"
'      .TextMatrix(0, 12) = ""
'
'      .Row = 0
'
'      'Column Alignment
'      For lnCtr = 0 To .Cols - 1
'         .Col = lnCtr
'         .CellFontBold = True
'         .ColAlignment(lnCtr) = 1
'         .CellAlignment = 1
'      Next
'
'      'Column Width
'      .ColWidth(0) = 450
'      .ColWidth(1) = 0 '1920
'      .ColWidth(2) = 1920 ' 2250
'      .ColWidth(3) = 0
'      .ColWidth(4) = 0 '2000
'      .ColWidth(5) = 0
'      .ColWidth(6) = 0
'      .ColWidth(7) = 1300 '0
'      .ColWidth(8) = 900 '650
'      .ColWidth(9) = 3350 '650
'      .ColWidth(10) = 0 '650
'      .ColWidth(11) = 0
'      .ColWidth(12) = 0
'
'      .Col = 1
'      .Row = 1
'   End With
'
'   If p_oRS.EOF Then Exit Sub
'
'   With MSFlexGrid1
'      .Rows = p_oRS.RecordCount + 1
'      For lnCtr = 0 To p_oRS.RecordCount - 1
'         .TextMatrix(lnCtr + 1, 0) = lnCtr + 1
'         .TextMatrix(lnCtr + 1, 1) = p_oRS("sSerialNo")
'         .TextMatrix(lnCtr + 1, 2) = IFNull(p_oRS("sCompnyNm"), "")
'         .TextMatrix(lnCtr + 1, 3) = IFNull(p_oRS("sReferNox"), "")
'         .TextMatrix(lnCtr + 1, 4) = IFNull(p_oRS("sSalesInv"), "")
'
'         If p_oRS("cLocation") = xeLocServiceCenter Then
'            .TextMatrix(lnCtr + 1, 5) = "SERVICE CENTER"
'         Else
'            Select Case p_oRS("cUnitType")
'            Case 0: .TextMatrix(lnCtr + 1, 5) = "Demo Unit"
'            Case 1: .TextMatrix(lnCtr + 1, 5) = "Regular"
'            Case 2: .TextMatrix(lnCtr + 1, 5) = "Free Unit"
'            Case 3: .TextMatrix(lnCtr + 1, 5) = "Live Unit"
'            Case 4: .TextMatrix(lnCtr + 1, 5) = "Service Unit"
'            End Select
'         End If
'         p_oRS.MoveNext
'      Next
'
'      If .Rows > 22 Then
'         .ColWidth(2) = 2670
'         .ColWidth(4) = 2850
'      End If
'
'      .ColAlignment(1) = 1
'      .ColAlignment(2) = 1
'      .ColAlignment(3) = 1
'      .ColAlignment(4) = 1
'      .ColAlignment(5) = 1
'   End With
'End Sub
'
'Private Sub Form_Load()
'   Dim nCtr As Integer
'   Dim lsOldProc As String
'
'   lsOldProc = "Form_Load"
'   'On Error GoTo errProc
'
'   Set oSkin = New clsFormSkin
'   Set oSkin.AppDriver = oApp
'   Set oSkin.Form = Me
'   oSkin.DisableClose = True
'   oSkin.ApplySkin xeFormLedger
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " & " )", True
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'   Set oSkin = Nothing
'   Set p_oRS = Nothing
'End Sub
'
'Private Sub MSFlexGrid1_GotFocus()
'   MSFlexGrid1.BackColp_oRSel = &HA4A36A
'End Sub
'
'Private Sub MSFlexGrid1_LostFocus()
'   MSFlexGrid1.BackColp_oRSel = &H8000000D
'End Sub
'
'Private Sub ShowError(ByVal lsProcName As String, Optional bEnd As Boolean = False)
'   With oApp
'      .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
'      If bEnd Then
'         .xShowError
'         End
'      Else
'         With Err
'            .Raise .Number, .Source, .Description
'         End With
'      End If
'   End With
'End Sub
'
