VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrcontrol.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmSerial_Ledger 
   BorderStyle     =   0  'None
   Caption         =   "IMEI No. Ledger"
   ClientHeight    =   7590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9045
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   9045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5130
      Left            =   120
      TabIndex        =   0
      Tag             =   "et0;eb0;et0;bc2"
      Top             =   2340
      Width           =   8550
      _ExtentX        =   15081
      _ExtentY        =   9049
      _Version        =   393216
      FocusRect       =   0
      HighLight       =   0
      Appearance      =   0
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
      Height          =   1755
      Left            =   120
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   6660
      _ExtentX        =   11748
      _ExtentY        =   3096
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   2
         Left            =   1245
         MaxLength       =   50
         TabIndex        =   8
         Top             =   1365
         Width           =   5115
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   1245
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1110
         Width           =   5115
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   0
         Left            =   1245
         TabIndex        =   2
         Top             =   210
         Width           =   2430
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   3
         Left            =   1245
         MaxLength       =   50
         TabIndex        =   1
         Top             =   855
         Width           =   5115
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier"
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   9
         Top             =   1365
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Brand-Model"
         Height          =   195
         Index           =   11
         Left            =   210
         TabIndex        =   6
         Top             =   870
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   10
         Left            =   210
         TabIndex        =   5
         Top             =   1125
         Width           =   795
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   1290
         Tag             =   "et0;ht2"
         Top             =   255
         Width           =   2430
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IMEI No."
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   4
         Top             =   255
         Width           =   630
      End
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   1755
      Left            =   6810
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   3096
      BackColor       =   12632256
      BorderStyle     =   1
      Begin xrControl.xrButton cmdButton 
         Height          =   405
         Left            =   90
         TabIndex        =   7
         Top             =   1230
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   714
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
      End
      Begin VB.Shape Shape2 
         Height          =   1125
         Index           =   0
         Left            =   90
         Top             =   75
         Width           =   1635
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00E0E0E0&
         Height          =   1050
         Index           =   1
         Left            =   135
         Top             =   105
         Width           =   1545
      End
   End
End
Attribute VB_Name = "frmSerial_Ledger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private oSkin As FormSkin
Private oRS As ADODB.Recordset

Dim psClient As String

Property Let BarrCode(BarrCode As String)
   psClient = BarrCode
End Property

Private Sub cmdButton_Click()
   Set oRS = Nothing
   Unload Me
End Sub

Private Sub Form_Load()
   Dim lnctr As Integer
   Dim lsSQL As String

   Set oSkin = New FormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.DisableClose = True
   oSkin.ApplySkin xeFormLedger

   Set oRS = New ADODB.Recordset
   If oRS.State = adStateOpen Then oRS.Close
  
   lsSQL = "SELECT" _
               & " c.sBarrcode, " _
               & " c.sDescript, " _
               & " d.sCompnyNm, " _
               & " e.sBrandNme, " _
               & " f.sModelNme, " _
               & " b.dTransact, " _
               & " g.sSourceNm, " _
               & " b.sSourceNo, " _
               & " b.nQtyInxxx, " _
               & " b.nQtyOutxx, " _
               & " b.nEntryNox, " _
               & " a.sBranchCd  "
   lsSQL = lsSQL _
            & " FROM CP_Inventory_Master a" _
               & " LEFT JOIN CP_Inventory_Ledger b" _
                  & " ON a.sStockIDx = b.sStockIDx" _
                  & " AND b.sbranchCd = '" & oApp.BranchCode & "'" _
               & " LEFT JOIN CP_Inventory c " _
                  & " ON a.sStockIdx = c.sStockIDx " _
               & " LEFT JOIN Client_Master d" _
                  & " ON c.sSupplier = d.sClientID" _
               & " LEFT JOIN Brand e" _
                  & " ON c.sBrandIDx = e.sBrandIDx" _
               & " LEFT JOIN Model f" _
                  & " ON c.sModelIDx = f.sModelIDx" _
               & " LEFT JOIN xxxTransactionSource g" _
                  & " ON b.sSourceCd = g.sSourceID"
                  
   lsSQL = lsSQL _
            & " WHERE c.sBarrcode = '" & psClient & "' " _
               & " AND a.sBranchCd = '" & oApp.BranchCode & "'" _
            & " ORDER BY b.nEntryNox"
   
   oRS.Open lsSQL, oApp.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
            
   If oRS.EOF Then Exit Sub
    
   txtField(0).Text = oRS("sBarrcode")
   txtField(1).Text = oRS("sDescript")
   txtField(2).Text = IIf(IsNull(oRS("sCompnyNm")), "", oRS("sCompnyNm"))
   txtField(3).Text = oRS("sBrandNme") & " - " & oRS("sModelNme")

   For lnctr = 0 To txtField.Count - 1
      txtField(lnctr).Locked = True
   Next
   
   With MSFlexGrid1
      .Rows = oRS.RecordCount + 1
      .Cols = 6
      .Font = "MS Sans Serif"
      
      'column title
      .TextMatrix(0, 1) = "Date"
      .TextMatrix(0, 2) = "Source"
      .TextMatrix(0, 3) = "Source No."
      .TextMatrix(0, 4) = "Qty. In"
      .TextMatrix(0, 5) = "Qty. Out"
      
      'column alignment
      .Row = 0
      For lnctr = 0 To .Cols - 1
         .Col = lnctr
         .CellFontBold = True
         .CellAlignment = 1
      Next

      For lnctr = 0 To oRS.RecordCount - 1
         .TextMatrix(lnctr + 1, 0) = lnctr + 1
         .TextMatrix(lnctr + 1, 1) = Format(oRS("dTransact"), "MMM-DD-YYYY")
         .TextMatrix(lnctr + 1, 2) = IIf(IsNull(oRS("sSourceNm")) = True, "", oRS("sSourceNm"))
         .TextMatrix(lnctr + 1, 3) = IIf(IsNull(oRS("sSourceNo")) = True _
         , "", oRS("sSourceNo"))
         .TextMatrix(lnctr + 1, 4) = Format(oRS("nQtyInxxx"), "##0")
         .TextMatrix(lnctr + 1, 5) = Format(oRS("nQtyOutxx"), "##0")
         oRS.MoveNext
      Next
      
      'column width
      .ColWidth(0) = 330
      .ColWidth(1) = 1300
      .ColWidth(2) = 2760
      .ColWidth(3) = 1200
      
      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      
      If .Rows < 22 Then
         .ColWidth(4) = 1450
         .ColWidth(5) = 1450
      Else
         .ColWidth(4) = 1350
         .ColWidth(5) = 1350
      End If
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
   Set oRS = Nothing
End Sub




