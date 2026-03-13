VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCP_Serial_Ledger 
   BorderStyle     =   0  'None
   Caption         =   "IMEI No. Ledger"
   ClientHeight    =   7605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5130
      Left            =   120
      TabIndex        =   0
      Tag             =   "et0;eb0;et0;bc2"
      Top             =   2340
      Width           =   7575
      _ExtentX        =   13361
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
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   3096
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   2
         Left            =   1245
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1380
         Width           =   4275
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
         Index           =   1
         Left            =   1245
         MaxLength       =   50
         TabIndex        =   1
         Top             =   1125
         Width           =   4275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Brand-Model"
         Height          =   195
         Index           =   11
         Left            =   210
         TabIndex        =   6
         Top             =   1140
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
         Top             =   1395
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
      Left            =   5850
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
      Begin VB.Image Image1 
         Height          =   660
         Left            =   405
         Picture         =   "frmCP_Serial_Ledger.frx":0000
         Stretch         =   -1  'True
         Top             =   270
         Width           =   1005
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
Attribute VB_Name = "frmCP_Serial_Ledger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private oSkin As FormSkin
Private oRs As ADODB.Recordset

Dim psClient As String

Property Let BarrCode(BarrCode As String)
   psClient = BarrCode
End Property

Private Sub cmdButton_Click()
   Set oRs = Nothing
   Unload Me
End Sub

Private Sub Form_Load()
   Dim lnctr As Integer
   Dim lssql As String

   Set oSkin = New FormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.DisableClose = True
   oSkin.ApplySkin xeFormLedger

   Set oRs = New ADODB.Recordset
   If oRs.State = adStateOpen Then oRs.Close
  
   lssql = "SELECT" _
               & " a.sSerialID, " _
               & " a.sIMEINoxx, " _
               & " b.sSourceNo, " _
               & " b.dTransact, " _
               & " c.sDescript, " _
               & " d.sBrandNme, " _
               & " e.sModelNme, " _
               & " f.sBranchNm, " _
               & " g.sSourceNm  " _
         & " FROM CP_Serial_Master a " _
            & " LEFT JOIN CP_Serial_Ledger b " _
               & " ON a.sSerialID = b.sSerialID " _
            & " LEFT JOIN CP_Inventory c " _
               & " ON a.sStockIDx = c.sStockIDx " _
            & " LEFT JOIN Brand d " _
               & " ON c.sBrandIDx = d.sBrandIDx " _
            & " LEFT JOIN Model e " _
               & " ON c.sModelIDx = e.sModelIDx " _
            & " LEFT JOIN Branch f " _
               & " ON b.sBranchCd = f.sBranchCd " _
            & " LEFT JOIN xxxTransactionSource g" _
               & " ON b.sSourceCd = g.sSourceID" _

   lssql = lssql _
            & " WHERE a.sSerialID = '" & psClient & "' " _
            & " ORDER BY b.nEntryNox"
   
   oRs.Open lssql, oApp.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
            
   If oRs.EOF Then Exit Sub
    
   txtField(0).Text = oRs("sIMEINoxx")
   txtField(1).Text = Trim(IIf(IsNull(oRs("sBrandNme")), "", oRs("sBrandNme")) & " - " & _
                     IIf(IsNull(oRs("sModelNme")), "", oRs("sModelNme")))
   txtField(2).Text = IIf(IsNull(oRs("sDescript")), "", oRs("sDescript"))

   For lnctr = 0 To txtField.Count - 1
      txtField(lnctr).Locked = True
   Next
   
   With MSFlexGrid1
      .Rows = oRs.RecordCount + 1
      .Cols = 5
      .Font = "MS Sans Serif"
      
      'column title
      .TextMatrix(0, 1) = "Date"
      .TextMatrix(0, 2) = "Source"
      .TextMatrix(0, 3) = "Source No."
      .TextMatrix(0, 4) = "Branch"
      
      'column alignment
      .Row = 0
      For lnctr = 0 To .Cols - 1
         .Col = lnctr
         .CellFontBold = True
         .CellAlignment = 1
      Next

      For lnctr = 0 To oRs.RecordCount - 1
         .TextMatrix(lnctr + 1, 0) = lnctr + 1
         .TextMatrix(lnctr + 1, 1) = Format(oRs("dTransact"), "MMM-DD-YYYY")
         .TextMatrix(lnctr + 1, 2) = IIf(IsNull(oRs("sSourceNm")) = True, "", oRs("sSourceNm"))
         .TextMatrix(lnctr + 1, 3) = IIf(IsNull(oRs("sSourceNo")) = True, "", oRs("sSourceNo"))
         .TextMatrix(lnctr + 1, 4) = IIf(IsNull(oRs("sBranchNm")) = True, "", oRs("sBranchNm"))
         oRs.MoveNext
      Next
      
      'column width
      .ColWidth(0) = 330
      .ColWidth(1) = 1100
      .ColWidth(2) = 2000
      .ColWidth(3) = 1200
      .ColWidth(4) = 2900
      
      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
   Set oRs = Nothing
End Sub




