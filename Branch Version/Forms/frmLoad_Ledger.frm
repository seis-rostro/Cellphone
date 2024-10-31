VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLoad_Ledger 
   BorderStyle     =   0  'None
   Caption         =   "Load Ledger"
   ClientHeight    =   7590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9675
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7590
   ScaleWidth      =   9675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5130
      Left            =   105
      TabIndex        =   7
      Tag             =   "et0;eb0;et0;bc2"
      Top             =   2340
      Width           =   9150
      _ExtentX        =   16140
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
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   7260
      _ExtentX        =   12806
      _ExtentY        =   3096
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   1245
         MaxLength       =   50
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1125
         Width           =   5760
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   0
         Left            =   1245
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   210
         Width           =   3795
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   2
         Left            =   1245
         MaxLength       =   30
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1380
         Width           =   5760
      End
      Begin MSComCtl2.Animation Progress 
         Height          =   495
         Left            =   5235
         TabIndex        =   8
         TabStop         =   0   'False
         Tag             =   "wt0;fb0"
         Top             =   0
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   873
         _Version        =   393216
         BackColor       =   2334974
         FullWidth       =   39
         FullHeight      =   33
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier"
         Height          =   195
         Index           =   12
         Left            =   210
         TabIndex        =   5
         Top             =   1395
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   10
         Left            =   210
         TabIndex        =   3
         Top             =   1140
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
         Width           =   3795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Barcode"
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   1
         Top             =   255
         Width           =   600
      End
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   1755
      Left            =   7395
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
         TabIndex        =   0
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
         Height          =   1065
         Left            =   120
         Picture         =   "frmLoad_Ledger.frx":0000
         Stretch         =   -1  'True
         Top             =   105
         Width           =   1590
      End
      Begin VB.Shape Shape2 
         Height          =   1125
         Index           =   0
         Left            =   90
         Top             =   75
         Width           =   1635
      End
   End
End
Attribute VB_Name = "frmLoad_Ledger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private oSkin As FormSkin
Private oRS As ADODB.Recordset
Private oRSAmount As ADODB.Recordset

Dim lnctr As Integer
Dim lsSQL As String
Dim Amount As String
Dim psClient As String

Property Let BarrCode(BarrCode As String)
   psClient = BarrCode
End Property

Private Sub cmdButton_Click()
   Set oRS = Nothing
   Set oRSAmount = Nothing
   Unload Me
End Sub

Private Sub Form_Activate()
   Progress.Open App.Path & "\images\BOOKS.AVI"
   Progress.Play
End Sub

Private Sub Form_Deactivate()
   cmdButton.SetFocus
   Progress.Stop
   Progress.Close
End Sub

Private Sub Form_Load()
   
   Set oSkin = New FormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.DisableClose = True
   oSkin.ApplySkin xeFormLedger

   Set oRS = New ADODB.Recordset
   Set oRSAmount = New ADODB.Recordset
   
   If oRS.State = adStateOpen Then oRS.Close
  
   lsSQL = "SELECT " _
               & " a.sBranchCd, " _
               & " c.sBarrcode, " _
               & " c.sDescript, " _
               & " d.sSupplyNm, " _
               & " b.dTransact, " _
               & " e.sSourceNm, " _
               & " b.sSourceNo, " _
               & " b.nQtyInxxx, " _
               & " b.nQtyOutxx, " _
               & " b.sReferNox, " _
               & " b.nQtyOnHnd, " _
               & " b.sPhoneNum, " _
               & " b.nEntryNox, " _
               & " f.sBranchCd  "
   lsSQL = lsSQL _
            & " FROM CP_Inventory_Master a" _
               & " LEFT JOIN ELoad_Ledger b" _
                  & " ON a.sStockIDx = b.sStockIDx" _
               & " LEFT JOIN Branch f " _
                  & " ON b.sBranchCd = f.sBranchCd" _
               & " LEFT JOIN CP_Inventory c " _
                  & " ON a.sStockIdx = c.sStockIDx " _
               & " LEFT JOIN Supplier d" _
                  & " ON c.sSupplier = d.sSupplyID" _
               & " LEFT JOIN xxxTransactionSource e" _
                  & " ON b.sSourceCd = e.sSourceID"
                  
   lsSQL = lsSQL _
            & " WHERE c.sBarrcode = '" & psClient & "' " _
               & " AND a.sBranchCd = '" & oApp.BranchCode & "'" _
               & " AND b.sBranchcd = '" & oApp.BranchCode & "'" _
            & " ORDER BY b.nEntryNox desc, b.stransnox desc "


   oRS.Open lsSQL, oApp.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
      
   If oRS.EOF Then Exit Sub
    
   txtfield(0).Text = oRS("sBarrcode")
   txtfield(1).Text = oRS("sDescript")
   txtfield(2).Text = IIf(IsNull(oRS("sSupplyNm")), "", oRS("sSupplyNm"))

   For lnctr = 0 To txtfield.Count - 1
      txtfield(lnctr).Locked = True
   Next
   
   With MSFlexGrid1
      .Rows = oRS.RecordCount + 1
      .Cols = 8
      .Font = "MS Sans Serif"
      
      'column title
      .TextMatrix(0, 1) = "Date"
      .TextMatrix(0, 2) = "Source"
      .TextMatrix(0, 3) = "Source No."
      .TextMatrix(0, 4) = "Reference No."
      .TextMatrix(0, 5) = "Qty. In"
      .TextMatrix(0, 6) = "Qty. Out"
      .TextMatrix(0, 7) = "Balance"

      
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
         If oRS("sSourceNo") = "" Then
            .TextMatrix(lnctr + 1, 3) = "Branch"
         Else
            .TextMatrix(lnctr + 1, 3) = IIf(IsNull(oRS("sSourceNo")) = True _
                                       , "", oRS("sSourceNo"))
         End If
         .TextMatrix(lnctr + 1, 4) = IIf(IsNull(oRS("sReferNox")), "", Trim(oRS("sReferNox")))
         .TextMatrix(lnctr + 1, 5) = Format(oRS("nQtyInxxx"), "#,###.##")
         .TextMatrix(lnctr + 1, 6) = Format(oRS("nQtyOutxx"), "#,###.##")
         .TextMatrix(lnctr + 1, 7) = Format(oRS("nQtyOnHnd"), "#,###.##")

         oRS.MoveNext
      Next
      
      'column width
      .ColWidth(0) = 400
      .ColWidth(1) = 1000
      If .Rows <= 20 Then
         .ColWidth(2) = 2300
      Else
         .ColWidth(2) = 1950
      End If
      .ColWidth(3) = 1000
      .ColWidth(4) = 1300
      .ColWidth(5) = 950
      .ColWidth(6) = 950
      .ColWidth(7) = 1200
      
      'column Alignment
      .ColAlignment(5) = 6
      .ColAlignment(6) = 6
      .ColAlignment(7) = 6
            
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
   Set oRS = Nothing
   Set oRSAmount = Nothing
End Sub
