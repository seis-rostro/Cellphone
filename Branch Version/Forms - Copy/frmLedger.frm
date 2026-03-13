VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLedger 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Supplier Ledger"
   ClientHeight    =   7680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9060
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   9060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   285
      Left            =   135
      TabIndex        =   10
      Top             =   7260
      Visible         =   0   'False
      Width           =   8520
      _ExtentX        =   15028
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5265
      Left            =   120
      TabIndex        =   9
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
      Width           =   6885
      _ExtentX        =   12144
      _ExtentY        =   2937
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   1590
         TabIndex        =   8
         Top             =   1200
         Width           =   5145
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   1590
         TabIndex        =   6
         Top             =   900
         Width           =   5145
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
         Left            =   1590
         TabIndex        =   2
         Top             =   135
         Width           =   1950
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   1590
         TabIndex        =   4
         Top             =   600
         Width           =   5145
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client ID"
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
         Left            =   210
         TabIndex        =   1
         Top             =   180
         Width           =   750
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1665
         Tag             =   "et0;ht2"
         Top             =   225
         Width           =   1950
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name"
         Height          =   195
         Index           =   10
         Left            =   210
         TabIndex        =   3
         Top             =   660
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Address"
         Height          =   195
         Index           =   11
         Left            =   210
         TabIndex        =   7
         Top             =   975
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company "
         Height          =   195
         Index           =   12
         Left            =   210
         TabIndex        =   5
         Top             =   1275
         Width           =   705
      End
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   1665
      Left            =   7035
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   2937
      BackColor       =   12632256
      BorderStyle     =   1
      Begin xrControl.xrButton cmdButton 
         Height          =   450
         Left            =   120
         TabIndex        =   0
         Top             =   1110
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
         Picture         =   "frmLedger.frx":0000
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00E0E0E0&
         Height          =   945
         Index           =   1
         Left            =   150
         Top             =   135
         Width           =   1290
      End
      Begin VB.Shape Shape2 
         Height          =   1005
         Index           =   0
         Left            =   120
         Top             =   105
         Width           =   1350
      End
   End
End
Attribute VB_Name = "frmLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Const pxeMODULENAME = "frmLedger"

Private oSkin As clsFormSkin
Private oRS As ADODB.Recordset

Dim psClient As String

Property Let ClientID(ClientID As String)
   psClient = ClientID
End Property

Private Sub cmdButton_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Dim nCtr As Integer
   Dim lsOldProc As String
   
   lsOldProc = "Form_Load"
   'On Error GoTo errProc

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.DisableClose = True
   oSkin.ApplySkin xeFormLedger
   
   For nCtr = 0 To txtField.Count - 1
      txtField(nCtr).Locked = True
   Next

   Set oRS = New ADODB.Recordset
   
   If oRS.State = adStateOpen Then oRS.Close
         
   oRS.Open "SELECT" _
               & "  a.sClientID" _
               & ", a.sCompnyNm" _
               & ", CONCAT(a.sLastName, ', ', a.sFrstName, ' ', a.sMiddName) as CustomerName" _
               & ", CONCAT(a.sAddressx, ', ', d.sTownName, ', ', e.sProvName) as Address" _
               & ", b.dTransact" _
               & ", c.sSourceNm" _
               & ", b.sSourceNo" _
               & ", b.nCreditxx" _
               & ", b.nDebitxxx" _
               & ", b.nABalance" _
            & " FROM Client_Master a" _
               & " LEFT JOIN Client_Ledger b" _
                  & " ON a.sClientID = b.sClientID" _
               & " LEFT JOIN xxxTransactionSource c" _
                  & " ON b.sSourceCd = c.sSourceID" _
               & " LEFT JOIN TownCity d" _
                  & " ON a.sTownIDxx = d.sTownIDxx" _
               & " LEFT JOIN Province e" _
                  & " ON d.sProvIDxx = e.sProvIDxx" _
            & " WHERE a.sClientID = " & strParm(psClient) _
               & " AND b.sBranchCd = " & strParm(oApp.BranchCode) _
            & " ORDER BY b.dTransact" _
            , oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
            
   If oRS.EOF Then Exit Sub
   
   txtField(0).Text = Format(oRS("sClientID"), "@@-@@@@@@")
   txtField(1).Text = IIf(IsNull(oRS("CustomerName")), "", oRS("CustomerName"))
   txtField(2).Text = IIf(IsNull(oRS("Address")), "", oRS("Address"))
   txtField(3).Text = IIf(IsNull(oRS("sCompnyNm")), "", oRS("sCompnyNm"))
   
   With MSFlexGrid1
      .Cols = 7
      .Rows = oRS.RecordCount + 1
      .Font = "MS Sans Serif"
      
      'column title
      .TextMatrix(0, 1) = "Date"
      .TextMatrix(0, 2) = "Source"
      .TextMatrix(0, 3) = "Source No."
      .TextMatrix(0, 4) = "Credit"
      .TextMatrix(0, 5) = "Debit"
      .TextMatrix(0, 6) = "Act. Bal."
      
      'column alignment
      .Row = 0
      For nCtr = 0 To .Cols - 1
         .Col = nCtr
         .CellFontBold = True
         .CellAlignment = 1
      Next
      
      For nCtr = 0 To oRS.RecordCount - 1
         .TextMatrix(nCtr + 1, 0) = nCtr + 1
         .TextMatrix(nCtr + 1, 1) = Format(oRS("dTransact"), "MMM-DD-YYYY")
         .TextMatrix(nCtr + 1, 2) = IIf(IsNull(oRS("sSourceNm")), "", oRS("sSourceNm"))
         .TextMatrix(nCtr + 1, 3) = IIf(IsNull(oRS("sSourceNo")) _
         , "", Format(oRS("sSourceNo"), "@@-@@@@@@@@"))
         .TextMatrix(nCtr + 1, 4) = Format(oRS("nCreditxx"), "#,##0.00")
         .TextMatrix(nCtr + 1, 5) = Format(oRS("nDebitxxx"), "#,##0.00")
         .TextMatrix(nCtr + 1, 6) = Format(oRS("nABalance"), "#,##0.00")
         oRS.MoveNext
      Next
      
      'column width
      .ColWidth(0) = 330
      .ColWidth(1) = 1300
      .ColWidth(3) = 1140
      .ColWidth(4) = 1100
      .ColWidth(5) = 1100
      .ColWidth(6) = 1100
      
      If .Rows < 22 Then
         .ColWidth(2) = 2390
      Else
         .ColWidth(2) = 2190
      End If
      
      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      .ColAlignment(4) = 6
      .ColAlignment(5) = 6
      .ColAlignment(6) = 6
      
      .Row = 1
      .Col = 1
      .ColSel = .Cols - 1
   End With

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

