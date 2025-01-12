VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmClientLedger 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Client Ledger"
   ClientHeight    =   7680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9060
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   9060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   240
      Left            =   120
      TabIndex        =   10
      Top             =   7245
      Width           =   8550
      _ExtentX        =   15081
      _ExtentY        =   423
      _Version        =   393216
      Appearance      =   0
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
         TabIndex        =   7
         Top             =   1185
         Width           =   5145
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   1590
         TabIndex        =   5
         Top             =   885
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
         TabIndex        =   1
         Top             =   105
         Width           =   1950
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   1590
         TabIndex        =   3
         Top             =   585
         Width           =   5145
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliet ID"
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
         TabIndex        =   0
         Top             =   150
         Width           =   645
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1680
         Tag             =   "et0;ht2"
         Top             =   210
         Width           =   1950
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name"
         Height          =   195
         Index           =   10
         Left            =   210
         TabIndex        =   2
         Top             =   645
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Address"
         Height          =   195
         Index           =   11
         Left            =   210
         TabIndex        =   4
         Top             =   960
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company "
         Height          =   195
         Index           =   12
         Left            =   210
         TabIndex        =   6
         Top             =   1260
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
         CausesValidation=   0   'False
         Height          =   465
         Left            =   120
         TabIndex        =   8
         Top             =   1110
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   820
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
         Picture         =   "frmClientLedger.frx":0000
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
Attribute VB_Name = "frmClientLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmClientLedger"

Private oSkin As clsFormSkin
Private oRS As ADODB.Recordset

Dim psClient As String
Dim pbLoaded As Boolean

Property Let ClientID(ClientID As String)
   psClient = ClientID
End Property

Private Sub cmdButton_Click()
   Unload Me
End Sub

Public Function browseLedger() As Boolean
   Dim lnCtr As Integer
   Dim lsOldProc As String
   
   lsOldProc = "BrowseLedger"
   'On Error GoTo errProc

   browseLedger = False
   
   If pbLoaded = False Then GoTo endProc
   Set oRS = New ADODB.Recordset
   If oRS.State = adStateOpen Then oRS.Close

   oRS.Open "SELECT" _
               & "  a.sClientID" _
               & ", a.sCompnyNm" _
               & ", CONCAT(a.sLastName, ', ', a.sFrstName, ' ', a.sMiddName) xCustName" _
               & ", CONCAT(a.sAddressx, ', ', d.sTownName, ', ', e.sProvName, ' ', d.sZippCode) as xAddressx" _
               & ", b.dTransact" _
               & ", c.sSourceNm" _
               & ", b.sSourceNo" _
               & ", b.nCreditxx" _
               & ", b.nDebitxxx" _
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
            & " ORDER BY b.dTransact,b.sSourceCd,b.sSourceNo DESC" _
            , oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   
   If Not oRS.EOF Then GoTo endProc
   
endProc:
   browseLedger = True
   Exit Function
errProc:
   Set oRS = Nothing
   ShowError lsOldProc & "( " & " )"
End Function

Private Sub Form_Load()
   Dim lnCtr As Integer
   Dim lsOldProc As String
   
   lsOldProc = "Form_Load"
   'On Error GoTo errProc

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   
   oSkin.ApplySkin xeFormLedger
   
   For lnCtr = 0 To txtField.Count - 1
      txtField(lnCtr).Locked = True
   Next
   
   pbLoaded = True
   If browseLedger Then
      ProgressBar1.Visible = True
      ProgressBar1.Max = oRS.RecordCount
      
      txtField(0).Text = Format(oRS("sClientID"), "@@@@-@@@@@@")
      txtField(1).Text = oRS("xCustName")
      txtField(2).Text = oRS("xAddressx")
      txtField(3).Text = IIf(IsNull(oRS("sCompnyNm")), "", oRS("sCompnyNm"))
   
      With MSFlexGrid1
         .Cols = 6
         .Font = "MS Sans Serif"
   
         'column title
         .TextMatrix(0, 1) = "Date"
         .TextMatrix(0, 2) = "Source"
         .TextMatrix(0, 3) = "Source No."
         .TextMatrix(0, 4) = "Credit"
         .TextMatrix(0, 5) = "Debit"
   
         'column alignment
         .Row = 0
         For lnCtr = 0 To .Cols - 1
            .Col = lnCtr
            .CellFontBold = True
            .CellAlignment = 1
         Next
         
         .Rows = IIf(oRS.RecordCount = 0, 2, oRS.RecordCount + 1)
         For lnCtr = 0 To oRS.RecordCount - 1
            .TextMatrix(lnCtr + 1, 0) = lnCtr + 1
            .TextMatrix(lnCtr + 1, 1) = Format(oRS("dTransact"), "DD-MMM-YYYY")
            .TextMatrix(lnCtr + 1, 2) = IIf(IsNull(oRS("sSourceNm")) = True, "", oRS("sSourceNm"))
            .TextMatrix(lnCtr + 1, 3) = IIf(IsNull(oRS("sSourceNo")) = True _
            , "", Format(oRS("sSourceNo"), "@@-@@@@@@"))
            .TextMatrix(lnCtr + 1, 4) = Format(oRS("nCreditxx"), "#,##0.00")
            .TextMatrix(lnCtr + 1, 5) = Format(oRS("nDebitxxx"), "#,##0.00")
   
            ProgressBar1.Value = lnCtr + 1
            If ProgressBar1.Value = ProgressBar1.Max Then ProgressBar1.Visible = False
            oRS.MoveNext
         Next
   
         'column width
         .ColWidth(0) = 330
         .ColWidth(1) = 1100
         .ColWidth(3) = 1200
         .ColWidth(4) = 1000
         .ColWidth(5) = 1000
         .ColAlignment(3) = 1
   
         If .Rows < 22 Then
            .ColWidth(2) = 3800
         Else
            .ColWidth(2) = 3600
         End If
         .Row = 1
         .Col = 1
         .ColSel = .Cols - 1
      End With
   End If
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
End Sub

Private Sub MSFlexGrid1_GotFocus()
   With MSFlexGrid1
      .BackColorSel = &HA4A36A
   End With
End Sub

Private Sub MSFlexGrid1_LostFocus()
   With MSFlexGrid1
      .BackColorSel = &H8000000D
   End With
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
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
