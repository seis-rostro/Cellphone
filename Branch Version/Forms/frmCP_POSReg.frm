VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0"; "xrGridControl.ocx"
Begin VB.Form frmCP_POSReg 
   BorderStyle     =   0  'None
   Caption         =   "Cellphone Sales"
   ClientHeight    =   7170
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11895
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   11895
   ShowInTaskbar   =   0   'False
   Begin xrGridEditor.GridEditor GridEditor1 
      Height          =   2205
      Left            =   240
      TabIndex        =   21
      Tag             =   "et0;eb0;et0;bc2"
      Top             =   3360
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   3889
      AllowBigSelection=   -1  'True
      AutoAdd         =   -1  'True
      AutoNumber      =   -1  'True
      BACKCOLOR       =   -2147483643
      BACKCOLORBKG    =   8421504
      BACKCOLORFIXED  =   -2147483633
      BACKCOLORSEL    =   -2147483635
      BORDERSTYLE     =   1
      COLS            =   2
      FILLSTYLE       =   0
      FIXEDCOLS       =   1
      FIXEDROWS       =   1
      FOCUSRECT       =   1
      EDITORBACKCOLOR =   -2147483643
      EDITORFORECOLOR =   -2147483640
      FORECOLOR       =   -2147483640
      FORECOLORFIXED  =   -2147483630
      FORECOLORSEL    =   -2147483634
      FORMATSTRING    =   ""
      Object.HEIGHT          =   2205
      GRIDCOLOR       =   12632256
      GRIDCOLORFIXED  =   0
      BeginProperty GRIDFONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GRIDLINES       =   1
      GRIDLINESFIXED  =   2
      GRIDLINEWIDTH   =   1
      MOUSEICON       =   "frmCP_POSReg.frx":0000
      MOUSEPOINTER    =   0
      REDRAW          =   -1  'True
      RIGHTTOLEFT     =   0   'False
      ROWS            =   2
      SCROLLBARS      =   3
      SCROLLTRACK     =   0   'False
      SELECTIONMODE   =   0
      Object.TOOLTIPTEXT     =   ""
      WORDWRAP        =   0   'False
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   5925
      Index           =   0
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   1110
      Width           =   10185
      _ExtentX        =   17965
      _ExtentY        =   10451
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   17
         Left            =   1170
         MaxLength       =   30
         TabIndex        =   27
         Top             =   5460
         Width           =   2640
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   9
         Left            =   7575
         MaxLength       =   30
         TabIndex        =   30
         Tag             =   "ht0;ft0"
         Text            =   "0.00"
         Top             =   4500
         Width           =   2430
      End
      Begin VB.CheckBox chkBox 
         Caption         =   "with Advance Payment"
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   3900
         TabIndex        =   28
         Tag             =   "et0;fb0"
         Top             =   5475
         Width           =   2265
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   8010
         MaxLength       =   6
         TabIndex        =   18
         Top             =   1230
         Width           =   1995
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   18
         Left            =   8010
         MaxLength       =   6
         TabIndex        =   20
         Top             =   1530
         Width           =   1995
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   8010
         MaxLength       =   10
         TabIndex        =   16
         Top             =   930
         Width           =   1995
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   1275
         TabIndex        =   10
         Top             =   915
         Width           =   5220
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   6
         Left            =   1170
         MaxLength       =   30
         TabIndex        =   23
         Top             =   4500
         Width           =   4650
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   5
         Left            =   7575
         MaxLength       =   30
         TabIndex        =   34
         Tag             =   "ht0"
         Text            =   "0.00"
         Top             =   5340
         Width           =   2430
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   645
         Index           =   7
         Left            =   1170
         MaxLength       =   128
         TabIndex        =   25
         Top             =   4800
         Width           =   4650
      End
      Begin VB.ComboBox cmbField 
         Height          =   315
         Index           =   1
         Left            =   8010
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   570
         Width           =   1995
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   1275
         MaxLength       =   50
         TabIndex        =   7
         Top             =   615
         Width           =   2310
      End
      Begin VB.CheckBox chkBox 
         Caption         =   "Company Sales"
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   3735
         TabIndex        =   8
         Tag             =   "et0;fb0"
         Top             =   645
         Width           =   2265
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   960
         Index           =   4
         Left            =   1275
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1215
         Width           =   5220
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1275
         TabIndex        =   5
         Top             =   210
         Width           =   1950
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Salesman"
         Height          =   285
         Index           =   14
         Left            =   420
         TabIndex        =   26
         Top             =   5490
         Width           =   690
      End
      Begin VB.Label lblAdvPayment 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7575
         TabIndex        =   32
         Top             =   5040
         Width           =   2430
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Adv. Payment"
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
         Left            =   6360
         TabIndex        =   31
         Top             =   5070
         Width           =   1185
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Term"
         Height          =   315
         Index           =   16
         Left            =   6855
         TabIndex        =   15
         Top             =   975
         Width           =   1110
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction No"
         Height          =   285
         Index           =   9
         Left            =   75
         TabIndex        =   4
         Top             =   225
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "UNKNOWN"
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
         Left            =   6885
         TabIndex        =   35
         Tag             =   "eb0;et0"
         Top             =   210
         Width           =   2955
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H8000000F&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   285
         Index           =   0
         Left            =   6795
         Tag             =   "et0;et0"
         Top             =   195
         Width           =   3165
      End
      Begin VB.Shape Shape4 
         Height          =   330
         Index           =   0
         Left            =   6765
         Top             =   165
         Width           =   3210
      End
      Begin VB.Shape Shape3 
         Height          =   390
         Index           =   0
         Left            =   6735
         Top             =   135
         Width           =   3270
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   285
         Index           =   12
         Left            =   165
         TabIndex        =   24
         Top             =   4860
         Width           =   690
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Receipt No."
         Height          =   285
         Index           =   10
         Left            =   6855
         TabIndex        =   19
         Top             =   1560
         Width           =   945
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   8
         Left            =   6360
         TabIndex        =   29
         Top             =   4545
         Width           =   720
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "A&mount Paid"
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
         Left            =   6360
         TabIndex        =   33
         Top             =   5415
         Width           =   1125
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Approved By"
         Height          =   330
         Index           =   7
         Left            =   165
         TabIndex        =   22
         Top             =   4545
         Width           =   1050
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Type"
         Height          =   285
         Index           =   5
         Left            =   6855
         TabIndex        =   13
         Top             =   615
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "S.I. No."
         Height          =   285
         Index           =   2
         Left            =   6855
         TabIndex        =   17
         Top             =   1260
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cust. Address"
         Height          =   195
         Index           =   11
         Left            =   75
         TabIndex        =   11
         Top             =   1275
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transact. Date"
         Height          =   285
         Index           =   1
         Left            =   75
         TabIndex        =   6
         Top             =   675
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name"
         Height          =   285
         Index           =   3
         Left            =   75
         TabIndex        =   9
         Top             =   975
         Width           =   1200
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   1350
         Tag             =   "et0;ht2"
         Top             =   300
         Width           =   1950
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   525
      Index           =   1
      Left            =   120
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   10185
      _ExtentX        =   17965
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
         Index           =   19
         Left            =   1275
         TabIndex        =   1
         Top             =   90
         Width           =   1950
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
         Index           =   20
         Left            =   4185
         MaxLength       =   50
         TabIndex        =   3
         Top             =   90
         Width           =   5820
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "S.I. &Number"
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
         Index           =   13
         Left            =   90
         TabIndex        =   0
         Top             =   135
         Width           =   1140
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
         Index           =   0
         Left            =   3300
         TabIndex        =   2
         Top             =   135
         Width           =   840
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   10530
      TabIndex        =   36
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
      Picture         =   "frmCP_POSReg.frx":001C
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   10545
      TabIndex        =   37
      Top             =   1260
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Void"
      AccessKey       =   "V"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCP_POSReg.frx":0796
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   10560
      TabIndex        =   38
      Top             =   1890
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Receipt"
      AccessKey       =   "R"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCP_POSReg.frx":0F10
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   10560
      TabIndex        =   39
      Top             =   3780
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
      Picture         =   "frmCP_POSReg.frx":168A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   10560
      TabIndex        =   40
      Top             =   2520
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Print"
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
      Picture         =   "frmCP_POSReg.frx":1E04
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   10560
      TabIndex        =   41
      Top             =   3150
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&J.O"
      AccessKey       =   "J"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCP_POSReg.frx":257E
   End
End
Attribute VB_Name = "frmCP_POSReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmCP_POSReg"

Private WithEvents oTrans As clsCPSales
Attribute oTrans.VB_VarHelpID = -1
'Private oFormGiveAway As frmGiveaway

Private oReceipt As ggcCPSales.Receipt
Private oSkin As clsFormSkin

Dim pbGridFocus As Boolean
Dim pbMoveCombo As Boolean
Dim pbHsSerial As Boolean
Dim pnIndex As Integer
Dim pnCtr As Integer

Dim pbSave As Boolean
Dim pbLoad As Boolean
Dim psUserName As String
Dim psUserIDxx As String
Dim psTransNox As String
Dim pnTtlAdj As Currency

Property Let TransNox(lsTransNox As String)
   psTransNox = lsTransNox
End Property

Private Sub chkBox_Click(Index As Integer)
'   If pbLoad = True Then Exit Sub
'   If Index = 1 Then
'      oTrans.Master("sAdvRefer") = chkBox(Index).Value
'      lblAdvPayment.Caption = Format(oTrans.Master("nAdvPaymx"), "#,##0.00")
'   End If
End Sub

Private Sub cmbField_Click(Index As Integer)
   With cmbField(Index)
      If .ListIndex < 0 Then .ListIndex = -1
      oTrans.Master(IIf(Index = 0, 14, 13)) = cmbField(Index).ListIndex

      If Index = 1 Then
         If cmbField(Index).ListIndex <> 3 Or cmbField(Index).ListIndex <> 4 Then
'            txtField(5).Text = ""
            oTrans.Master("sTermIDxx") = ""
         End If
'         txtField(5).Enabled = cmbField(Index).ListIndex = 3 Or cmbField(Index).ListIndex = 4
      End If
   End With
End Sub

Private Sub cmbField_GotFocus(Index As Integer)
   With cmbField(Index)
      .BackColor = oApp.getColor("HT1")
   End With

   pbMoveCombo = True
End Sub

Private Sub cmbField_LostFocus(Index As Integer)
   With cmbField(Index)
      .BackColor = oApp.getColor("EB")
   End With
   pbMoveCombo = False
End Sub

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   Dim lnRep As Integer
   Dim lsAppvID As String
   Dim lsAppvName As String
   Dim lnAppvRights As Integer
   Dim lbGetApproval As Boolean
   
   lsOldProc = "cmdButton_Click"
   On Error GoTo errProc
   
   lnAppvRights = oApp.UserLevel
   lsAppvID = oApp.UserID
   lbGetApproval = False

   txtField_LostFocus pnIndex
   With GridEditor1
      Select Case Index
      Case 0 'browse
         If oTrans.SearchTransaction() Then
            LoadMaster
            LoadDetail
         Else
            If txtField(0).Text = "" Then ClearFields
         End If

         txtField(pnIndex).SetFocus
         .Refresh
      Case 1 'Receipt
         If CDbl(txtField(5).Text) + CDbl(lblAdvPayment) > 0# And _
            CDbl(txtField(9).Text) <> 0# And _
            txtField(3).Text <> "" Then
            Call Receipt
         Else
            MsgBox "Unable to Load Receipt!!!" & vbCrLf & _
                   "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
         End If
      Case 2 'Close
         Unload Me
      Case 3
         If DateDiff("d", oApp.ServerDate, oTrans.Master("dModified")) = 0 Then
            If oApp.UserLevel >= xeEncoder Then
               If oTrans.Master("cTranStat") = xeStateOpen Or oTrans.Master("cTranStat") = xeStateClosed Then
                  If oTrans.CancelTransaction = True Then
                     MsgBox "Transaction Cancelled Successfully!", vbInformation
                  Else
                     MsgBox "Unable to Cancel Transaction!", vbCritical, "Warning"
                  End If
               End If
            End If
         Else
            If oApp.UserLevel > xeManager Then
               If oTrans.Master("cTranStat") = xeStateOpen Or oTrans.Master("cTranStat") = xeStateClosed Then
                  If oTrans.CancelTransaction = True Then
                     MsgBox "Transaction Cancelled Successfully!", vbInformation
                  Else
                     MsgBox "Unable to Cancel Transaction!", vbCritical, "Warning"
                  End If
               End If
            Else
               MsgBox "Unable to Cancel Transaction!", vbCritical, "Warning"
            End If
         End If
      Case 4 'Print
         If oTrans.Master("cTranStat") = xeStateOpen Then
            lnRep = MsgBox("Print Transaction!!!", vbYesNo + vbQuestion, "Confirm")
            If lnRep = vbYes Then
               If oTrans.Master("sJobOrder") <> "" Then
                  If Not PrintJOSI Then MsgBox "Unable to Print Transaction!!!", vbCritical, "Warning"
               Else
                  If Not PrintTrans Then MsgBox "Unable to Print Transaction!!!", vbCritical, "Warning"
               End If
            End If
         ElseIf oTrans.Master("cTranStat") = xeStateClosed Then
            If oApp.UserLevel <> xeManager Or oApp.UserLevel <> xeEngineer Then lbGetApproval = True
            
            If lbGetApproval Then
               If GetApproval(oApp, lnAppvRights, lsAppvID, lsAppvName, oApp.MenuName) = False Then GoTo endProc
               '2015-04-11
               '(she) change lnAppvRights to manager requested by tatay rex
               If lnAppvRights < xeManager Then GoTo endProc
            End If
            
            If oTrans.Master("sJobOrder") <> "" Then
               If Not PrintJOSI Then MsgBox "Unable to Print Transaction!!!", vbCritical, "Warning"
            Else
               If Not PrintTrans Then MsgBox "Unable to Print Transaction!!!", vbCritical, "Warning"
            End If
         Else
            MsgBox "Unable to Print Transaction!!!" & vbCrLf & _
                   "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
         End If
         Call LoadMaster
         Call LoadDetail
      Case 5 'JO
         If oTrans.Master("sJobOrder") <> "" Then
            Call LoadJO
         End If
      End Select
   End With

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0
   
   If psTransNox <> "" Then
      If oTrans.OpenTransaction(psTransNox) Then
         Call LoadMaster
         Call LoadDetail
         Call cmdButton_Click(5)
      End If
   End If
   
   With GridEditor1
      .Refresh
   End With
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String

   lsOldProc = "Form_Load"
    On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oTrans = New clsCPSales
   Set oTrans.AppDriver = oApp

   oTrans.Branch = oApp.BranchCode
   oTrans.InitTransaction
   oTrans.NewTransaction

   Set oReceipt = New ggcCPSales.Receipt
   Set oReceipt.AppDriver = oApp

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualRight

   InitForm
   ClearFields
   initButton xeModeReady

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oReceipt = Nothing
   Set oTrans = Nothing
   Set oSkin = Nothing
End Sub

Private Sub GridEditor1_EditorValidate(Cancel As Boolean)
   Dim lnPercent As Integer
   Dim lnDiscount As Variant
   
   With GridEditor1
      Select Case .Col
      Case 3, 4, 5, 6
         If .Col = 5 Then
            If Not IsNumeric(lnDiscount) Then
               .TextMatrix(.Row, .Col) = 0
            Else
               lnDiscount = .TextMatrix(.Row, .Col)
               lnPercent = InStr(lnDiscount, "%")
               If lnPercent > 0 Then lnDiscount = Left(lnDiscount, lnPercent - 1)
   
               If lnDiscount > 99 Then lnDiscount = 0
            End If
            .TextMatrix(.Row, .Col) = lnDiscount & "%"
         Else
            oTrans.Detail(.Row - 1, .Col) = CDbl(.TextMatrix(.Row, .Col))
         End If
         .TextMatrix(.Row, 7) = CDbl(.TextMatrix(.Row, 3)) * CDbl(.TextMatrix(.Row, 4))
         .TextMatrix(.Row, 7) = (CDbl(.TextMatrix(.Row, 7)) * _
                                    (100 - CDbl(Left(.TextMatrix(.Row, 5), _
                                       Len(.TextMatrix(.Row, 5)) - 1))) / 100) - _
                                    CDbl(.TextMatrix(.Row, 6))
         ComputeTotal
      Case Else
         oTrans.Detail(.Row - 1, .Col) = .TextMatrix(.Row, .Col)
      End Select
   End With
End Sub

Private Sub GridEditor1_AddingRow(Cancel As Boolean)
   With GridEditor1
      If .TextMatrix(.Row, 1) = "" Then
         Cancel = True
      ElseIf .TextMatrix(.Row, 7) = "0" Then
         If cmbField(1).ListIndex = 0 Then Cancel = True
      Else
         If Not Cancel Then oTrans.addDetail
      End If
   End With
End Sub

Private Sub GridEditor1_GotFocus()
   With GridEditor1
      .EditorBackColor = oApp.getColor("HT1")
   End With
   pbGridFocus = True
End Sub

Private Sub GridEditor1_LostFocus()
   With GridEditor1
      .Col = 1
      .TopRow = 1
      .LeftCol = 1
      .EditorBackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub oTrans_MasterRetrieved(ByVal Index As Integer)
   Select Case Index
   Case 9
      txtField(Index).Text = Format(oTrans.Master(Index), "#,##0.00")
   Case 11, 12
   Case Else
      txtField(Index).Text = oTrans.Master(Index)
   End Select
End Sub

Private Sub oTrans_DetailRetrieved(ByVal Index As Integer)
   Dim lnSubTotal As Currency
   
   With GridEditor1
      If Index = 5 Then
         lnSubTotal = CDbl(oTrans.Detail(.Row - 1, 3)) * CDbl(oTrans.Detail(.Row - 1, 4))
         .TextMatrix(.Row, 7) = lnSubTotal - (lnSubTotal * _
                                   (oTrans.Detail(.Row - 1, 5) / 100) - oTrans.Detail(.Row - 1, 6))
      
         .TextMatrix(.Row, Index) = Format(oTrans.Detail(.Row - 1, Index), "#0.00") & "%"
      Else
         If Index <> 7 Then .TextMatrix(.Row, Index) = oTrans.Detail(.Row - 1, Index)
      End If
      Call ComputeTotal
   End With
End Sub

Private Sub InitForm()
   With GridEditor1
      .Rows = 2
      .Cols = 8
      .Font = "MS Sans Serif"

      'column title
      .TextMatrix(0, 1) = "BarrCode"
      .TextMatrix(0, 2) = "Description"
      .TextMatrix(0, 3) = "Qty"
      .TextMatrix(0, 4) = "Unit Price"
      .TextMatrix(0, 5) = "Dsc Rte"
      .TextMatrix(0, 6) = "Dsc Amt"
      .TextMatrix(0, 7) = "Total"

      .Row = 0

      'column alignment
      For pnCtr = 0 To .Cols - 1
         .Col = pnCtr
         .CellFontBold = True
         .CellAlignment = 3
      Next

      'column width
      .ColWidth(0) = 330
      .ColWidth(1) = 2030
      .ColWidth(3) = 500
      .ColWidth(4) = 1000
      .ColWidth(5) = 800
      .ColWidth(6) = 800
      .ColWidth(7) = 1000
      .ColDefault(7) = 0
      

      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 6
      .ColAlignment(4) = 6
      .ColAlignment(5) = 6
      .ColAlignment(6) = 6
      .ColAlignment(7) = 6
      
      .ColFormat(4) = "#,##0.00"
      .ColFormat(6) = "#,##0.00"
      .ColFormat(7) = "#,##0.00"
      .ColNumberOnly(7) = True
      

      .Row = 1
      .Col = 1
   End With

'   cmbField(0).List(0) = "Gift"
'   cmbField(0).List(1) = "Sales"
'   cmbField(0).List(2) = "Replacement"

   cmbField(1).List(0) = "Cash"
   cmbField(1).List(1) = "Cash Balance"
   cmbField(1).List(2) = "Installment"
   cmbField(1).List(3) = "Term"

   txtField(2).MaxLength = oTrans.MasFldSize(2)
   txtField(8).MaxLength = oTrans.MasFldSize(8)
   txtField(9).MaxLength = oTrans.MasFldSize(9)
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      If Index = 1 Then .Text = Format(.Text, "MM/DD/YYYY")
      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = oApp.getColor("HT1")
   End With

   pbGridFocus = False
   pnIndex = Index
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String

   lsOldProc = "txtField_KeyDown"
    On Error GoTo errProc

   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
      With txtField(Index)
         Select Case Index
         Case 3
            oTrans.Master(Index) = .Text
         Case 8, 17
            If KeyCode = vbKeyF3 Then
               oTrans.SearchMaster Index, .Text
               If .Text <> "" Then SetNextFocus
            Else
               If .Text <> "" Then oTrans.SearchMaster Index, .Text
            End If
         Case 19, 20
            Call txtField_Validate(Index, False)
         End Select
      End With
      KeyCode = 0
   End If

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index _
                       & ", " & KeyCode _
                       & ", " & Shift & " )", True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim lnReply As Integer
   Dim lsApprvID As String
   Dim lsApprvNm As String
   
   Select Case KeyCode
   Case vbKeyReturn, vbKeyUp, vbKeyDown
      If pbMoveCombo And KeyCode <> vbKeyReturn Then
         Exit Sub
      End If

      Select Case KeyCode
      Case vbKeyReturn, vbKeyDown
         If GetFocus = GridEditor1.hwnd Then Exit Sub
         SetNextFocus
      Case vbKeyUp
         SetPreviousFocus
      End Select
   Case vbKeyF8
      If txtField(0).Text <> "" And oTrans.EditMode = xeModeReady Then
         If DateDiff("d", oApp.ServerDate, oTrans.Master("dModified")) = 0 And oTrans.Master("cTranStat") = xeStateOpen Then
            If oApp.UserLevel >= xeManager Then
               If oTrans.DeleteTransaction Then ClearFields
            End If
         Else
            MsgBox "Unable to Delete Transaction!", vbCritical, "Warning"
            Exit Sub
         End If
      End If
   Case vbKeyF12
      oTrans.ViewModify
   End Select
End Sub

Private Sub initButton(lnStat As Integer)
   Dim lbShow As Boolean

   lbShow = IIf(lnStat = xeModeReady, False, True)
   xrFrame1(1).Enabled = Not lbShow
'   cmdButton(4).Visible = Not lbShow
'   cmdButton(5).Visible = Not lbShow
'   cmdButton(6).Visible = Not lbShow
'   cmdButton(8).Visible = Not lbShow
'
'   cmdButton(0).Visible = lbShow
'   cmdButton(1).Visible = lbShow
'   cmdButton(2).Visible = lbShow
'   cmdButton(3).Visible = lbShow

'   cmbField(0).Enabled = lbShow
   cmbField(1).Enabled = lbShow

   '   chkBox(0).Enabled = lbShow
   chkBox(1).Enabled = lbShow

   With GridEditor1
      For pnCtr = 1 To .Cols - 1
         .ColEnabled(pnCtr) = lbShow
      Next
   End With
End Sub

Private Function ComputeTotal() As Double
   Dim lnCtr As Integer
   Dim lnSum As Double

   lnSum = 0#
   With GridEditor1
      For lnCtr = 1 To .Rows - 1
         lnSum = CDbl(.TextMatrix(lnCtr, 7)) + lnSum
      Next
   End With

   txtField(9).Text = Format(lnSum, "#,##0.00")
'   txtField(5).Text = txtField(9).Text

   oTrans.Master("nAmtPaidx") = CDbl(txtField(5).Text)
   oTrans.Master("nTranTotl") = CDbl(txtField(9).Text)
End Function

Private Sub GridEditor1_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String

   lsOldProc = "GridEditor1_KeyDown"
    On Error GoTo errProc

   If KeyCode = vbKeyF3 Then
      With GridEditor1
         Select Case .Col
         Case 1, 2
            If oTrans.searchDetail(.Row - 1, .Col, .TextMatrix(.Row, .Col)) Then
               .Col = 3
            Else
               .TextMatrix(.Row, .Col) = oTrans.Detail(.Row - 1, .Col)
            End If

            .SetFocus
            .Refresh
         End Select
      End With
      KeyCode = 0
   End If

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " _
                       & "  " & KeyCode _
                       & ", " & Shift _
                       & " )", True
End Sub

Private Sub ClearFields()
   For pnCtr = 0 To 20
      Select Case pnCtr
      Case 1
         txtField(pnCtr).Text = Format(oTrans.Master("dTransact"), "MMMM DD, YYYY")
      Case 5, 9
         txtField(pnCtr).Text = "0.00"
      Case 10 To 16
      Case Else
         txtField(pnCtr).Text = ""
         txtField(pnCtr).Tag = txtField(pnCtr).Text
      End Select
   Next

   lblAdvPayment.Caption = "0.00"

'   cmbField(0).ListIndex = oTrans.Master("cTranType")
'   cmbField(1).ListIndex = oTrans.Master("cPaymForm")

'   chkBox(0).Value = oTrans.Master("cCustType")
'   chkBox(1).Value = oTrans.Master("sAdvRefer")

   With GridEditor1
      .LeftCol = 1
      .TopRow = 1
      .Rows = 2
      .Row = 1
      .Col = 1

      .TextMatrix(1, 1) = ""
      .TextMatrix(1, 2) = ""
      .TextMatrix(1, 3) = 0
      .TextMatrix(1, 4) = 0#
      .TextMatrix(1, 5) = "0.00%"
      .TextMatrix(1, 6) = 0#
      .TextMatrix(1, 7) = 0#

      .ColWidth(2) = 3400
   End With

   pbSave = False
   pbLoad = False
   oReceipt.InitReceipt
End Sub

Private Function isEntryOk() As Boolean
   Dim lnCtr As Integer
   Dim lsUserID As String, lsUserName As String, lsOldProc As String
   Dim lnUserRights As Integer, lnRep As String

'   If txtField(3).Text = "" Then
'      MsgBox "Customer not found!!!", vbCritical, "Warning"
'      txtField(3).SetFocus
'      GoTo endProc
'   End If
   
   ' if sales transaction, check payment vs payment type
   Select Case cmbField(1).ListIndex
   Case 0   ' Cash Sales
      If CDbl(txtField(5).Text) + CDbl(lblAdvPayment) <> CDbl(txtField(9).Text) Then
         MsgBox "Invalid amount paid!!!" & vbCrLf & _
            "Transaction must fit from payment type!!!", vbCritical, "Warning"
         txtField(5).SetFocus
         GoTo endProc
      End If
   Case 1, 2   ' 1 - Cash Balance; 2 - Installment
      If CDbl(txtField(9).Text) <= CDbl(txtField(5).Text) + CDbl(lblAdvPayment) Or _
         CDbl(txtField(5).Text) + CDbl(lblAdvPayment) = 0# Then
         MsgBox "Invalid amount paid!!!" & vbCrLf & _
            "Transaction must fit from payment type!!!", vbCritical, "Warning"
         txtField(5).SetFocus
         GoTo endProc
      End If

      If cmbField(1).ListIndex = 2 Then

      End If
   Case 3 ' Term
      If txtField(5).Text = "" Then
         MsgBox "Invalid Term Detected!!!" & vbCrLf & _
               "Please Verify your entry then try again!!!", vbCritical, "Warning"
         txtField(5).SetFocus
         GoTo endProc
      End If

      If CDbl(txtField(5).Text) + CDbl(lblAdvPayment.Caption) > 0 Then
         MsgBox "Invalid amount paid!!!" & vbCrLf & _
               "Transaction must fit from payment type!!!", vbCritical, "Warning"
         txtField(5).SetFocus
         GoTo endProc
      End If
   End Select

   ' just include this code, coz the sales object create an entry to
   '  the application no despite the non-existence of an application
'   If cmbField(1).ListIndex = 1 And txtField(10).Text = "" Or _
'      cmbField(1).ListIndex = 2 And txtField(10).Text = "" Then
'      MsgBox "Invalid Application No!!!", vbCritical, "Warning"
'      txtField(10).SetFocus
'      GoTo endProc
'   ElseIf cmbField(1).ListIndex = 0 Then
'      txtField(10).Text = ""
'      oTrans.Master("sApplicNo") = ""
'   End If

   With GridEditor1
      If .TextMatrix(1, 1) = "" Then
         MsgBox "Detail is required!!!", vbCritical, "Warning"
         .SetFocus
         .Row = 1
         .Col = 1
         GoTo endProc
      End If
   End With

'   Select Case cmbField(0).ListIndex
'   Case 0, 2
'      ' check the amount paid vs the type of transaction
'      '  0 - for gift
'      '  2 - for replacement
'      If CDbl(txtField(5).Text) > CDbl(txtField(12).Text) Then
'         MsgBox "Invalid amount paid!!!" & vbCrLf & _
'            "Transaction must fit from payment type!!!", vbCritical, "Warning"
'         cmbField(0).SetFocus
'         GoTo endProc
'      End If
'   Case 1
'      '  1 - for sales
'      '      sales must have a payment
'      If CDbl(txtField(9).Text) = 0# And CDbl(lblAdvPayment) = 0 Then
'         MsgBox "Invalid amount paid!!!" & vbCrLf & _
'            "Transaction must fit from payment type!!!", vbCritical, "Warning"
'         cmbField(0).SetFocus
'         GoTo endProc
'      End If
'
'      If CDbl(txtField(5).Text) + CDbl(lblAdvPayment) > CDbl(txtField(9).Text) Then
'         MsgBox "Invalid amount paid!!!" & vbCrLf & _
'            "Transaction must fit from payment type!!!", vbCritical, "Warning"
'         cmbField(1).SetFocus
'         GoTo endProc
'      End If
'
'      ' if sales transaction, check payment vs payment type
'      Select Case cmbField(1).ListIndex
'      Case 0   ' Cash Sales
'         If CDbl(txtField(5).Text) + CDbl(lblAdvPayment) <> CDbl(txtField(9).Text) Then
'            MsgBox "Invalid amount paid!!!" & vbCrLf & _
'               "Transaction must fit from payment type!!!", vbCritical, "Warning"
'            txtField(5).SetFocus
'            GoTo endProc
'         End If
'      Case 1, 2   ' 1 - Cash Balance; 2 - Installment
'         If CDbl(txtField(9).Text) <= CDbl(txtField(5).Text) + CDbl(lblAdvPayment) Or _
'            CDbl(txtField(5).Text) + CDbl(lblAdvPayment) = 0# Then
'            MsgBox "Invalid amount paid!!!" & vbCrLf & _
'               "Transaction must fit from payment type!!!", vbCritical, "Warning"
'            txtField(5).SetFocus
'            GoTo endProc
'         End If
'
'         If cmbField(1).ListIndex = 2 Then
'
'         End If
'      Case 3 ' Term
'         If txtField(5).Text = "" Then
'            MsgBox "Invalid Term Detected!!!" & vbCrLf & _
'                  "Please Verify your entry then try again!!!", vbCritical, "Warning"
'            txtField(5).SetFocus
'            GoTo endProc
'         End If
'
'         If CDbl(txtField(5).Text) + CDbl(lblAdvPayment.Caption) > 0 Then
'            MsgBox "Invalid amount paid!!!" & vbCrLf & _
'                  "Transaction must fit from payment type!!!", vbCritical, "Warning"
'            txtField(5).SetFocus
'            GoTo endProc
'         End If
'      Case 4 ' Term With Registration
'         If txtField(8).Text = "" Then
'            MsgBox "Invalid Term Detected!!!" & vbCrLf & _
'                  "Please Verify your entry then try again!!!", vbCritical, "Warning"
'            txtField(8).SetFocus
'            GoTo endProc
'         End If
'      End Select
'   End Select
   isEntryOk = True

endProc:
   Exit Function
End Function

'Private Function Receipt() As Boolean
'   Dim lnCheckAmt As Currency
'   Dim lnCashAmtx As Currency
'   Dim lnCardAmtx As Currency
'   Dim lnTotalAmt As Currency
'   Dim lsOldProc As String
'
'   lsOldProc = "Receipt"
'   On Error GoTo errProc
'
'   With oReceipt
'      lnCheckAmt = oTrans.Receipt("nCheckAmt")
'      lnCardAmtx = oTrans.Receipt("nCardAmtx")
'
'      .Client = oTrans.Client
'      .Client.Master("sCompnyNm") = oTrans.Master("xFullName")
'      .Client.Master("sClientId") = oTrans.Master("sClientID")
'      .UserName = oTrans.Master("xSalesman")
'      .CashAmount = oTrans.Master("nCashAmtx")
'      .AmountPaid = oTrans.Master("nTranTotl")
'      .InvoiceDate = oTrans.Master("dTransact")
'      .Remarks = oTrans.Master("sRemarksx")
'      .InvoiceNo = oTrans.Master("sSalesInv")
'      .ORNo = oTrans.Master("sORNoxxxx")
'
'      If lnCheckAmt > 0 Then
'         .Checks("sBankIDxx") = oTrans.Checks("sBankIDxx")
'         .Checks("sCheckNox") = oTrans.Checks("sCheckNox")
'         .Checks("dCheckDte") = oTrans.Checks("dTransact")
'         .Checks("nAmountxx") = oTrans.Checks("nAmountxx")
'         .Checks("sAcctNoxx") = oTrans.Checks("sAcctNoxx")
'         .Checks("sBankName") = oTrans.Checks("sBankName")
'      Else
'         .Checks("sBankIDxx") = ""
'         .Checks("sCheckNox") = ""
'         .Checks("dCheckDte") = oTrans.Master("dTransact")
'         .Checks("nAmountxx") = 0#
'         .Checks("sAcctNoxx") = ""
'         .Checks("sBankName") = ""
'      End If
'
'      If lnCardAmtx > 0 Then
'         .Cards("sBankIDxx") = oTrans.Card(pnCtr, "sBankIDxx")
'         .Cards("sCardIDxx") = oTrans.Card(pnCtr, "sCardIDxx")
'         .Cards("sCardNoxx") = oTrans.Card(pnCtr, "sCardNoxx")
'         .Cards("sApproval") = oTrans.Card(pnCtr, "sApproval")
'         .Cards("nCardAmtx") = oTrans.Card(pnCtr, "nCardAmtx")
'      Else
'         .Cards("sBankIDxx") = ""
'         .Cards("sCardIDxx") = ""
'         .Cards("sCardNoxx") = ""
'         .Cards("sApproval") = ""
'         .Cards("nCardAmtx") = 0#
'         .Cards("sBankName") = ""
'         .Cards("sCardName") = ""
'      End If
'      .HasSerial = pbHsSerial
'      .ShowReceipt
'
'      If Not .Cancelled Then
'         If .CheckAmount > 0 Then
'            oTrans.Checks("sBankIDxx") = .Checks("sBankIDxx")
'            oTrans.Checks("sCheckNox") = .Checks("sCheckNox")
'            oTrans.Checks("dCheckDte") = .Checks("dCheckDte")
'            oTrans.Checks("nAmountxx") = .Checks("nAmountxx")
'            oTrans.Checks("sAcctNoxx") = .Checks("sAcctNoxx")
'         End If
'
'         If .CardAmount > 0 Then
'            oTrans.Card("sBankIDxx") = .Cards("sBankIDxx")
'            oTrans.Card("sCardIDxx") = .Cards("sCardIDxx")
'            oTrans.Card("sCardNoxx") = .Cards("sCardNoxx")
'            oTrans.Card("sApproval") = .Cards("sApproval")
'            oTrans.Card("nAmountxx") = .Cards("nAmountxx")
'         End If
'
'         oTrans.Receipt("nTranTotl") = oTrans.Master("nTranTotl")
'         oTrans.Receipt("nCashAmtx") = oTrans.Master("nCashAmtx")
'         oTrans.Receipt("nCheckAmt") = oTrans.Checks("nAmountxx")
'         oTrans.Receipt("nCardAmtx") = oTrans.Cards("nAmountxx")
'         oTrans.Receipt("sRemarksx") = oTrans.Master("sRemarksx")
'
'         oTrans.Master("sSalesman") = .UserID
'         oTrans.Master("sRemarksx") = .Remarks
'         oTrans.Master("nCashAmtx") = .CashAmount
'         oTrans.Master("nAmtPaidx") = .AmountPaid
'         oTrans.Master("sSalesInv") = .InvoiceNo
'         oTrans.Master("dTransact") = .InvoiceDate
'         oTrans.Master("sORNoxxxx") = .ORNo
'
'         txtField(2).Text = oTrans.Master("sSalesInv")
'         txtField(3).Text = oReceipt.Client.Master("sCompnyNm")
'         txtField(17).Text = .UserName
'
'         oTrans.Client = oReceipt.Client
'         oTrans.Master("sClientID") = oReceipt.Client("sClientId")
'         psUserName = .UserName
'         psUserIDxx = .UserID
'         Receipt = True
'      End If
'   End With
'
'endProc:
'   Exit Function
'errProc:
'   ShowError lsOldProc & "( " & " )", True
'End Function

Private Function Receipt() As Boolean
   Dim lnCheckAmt As Currency
   Dim lnCashAmtx As Currency
   Dim lnCardAmtx As Currency
   Dim lnTotalAmt As Currency
   Dim lsOldProc As String

   lsOldProc = "Receipt"
   'On Error GoTo errProc

   With oReceipt
      lnCheckAmt = oTrans.Receipt("nCheckAmt")
      lnCardAmtx = oTrans.Receipt("nCardAmtx")
      .Client = oTrans.Client
      .Sales = oTrans
      .UserName = psUserName
      .UserID = psUserIDxx
      .CashAmount = oTrans.Master("nCashAmtx")
      .AmountPaid = oTrans.Master("nTranTotl")
      .InvoiceDate = oTrans.Master("dTransact")
      .UserID = oTrans.Master("sSalesman")
      .Remarks = oTrans.Master("sRemarksx")
      .InvoiceNo = oTrans.Master("sSalesInv")
      .ORNo = oTrans.Master("sORNoxxxx")

      If lnCheckAmt > 0 Then
         .Checks("sBankIDxx") = oTrans.Checks("sBankIDxx")
         .Checks("sCheckNox") = oTrans.Checks("sCheckNox")
         .Checks("dCheckDte") = oTrans.Checks("dTransact")
         .Checks("nAmountxx") = oTrans.Checks("nAmountxx")
         .Checks("sAcctNoxx") = oTrans.Checks("sAcctNoxx")
      Else
         .Checks("sBankIDxx") = ""
         .Checks("sCheckNox") = ""
         .Checks("dCheckDte") = oTrans.Master("dTransact")
         .Checks("nAmountxx") = 0#
         .Checks("sAcctNoxx") = ""
      End If

      If lnCardAmtx > 0 Then
         .Cards("sBankIDxx") = oTrans.Card(pnCtr, "sBankIDxx")
         .Cards("sCardIDxx") = oTrans.Card(pnCtr, "sCardIDxx")
         .Cards("sCardNoxx") = oTrans.Card(pnCtr, "sCardNoxx")
         .Cards("sApproval") = oTrans.Card(pnCtr, "sApproval")
         .Cards("nCardAmtx") = oTrans.Card(pnCtr, "nCardAmtx")
      Else
         .Cards("sBankIDxx") = ""
         .Cards("sCardIDxx") = ""
         .Cards("sCardNoxx") = ""
         .Cards("sApproval") = ""
         .Cards("nCardAmtx") = 0#
      End If
      .HasSerial = pbHsSerial
      .ShowReceipt

      If Not .Cancelled Then
         txtField(4).Text = .UserName
         oTrans.Receipt("nCheckAmt") = 0#
         oTrans.Receipt("nCardAmtx") = 0#
         
         If .CheckAmount > 0 Then
            oTrans.Checks("sBankIDxx") = .Checks("sBankIDxx")
            oTrans.Checks("sCheckNox") = .Checks("sCheckNox")
            oTrans.Checks("dCheckDte") = .Checks("dCheckDte")
            oTrans.Checks("nAmountxx") = .Checks("nAmountxx")
            oTrans.Checks("sAcctNoxx") = .Checks("sAcctNoxx")
            
            oTrans.Receipt("nCheckAmt") = oTrans.Checks("nAmountxx")
         End If

         If .CardAmount > 0 Then
            oTrans.Card(pnCtr, "sBankIDxx") = .Cards("sBankIDxx")
            oTrans.Card(pnCtr, "sCardIDxx") = .Cards("sCardIDxx")
            oTrans.Card(pnCtr, "sCardNoxx") = .Cards("sCardNoxx")
            oTrans.Card(pnCtr, "sApproval") = .Cards("sApproval")
            oTrans.Card(pnCtr, "nAmountxx") = .Cards("nAmountxx")
            
            oTrans.Receipt("nCardAmtx") = oTrans.Card(pnCtr, "nAmountxx")
         End If

         oTrans.Receipt("nTranTotl") = oTrans.Master("nTranTotl")
         oTrans.Receipt("nCashAmtx") = oTrans.Master("nCashAmtx")
         oTrans.Receipt("sRemarksx") = oTrans.Master("sRemarksx")

         oTrans.Master("sSalesman") = .UserID
         oTrans.Master("sRemarksx") = .Remarks
         oTrans.Master("nCashAmtx") = .CashAmount
         oTrans.Master("nAmtPaidx") = .AmountPaid
         oTrans.Master("sSalesInv") = .InvoiceNo
         oTrans.Master("dTransact") = .InvoiceDate
         oTrans.Master("sORNoxxxx") = .ORNo

         txtField(1).Text = oTrans.Master("sSalesInv")
         txtField(7).Text = oReceipt.Client.Master("sCompnyNm")

         oTrans.Client = oReceipt.Client
         oTrans.Master("sClientID") = oReceipt.Client.Master("sClientID")
         oTrans.Master("xFullName") = oReceipt.Client.Master("sCompnyNm")
         oTrans.Master("xAddressx") = oReceipt.Client.Master("sAddressx") & ", " & oReceipt.Client.Master("sTownName") & ", " & oReceipt.Client.Master("sProvName") & " " & oReceipt.Client.Master("sZippCode")

         psUserName = .UserName
         psUserIDxx = .UserID
         Receipt = True
      End If
   End With

endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )", True
End Function

Private Function getGiveAways() As String
   Dim lrs As Recordset
   Dim lsSQL As String
   
   lsSQL = "SELECT" & _
               "  a.sStockIDx" & _
               ", b.sDescript" & _
               ", a.nQuantity" & _
               ", a.nGivenxxx" & _
            " FROM CP_SO_GiveAways a" & _
               " LEFT JOIN CP_Inventory b" & _
                  " ON a.sStockIDx = b.sStockIDx" & _
            " WHERE a.cGAwyStat <> '2'" & _
               " AND a.sTransNox = " & strParm(oTrans.Master("sTransNox")) & _
            " ORDER BY a.nGivenxxx DESC"
   
   
      Set lrs = New Recordset
      lrs.Open lsSQL, oApp.Connection, adOpenStatic, adLockOptimistic, adCmdText
      Set lrs.ActiveConnection = Nothing
      
      With lrs
      lsSQL = ""
      
      If .EOF Then GoTo endProc
      
      lsSQL = "Giveaways: "
      .MoveFirst
      Do Until .EOF
         If lrs("nGivenxxx") = 1 Then
            lsSQL = lsSQL & lrs("sDescript") & "; "
         End If
         .MoveNext
      Loop
      
      lsSQL = "To Follow: "
      .MoveFirst
      Do Until .EOF
         If lrs("nGivenxxx") = 0 Then
            lsSQL = lsSQL & lrs("sDescript") & "; "
         End If
         .MoveNext
      Loop
   End With
endProc:
   getGiveAways = lsSQL
End Function

Private Function getAccesories() As String
   Dim lrs As Recordset
   Dim lsSQL As String
   
   
   lsSQL = "SELECT" & _
               "  b.sDescript" & _
               ", a.sSerialNo" & _
            " FROM CP_SO_Accessories a" & _
               " LEFT JOIN CP_Accessories b" & _
                  " ON a.sAccessID = b.sAccessID" & _
            " WHERE a.sTransNox = " & strParm(oTrans.Master("sTransNox"))
   
   Set lrs = New Recordset
   lrs.Open lsSQL, oApp.Connection, adOpenStatic, adLockOptimistic, adCmdText
   Set lrs.ActiveConnection = Nothing
   
   lsSQL = ""
   
   If lrs.EOF Then GoTo endProc
   
   lsSQL = "Accessories: "
   Do Until lrs.EOF
      lsSQL = lsSQL & lrs("sDescript") & " - " + lrs("sSerialNo") & "; "
               
      lrs.MoveNext
   Loop
   
endProc:
   getAccesories = lsSQL
End Function

Function PrintTrans() As Boolean
      Dim lrs As New ADODB.Recordset
   Dim lnCtr As Integer
   Dim lsIMEI As String
   Dim lReplAmt As String
   Dim lsOldProc As String
   Dim loModel As Recordset
   
   Dim lsSQL As String
   Dim lrsCOInv As Recordset
   Dim lnFinAmt As Currency
   Dim lsFinTrans As String
   Dim lnTotalDisc As Currency 'she temporary for mobile fest
   Dim lnSelPrice As Double
   
   lsOldProc = "printTrans"
   ''On Error GoTo errProc
   
   PrintTrans = False
   
   Set lrs = New ADODB.Recordset
   
   lrs.Fields.Append "nField01", adInteger, 3
   lrs.Fields.Append "sField01", adVarChar, 50
   lrs.Fields.Append "sField02", adVarChar, 60
   lrs.Fields.Append "sField03", adVarChar, 30
   lrs.Fields.Append "lField01", adCurrency
   lrs.Fields.Append "lField02", adCurrency
   lrs.Fields.Append "lField03", adCurrency
   lrs.Fields.Append "lField04", adCurrency
   lrs.Fields.Append "lField05", adCurrency
   lrs.Fields.Append "lField06", adCurrency
   lrs.Open
      
   lsSQL = "SELECT a.sTransNox" & _
            ", b.sBarrCode" & _
            ", c.sModelNme" & _
            ", c.sModelCde" & _
            ", d.sColorNme" & _
            ", a.nUnitPrce" & _
            ", a.nDiscRate" & _
            ", a.nQuantity" & _
            ", b.sDescript" & _
            ", b.cHsSerial" & _
            ", a.sSerialID" & _
            ", e.sSerialNo" & _
            ", f.nReplAmtx" & _
            ", f.nAmtPaidX" & _
            ", g.sTransNox `sFinTrans`" & _
            ", h.sCompnyNm" & _
            ", g.nFinAmtxx" & _
            ", a.nDiscAmtx"
            
    lsSQL = lsSQL & _
         " FROM CP_SO_Detail a" & _
               " LEFT JOIN CP_Inventory_Serial e" & _
                  " ON a.sSerialID = e.sSerialID" & _
            ", CP_Inventory b" & _
               " LEFT JOIN Color d" & _
                  " ON b.sColorIDx = d.sColorIDx" & _
            ", CP_Model c" & _
            ", CP_SO_Master f" & _
               " LEFT JOIN CP_SO_Finance g" & _
                  " ON f.sTransNox = g.sTransNox" & _
               " LEFT JOIN Client_Master h" & _
                  " ON g.sClientID = h.sClientID" & _
         " WHERE a.sTransNox = " & strParm(oTrans.Master("sTransNox")) & _
            " AND a.sTransNox = f.sTransNox" & _
            " AND a.sStockIDx = b.sStockIDx" & _
            " AND b.sModelIDx = c.sModelIDx" & _
         " ORDER BY a.nEntryNox"
   
   Set loModel = New Recordset
   loModel.Open lsSQL, oApp.Connection, , , adCmdText
   
   lsIMEI = "Unit IMEI: "
   lReplAmt = "PR Amt: "
   lnFinAmt = 0#
   lsFinTrans = ""
   lnTotalDisc = 0#
'   If IFNull(loModel("sFinTrans"), "") <> "" Then
'      lnFinAmt = loModel("nFinAmtxx")
'   End If
   
   With loModel
      .MoveFirst
      
      Do Until .EOF
         lrs.AddNew
         lrs("nField01").Value = loModel("nQuantity")
         
         'she 2015 - 4 - 10
         'print barrcode if <> serialize
         If loModel("cHsSerial") = xeYes Then
            lrs("sField01").Value = IFNull(loModel("sModelCde"), loModel("sModelNme"))
            lrs("sField02").Value = loModel("sBarrCode") & " " & IFNull(loModel("sColorNme")) & ";"
         Else
            lrs("sField01").Value = loModel("sBarrCode")
            lrs("sField02").Value = loModel("sDescript")
         End If
         
         If IFNull(loModel("nReplAmtx"), 0#) <> 0# Then
            lrs("sField03").Value = lReplAmt
            lrs("lField03").Value = loModel("nReplAmtx")
         Else
            lrs("lField03").Value = 0#
         End If
         
         If IFNull(loModel("sFinTrans"), "") <> "" Then
            If lsFinTrans = loModel("sFinTrans") Then
               lnFinAmt = lnFinAmt - loModel("nFinAmtxx")
            Else
               lnFinAmt = loModel("nFinAmtxx")
            End If
         End If
         
         'mac 2022-03-24
         lnSelPrice = loModel("nUnitPrce")
         If CDbl(loModel("nDiscRate")) > 0# Then
            lnSelPrice = loModel("nQuantity") * CDbl(loModel("nUnitPrce")) + CDbl(loModel("nDiscAmtx"))
            lnSelPrice = (lnSelPrice * 100) / (100 - CDbl(loModel("nDiscRate")))
         End If
         'mac 2022-03-24

         If IFNull(loModel("sFinTrans"), "") <> "" Then
            lrs("lField01").Value = CDbl(lnSelPrice) - (loModel("nQuantity") * CDbl(lnSelPrice * CDbl(loModel("nDiscRate") / 100))) - CDbl(loModel("nDiscAmtx"))
'            lrs("lField02").Value = CDbl(lrs("lField01").Value) * CDbl(loModel("nQuantity"))
            lrs("lField02").Value = CDbl(lrs("lField01").Value) * CDbl(loModel("nQuantity"))
            
            'she 2022-03-23 to get the total discount amount
            lnTotalDisc = lnTotalDisc + (loModel("nQuantity") * CDbl(lnSelPrice * CDbl(loModel("nDiscRate") / 100))) + CDbl(loModel("nDiscAmtx"))
'            lrs("lField05").Value = (loModel("nQuantity") * CDbl(loModel("nUnitPrce") * CDbl(loModel("nDiscRate") / 100))) + CDbl(loModel("nDiscAmtx"))
'            lrs("lField06").Value = CDbl(lrs("lField02").Value - lrs("lField05").Value)
            
            lnFinAmt = lnFinAmt - lrs("lField05")
            'lnFinAmt = lnFinAmt - (CDbl(loModel("nUnitPrce")) * (100 - (CDbl(loModel("nDiscRate")))) / 100)
         Else
            lrs("lField01").Value = CDbl(lnSelPrice) - (loModel("nQuantity") * CDbl(lnSelPrice * CDbl(loModel("nDiscRate") / 100))) - CDbl(loModel("nDiscAmtx"))
'            lrs("lField02").Value = CDbl(lrs("lField01").Value) * loModel("nQuantity")
            lrs("lField02").Value = CDbl(lrs("lField01").Value) * CDbl(loModel("nQuantity"))
            
            'she 2022-03-23 to get the total discount amount
            lnTotalDisc = lnTotalDisc + (loModel("nQuantity") * CDbl(lnSelPrice * CDbl(loModel("nDiscRate") / 100))) + CDbl(loModel("nDiscAmtx"))

'            lrs("lField05").Value = (loModel("nQuantity") * CDbl(loModel("nUnitPrce") * CDbl(loModel("nDiscRate") / 100))) + CDbl(loModel("nDiscAmtx"))
'            lrs("lField06").Value = CDbl(lrs("lField02").Value - lrs("lField05").Value)
         End If
         
         If loModel("cHsSerial") = xeYes Then
            lsIMEI = lsIMEI & ";" & IFNull(loModel("sSerialNo"))
         End If
         
         lsFinTrans = IFNull(loModel("sFinTrans"))

         .MoveNext
      Loop
   End With
   
   pnTtlAdj = 0#
   
   lsSQL = "SELECT" & _
            " b.sTransNox" & _
            " FROM CP_CO_Master a" & _
            ", AR_Payment_Detail b" & _
            " WHERE a.sTransNox = b.sReferNox" & _
            " AND a.sReferNox = " & strParm(oTrans.Master("sTransNox")) & _
            " AND a.cTranStat = '4' "
          
   Debug.Print lsSQL
   Set lrsCOInv = New Recordset
   lrsCOInv.Open lsSQL, oApp.Connection, , , adCmdText
   
   If Not lrsCOInv.EOF Then
      Call ComputeAdjustment(lrsCOInv("sTransNox"))
      lrs("lField04").Value = Format(pnTtlAdj, "#,##0.00")
   Else
      lrs("lField04").Value = 0#
   End If

   ' assign important info to the report
   Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\CP_SI.rpt")
   oReport.DiscardSavedData
   oReport.FieldMappingType = crAutoFieldMapping
   oReport.Database.SetDataSource lrs

   loModel.Requery
   With oReceipt
      oReport.Sections("PH").ReportObjects("txtCustomer").SetText oTrans.Master("xFullName")
      oReport.Sections("PH").ReportObjects("txtDate").SetText Format(oTrans.Master("dTransact"), "MMM-DD-YYYY")
      oReport.Sections("PH").ReportObjects("txtAddress").SetText oTrans.Master("xAddressx")
      oReport.Sections("PH").ReportObjects("txtTIN").SetText ""
      oReport.Sections("PH").ReportObjects("txtBusiness").SetText ""
      oReport.Sections("PH").ReportObjects("txtTerm").SetText IFNull(oTrans.Master("sTermName"), "")
      oReport.Sections("PH").ReportObjects("txtPrepared").SetText Trim(txtField(17))
      oReport.Sections("RF").ReportObjects("txtIMEI").SetText lsIMEI
'      oReport.Sections("RF").ReportObjects("txtDiscount").SetText "Total Discount:" & Format(lnTotalDisc, "#,##0.00")
      oReport.Sections("RF").ReportObjects("txtAccessories").SetText getAccesories
'      oReport.Sections("RF").ReportObjects("txtGiveaways").SetText getGiveAways
      If IFNull(loModel("sFinTrans"), "") <> "" Then
         oReport.Sections("RF").ReportObjects("txtRemarks").SetText loModel("sCompnyNm") & " " & Format(loModel("nFinAmtxx"), "#,##0.00")
      Else
         oReport.Sections("RF").ReportObjects("txtRemarks").SetText oTrans.Master("sRemarksx")
      End If
   End With
   
   oReport.PrintOutEx False, 1
   lrs.Close
   PrintTrans = True

endProc:
   oTrans.CloseTransaction oTrans.Master(0)
   Set oReport = Nothing
   Set lrs = Nothing
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"
   
   
   
'   Dim lrs As New ADODB.Recordset
'   Dim lnCtr As Integer
'   Dim lsIMEI As String
'   Dim lReplAmt As String
'   Dim lsOldProc As String
'
'   Dim lrsCOInv As Recordset
'   Dim lsSQL As String
'   Dim loModel As Recordset
'   Dim lnFinAmt As Currency
'   Dim lsFinTrans As String
'   Dim lnTotalDisc As Currency 'she temporary for mobile fest
'
'   lsOldProc = "printTrans"
'   On Error GoTo errProc
'
'   PrintTrans = False
'
'   Set lrs = New ADODB.Recordset
'
'   lrs.Fields.Append "nField01", adInteger, 3
'   lrs.Fields.Append "sField01", adVarChar, 50
'   lrs.Fields.Append "sField02", adVarChar, 60
'   lrs.Fields.Append "sField03", adVarChar, 30
'   lrs.Fields.Append "sField04", adVarChar, 30
'   lrs.Fields.Append "lField01", adCurrency
'   lrs.Fields.Append "lField02", adCurrency
'   lrs.Fields.Append "lField03", adCurrency
'   lrs.Fields.Append "lField04", adCurrency
'   lrs.Fields.Append "lField05", adCurrency
'   lrs.Fields.Append "lField06", adCurrency
'   lrs.Open
'
'   lsSQL = "SELECT a.sTransNox" & _
'            ", b.sBarrCode" & _
'            ", c.sModelNme" & _
'            ", c.sModelCde" & _
'            ", d.sColorNme" & _
'            ", a.nUnitPrce" & _
'            ", a.nDiscRate" & _
'            ", a.nQuantity" & _
'            ", b.sDescript" & _
'            ", b.cHsSerial" & _
'            ", a.sSerialID" & _
'            ", e.sSerialNo" & _
'            ", f.nReplAmtx" & _
'            ", f.nAmtPaidX" & _
'            ", g.sTransNox `sFinTrans`" & _
'            ", h.sCompnyNm" & _
'            ", g.nFinAmtxx" & _
'            ", a.nDiscAmtx"
'
'    lsSQL = lsSQL & _
'         " FROM CP_SO_Detail a" & _
'               " LEFT JOIN CP_Inventory_Serial e" & _
'                  " ON a.sSerialID = e.sSerialID" & _
'            ", CP_Inventory b" & _
'               " LEFT JOIN Color d" & _
'                  " ON b.sColorIDx = d.sColorIDx" & _
'            ", CP_Model c" & _
'            ", CP_SO_Master f" & _
'               " LEFT JOIN CP_SO_Finance g" & _
'                  " ON f.sTransNox = g.sTransNox" & _
'               " LEFT JOIN Client_Master h" & _
'                  " ON g.sClientID = h.sClientID" & _
'         " WHERE a.sTransNox = " & strParm(oTrans.Master("sTransNox")) & _
'            " AND a.sTransNox = f.sTransNox" & _
'            " AND a.sStockIDx = b.sStockIDx" & _
'            " AND b.sModelIDx = c.sModelIDx" & _
'            " AND f.cTranStat <> 3" & _
'         " ORDER BY a.nEntryNox"
'
'      Set loModel = New Recordset
'      loModel.Open lsSQL, oApp.Connection, , adCmdText
'   Debug.Print lsSQL
'
'   With GridEditor1
'      lsIMEI = "Unit IMEI: "
'      lReplAmt = "PR Amt: "
'      lnFinAmt = 0#
'      lnTotalDisc = 0#
'      lsFinTrans = ""
'
'      For lnCtr = 1 To .Rows - 1
'         lrs.AddNew
'         lrs("nField01").Value = IFNull(.TextMatrix(lnCtr, 3), 1)
'         'she 2015-04-10
'         'print barrcode if <> serialize
'         If oTrans.Detail(lnCtr - 1, "cHsSerial") = xeYes Then
'            lrs("sField01").Value = oTrans.Detail(lnCtr - 1, "sModelNme")
'            lrs("sField02").Value = oTrans.Detail(lnCtr - 1, "sBarrCode") & " " & IFNull(oTrans.Detail(lnCtr - 1, "sColorNme"))
'         Else
'            lrs("sField01").Value = .TextMatrix(lnCtr, 1)
'            lrs("sField02").Value = Left(.TextMatrix(lnCtr, 2), 30)
'         End If
'
'         If oTrans.Master("nReplAmtx") <> 0# Then
'            lrs("sField03").Value = lReplAmt
'            lrs("lField03").Value = oTrans.Master("nReplAmtx")
'         Else
'            lrs("lField03").Value = 0#
'         End If
'
'         If IFNull(loModel("sFinTrans"), "") <> "" Then
'            If lsFinTrans = loModel("sFinTrans") Then
'               lnFinAmt = lnFinAmt - loModel("nFinAmtxx")
'            Else
'               lnFinAmt = loModel("nFinAmtxx")
'            End If
'         End If
'
'         If IFNull(loModel("sFinTrans"), "") <> "" Then
'            lrs("lField01").Value = CDbl(.TextMatrix(lnCtr, 4)) - .TextMatrix(lnCtr, 3) * (CDbl(.TextMatrix(lnCtr, 4) * Replace(.TextMatrix(lnCtr, 5), "%", "") / 100) + .TextMatrix(lnCtr, 6))
''            lrs("lField02").Value = CDbl(lrs("lField01").Value) * .TextMatrix(lnCtr, 3)
'            lrs("lField02").Value = lrs("lField01").Value * .TextMatrix(lnCtr, 3)
'
'            'she 2022-03-23 to get the total discount amount
'            lnTotalDisc = lnTotalDisc + .TextMatrix(lnCtr, 3) * (CDbl(.TextMatrix(lnCtr, 4) * Replace(.TextMatrix(lnCtr, 5), "%", "") / 100) + .TextMatrix(lnCtr, 6))
'
''            lrs("lField05").Value = .TextMatrix(lnCtr, 3) * (CDbl(.TextMatrix(lnCtr, 4) * Replace(.TextMatrix(lnCtr, 5), "%", "") / 100) + .TextMatrix(lnCtr, 6))
''            lrs("lField06").Value = CDbl(lrs("lField01").Value - lrs("lField05").Value)
'
'            lnFinAmt = lnFinAmt - lrs("lField05").Value
'            'lnFinAmt = lnFinAmt - CDbl(loModel("nUnitPrce")) * ((100 - (CDbl(loModel("nDiscRate")))) / 100)
'         Else
'            Debug.Print CDbl(.TextMatrix(lnCtr, 4)) - ((CDbl(.TextMatrix(lnCtr, 4) * Replace(.TextMatrix(lnCtr, 5), "%", "") / 100)) + .TextMatrix(lnCtr, 6))
'            lrs("lField01").Value = CDbl(.TextMatrix(lnCtr, 4)) - ((CDbl(.TextMatrix(lnCtr, 4) * Replace(.TextMatrix(lnCtr, 5), "%", "") / 100)) + .TextMatrix(lnCtr, 6))
''            lrs("lField02").Value = CDbl(.TextMatrix(lnCtr, 4)) * .TextMatrix(lnCtr, 3)
'            lrs("lField02").Value = lrs("lField01").Value * .TextMatrix(lnCtr, 3)
'
'            'she 2022-03-23 to get the total discount amount
'            lnTotalDisc = lnTotalDisc + (.TextMatrix(lnCtr, 3) * (CDbl(.TextMatrix(lnCtr, 4) * Replace(.TextMatrix(lnCtr, 5), "%", "") / 100))) + .TextMatrix(lnCtr, 6)
'
''            lrs("lField05").Value = (.TextMatrix(lnCtr, 3) * (CDbl(.TextMatrix(lnCtr, 4) * Replace(.TextMatrix(lnCtr, 5), "%", "") / 100))) + .TextMatrix(lnCtr, 6)
''            lrs("lField06").Value = CDbl(lrs("lField02").Value - lrs("lField05").Value)
'         End If
'
'         If oTrans.Detail(lnCtr - 1, "cHsSerial") = xeYes Then
'            lsIMEI = lsIMEI & " " & .TextMatrix(lnCtr, 1)
'         End If
'
'         lsFinTrans = IFNull(loModel("sFinTrans"))
'
'      Next
'   End With
'
'   'she 2015-12-01 4:29 pm
'   'show adjustment from charge invoice discount
'   pnTtlAdj = 0#
'
'   lsSQL = "SELECT" & _
'            " b.sTransNox" & _
'            " FROM CP_CO_Master a" & _
'            ", AR_Payment_Detail b" & _
'            " WHERE a.sTransNox = b.sReferNox" & _
'            " AND a.sReferNox = " & strParm(oTrans.Master("sTransNox")) & _
'            " AND a.cTranStat = '4' "
'
'   Debug.Print lsSQL
'   Set lrsCOInv = New Recordset
'   lrsCOInv.Open lsSQL, oApp.Connection, , , adCmdText
'
'   If Not lrsCOInv.EOF Then
'      Call ComputeAdjustment(lrsCOInv("sTransNox"))
'      lrs("lField04").Value = Format(pnTtlAdj, "#,##0.00")
'   Else
'      lrs("lField04").Value = 0#
'   End If
'
'   ' assign important info to the report
'   Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\CP_SI.rpt")
'   oReport.DiscardSavedData
'   oReport.FieldMappingType = crAutoFieldMapping
'   oReport.Database.SetDataSource lrs
'
'   With oReceipt
'      oReport.Sections("PH").ReportObjects("txtCustomer").SetText txtField(3)
'      oReport.Sections("PH").ReportObjects("txtDate").SetText Format(txtField(1), "MMM-DD-YYYY")
'      oReport.Sections("PH").ReportObjects("txtAddress").SetText Trim(txtField(4))
'      oReport.Sections("PH").ReportObjects("txtTIN").SetText ""
'      oReport.Sections("PH").ReportObjects("txtBusiness").SetText ""
'      oReport.Sections("PH").ReportObjects("txtTerm").SetText txtField(8)
'      oReport.Sections("PH").ReportObjects("txtPrepared").SetText txtField(17)
'      oReport.Sections("RF").ReportObjects("txtIMEI").SetText lsIMEI
''       oReport.Sections("RF").ReportObjects("txtDiscount").SetText "Total Discount:" & Format(lnTotalDisc, "#,##0.00")
'      oReport.Sections("RF").ReportObjects("txtAccessories").SetText getAccesories
''      oReport.Sections("RF").ReportObjects("txtGiveaways").SetText getGiveAways
'      If pnTtlAdj <> 0# Then oReport.Sections("RF").ReportObjects("txtDiscount").SetText "Discount"
'      If IFNull(loModel("sFinTrans"), "") <> "" Then
'         oReport.Sections("RF").ReportObjects("txtRemarks").SetText loModel("sCompnyNm") & " " & Format(loModel("nFinAmtxx"), "#,###0.00")
'      Else
'         oReport.Sections("RF").ReportObjects("txtRemarks").SetText oTrans.Master("sRemarksx")
'      End If
'   End With
'
'   oReport.PrintOutEx False, 1
'   lrs.Close
'   PrintTrans = True
'
'endProc:
'   If oTrans.Master("cTranStat") = xeStateOpen Then
'      oTrans.CloseTransaction oTrans.Master(0)
'   End If
'   Set oReport = Nothing
'   Set lrs = Nothing
'   Exit Function
'errProc:
'   ShowError lsOldProc & "( " & " )"
End Function

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   Dim lsOldProc As String

   lsOldProc = "txtField_Validate"
   On Error GoTo errProc

   With txtField(Index)
      .Text = TitleCase(.Text)

      Select Case Index
      Case 1
         If Not IsDate(.Text) Then .Text = oApp.ServerDate
         .Text = Format(.Text, "MMMM DD, YYYY")

         If DateDiff("d", CDate(.Text), oApp.ServerDate) < 0 Then
            .Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
         End If
         oTrans.Master(Index) = .Text
      Case 2, 18
         .Text = Format(.Text, ">")
         oTrans.Master(Index) = .Text
      Case 3
         oTrans.Master(Index) = .Text
         GridEditor1.Refresh

         lblAdvPayment.Caption = "0.00"
      Case 5
         If Not IsNumeric(.Text) Then .Text = 0#
         If .Text > 99999999.99 Then .Text = 0#
         .Text = Format(.Text, "#,##0.00")
         oTrans.Master(Index) = CDbl(.Text)
      Case 10
         If Not (cmbField(1).ListIndex = 2 Or cmbField(1).ListIndex = 1) Then
            oTrans.Master("sApplicNo") = ""
            .Text = ""
         End If
      Case 19, 20
         If Trim(.Text) = "" And Trim(txtField(0).Text) = "" Then
            ClearFields
            Exit Sub
         End If

         If Trim(.Text) <> Trim(.Tag) Then
            If oTrans.SearchTransaction(.Text, IIf(Index = 19, True, False)) Then
               LoadMaster
               LoadDetail
            Else
               If Index = 19 Then
                  ClearFields
                  Exit Sub
               Else
               End If
            End If
         End If
      Case Else
         If Index < 19 Then oTrans.Master(Index) = .Text
      End Select
   End With

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " _
                       & "  " & Index _
                       & ", " & Cancel _
                       & " )", True
End Sub

Private Sub LoadMaster()
   For pnCtr = 0 To 19
      Select Case pnCtr
      Case 0
         txtField(pnCtr).Text = Format(oTrans.Master(pnCtr), "@@@@-@@@@@@")
      Case 2, 19
         txtField(pnCtr).Text = oTrans.Master("sSalesInv")
      Case 3
         txtField(pnCtr).Text = oTrans.Master("xFullName")
         txtField(20).Text = txtField(pnCtr).Text
         txtField(20).Tag = txtField(20).Text
      Case 1
         txtField(pnCtr).Text = Format(oTrans.Master(pnCtr), "MMMM DD, YYYY")
      Case 5
         txtField(pnCtr).Text = Format(oTrans.Master("nTranTotl"), "#,##0.00")
      Case 9
         txtField(pnCtr).Text = Format(oTrans.Master(pnCtr), "#,##0.00")
      Case 10 To 16
      Case Else
         txtField(pnCtr).Text = IIf(IsNull(oTrans.Master(pnCtr)), "", oTrans.Master(pnCtr))
      End Select
   Next

   cmbField(1).ListIndex = oTrans.Master(13)
   pbLoad = True

   If CStr(oTrans.Master("nReplAmtx")) > 0 Then
      lblAdvPayment.Caption = Format(oTrans.Master("nReplAmtx"), "#,##0.00")
   End If

   Label2.Caption = Format(TransStat(oTrans.Master("cTranStat")), ">")
End Sub

Private Sub LoadDetail()
   Dim lnRow As Integer
   Dim lnCol As Integer
   Dim lnSubTotal As Currency

   With GridEditor1
      .Rows = oTrans.ItemCount + 1

      For lnRow = 0 To oTrans.ItemCount - 1
         For lnCol = 1 To .Cols - 1
            If lnCol = 5 Then
               .TextMatrix(lnRow + 1, lnCol) = Format(oTrans.Detail(lnRow, lnCol), "#0.00") & "%"
            Else
               .TextMatrix(lnRow + 1, lnCol) = oTrans.Detail(lnRow, lnCol)
            End If
         Next
'         lnSubTotal = CDbl(oTrans.Detail(lnRow, 3)) * CDbl(oTrans.Detail(lnRow, 4))
'         .TextMatrix(lnRow + 1, 7) = lnSubTotal - (lnSubTotal * _
'                                   (oTrans.Detail(lnRow, 5) / 100) - oTrans.Detail(lnRow, 6))
'         lnSubTotal = 0#
'         If oTrans.Detail(lnRow, 5) <> 0# Then
'            lnSubTotal = oTrans.Detail(lnRow, 3) * CDbl(oTrans.Detail(lnRow, 4) * oTrans.Detail(lnRow, 5) / 100)
'         End If
'
'         lnSubTotal = oTrans.Detail(lnRow, 3) * CDbl(oTrans.Detail(lnRow, 4)) - lnSubTotal - oTrans.Detail(lnRow, 6)

         'mac 2022-03-24
         lnSubTotal = oTrans.Detail(lnRow, 3) * CDbl(oTrans.Detail(lnRow, 4)) + oTrans.Detail(lnRow, 6)
         If oTrans.Detail(lnRow, 5) <> 0# Then
            lnSubTotal = (100 * lnSubTotal) / (100 - CDbl(oTrans.Detail(lnRow, 5)))
         End If
         
         .TextMatrix(lnRow + 1, 4) = lnSubTotal
         
         lnSubTotal = 0#
         If oTrans.Detail(lnRow, 5) <> 0# Then
            lnSubTotal = .TextMatrix(lnRow + 1, 3) * CDbl(.TextMatrix(lnRow + 1, 4)) * oTrans.Detail(lnRow, 5) / 100
         End If

         lnSubTotal = .TextMatrix(lnRow + 1, 3) * CDbl(.TextMatrix(lnRow + 1, 4)) - lnSubTotal - oTrans.Detail(lnRow, 6)
          
         .TextMatrix(lnRow + 1, 7) = lnSubTotal
         'mac 2022-03-24
      Next
      Call ComputeTotal
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

Private Sub LoadJO()
   Dim lorJO As Recordset
   Dim lsSQL As String
   Dim loFormPayment As frmJOPayment
   
   lsSQL = "SELECT sSalesTrn" & _
               ", sTransNox" & _
               ", nLaborAmt" & _
               ", nPartsAmt" & _
               ", nMiscChrg" & _
               ", nTranTotl" & _
               ", sSalesInv" & _
               ", dTransact" & _
            " FROM CP_JobOrder_Master " & _
            " WHERE sSalesTrn = " & strParm(oTrans.Master("sTransNox"))
            
   Set lorJO = New Recordset
   lorJO.Open lsSQL, oApp.Connection, , , adCmdText
   
   Set loFormPayment = New frmJOPayment
   Set loFormPayment.AppDriver = oApp
   
   If Not lorJO.EOF Then
   
      With loFormPayment
         .txtField(3).Enabled = False
         .txtField(4).Enabled = False
         .txtField(0).Text = Format(lorJO("nLaborAmt"), "#,##0.00")
         .txtField(1).Text = Format(lorJO("nPartsAmt"), "#,##0.00")
         .txtField(2).Text = Format(lorJO("nMiscChrg"), "#,##0.00")
         .txtField(3).Text = lorJO("sSalesInv")
         .txtField(4).Text = Format(lorJO("dTransact"), "MMM DD, YYYY")
         .lblTotal.Caption = Format(lorJO("nTranTotl"), "#,##0.00")
         .Show 1
      End With
   End If
End Sub

Function PrintJOSI() As Boolean
   Dim lrs As New ADODB.Recordset
   Dim loRecd As Recordset
   Dim lnCtr As Integer
   Dim lsOldProc As String
   Dim lsSQL As String
   
   lsOldProc = "PrintJOSI"
   On Error GoTo errProc
   
   PrintJOSI = False
   
   Set lrs = New ADODB.Recordset
   
   lrs.Fields.Append "nField01", adInteger, 3
   lrs.Fields.Append "sField01", adVarChar, 50
   lrs.Fields.Append "sField02", adVarChar, 60
   lrs.Fields.Append "sField03", adVarChar, 20
   lrs.Fields.Append "lField01", adCurrency
   lrs.Fields.Append "lField02", adCurrency
   lrs.Fields.Append "lField03", adCurrency
   lrs.Open
   
   lsSQL = "SELECT sSalesTrn" & _
               ", sTransNox" & _
               ", nLaborAmt" & _
               ", nPartsAmt" & _
               ", nMiscChrg" & _
               ", nTranTotl" & _
               ", sSalesInv" & _
               ", dTransact" & _
            " FROM CP_JobOrder_Master " & _
            " WHERE sSalesTrn = " & strParm(oTrans.Master("sTransNox"))
         
   Set loRecd = New Recordset
   loRecd.Open lsSQL, oApp.Connection, , , adCmdText
   

   
      lrs.AddNew
         lrs("lField01").Value = loRecd("nLaborAmt")
         lrs("lField02").Value = loRecd("nPartsAmt")
         lrs("lField03").Value = loRecd("nMiscChrg")

   ' assign important info to the report
   Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\JO_SI.rpt")
   oReport.DiscardSavedData
   oReport.FieldMappingType = crAutoFieldMapping
   oReport.Database.SetDataSource lrs
   
   With oTrans
      oReport.Sections("PH").ReportObjects("txtCustomer").SetText oTrans.Master("xFullName")
      oReport.Sections("PH").ReportObjects("txtDate").SetText Format(loRecd("dTransact"), "MMM-DD-YYYY")
      oReport.Sections("PH").ReportObjects("txtAddress").SetText oTrans.Master("xAddressx")
      oReport.Sections("PH").ReportObjects("txtTIN").SetText ""
      oReport.Sections("PH").ReportObjects("txtBusiness").SetText ""
      oReport.Sections("PH").ReportObjects("txtTerm").SetText ""
      oReport.Sections("PH").ReportObjects("txtPrepared").SetText oApp.UserName
      oReport.Sections("RF").ReportObjects("Labor").SetText "Labor"
      oReport.Sections("RF").ReportObjects("Parts").SetText "Parts"
      oReport.Sections("RF").ReportObjects("misc").SetText "Other Charges"
   End With
   
   oReport.PrintOutEx False, 1
   lrs.Close
   PrintJOSI = True

endProc:
   Set oReport = Nothing
   Set lrs = Nothing
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"
End Function

Private Sub ComputeAdjustment(lsTransNox As String)
   Dim lsSQL As String
   Dim lorec As Recordset
   
   lsSQL = "SELECT" & _
            " SUM(nCredtAmt) `nCredtAmt`" & _
            " FROM AR_Payment_Detail" & _
            " WHERE sTransNox = " & strParm(lsTransNox) & _
            " AND sSourceCD = 'CPCm' " & _
            " GROUP BY sTransNox"
   Set lorec = New Recordset
   lorec.Open lsSQL, oApp.Connection, , , adCmdText
   
   pnTtlAdj = lorec("nCredtAmt")
End Sub
