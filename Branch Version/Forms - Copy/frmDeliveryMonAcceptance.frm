VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmDeliveryMonAcceptance 
   BorderStyle     =   0  'None
   Caption         =   "Guanzon Delivery Posting"
   ClientHeight    =   8220
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8220
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   5340
      Left            =   120
      Tag             =   "wt0;fb0"
      Top             =   585
      Width           =   10290
      _ExtentX        =   18150
      _ExtentY        =   9419
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   11
         Left            =   1185
         TabIndex        =   25
         Text            =   "September 25, 2015"
         Top             =   3795
         Width           =   1650
      End
      Begin VB.TextBox txtArrTime 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   2880
         TabIndex        =   26
         Text            =   "04:30 AM"
         Top             =   3795
         Width           =   1050
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   10
         Left            =   1185
         TabIndex        =   22
         Text            =   "September 25, 2015"
         Top             =   3465
         Width           =   1650
      End
      Begin VB.TextBox txtBrdTime 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   2880
         TabIndex        =   23
         Text            =   "04:30 AM"
         Top             =   3465
         Width           =   1050
      End
      Begin VB.TextBox txtDepTIme 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   2880
         TabIndex        =   20
         Text            =   "04:30 AM"
         Top             =   3135
         Width           =   1050
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   9
         Left            =   1185
         TabIndex        =   19
         Text            =   "September 25, 2015"
         Top             =   3135
         Width           =   1650
      End
      Begin VB.TextBox txtEstTime 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   2880
         TabIndex        =   17
         Text            =   "04:30 AM"
         Top             =   2805
         Width           =   1050
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   7
         Left            =   1185
         TabIndex        =   16
         Text            =   "September 25, 2015"
         Top             =   2805
         Width           =   1650
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   6
         Left            =   1185
         TabIndex        =   13
         Text            =   "September 25, 2015"
         Top             =   2475
         Width           =   1650
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   1065
         Index           =   8
         Left            =   1170
         TabIndex        =   28
         Text            =   "This is a test. Don't take things as it is."
         Top             =   4125
         Width           =   3675
      End
      Begin VB.TextBox txtDate 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   2880
         TabIndex        =   14
         Text            =   "04:30 AM"
         Top             =   2475
         Width           =   1050
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   5
         Left            =   1185
         TabIndex        =   9
         Text            =   "Dacasin, Princess Joy"
         Top             =   1785
         Width           =   3660
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   4
         Left            =   1185
         TabIndex        =   7
         Text            =   "Cuison, Michael Torres"
         Top             =   1440
         Width           =   3660
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   1185
         TabIndex        =   5
         Text            =   "Adversalo, Rex Soriano"
         Top             =   1095
         Width           =   3660
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   1185
         TabIndex        =   11
         Text            =   "UEMI Tuguegarao - Multi"
         Top             =   2130
         Width           =   3660
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
         Height          =   315
         Index           =   0
         Left            =   1170
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   "M001-12-000001"
         Top             =   120
         Width           =   1860
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   1185
         TabIndex        =   3
         Text            =   "AVX 897609"
         Top             =   750
         Width           =   1395
      End
      Begin xrControl.xrFrame xrFrame2 
         Height          =   4410
         Left            =   4980
         Tag             =   "wt0;fb0"
         Top             =   780
         Width           =   5145
         _ExtentX        =   9075
         _ExtentY        =   7779
         BackColor       =   12632256
         ClipControls    =   0   'False
         BorderStyle     =   3
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
            Height          =   3960
            Left            =   105
            TabIndex        =   30
            Top             =   315
            Width           =   4920
            _ExtentX        =   8678
            _ExtentY        =   6985
            _Version        =   393216
            FocusRect       =   0
            AllowUserResizing=   2
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ROUTES(BRANCH) OF DELIVERY"
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
            TabIndex        =   29
            Top             =   75
            Width           =   2940
         End
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Arrival"
         Height          =   195
         Index           =   12
         Left            =   165
         TabIndex        =   24
         Top             =   3840
         Width           =   435
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BroadCast"
         Height          =   195
         Index           =   11
         Left            =   165
         TabIndex        =   21
         Top             =   3510
         Width           =   735
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Departure"
         Height          =   195
         Index           =   8
         Left            =   165
         TabIndex        =   18
         Top             =   3180
         Width           =   705
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Est. Arrival"
         Height          =   195
         Index           =   7
         Left            =   165
         TabIndex        =   15
         Top             =   2850
         Width           =   750
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
         Left            =   7200
         TabIndex        =   37
         Tag             =   "eb0;et0"
         Top             =   225
         Width           =   2865
      End
      Begin VB.Shape Shape3 
         Height          =   390
         Index           =   0
         Left            =   7140
         Top             =   165
         Width           =   2985
      End
      Begin VB.Shape Shape4 
         Height          =   330
         Index           =   0
         Left            =   7170
         Top             =   195
         Width           =   2925
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   195
         Index           =   10
         Left            =   165
         TabIndex        =   27
         Top             =   4275
         Width           =   630
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Porter #2"
         Height          =   195
         Index           =   6
         Left            =   165
         TabIndex        =   8
         Top             =   1860
         Width           =   660
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Porter #1"
         Height          =   195
         Index           =   5
         Left            =   165
         TabIndex        =   6
         Top             =   1515
         Width           =   660
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Driver"
         Height          =   195
         Index           =   4
         Left            =   165
         TabIndex        =   4
         Top             =   1170
         Width           =   420
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Destination"
         Height          =   195
         Index           =   3
         Left            =   165
         TabIndex        =   10
         Top             =   2205
         Width           =   795
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Trans,"
         Height          =   195
         Index           =   1
         Left            =   165
         TabIndex        =   12
         Top             =   2520
         Width           =   840
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trans. No."
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
         Left            =   165
         TabIndex        =   0
         Top             =   165
         Width           =   915
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   315
         Left            =   1275
         Tag             =   "et0;ht2"
         Top             =   225
         Width           =   1860
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vehicle No"
         Height          =   195
         Index           =   2
         Left            =   165
         TabIndex        =   2
         Top             =   825
         Width           =   780
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   315
         Left            =   7200
         Tag             =   "et0;ht2"
         Top             =   195
         Width           =   2880
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   10650
      TabIndex        =   32
      Top             =   585
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
      Picture         =   "frmDeliveryMonAcceptance.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   10650
      TabIndex        =   36
      Top             =   3105
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "Cl&ose"
      AccessKey       =   "o"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmDeliveryMonAcceptance.frx":077A
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2100
      Left            =   135
      TabIndex        =   31
      Top             =   5955
      Width           =   10260
      _ExtentX        =   18098
      _ExtentY        =   3704
      _Version        =   393216
      FocusRect       =   0
      SelectionMode   =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   10650
      TabIndex        =   33
      Top             =   1845
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "Arri&val"
      AccessKey       =   "v"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmDeliveryMonAcceptance.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   10650
      TabIndex        =   34
      Top             =   1215
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Inform"
      AccessKey       =   "I"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmDeliveryMonAcceptance.frx":17CE
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   10650
      TabIndex        =   35
      Top             =   2475
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
      Picture         =   "frmDeliveryMonAcceptance.frx":20A8
   End
End
Attribute VB_Name = "frmDeliveryMonAcceptance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pxeMODULENAME = "frmDeliveryMonitoring"
Private WithEvents oTrans As clsDeliveryMonitoring
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin
Private p_sDivision As String

Property Let Division(ByVal Value As String)
   p_sDivision = Value
End Property

Private Sub cmdButton_Click(Index As Integer)
   Dim lnRep As Integer

   Select Case Index
   Case 0   'Browse
      If oTrans.SearchTransaction() Then
         Call loadFields
      End If
   Case 1   'BroadCast
      If oTrans.BroadCast(oTrans.Master("sTransNox")) Then
         MsgBox "Transact BroadCast Successfully...", vbInformation, "INFORMATION"
         Call ClearFields
      Else
         MsgBox "Unable to update transaction!!!" & vbCrLf & _
                  "Please contact GGC SSG/SEG for assistance!!!", vbCritical, "WARNING"
      End If
   Case 2   'Closed
      Unload Me
   Case 3   'Arrival
      If oTrans.PostTransaction(oTrans.Master("sTransNox")) Then
         MsgBox "Transact Post Successfully...", vbInformation, "INFORMATION"
         txtField(11) = Format(oApp.ServerDate, "MMMM DD, YYYY")
         txtArrTime = Format(oApp.ServerDate, "HH:MM AM/PM")
      Else
         MsgBox "Unable to update transaction!!!" & vbCrLf & _
                  "Please contact GGC SSG/SEG for assistance!!!", vbCritical, "WARNING"
      End If
   Case 4   'Cancel
      If oTrans.CancelTransaction(oTrans.Master("sTransNox")) Then
         MsgBox "Transact Cancelled Successfully...", vbInformation, "INFORMATION"
      Else
         MsgBox "Unable to update transaction!!!" & vbCrLf & _
                  "Please contact GGC SSG/SEG for assistance!!!", vbCritical, "WARNING"
      End If
   End Select
End Sub

Private Sub Form_Activate()
   MSFlexGrid1.Refresh
   MSFlexGrid2.Refresh
   oApp.MenuName = Me.Tag
   Me.ZOrder 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyReturn, vbKeyDown
      SetNextFocus
   Case vbKeyUp
      SetPreviousFocus
   End Select
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String

   lsOldProc = "Form_Load"
   '''On Error GoTo errProc
   
   CenterChildForm mdiMain, Me

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualRight
   
   Set oTrans = New clsDeliveryMonitoring
   Set oTrans.AppDriver = oApp
   oTrans.Division = p_sDivision
   oTrans.TransStatus = 10
   oTrans.InitTransaction
   oTrans.NewTransaction
   
   InitGridDetail
   InitGridRoute
   ClearFields

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
End Sub

Private Sub InitGridRoute()
   Dim lnCtr As Integer
   
   With MSFlexGrid2
      .Rows = 2
      .Cols = 4
      .Font = "MS Sans Serif"

      'column title
      .TextMatrix(0, 0) = "No"
      .TextMatrix(0, 1) = "Branch/Route"
      .TextMatrix(0, 2) = "Arrival"
      .TextMatrix(0, 3) = "Departure"
      .Row = 0

      'column alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = 1
      Next

      'column width
      .ColWidth(0) = 450
      .ColWidth(1) = 2900
      .ColWidth(2) = 1500
      .ColWidth(3) = 1500
      
      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      
      .TextMatrix(1, 0) = "1"
      
      .Row = 1
      .Col = 1
      .ColSel = .Cols - 1
   End With
End Sub

Private Sub InitGridDetail()
   Dim lnCtr As Integer
   With MSFlexGrid1
      .Rows = 2
      .Cols = 5
      .Font = "MS Sans Serif"

      'column title
      .TextMatrix(0, 0) = "No"
      .TextMatrix(0, 1) = "Transfer"
      .TextMatrix(0, 2) = "Refer No"
      .TextMatrix(0, 3) = "Recepient"
      .TextMatrix(0, 4) = "Remarks"
      .Row = 0

      'column alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = 1
      Next

      'column width
      .ColWidth(0) = 450
      .ColWidth(1) = 2500
      .ColWidth(2) = 1500
      .ColWidth(3) = 3200
      .ColWidth(4) = 2500

      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      .ColAlignment(4) = 1
      
      .TextMatrix(1, 0) = "1"
      
      .Row = 1
      .Col = 1
      .ColSel = .Cols - 1
   End With
End Sub

Private Sub loadFields()
   Dim loTxt As TextBox
   Dim lnRow As Integer
   Dim lnCtr As Integer
   Dim lnCur As Integer
   
   For Each loTxt In txtField
      Select Case loTxt.Index
      Case 0
         loTxt.Text = Format(oTrans.Master(loTxt.Index), "@@@@@@-@@@@@@")
      Case 6, 7, 9, 10, 11
         loTxt.Text = Format(oTrans.Master(loTxt.Index), "MMMM DD, YYYY")
      Case Else
         loTxt.Text = oTrans.Master(loTxt.Index)
      End Select
   Next
   
   txtDate = Format(oTrans.Master("dTransact"), "HH:MM AM/PM")
   txtEstTime = Format(oTrans.Master("dEArrival"), "HH:MM AM/PM")
   txtDepTime = Format(oTrans.Master("dDepartre"), "HH:MM AM/PM")
   txtBrdTime = Format(oTrans.Master("dBroadcst"), "HH:MM AM/PM")
   txtArrTime = Format(oTrans.Master("dArrivalx"), "HH:MM AM/PM")
   
   With MSFlexGrid1
      If oTrans.ItemCount = 0 Then
         .Rows = 2
         
         .TextMatrix(1, 1) = ""
         .TextMatrix(1, 2) = ""
         .TextMatrix(1, 3) = ""
         .TextMatrix(1, 4) = ""
      Else
         .Rows = oTrans.ItemCount + 1
         lnCur = 0
         For lnCtr = 0 To oTrans.RouteCount - 1
            If Trim(oTrans.Route(lnCtr, "sBranchCd")) <> "" Then
               For lnRow = 0 To oTrans.ItemCount(oTrans.Route(lnCtr, "sBranchCd")) - 1
                  .TextMatrix(lnCur + 1, 0) = lnCur + 1
                  .TextMatrix(lnCur + 1, 1) = oTrans.Route(lnCtr, "sBranchNm")
                  .TextMatrix(lnCur + 1, 2) = oTrans.Detail(lnRow, "sDescript", oTrans.Route(lnCtr, "sBranchCd"))
                  .TextMatrix(lnCur + 1, 3) = oTrans.Detail(lnRow, "sReferNox", oTrans.Route(lnCtr, "sBranchCd"))
                  .TextMatrix(lnCur + 1, 4) = oTrans.Detail(lnRow, "sDestinat", oTrans.Route(lnCtr, "sBranchCd"))
                  lnCur = lnCur + 1
               Next
            End If
         Next
      End If
      
      .Row = 1
      .Col = 1
      .ColSel = .Cols - 1
   End With

   With MSFlexGrid2
      .Rows = oTrans.RouteCount + 1
      
      For lnRow = 0 To oTrans.RouteCount - 1
         .TextMatrix(lnRow + 1, 0) = lnRow + 1
         .TextMatrix(lnRow + 1, 1) = IFNull(oTrans.Route(lnRow, "sBranchNm"), "")
         .TextMatrix(lnRow + 1, 2) = IIf(IsNull(oTrans.Route(lnRow, "dArrivedx")), "", Format(oTrans.Route(lnRow, "dArrivedx"), "MM-DD-YY HH:MM AM/PM"))
         .TextMatrix(lnRow + 1, 3) = IIf(IsNull(oTrans.Route(lnRow, "dDepartre")), "", Format(oTrans.Route(lnRow, "dDepartre"), "MM-DD-YY HH:MM AM/PM"))
      Next
      
      .Row = 1
      .Col = 1
      .ColSel = .Cols - 1
   End With
   Label2.Caption = Format(TransStat(oTrans.Master("cTranStat")), ">")
End Sub

Private Sub ClearFields()
   Dim loTxt As TextBox
   
   For Each loTxt In txtField
      loTxt.Text = ""
   Next
   
   txtDate = ""
   txtEstTime = ""
   txtDepTime = ""
   txtBrdTime = ""
   txtArrTime = ""
   
   With MSFlexGrid1
      .Rows = 2
      
      .TextMatrix(1, 0) = 1
      .TextMatrix(1, 1) = ""
      .TextMatrix(1, 2) = ""
      .TextMatrix(1, 3) = ""
      .TextMatrix(1, 4) = ""
      
      .Row = 1
      .Col = 1
      .ColSel = .Cols - 1
   End With
   
   With MSFlexGrid2
      .Rows = 2
      
      .TextMatrix(1, 0) = 1
      .TextMatrix(1, 1) = ""
      
      .Row = 1
      .Col = 1
      .ColSel = .Cols - 1
   End With
   Label2 = "UNKNOWN"
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

