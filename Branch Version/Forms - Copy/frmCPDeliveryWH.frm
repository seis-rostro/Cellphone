VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCPDelivery 
   BorderStyle     =   0  'None
   Caption         =   "Motorcycle Delivery"
   ClientHeight    =   9630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15255
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9630
   ScaleWidth      =   15255
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   90
      TabIndex        =   25
      Top             =   540
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&New"
      AccessKey       =   "N"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCPDelivery.frx":0000
   End
   Begin xrControl.xrFrame xrFrame4 
      Height          =   2145
      Left            =   1590
      Tag             =   "wt0;fb0"
      Top             =   2565
      Width           =   5220
      _ExtentX        =   9208
      _ExtentY        =   3784
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   1095
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "9999999999999999999999999"
         Top             =   90
         Width           =   3990
      End
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   18
         Left            =   1095
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   "Mmm dd, yyyy"
         Top             =   1650
         Width           =   3990
      End
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   17
         Left            =   1095
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Text            =   "Mmm dd, yyyy"
         Top             =   1260
         Width           =   3990
      End
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   1095
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Text            =   "Mmm dd, yyyy"
         Top             =   870
         Width           =   3990
      End
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   1095
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   "Mmm dd, yyyy"
         Top             =   480
         Width           =   3990
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Bar Code"
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
         Index           =   9
         Left            =   150
         TabIndex        =   10
         Top             =   165
         Width           =   780
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Color"
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
         Index           =   6
         Left            =   480
         TabIndex        =   18
         Top             =   1725
         Width           =   450
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
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
         Index           =   5
         Left            =   435
         TabIndex        =   16
         Top             =   1335
         Width           =   495
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Brand"
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
         Index           =   4
         Left            =   435
         TabIndex        =   14
         Top             =   945
         Width           =   495
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Desc."
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
         Index           =   3
         Left            =   450
         TabIndex        =   12
         Top             =   555
         Width           =   480
      End
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   945
      Left            =   1590
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   5220
      _ExtentX        =   9208
      _ExtentY        =   1667
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   1380
         TabIndex        =   3
         Text            =   "M001-12-000001"
         Top             =   480
         Width           =   2055
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   1380
         TabIndex        =   1
         Text            =   "GMC Dagupan - Honda"
         Top             =   90
         Width           =   3705
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reference No."
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
         Index           =   8
         Left            =   105
         TabIndex        =   2
         Top             =   555
         Width           =   1200
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Destination"
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
         Index           =   7
         Left            =   345
         TabIndex        =   0
         Top             =   165
         Width           =   960
      End
   End
   Begin xrControl.xrFrame xrFrame3 
      Height          =   4170
      Left            =   6825
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   7355
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   4080
         Left            =   30
         TabIndex        =   20
         Top             =   30
         Width           =   8205
         _ExtentX        =   14473
         _ExtentY        =   7197
         _Version        =   393216
         FocusRect       =   0
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1050
      Left            =   1590
      Tag             =   "wt0;fb0"
      Top             =   1500
      Width           =   5220
      _ExtentX        =   9208
      _ExtentY        =   1852
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   1095
         TabIndex        =   9
         Text            =   "Mmm dd, yyyy"
         Top             =   570
         Width           =   1860
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   1095
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "M001-12-000001"
         Top             =   90
         Width           =   1860
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   3660
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "Mmm dd, yyyy"
         Top             =   90
         Width           =   1425
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Refer. No."
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
         Index           =   2
         Left            =   135
         TabIndex        =   8
         Top             =   645
         Width           =   825
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
         Left            =   3210
         TabIndex        =   6
         Top             =   165
         Width           =   390
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   360
         Left            =   1185
         Tag             =   "et0;ht2"
         Top             =   180
         Width           =   1860
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trans. No."
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
         Left            =   105
         TabIndex        =   4
         Top             =   165
         Width           =   855
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4725
      Left            =   1575
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   4740
      Width           =   13560
      _ExtentX        =   23918
      _ExtentY        =   8334
      _Version        =   393216
      FillStyle       =   1
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   90
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   1800
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
      Picture         =   "frmCPDelivery.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   90
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   540
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&OK"
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
      Picture         =   "frmCPDelivery.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   90
      TabIndex        =   24
      Top             =   1170
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Del. Row"
      AccessKey       =   "D"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCPDelivery.frx":166E
   End
End
Attribute VB_Name = "frmCPDelivery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pxeMODULENAME = "frmCPDelivery"
Private Const pxeVisibleRow = 20

Private oSkin As clsFormSkin
Private WithEvents oTrans As clsCPTransfer
Attribute oTrans.VB_VarHelpID = -1

Private pnActiveRow As Integer
Private pbControl As Boolean
Private pnIndex As Integer
Private pbLoaded As Boolean
Private poRS As Recordset
Private pnCtr As Integer
Private pbIsCPUnit As Boolean

Property Let IsCPUnit(ByVal Value As Boolean)
   pbIsCPUnit = Value
End Property

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   
   lsOldProc = "cmdButton_Click"
   On Error Goto errProc
   
   Select Case Index
      Case 0 'OK/save
         If oTrans.SaveTransaction Then
            MsgBox "Transaction Saved Successfully.", vbInformation, "Notice"
            If MsgBox("Do you want to print transaction?", _
               vbQuestion + vbYesNo, "Confirm") = vbYes Then
               
               Call PrintTrans
            End If
            InitTransaction
         Else
            MsgBox "Unable to Save Transaction.", vbInformation, "Notice"
         End If
      Case 1 'del
         Call DeleteDetail
         txtOthers(1).SetFocus
      Case 2 'cancel
         Unload Me
      Case 3 'new
         InitTransaction
   End Select

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Activate()
   Dim lsOldProc As String
   
   lsOldProc = "Form_Activate"
   On Error Goto errProc
   
   oApp.MenuName = Me.Tag
   Me.ZOrder 0

   If Not pbLoaded Then pbLoaded = True

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   With MSFlexGrid1
      Select Case KeyCode
         Case vbKeyReturn, vbKeyDown
            SetNextFocus
         Case vbKeyUp
            If GetFocus = txtField(2).hwnd Then Exit Sub
            SetPreviousFocus
      End Select
   End With
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   If pbControl Then
      If KeyCode = pbControl Then pbControl = False
   End If
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String
   lsOldProc = "Form_Load"
   
   On Error Goto errProc
   
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualLeft
   CenterChildForm mdiMain, Me
   
   Set oTrans = New clsCPTransfer
   Set oTrans.AppDriver = oApp
   oTrans.Branch = oApp.BranchCode

   Call ClearFields
   Call ClearOthers
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub InitGrid()
   Dim lnCtr As Integer

   With MSFlexGrid1
      .Cols = 6
      .Rows = 2
      .Clear
      
      pnActiveRow = 0
      .Row = 0
      .TextMatrix(0, 0) = "No."
      .TextMatrix(0, 1) = "Bar Code"
      .TextMatrix(0, 2) = "Description"
      .TextMatrix(0, 3) = "Brand"
      .TextMatrix(0, 4) = "Model"
      .TextMatrix(0, 5) = "Color"

      .Row = 0
      'Column Width
      .ColWidth(0) = 650
      .ColWidth(1) = 2500
      .ColWidth(2) = 3500
      .ColWidth(3) = 2500
      .ColWidth(4) = 2150
      .ColWidth(5) = 2150

      'Column Alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = 3
      Next
      .Row = 1
      'Column Alignment
      .TextMatrix(1, 0) = 1
      
      .ColAlignment(0) = flexAlignCenterCenter
      .ColAlignment(1) = flexAlignLeftCenter
      .ColAlignment(2) = flexAlignLeftCenter
      .ColAlignment(3) = flexAlignLeftCenter
      .ColAlignment(4) = flexAlignLeftCenter
      .ColAlignment(5) = flexAlignLeftCenter
      
      .Col = 0
      .ColSel = .Cols - 1
   End With
End Sub

Private Sub InitGrid2()
   Dim lnCtr As Integer

   With MSFlexGrid2
      .Cols = 7
      .Rows = 2
      
      .Clear
      
      .TextMatrix(0, 0) = "No."
      .TextMatrix(0, 1) = "Brand"
      .TextMatrix(0, 2) = "Model"
      .TextMatrix(0, 3) = "Color"
      .TextMatrix(0, 4) = "Rec"
      .TextMatrix(0, 5) = "Qty"
      .TextMatrix(0, 6) = "Iss"
      .TextMatrix(1, 0) = "1"
   
      .Row = 0
      'Column Width
      .ColWidth(0) = 442
      .ColWidth(1) = 2000
      .ColWidth(2) = 2000
      .ColWidth(3) = 1600
      .ColWidth(4) = 700
      .ColWidth(5) = 700
      .ColWidth(6) = 700

      'Column Alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = 3
      Next
      
      .Row = 1
      .ColAlignment(0) = flexAlignLeftCenter
      .ColAlignment(1) = flexAlignLeftCenter
      .ColAlignment(2) = flexAlignLeftCenter
      .ColAlignment(3) = flexAlignLeftCenter
      .ColAlignment(4) = flexAlignRightCenter
      .ColAlignment(5) = flexAlignRightCenter
      .ColAlignment(6) = flexAlignRightCenter

      .Col = 1
      .ColSel = .Cols - 1
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
   pnActiveRow = 0
   pbLoaded = False
End Sub


Private Sub ClearFields()
   Dim lotxt As TextBox
   
   For Each lotxt In txtField
      lotxt = ""
   Next
   
   For Each lotxt In txtOthers
      lotxt = ""
   Next
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

Private Sub MSFlexGrid1_Click()
   With MSFlexGrid1
      If pbLoaded Then setDetailInfo
      
      .Col = 0
      .ColSel = .Cols - 1
      
      txtOthers(1).SetFocus
   End With
End Sub

Private Sub MSFlexGrid1_RowColChange()
   If pbLoaded Then txtOthers(1).SetFocus
End Sub

Private Sub MSFlexGrid2_GotFocus()
   txtOthers(1).SetFocus
End Sub

Private Sub oTrans_DetailRetrieved(ByVal Index As Integer)
   With MSFlexGrid1
      Select Case Index
         Case 1
            .TextMatrix(.Row, 1) = oTrans.Detail(.Row - 1, Index)
            txtOthers(Index) = oTrans.Detail(.Row - 1, Index)
         Case 2
            .TextMatrix(.Row, 2) = oTrans.Detail(.Row - 1, Index)
            txtOthers(Index) = oTrans.Detail(.Row - 1, Index)
         Case 3
            .TextMatrix(.Row, 3) = oTrans.Detail(.Row - 1, Index)
            txtOthers(Index) = oTrans.Detail(.Row - 1, Index)
         Case 17
            .TextMatrix(.Row, 4) = oTrans.Detail(.Row - 1, Index)
            txtOthers(Index) = oTrans.Detail(.Row - 1, Index)
         Case 18
            .TextMatrix(.Row, 5) = oTrans.Detail(.Row - 1, Index)
            txtOthers(Index) = oTrans.Detail(.Row - 1, Index)
            Exit Sub
      End Select
   End With
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   Select Case Index
      Case 2, 4
         Call HighlightOn(Me.txtField(Index))
      Case 3
         If Len(txtField(Index)) <> 0 Then
            txtOthers(1).SetFocus
         Else
            Call HighlightOn(Me.txtField(Index))
         End If
   End Select
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   With oTrans
      Select Case KeyCode
         Case vbKeyReturn
            Select Case Index
               Case 2
               Case 4
                  If .SearchAcceptance(txtField(Index), True) Then
                     LoadMaster
                     LoadStockRequest
                  Else
                     txtField(Index) = ""
                  End If
            End Select
         Case vbKeyF3
            Select Case Index
               Case 2
                  If .SearchMaster(2, txtField(Index)) Then LoadMaster
               Case 4
                  If .SearchAcceptance(txtField(Index), False) Then
                     InitGrid
                     InitGrid2
                     LoadMaster
                     LoadStockRequest
                  End If
            End Select
      End Select
   End With
End Sub

Private Sub LoadMaster()
   With oTrans
      txtField(0) = Format(.Master("sTransNox"), "@@@@-@@-@@@@@@")
      txtField(1) = strLongDate(.Master("dTransact"))
      txtField(2) = IFNull(.Master(2))
      txtField(3) = .StockReqSourceNo
      txtField(4) = txtField(3)
      txtField(4).Tag = txtField(3)
      
      If Not pbLoaded Then Exit Sub
      If Len(txtField(3)) = 0 Then
         txtField(3).Locked = False
      Else
         txtField(3).Locked = True
         xrFrame2.Enabled = Len(txtField(2)) = 0
      End If
   End With
End Sub
Private Sub LoadStockRequest()
   Dim lsOldProc As String
   Dim lnCtr As Integer
   Dim lnRow As Integer
   
   lsOldProc = "LoadStockRequest"
   On Error Goto errProc
   
   With MSFlexGrid2
      Set poRS = New Recordset
      With poRS.Fields
         .Append "nQuantity", adInteger
         .Append "nRecOrder", adInteger
         .Append "nIssuedxx", adInteger
         .Append "sBarrCode", adVarChar, 9
         .Append "sBrandNme", adVarChar, 30
         .Append "sModelNme", adVarChar, 30
         .Append "sColorNme", adVarChar, 30
         poRS.Open
      End With
      
      lnRow = oTrans.StockOrderCount
      
      For lnCtr = 0 To lnRow - 1
         poRS.AddNew
         poRS("sBrandNme") = oTrans.StockDetail(lnCtr, "sBrandNme")
         poRS("sModelNme") = oTrans.StockDetail(lnCtr, "sModelNme")
         poRS("sColorNme") = oTrans.StockDetail(lnCtr, "sColorNme")
         poRS("nRecOrder") = oTrans.StockDetail(lnCtr, "nRecOrder")
         poRS("nQuantity") = oTrans.StockDetail(lnCtr, "nQuantity")
         poRS("sBarrCode") = oTrans.StockDetail(lnCtr, "sBarrCode")
         poRS("nIssuedxx") = 0
      Next
      
      .Rows = poRS.RecordCount + 1
      
      If .Rows > 15 Then
         .ColWidth(2) = 2150
      Else
         .ColWidth(2) = 2400
      End If
      
      poRS.MoveFirst
      lnCtr = 1
      Do Until poRS.EOF
         .TextMatrix(lnCtr, 0) = lnCtr
         .TextMatrix(lnCtr, 1) = poRS("sBrandNme")
         .TextMatrix(lnCtr, 2) = poRS("sModelNme")
         .TextMatrix(lnCtr, 3) = poRS("sColorNme")
         .TextMatrix(lnCtr, 4) = poRS("nRecOrder")
         .TextMatrix(lnCtr, 5) = poRS("nQuantity")
         .TextMatrix(lnCtr, 6) = poRS("nIssuedxx")
         lnCtr = lnCtr + 1
         poRS.MoveNext
      Loop
   End With
   
   If txtField(3) = "" Then
      txtField(3).SetFocus
   Else
      txtOthers(1).SetFocus
   End If
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   Select Case Index
      Case 2, 3, 4
         Call HighlightOff(Me.txtField(Index))
   End Select
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   With oTrans
      Select Case Index
         Case 2
            .Master("sDestinat") = .Master("sDestinat")
         Case 3
            .Master("sReferNox") = txtField(Index)
      End Select
   End With
End Sub

Private Sub txtOthers_GotFocus(Index As Integer)
   If Index = 1 Then Call HighlightOn(Me.txtOthers(1))
End Sub

Private Sub txtOthers_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If Index < 2 Then Exit Sub
   With MSFlexGrid1
      Select Case KeyCode
         Case vbKeyF3
            If oTrans.SearchDetail(.Row - 1, Index, txtOthers(1)) Then
               txtOthers(1).SetFocus
            End If
         Case vbKeyReturn
            If .Row = .Rows - 1 Then
               Call AddDetail
            Else
               .Row = .Rows - 1
               .Col = 0
               .ColSel = .Cols - 1
               setDetailInfo
            End If
      End Select
   End With
End Sub

Private Sub AddDetail()
   Dim lsBarrCode As String
   
   With MSFlexGrid1
      If oTrans.Detail(.Row - 1, "sBarrCode") = "" Then Exit Sub
      lsBarrCode = oTrans.Detail(.Row - 1, "sBarrCode")
      
      'find matched engine # on mc order
      poRS.MoveFirst
      poRS.Find "sMCInvIDx = " & strParm(lsMCInvIDx), 0, adSearchForward, adBookmarkFirst
      If Not poRS.EOF Then
         If .Row = .Rows - 1 Then
            poRS("nIssuedxx") = poRS("nIssuedxx") + 1
            With MSFlexGrid2
               .TextMatrix(poRS.AbsolutePosition, 6) = poRS("nIssuedxx")
               .Row = poRS.AbsolutePosition
               If .Row > 14 Then .TopRow = .Row - 13
               .Col = 1
               .ColSel = .Cols - 1
            End With
         End If
      End If
      poRS.Cancel
      
      If oTrans.AddDetail Then
      
         .Rows = .Rows + 1
         .Row = .Rows - 1
         .Col = 0
         .ColSel = .Cols - 1
         
         If .Rows > 16 Then
            .ColWidth(1) = 2750
            .TopRow = .Rows - 14
         Else
            .ColWidth(1) = 3000
         End If
         
         .TextMatrix(.Row, 0) = .Row
         ClearOthers
      End If
   End With
End Sub

Private Sub setDetailInfo()
   Dim lnRow As Integer
   
   lnRow = MSFlexGrid1.Row
   
   With oTrans
      txtOthers(1) = oTrans.Detail(lnRow - 1, "sEngineNo")
      txtOthers(2) = oTrans.Detail(lnRow - 1, "sFrameNox")
      txtOthers(3) = oTrans.Detail(lnRow - 1, "sModelNme")
      txtOthers(4) = oTrans.Detail(lnRow - 1, "sColorNme")
      txtOthers(9) = oTrans.Detail(lnRow - 1, "sSerialID")
   End With
End Sub

Private Sub txtOthers_LostFocus(Index As Integer)
   If Index = 1 Then Call HighlightOff(Me.txtOthers(1))
End Sub
Private Sub ClearOthers()
   Dim lotxt As TextBox
   
   For Each lotxt In txtOthers
      lotxt = ""
   Next
End Sub

Private Sub DeleteDetail()
   Dim lnCtr As Integer
   Dim lnRow As Integer
   Dim lsMCInvIDx As String
   
   With MSFlexGrid1
      lsMCInvIDx = oTrans.Detail(.Row - 1, "sMCInvIDx")
      'find matched engine # on mc order
      poRS.MoveFirst
      poRS.Find "sMCInvIDx = " & strParm(lsMCInvIDx), 0, adSearchForward, adBookmarkFirst
         
      If oTrans.DeleteDetail(.Row - 1) Then
         lnRow = oTrans.ItemCount
         
         If Not poRS.EOF Then
            poRS("nIssuedxx") = poRS("nIssuedxx") - 1
            MSFlexGrid2.TextMatrix(poRS.AbsolutePosition, 6) = poRS("nIssuedxx")
         End If
         poRS.Cancel
         
         If lnRow = 0 Then
            oTrans.AddDetail
            lnRow = 1
         Else
            If oTrans.Detail(oTrans.ItemCount - 1, "sEngineNo") <> "" Then oTrans.AddDetail
         End If
         
         InitGrid
         
         lnRow = oTrans.ItemCount
         .Rows = lnRow + 1
         
         For lnCtr = 0 To lnRow - 1
            .TextMatrix(lnCtr + 1, 0) = lnCtr + 1
            .TextMatrix(lnCtr + 1, 1) = IFNull(oTrans.Detail(lnCtr, "sBrandNme"))
            .TextMatrix(lnCtr + 1, 2) = IFNull(oTrans.Detail(lnCtr, "sModelNme"))
            .TextMatrix(lnCtr + 1, 3) = IFNull(oTrans.Detail(lnCtr, "sColorNme"))
            .TextMatrix(lnCtr + 1, 4) = IFNull(oTrans.Detail(lnCtr, "sEngineNo"))
            .TextMatrix(lnCtr + 1, 5) = IFNull(oTrans.Detail(lnCtr, "sFrameNox"))
         Next
         
         .Row = .Rows - 1
         .Col = 0
         .ColSel = .Cols - 1
         
         setDetailInfo
      End If
   End With
End Sub

Private Sub InitTransaction()
   With oTrans
      oTrans.IsCellphoneUnits = pbIsCPUnit
      oTrans.InitTransaction
      oTrans.NewTransaction
   End With
   cmdButton(3).Visible = False
   
   Call ClearFields
   Call InitGrid
   Call InitGrid2
   Call LoadMaster
   txtField(2).SetFocus
End Sub

Private Function PrintTrans() As Boolean
   Dim lrs As ADODB.Recordset
   Dim lors As ADODB.Recordset
   Dim lnCtr As Integer
   Dim lsMCInvID As String
   Dim lasMCInv() As String
   Dim lanMCInv() As Integer
   Dim lsIncluded As String
   Dim lsExcluded As String
   Dim lnQuantity As Integer
   Dim lnGivenxxx  As Integer
   Dim lsAcsModID As String
   Dim lasAcsMod() As String
   Dim lbFirst As Boolean
   Dim lsOldProc As String
   
   lsOldProc = "PrinTrans"
   On Error Goto errProc

   PrintTrans = True
   
   Set lrs = New ADODB.Recordset

   lrs.Fields.Append "nField01", adInteger, 5
   lrs.Fields.Append "sField01", adVarChar, 120
   lrs.Fields.Append "sField02", adVarChar, 50
   lrs.Fields.Append "sField03", adVarChar, 50
   lrs.Fields.Append "sField04", adVarChar, 30
   lrs.Open

   With oTrans
      lsMCInvID = ""
      For pnCtr = 0 To .ItemCount - 1
         If .Detail(pnCtr, "sEngineNo") = "" Then Exit For

         If InStr(1, lsMCInvID, .Detail(pnCtr, "sMCInvIDx"), vbTextCompare) = 0 Then
            lsMCInvID = lsMCInvID & "�" & .Detail(pnCtr, "sMCInvIDx")
         End If
      Next
      
      lsMCInvID = Mid(lsMCInvID, 2)
      lasMCInv = Split(lsMCInvID, "�")
      ReDim lanMCInv(UBound(lasMCInv)) As Integer
      
      For pnCtr = 0 To .ItemCount - 1
         If .Detail(pnCtr, "sEngineNo") = "" Then Exit For
         
         lnCtr = InStr(1, lsMCInvID, .Detail(pnCtr, "sMCInvIDx"))
         lnCtr = lnCtr \ 9
         If lnCtr > 8 Then lnCtr = lnCtr - 1
         lanMCInv(lnCtr) = lanMCInv(lnCtr) + 1
      Next

      For pnCtr = 0 To UBound(lanMCInv)
         lbFirst = True
         For lnCtr = 0 To .ItemCount - 1
            If .Detail(lnCtr, "sMCInvIDx") = lasMCInv(pnCtr) Then
               lrs.AddNew
               If lbFirst Then
                  lrs("nField01").Value = lanMCInv(pnCtr)
                  lrs("sField01").Value = .Detail(lnCtr, "sModelNme")
                  lbFirst = False
               End If
               lrs("sField02").Value = .Detail(lnCtr, "sEngineNo")
               lrs("sField03").Value = .Detail(lnCtr, "sFrameNox")
               lrs("sField04").Value = .Detail(lnCtr, "sCompnyCd")
            End If
         Next
      Next
      
      lsAcsModID = ""
      For pnCtr = 0 To .AccCount - 1
         If InStr(1, lsAcsModID, .Accessory(pnCtr, "sDescript"), vbTextCompare) = 0 Then
            lsAcsModID = lsAcsModID & "�" & .Accessory(pnCtr, "sDescript")
         End If
      Next
      
      lsAcsModID = Mid(lsAcsModID, 2)
      lasAcsMod = Split(lsAcsModID, "�")

      lsIncluded = ""
      lsExcluded = ""
      For pnCtr = 0 To UBound(lasAcsMod)
         lnGivenxxx = 0
         lnQuantity = 0
         For lnCtr = 0 To .AccCount - 1
            If .Accessory(lnCtr, "sDescript") = lasAcsMod(pnCtr) Then
               lnQuantity = lnQuantity + .Accessory(lnCtr, "nQuantity")
               lnGivenxxx = lnGivenxxx + .Accessory(lnCtr, "nGivenxxx")
            End If
         Next
         If lnGivenxxx > 0 Then lsIncluded = lsIncluded & ", " & Trim(Str(lnGivenxxx)) & " " & lasAcsMod(pnCtr)
         If lnQuantity > lnGivenxxx Then
            lsExcluded = lsExcluded & ", " & Trim(Str(lnQuantity - lnGivenxxx)) & " " & lasAcsMod(pnCtr)
         End If
      Next
      
      If lsIncluded <> "" Then lsIncluded = Mid(lsIncluded, 2)
      If lsExcluded <> "" Then lsExcluded = Mid(lsExcluded, 2)
   End With

   'assign important info to the report
   Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\Inter-BranchStockTransfer.rpt")
   oReport.DiscardSavedData
   oReport.FieldMappingType = crAutoFieldMapping
   oReport.Database.SetDataSource lrs

   Set lors = New ADODB.Recordset
   If lors.State = adStateOpen Then lors.Close
   
   lors.Open "SELECT" _
               & "  CONCAT(a.sAddressx, ', ', b.sTownName, ', ', c.sProvName, ' ', b.sZippCode) as Address" _
               & ", d.sCompnyNm" _
            & " From Branch a" _
               & ", TownCity b" _
               & ", Province c" _
               & ", Company d" _
            & " WHERE a.sBranchCd = " & strParm(oTrans.Master("sDestinat")) _
               & " AND a.sTownIDxx = b.sTownIDxx" _
               & " AND b.sProvIDxx = c.sProvIDxx" _
               & " AND a.sCompnyID = d.sCompnyID" _
            , oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   
   oReport.Sections("RH").ReportObjects("txtRefNo").SetText "MC" & "-" & Right(oTrans.Master("sTransNox"), 8)
   oReport.Sections("RH").ReportObjects("txtDate").SetText txtField(1).Text
   oReport.Sections("PH").ReportObjects("txtTo").SetText lors("sCompnyNm")
   oReport.Sections("PH").ReportObjects("txtToAddress").SetText lors("Address")
   oReport.Sections("PH").ReportObjects("txtFrom").SetText oApp.ClientName
   oReport.Sections("PH").ReportObjects("txtFromAddress").SetText oApp.Address & ", " & oApp.TownCity & ", " & oApp.Province & " " & oApp.ZippCode
   oReport.Sections("PF").ReportObjects("txtPrepared").SetText oApp.UserName
   
   If lsIncluded = "" Then
      lsIncluded = lsExcluded
      lsExcluded = ""
   End If
   
   If lsExcluded <> "" Then
      oReport.Sections("RF").ReportObjects("txtNote").SetText "Accessories" & "(" & lsIncluded & " )" _
                                                              & ", " & vbCrLf & "The Items" _
                                                              & "(" & lsExcluded & " ) " & "will follow..." & vbCrLf & txtField(4).Text
   Else
      oReport.Sections("RF").ReportObjects("txtNote").SetText "With Complete Accessories" & "(" & lsIncluded & " )" & vbCrLf & txtField(4).Text
   End If
   
   oReport.PrintOutEx False, 1
   lors.Close

endPoc:
   Call oTrans.CloseTransaction(oTrans.Master(0))
   Set oReport = Nothing
   Set lrs = Nothing
   Set lors = Nothing
   Exit Function
errProc:
   PrintTrans = False
   ShowError lsOldProc & "( " & " )"
End Function

