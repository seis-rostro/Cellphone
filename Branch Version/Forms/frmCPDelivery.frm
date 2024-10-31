VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmCPDelivery 
   BorderStyle     =   0  'None
   Caption         =   "Delivery"
   ClientHeight    =   9315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15255
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9315
   ScaleWidth      =   15255
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   90
      TabIndex        =   20
      TabStop         =   0   'False
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
      Height          =   990
      Left            =   1590
      Tag             =   "wt0;fb0"
      Top             =   3405
      Width           =   5220
      _ExtentX        =   9208
      _ExtentY        =   1746
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
         Left            =   1380
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   "00000000000000000000"
         Top             =   90
         Width           =   3720
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
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Text            =   "Mmm dd, yyyy"
         Top             =   480
         Width           =   3720
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IMEI"
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
         Left            =   945
         TabIndex        =   12
         Top             =   150
         Width           =   345
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
         Left            =   810
         TabIndex        =   14
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
         Enabled         =   0   'False
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
         Index           =   5
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
         Caption         =   "Source No."
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
         Left            =   345
         TabIndex        =   2
         Top             =   555
         Width           =   930
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
      Height          =   3855
      Left            =   6825
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   6800
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   3765
         Left            =   30
         TabIndex        =   22
         Top             =   30
         Width           =   8205
         _ExtentX        =   14473
         _ExtentY        =   6641
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
      Height          =   1890
      Left            =   1590
      Tag             =   "wt0;fb0"
      Top             =   1500
      Width           =   5220
      _ExtentX        =   9208
      _ExtentY        =   3334
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
         Height          =   825
         Index           =   4
         Left            =   1095
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   960
         Width           =   3990
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         TabIndex        =   7
         Text            =   "Mmm dd, yyyy"
         Top             =   90
         Width           =   1425
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
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
         Index           =   10
         Left            =   135
         TabIndex        =   10
         Top             =   990
         Width           =   765
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
      Height          =   4770
      Left            =   1605
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   4440
      Width           =   13515
      _ExtentX        =   23839
      _ExtentY        =   8414
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
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2430
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
      TabIndex        =   17
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
      TabIndex        =   18
      TabStop         =   0   'False
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
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   90
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1170
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
      Picture         =   "frmCPDelivery.frx":1DE8
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   90
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Order"
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
      Picture         =   "frmCPDelivery.frx":2562
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
Private WithEvents oTrans As clsCPStockIssue
Attribute oTrans.VB_VarHelpID = -1

Private pnPrintRow As Integer
Private poPrinter As clsPrintDirect
Private Const pxeMaxLine As Integer = 65

Private pnActiveRow As Integer
Private pbControl As Boolean
Private pnIndex As Integer
Private pbLoaded As Boolean
Private poRS As Recordset
Private pnCtr As Integer
Dim pbClosedTrans As Boolean

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String

   lsOldProc = "cmdButton_Click"
   ''On Error GoTo errProc

   Select Case Index
      Case 0 'OK/save
         If oTrans.Master("sDestinat") <> "" And txtField(2).Text <> "" Then
            With oTrans
               If .Detail(.ItemCount - 1, "sbarrcode") = "" Then .deleteDetail (.ItemCount - 1)
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
            End With
         Else
            MsgBox "Invalid Destination Branch!!!", vbCritical, "Warning"
            txtField(2).SetFocus
         End If
      Case 1 'del
         Call deleteDetail
         txtOthers(1).SetFocus
      Case 2 'cancel
         Unload Me
      Case 3 'new
         InitTransaction
      Case 5 'Print orders
         If MsgBox("Do you want to print transaction?", _
                  vbQuestion + vbYesNo, "Confirm") = vbYes Then
   
            Call PrintOrders
         Else
            MsgBox "No Orders to Print!!!"
         End If
   End Select

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Activate()
   Dim lsOldProc As String

   lsOldProc = "Form_Activate"
   ''On Error GoTo errProc

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
            If GetFocus = txtOthers(1).hwnd Then Exit Sub
            SetNextFocus
         Case vbKeyUp
            If GetFocus = txtField(2).hwnd Then Exit Sub
            SetPreviousFocus
         Case vbKeyF12
            'Call CPTransfer
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

   ''On Error GoTo errProc

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualLeft
   CenterChildForm mdiMain, Me

   Set oTrans = New clsCPStockIssue
   Set oTrans.AppDriver = oApp
   oTrans.Branch = oApp.BranchCode
   Call InitTransaction
   pnPrintRow = 0
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub InitGrid()
   Dim lnCtr As Integer

   With MSFlexGrid1
      .Cols = 7
      .Rows = 2
      .Clear

      pnActiveRow = 0
      .Row = 0
      .TextMatrix(0, 0) = "No."
      .TextMatrix(0, 1) = "IMEI"
      .TextMatrix(0, 2) = "Description"
      .TextMatrix(0, 3) = "Brand"
      .TextMatrix(0, 4) = "Model"
      .TextMatrix(0, 5) = "Color"
      .TextMatrix(0, 6) = "Pur Prc"

      .Row = 0
      'Column Width
      .ColWidth(0) = 650
      .ColWidth(1) = 2200
      .ColWidth(2) = 3200
      .ColWidth(3) = 2200
      .ColWidth(4) = 2010
      .ColWidth(5) = 1900
      .ColWidth(6) = 1200

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
      .Cols = 8
      .Rows = 2

      .Clear

      .TextMatrix(0, 0) = "No."
      .TextMatrix(0, 1) = "Brand"
      .TextMatrix(0, 2) = "Model"
      .TextMatrix(0, 3) = "Color"
      .TextMatrix(0, 4) = "QOH"
      .TextMatrix(0, 5) = "Rec"
      .TextMatrix(0, 6) = "Req"
      .TextMatrix(0, 7) = "Iss"
      .TextMatrix(1, 0) = "1"

      .Row = 0
      'Column Width
      .ColWidth(0) = 442
      .ColWidth(1) = 1750
      .ColWidth(2) = 1750
      .ColWidth(3) = 1400
      .ColWidth(4) = 700
      .ColWidth(5) = 700
      .ColWidth(6) = 700
      .ColWidth(7) = 700
      
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
   pnActiveRow = 0
   pbLoaded = False
   Set oTrans = Nothing
End Sub

Private Sub clearFields()
   Dim loTxt As TextBox

   For Each loTxt In txtField
      loTxt = ""
   Next

   For Each loTxt In txtOthers
      loTxt = ""
   Next
   
   pbClosedTrans = False
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
   If pbLoaded Then
      txtOthers(1).SetFocus
   End If
End Sub

Private Sub MSFlexGrid2_Click()
   With MSFlexGrid2
      .Col = 1
      .ColSel = .Cols - 1
   End With
End Sub

Private Sub MSFlexGrid2_GotFocus()
   txtOthers(1).SetFocus
End Sub

Private Sub oTrans_DetailRetrieved(ByVal Index As Integer)
   With MSFlexGrid1
      Select Case Index
         Case 1
            .TextMatrix(.Row, 1) = IFNull(oTrans.Detail(.Row - 1, "sSerialNo"))
            txtOthers(Index) = IFNull(oTrans.Detail(.Row - 1, Index))
         Case 2
            .TextMatrix(.Row, 2) = oTrans.Detail(.Row - 1, "sDescript")
            txtOthers(Index) = oTrans.Detail(.Row - 1, Index)
         Case 3
            .TextMatrix(.Row, 3) = oTrans.Detail(.Row - 1, "sBrandNme")
         Case 4
            .TextMatrix(.Row, 4) = IFNull(oTrans.Detail(.Row - 1, "sModelNme"), "") 'oTrans.Detail(.Row - 1, "sModelCde"))
         Case 5
            .TextMatrix(.Row, 5) = oTrans.Detail(.Row - 1, "sColorNme")
         Case 6
            .TextMatrix(.Row, 6) = IFNull(oTrans.Detail(.Row - 1, "nPurPrice"), 0#)
      End Select
   End With
End Sub

Private Sub oTrans_MasterRetrieved(ByVal Index As Integer)
   Select Case Index
   Case 1
      txtField(Index) = strLongDate(oTrans.Master("dTransact"))
   Case 2
      txtField(Index) = IFNull(oTrans.Master("sBranchNm"))
   Case 9
      txtField(Index) = IFNull(oTrans.Master("sSourceNo"))
      txtField(Index).Tag = txtField(Index)
   End Select
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   Select Case Index
      Case 1, 2, 4
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
            .SearchMaster 2, txtField(Index)
            If oTrans.Master("sSourceNo") <> "" Then
               InitGrid
               InitGrid2
               LoadStockRequest
            End If
         Case 5
            If txtField(Index) <> "" Then
               If txtField(Index).Tag <> txtField(Index) Then
                  If .SearchStockOrder(txtField(Index), True) Then
                     InitGrid
                     InitGrid2
                     LoadStockRequest
                  End If
               End If
            End If
            txtField(Index).Tag = txtField(Index)
         End Select
      Case vbKeyF3
         Select Case Index
         Case 2
            .SearchMaster 2, txtField(Index)
            If oTrans.Master("sSourceNo") <> "" Then
               InitGrid
               InitGrid2
               LoadStockRequest
            End If
         Case 5
            
            If .SearchStockOrder(txtField(Index), False) Then
               InitGrid
               InitGrid2
               LoadStockRequest
            End If
            
            txtField(Index) = txtField(Index).Tag
            If txtField(Index) = "" Then InitGrid2
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
   ''On Error GoTo errProc

   With MSFlexGrid2
      Set poRS = New Recordset
      With poRS.Fields
         .Append "sStockIDx", adVarChar, 12
         .Append "nQtyOnHnd", adInteger
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

      txtField(5) = Format(oTrans.StockReqSourceNo, "@@@@@@-@@@@@@")
      For lnCtr = 0 To lnRow - 1
         poRS.AddNew
         poRS("sStockIDx") = oTrans.StockDetail(lnCtr, "sStockIDx")
         poRS("nQtyOnHnd") = oTrans.StockDetail(lnCtr, "nQtyOnHnd")
         poRS("sBrandNme") = oTrans.StockDetail(lnCtr, "sBrandNme")
         poRS("sModelNme") = IFNull(oTrans.StockDetail(lnCtr, "sModelNme"), oTrans.StockDetail(lnCtr, "sModelCde"))
         poRS("sColorNme") = oTrans.StockDetail(lnCtr, "sColorNme")
         poRS("nQtyOnHnd") = oTrans.StockDetail(lnCtr, "nQtyOnHnd")
         poRS("nRecOrder") = oTrans.StockDetail(lnCtr, "nRecOrder")
         poRS("nQuantity") = oTrans.StockDetail(lnCtr, "nQuantity")
'         poRS("sBarrCode") = oTrans.StockDetail(lnCtr, "sBarrCode")
         poRS("nIssuedxx") = 0
         
         Debug.Print poRS("sStockIDx")
      Next

      .Rows = poRS.RecordCount + 1
      .Rows = IIf(.Rows = 1, 2, .Rows)

      If .Rows > 14 Then
         .ColWidth(2) = 1500
      Else
         .ColWidth(2) = 1750
      End If

      If Not poRS.EOF Then poRS.MoveFirst
      lnCtr = 1
      Do Until poRS.EOF
         .TextMatrix(lnCtr, 0) = lnCtr
         .TextMatrix(lnCtr, 1) = poRS("sBrandNme")
         .TextMatrix(lnCtr, 2) = poRS("sModelNme")
         .TextMatrix(lnCtr, 3) = poRS("sColorNme")
         .TextMatrix(lnCtr, 4) = poRS("nQtyOnHnd")
         .TextMatrix(lnCtr, 5) = poRS("nRecOrder")
         .TextMatrix(lnCtr, 6) = poRS("nQuantity")
         .TextMatrix(lnCtr, 7) = poRS("nIssuedxx")
         lnCtr = lnCtr + 1
         poRS.MoveNext
      Loop
   End With
   
   txtField(3) = oTrans.SourceNo
   txtField(5) = txtField(3)
   txtField(5).Tag = txtField(3)

'   If txtField(3) = "" Then
'      txtField(3).SetFocus
'   Else
      txtOthers(1).SetFocus
'   End If
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   Select Case Index
      Case 1, 2, 3, 4
         Call HighlightOff(Me.txtField(Index))
   End Select
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   With oTrans
      Select Case Index
      Case 1
         .Master("dTransact") = txtField(Index)
      Case 2
         .Master("sBranchNm") = txtField(Index)
      Case 5
         Call txtField_KeyDown(Index, vbKeyReturn, 0)
      Case Else
         .Master(Index) = txtField(Index)
      End Select
   End With
End Sub

Private Sub txtOthers_GotFocus(Index As Integer)
   Select Case Index
      Case 1
         Call HighlightOn(Me.txtOthers(Index))
   End Select
End Sub

Private Sub txtOthers_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   With MSFlexGrid1
      If oTrans.Master("sBranchNm") = "" Or IsNull(oTrans.Master("sBranchNm")) Then Exit Sub
      Select Case Index
         Case 1
            Select Case KeyCode
               Case vbKeyF3
                  If oTrans.searchDetail(.Row - 1, IIf(Index = 1, "xReferNox", "sDescript"), txtOthers(1)) Then
                     txtOthers(1).SetFocus
                  End If
               Case vbKeyReturn
                  If .Row = .Rows - 1 Then
                     If oTrans.searchDetail(.Row - 1, IIf(Index = 1, "xReferNox", "sDescript"), txtOthers(1)) Then
                        txtOthers(1).SetFocus
                     End If
                     Call addDetail
                  Else
                     .Row = .Rows - 1
                     .Col = 0
                     .ColSel = .Cols - 1
                     setDetailInfo
               End If
            End Select
      End Select
   End With
End Sub

Private Sub addDetail()
   Dim lsStockIDx As String
   
   With MSFlexGrid1
      If oTrans.Detail(.Row - 1, "sStockIDx") = "" Then Exit Sub
      lsStockIDx = oTrans.Detail(.Row - 1, "sStockIDx")
      
      'find matched reference # on cp order
      If Not TypeName(poRS) = "Nothing" Then
         If poRS.RecordCount > 0 Then
            poRS.MoveFirst
            poRS.Find "sStockIDx = " & strParm(lsStockIDx), 0, adSearchForward, adBookmarkFirst
            If Not poRS.EOF Then
               If .Row = .Rows - 1 Then
                  poRS("nIssuedxx") = poRS("nIssuedxx") + 1
                  
                  If poRS("nIssuedxx") > poRS("nQuantity") Then
                     MsgBox "Quantity to issue exceeds the request."
                  End If
                  
                  With MSFlexGrid2
                     .TextMatrix(poRS.AbsolutePosition, 7) = poRS("nIssuedxx")
                     .Row = poRS.AbsolutePosition
                     If .Row > 14 Then .TopRow = .Row - 13
                     .Col = 1
                     .ColSel = .Cols - 1
                  End With
               End If
            End If
            poRS.Cancel
         End If
      End If
      
      If oTrans.addDetail Then
         .Rows = .Rows + 1
         .Row = .Rows - 1
         .Col = 0
         .ColSel = .Cols - 1
         
         If .Rows > 16 Then
            .ColWidth(1) = 2595
            .ColWidth(2) = 2595
            .TopRow = .Rows - 14
         Else
'            .ColWidth(1) = 2500
'            .ColWidth(2) = 3200
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
      txtOthers(1) = oTrans.Detail(lnRow - 1, "xReferNox")
      txtOthers(2) = oTrans.Detail(lnRow - 1, "sDescript")
   End With
End Sub

Private Sub txtOthers_LostFocus(Index As Integer)
   Select Case Index
      Case 1
         Call HighlightOff(Me.txtOthers(Index))
   End Select
End Sub
Private Sub ClearOthers()
   Dim loTxt As TextBox

   For Each loTxt In txtOthers
      loTxt = ""
   Next
End Sub

Private Sub deleteDetail()
   Dim lnCtr As Integer
   Dim lnRow As Integer
   Dim lsStockIDx As String

   With MSFlexGrid1
      lsStockIDx = oTrans.Detail(.Row - 1, "sStockIDx")
      'find matched engine # on mc order
      If TypeName(poRS) = "Nothing" Then GoTo NopoRS
      
      poRS.Find "sStockIDx = " & strParm(lsStockIDx), 0, adSearchForward, adBookmarkFirst

      If oTrans.deleteDetail(.Row - 1) Then
         lnRow = oTrans.ItemCount

         If Not poRS.EOF Then
            poRS("nIssuedxx") = poRS("nIssuedxx") - 1
            MSFlexGrid2.TextMatrix(poRS.AbsolutePosition, 7) = poRS("nIssuedxx")
         End If
         poRS.Cancel

NopoRS:
         lnRow = oTrans.ItemCount
         If lnRow = 0 Then
            oTrans.addDetail
            lnRow = 1
         Else
            If oTrans.Detail(MSFlexGrid1.Row - 1, "sBarrCode") <> "" Then
               Call oTrans.deleteDetail(MSFlexGrid1.Row - 1)
               If oTrans.ItemCount = 0 Then Call oTrans.addDetail
            End If
            'If oTrans.Detail(oTrans.ItemCount - 1, "sReferNox") <> "" Then oTrans.AddDetail
         End If

         InitGrid

         lnRow = oTrans.ItemCount
         .Rows = lnRow + 1

         For lnCtr = 0 To lnRow - 1
            .TextMatrix(lnCtr + 1, 0) = lnCtr + 1
            .TextMatrix(lnCtr + 1, 1) = IFNull(oTrans.Detail(lnCtr, "sBarrcode"))
            .TextMatrix(lnCtr + 1, 2) = IFNull(oTrans.Detail(lnCtr, "sDescript"))
            .TextMatrix(lnCtr + 1, 3) = IFNull(oTrans.Detail(lnCtr, "sBrandNme"))
            .TextMatrix(lnCtr + 1, 4) = IFNull(oTrans.Detail(lnCtr, "sModelNme"), oTrans.Detail(lnCtr, "sModelCde"))
            .TextMatrix(lnCtr + 1, 5) = IFNull(oTrans.Detail(lnCtr, "sColorNme"))
         Next

         .Row = .Rows - 1
         .Col = 0
         .ColSel = .Cols - 1

         setDetailInfo
      End If
   End With
End Sub

Private Sub InitTransaction()
   Call clearFields
   Call InitGrid
   Call InitGrid2

   oTrans.IsCellphoneUnits = True
   oTrans.InitTransaction
   oTrans.NewTransaction
   Call LoadMaster
   cmdButton(3).Visible = False
End Sub

Private Function BranchAutomate(ByVal sBranchCd As String) As Boolean
   Dim lrs As Recordset
   
   Set lrs = New Recordset
   lrs.Open "SELECT * FROM Branch" & _
               " WHERE sBranchCd = " & strParm(sBranchCd) & _
                  " AND cAutomate = " & strParm(xeYes) _
   , oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   
   If Not lrs.EOF Then BranchAutomate = True
   Set lrs = Nothing
End Function

Private Function PrintTrans() As Boolean
   Dim loreport As frmRepViewer
   Dim lrs As Recordset
   Dim loRS As Recordset
   Dim lnCtr As Integer
   Dim lsOldProc As String
   Dim lnTotlWSerial As Double
   Dim lnTotlWOSerial As Double
   Dim lsSourceNo As String
   
   lsOldProc = "PrinTrans"
   ''On Error GoTo errProc

   PrintTrans = True
   
   Set lrs = New ADODB.Recordset

   lrs.Fields.Append "nField01", adInteger, 3
   lrs.Fields.Append "nField02", adChar, 1
   lrs.Fields.Append "sField01", adVarChar, 20
   lrs.Fields.Append "sField02", adVarChar, 128
   lrs.Fields.Append "sField03", adVarChar, 30
   lrs.Fields.Append "sField04", adVarChar, 12
   lrs.Fields.Append "sField05", adVarChar, 100
   lrs.Open

   lsSourceNo = IFNull(oTrans.StockReqSourceNo, "")
   
   With oTrans
      lnTotlWOSerial = 0
      lnTotlWSerial = 0
      
      For lnCtr = 0 To .ItemCount - 1
         lrs.AddNew
         lrs.Fields("nField01") = oTrans.Detail(lnCtr, "nQuantity")
         lrs.Fields("nField02") = oTrans.Detail(lnCtr, "cHsSerial")
         lrs.Fields("sField01") = oTrans.Detail(lnCtr, "sBarrCode")
         lrs.Fields("sField02") = oTrans.Detail(lnCtr, "sDescript")
         lrs.Fields("sField03") = oTrans.Detail(lnCtr, "sSerialNo")
         lrs.Fields("sField04") = oTrans.Detail(lnCtr, "sStockIDx")
         lrs.Fields("sField05") = oTrans.Detail(lnCtr, "sBrandNme")
         If oTrans.Detail(lnCtr, "cHsSerial") = xeYes Then
            lnTotlWSerial = lnTotlWSerial + 1
         Else
            lnTotlWOSerial = lnTotlWOSerial + CDbl(oTrans.Detail(lnCtr, "nQuantity"))
         End If
      Next
      lrs.Sort = "nField02 DESC,sField05,sField05,sField03"
   End With

   'assign important info to the report
   Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\CP_Transfer.rpt")
   oReport.DiscardSavedData
   oReport.FieldMappingType = crAutoFieldMapping
   oReport.Database.SetDataSource lrs

   Set loRS = New ADODB.Recordset
   If loRS.State = adStateOpen Then loRS.Close

   loRS.Open "SELECT" _
               & "  a.sAddressx" _
               & ", CONCAT(b.sTownName, ', ' , c.sProvName, ' ' , b.sZippCode) xTownName" _
               & ", a.sBranchNm" _
            & " FROM Branch a" _
               & " LEFT JOIN TownCity b" _
                  & " LEFT JOIN Province c" _
                     & " ON b.sProvIDxx = c.sProvIDxx" _
                  & " ON a.sTownIDxx = b.sTownIDxx" _
            & " WHERE a.sBranchCd = " & strParm(oTrans.Master("sDestinat")) _
            , oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText

   oReport.Sections("PHa").ReportObjects("txtRefNo").SetText "CP" & "-" & Right(oTrans.Master("sTransNox"), 10)
   oReport.Sections("PHa").ReportObjects("txtDate").SetText txtField(1).Text
   oReport.Sections("PHb").ReportObjects("txtTo").SetText loRS("sBranchNm")
   oReport.Sections("PHb").ReportObjects("txtToAddress").SetText loRS("sAddressx") & IFNull(loRS("xTownName"), "")
   oReport.Sections("PHb").ReportObjects("txtFrom").SetText oApp.ClientName
   oReport.Sections("PHb").ReportObjects("txtFromAddress").SetText oApp.Address & ", " & oApp.TownCity & ", " & oApp.Province & " " & oApp.ZippCode
   oReport.Sections("RFb").ReportObjects("txtRemarks").SetText lsSourceNo & " " & txtField(4).Text
   oReport.Sections("RFb").ReportObjects("txtWithSerial").SetText IIf(lnTotlWSerial = 0, "", Format(lnTotlWSerial, "#,##0"))
   oReport.Sections("RFb").ReportObjects("txtWOutSerial").SetText IIf(lnTotlWOSerial = 0, "", Format(lnTotlWOSerial, "#,##0"))
   oReport.Sections("PF").ReportObjects("txtRptUser").SetText oApp.UserName

   Set loreport = New frmRepViewer
   Set loreport.ReportSource = oReport
   loreport.Show
   
   PrintTrans = True

endPoc:
   If Not pbClosedTrans Then
      If BranchAutomate(oTrans.Master("sDestinat")) Then
         If oTrans.CloseTransaction(oTrans.Master(0)) Then pbClosedTrans = True
      End If
   End If
   Set loreport = Nothing
   Set oReport = Nothing
   Set lrs = Nothing
   Set loRS = Nothing
   Exit Function
errProc:
   PrintTrans = False
   ShowError lsOldProc & "( " & " )"
End Function

'she 2016-03-01 10:53 am
Private Function PrintOrders() As Boolean
   Dim loRS As Recordset
   Dim lsSQL As String
   Dim lsLineStr As String
   Dim lsTransNox As String
  
   lsSQL = "SELECT g.sBranchNm" & _
               ", a.sTransNox" & _
               ", a.dTransact" & _
               ", a.sRemarksX" & _
               ", e.sBrandNme" & _
               ", d.sModelNme" & _
               ", d.sModelCde" & _
               ", f.sColorNme" & _
               ", b.nQtyOnHnd" & _
               ", b.nQuantity" & _
            " FROM CP_Stock_Request_Master a" & _
            ", CP_Stock_Request_Detail b" & _
            ", CP_Inventory c" & _
                  " LEFT JOIN CP_Model d" & _
                     " ON c.sModelIDx = d.sModelIDx" & _
                  " LEFT JOIN CP_Brand e" & _
                     " ON c.sBrandIDx = e.sBrandIDx" & _
                  " LEFT JOIN Color f" & _
                     " ON c.sColorIDx = f.sColorIDx" & _
            ", Branch g" & _
            " WHERE a.sTransNox = b.sTransNox" & _
            " AND b.sStockIDx = c.sStockIDx" & _
            " AND LEFT(a.sTransNox,4) = g.sBranchCd" & _
            " AND a.sTransNox = " & strParm(oTrans.StockReqSourceNo) & _
            " ORDER BY e.sBrandNme, d.sModelCde,d.sModelNme,f.sColorNme"
   
   Set loRS = New Recordset
   loRS.Open lsSQL, oApp.Connection, , , adCmdText
   
   If loRS.EOF Then GoTo endProc
   
   
   Set poPrinter = New clsPrintDirect
   With poPrinter
      .FontName = "Draft 20cpi"
      .FontSize = 10
      
      If Not .BegPrint() Then GoTo endProc
      
      If pnPrintRow = 65 Or pnPrintRow = 0 Then
         If lsTransNox <> loRS("sTransNox") Then
            lsLineStr = padRight("Brand", 15) & " " & _
                                 padRight("Code", 25) & " " & _
                                 padRight("Model", 25) & " " & _
                                 padRight("Color", 15) & " " & _
                                 padRight("QOH", 7) & " " & _
                                 padRight("Order", 7)
   
               .PrintText pnPrintRow, 2, lsLineStr
               pnPrintRow = pnPrintRow + 1
   
               lsLineStr = padRight(loRS("sBranchNm"), 20) & " " & _
                              Format(loRS("sTransNox"), "@@@@-@@-@@@@@@") & "   " & _
                              Format(loRS("dTransact"), "MMM DD, YYYY")
               .PrintText pnPrintRow, 2, lsLineStr
               pnPrintRow = pnPrintRow + 2
         End If
      End If
      
      Do Until loRS.EOF
          lsLineStr = Left(padRight(Trim(loRS("sBrandNme") + "_______________"), 15), 15) & _
                     Left(padRight(Trim(loRS("sModelCde") + "____________________"), 25), 25) & _
                     Left(padRight(Trim(loRS("sModelNme") + "____________________"), 25), 25) & _
                     Left(padRight(Trim(loRS("sColorNme") + "_______________"), 15), 15) & _
                     Left(padRight(Trim(Format(loRS("nQtyOnHnd"), "#0") + "_______"), 7), 7) & _
                     padRight(Trim(Format(loRS("nQuantity"), "#0")), 7)
            
            .PrintText pnPrintRow, 2, lsLineStr
            pnPrintRow = pnPrintRow + 1
            
            lsTransNox = loRS("sTransNox")
      loRS.MoveNext
      Loop
      .EndPrint
   End With
   
   PrintOrders = True
   
'   If poPrinter Is Nothing Then
'      Set poPrinter = New clsPrintDirect
'      With poPrinter
'         .FontName = "Draft 15cpi"
'         .FontSize = 11
'
'         If Not .BegPrint() Then GoTo endProc
'      End With
'   End If
'
'
'   With poPrinter
'      Do Until loRS.EOF
'         If lsTransNox <> loRS("sTransNox") And pnPrintRow < pxeMaxLine Then
'            lsLineStr = padRight("Brand", 15) & " " & _
'                              padRight("Code", 25) & " " & _
'                              padRight("Model", 25) & " " & _
'                              padRight("Color", 15) & " " & _
'                              padRight("QOH", 7) & " " & _
'                              padRight("Order", 7)
'
'            .PrintText pnPrintRow, 2, lsLineStr
'            pnPrintRow = pnPrintRow + 1
'
'            lsLineStr = padRight(loRS("sBranchNm"), 20) & " " & _
'                           Format(loRS("sTransNox"), "@@@@-@@-@@@@@@") & "   " & _
'                           Format(loRS("dTransact"), "MMM DD, YYYY")
'            lsSQL = loRS("sTransNox")
'
'            .PrintText pnPrintRow, 2, lsLineStr
'            pnPrintRow = pnPrintRow + 2
'         End If
'
'         lsLineStr = Left(padRight(Trim(loRS("sBrandNme") + "_______________"), 15), 15) & _
'                     Left(padRight(Trim(loRS("sModelCde") + "____________________"), 25), 25) & _
'                     Left(padRight(Trim(loRS("sModelNme") + "____________________"), 25), 25) & _
'                     Left(padRight(Trim(loRS("sColorNme") + "_______________"), 15), 15) & _
'                     Left(padRight(Trim(Format(loRS("nQtyOnHnd"), "#0") + "_______"), 7), 7) & _
'                     padRight(Trim(Format(loRS("nQuantity"), "#0")), 7)
'
'         .PrintText pnPrintRow, 2, lsLineStr
'         pnPrintRow = pnPrintRow + 1
'
'         If pnPrintRow <= pxeMaxLine Then
'            pnPrintRow = 0
'            lsLineStr = padRight("Brand", 15) & " " & _
'                              padRight("Code", 25) & " " & _
'                              padRight("Model", 25) & " " & _
'                              padRight("Color", 15) & " " & _
'                              padRight("QOH", 7) & " " & _
'                              padRight("Order", 7)
'
'            .PrintText pnPrintRow, 2, lsLineStr
'            pnPrintRow = pnPrintRow + 1
'
'            lsLineStr = padRight(loRS("sBranchNm"), 20) & " " & _
'                           Format(loRS("sTransNox"), "@@@@-@@-@@@@@@") & "   " & _
'                           Format(loRS("dTransact"), "MMM DD, YYYY")
'            lsSQL = loRS("sTransNox")
'
'            .PrintText pnPrintRow, 2, lsLineStr
'            pnPrintRow = pnPrintRow + 2
'         End If
'
'         lsTransNox = loRS("sTransNox")
'
'         loRS.MoveNext
'      Loop
''         .EndPrint
'   End With
'
'   PrintOrders = True
   
endProc:
   Exit Function
End Function

Private Function CPTransfer() As Boolean
   Dim lsSQL As String
   Dim loRS As Recordset
   Dim loTrans As clsCPStockIssue
   
   lsSQL = "SELECT c.sSerialNo" & _
            " FROM CP_Transfer_Master a" & _
               " LEFT JOIN CP_Transfer_Detail b ON a.sTransNox = b.sTransNox" & _
               " LEFT JOIN CP_Inventory_Serial c ON b.sSerialID = c.sSerialID" & _
            " WHERE sDestinat LIKE 'C0M2%'" & _
               " AND cTranStat = '2'" & _
               " AND dReceived = '2023-02-07'"
      
   Set loRS = New Recordset
   loRS.Open lsSQL, oApp.Connection, , , adCmdText
   Set loRS.ActiveConnection = Nothing
   
   Set loTrans = New clsCPStockIssue
   Set loTrans.AppDriver = oApp
   
   With loTrans
      .Branch = oApp.BranchCode
      .IsCellphoneUnits = True
      .InitTransaction
      .NewTransaction
      
      .SearchMaster 2, "GCC - MP Warehouse"
      .Master("dTransact") = oApp.ServerDate
      .Master("sRemarksx") = "For Guanzon Festival"
      
      Do Until loRS.EOF
         '.Detail(.ItemCount - 1, "xrefernox") = loRS("sSerialNo")
         If .searchDetail(.ItemCount - 1, "xrefernox", loRS("sSerialNo")) Then
            If Not loRS.EOF Then .addDetail
         End If
         loRS.MoveNext
      Loop
      If .Detail(.ItemCount - 1, "sStockIDx") = "" Then
         .deleteDetail (.ItemCount - 1)
      End If
      MsgBox .Detail(.ItemCount - 1, "sStockIDx"), , .ItemCount
      
      If .SaveTransaction Then
         MsgBox "Tapos na po!"
      End If
   End With
End Function
