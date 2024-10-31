VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmCPInvUnitRequest 
   BorderStyle     =   0  'None
   Caption         =   "Unit Stock Request w/ ROQ Computation"
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
      TabIndex        =   0
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
      Picture         =   "frmCPInvUnitRequest.frx":0000
   End
   Begin xrControl.xrFrame xrFrame3 
      Height          =   3105
      Left            =   6825
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   5477
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   3015
         Left            =   30
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   30
         Width           =   8205
         _ExtentX        =   14473
         _ExtentY        =   5318
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
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5520
      Left            =   1575
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   3675
      Width           =   13560
      _ExtentX        =   23918
      _ExtentY        =   9737
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
      TabIndex        =   27
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
      Picture         =   "frmCPInvUnitRequest.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   90
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   1170
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&History"
      AccessKey       =   "H"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCPInvUnitRequest.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   90
      TabIndex        =   25
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
      Picture         =   "frmCPInvUnitRequest.frx":1F86
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   3105
      Left            =   1575
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   5220
      _ExtentX        =   9208
      _ExtentY        =   5477
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.Frame Frame1 
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   105
         TabIndex        =   28
         Top             =   2745
         Width           =   2805
         Begin VB.Label lblField 
            BackStyle       =   0  'Transparent
            Caption         =   "*Active Model"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Index           =   7
            Left            =   75
            TabIndex        =   29
            Top             =   0
            Width           =   2760
         End
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
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
         Index           =   5
         Left            =   4485
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Text            =   "A"
         Top             =   1905
         Width           =   600
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
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
         Index           =   9
         Left            =   2310
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Text            =   "00"
         Top             =   2295
         Width           =   600
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
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
         Index           =   8
         Left            =   1095
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Text            =   "00"
         Top             =   2295
         Width           =   600
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
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
         Index           =   7
         Left            =   2310
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Text            =   "00"
         Top             =   1905
         Width           =   600
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
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
         Index           =   6
         Left            =   1095
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "00"
         Top             =   1905
         Width           =   600
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
         Left            =   1095
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   "United Excelsior Marketing Inc"
         Top             =   750
         Width           =   3990
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
         Index           =   2
         Left            =   1095
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "United Excelsior Marketing Inc"
         Top             =   1125
         Width           =   3990
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
         Index           =   3
         Left            =   1095
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "United Excelsior Marketing Inc"
         Top             =   1500
         Width           =   3990
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   10
         Left            =   3945
         MaxLength       =   3
         TabIndex        =   24
         Text            =   "22"
         Top             =   2295
         Width           =   1140
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
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   "M001-12-000001"
         Top             =   180
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
         Index           =   11
         Left            =   3660
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   "Mmm dd, yyyy"
         Top             =   180
         Width           =   1425
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Classification"
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
         Left            =   3255
         TabIndex        =   15
         Top             =   1980
         Width           =   1125
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ROQ"
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
         Left            =   1860
         TabIndex        =   19
         Top             =   2370
         Width           =   405
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AMC"
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
         Left            =   585
         TabIndex        =   17
         Top             =   2370
         Width           =   375
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QOH"
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
         Index           =   3
         Left            =   1860
         TabIndex        =   13
         Top             =   1980
         Width           =   405
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "On Trnsit"
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
         Left            =   210
         TabIndex        =   11
         Top             =   1980
         Width           =   750
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
         Index           =   13
         Left            =   465
         TabIndex        =   5
         Top             =   825
         Width           =   495
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
         Index           =   14
         Left            =   465
         TabIndex        =   7
         Top             =   1200
         Width           =   495
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
         Index           =   15
         Left            =   510
         TabIndex        =   9
         Top             =   1575
         Width           =   450
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Order"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   16
         Left            =   3045
         TabIndex        =   21
         Top             =   2310
         Width           =   810
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
         TabIndex        =   3
         Top             =   255
         Width           =   390
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   360
         Left            =   1185
         Tag             =   "et0;ht2"
         Top             =   270
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
         TabIndex        =   1
         Top             =   255
         Width           =   855
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "PRESS F5 TO ADD ITEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   360
      Left            =   9645
      TabIndex        =   30
      Top             =   9225
      Width           =   5490
   End
End
Attribute VB_Name = "frmCPInvUnitRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pxeMODULENAME = "frmCPInvUnitRequest"
Private Const pxeVisibleRow = 20

Private oSkin As clsFormSkin
Private WithEvents oTrans As clsCPUnitOrder
Attribute oTrans.VB_VarHelpID = -1

Private poRSOrigROQ As Recordset    'serves as the original ROQ recorset acquired from class
Private poRSRecOrder As Recordset   'serves as the filtered ROQ result
Private poRS As Recordset           'serves as the Request Recordset

Private pnCtr As Integer
Private pnROQ As Integer
Private pnIndex As Integer
Private pnActiveRow As Integer

Private pbLoaded As Boolean
Private pbScroll As Boolean
Private pbClickd As Boolean
Private pbControl As Boolean
Private pbByModel As Boolean

Private pbInitialized As Boolean

Property Let AddModel(ByVal lors As Recordset)
   Call AddStock(lors)
End Property

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String

   lsOldProc = "cmdButton_Click"
   ''On Error GoTo errProc

   Select Case Index
      Case 0 'OK/save
         If SaveTransaction Then
            MsgBox "Transaction Saved Successfully.", vbInformation, "Notice"
            InitTransaction
         Else
            MsgBox "Unable to Save Transaction.", vbInformation, "Notice"
         End If
      Case 1 'history
         If pnROQ <= 0 Or pbInitialized = False Then Exit Sub
         
         With frmCPUnitHistory
            .Brand = poRSRecOrder.Fields("sBrandNme")
            .Model = poRSRecOrder.Fields("sModelNme")
            .IsCPUnit = True
            
            If pbByModel Then
               .History = oTrans.GetHistoryModel(poRSRecOrder.Fields("sModelIDx"))
            Else
               .History = oTrans.GetHistory(poRSRecOrder.Fields("sStockIDx"))
            End If
            
            .Show vbModal
            txtField(10).SetFocus
         End With
         
      Case 2 'close
         Unload Me
      Case 3 'new
         If Not InitTransaction Then Unload Me
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
   
   If pbLoaded Then Exit Sub

   oApp.MenuName = Me.Tag
   Me.ZOrder 0

   If Not pbLoaded Then pbLoaded = True
   If Not pbScroll Then pbScroll = True
   pbClickd = False

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String
   
   lsOldProc = "Form_KeyDown"
   ''On Error GoTo errProc

   With MSFlexGrid1
      Select Case KeyCode
         Case vbKeyReturn, vbKeyDown
            If GetFocus = MSFlexGrid2.hwnd Then
            Else
               If pnIndex = 10 And pnActiveRow < .Rows - 1 Then
                  ' this does not trigger lost focus or validate
                  If pnIndex = 10 Then
                     Call txtField_Validate(pnIndex, False)
                     Exit Sub
                  End If
                  .Row = pnActiveRow + 1

                  Call HiglightRow(Me.MSFlexGrid1, .Row, 1)
                  pnActiveRow = .Row
                  If .Row > 19 And Not .RowIsVisible(.Row) Then _
                     .TopRow = .Row - (pxeVisibleRow - 3)

                  Call setFieldInfo
                  With txtField(10)
                     .SelStart = 0
                     .SelLength = Len(.Text)
                     .SetFocus
                  End With
                  Exit Sub
               Else
                  SetNextFocus
               End If
            End If
         Case vbKeyUp
            If pbControl Then
               If .Row > 2 Then
                  ' this does not trigger lost focus or validate
                  If pnIndex = 10 Then
                     Call txtField_Validate(pnIndex, False)
                  End If

                  If .Row = .TopRow Then .TopRow = .TopRow - 1

                  .Row = pnActiveRow - 1
                  Call HiglightRow(Me.MSFlexGrid1, .Row, 1)
                  pnActiveRow = .Row

                  Call setFieldInfo
                  With txtField(10)
                     .SelStart = 0
                     .SelLength = Len(.Text)
                     .SetFocus
                  End With
               End If
            Else
               SetPreviousFocus
            End If
         Case vbKeyF5
            If TypeName(poRSOrigROQ) = "Nothing" Then Exit Sub
         
            With frmSearchUnit
               Set .ROQ = poRSOrigROQ
               .Show vbModal
               Call LoadRecOrder
               
               Call setFieldInfo
            End With
         Case vbKeyControl
            pbControl = True
            KeyCode = 0
         Case vbKeyEscape
            txtField(1) = ""
            txtField(2) = ""
            txtField(3) = ""
      End Select
   End With

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String
   
   lsOldProc = "Form_KeyUp"
   ''On Error GoTo errProc
   
   If pbControl Then
      If KeyCode = pbControl Then pbControl = False
   End If

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
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

   Set oTrans = New clsCPUnitOrder
   Set oTrans.AppDriver = oApp
   oTrans.DisplayMessage = True
   oTrans.Branch = oApp.BranchCode

   cmdButton(3).Visible = True
   xrFrame1.Enabled = False
   ClearFields
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub addDetail()
   Dim lsOldProc As String

   lsOldProc = "AddDetail"
   ''On Error GoTo errProc

   With poRSRecOrder
      poRS.AddNew
      poRS("sBrandIDx") = .Fields("sBrandIDx")
      poRS("sModelIDx") = .Fields("sModelIDx")
      poRS("sColorIDx") = .Fields("sColorIDx")
      poRS("nQuantity") = .Fields("nQuantity")
      poRS("sStockIDx") = .Fields("sStockIDx")
      poRS("sBrandNme") = .Fields("sBrandNme")
      poRS("sModelNme") = .Fields("sModelNme")
      poRS("sColorNme") = .Fields("sColorNme")
      poRS("cCategory") = IIf(pbByModel, "Mod", "Inv")

      If pbByModel = False Then
         poRS("cClassify") = .Fields("cClassify")
         poRS("nAveMonsl") = .Fields("nAveMonsl")
         poRS("nRecOrder") = .Fields("nRecOrder")
         poRS("nQtyOnHnd") = .Fields("nQtyOnHnd")
         poRS("nOnTranst") = .Fields("nOnTranst")
      Else
         poRS("cClassify") = .Fields("cClassMdl")
         poRS("nAveMonsl") = .Fields("nAveMonMd")
         poRS("nRecOrder") = .Fields("nRecOrdMd")
         poRS("nQtyOnHnd") = .Fields("nQtyOnHMd")
         poRS("nOnTranst") = .Fields("nOnTrnsMd")
      End If
   End With

   Call LoadDetail

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub AddStock(ByVal lors As Recordset)
   Dim lsOldProc As String

   lsOldProc = "AddStock"
   ''On Error GoTo errProc
   
   With lors
      If Not .EOF Then
         Call poRSRecOrder.Find("sStockIDx = " & strParm(.Fields("sStockIDx")), 0, adSearchForward, adBookmarkFirst)
         If poRSRecOrder.EOF Then
            poRSRecOrder.Cancel
            poRSRecOrder.AddNew
            poRSRecOrder.Fields("sStockIDx") = .Fields("sStockIDx")
            poRSRecOrder.Fields("sModelIDx") = .Fields("sModelIDx")
            poRSRecOrder.Fields("sBrandNme") = .Fields("sBrandNme")
            poRSRecOrder.Fields("sModelNme") = .Fields("sModelNme")
            poRSRecOrder.Fields("sColorNme") = .Fields("sColorNme")
            poRSRecOrder.Fields("nAveMonSl") = .Fields("nAveMonSl")
            poRSRecOrder.Fields("nAveMonMd") = .Fields("nAveMonMd")
            poRSRecOrder.Fields("nMinLevel") = .Fields("nMinLevel")
            poRSRecOrder.Fields("nMaxLevel") = .Fields("nMaxLevel")
            poRSRecOrder.Fields("cClassify") = .Fields("cClassify")
            poRSRecOrder.Fields("cClassMdl") = .Fields("cClassMdl")
            poRSRecOrder.Fields("nQuantity") = .Fields("nQuantity")
            poRSRecOrder.Fields("nOnTranst") = .Fields("nOnTranst")
            poRSRecOrder.Fields("nOnTrnsMd") = .Fields("nOnTrnsMd")
            poRSRecOrder.Fields("nRecOrder") = .Fields("nRecOrder")
            poRSRecOrder.Fields("nRecOrdMd") = .Fields("nRecOrdMd")
            poRSRecOrder.Fields("nQtyOnHnd") = .Fields("nQtyOnHnd")
            poRSRecOrder.Fields("nQtyOnHMd") = .Fields("nQtyOnHMd")
            poRSRecOrder.Fields("sBrandIDx") = .Fields("sBrandIDx")
            poRSRecOrder.Fields("sColorIDx") = .Fields("sColorIDx")
         Else
            MsgBox "Item was already on the list.", vbInformation, "Notice"
         End If
      End If
   End With

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub LoadDetail()
   Dim lsOldProc As String
   Dim lnCtr As Integer
   Dim lnRow As Integer

   lsOldProc = "LoadDetail"
   ''On Error GoTo errProc

   With MSFlexGrid2
      InitGrid2
      lnRow = poRS.RecordCount

      .Rows = lnRow + 1

      If .Rows > 11 Then
         .ColWidth(1) = 2490
      Else
         .ColWidth(1) = 2750
      End If

      lnCtr = 0
      poRS.Sort = "sModelNme ASC"
      If poRS.RecordCount <> 0 Then poRS.MoveFirst
      Do Until poRS.EOF
         DoEvents
         .TextMatrix(lnCtr + 1, 0) = lnCtr + 1
         .TextMatrix(lnCtr + 1, 1) = poRS("sModelNme")
         .TextMatrix(lnCtr + 1, 2) = poRS("sColorNme")
         .TextMatrix(lnCtr + 1, 3) = poRS("nQtyOnHnd")
         .TextMatrix(lnCtr + 1, 4) = poRS("nAveMonsl")
         .TextMatrix(lnCtr + 1, 5) = poRS("nRecOrder")
         .TextMatrix(lnCtr + 1, 6) = poRS("nQuantity")

         lnCtr = lnCtr + 1
         poRS.MoveNext
      Loop
      
      .Col = 1
      .ColSel = .Cols - 1
   End With

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub UpdateDetail()
   Dim lsOldProc As String

   lsOldProc = "UpdateDetail"
   ''On Error GoTo errProc

   With MSFlexGrid2
      poRS.Move .Row - 1, adBookmarkFirst
      poRS("nQuantity") = poRSRecOrder("nQuantity")
      .TextMatrix(.Row, 6) = poRS("nQuantity")
   End With

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub deleteDetail()
   Dim lsOldProc As String

   lsOldProc = "DeleteDetail"
   ''On Error GoTo errProc

   poRS.Move MSFlexGrid2.Row - 1, adBookmarkFirst
   poRS.Delete

   Call LoadDetail

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub InitGrid()
   Dim lsOldProc As String
   Dim lnCtr As Integer

   lsOldProc = "InitGrid"
   ''On Error GoTo errProc

   With MSFlexGrid1
      .Cols = 10
      .Rows = 3
      .FixedRows = 2
      .Clear

      pnActiveRow = 0
      .Row = 0
      .TextMatrix(0, 0) = " "
      .TextMatrix(0, 1) = " "
      .TextMatrix(0, 2) = " "
      .TextMatrix(0, 3) = " "
      .TextMatrix(0, 4) = "AMC"
      .TextMatrix(0, 5) = "AMC"
      .TextMatrix(0, 6) = "ROQ"
      .TextMatrix(0, 7) = "ROQ"
      .TextMatrix(0, 8) = "QOH"
      .TextMatrix(0, 9) = "QOH"
      'Row 1
      .TextMatrix(1, 0) = "No"
      .TextMatrix(1, 1) = "Brand"
      .TextMatrix(1, 2) = "Model"
      .TextMatrix(1, 3) = "Color"
      .TextMatrix(1, 4) = "Model"
      .TextMatrix(1, 5) = "Inv."
      .TextMatrix(1, 6) = "Model"
      .TextMatrix(1, 7) = "Inv."
      .TextMatrix(1, 8) = "Model"
      .TextMatrix(1, 9) = "Inv."

      .MergeCells = flexMergeFree 'disables colsel procedure
      .MergeRow(0) = True

      .Row = 0
      'Column Width
      .ColWidth(0) = 530
      .ColWidth(1) = 2975 '2725
      .ColWidth(2) = 3380
      .ColWidth(3) = 1800
      .ColWidth(4) = 800
      .ColWidth(5) = 800
      .ColWidth(6) = 800
      .ColWidth(7) = 800
      .ColWidth(8) = 800
      .ColWidth(9) = 800

      'Column Alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = 3
      Next
      .Row = 1
      'Column Alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = 3
      Next

      .Row = 2
      .ColAlignment(0) = flexAlignLeftCenter
      .ColAlignment(1) = flexAlignLeftCenter
      .ColAlignment(2) = flexAlignLeftCenter
      .ColAlignment(3) = flexAlignLeftCenter
      .ColAlignment(4) = flexAlignRightCenter
      .ColAlignment(5) = flexAlignRightCenter
      .ColAlignment(6) = flexAlignRightCenter
      .ColAlignment(7) = flexAlignRightCenter
      .ColAlignment(8) = flexAlignRightCenter
      .ColAlignment(9) = flexAlignRightCenter

      Call HiglightRow(Me.MSFlexGrid1, .Row, 1)
      pnActiveRow = .Row
   End With

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub InitGrid2()
   Dim lsOldProc As String
   Dim lnCtr As Integer

   lsOldProc = "InitGrid2"
   ''On Error GoTo errProc

   With MSFlexGrid2
      .Cols = 7
      .Rows = 2

      .Clear

      .TextMatrix(0, 0) = "No."
      .TextMatrix(0, 1) = "Model"
      .TextMatrix(0, 2) = "Color"
      .TextMatrix(0, 3) = "QOH"
      .TextMatrix(0, 4) = "AMC"
      .TextMatrix(0, 5) = "ROQ"
      .TextMatrix(0, 6) = "QTY"
      .TextMatrix(1, 0) = "1"

      .Row = 0
      'Column Width
      .ColWidth(0) = 442
      .ColWidth(1) = 2750 '2490
      .ColWidth(2) = 1800
      .ColWidth(3) = 800
      .ColWidth(4) = 800
      .ColWidth(5) = 800
      .ColWidth(6) = 800

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
      .ColAlignment(3) = flexAlignRightCenter
      .ColAlignment(4) = flexAlignRightCenter
      .ColAlignment(5) = flexAlignRightCenter
      .ColAlignment(6) = flexAlignRightCenter

      If Not pbLoaded Then
         .Col = 1
         .ColSel = .Cols - 1
      End If
   End With

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Dim lsOldProc As String

   lsOldProc = "Form_Unload"
   ''On Error GoTo errProc
   
   Set oSkin = Nothing
   pnActiveRow = 0

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub MSFlexGrid1_Click()
   Dim lsOldProc As String

   lsOldProc = "MSFlexGrid1_Click"
   ''On Error GoTo errProc
   
   pbClickd = True
   
   If pnROQ <= 0 Then GoTo endProc
   
   
   Call HiglightRow(Me.MSFlexGrid1, MSFlexGrid1.Row, 1)
   
   Call setFieldInfo(True)
endProc:
   With txtField(10)
      .SelStart = 0
      .SelLength = Len(.Text)
      If .Enabled Then .SetFocus
   End With
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub MSFlexGrid1_DblClick()
   Dim lsOldProc As String

   lsOldProc = "MSFlexGrid1_DblClick"
   ''On Error GoTo errProc

   With MSFlexGrid1
      If .MouseRow > 0 Then
         Call cmdButton_Click(1)
      End If
   End With

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub MSFlexGrid1_Scroll()
   Dim lsOldProc As String

   lsOldProc = "MSFlexGrid1_Scroll"
   ''On Error GoTo errProc

'Mac PH (09.28.12)
   With MSFlexGrid1
      If Not pbScroll Then
         If .Row = .TopRow + 17 Then
            If Not .RowIsVisible(.Row + 1) Then
               pbScroll = False
               .TopRow = .Row
            End If
         End If
      End If

      .Row = .TopRow
      pbScroll = True
      Call HiglightRow(Me.MSFlexGrid1, .Row, 1)
      pnActiveRow = .Row
   End With

   Call setFieldInfo(pbClickd)
   With txtField(10)
      .SelStart = 0
      .SelLength = Len(.Text)
      .SetFocus
   End With

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub MSFlexGrid2_Click()
   Dim lsOldProc As String

   lsOldProc = "MSFlexGrid2_Click"
   ''On Error GoTo errProc
   
   If pnROQ <= 0 Then GoTo endWithFocus

   Call setDetailInfo

endWithFocus:
   With MSFlexGrid2
      .Col = 1
      .ColSel = .Cols - 1
   End With

   With txtField(10)
      .SelStart = 0
      .SelLength = Len(.Text)
      .SetFocus
   End With
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub HiglightRow(ByVal grid As MSFlexGrid, _
                           ByVal Row As Integer, _
                           ByVal colstart As Integer)
   Dim lsOldProc As String
   Dim lnCtr As Integer

   lsOldProc = "HiglightRow"
   ''On Error GoTo errProc

   With grid
      If Row < 2 Then Exit Sub

      If Row <> pnActiveRow Then
         .Row = pnActiveRow
         For lnCtr = 1 To .Cols - 1
            .Col = lnCtr
            .CellFontBold = False
            .CellForeColor = &H0&
         Next

         .Row = Row
         For lnCtr = 1 To .Cols - 1
            .Col = lnCtr
            .CellFontBold = True
            .CellForeColor = &H8000000D
         Next
      End If
   End With

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub MSFlexGrid2_DblClick()
   Dim lsOldProc As String

   lsOldProc = "MSFlexGrid2_DblClick"
   ''On Error GoTo errProc
   
   If pnROQ <= 0 Then Exit Sub
   
   With MSFlexGrid2
      If .MouseRow = .Rows - 1 Then Exit Sub
   End With

   With frmCPUnitHistory
      .Brand = poRS.Fields("sBrandNme")
      .Model = poRS.Fields("sModelNme")
      .IsCPUnit = True
      
      If poRS.Fields("cCategory") = "Mod" Then
         .History = oTrans.GetHistoryModel(poRSRecOrder.Fields("sModelIDx"))
      Else
         .History = oTrans.GetHistory(poRSRecOrder.Fields("sStockIDx"))
      End If
      
      .Show vbModal
      txtField(10).SetFocus
   End With

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   Dim lsOldProc As String

   lsOldProc = "MSFlexGrid2_Click"
   ''On Error GoTo errProc
   
   With txtField(10)
'      .SetFocus
      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = oApp.getColor("HT1")
   End With

   pnIndex = Index

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String

   lsOldProc = "txtField_KeyDown"
   ''On Error GoTo errProc
   
   Select Case KeyCode
   Case vbKeyDown, vbKeyUp
      If pbControl Then KeyCode = 0
   Case vbKeyReturn
      If Index = 10 Then
         With poRSRecOrder
            If pnROQ <= 0 Then Exit Sub
            If txtField(10) = "" Then txtField(10) = "0"
            If .Fields("nQuantity") = 0 And txtField(10) = 0 Then Exit Sub
         
            'validation on inventory type
            If .Fields("cInvTypex") = "0" Or .Fields("cInvTypex") = "3" Then
               MsgBox "Unable to order phased out or stop model.", vbCritical, "Warning"
               Exit Sub
            End If
         
            .Fields("nQuantity") = IIf(IsNumeric(txtField(10)), txtField(10), 0)
            
            If .Fields("nQuantity") <> 0 And _
                  findCPOnOrder(.Fields("sModelNme"), .Fields("sColorNme")) = False Then
               Call addDetail
            ElseIf findCPOnOrder(.Fields("sModelNme"), .Fields("sColorNme")) = True Then
               If .Fields("nQuantity") = 0 Then
                  Call deleteDetail
               Else
                  Call UpdateDetail
               End If
            End If
         End With
      End If
   End Select

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub txtField_KeyPress(Index As Integer, keyascii As Integer)
   Dim lsOldProc As String

   lsOldProc = "txtField_KeyPress"
   ''On Error GoTo errProc
   
   If Index <> 10 Then Exit Sub
   Select Case keyascii
      Case vbKey0 To vbKey9
      Case vbKeyBack, vbKeyClear, vbKeyDelete
      Case vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown, vbKeyTab, vbKeyReturn
      Case Else
         keyascii = 0
         Beep
   End Select

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   Dim lsOldProc As String

   lsOldProc = "txtField_LostFocus"
   ''On Error GoTo errProc
   
   With txtField(Index)
      Select Case Index
         Case 10
            .BackColor = oApp.getColor("EB")

            If Index = 10 Then Exit Sub
      End Select
   End With
   pnIndex = 0

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   Dim lsOldProc As String

   lsOldProc = "txtField_Validate"
   ''On Error GoTo errProc
   
   Select Case Index
      Case 10
         With poRSRecOrder
            If pnROQ <= 0 Then Exit Sub
            If txtField(10) = "" Then txtField(10) = "0"
            If .Fields("nQuantity") = 0 And txtField(10) = 0 Then Exit Sub
         
            'validation on inventory type
            If .Fields("cInvTypex") = "0" Or .Fields("cInvTypex") = "3" Then
               MsgBox "Unable to order phased out or stop model.", vbCritical, "Warning"
               Exit Sub
            End If
         
            .Fields("nQuantity") = IIf(IsNumeric(txtField(10)), txtField(10), 0)
            
            If .Fields("nQuantity") <> 0 And _
                  findCPOnOrder(.Fields("sModelNme"), .Fields("sColorNme")) = False Then
               Call addDetail
            ElseIf findCPOnOrder(.Fields("sModelNme"), .Fields("sColorNme")) = True Then
               If .Fields("nQuantity") = 0 Then
                  Call deleteDetail
               Else
                  Call UpdateDetail
               End If
            End If
         End With
   End Select

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub setFieldInfo(Optional ByVal lbClick As Boolean = False)
   Dim lsOldProc As String

   lsOldProc = "setFieldInfo"
   ''On Error GoTo errProc
   
   Call ClearFields
   If Not lbClick Then
      If txtField(1) <> "" And txtField(2) <> "" And txtField(3) <> "" Then
         pbByModel = False
      ElseIf (txtField(1) <> "" And txtField(2) <> "" And txtField(3) = "") Or _
         (txtField(1) <> "" And txtField(2) = "" And txtField(3) = "") Then
         pbByModel = True
      End If
   Else
      pbByModel = False
   End If

   With poRSRecOrder
      .Move MSFlexGrid1.Row - 2, adAffectCurrent

      txtField(10) = .Fields("nQuantity")
      'MAC - Inventory Type (06-08-13)
      Select Case .Fields("cInvTypex")
         Case "0"
            lblField(7).Caption = "*Stop Model"
         Case "1"
            lblField(7).Caption = "*Active Model"
         Case "2", ""
            lblField(7).Caption = "*Push Model"
         Case "3"
            lblField(7).Caption = "*Phased Out"
      End Select
      
      If pbByModel = False Then
         txtField(1) = .Fields("sBrandNme")
         txtField(2) = .Fields("sModelNme")
         txtField(3) = .Fields("sColorNme")
         txtField(5) = .Fields("cClassify")
         txtField(6) = .Fields("nOnTranst")
         txtField(7) = .Fields("nQtyOnHnd")
         txtField(8) = .Fields("nAveMonSl")
         txtField(9) = .Fields("nRecOrder")
      Else
         txtField(5) = .Fields("cClassMdl")
         txtField(6) = .Fields("nOnTrnsMd")
         txtField(7) = .Fields("nQtyOnHMd")
         txtField(8) = .Fields("nAveMonMd")
         txtField(9) = .Fields("nRecOrdMd")
      End If

      pnActiveRow = MSFlexGrid1.Row

      Call findCPOnOrder(.Fields("sModelNme"), .Fields("sColorNme"))
   End With

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub setDetailInfo()
   Dim lsOldProc As String

   lsOldProc = "setDetailInfo"
   ''On Error GoTo errProc

   With poRS
      If .RecordCount <= 0 Then Exit Sub
      .Move MSFlexGrid2.Row - 1, adAffectCurrent

      txtField(1) = .Fields("sBrandNme")
      txtField(2) = .Fields("sModelNme")
      txtField(3) = .Fields("sColorNme")
      txtField(10) = .Fields("nQuantity")
      txtField(5) = .Fields("cClassify")
      txtField(6) = .Fields("nOnTranst")
      txtField(7) = .Fields("nQtyOnHnd")
      txtField(8) = .Fields("nAveMonSl")
      txtField(9) = .Fields("nRecOrder")

      Call findCellphone("sBrandNme = " & strParm(poRS("sBrandNme")), _
                     "sModelNme = " & strParm(poRS("sModelNme")), _
                     "sColorNme = " & strParm(poRS("sColorNme")), False)
      pnActiveRow = MSFlexGrid1.Row
   End With

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub ClearFields()
   Dim lsOldProc As String
   Dim loTxt As TextBox

   lsOldProc = "txtField_KeyPress"
   ''On Error GoTo errProc

   txtField(1) = ""
   txtField(2) = ""
   txtField(3) = ""
   txtField(5) = ""
   txtField(6) = ""
   txtField(7) = ""
   txtField(8) = ""
   txtField(9) = ""
   txtField(10) = ""
   
   lblField(7).Caption = ""

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub LoadMaster()
   Dim lsOldProc As String

   lsOldProc = "txtField_KeyPress"
   ''On Error GoTo errProc
   
   With oTrans
      txtField(0) = Format(.Master("sTransNox"), "@@@@-@@-@@@@@@")
      txtField(11) = strLongDate(.Master("dTransact"))
   End With

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub LoadRecOrder()
   Dim lnCtr As Integer
   Dim lsOldProc As String

   lsOldProc = "LoadRecOrder"
   ''On Error GoTo errProc

   With MSFlexGrid1
      .Rows = poRSRecOrder.RecordCount + 2

      If .Rows > 20 Then
         .ColWidth(1) = 2725
      Else
         .ColWidth(1) = 2975
      End If
   End With

   With poRSRecOrder
      pnROQ = .RecordCount
      
      If Not pbLoaded Then
         If pnROQ <= 0 Then
            MsgBox "No System Recommendation for Stock Order.", vbInformation, "Notice"
            GoTo endProc
         End If
      Else
         If pnROQ <= 0 Then
            GoTo endProc
         End If
      End If
   
      .MoveFirst
      lnCtr = 0
      Do Until .EOF
         DoEvents
         MSFlexGrid1.TextMatrix(lnCtr + 2, 0) = lnCtr + 1
         MSFlexGrid1.TextMatrix(lnCtr + 2, 1) = .Fields("sBrandNme")
         MSFlexGrid1.TextMatrix(lnCtr + 2, 2) = .Fields("sModelNme")
         MSFlexGrid1.TextMatrix(lnCtr + 2, 3) = .Fields("sColorNme")
         MSFlexGrid1.TextMatrix(lnCtr + 2, 4) = .Fields("nAveMonMd")
         MSFlexGrid1.TextMatrix(lnCtr + 2, 5) = .Fields("nAveMonsl")
         MSFlexGrid1.TextMatrix(lnCtr + 2, 6) = .Fields("nRecOrdMd")
         MSFlexGrid1.TextMatrix(lnCtr + 2, 7) = .Fields("nRecOrder")
         MSFlexGrid1.TextMatrix(lnCtr + 2, 8) = .Fields("nQtyOnHMd")
         MSFlexGrid1.TextMatrix(lnCtr + 2, 9) = .Fields("nQtyOnHnd")
         lnCtr = lnCtr + 1
         .MoveNext
      Loop
      
      .MoveFirst
      
'iMac 2016.09.13
'  changed the way i highlight the detail
'      pnActiveRow = 0
'      Call HiglightRow(MSFlexGrid1, 2, 1)

      MSFlexGrid1.Row = MSFlexGrid1.Rows - 1
      MSFlexGrid1_Click
   End With

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub
Private Sub findCellphone(ByVal sBrandNme As String, _
                              Optional ByVal sModelNme As String = "", _
                              Optional ByVal sColorNme As String = "", _
                              Optional ByVal bDisplayV As Boolean = True)
   Dim lnPos As Integer
   Dim lsOldProc As String

   lsOldProc = "findCellphone"
   ''On Error GoTo errProc

   With poRSRecOrder
      Call .Find(sBrandNme, 0, adSearchForward, 1)
      If Not .EOF Then
         lnPos = .AbsolutePosition
      Else
         lnPos = 1
      End If

      If sModelNme <> "" Then
         Call .Find(sModelNme, lnPos - 1, adSearchForward, 1)
         If Not .EOF Then
            lnPos = .AbsolutePosition
         Else
            lnPos = lnPos
         End If


         If sColorNme <> "" Then
            Call .Find(sColorNme, lnPos - 1, adSearchForward, 1)
            If Not .EOF Then
               lnPos = .AbsolutePosition
            Else
               lnPos = lnPos
            End If
         End If
      End If

      .Cancel
   End With

   With MSFlexGrid1
      .Row = lnPos + 1
      If .Row > 20 Then .TopRow = .Row - 18
   End With

   Call HiglightRow(Me.MSFlexGrid1, MSFlexGrid1.Row, 1)
   pbClickd = False
   If bDisplayV Then Call setFieldInfo

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Function findCPOnOrder(ByVal sModelNme As String, _
                              ByVal sColorNme As String) As Boolean

   Dim lnPos As Integer
   Dim lsOldProc As String

   lsOldProc = "findCPOnOrder"
   ''On Error GoTo errProc

   If TypeName(poRS) = "Nothing" Then Exit Function
   
   With poRS
      Call .Find("sModelNme = " & strParm(sModelNme), 0, adSearchForward, adBookmarkFirst)
      If Not .EOF Then
         lnPos = .AbsolutePosition

         Call .Find("sColorNme = " & strParm(sColorNme), lnPos - 1, adSearchForward, adBookmarkFirst)
         If Not .EOF Then
            If .Fields("sModelNme") = sModelNme Then
               lnPos = .AbsolutePosition
               findCPOnOrder = True
            Else
               lnPos = 0
            End If
         Else
            lnPos = 0
         End If
      Else
         lnPos = 0
      End If
      .Cancel
   End With

   With MSFlexGrid2
      .Row = IIf(lnPos = 0, .Rows - 1, lnPos)
      If .Row > 10 Then .TopRow = .Row - 9
      .Col = 1
      .ColSel = .Cols - 1
   End With

endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )", True
End Function

Private Function SaveTransaction() As Boolean
   Dim lnCtr As Integer
   Dim lsOldProc As String

   lsOldProc = "SaveTransaction"
   ''On Error GoTo errProc

   If poRS.RecordCount = 0 Then GoTo endProc

   'pass detail to class
   With oTrans
      poRS.MoveFirst
      For lnCtr = 0 To poRS.RecordCount - 1
         If poRS.Fields("sBrandIDx") <> "" Then
            .Detail(lnCtr, "sBrandIDx") = poRS("sBrandIDx")
            .Detail(lnCtr, "sModelIDx") = poRS("sModelIDx")
            .Detail(lnCtr, "sColorIDx") = poRS("sColorIDx")
            .Detail(lnCtr, "nQuantity") = poRS("nQuantity")
            .Detail(lnCtr, "nRecOrder") = poRS("nRecOrder")
            .Detail(lnCtr, "nQtyOnHnd") = poRS("nQtyOnHnd")
            .Detail(lnCtr, "cClassify") = poRS("cClassify")
            .Detail(lnCtr, "sStockIDx") = poRS("sStockIDx")
            .addDetail
         End If
         
         poRS.MoveNext
      Next
      
      'save
      SaveTransaction = .SaveTransaction
      Set poRS = Nothing
      Set poRSRecOrder = Nothing
   End With

endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )", True
End Function

Private Function InitTransaction() As Boolean
   Dim lsOldProc As String

   lsOldProc = "InitTransaction"
   ''On Error GoTo errProc
   
   pbLoaded = False
   ClearFields
   InitGrid
   InitGrid2

   oTrans.InitTransaction
   oTrans.NewTransaction
   Set poRSRecOrder = oTrans.LoadRecOrderUnits
   Set poRSOrigROQ = oTrans.LoadRecOrderUnits
   
   'Mac PH (09.04.12)
   If TypeName(poRSRecOrder) = "Nothing" Then Exit Function

   Call FilterRecOrder

   LoadMaster
   LoadRecOrder

   Set poRS = New Recordset
   With poRS.Fields
      .Append "sBrandIDx", adVarChar, 7
      .Append "sModelIDx", adVarChar, 9
      .Append "sColorIDx", adVarChar, 7
      .Append "nQuantity", adInteger
      .Append "nRecOrder", adInteger
      .Append "nQtyOnHnd", adInteger
      .Append "cClassify", adChar, 1
      .Append "sStockIDx", adVarChar, 12
      .Append "nAveMonsl", adInteger
      .Append "nOnTranst", adInteger
      .Append "sBrandNme", adVarChar, 30
      .Append "sModelNme", adVarChar, 30
      .Append "sColorNme", adVarChar, 30
      .Append "cCategory", adChar, 3
      poRS.Open
   End With

   pbLoaded = True
   pbInitialized = True

   cmdButton(3).Visible = False
   xrFrame1.Enabled = True
   txtField(1).SetFocus

   InitTransaction = True

endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )", True
End Function

Private Sub FilterRecOrder()
   Dim lsOldProc As String

   lsOldProc = "FilterRecOrder"
   ''On Error GoTo errProc
   
   pnCtr = 0
   
   With poRSRecOrder
      .MoveFirst

      Do Until .EOF
         If .Fields("nRecOrder") <= 0 Or .Fields("nRecOrdMd") <= 0 Then
            .Delete
         End If

         .MoveNext
      Loop
   End With

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub oTrans_OthersRetrieved(ByVal Index As Integer, ByVal Value As Variant)
   Dim lsCondition As String
   Dim lsOldProc As String

   lsOldProc = "oTrans_OthersRetrieved"
   ''On Error GoTo errProc

   txtField(Index) = Value

   Select Case Index
      Case 1
         txtField(2) = ""
         txtField(3) = ""
      Case 2
         txtField(3) = ""
   End Select

   Call findCellphone("sBrandNme = " & strParm(txtField(1)), _
                        IIf(txtField(2) <> "", "sModelNme = " & strParm(txtField(2)), ""), _
                        IIf(txtField(3) <> "", "sColorNme = " & strParm(txtField(3)), ""))

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
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
