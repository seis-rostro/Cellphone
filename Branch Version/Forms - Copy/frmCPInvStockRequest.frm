VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmCPInvStockRequest 
   BorderStyle     =   0  'None
   Caption         =   "Stock Request w/ ROQ Computation"
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
      TabIndex        =   23
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
      Picture         =   "frmCPInvStockRequest.frx":0000
   End
   Begin xrControl.xrFrame xrFrame3 
      Height          =   3120
      Left            =   6825
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   5503
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   3030
         Left            =   30
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   30
         Width           =   8205
         _ExtentX        =   14473
         _ExtentY        =   5345
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
      Height          =   5505
      Left            =   1575
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   3675
      Width           =   13560
      _ExtentX        =   23918
      _ExtentY        =   9710
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
      TabIndex        =   28
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
      Picture         =   "frmCPInvStockRequest.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   90
      TabIndex        =   27
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
      Picture         =   "frmCPInvStockRequest.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   90
      TabIndex        =   26
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
      Picture         =   "frmCPInvStockRequest.frx":1F86
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   3120
      Left            =   1605
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   5503
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.Frame Frame1 
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   105
         TabIndex        =   29
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
            TabIndex        =   22
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
         TabIndex        =   15
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
         TabIndex        =   19
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
         TabIndex        =   17
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
         TabIndex        =   13
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
         TabIndex        =   11
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
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1132
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
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1500
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
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   750
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
         TabIndex        =   21
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
         TabIndex        =   1
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
         TabIndex        =   3
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
         TabIndex        =   14
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
         TabIndex        =   18
         Top             =   2370
         Width           =   405
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
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
         TabIndex        =   16
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
         TabIndex        =   12
         Top             =   1980
         Width           =   405
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
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
         TabIndex        =   10
         Top             =   1980
         Width           =   750
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BarrCode"
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
         Left            =   165
         TabIndex        =   6
         Top             =   1200
         Width           =   795
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descript"
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
         Left            =   270
         TabIndex        =   8
         Top             =   1568
         Width           =   690
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
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
         Index           =   15
         Left            =   465
         TabIndex        =   4
         Top             =   818
         Width           =   495
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
         TabIndex        =   20
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
         TabIndex        =   2
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
         TabIndex        =   0
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
Attribute VB_Name = "frmCPInvStockRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pxeMODULENAME = "frmCPInvStockRequest"
Private Const pxeVisibleRow = 20

Private oSkin As clsFormSkin
Private WithEvents oTrans As clsCPStockOrder
Attribute oTrans.VB_VarHelpID = -1
Private poRSRecOrder As Recordset
Private poRSOrigROQ As Recordset
Private poRS As Recordset

Private pnActiveRow As Integer
Private pbControl As Boolean
Private pnIndex As Integer
Private pnROQ As Integer
Private pnCtr As Integer
Private pbByModel As Boolean
Private pbLoaded As Boolean
Private pbScroll As Boolean
Private pbClickd As Boolean

Private pbInitTransaction As Boolean

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
         If Not pbInitTransaction Then Exit Sub
         If pnROQ <= 0 Then Exit Sub
         With frmCPOrderHistory
            .Brand = poRSRecOrder.Fields("sBrandNme")
            .Description = poRSRecOrder.Fields("sDescript")
            .History = oTrans.GetHistory(poRSRecOrder.Fields("sStockIDx"))
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

   If pbLoaded Then Exit Sub
   
   lsOldProc = "Form_Activate"
   ''On Error GoTo errProc

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
   With MSFlexGrid1
      Select Case KeyCode
         Case vbKeyReturn, vbKeyDown
            If GetFocus = MSFlexGrid2.hwnd Then

            Else
               If pnIndex = 10 And pnActiveRow < .Rows - 1 Then
                  ' this does not trigger lost focus or validate
                  If pnIndex = 10 Then
                     Call txtField_KeyDown(pnIndex, vbKeyReturn, 0)
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
                     Call txtField_KeyDown(pnIndex, vbKeyReturn, 0)
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
         
            With frmSearchStock
               Set .ROQ = poRSOrigROQ
               .Show vbModal
               Call LoadRecOrder
            End With
         Case vbKeyControl
            pbControl = True
            KeyCode = 0
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

   Set oTrans = New clsCPStockOrder
   Set oTrans.AppDriver = oApp
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
      poRS("sBarrCode") = .Fields("sBarrCode")
      poRS("sDescript") = .Fields("sDescript")
      poRS("sBrandNme") = .Fields("sBrandNme")
      poRS("nQuantity") = .Fields("nQuantity")
      poRS("sStockIDx") = .Fields("sStockIDx")
      poRS("sBrandIDx") = .Fields("sBrandIDx")
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
      poRS.Sort = "sDescript"
      If poRS.RecordCount <> 0 Then poRS.MoveFirst
      Do Until poRS.EOF
         DoEvents
         .TextMatrix(lnCtr + 1, 0) = lnCtr + 1
         .TextMatrix(lnCtr + 1, 1) = poRS("sBarrCode")
         .TextMatrix(lnCtr + 1, 2) = poRS("sDescript")
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
   Dim lnCtr As Integer

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
      .TextMatrix(1, 1) = "Barrcode"
      .TextMatrix(1, 2) = "Description"
      .TextMatrix(1, 3) = "Brand"
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
      .ColWidth(1) = 2975
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
End Sub

Private Sub InitGrid2()
   Dim lnCtr As Integer

   With MSFlexGrid2
      .Cols = 7
      .Rows = 2

      .Clear

      .TextMatrix(0, 0) = "No."
      .TextMatrix(0, 1) = "BarrCode"
      .TextMatrix(0, 2) = "Descript"
      .TextMatrix(0, 3) = "QOH"
      .TextMatrix(0, 4) = "AMC"
      .TextMatrix(0, 5) = "ROQ"
      .TextMatrix(0, 6) = "QTY"
      .TextMatrix(1, 0) = "1"

      .Row = 0
      'Column Width
      .ColWidth(0) = 442
      .ColWidth(1) = 2750 '2490
      .ColWidth(2) = 2160
      .ColWidth(3) = 700
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
      .ColAlignment(3) = flexAlignRightCenter
      .ColAlignment(4) = flexAlignRightCenter
      .ColAlignment(5) = flexAlignRightCenter
      .ColAlignment(6) = flexAlignRightCenter

      If Not pbLoaded Then
         .Col = 1
         .ColSel = .Cols - 1
      End If
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
   pnActiveRow = 0
End Sub

Private Sub MSFlexGrid1_Click()
   pbClickd = True
   
   If pnROQ <= 0 Then GoTo endProc
   poRSRecOrder.Move MSFlexGrid1.Row - 2, adBookmarkFirst
   
   Call setFieldInfo(True)

endProc:
   Call HiglightRow(Me.MSFlexGrid1, MSFlexGrid1.Row, 1)

   With txtField(10)
      .SelStart = 0
      .SelLength = Len(.Text)
      .SetFocus
   End With
End Sub

Private Sub MSFlexGrid1_DblClick()
   If pnROQ <= 0 Then Exit Sub

   With MSFlexGrid1
      poRSRecOrder.Move .Row - 2, adBookmarkFirst
      If .MouseRow > 0 Then
         Call cmdButton_Click(1)
      End If
   End With
End Sub

Private Sub MSFlexGrid1_Scroll()
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

   With txtField(10)
      .SelStart = 0
      .SelLength = Len(.Text)
      .SetFocus
   End With
End Sub

Private Sub MSFlexGrid1_SelChange()
   With MSFlexGrid1
      Call HiglightRow(Me.MSFlexGrid1, .Row, 1)
      pnActiveRow = .Row
      
   End With
End Sub

Private Sub MSFlexGrid2_Click()
   
   With MSFlexGrid2
      .Col = 1
      .ColSel = .Cols - 1
   
      If .MouseRow = .Rows - 1 Then GoTo endWithFocus
      
      poRS.Move .Row - 1, adBookmarkFirst
      
      Call setDetailInfo

      Call findCellphone("", .TextMatrix(.Row, 1), "")
   End With
endWithFocus:

   With txtField(10)
      .SelStart = 0
      .SelLength = Len(.Text)
      .SetFocus
   End With
End Sub

Private Sub HiglightRow(ByVal grid As MSFlexGrid, _
                           ByVal Row As Integer, _
                           ByVal colstart As Integer)

   'Mac PH (09.28.12)
   Dim lnCtr As Integer

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
End Sub

Private Sub MSFlexGrid2_DblClick()
   With MSFlexGrid2
      If .MouseRow = .Rows - 1 Then Exit Sub
      
      If poRS.RecordCount > 0 Then
         poRS.Move .Row - 1, adBookmarkFirst
      Else
         Exit Sub
      End If
   End With

   With frmCPOrderHistory
      .Brand = poRS.Fields("sBrandNme")
      .Description = poRS.Fields("sDescript")
      .IsCPUnit = False
      .History = oTrans.GetHistory(poRSRecOrder.Fields("sStockIDx"))
      .Show vbModal
      txtField(10).SetFocus
   End With
End Sub

Private Sub oTrans_OthersRetrieved(ByVal Index As Integer, ByVal Value As Variant)
   Select Case Index
      Case 1
         txtField(Index) = Value
         
         Call findCellphone("", _
                              txtField(1), _
                              "")
      Case Else
         txtField(Index) = Value
   End Select
   
   
   txtField(10).SetFocus
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(10)
      .SetFocus
      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = oApp.getColor("HT1")
   End With
   
   pnIndex = Index
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode <> vbKeyReturn Then
      KeyCode = 0
      Exit Sub
   End If
   
   Select Case Index
      Case 10
         If txtField(10) = "" Then txtField(10) = "0"
         With poRSRecOrder
            If .EOF Then Exit Sub
            If .Fields("nQuantity") = 0 And txtField(10) = "0" Then Exit Sub

            'validation on inventory type
            If .Fields("cInvTypex") = "0" Or .Fields("cInvTypex") = "3" Then
               MsgBox "Unable to order phased out or stop model.", vbCritical, "Warning"
               Exit Sub
            End If

            .Fields("nQuantity") = IIf(IsNumeric(txtField(10)), txtField(10), 0)

            If .Fields("nQuantity") <> 0 And _
                  findCPOnOrder(.Fields("sBarrCode")) = False Then
               Call addDetail
            ElseIf findCPOnOrder(.Fields("sBarrCode")) = True Then
               If .Fields("nQuantity") = 0 Then
                  Call deleteDetail
               Else
                  Call UpdateDetail
               End If
            End If
         End With
   End Select
End Sub

Private Sub txtField_KeyPress(Index As Integer, keyascii As Integer)
   If Index <> 10 Then Exit Sub
   Select Case keyascii
      Case vbKey0 To vbKey9
      Case vbKeyBack, vbKeyClear, vbKeyDelete
      Case vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown, vbKeyTab, vbKeyReturn
      Case Else
         keyascii = 0
         Beep
   End Select
End Sub

'Private Sub txtField_LostFocus(Index As Integer)
'   With txtField(Index)
'      Select Case Index
'         Case 10
'            .BackColor = oApp.getColor("EB")
'
'            If Index = 10 Then Exit Sub
'      End Select
'   End With
'   pnIndex = 0
'End Sub

Private Sub setFieldInfo(Optional ByVal lbClick As Boolean = False)
   Dim lsOldProc As String

   lsOldProc = "setFieldInfo"
   ''On Error GoTo errProc
   
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
      
      txtField(3) = .Fields("sBrandNme")
      txtField(1) = .Fields("sBarrCode")
      txtField(2) = .Fields("sDescript")
      
      If pbByModel = False Then
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

      Call findCPOnOrder(.Fields("sBarrCode"))
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
      txtField(10) = .Fields("nQuantity")
      txtField(5) = .Fields("cClassify")
      txtField(6) = .Fields("nOnTranst")
      txtField(7) = .Fields("nQtyOnHnd")
      txtField(8) = .Fields("nAveMonSl")
      txtField(9) = .Fields("nRecOrder")
   End With
endProc:
   pnActiveRow = MSFlexGrid1.Row
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub ClearFields()
   Dim loTxt As TextBox
   
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
End Sub

Private Sub LoadMaster()
   With oTrans
      txtField(0) = Format(.Master("sTransNox"), "@@@@-@@-@@@@@@")
      txtField(11) = strLongDate(.Master("dTransact"))
   End With
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
'            MsgBox "Selected Item was not on ROQ.", vbInformation, "Notice"
            GoTo endProc
         End If
      End If
            
      .MoveFirst
      lnCtr = 0
      Do Until .EOF
         DoEvents
         MSFlexGrid1.TextMatrix(lnCtr + 2, 0) = lnCtr + 1
         MSFlexGrid1.TextMatrix(lnCtr + 2, 1) = .Fields("sBarrCode")
         MSFlexGrid1.TextMatrix(lnCtr + 2, 2) = .Fields("sDescript")
         MSFlexGrid1.TextMatrix(lnCtr + 2, 3) = .Fields("sBrandNme")
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
      
      pnActiveRow = lnCtr
      Call HiglightRow(MSFlexGrid1, 2, 1)
   End With

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub findCellphone(ByVal sBrandNme As String, _
                              Optional ByVal sBarrCode As String = "", _
                              Optional ByVal sDescript As String = "", _
                              Optional ByVal bDisplayV As Boolean = True)
   Dim lnPos As Integer

   With poRSRecOrder
      If sBrandNme <> "" Then
         sBrandNme = " sBrandNme LIKE " & strParm(sBrandNme & "%")
         
         Call .Find(sBrandNme, 0, adSearchForward, 1)
            If Not .EOF Then
            lnPos = .AbsolutePosition
            txtField(3) = .Fields("sBrandNme")
         End If
      End If
      
      If sBarrCode <> "" Then
         sBarrCode = " sBarrCode LIKE " & strParm("%" & sBarrCode & "%")
      
         Call .Find(sBarrCode, IIf(lnPos > 0, lnPos - 1, 0), adSearchForward, 1)
         If Not .EOF Then
            lnPos = .AbsolutePosition
            txtField(1) = .Fields("sBarrCode")
         End If
      End If
      
      If sDescript <> "" Then
         sDescript = " sDescript LIKE " & strParm("%" & sDescript & "%")
               
         Call .Find(sDescript, IIf(lnPos > 0, lnPos - 1, 0), adSearchForward, 1)
         If Not .EOF Then
            lnPos = .AbsolutePosition
            txtField(2) = .Fields("sDescript")
         End If
      End If

      .Cancel
   End With
   
   If lnPos = 0 Then Exit Sub
   With MSFlexGrid1
      .Row = lnPos + 1
      .TopRow = .Row
   End With

   Call HiglightRow(Me.MSFlexGrid1, MSFlexGrid1.Row, 1)
   pbClickd = False
   If bDisplayV Then Call setFieldInfo
End Sub

Private Function findCPOnOrder(ByVal sBarrCode As String) As Boolean

   Dim lnPos As Integer

   With poRS
      Call .Find("sBarrCode = " & strParm(sBarrCode), 0, adSearchForward, adBookmarkFirst)
      If Not .EOF Then
         lnPos = .AbsolutePosition
         
         findCPOnOrder = True
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
         If poRS.Fields("sBarrCode") = "" Then GoTo move2Next

         .Detail(lnCtr, "nQuantity") = poRS("nQuantity")
         .Detail(lnCtr, "nRecOrder") = poRS("nRecOrder")
         .Detail(lnCtr, "nQtyOnHnd") = poRS("nQtyOnHnd")
         .Detail(lnCtr, "cClassify") = poRS("cClassify")
         .Detail(lnCtr, "sStockIDx") = poRS("sStockIDx")
         .Detail(lnCtr, "sBrandIDx") = poRS("sBrandIDx")
         .addDetail
move2Next:
         poRS.MoveNext
      Next
   'save
      If .SaveTransaction Then
         SaveTransaction = True
         Set poRS = Nothing
         Set poRSRecOrder = Nothing
         Set poRSOrigROQ = Nothing
      End If
   End With

endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )", True
End Function

Private Function InitTransaction() As Boolean
   pbLoaded = False
   ClearFields
   InitGrid
   InitGrid2

   oTrans.InitTransaction
   oTrans.NewTransaction
   Set poRSRecOrder = oTrans.LoadRecOrderOthers
   Set poRSOrigROQ = oTrans.LoadRecOrderOthers
   
   If TypeName(poRSRecOrder) = "Nothing" Then Exit Function
   
   'delete entries without system recommendation
   Call FilterRecOrder
   
   LoadMaster
   LoadRecOrder

   Set poRS = New Recordset
   With poRS.Fields
      .Append "sStockIDx", adVarChar, 12
      .Append "sBarrCode", adVarChar, 20
      .Append "sDescript", adVarChar, 50
      .Append "nQuantity", adInteger
      .Append "nRecOrder", adInteger
      .Append "nQtyOnHnd", adInteger
      .Append "cClassify", adChar, 1
      .Append "nAveMonsl", adInteger
      .Append "nOnTranst", adInteger
      .Append "sBrandNme", adVarChar, 30
      .Append "cCategory", adChar, 3
      .Append "sBrandIDx", adChar, 7
      
      poRS.Open
   End With

   pbLoaded = True
   pbInitTransaction = True

   cmdButton(3).Visible = False
   xrFrame1.Enabled = True
   txtField(10).SetFocus

   InitTransaction = True
End Function

Private Sub FilterRecOrder()
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
End Sub

Private Sub AddStock(ByVal lors As Recordset)
   With lors
      If Not .EOF Then
         Call poRSRecOrder.Find("sStockIDx = " & strParm(.Fields("sStockIDx")), 0, adSearchForward, adBookmarkFirst)
         If poRSRecOrder.EOF Then
            poRSRecOrder.AddNew
            poRSRecOrder.Fields("sStockIDx") = .Fields("sStockIDx")
            poRSRecOrder.Fields("sBarrCode") = .Fields("sBarrCode")
            poRSRecOrder.Fields("sDescript") = .Fields("sDescript")
            poRSRecOrder.Fields("sBrandNme") = .Fields("sBrandNme")
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
         Else
            MsgBox "Item was already on the list.", vbInformation, "Notice"
         End If
      End If

      poRSRecOrder.Cancel
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
