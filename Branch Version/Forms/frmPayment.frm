VERSION 5.00
Begin VB.Form frmPayment 
   BorderStyle     =   0  'None
   Caption         =   "Payment"
   ClientHeight    =   7815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10095
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "frmPayment.frx":0000
   ScaleHeight     =   7815
   ScaleWidth      =   10095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtField 
      Height          =   360
      Index           =   11
      Left            =   6315
      TabIndex        =   30
      Text            =   "0000-0000-0000-0000"
      Top             =   5160
      Width           =   1695
   End
   Begin VB.TextBox txtField 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   12
      Left            =   8280
      TabIndex        =   27
      Text            =   "000,000.00"
      Top             =   5160
      Width           =   1515
   End
   Begin VB.TextBox txtField 
      Height          =   360
      Index           =   10
      Left            =   6315
      TabIndex        =   22
      Text            =   "0000-0000-0000-0000"
      Top             =   4755
      Width           =   1695
   End
   Begin VB.TextBox txtField 
      Height          =   360
      Index           =   9
      Left            =   6315
      TabIndex        =   21
      Top             =   4350
      Width           =   1695
   End
   Begin VB.TextBox txtField 
      Height          =   360
      Index           =   8
      Left            =   6315
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   3945
      Width           =   3480
   End
   Begin VB.TextBox txtField 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   7
      Left            =   3345
      TabIndex        =   18
      Text            =   "000,000.00"
      Top             =   5160
      Width           =   1515
   End
   Begin VB.TextBox txtField 
      Height          =   360
      Index           =   4
      Left            =   1380
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   3945
      Width           =   3480
   End
   Begin VB.TextBox txtField 
      Height          =   360
      Index           =   5
      Left            =   3345
      TabIndex        =   16
      Text            =   "December 31, 2008"
      Top             =   4350
      Width           =   1515
   End
   Begin VB.TextBox txtField 
      Height          =   360
      Index           =   6
      Left            =   3345
      TabIndex        =   15
      Text            =   "0000000000"
      Top             =   4755
      Width           =   1515
   End
   Begin VB.TextBox txtField 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   1
      Left            =   2280
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   1890
      Width           =   7590
   End
   Begin VB.TextBox txtField 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Index           =   3
      Left            =   6315
      TabIndex        =   8
      Text            =   "0.00"
      Top             =   2820
      Width           =   3555
   End
   Begin VB.TextBox txtField 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   165
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   885
      Width           =   3105
   End
   Begin VB.TextBox txtField 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   2
      Left            =   2280
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   2355
      Width           =   7590
   End
   Begin VB.Label lblField 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "F6-Cancel"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   420
      Index           =   22
      Left            =   8700
      TabIndex        =   37
      Top             =   7320
      Width           =   1305
   End
   Begin VB.Label lblField 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "F1-Help"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   420
      Index           =   21
      Left            =   1725
      TabIndex        =   36
      Top             =   7320
      Width           =   1305
   End
   Begin VB.Label lblField 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "F2-Find"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   420
      Index           =   20
      Left            =   3120
      TabIndex        =   35
      Top             =   7320
      Width           =   1305
   End
   Begin VB.Label lblField 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "F5-Ok"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   420
      Index           =   17
      Left            =   7305
      TabIndex        =   34
      Top             =   7320
      Width           =   1305
   End
   Begin VB.Label lblField 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "F4-Card"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   420
      Index           =   18
      Left            =   5910
      TabIndex        =   33
      Top             =   7320
      Width           =   1305
   End
   Begin VB.Label lblField 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "F3-Check"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   420
      Index           =   19
      Left            =   4515
      TabIndex        =   32
      Top             =   7320
      Width           =   1305
   End
   Begin VB.Label lblField 
      BackColor       =   &H001778E7&
      BackStyle       =   0  'Transparent
      Caption         =   "Approval No:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Index           =   14
      Left            =   5190
      TabIndex        =   31
      Top             =   5235
      Width           =   1185
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   4
      X1              =   9870
      X2              =   9870
      Y1              =   3735
      Y2              =   5655
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H001778E7&
      BackStyle       =   0  'Transparent
      Caption         =   "Change Amount:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   555
      Index           =   16
      Left            =   3795
      TabIndex        =   29
      Top             =   6060
      Width           =   2460
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   3
      X1              =   11640
      X2              =   11640
      Y1              =   3690
      Y2              =   5610
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   2
      X1              =   5100
      X2              =   5100
      Y1              =   3870
      Y2              =   5655
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   6
      X1              =   9870
      X2              =   5085
      Y1              =   5655
      Y2              =   5655
   End
   Begin VB.Label lblField 
      BackColor       =   &H001778E7&
      BackStyle       =   0  'Transparent
      Caption         =   " Amount:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Index           =   15
      Left            =   8235
      TabIndex        =   28
      Top             =   4935
      Width           =   1515
   End
   Begin VB.Label lblField 
      BackColor       =   &H001778E7&
      BackStyle       =   0  'Transparent
      Caption         =   "Credit Card"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Index           =   10
      Left            =   5085
      TabIndex        =   26
      Top             =   3615
      Width           =   1170
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   5
      X1              =   9870
      X2              =   6330
      Y1              =   3750
      Y2              =   3750
   End
   Begin VB.Label lblField 
      BackColor       =   &H001778E7&
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Name:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Index           =   11
      Left            =   5205
      TabIndex        =   25
      Top             =   3990
      Width           =   1065
   End
   Begin VB.Label lblField 
      BackColor       =   &H001778E7&
      BackStyle       =   0  'Transparent
      Caption         =   "Card No:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Index           =   13
      Left            =   5190
      TabIndex        =   24
      Top             =   4815
      Width           =   900
   End
   Begin VB.Label lblField 
      BackColor       =   &H001778E7&
      BackStyle       =   0  'Transparent
      Caption         =   "Card Type:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Index           =   12
      Left            =   5190
      TabIndex        =   23
      Top             =   4410
      Width           =   1065
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   4995
      X2              =   4995
      Y1              =   3735
      Y2              =   5655
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   0
      X1              =   225
      X2              =   225
      Y1              =   3870
      Y2              =   5655
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   4
      X1              =   4995
      X2              =   210
      Y1              =   5655
      Y2              =   5640
   End
   Begin VB.Label lblField 
      BackColor       =   &H001778E7&
      BackStyle       =   0  'Transparent
      Caption         =   "Check Amt.:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Index           =   9
      Left            =   2235
      TabIndex        =   19
      Top             =   5235
      Width           =   1065
   End
   Begin VB.Line Line7 
      BorderColor     =   &H0013B8FD&
      BorderWidth     =   3
      Index           =   2
      X1              =   10065
      X2              =   10065
      Y1              =   465
      Y2              =   6885
   End
   Begin VB.Line Line7 
      BorderColor     =   &H0013B8FD&
      BorderWidth     =   3
      Index           =   0
      X1              =   15
      X2              =   15
      Y1              =   1575
      Y2              =   6870
   End
   Begin VB.Label lblField 
      BackColor       =   &H001778E7&
      BackStyle       =   0  'Transparent
      Caption         =   "Check Date:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Index           =   7
      Left            =   2235
      TabIndex        =   13
      Top             =   4410
      Width           =   1065
   End
   Begin VB.Label lblField 
      BackColor       =   &H001778E7&
      BackStyle       =   0  'Transparent
      Caption         =   "Check No:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Index           =   8
      Left            =   2235
      TabIndex        =   12
      Top             =   4815
      Width           =   900
   End
   Begin VB.Label lblField 
      BackColor       =   &H001778E7&
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Name:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Index           =   6
      Left            =   285
      TabIndex        =   11
      Top             =   3990
      Width           =   1065
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   3
      X1              =   4995
      X2              =   1335
      Y1              =   3750
      Y2              =   3750
   End
   Begin VB.Label lblField 
      BackColor       =   &H001778E7&
      BackStyle       =   0  'Transparent
      Caption         =   "Check Info"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Index           =   5
      Left            =   195
      TabIndex        =   10
      Top             =   3615
      Width           =   1170
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      Height          =   1005
      Index           =   0
      Left            =   15
      Top             =   465
      Width           =   3375
   End
   Begin VB.Line Line4 
      BorderColor     =   &H0013B8FD&
      BorderWidth     =   3
      Index           =   2
      X1              =   10140
      X2              =   30
      Y1              =   6870
      Y2              =   6870
   End
   Begin VB.Line Line4 
      BorderColor     =   &H0013B8FD&
      BorderWidth     =   3
      Index           =   1
      X1              =   3510
      X2              =   15
      Y1              =   1590
      Y2              =   1590
   End
   Begin VB.Line Line7 
      BorderColor     =   &H0013B8FD&
      BorderWidth     =   3
      Index           =   1
      X1              =   3510
      X2              =   3510
      Y1              =   435
      Y2              =   1590
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H001778E7&
      BackStyle       =   0  'Transparent
      Caption         =   "Cash Amount:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   555
      Index           =   3
      Left            =   3795
      TabIndex        =   9
      Top             =   2940
      Width           =   2460
   End
   Begin VB.Label lblField 
      BackColor       =   &H001778E7&
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Index           =   2
      Left            =   300
      TabIndex        =   7
      Top             =   2445
      Width           =   1890
   End
   Begin VB.Label lblField 
      BackColor       =   &H001778E7&
      BackStyle       =   0  'Transparent
      Caption         =   "Name/Barcode:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Index           =   1
      Left            =   300
      TabIndex        =   6
      Top             =   1935
      Width           =   1875
   End
   Begin VB.Line Line4 
      BorderColor     =   &H0013B8FD&
      BorderWidth     =   3
      Index           =   0
      X1              =   10500
      X2              =   3495
      Y1              =   450
      Y2              =   450
   End
   Begin VB.Label lblField 
      BackColor       =   &H001778E7&
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Person:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Index           =   0
      Left            =   165
      TabIndex        =   5
      Top             =   585
      Width           =   1485
   End
   Begin VB.Label lblTotalAmount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000,000.00"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1035
      Left            =   6330
      TabIndex        =   4
      Top             =   600
      Width           =   3555
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H001778E7&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   555
      Index           =   4
      Left            =   3795
      TabIndex        =   2
      Top             =   660
      Width           =   2460
   End
   Begin VB.Label lblChangeAmount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000,000.00"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   750
      Left            =   6315
      TabIndex        =   1
      Top             =   5865
      Width           =   3555
   End
   Begin VB.Image Image1 
      Height          =   11520
      Left            =   0
      Picture         =   "frmPayment.frx":8CC6
      Top             =   375
      Width           =   15360
   End
End
Attribute VB_Name = "frmPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME As String = "frmPayment"
Private WithEvents oTrans As clsCPSales
Attribute oTrans.VB_VarHelpID = -1

Private Enum xePaymentType
   xeCashOnly = 0
   xeOthers = 1
End Enum

Private oSkin As clsFormSkin
Dim pnCtr As Integer
Dim pbCancelled As Boolean

Property Set PaymentTrans(loPayment As clsCPSales)
   Set oTrans = loPayment
End Property

Private Sub Command1_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0
End Sub

Property Get Cancelled() As Boolean
   Cancelled = pbCancelled
End Property

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyReturn, vbKeyUp, vbKeyDown
      Select Case KeyCode
      Case vbKeyReturn, vbKeyDown
         SetNextFocus
      Case vbKeyUp
         SetPreviousFocus
      End Select
   Case vbKeyF6
      pbCancelled = True
      Unload Me
   End Select
End Sub

Private Sub Form_Load()
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.DisableClose = True
   oSkin.ApplySkin xeFormMaintenance
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
End Sub

Private Function isEntryOK() As Boolean

EntryOK:
   isEntryOK = True
   Exit Function
EntryNotOK:
   isEntryOK = False
End Function

Private Sub ShowOtherInfo(lnStat As Integer)
   Dim lbShow As Boolean

   lbShow = IIf(lnStat = 0, False, True)
   
'   lblField(5).Visible = lbShow
'   lblField(6).Visible = lbShow
'   lblField(7).Visible = lbShow
'   lblField(8).Visible = lbShow
'   lblField(9).Visible = lbShow
'   lblField(10).Visible = lbShow
'   lblField(11).Visible = lbShow
'   lblField(12).Visible = lbShow
'   lblField(13).Visible = lbShow
'   lblField(14).Visible = lbShow
'   lblField(15).Visible = lbShow
'
'   txtField(4).Visible = lbShow
'   txtField(5).Visible = lbShow
'   txtField(6).Visible = lbShow
'   txtField(7).Visible = lbShow
'   txtField(8).Visible = lbShow
'   txtField(9).Visible = lbShow
'   txtField(10).Visible = lbShow
'   txtField(11).Visible = lbShow
'   txtField(12).Visible = lbShow
   
   txtField(4).Enabled = lbShow
   txtField(5).Enabled = lbShow
   txtField(6).Enabled = lbShow
   txtField(7).Enabled = lbShow
   txtField(8).Enabled = lbShow
   txtField(9).Enabled = lbShow
   txtField(10).Enabled = lbShow
   txtField(11).Enabled = lbShow
   txtField(12).Enabled = lbShow
   
   Line1(0).Visible = lbShow
   Line1(1).Visible = lbShow
   Line4(3).Visible = lbShow
   Line4(4).Visible = lbShow

   Line1(2).Visible = lbShow
   Line1(4).Visible = lbShow
   Line4(5).Visible = lbShow
   Line4(6).Visible = lbShow

   
   If lbShow Then
      Line4(2).Y1 = 6870
      Line4(2).Y2 = 6855

      Line7(0).Y2 = 6870
      Line7(2).Y2 = 6885

      lblChangeAmount.Top = 5865
      lblField(16).Top = 6060
      Me.Height = 6900
   Else
      Line4(2).Y1 = 4500
      Line4(2).Y2 = 4485

      Line7(0).Y2 = 4470
      Line7(2).Y2 = 4470

      lblChangeAmount.Top = 3525
      lblField(16).Top = 3720
      Me.Height = 4530
   End If
End Sub

Private Sub EnableCheckInfo(lnStat As Integer)
   Dim lbShow As Boolean

   lbShow = IIf(lnStat = 0, False, True)
   
   txtField(4).Enabled = lbShow
   txtField(5).Enabled = lbShow
   txtField(6).Enabled = lbShow
   txtField(7).Enabled = lbShow
   
   If lbShow Then
      txtField(4).BackColor = lbShow
      txtField(5).BackColor = lbShow
      txtField(6).BackColor = lbShow
      txtField(7).BackColor = lbShow
      
      lblField(5).ForeColor = lbShow
      lblField(6).ForeColor = lbShow
      lblField(7).ForeColor = lbShow
      lblField(8).ForeColor = lbShow
      lblField(9).ForeColor = lbShow

   Else
      txtField(4).BackColor = lbShow
      txtField(5).BackColor = lbShow
      txtField(6).BackColor = lbShow
      txtField(7).BackColor = lbShow
      
      lblField(5).ForeColor = lbShow
      lblField(6).ForeColor = lbShow
      lblField(7).ForeColor = lbShow
      lblField(8).ForeColor = lbShow
      lblField(9).ForeColor = lbShow
   End If
End Sub

Private Sub EnableCardInfo(lnStat As Integer)
   Dim lbShow As Boolean

   lbShow = IIf(lnStat = 0, False, True)
   
   txtField(8).Enabled = lbShow
   txtField(9).Enabled = lbShow
   txtField(10).Enabled = lbShow
   txtField(11).Enabled = lbShow
   txtField(12).Enabled = lbShow
   
   If lbShow Then
      txtField(8).BackColor = lbShow
      txtField(9).BackColor = lbShow
      txtField(10).BackColor = lbShow
      txtField(11).BackColor = lbShow
      txtField(12).BackColor = lbShow
      
      lblField(10).ForeColor = lbShow
      lblField(11).ForeColor = lbShow
      lblField(12).ForeColor = lbShow
      lblField(13).ForeColor = lbShow
      lblField(14).ForeColor = lbShow
      lblField(15).ForeColor = lbShow
   Else
      txtField(8).BackColor = lbShow
      txtField(9).BackColor = lbShow
      txtField(10).BackColor = lbShow
      txtField(11).BackColor = lbShow
      txtField(12).BackColor = lbShow
      
      lblField(10).ForeColor = lbShow
      lblField(11).ForeColor = lbShow
      lblField(12).ForeColor = lbShow
      lblField(13).ForeColor = lbShow
      lblField(14).ForeColor = lbShow
      lblField(15).ForeColor = lbShow
   End If
End Sub

Private Sub oTrans_MasterRetrieved(ByVal Index As Integer)
   Select Case Index
   Case 3
      txtField(1).Text = oTrans.Master(Index)
   Case 18
      txtField(0).Text = Format(oApp.getLogName(oTrans.Master(Index)), ">")
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

