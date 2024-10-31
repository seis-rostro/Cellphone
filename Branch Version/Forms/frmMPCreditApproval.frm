VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmMPCreditApproval 
   BorderStyle     =   0  'None
   Caption         =   "Credit Application Approval"
   ClientHeight    =   7500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11940
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7500
   ScaleWidth      =   11940
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   6450
      Index           =   0
      Left            =   1545
      Tag             =   "wt0;fb0"
      Top             =   975
      Width           =   10320
      _ExtentX        =   18203
      _ExtentY        =   11377
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.Frame Frame3 
         Caption         =   "Spouse Info"
         Enabled         =   0   'False
         Height          =   1335
         Left            =   90
         TabIndex        =   55
         Tag             =   "wt0;fb0"
         Top             =   4995
         Width           =   10095
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   87
            Left            =   1125
            TabIndex        =   42
            Top             =   240
            Width           =   3660
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   585
            Index           =   90
            Left            =   1125
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   540
            Width           =   3660
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   88
            Left            =   6375
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   240
            Width           =   1635
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   89
            Left            =   6375
            Locked          =   -1  'True
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   540
            Width           =   3345
         End
         Begin VB.TextBox txtOthers 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   8730
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtWaysMn 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Index           =   85
            Left            =   6375
            MaxLength       =   50
            TabIndex        =   52
            Top             =   855
            Width           =   3345
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            Height          =   195
            Index           =   23
            Left            =   120
            TabIndex        =   41
            Top             =   285
            Width           =   420
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            Height          =   195
            Index           =   22
            Left            =   120
            TabIndex        =   43
            Top             =   585
            Width           =   570
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Birth Date"
            Height          =   195
            Index           =   21
            Left            =   5340
            TabIndex        =   45
            Top             =   285
            Width           =   705
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Birth Place"
            Height          =   195
            Index           =   20
            Left            =   5340
            TabIndex        =   49
            Top             =   585
            Width           =   765
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Age"
            Height          =   195
            Index           =   19
            Left            =   8145
            TabIndex        =   47
            Top             =   285
            Width           =   285
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Occupation"
            Height          =   195
            Index           =   5
            Left            =   5355
            TabIndex        =   51
            Top             =   915
            Width           =   825
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " Application Information "
         Enabled         =   0   'False
         Height          =   2820
         Left            =   75
         TabIndex        =   54
         Tag             =   "wt0;fb0"
         Top             =   2145
         Width           =   10125
         Begin VB.CheckBox chkField 
            Caption         =   "With Financer"
            Height          =   285
            Left            =   6930
            TabIndex        =   40
            Tag             =   "wt0;fb0"
            Top             =   2400
            Width           =   1470
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   99
            Left            =   6930
            TabIndex        =   37
            Top             =   1785
            Width           =   2805
         End
         Begin VB.TextBox txtField 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Index           =   14
            Left            =   6930
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   2085
            Width           =   1470
         End
         Begin VB.TextBox txtField 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   12
            Left            =   6930
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   35
            TabStop         =   0   'False
            Text            =   "0.00"
            Top             =   1485
            Width           =   2805
         End
         Begin VB.TextBox txtField 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   9
            Left            =   1575
            MaxLength       =   50
            TabIndex        =   29
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   2415
            Width           =   2805
         End
         Begin VB.TextBox txtField 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   8
            Left            =   1575
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   27
            TabStop         =   0   'False
            Text            =   "0.00"
            Top             =   2115
            Width           =   2805
         End
         Begin VB.TextBox txtField 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   13
            Left            =   6930
            MaxLength       =   50
            TabIndex        =   33
            Top             =   1170
            Width           =   1500
         End
         Begin VB.TextBox txtField 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   10
            Left            =   6930
            MaxLength       =   50
            TabIndex        =   31
            Text            =   "0.00"
            Top             =   855
            Width           =   2805
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   97
            Left            =   1575
            MaxLength       =   50
            TabIndex        =   25
            Top             =   1815
            Width           =   2805
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   7
            Left            =   1575
            MaxLength       =   50
            TabIndex        =   23
            Top             =   1500
            Width           =   2805
         End
         Begin VB.ComboBox cmbField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   1
            ItemData        =   "frmMPCreditApproval.frx":0000
            Left            =   1560
            List            =   "frmMPCreditApproval.frx":0002
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   1170
            Width           =   2805
         End
         Begin VB.ComboBox cmbField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   0
            ItemData        =   "frmMPCreditApproval.frx":0004
            Left            =   1575
            List            =   "frmMPCreditApproval.frx":0006
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   855
            Width           =   2805
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   3
            Left            =   1575
            MaxLength       =   50
            TabIndex        =   17
            Top             =   240
            Width           =   2805
         End
         Begin xrControl.xrButton xrButton1 
            Height          =   285
            Left            =   8535
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   1170
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   503
            Caption         =   "Auto-C&ompute"
            AccessKey       =   "o"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
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
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   6930
            TabIndex        =   61
            Tag             =   "eb0;et0"
            Top             =   315
            Width           =   2850
         End
         Begin VB.Shape Shape3 
            Height          =   390
            Index           =   0
            Left            =   6900
            Top             =   240
            Width           =   2925
         End
         Begin VB.Shape Shape4 
            Height          =   330
            Index           =   0
            Left            =   6930
            Top             =   270
            Width           =   2865
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Credit Investigator"
            Height          =   195
            Index           =   95
            Left            =   5550
            TabIndex        =   36
            Top             =   1815
            Width           =   1275
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quick Match No."
            Height          =   195
            Index           =   12
            Left            =   5550
            TabIndex        =   38
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Monthly Amort."
            Height          =   195
            Index           =   10
            Left            =   5550
            TabIndex        =   34
            Top             =   1485
            Width           =   1050
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PN Value"
            Height          =   195
            Index           =   7
            Left            =   270
            TabIndex        =   28
            Top             =   2460
            Width           =   675
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Gross Price"
            Height          =   195
            Index           =   6
            Left            =   270
            TabIndex        =   26
            Top             =   2160
            Width           =   810
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Term"
            Height          =   195
            Index           =   11
            Left            =   5550
            TabIndex        =   32
            Top             =   1200
            Width           =   360
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Down Payment"
            Height          =   195
            Index           =   8
            Left            =   5550
            TabIndex        =   30
            Top             =   915
            Width           =   1080
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Model"
            Height          =   195
            Index           =   9
            Left            =   270
            TabIndex        =   24
            Top             =   1815
            Width           =   435
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Unit Applied"
            Height          =   195
            Index           =   3
            Left            =   270
            TabIndex        =   20
            Top             =   1200
            Width           =   855
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Application Type"
            Height          =   195
            Index           =   2
            Left            =   270
            TabIndex        =   18
            Top             =   915
            Width           =   1185
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Application Date"
            Height          =   195
            Index           =   1
            Left            =   270
            TabIndex        =   16
            Top             =   255
            Width           =   1170
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Specify"
            Height          =   195
            Index           =   4
            Left            =   270
            TabIndex        =   22
            Top             =   1485
            Width           =   525
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'Transparent
            FillStyle       =   0  'Solid
            Height          =   285
            Index           =   0
            Left            =   6960
            Tag             =   "et0;et0"
            Top             =   300
            Width           =   2820
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   " Applicant Info "
         Enabled         =   0   'False
         Height          =   1575
         Left            =   75
         TabIndex        =   53
         Tag             =   "wt0;fb0"
         Top             =   510
         Width           =   10095
         Begin VB.TextBox txtWaysMn 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Index           =   82
            Left            =   5820
            MaxLength       =   50
            TabIndex        =   15
            Top             =   1155
            Width           =   2865
         End
         Begin VB.ComboBox cmbField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   2
            ItemData        =   "frmMPCreditApproval.frx":0008
            Left            =   1125
            List            =   "frmMPCreditApproval.frx":000A
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   1155
            Width           =   1710
         End
         Begin VB.TextBox txtOthers 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   7725
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   82
            Left            =   5835
            Locked          =   -1  'True
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   540
            Width           =   2865
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   81
            Left            =   5835
            MaxLength       =   50
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   240
            Width           =   1380
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   585
            Index           =   83
            Left            =   1125
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   540
            Width           =   3660
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   80
            Left            =   1125
            TabIndex        =   3
            Top             =   240
            Width           =   3660
         End
         Begin xrControl.xrFrame xrFrame2 
            Height          =   1215
            Left            =   8760
            Top             =   225
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   2143
            BackColor       =   12632256
            Begin VB.Image imgField 
               Height          =   1095
               Left            =   30
               Picture         =   "frmMPCreditApproval.frx":000C
               Stretch         =   -1  'True
               Top             =   45
               Width           =   1095
            End
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Occupation"
            Height          =   195
            Index           =   15
            Left            =   4875
            TabIndex        =   14
            Top             =   1245
            Width           =   825
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Civil Status"
            Height          =   195
            Index           =   24
            Left            =   120
            TabIndex        =   6
            Top             =   1245
            Width           =   780
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Age"
            Height          =   195
            Index           =   81
            Left            =   7320
            TabIndex        =   10
            Top             =   285
            Width           =   285
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Birth Place"
            Height          =   195
            Index           =   18
            Left            =   4860
            TabIndex        =   12
            Top             =   585
            Width           =   765
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Birth Date"
            Height          =   195
            Index           =   17
            Left            =   4860
            TabIndex        =   8
            Top             =   285
            Width           =   705
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            Height          =   195
            Index           =   14
            Left            =   120
            TabIndex        =   4
            Top             =   585
            Width           =   570
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            Height          =   195
            Index           =   13
            Left            =   120
            TabIndex        =   2
            Top             =   285
            Width           =   420
         End
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   1
         Top             =   120
         Width           =   1710
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   285
         Left            =   1365
         Tag             =   "et0;ht2"
         Top             =   195
         Width           =   1710
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction No."
         Height          =   195
         Index           =   0
         Left            =   105
         TabIndex        =   0
         Top             =   165
         Width           =   1140
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   90
      TabIndex        =   57
      Top             =   510
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
      Picture         =   "frmMPCreditApproval.frx":4CA0
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   90
      TabIndex        =   58
      Top             =   1770
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&DisApprv"
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
      Picture         =   "frmMPCreditApproval.frx":541A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   90
      TabIndex        =   59
      Top             =   1140
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Approve"
      AccessKey       =   "A"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmMPCreditApproval.frx":5B94
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   90
      TabIndex        =   60
      Top             =   2400
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
      Picture         =   "frmMPCreditApproval.frx":630E
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   450
      Index           =   1
      Left            =   1545
      Tag             =   "wt0;fb0"
      Top             =   510
      Width           =   10305
      _ExtentX        =   18177
      _ExtentY        =   794
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   4725
         TabIndex        =   63
         Top             =   60
         Width           =   5460
      End
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   1245
         TabIndex        =   62
         Top             =   60
         Width           =   1710
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction No."
         Height          =   285
         Index           =   19
         Left            =   60
         TabIndex        =   65
         Top             =   105
         Width           =   1155
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Full Name"
         Height          =   285
         Index           =   9
         Left            =   3855
         TabIndex        =   64
         Top             =   105
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmMPCreditApproval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmCreditAppRegNeo"

Private oSkin As clsFormSkin
Private WithEvents oTrans As ggcLRApplication.clsLRApplicationMP
Attribute oTrans.VB_VarHelpID = -1

Dim psObjectNme As String

Dim pnCtr As Integer
Dim pnIndex As Integer
Dim panCmdHwnd(3) As Long

Dim pnRow As Integer

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String

   lsOldProc = "Form_Load"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   '''On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oTrans = New ggcLRApplication.clsLRApplicationMP
   Set oTrans.AppDriver = oApp
   oTrans.UnitApplied = 3 'Financing MP
   oTrans.InitTransaction
   oTrans.LoadMode = 1
   
   oTrans.Filter = ""
   If oApp.ProductID <> "LRTrackr" Then oTrans.Filter = "a.sTransNox LIKE " & strParm(oApp.BranchCode & "%")
   oTrans.Filter = oTrans.Filter & IIf(Trim(oTrans.Filter) = "", "", " AND ") & " (a.cTranStat IN ('0','1') OR (a.cTranStat IN ('3', '2') AND sApproved = ''))"
      
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualLeft

   InitEntry
   InitValue
   initButton xeModeReady

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub InitEntry()
   Dim loTxt As TextBox

   txtField(0).Enabled = False

   'set maximum lenth to txtField
   For Each loTxt In txtField
      Select Case loTxt.Index
      Case 15, 25, 16, 17, 7, 20, 21
         loTxt.MaxLength = oTrans.MasterMasFldSize(loTxt.Index)
      Case Else
         loTxt.MaxLength = 0
      End Select
   Next

   'set maximum lenth to txtWaysMn
   For Each loTxt In txtWaysMn
      Select Case loTxt.Index
      Case 0, 1, 3, 6, 10, 11, 13, 25, 26, 28, 31, 35, 36, 38, 44
         loTxt.MaxLength = oTrans.WayMeansMasFldSize(loTxt.Index)
      Case Else
         loTxt.MaxLength = 0
      End Select
   Next

   cmbField(0).List(0) = "New Customer"
   cmbField(0).List(1) = "Repeat Customer"
   
   cmbField(1).List(0) = "Financing MP"
'   cmbField(1).List(0) = "Motorcycle"
'   cmbField(1).List(1) = "Sidecar"
'   cmbField(1).List(2) = "Others"

   'Applicant status
   cmbField(2).List(0) = "Single"
   cmbField(2).List(1) = "Married"
   cmbField(2).List(2) = "Separated"
   cmbField(2).List(3) = "Widowed"
   cmbField(2).List(4) = "Single-Parent"
   cmbField(2).List(5) = "Single-Parent w/ live in partner"
   
   panCmdHwnd(0) = cmdButton(1).hwnd   ' Search
   panCmdHwnd(1) = cmdButton(2).hwnd   ' Delete
   panCmdHwnd(2) = cmdButton(3).hwnd   ' Cancel
   panCmdHwnd(3) = cmdButton(0).hwnd   ' Save
End Sub

Private Sub oTrans_LoadData()
   LoadMaster
End Sub

Private Sub oTrans_MasterRetrieved(ByVal Index As Variant)
   Select Case Index
   'Applicant Info
   Case 80
      txtField(80) = oTrans.Master("sClientNm") 'Name
      txtField(83) = oTrans.Master("sAddress1") 'Address
      txtField(81) = Format(oTrans.Master("dBrthDte1"), "Mmm. DD, YYYY") 'Birth Date
      txtField(82) = oTrans.Master("sBrthPlc1") 'Birth Place
      If IsDate(oTrans.Master("dBrthDte1")) Then
         txtOthers(0) = Format(DateDiff("M", oTrans.Master("dBrthDte1"), oTrans.Master("dAppliedx")) / 12, "0.00") & " yrs"
      Else
         txtOthers(0) = "yrs"
      End If
      cmbField(2).ListIndex = oTrans.Master("cCvlStat1")
      txtField(86) = oTrans.Master("sPhoneNo1")
   'Spouse Info
   Case 87
      txtField(87) = oTrans.Master("sSpouseNm") 'Name
      txtField(90) = oTrans.Master("sAddress2") 'Address
      txtField(88) = Format(oTrans.Master("dBrthDte2"), "Mmm. DD, YYYY") 'Birth Date
      txtField(89) = oTrans.Master("sBrthPlc2") 'Birth Place
      If IsDate(oTrans.Master("dBrthDte2")) Then txtOthers(1) = Format(DateDiff("M", oTrans.Master("dBrthDte2"), oTrans.Master("dAppliedx")) / 12, "0.00") & " yrs"
      txtField(91) = oTrans.Master("sPhoneNo2")
   'Comaker Info
   Case 92
      txtField(92) = oTrans.Master("sComakrNm") 'Name
      txtField(94) = oTrans.Master("sAddress3") 'Address
      txtField(96) = oTrans.Master("sPhoneNo3")
   Case 84, 98
   Case 8, 9, 10, 12, 13
      txtField(Index).Text = Format(oTrans.Master(Index), "#,##0.00")
   Case Else
      txtField(Index).Text = oTrans.Master(Index)
   End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
   Set oTrans = Nothing
End Sub

Private Sub cmdButton_Click(Index As Integer)
   Dim lnRep As Integer
   Dim lsOldProc As String
   
   lsOldProc = "cmdButton_Click"
   '''On Error GoTo errProc
   
   Select Case Index
   Case 0
      If oTrans.SearchTransaction Then
         LoadMaster
      End If
   Case 1
      If txtField(0).Text <> "" Then
         lnRep = MsgBox("Approved Customer Credit Application!!!", vbQuestion + vbYesNo, "Confirm")
   
         If lnRep = vbYes Then
            If oTrans.ApproveTransaction Then
               MsgBox "Transaction Saved Successfully!!!", vbInformation, "Notice"
               Label2.Caption = Format(ApplStat(oTrans.Master("cTranStat")), ">")
            Else
               MsgBox "Unable to Approved Customer Credit Application!!!", vbCritical, "Warning"
            End If
         End If
      Else
         MsgBox "Unable to Approved Customer Credit Application!!!" & vbCrLf & _
                "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
      End If
   Case 2
      If txtField(0).Text <> "" Then
         lnRep = MsgBox("Disapproved Customer Credit Application!!!", vbQuestion + vbYesNo, "Confirm")
   
         If lnRep = vbYes Then
            If oTrans.DisApproveTransaction Then
               MsgBox "Transaction Saved Successfully!!!", vbInformation, "Notice"
               Label2.Caption = Format(ApplStat(oTrans.Master("cTranStat")), ">")
            Else
               MsgBox "Unable to Disapproved Customer Credit Application!!!", vbCritical, "Warning"
            End If
         End If
      Else
         MsgBox "Unable to Disapproved Customer Credit Application!!!" & vbCrLf & _
                "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
      End If
   Case 3
      Unload Me
   End Select

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Sub initButton(lnStat As Integer)
   Dim lbShow As Boolean

   lbShow = IIf(lnStat = 0, False, True)
   
   xrFrame1(0).Enabled = lbShow

   'Disable entry for name, spouse, comaker, financer, and dependent if already entered
   'to update use the command button
   'Kalyptus - 2009.05.01
   If lnStat = xeModeUpdate Then
      'Applicant
      If txtField(80).Text <> Empty Then
         txtField(80).Locked = True
      Else
         txtField(80).Locked = False
      End If

      'Spouse
      If txtField(87).Text <> Empty Then
         txtField(87).Locked = True
      Else
         txtField(87).Locked = False
      End If

      'Comaker
      If txtField(92).Text <> Empty Then
         txtField(92).Locked = True
      Else
         txtField(92).Locked = False
      End If

      'Financer
      If txtWaysMn(86).Text <> Empty Then
         txtWaysMn(86).Locked = True
      Else
         txtWaysMn(86).Locked = False
      End If
   Else
   End If
End Sub

Private Sub oTrans_WaysMeans(ByVal Index As Variant)
   Select Case Index
   Case 0      'Just a dummy
   Case 46, 47, 16, 17
      txtWaysMn(Index).Text = Format(oTrans.WaysMeans(Index), "#,##0.00")
   Case 86
      txtWaysMn(86).Text = oTrans.WaysMeans("sFinancer")  'Name
      txtWaysMn(87).Text = oTrans.WaysMeans("sAddress2")  'Address
      txtWaysMn(91).Text = oTrans.WaysMeans("sPhoneNo2")
   Case Else
      txtWaysMn(Index).Text = oTrans.WaysMeans(Index)
   End Select
End Sub

Private Sub InitValue()
   Dim loTxt As TextBox
   Dim lnCtr As Integer

   'Load Value from the master table
   For Each loTxt In txtField
      Select Case loTxt.Index
      Case 8 To 10, 12
         loTxt.Text = ".00"
      Case Else
         loTxt.Text = ""
      End Select
   Next

   'Load Value from the ways and means
   For Each loTxt In txtWaysMn
      Select Case loTxt.Index
      Case 7, 8, 32, 33, 40, 15 To 24
         loTxt.Text = "0.00"
      Case Else
         loTxt.Text = oTrans.Master(loTxt.Index)
      End Select
   Next

   txtOthers(0).Text = "yrs"
   txtOthers(1).Text = "yrs"
'   txtOthers(2).Text = "0.00"
'   txtOthers(3).Text = "0.00"
'   txtOthers(4).Text = "0.00"

   cmbField(0).ListIndex = 0
   cmbField(1).ListIndex = 0
   cmbField(2).ListIndex = 0
'   cmbField(3).ListIndex = 0
'   cmbField(4).ListIndex = 0

   Label2.Caption = "UNKNOWN"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyReturn, vbKeyUp, vbKeyDown
      If psObjectNme = "cmbField" And KeyCode <> vbKeyReturn Then
         Exit Sub
      End If

      Select Case KeyCode
      Case vbKeyReturn, vbKeyDown
'         If GetFocus = GridEditor1.hwnd Or GetFocus = GridEditor2.hwnd Then Exit Sub
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

Private Sub ComInEx()
   Dim lnTotExpense As Double
   Dim lnCtr As Integer

   lnTotExpense = 0#
   For lnCtr = 18 To 24
      lnTotExpense = lnTotExpense + IFNull(oTrans.WaysMeans(lnCtr), 0)
   Next
'   txtOthers(3).Text = Format(lnTotExpense, "#,##0.00")
'   txtOthers(4).Text = Format(IFNull(oTrans.WaysMeans("nMonGross"), 0) + IFNull(oTrans.WaysMeans("nMonOther"), 0) - lnTotExpense, "#,##0.00")
End Sub

Private Sub txtOthers_GotFocus(Index As Integer)
   psObjectNme = "txtOthers"
End Sub

Private Sub LoadMaster()
   Dim loTxt As TextBox
   Dim lnCtr As Integer
   Dim lsOldProc As String

   lsOldProc = "LoadMaster()"
   Debug.Print pxeMODULENAME & "." & lsOldProc

   'Load Value from the master table
   For Each loTxt In txtField
      Select Case loTxt.Index
      Case 0
         loTxt = Format(oTrans.Master(loTxt.Index), IIf(Len(oApp.BranchCode) = 2, "@@@@-@@@@@@", "@@@@@@-@@@@@@"))
         txtSearch(0) = oTrans.Master(0)
         txtSearch(1) = oTrans.Master(80)
         txtSearch(0).Tag = oTrans.Master(0)
         txtSearch(1).Tag = oTrans.Master(80)
      Case 8 To 10, 12
         loTxt.Text = Format(oTrans.Master(loTxt.Index), "#,##0.00")
      Case 11
         loTxt.Text = CInt(oTrans.Master(loTxt.Index))
      Case 3, 81, 88
         loTxt.Text = Format(oTrans.Master(loTxt.Index), "Mmm DD, YYYY")
      Case Else
         loTxt.Text = IFNull(oTrans.Master(loTxt.Index))
      End Select
   Next

   'Load Value from the ways and means
   For Each loTxt In txtWaysMn
      Select Case loTxt.Index
      Case 7, 8, 32, 33, 40, 15 To 24, 46, 47
         loTxt.Text = Format(IFNull(oTrans.WaysMeans(loTxt.Index), 0), "#,##0.00")
      Case Else
         loTxt.Text = IFNull(oTrans.WaysMeans(loTxt.Index))
      End Select
   Next

   'Load Age of Applicant based on the transaction date
   If IsDate(oTrans.Master("dBrthDte1")) Then
      txtOthers(0) = Format(DateDiff("M", oTrans.Master("dBrthDte1"), oTrans.Master("dAppliedx")) / 12, "0.00") & " yrs"
   Else
      txtOthers(0) = "yrs"
   End If

   'Load age of spouse based on transaction
   If IsDate(oTrans.Master("dBrthDte2")) Then
      txtOthers(1) = Format(DateDiff("M", oTrans.Master("dBrthDte2"), oTrans.Master("dAppliedx")) / 12, "0.00") & " yrs"
   Else
      txtOthers(1) = "yrs"
   End If

   'Load Income/Expenses
'   txtOthers(2).Text = Format(IFNull(oTrans.WaysMeans("nMonGross"), 0) + IFNull(oTrans.WaysMeans("nMonOther"), 0), "#,##0.00")
   ComInEx

   cmbField(0).ListIndex = IFNull(oTrans.Master("cApplType"), 0)
   cmbField(1).ListIndex = IFNull(oTrans.Master("cUnitAppl") - 3, 0)
   cmbField(2).ListIndex = IIf(IsNull(oTrans.Master("cCvlStat1")), -1, oTrans.Master("cCvlStat1"))

   Label2.Caption = Format(ApplStat(oTrans.Master("cTranStat")), ">")
End Sub

Private Sub txtSearch_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String

   lsOldProc = "txtSearch_KeyDown"
   '''On Error GoTo errProc

   If KeyCode = vbKeyReturn Then
      If txtSearch(Index).Text <> txtSearch(Index).Tag Then
         If oTrans.SearchTransaction(IIf(Index = 0, CodeFormat(oApp.BranchCode, txtSearch(Index).Text) _
            , txtSearch(Index).Text) _
            , IIf(Index = 0, True, False)) Then
            LoadMaster
'            LoadDetail
         End If
      End If
   End If

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " _
                       & "  " & Index _
                       & ", " & KeyCode _
                       & ", " & Shift _
                       & " )", True
End Sub

Private Sub txtSearch_Validate(Index As Integer, Cancel As Boolean)
   Dim lsOldProc As String

   lsOldProc = "txtSearch_Validate"
   '''On Error GoTo errProc

   If txtSearch(Index).Text <> "" Then
      If txtSearch(Index).Text <> txtSearch(Index).Tag Then
         If oTrans.SearchTransaction(IIf(Index = 0, CodeFormat(oApp.BranchCode, txtSearch(Index).Text) _
            , txtSearch(Index).Text) _
            , IIf(Index = 0, True, False)) Then
            LoadMaster
'            LoadDetail
         End If
      End If
   End If
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " _
                       & "  " & Index _
                       & ", " & Cancel _
                       & " )", True
End Sub

Private Function isEntryOk() As Boolean
   Dim lbSame As Boolean
   Dim lnCtr As Integer

   'Test valid entries
   'Test validity of spouse/comaker
   lbSame = False
   If txtField(80).Text = txtField(87).Text Or _
      txtField(80).Text = txtField(92).Text Or _
      (txtField(87).Text = txtField(92).Text And txtField(87) <> "") Then
      lbSame = True
   End If

   If lbSame Then
      MsgBox "Applicant/Spouse/Co-Maker must be different person!!!" & vbCrLf & _
             "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
      GoTo EntryNotOK
   End If

   'Transfer contents of combo boxes
   oTrans.Master(1) = cmbField(0).ListIndex
   oTrans.Master(5) = cmbField(1).ListIndex
'   oTrans.Master(42) = cmbField(2).ListIndex
   oTrans.WaysMeans(9) = IIf(cmbField(3).ListIndex = -1, "", cmbField(3).ListIndex)
   oTrans.WaysMeans(34) = IIf(cmbField(4).ListIndex = -1, "", cmbField(4).ListIndex)

EntryOK:
   isEntryOk = True
   Exit Function
EntryNotOK:
   isEntryOk = False
End Function
