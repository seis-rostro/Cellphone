VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmMPCreditTempAppNeo 
   BorderStyle     =   0  'None
   Caption         =   "Credit Application"
   ClientHeight    =   7050
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11940
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7050
   ScaleWidth      =   11940
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   6450
      Index           =   0
      Left            =   1545
      Tag             =   "wt0;fb0"
      Top             =   510
      Width           =   10320
      _ExtentX        =   18203
      _ExtentY        =   11377
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.Frame Frame3 
         Caption         =   "Spouse Info"
         Height          =   1335
         Left            =   90
         TabIndex        =   61
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
         Height          =   2820
         Left            =   75
         TabIndex        =   60
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
            ItemData        =   "frmCreditTempAppNeo.frx":0000
            Left            =   1560
            List            =   "frmCreditTempAppNeo.frx":0002
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
            Left            =   1575
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
            TabIndex        =   62
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
      End
      Begin VB.Frame Frame1 
         Caption         =   " Applicant Info "
         Height          =   1575
         Left            =   75
         TabIndex        =   59
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
            ItemData        =   "frmCreditTempAppNeo.frx":0004
            Left            =   1125
            List            =   "frmCreditTempAppNeo.frx":0006
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
               Picture         =   "frmCreditTempAppNeo.frx":0008
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
      TabIndex        =   53
      Top             =   510
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Save"
      AccessKey       =   "S"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCreditTempAppNeo.frx":4C9C
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   90
      TabIndex        =   55
      Top             =   510
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
      Picture         =   "frmCreditTempAppNeo.frx":5416
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   90
      TabIndex        =   54
      Top             =   1140
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "Searc&h"
      AccessKey       =   "h"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCreditTempAppNeo.frx":5B90
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   90
      TabIndex        =   56
      Top             =   2400
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
      Picture         =   "frmCreditTempAppNeo.frx":630A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   105
      TabIndex        =   58
      Top             =   1140
      Width           =   1230
      _ExtentX        =   2170
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
      Picture         =   "frmCreditTempAppNeo.frx":6A84
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   90
      TabIndex        =   57
      Top             =   1770
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
      Picture         =   "frmCreditTempAppNeo.frx":71FE
   End
End
Attribute VB_Name = "frmMPCreditTempAppNeo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmCreditTempAppNeo"

Private oSkin As clsFormSkin
Private WithEvents oTrans As ggcLRApplication.clsLRApplication
Attribute oTrans.VB_VarHelpID = -1

Dim pnCtr As Integer
Dim pnIndex As Integer

Dim psObjectNme As String

Dim panCmdHwnd(3) As Long

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String

   lsOldProc = "Form_Load"
   ''On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oTrans = New ggcLRApplication.clsLRApplication
   Set oTrans.AppDriver = oApp
   
   oTrans.UnitApplied = 3 'Financing MP
   oTrans.InitTransaction

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualLeft
   
   oTrans.NewTransaction
   InitEntry
   InitValue
   initButton xeModeAddNew

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub cmbField_GotFocus(Index As Integer)
   psObjectNme = "cmbField"
End Sub

Private Sub cmbField_LostFocus(Index As Integer)
   If Not Index = 2 Then
      If cmbField(Index).ListIndex < 0 Then cmbField(Index).ListIndex = -1
   End If
End Sub

Private Sub InitEntry()
   txtField(0).Enabled = False


   cmbField(0).List(0) = "New Customer"
   cmbField(0).List(1) = "Repeat Customer"

   cmbField(1).List(0) = "Financing MP"
'   cmbField(1).List(0) = "Motorcycle"
'   cmbField(1).List(1) = "Sidecar"
'   cmbField(1).List(2) = "Others"

   cmbField(2).List(0) = "Single"
   cmbField(2).List(1) = "Married"
   cmbField(2).List(2) = "Separated"
   cmbField(2).List(3) = "Widowed"

   panCmdHwnd(0) = cmdButton(1).hwnd   ' Search
   panCmdHwnd(1) = cmdButton(2).hwnd   ' Delete
   panCmdHwnd(2) = cmdButton(3).hwnd   ' Cancel
   panCmdHwnd(3) = cmdButton(0).hwnd   ' Save
End Sub

'TODO: Transfer the capturing of image in the ggcClient.clsNeoClient
Private Sub imgField_Click()
   Dim lsSQL As String
   Dim lnRet As Integer
   Dim lsCamPath As String
   Dim lsTmpPath As String
   Dim loFrm As frmWebCam
   
   lsCamPath = oApp.AppPath & "/Temp/WEBCAM.BMP"
   lsTmpPath = oApp.AppPath & "/Temp/client/" & oTrans.Master("sTransNox") & Format(oTrans.Master("dImageDte"), "???") & ".BMP"  '????
      
   If Not FileExists(oApp.AppPath & "/Temp/client/") Then
      MkDir oApp.AppPath & "/Temp/client/"
   End If
   
   'Check if file exist
   If (FileExists(lsCamPath)) Then
      Kill lsCamPath
   End If
      
   Set loFrm = New frmWebCam
   loFrm.Show vbModal
      
   If (FileExists(lsTmpPath)) Then
      Kill lsTmpPath
   End If
      
   'Check if file exist
   If (FileExists(lsCamPath)) Then
      'Move the image file
      Name lsCamPath As lsTmpPath
      imgField.Picture = LoadPicture(lsTmpPath)
   Else
      imgField.Picture = Nothing
   End If
End Sub

Function FileExists(ByVal sFileName As String) As Boolean
   Dim intReturn As Integer
   On Error GoTo FileExists_Error
   intReturn = GetAttr(sFileName)
   FileExists = True
Exit Function
FileExists_Error:
    FileExists = False
End Function

Private Sub oTrans_MasterRetrieved(ByVal Index As Variant)
   Select Case Index
   Case 80 'Applicant
      txtField(80) = oTrans.Master("sClientNm") 'Name
      txtField(83) = oTrans.Master("sAddress1") 'Address
      txtField(81) = Format(oTrans.Master("dBrthDte1"), "Mmm. DD, YYYY") 'Birth Date
      txtField(82) = oTrans.Master("sBrthPlc1") 'Birth Place
      If IsDate(oTrans.Master("dBrthDte1")) Then txtOthers(0) = Format(DateDiff("M", oTrans.Master("dBrthDte1"), oTrans.Master("dAppliedx")) / 12, "0.00") & " yrs"
      cmbField(2).ListIndex = IIf(IsNumeric(oTrans.Master("cCvlStat1")), oTrans.Master("cCvlStat1"), -1)
      txtWaysMn(82) = oTrans.WaysMeans(82)
      Debug.Print oTrans.Master("cCvlStat1")
   Case 87
      txtField(87) = oTrans.Master("sSpouseNm") 'Name
      txtField(90) = oTrans.Master("sAddress2") 'Address
      txtField(88) = Format(oTrans.Master("dBrthDte2"), "Mmm. DD, YYYY") 'Birth Date
      txtField(89) = oTrans.Master("sBrthPlc2") 'Birth Place
      If IsDate(oTrans.Master("dBrthDte2")) Then txtOthers(1) = Format(DateDiff("M", oTrans.Master("dBrthDte2"), oTrans.Master("dAppliedx")) / 12, "0.00") & " yrs"
      txtWaysMn(85) = oTrans.WaysMeans(85)
   Case 8, 9, 10, 12
      txtField(Index).Text = Format(oTrans.Master(Index), "#,##0.00")
   Case 84
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
   ''On Error GoTo errProc

   Select Case Index
   Case 0
      If isEntryOk Then
         If oTrans.SaveTransaction = True Then
            MsgBox "Transaction Saved Successfully!!!", vbInformation, "Notice"
            initButton xeModeReady
         Else
            MsgBox "Unable to Save Transaction!!!", vbCritical, "Warning"
         End If
      End If
   Case 1
      If psObjectNme = "txtField" Then
         Select Case pnIndex
         Case 80, 87  'applicant,spouse,model,credit investigator
            If txtField(Index).Text <> "" Then oTrans.SearchMaster Index, txtField(Index).Text
         Case 97, 99
            oTrans.SearchMaster Index, txtField(Index).Text
         End Select
         txtField(pnIndex).SetFocus
      ElseIf psObjectNme = "txtWaysMn" Then
         Select Case pnIndex
         Case 82, 85
            oTrans.SearchOnWays pnIndex, txtWaysMn(pnIndex)
            txtWaysMn(pnIndex).SetFocus
         End Select
      End If
   Case 3
      lnRep = MsgBox("Transaction is in Update Mode!!!" & vbCrLf & _
                     "Do you want to Cancel Transaction!!!", vbYesNo + vbQuestion, "Confirm")

      If lnRep = vbYes Then
         oTrans.NewTransaction
         initButton xeModeReady
         InitValue
      Else
         txtField(pnIndex).SetFocus
      End If
   Case 4
      oTrans.NewTransaction
      initButton xeModeAddNew
      InitValue
   Case 5
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
   cmdButton(0).Visible = lbShow
   cmdButton(1).Visible = lbShow
   cmdButton(2).Visible = lbShow
   cmdButton(3).Visible = lbShow

   cmdButton(4).Visible = Not lbShow
   cmdButton(5).Visible = Not lbShow

   xrFrame1(0).Enabled = lbShow

   For pnCtr = 1 To xrFrame1.Count - 1
      xrFrame1(pnCtr).Enabled = lbShow
   Next

   If Not lbShow Then cmdButton(4).SetFocus
End Sub

Private Sub oTrans_WaysMeans(ByVal Index As Variant)
   Select Case Index
   Case 82, 85
      txtWaysMn(Index).Text = oTrans.WaysMeans(Index)
   End Select
End Sub

Private Sub InitValue()
   Dim lnCtr As Integer
   Dim lotxt As TextBox

   'Load data from clsLRApplication
   For Each lotxt In txtField
      Select Case lotxt.Index
      Case 0
         lotxt.Text = Format(oTrans.Master(lotxt.Index), IIf(Len(oApp.BranchCode) = 2, "@@@@-@@@@@@", "@@@@@@-@@@@@@"))
      Case 8 To 10, 12
         lotxt.Text = Format(oTrans.Master(lotxt.Index), "#,##0.00")
      Case 11
         lotxt.Text = CInt(oTrans.Master(lotxt.Index))
      Case 81, 88
         lotxt.Text = Format(oTrans.Master(lotxt.Index), "Mmmm DD, YYYY")
      Case Else
         lotxt.Text = oTrans.Master(lotxt.Index)
      End Select
   Next

   'Load data from clsLRWaysMeans
   For Each lotxt In txtWaysMn
      lotxt.Text = oTrans.Master(lotxt.Index)
   Next

   txtOthers(0).Text = "yrs."
   txtOthers(1).Text = "yrs."
   oTrans.Master("cUnitAppl") = 3
   
   cmbField(0).ListIndex = oTrans.Master("cApplType")
   cmbField(1).ListIndex = oTrans.Master("cUnitAppl") - 3
   cmbField(2).ListIndex = oTrans.Master("cCvlStat1")
   
   'enable QM temporarily for the LR Tracker
   If oApp.ProductID = "LRTrackr" Then
      txtField(14).Locked = False
      txtField(14).TabStop = True
   Else
      txtField(14).Locked = True
      txtField(14).TabStop = False
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyReturn, vbKeyUp, vbKeyDown
      If psObjectNme = "cmbField" And KeyCode <> vbKeyReturn Then
         Exit Sub
      End If

      Select Case KeyCode
      Case vbKeyReturn, vbKeyDown
         SetNextFocus
      Case vbKeyUp
         SetPreviousFocus
      End Select
   End Select
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   If txtField(Index) <> Empty Then
      txtField(Index).SelStart = 0
      txtField(Index).SelLength = Len(txtField(Index).Text)
   End If
   psObjectNme = "txtField"
   pnIndex = Index
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String

   lsOldProc = "txtField_KeyDown"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   ''On Error GoTo errProc

   If KeyCode = vbKeyF3 Or KeyCode = vbKeyTab Or KeyCode = vbKeyReturn Then
      Select Case Index
      Case 80, 87  'applicant,spouse,model,credit investigator
         If txtField(Index).Text <> "" Then oTrans.SearchMaster Index, txtField(Index).Text
      Case 97, 99
         oTrans.SearchMaster Index, txtField(Index).Text
      End Select
      KeyCode = 0
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

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   Dim lsOldProc As String

   lsOldProc = "txtField_Validate"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   ''On Error GoTo errProc

   If Index < 80 And txtField(Index).Locked = False Then
      With txtField(Index)
         Select Case Index
         Case 3   'Application date
            If Not IsDate(.Text) Then .Text = oApp.ServerDate
            .Text = Format(txtField(Index).Text, "MMMM DD, YYYY")
         Case 10  'Case Down payment
            If Not IsNumeric(.Text) Then .Text = "0.00"
            .Text = Format(txtField(Index).Text, "#,##0.00")
         Case 13  'Case term
            If Not IsNumeric(.Text) Then .Text = "0"
         End Select
      End With

      'save the data to the otrans object
      oTrans.Master(Index) = txtField(Index).Text
   End If

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " _
                       & "  " & Index _
                       & ", " & Cancel _
                       & " )", True

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

Private Sub txtOthers_GotFocus(Index As Integer)
   psObjectNme = "txtOthers"
End Sub

Private Sub txtWaysMn_GotFocus(Index As Integer)
   If txtWaysMn(Index) <> Empty Then
      txtWaysMn(Index).SelStart = 0
      txtWaysMn(Index).SelLength = Len(txtWaysMn(Index).Text)
   End If

   psObjectNme = "txtWaysMn"
   pnIndex = Index
End Sub

Private Sub txtWaysMn_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String

   lsOldProc = "txtField_KeyDown"
   ''On Error GoTo errProc

   If KeyCode = vbKeyF3 Then
      Select Case Index
      Case 82, 85 'Applicant's Occupation, Spouse Occupation
         'If txtWaysMn(Index).Text <> "" Then oTrans.SearchOnWays Index, txtWaysMn(Index).Text
         oTrans.SearchOnWays Index, txtWaysMn(Index).Text
      End Select
      If txtWaysMn(Index).Text <> "" Then SetNextFocus
      KeyCode = 0
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

Private Function isEntryOk() As Boolean
   Dim lbSame As Boolean
   Dim lnCtr As Integer

   oTrans.Master("cApplType") = cmbField(0).ListIndex
   oTrans.Master("cUnitAppl") = cmbField(1).ListIndex + 3
   oTrans.Master("cWithFinx") = chkField.Value

   If oTrans.WaysMeans("sPosition") = "" Then
      MsgBox "Invalid Occupation Detected...", vbCritical, "Entry Error"
   End If

   If cmbField(1).ListIndex = 1 Then
      If oTrans.Master("sModelIDx") = "" Then
         MsgBox "Invalid Model Detected...", vbCritical, "Entry Error"
         isEntryOk = False
      End If
   Else
      isEntryOk = True
   End If
End Function

Private Sub txtWaysMn_Validate(Index As Integer, Cancel As Boolean)
   Dim lsOldProc As String

   lsOldProc = "txtWaysMn_Validate"
   ''On Error GoTo errProc

   Select Case Index
   Case 82, 85 'Applicant's Occupation, Spouse Occupation
      'If txtWaysMn(Index).Text <> "" Then oTrans.SearchOnWays Index, txtWaysMn(Index).Text
      oTrans.SearchOnWays Index, txtWaysMn(Index).Text
   End Select
   If txtWaysMn(Index).Text <> "" Then SetNextFocus

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " _
                       & "  " & Index _
                       & ", " & Cancel _
                       & " )", True
End Sub

Private Sub xrButton1_Click()
   If txtField(9).Locked = True Then
      Call oTrans.AutoCompute(oTrans.Master("sModelIDx"), oTrans.Master("nPNValueX"), oTrans.Master("nDownPaym"), oTrans.Master("nAcctTerm"))
   Else
      oTrans.Master(12) = oTrans.Master("nPNValueX") / IIf(oTrans.Master("nAcctTerm") > 0, oTrans.Master("nAcctTerm"), 1)
   End If
   
   oTrans.Master(8) = oTrans.Master("nPNValueX") + oTrans.Master("nDownPaym")
   
End Sub
