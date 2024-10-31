VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmCashierTransReg 
   BorderStyle     =   0  'None
   Caption         =   "Office Collection"
   ClientHeight    =   8625
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10500
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   10500
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   540
      Index           =   1
      Left            =   120
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   8790
      _ExtentX        =   15505
      _ExtentY        =   953
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
         Index           =   80
         Left            =   1080
         TabIndex        =   1
         Top             =   90
         Width           =   1725
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
         Index           =   81
         Left            =   3915
         TabIndex        =   3
         Top             =   90
         Width           =   4680
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&O.R. No."
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
         Index           =   21
         Left            =   165
         TabIndex        =   0
         Top             =   120
         Width           =   825
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Customer"
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
         Left            =   2970
         TabIndex        =   2
         Top             =   120
         Width           =   795
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   7
      Left            =   9150
      TabIndex        =   66
      Top             =   4770
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
      Picture         =   "frmCashierTransReg.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   9150
      TabIndex        =   60
      Top             =   2250
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
      Picture         =   "frmCashierTransReg.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   9150
      TabIndex        =   61
      Top             =   2880
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Update"
      AccessKey       =   "U"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCashierTransReg.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   9150
      TabIndex        =   62
      Top             =   3510
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
      Picture         =   "frmCashierTransReg.frx":166E
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   9150
      TabIndex        =   64
      Top             =   2880
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
      Picture         =   "frmCashierTransReg.frx":1DE8
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   600
      Index           =   2
      Left            =   9150
      TabIndex        =   67
      Top             =   4770
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
      Picture         =   "frmCashierTransReg.frx":2562
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   600
      Index           =   1
      Left            =   9150
      TabIndex        =   65
      Top             =   3510
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
      Picture         =   "frmCashierTransReg.frx":2CDC
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   7395
      Index           =   0
      Left            =   120
      Tag             =   "wt0;fb0"
      Top             =   1110
      Width           =   8790
      _ExtentX        =   15505
      _ExtentY        =   13044
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   44
         Left            =   1650
         TabIndex        =   76
         Top             =   3750
         Width           =   6870
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   40
         Left            =   1650
         TabIndex        =   57
         Top             =   6600
         Width           =   6855
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   41
         Left            =   1650
         TabIndex        =   59
         Top             =   6915
         Width           =   6855
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   1650
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1545
         Width           =   6855
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   1650
         MaxLength       =   128
         TabIndex        =   19
         Top             =   1860
         Width           =   6855
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   16
         Left            =   6075
         MaxLength       =   30
         TabIndex        =   37
         Tag             =   "ht0"
         Top             =   4605
         Width           =   2430
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   1650
         TabIndex        =   15
         Top             =   1230
         Width           =   6855
      End
      Begin VB.ComboBox cmbField 
         Height          =   315
         ItemData        =   "frmCashierTransReg.frx":3456
         Left            =   6315
         List            =   "frmCashierTransReg.frx":3458
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   495
         Width           =   2310
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   1665
         MaxLength       =   50
         TabIndex        =   7
         Top             =   495
         Width           =   2310
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   17
         Left            =   1665
         TabIndex        =   9
         Top             =   855
         Width           =   2310
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   6315
         MaxLength       =   6
         TabIndex        =   13
         Top             =   855
         Width           =   2295
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   6
         Left            =   1650
         MaxLength       =   50
         TabIndex        =   21
         Top             =   2340
         Width           =   2250
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   14
         Left            =   1665
         MaxLength       =   30
         TabIndex        =   31
         Top             =   4380
         Width           =   1860
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   13
         Left            =   1665
         MaxLength       =   30
         TabIndex        =   29
         Top             =   4080
         Width           =   1860
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   15
         Left            =   1665
         MaxLength       =   30
         TabIndex        =   33
         Top             =   4680
         Width           =   1860
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   9
         Left            =   1665
         MaxLength       =   50
         TabIndex        =   26
         Top             =   2940
         Width           =   2250
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
         Height          =   285
         Index           =   0
         Left            =   1665
         TabIndex        =   5
         Top             =   75
         Width           =   2310
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   30
         Left            =   1680
         MaxLength       =   30
         TabIndex        =   39
         Top             =   5550
         Width           =   1425
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   32
         Left            =   1680
         MaxLength       =   30
         TabIndex        =   43
         Top             =   6150
         Width           =   1425
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   31
         Left            =   1680
         MaxLength       =   30
         TabIndex        =   41
         Top             =   5850
         Width           =   1425
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   33
         Left            =   4380
         MaxLength       =   30
         TabIndex        =   45
         Top             =   5550
         Width           =   1425
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   35
         Left            =   4380
         MaxLength       =   30
         TabIndex        =   49
         Top             =   6150
         Width           =   1425
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   34
         Left            =   4380
         MaxLength       =   30
         TabIndex        =   47
         Top             =   5850
         Width           =   1425
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   36
         Left            =   7080
         MaxLength       =   30
         TabIndex        =   51
         Top             =   5550
         Width           =   1425
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   38
         Left            =   7080
         MaxLength       =   30
         TabIndex        =   55
         Top             =   6150
         Width           =   1425
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   37
         Left            =   7080
         MaxLength       =   30
         TabIndex        =   53
         Top             =   5850
         Width           =   1425
      End
      Begin VB.ComboBox cmbcRegisFrm 
         Height          =   315
         ItemData        =   "frmCashierTransReg.frx":345A
         Left            =   6210
         List            =   "frmCashierTransReg.frx":345C
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   2370
         Width           =   2310
      End
      Begin xrControl.xrFrame xrFrame2 
         Height          =   945
         Index           =   0
         Left            =   1590
         Tag             =   "wt0;fb0"
         Top             =   2595
         Width           =   6915
         _ExtentX        =   12197
         _ExtentY        =   1667
         BorderStyle     =   4
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   10
            Left            =   75
            MaxLength       =   50
            TabIndex        =   73
            Top             =   645
            Width           =   2250
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   12
            Left            =   4590
            MaxLength       =   50
            TabIndex        =   72
            Top             =   645
            Width           =   2325
         End
         Begin VB.OptionButton OptField 
            Caption         =   "Solo"
            Height          =   195
            Index           =   1
            Left            =   60
            TabIndex        =   71
            Tag             =   "et0;fb0"
            Top             =   90
            Width           =   630
         End
         Begin VB.OptionButton OptField 
            Caption         =   "Tricycle"
            Height          =   195
            Index           =   0
            Left            =   885
            TabIndex        =   70
            Tag             =   "et0;fb0"
            Top             =   90
            Width           =   930
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   11
            Left            =   4590
            MaxLength       =   50
            TabIndex        =   69
            Top             =   240
            Width           =   2325
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Color"
            Height          =   285
            Index           =   14
            Left            =   4035
            TabIndex        =   75
            Top             =   660
            Width           =   405
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Model"
            Height          =   285
            Index           =   15
            Left            =   3960
            TabIndex        =   74
            Top             =   300
            Width           =   525
         End
      End
      Begin VB.Label lblAdvPayment 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6075
         TabIndex        =   79
         Top             =   5025
         Width           =   2430
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   30
         Left            =   4830
         TabIndex        =   78
         Top             =   5040
         Width           =   1185
      End
      Begin VB.Label lblOtherC 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&PAID BY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   300
         TabIndex        =   77
         Top             =   3765
         Width           =   900
      End
      Begin VB.Shape Shape2 
         Height          =   720
         Index           =   2
         Left            =   105
         Top             =   6540
         Width           =   8505
      End
      Begin VB.Label lblOtherC 
         BackStyle       =   0  'Transparent
         Caption         =   "Co Owner #1"
         Height          =   285
         Index           =   0
         Left            =   300
         TabIndex        =   56
         Top             =   6645
         Width           =   1200
      End
      Begin VB.Label lblOtherC 
         BackStyle       =   0  'Transparent
         Caption         =   "Co Owner #2"
         Height          =   285
         Index           =   1
         Left            =   300
         TabIndex        =   58
         Top             =   6945
         Width           =   1200
      End
      Begin VB.Label lblcRegisFrm 
         BackStyle       =   0  'Transparent
         Caption         =   "Regis. Form"
         Height          =   285
         Left            =   5160
         TabIndex        =   22
         Top             =   2400
         Width           =   960
      End
      Begin VB.Shape Shape3 
         Height          =   390
         Index           =   3
         Left            =   5280
         Top             =   45
         Width           =   3315
      End
      Begin VB.Shape Shape4 
         Height          =   330
         Index           =   0
         Left            =   5310
         Top             =   75
         Width           =   3255
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
         Height          =   225
         Left            =   5565
         TabIndex        =   68
         Tag             =   "eb0;et0"
         Top             =   150
         Width           =   2910
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name"
         Height          =   285
         Index           =   29
         Left            =   315
         TabIndex        =   14
         Top             =   1275
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Date"
         Height          =   285
         Index           =   28
         Left            =   135
         TabIndex        =   6
         Top             =   540
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cust. Address"
         Height          =   195
         Index           =   11
         Left            =   315
         TabIndex        =   16
         Top             =   1590
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Trans. Type"
         Height          =   285
         Index           =   6
         Left            =   5265
         TabIndex        =   10
         Top             =   555
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A&mt.Tendered"
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
         Index           =   4
         Left            =   4815
         TabIndex        =   36
         Top             =   4725
         Width           =   1200
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
         Left            =   4770
         TabIndex        =   34
         Top             =   4185
         Width           =   720
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Account No."
         Height          =   285
         Index           =   10
         Left            =   150
         TabIndex        =   8
         Top             =   885
         Width           =   945
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   285
         Index           =   12
         Left            =   315
         TabIndex        =   18
         Top             =   1890
         Width           =   690
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   6075
         TabIndex        =   35
         Tag             =   "ht0;ft0"
         Top             =   4065
         Width           =   2430
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "O.R. No."
         Height          =   285
         Index           =   3
         Left            =   5280
         TabIndex        =   12
         Top             =   900
         Width           =   930
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Insurance Type"
         Height          =   285
         Index           =   2
         Left            =   300
         TabIndex        =   20
         Top             =   2340
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Engine No"
         Height          =   285
         Index           =   13
         Left            =   600
         TabIndex        =   25
         Top             =   3000
         Width           =   795
      End
      Begin VB.Shape Shape2 
         Height          =   1020
         Index           =   0
         Left            =   105
         Top             =   1185
         Width           =   8505
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Others"
         Height          =   285
         Index           =   5
         Left            =   315
         TabIndex        =   32
         Top             =   4710
         Width           =   660
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Discoun&t"
         Height          =   285
         Index           =   16
         Left            =   315
         TabIndex        =   30
         Top             =   4440
         Width           =   705
      End
      Begin VB.Shape Shape3 
         Height          =   1410
         Index           =   0
         Left            =   105
         Top             =   2235
         Width           =   8505
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Amount"
         Height          =   285
         Index           =   17
         Left            =   315
         TabIndex        =   28
         Top             =   4155
         Width           =   660
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Registration Type"
         Height          =   285
         Index           =   18
         Left            =   300
         TabIndex        =   24
         Top             =   2670
         Width           =   1245
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Frame No"
         Height          =   285
         Index           =   7
         Left            =   600
         TabIndex        =   27
         Top             =   3315
         Width           =   795
      End
      Begin VB.Shape Shape3 
         Height          =   1755
         Index           =   1
         Left            =   105
         Top             =   3690
         Width           =   8505
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction No."
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
         Index           =   9
         Left            =   135
         TabIndex        =   4
         Top             =   120
         Width           =   1395
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   285
         Left            =   1755
         Tag             =   "et0;ht2"
         Top             =   180
         Width           =   2310
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "PN Value"
         Height          =   285
         Index           =   19
         Left            =   300
         TabIndex        =   38
         Top             =   5550
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Gross Amount"
         Height          =   285
         Index           =   1
         Left            =   300
         TabIndex        =   42
         Top             =   6150
         Width           =   1080
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Mon Amort"
         Height          =   285
         Index           =   22
         Left            =   3360
         TabIndex        =   44
         Top             =   5550
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Balance"
         Height          =   285
         Index           =   23
         Left            =   6090
         TabIndex        =   50
         Top             =   5550
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Term"
         Height          =   285
         Index           =   24
         Left            =   3360
         TabIndex        =   48
         Top             =   6150
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cash Balance"
         Height          =   285
         Index           =   25
         Left            =   3360
         TabIndex        =   46
         Top             =   5850
         Width           =   1005
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount Due"
         Height          =   285
         Index           =   26
         Left            =   6090
         TabIndex        =   52
         Top             =   5850
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Rebate"
         Height          =   285
         Index           =   27
         Left            =   6090
         TabIndex        =   54
         Top             =   6150
         Width           =   915
      End
      Begin VB.Shape Shape3 
         Height          =   1005
         Index           =   2
         Left            =   105
         Top             =   5490
         Width           =   8505
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Down Payment"
         Height          =   285
         Index           =   0
         Left            =   300
         TabIndex        =   40
         Top             =   5850
         Width           =   1095
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H8000000F&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   285
         Index           =   1
         Left            =   5340
         Tag             =   "et0;et0"
         Top             =   105
         Width           =   3195
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   6
      Left            =   9150
      TabIndex        =   63
      Top             =   4140
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "Print"
      AccessKey       =   "Print"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCashierTransReg.frx":345E
   End
End
Attribute VB_Name = "frmCashierTransReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmCashierTransReg"

Private WithEvents oTrans As clsORReceipt
Attribute oTrans.VB_VarHelpID = -1
Private oReceipt As ggcReceipt.Receipt
Private oSkin As clsFormSkin

Dim pnIndex As Integer, pnCtr As Integer
Dim pbMoveCombo As Boolean
Dim psTransNox As String

Property Let TransNo(ByVal Value As String)
   psTransNox = Value
End Property

Private Sub cmbcRegisFrm_Click()
   oTrans.Master("cRegisFrm") = cmbcRegisFrm.ListIndex

   If cmbcRegisFrm.ListIndex = 1 Or cmbcRegisFrm.ListIndex = 2 Then
      txtField(40).Enabled = True
      txtField(41).Enabled = True
   Else
      txtField(40).Enabled = False
      txtField(41).Enabled = False
   End If
End Sub

Private Sub cmbcRegisFrm_Validate(Cancel As Boolean)
   oTrans.Master("cRegisFrm") = cmbcRegisFrm.ListIndex
End Sub

Private Sub cmbField_Click()
   Dim lbEnabled As Boolean

   With cmbField
      Select Case .ListIndex
      Case Is < 0
         .ListIndex = -1
      Case 0
         lbEnabled = False
         lblOtherC(0) = "Co Owner #1"
         lblOtherC(1) = "Co Owner #2"
      Case 1, 2, 3
         lbEnabled = False
         lblOtherC(0) = "Co Owner #1"
         lblOtherC(1) = "Co Owner #2"
      Case 4, 5, 6, 7, 8, 9
         lbEnabled = True
         lblOtherC(0) = "Co Owner #1"
         lblOtherC(1) = "Co Owner #2"
      End Select
      oTrans.Master("cTranType") = .ListIndex
   End With
   
   For pnCtr = 6 To 12
      If Not (pnCtr = 7 Or pnCtr = 8) Then
         txtField(pnCtr).Enabled = lbEnabled
               'control the existence of the Regis Form combo here
      End If
   Next

   If oTrans.Master("cTranType") = 4 _
            Or oTrans.Master("cTranType") = 5 _
            Or oTrans.Master("cTranType") = 6 Then
      cmbcRegisFrm.Visible = True
      lblcRegisFrm.Visible = True
      cmbcRegisFrm.ListIndex = 0
   Else
      cmbcRegisFrm.Visible = False
      lblcRegisFrm.Visible = False
      cmbcRegisFrm.ListIndex = 0
   End If
   
   If Not lbEnabled Then
      OptField(0).Value = False
      OptField(1).Value = False
   End If
End Sub

Private Sub cmbField_GotFocus()
   With cmbField
      .BackColor = oApp.getColor("HT1")
   End With
   pbMoveCombo = True
End Sub

Private Sub cmbField_LostFocus()
   With cmbField
      .BackColor = oApp.getColor("EB")
   End With
   pbMoveCombo = False
End Sub

Private Sub cmbcRegisFrm_GotFocus()
   With cmbcRegisFrm
      .BackColor = oApp.getColor("EB")
   End With
   pbMoveCombo = True
End Sub

Private Sub cmbcRegisFrm_LostFocus()
   With cmbcRegisFrm
      .BackColor = oApp.getColor("EB")
   End With
   pbMoveCombo = False
End Sub

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   Dim lnRep As Integer
   Dim lnStat As Integer
   Dim lbApproved As Boolean
   Dim lnUserRights As Integer
   Dim lsUserID As String, lsUserName As String
   
   lsOldProc = "cmdButton_Click"
   ''On Error GoTo errProc

   If Not pbMoveCombo Then txtField_LostFocus pnIndex
   Select Case Index
   Case 0
      If isEntryOk Then
         If oTrans.Master("nTranTotl") > 0# Then Call Receipt
         If oReceipt.Cancelled Then Exit Sub
         oTrans.Master("sSystemCd") = "GL"

         If oTrans.SaveTransaction = True Then
            MsgBox "Record Updated Successfully!!!", vbInformation, "Notice"
            initButton xeModeReady
         Else
            MsgBox "Unable to Update Record!!!", vbCritical, "Warning"
         End If
      End If
   Case 1
      Select Case pnIndex
      Case 3, 40, 41
         oTrans.Master(pnIndex) = txtField(pnIndex).Text
      Case Else
         oTrans.SearchMaster pnIndex
         txtField(pnIndex).SetFocus
      End Select
      txtField(pnIndex).SetFocus
   Case 2
      lnRep = MsgBox("Transaction is in Update Mode!!!" & vbCrLf & _
                     "Do you want to Cancel Transaction!!!", vbYesNo + vbQuestion, "Confirm")

      If lnRep = vbYes Then
         If oTrans.Master("sClientID") = "" Then
            ClearFields
         Else
            If oTrans.OpenTransaction(oTrans.Master("sTransNox")) Then LoadMaster
         End If
         initButton xeModeReady
      Else
         txtField(pnIndex).SetFocus
      End If
   Case 3
      If oTrans.SearchTransaction() Then
         LoadMaster
      Else
         If txtField(0).Text = "" Then ClearFields
      End If

      txtField(pnIndex).SetFocus
   Case 4
      If Trim(txtField(0).Text) <> "" Then
         If IFNull(oApp.getConfiguration("cRealTime", oTrans.Branch), "0") = "1" Then
            MsgBox "Update of transaction is not allowed for this branch!!!", vbCritical, "Warning"
         Else
               If oTrans.UpdateTransaction Then
                  initButton xeModeUpdate
                  txtField(1).SetFocus
               Else
                  MsgBox "Unable to Update Transaction!!!", vbCritical, "Warning"
               End If
         End If
      Else
         MsgBox "No Transaction is loaded!!!" & vbCrLf & _
                "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
      End If
   Case 5
      If Trim(txtField(0).Text) <> "" Then
               If oTrans.CancelTransaction Then
                  MsgBox "Transaction Cancelled Successfully!!!", vbInformation, "Notice"
                  Label2.Caption = Format(TransStat(oTrans.Master("cTranStat")), ">")
               Else
                  MsgBox "Unable to Cancel Transaction!!!", vbCritical, "Warning"
               End If
      Else
         MsgBox "No Transaction is loaded!!!" & vbCrLf & _
                "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
      End If
   Case 6
      If IFNull(oApp.getConfiguration("cRealTime", oTrans.Branch), "0") = "1" Then
         
         If oTrans.EditMode = xeModeReady Then
            'kalyptus - 2014.09.10 01:39pm
            'Allow print of receipt if entry on GCard precedes printing
            lnStat = oTrans.Master("cTranStat")
            If lnStat >= 4 Then
               lnStat = lnStat Xor 4
            End If
            
            If lnStat <= xeStateClosed Then
               lnRep = MsgBox("Print Transaction!!!", vbYesNo + vbQuestion, "Confirm")
               If lnRep = vbYes Then
                  
                  
                  'kalyptus - 2014.09.10 01:39pm
                  'Allow reprinting print of receipt with managers approval
                  lbApproved = True
                  If lnStat = xeStateClosed Then
                     If oApp.UserLevel < xeManager Then
                        MsgBox "Reprinting of Receipt requires a MANAGER account..." & vbCrLf & _
                               "For inquiries please contact the MIS Department!", vbInformation + vbOKOnly
                        If Not GetApproval(oApp, lnUserRights, lsUserID, lsUserName, oApp.MenuName) Then GoTo endProc
                     Else
                        lnUserRights = oApp.UserLevel
                     End If
                  
                     If lnUserRights < xeManager Then
                        MsgBox "Reprinting of Receipt requires a MANAGER account..." & vbCrLf & _
                               "Can't proceed because of insufficient RIGHT!", vbInformation + vbOKOnly
                        lbApproved = False
                     End If
                  
                  End If
                                    
                  If lbApproved Then
                     MsgBox "Please mount the OR..."
                    'she 11-26-2020
                    If oTrans.Master("cTranType") = 9 Then 'Misc
                        If Not PrintOR Then
                           MsgBox "Unable to Print OR!!!", vbCritical, "Warning"
                        Else
                           If oTrans.CloseTransaction = False Then GoTo endProc
                        End If
                    Else
                        If Not PrintAR Then
                            MsgBox "Unable to Print OR!!!", vbCritical, "Warning"
                        Else
                            If oTrans.CloseTransaction = False Then GoTo endProc
                        End If
                    End If
'                     If Not PrintOR Then
'                        MsgBox "Unable to Print OR!!!", vbCritical, "Warning"
'                     Else
'                        If lnStat = xeStateOpen Then
'                           If oTrans.CloseTransaction = False Then GoTo endProc
'                        End If
'                     End If
                  End If
                  
               End If
            End If
         
         End If
      End If
   Case 7
      Unload Me
   End Select

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String

   lsOldProc = "Form_Load"
   ''On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oTrans = New clsORReceipt
   Set oTrans.AppDriver = oApp
   
   oTrans.Branch = oApp.BranchCode
   oTrans.InitTransaction
'   oTrans.NewTransaction

   Set oReceipt = New ggcReceipt.Receipt
   Set oReceipt.AppDriver = oApp
   oReceipt.InitReceipt
   oReceipt.CheckFieldEnabled = False

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransMaintenance

   InitForm
   ClearFields
   cmbField.ListIndex = 9
   cmbField.Enabled = False
   initButton xeModeReady
   
   If psTransNox <> "" Then
      If oTrans.OpenTransaction(psTransNox) Then LoadMaster
   End If
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   psTransNox = ""

   Set oTrans = Nothing
   Set oSkin = Nothing
   Set oReceipt = Nothing
End Sub

Private Sub oTrans_MasterRetrieved(ByVal Index As Variant)
   Debug.Print oTrans.Master(Index)
   Select Case Index
   Case 7
      If IsNull(oTrans.Master(Index)) Then
         cmbcRegisFrm.ListIndex = -1
         Exit Sub
      End If
      cmbcRegisFrm.ListIndex = oTrans.Master(Index)
   Case 8
      If IsNull(oTrans.Master(Index)) Then
         OptField(0).Value = False
         OptField(1).Value = False
         Exit Sub
      End If
      OptField(oTrans.Master(Index)).Value = True
   Case 16, 30 To 38
      txtField(Index).Text = Format(oTrans.Master(Index), "#,##0.00")
   Case Else
      txtField(Index).Text = IIf(IsNull(oTrans.Master(Index)), "", oTrans.Master(Index))
   End Select
   
   txtField(3).Tag = txtField(3).Text
End Sub

Private Sub ClearFields()
   Dim loTxt As TextBox
   
   For Each loTxt In txtField
      pnCtr = loTxt.Index
      Select Case pnCtr
      Case 1
         txtField(pnCtr).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
      Case 7, 8, 18 To 29
      Case 13, 14, 15, 16, 30 To 34, 36, 37, 38
         txtField(pnCtr).Text = "0.00"
      Case 35
         txtField(pnCtr).Text = 0
      Case Else
         txtField(pnCtr).Text = ""
         txtField(pnCtr).Tag = ""
      End Select
   Next

   lblTotal.Caption = "0.00"
   Label2.Caption = "UNKNOWN"

   cmbField.ListIndex = 9
   oTrans.Master("cTranType") = 1
   
   cmbcRegisFrm.Visible = False
   lblcRegisFrm.Visible = False
   cmbcRegisFrm.ListIndex = 0
   
   txtField(40).Enabled = False
   txtField(41).Enabled = False
   
   xrFrame2(0).Enabled = False
End Sub

Private Sub InitForm()
   cmbField.List(0) = "MC Sales"
   cmbField.List(1) = "Monthly Payment"
   cmbField.List(2) = "Cash Balance"
   cmbField.List(3) = "Down Balance"

   cmbField.List(4) = "Registration"
   cmbField.List(5) = "Insurance"
   cmbField.List(6) = "Change Class"
   cmbField.List(7) = "Deed of Sale"
   cmbField.List(8) = "Release"
   cmbField.List(9) = "Miscellaneous"

   cmbcRegisFrm.List(0) = "New"
   cmbcRegisFrm.List(1) = "Transfer and Renewal"
   cmbcRegisFrm.List(2) = "Transfer"
   cmbcRegisFrm.List(3) = "Renewal"
   cmbcRegisFrm.List(4) = "Others"
   cmbcRegisFrm.List(5) = "Additional"
   cmbcRegisFrm.List(6) = "MAPFRE"
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      If Index = 1 Then .Text = Format(.Text, "MM/DD/YYYY")
      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = oApp.getColor("HT1")
   End With
   pnIndex = Index
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String

   lsOldProc = "txtField_KeyDown"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   ''On Error GoTo errProc

   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
      With txtField(Index)
         Select Case Index
         Case 3, 40, 41
            oTrans.Master(Index) = txtField(Index).Text
         Case 6, 9, 11, 12, 17
            If KeyCode = vbKeyF3 Then
               oTrans.SearchMaster Index, txtField(Index).Text
               If .Text <> "" Then SetNextFocus
            Else
               If .Text <> "" Then oTrans.SearchMaster Index, txtField(Index).Text
            End If
         Case 80, 81
            Call txtField_Validate(Index, False)
         End Select
      End With
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyReturn, vbKeyUp, vbKeyDown
      If pbMoveCombo = True And KeyCode <> vbKeyReturn Then
         Exit Sub
      End If

      Select Case KeyCode
      Case vbKeyReturn, vbKeyDown
         SetNextFocus
      Case vbKeyUp
         SetPreviousFocus
      End Select
   Case vbKeyF8
      If txtField(0).Text <> "" And oTrans.EditMode = xeModeReady Then
         If oTrans.DeleteTransaction Then ClearFields
      End If
   Case vbKeyF12
      oTrans.ViewModify
   End Select
End Sub

Private Sub initButton(lnStat As Integer)
   Dim lbShow As Boolean

   lbShow = IIf(lnStat = 0, False, True)
   cmdButton(3).Visible = Not lbShow
   cmdButton(4).Visible = Not lbShow
   cmdButton(5).Visible = Not lbShow
   cmdButton(7).Visible = Not lbShow
   xrFrame1(1).Enabled = Not lbShow
   
   cmdButton(0).Visible = lbShow
   cmdButton(1).Visible = lbShow
   cmdButton(2).Visible = lbShow
   xrFrame1(0).Enabled = lbShow
'   cmbField.Enabled = lbShow
End Sub

Private Sub Receipt()
   Dim lnCol As Integer
   Dim lnRow As Integer

   With oReceipt
      .AllowEmptyOR = True
      .ORNo = txtField(2).Text
      .ReceiveFrom = txtField(3).Text
      .Address = txtField(4).Text
      .TranDate = txtField(1).Text
      .CashAmount = oTrans.Master("nTranTotl")
'      .CheckAmount = oTrans.Checks("nAmountxx")
'      If oTrans.Checks("nAmountxx") > 0 Then .CashAmount = oTrans.Master("nTranTotl") - oTrans.Checks("nAmountxx")
      .AmountPaid = oTrans.Master("nTranTotl")
      .Remarks = txtField(5).Text
      .SystemCd = "GL"
'      .PRNo = oTrans.PRNumber

'      If oTrans.Checks("nAmountxx") > 0# Then
'         oReceipt.Checks("sCheckNox") = oTrans.Checks("sCheckNox")
'         oReceipt.Checks("sAcctNoxx") = oTrans.Checks("sAcctNoxx")
'         oReceipt.Checks("sBankIDxx") = oTrans.Checks("sBankIDxx")
'         oReceipt.Checks("dCheckDte") = oTrans.Checks("dCheckDte")
'         oReceipt.Checks("nAmountxx") = oTrans.Checks("nAmountxx")
'         oReceipt.Checks("sPRNoxxxx") = oTrans.Checks("sPRNoxxxx")
'         oTrans.PRNumber = .PRNo
'      Else
'         oReceipt.Checks("sCheckNox") = ""
'         oReceipt.Checks("sAcctNoxx") = ""
'         oReceipt.Checks("sBankIDxx") = ""
'         oReceipt.Checks("dCheckDte") = oApp.ServerDate
'         oReceipt.Checks("nAmountxx") = 0#
'         oReceipt.Checks("sPRNoxxxx") = ""
'         .CheckAmount = 0
'      End If
      .ShowReceipt
   End With

   If oTrans.EditMode <> xeModeUpdate Then Exit Sub
   If Not oReceipt.Cancelled Then
      txtField(16).Text = Format(CDbl(oReceipt.AmountPaid), "#,##0.00")

      With oTrans
'         If oReceipt.CheckAmount > 0# Then
'            .Checks("sCheckNox") = oReceipt.Checks("sCheckNox")
'            .Checks("sAcctNoxx") = oReceipt.Checks("sAcctNoxx")
'            .Checks("sBankIDxx") = oReceipt.Checks("sBankIDxx")
'            .Checks("dCheckDte") = oReceipt.Checks("dCheckDte")
'            .Checks("nAmountxx") = oReceipt.Checks("nAmountxx")
'         End If

         .Master("sORNoxxxx") = oReceipt.ORNo
         .Master("sRemarksx") = oReceipt.Remarks
'         .Checks("nAmountxx") = oReceipt.CheckAmount
'         .PRNumber = oReceipt.PRNo
      End With

      txtField(2).Text = oReceipt.ORNo
      txtField(5).Text = oReceipt.Remarks
   End If
End Sub

Private Function isEntryOk() As Boolean
   If txtField(3).Text = "" Then
      MsgBox "Customer not found!!!", vbCritical, "Warning"
      txtField(3).SetFocus
      GoTo EntryNotOK
   End If

   If cmbField.ListIndex < 0 Then
      MsgBox "Please Specify Transaction Type!!!", vbCritical, "Warning"
      cmbField.SetFocus
      GoTo EntryNotOK
   End If

   Select Case cmbField.ListIndex
   Case 1, 2, 3
      If txtField(17).Text = "" Then
         MsgBox "Account No. is Required!!!" & vbCrLf & _
         "Please Specify to Complete Transaction!!!", vbCritical, "Warning"
         txtField(17).SetFocus
         GoTo EntryNotOK
      End If

'      If txtField(2).Text = "" Then
'         MsgBox "O.R. No. is Required!!!" & vbCrLf & _
'                "Please Specify to Complete Transaction!!!", vbCritical, "Warning"
'         txtField(2).SetFocus
'         GoTo EntryNotOK
'      End If
   Case 4, 5, 6, 7, 8
      If txtField(2).Text = "" Then
         MsgBox "O.R. No. is Required!!!" & vbCrLf & _
         "Please Specify to Complete Transaction!!!", vbCritical, "Warning"
         txtField(2).SetFocus
         GoTo EntryNotOK
      End If

      If txtField(9).Text = "" Then
         MsgBox "Engine No is Required!!!" & vbCrLf & _
                  "Please specify a Valid OR No to Complete Transaction!!!", vbCritical, "Warning"
         txtField(9).SetFocus
         GoTo EntryNotOK
      End If
   End Select

   If lblTotal.Caption = 0# Then
      MsgBox "No Amount to Paid!!!" & vbCrLf & _
             "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
      GoTo EntryNotOK
   End If

   If txtField(16).Text = 0# Then
      MsgBox "Invalid Amount Tendered!!!" & vbCrLf & _
      "Pleae Enter Correct Amount!!!", vbCritical, "Warning"
      GoTo EntryNotOK
   End If

   oTrans.Master("nTranTotl") = CDbl(lblTotal.Caption)
   With cmbField
      If .ListIndex = 3 Or .ListIndex = 4 Then
         If OptField(0).Value Then oTrans.Master("cRegisTyp") = 0
         If OptField(1).Value Then oTrans.Master("cRegisTyp") = 1
      End If
   End With

EntryOK:
   isEntryOk = True
   Exit Function
EntryNotOK:
   isEntryOk = False
End Function

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   Dim lsOldProc As String

   lsOldProc = "txtField_Validate"
   ''On Error GoTo errProc

   With txtField(Index)
      .Text = TitleCase(.Text)

      Select Case Index
      Case 1
         If Not IsDate(.Text) Then .Text = oApp.ServerDate
         .Text = Format(.Text, "MMMM DD, YYYY")
      Case 2
         If Not IsNumeric(.Text) Then txtField(Index).Text = ""
      Case 6, 9
         If Index = 9 Then
            If Trim(txtField(6).Text) <> "" And _
               Trim(txtField(9).Text) = "" Then txtField(6).Text = ""
         End If

         Select Case cmbField.ListIndex
         Case -1, 1, 2, 3
            txtField(6).Text = ""
            txtField(9).Text = ""
            txtField(10).Text = ""
            txtField(11).Text = ""
            txtField(12).Text = ""
            OptField(0).Value = False
            OptField(1).Value = False
         End Select
      Case 13, 14, 15
         If IsNumeric(.Text) = False Then
            .Text = "0.00"
         Else
            If .Text > 999999.99 Then
               .Text = "0.00"
            Else
               .Text = Format(.Text, "#,##0.00")
            End If
         End If

         lblTotal.Caption = Format((CDbl(txtField(13).Text) + CDbl(txtField(15).Text)), "#,##0.00")
         txtField(16).Text = lblTotal.Caption
      Case 16
         If IsNumeric(.Text) = False Then
            .Text = "0.00"
         Else
            If .Text > 99999999.99 Then
               .Text = "0.00"
            Else
               .Text = Format(.Text, "#,##0.00")
            End If
         End If
      Case 80, 81
         If .Text = "" Then
            ClearFields
            Exit Sub
         End If
                  
         If Trim(UCase(.Text)) <> Trim(UCase(.Tag)) Then
            If oTrans.SearchTransaction(.Text, IIf(Index = 80, True, False)) Then
               LoadMaster
            Else
               ClearFields
               .SetFocus
            End If
         End If
      End Select
      
      Select Case Index
      Case 13, 14, 15
         oTrans.Master(Index) = CDbl(.Text)
      Case 17
         If .Text <> "" Then oTrans.Master(Index) = .Text
      Case 13, 16
      Case Else
         If Index < 80 Then oTrans.Master(Index) = .Text
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
   Dim loTxt As TextBox
   With oTrans
      For Each loTxt In txtField
         pnCtr = loTxt.Index
         Select Case pnCtr
         Case 0
            txtField(pnCtr).Text = Format(.Master("sTransNox"), IIf(Len(oApp.BranchCode) = 2, "@@@@-@@@@@@", "@@@@@@-@@@@@@"))
         Case 2, 80
            txtField(pnCtr).Text = .Master("sORNoxxxx")
            txtField(pnCtr).Tag = txtField(pnCtr).Text
         Case 1
            txtField(pnCtr).Text = Format(.Master(pnCtr), "MMMM DD, YYYY")
         Case 3, 81
            txtField(pnCtr).Text = .Master("xFullName")
            txtField(pnCtr).Tag = txtField(pnCtr).Text
         Case 7, 8, 18 To 29
         Case 13, 14, 15
            txtField(pnCtr).Text = Format(.Master(pnCtr), "#,##0.00")
         Case 16
            txtField(pnCtr).Text = Format(.Master("nTranTotl") - .AdvancePayment, "#,##0.00")
         Case 30 To 34, 36, 37, 38
            txtField(pnCtr).Text = IIf(IsNull(.Master(pnCtr)), "0.00", .Master(pnCtr))
         Case 35
            txtField(pnCtr).Text = IIf(IsNull(.Master(pnCtr)), "0", .Master(pnCtr))
         Case Else
            txtField(pnCtr).Text = IIf(IsNull(.Master(pnCtr)), "", .Master(pnCtr))
         End Select
      Next

      Label2.Caption = Format(TransStat(oTrans.Master("cTranStat")), ">")
      lblTotal.Caption = Format(.Master("nTranTotl"), "#,##0.00")
      lblAdvPayment.Caption = Format(.AdvancePayment, "#,##0.00")
      
      Select Case oTrans.Master("cTranType")
      Case "o"
         cmbField.ListIndex = 9
      Case Else
         cmbField.ListIndex = oTrans.Master("cTranType")
      End Select
      
      txtField(40).Enabled = False
      txtField(41).Enabled = False
      
      If cmbField.ListIndex = 4 Or cmbField.ListIndex = 5 Or cmbField.ListIndex = 6 Then
         OptField(IIf(Trim(oTrans.Master("cRegisTyp")) = "", 0, oTrans.Master("cRegisTyp"))).Value = True
         cmbcRegisFrm.ListIndex = oTrans.Master("cRegisFrm")
      
         If cmbcRegisFrm.ListIndex = 1 Or cmbcRegisFrm.ListIndex = 2 Then
            txtField(40).Enabled = True
            txtField(41).Enabled = True
         End If
      End If
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

Private Function PrintOR() As Boolean
   Dim lors As ADODB.Recordset
   Dim lsOldProc As String
   Dim lnCtr As Integer
   Dim p_oRepViewer As frmRepPreview
   Dim lsRemarks1 As String
   Dim lnTranAmtx As Double
   Dim lsSQL As String
   
   lsOldProc = "PrintOR"
   ''On Error GoTo errProc

   With oTrans
      'Initialize the report
      If .Master("sSourceCd") = "SPOR" Then
         Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\OfficialReceiptSPOR.rpt")
      Else
         Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\OfficialReceipt.rpt")
      End If
      
      oReport.DiscardSavedData
      oReport.FieldMappingType = crAutoFieldMapping
      'Disable parameter prompting
      oReport.EnableParameterPrompting = False
'      oReport.Database.SetDataSource loRS
      
      'Set Header
      oReport.Sections("D").ReportObjects("txtCustName").SetText .Master("xFullName")
      oReport.Sections("D").ReportObjects("txtAddressx").SetText .Master("xAddressx")
      oReport.Sections("D").ReportObjects("txtTinNo").SetText ""
      oReport.Sections("D").ReportObjects("txtDtIssued").SetText Format(.Master("dTransact"), "DD-MMM-YYYY")
      'Set details
      lnTranAmtx = .Master("nTranAmtx") + .Master("nOthersxx")
      
      Select Case .Master("cTranType")
      Case "0"
'         lsRemarks1 = .Master("sModelNme") & "/" & .Master("sColorNme") & " Engine#: " & .Master("sEngineNo")
      Case "4", "5", "6", "7", "8"
         lsRemarks1 = "Engine#: " & .Master("sEngineNo")
      Case "1"
         oReport.Sections("D").ReportObjects("txtTranType").SetText "Payment"
         lsRemarks1 = "AcctNmbr#: " & .Master("sAcctNmbr")
      Case "2", "3"
         lsRemarks1 = "AcctNmbr#: " & .Master("sAcctNmbr")
      End Select
      
      If oTrans.Master("cTranType") = 4 Or oTrans.Master("cTranType") = 5 Then
         lsRemarks1 = lsRemarks1 & " - " & cmbcRegisFrm.List(IFNull(.Master("cRegisFrm"), 0))
      End If
      
      If .Master("sSourceCd") = "MCCO" Then
         oReport.Sections("D").ReportObjects("txtTranType").SetText "MC Reservation"
      ElseIf .Master("sSourceCd") = "SPOR" Then
         oReport.Sections("D").ReportObjects("txtTranType").SetText "SP Reservation"
      ElseIf .Master("sSourceCd") = "SPJO" Then
         oReport.Sections("D").ReportObjects("txtTranType").SetText "Payment"
         lsRemarks1 = "Services"
      ElseIf .Master("sSourceCd") = "MCSl" Then
         If .Master("nTranAmtx") = 0 Then
            ' XerSys - 2014-10-17
            '  Allowing OR for Check Payment must retrieve the amount through the parent transaction
            lsSQL = "SELECT nAmtPaidx" & _
                     " FROM MC_SO_Master" & _
                     " WHERE sTransNox = " & strParm(.Master("sReferNox"))
            Set lors = New Recordset
            lors.Open lsSQL, oApp.Connection, , , adCmdText
            
            If Not lors.EOF Then
               lnTranAmtx = lors("nAmtPaidx")
            End If
         End If
         oReport.Sections("D").ReportObjects("txtTranType").SetText cmbField.List(Val(.Master("cTranType")))
      Else
         oReport.Sections("D").ReportObjects("txtTranType").SetText cmbField.List(Val(.Master("cTranType")))
      End If
      
      oReport.Sections("D").ReportObjects("txtRemarks1").SetText lsRemarks1
      oReport.Sections("D").ReportObjects("txtAmtPaidx").SetText Format(lnTranAmtx + .Master("nDiscount"), "#,#0.00")
      oReport.Sections("D").ReportObjects("txtDiscount").SetText IIf(.Master("nDiscount") = 0, "-", Format(.Master("nDiscount"), "#,#0.00"))
      oReport.Sections("D").ReportObjects("nTranTotal").SetText Format(lnTranAmtx, "#,#0.00")
      
      If oTrans.Master("cTranType") = 0 Then
         If .Master("sSourceCd") = "MCCO" Then
            oReport.Sections("D").ReportObjects("txtRemarks2").SetText ""
            oReport.Sections("D").ReportObjects("txtRemarks3").SetText IFNull(.Master("sRemarksx"))
         ElseIf .Master("sSourceCd") = "SPOR" Then
            oReport.Sections("D").ReportObjects("txtRemarks2").SetText ""
            oReport.Sections("D").ReportObjects("txtRemarks3").SetText IFNull(.Master("sRemarksx"))
         Else
            oReport.Sections("D").ReportObjects("txtRemarks2").SetText getGiveAways(.Master("sReferNox"))
            oReport.Sections("D").ReportObjects("txtRemarks3").SetText IFNull(.Master("sRemarksx"))
         End If
      Else
         oReport.Sections("D").ReportObjects("txtRemarks2").SetText IFNull(.Master("sRemarksx"))
         If txtField(44).Text <> "" Then
            oReport.Sections("D").ReportObjects("txtRemarks3").SetText "Paid By: " + txtField(44).Text
         Else
            oReport.Sections("D").ReportObjects("txtRemarks3").SetText ""
         End If
      End If
   
      oReport.Sections("D").ReportObjects("txtValClt").SetText .Master("xFullName")
      oReport.Sections("D").ReportObjects("txtValDoc").SetText "OR# " + .Master("sORNoxxxx")
      oReport.Sections("D").ReportObjects("txtValAmt").SetText Format(.AdvancePayment - lnTranAmtx, "#,#0.00")
      oReport.Sections("D").ReportObjects("txtValDate").SetText Format(.Master("dTransact"), "DD-MMM-YYYY")
      oReport.ParameterFields.GetItemByName("nSaleTotl").AddCurrentValue lnTranAmtx
      
      oReport.Sections("D").ReportObjects("txtAdvPaym").SetText Format(.AdvancePayment, "#,#0.00")
      oReport.Sections("D").ReportObjects("txtDue").SetText Format(lnTranAmtx - .AdvancePayment, "#,#0.00")
      
      oReport.PrintOutEx False, 1, , 1, 1
      
'      Set p_oRepViewer = New frmRepPreview
'      With p_oRepViewer
'         .CRViewer91.ReportSource = oReport
'         .Show
'         .CRViewer91.ViewReport
'
'         While .CRViewer91.IsBusy
'            DoEvents
'         Wend
'      End With
   End With
   
   PrintOR = True

endProc:
   Set oReport = Nothing
   Set lors = Nothing
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"
End Function

Private Function getGiveAways(ByVal fsTransNox As String) As String
   Dim lsSQL As String
   Dim lors As Recordset
   Dim lsReleased As String
   Dim lsIssuance As String
   
   'iMac 2017.05.31
   If oTrans.Master("sSourceCd") = "SPOR" Then
      getGiveAways = ""
      Exit Function
   End If
   
   If oTrans.Master("sSourceCd") = "CChk" And oTrans.Master("cTranType") = 0 Then
      lsSQL = "SELECT e.nQuantity, e.nGivenxxx, f.sDescript" & _
          " FROM Receipt_Master a" & _
          ", Checks_Received b" & _
          ", Provisionary_Receipt_Master c" & _
          ", MC_SO_Master d" & _
          ", MC_SO_GiveAways e" & _
               " LEFT JOIN Spareparts f ON e.sPartsIDx = f.sPartsIDx" & _
          " WHERE a.sReferNox = " & strParm(fsTransNox) & _
            " AND a.sSerialID = " & strParm(oTrans.Master("sSerialID")) & _
            " AND a.sReferNox = b.sTransNox" & _
            " AND b.sReferNox = c.sTransNox" & _
            " AND c.sReferNox = d.sTransNox" & _
            " AND d.sTransNox = e.sTransNox" & _
            " AND a.cTranType = 0 " & _
            " AND a.sSourceCd = 'CChk' " & _
            " AND e.cGAwyStat IN ('0', '3')" & _
          " ORDER BY e.nGivenxxx  DESC"
   Else
      lsSQL = "SELECT a.nQuantity, a.nGivenxxx, b.sDescript" & _
          " FROM MC_SO_GiveAways a" & _
               " LEFT JOIN Spareparts b ON a.sPartsIDx = b.sPartsIDx" & _
          " WHERE a.sTransNox = " & strParm(fsTransNox) & _
            " AND a.cGAwyStat IN ('0', '3')" & _
          " ORDER BY a.nGivenxxx  DESC"
   End If
   Set lors = oApp.Connection.Execute(lsSQL, , adCmdText)
   Debug.Print lsSQL
   Do Until lors.EOF

      If lors("nGivenxxx") > 0 Then
         lsReleased = lsReleased & ", " & IIf(lors("nGivenxxx") = 1, "", lors("nGivenxxx")) & " " & lors("sDescript")
      End If

      If (lors("nQuantity") - lors("nGivenxxx")) > 0 Then
         lsIssuance = lsIssuance & ", " & IIf((lors("nQuantity") - lors("nGivenxxx")) = 1, "", (lors("nQuantity") - lors("nGivenxxx"))) & " " & lors("sDescript")
      End If
      
      lors.MoveNext
   Loop
   
   If Len(lsReleased) > 0 Then
      getGiveAways = "GIVEAWAYS:" & Mid(lsReleased, 2)
      If Len(lsIssuance) > 0 Then
         getGiveAways = getGiveAways & " / FOR ISSUANCE:" & Mid(lsIssuance, 2)
      End If
   Else
      If Len(lsIssuance) > 0 Then
         getGiveAways = "FOR ISSUANCE:" & Mid(lsIssuance, 2)
      End If
   End If
End Function

Private Function PrintAR() As Boolean
    Dim lors As ADODB.Recordset
   Dim lsOldProc As String
   Dim lnCtr As Integer
   Dim p_oRepViewer As frmRepPreview
   Dim lsRemarks1 As String
   Dim lnTranAmtx As Double
   Dim lsReleased As String
   Dim ls4Release As String
   Dim lsSQL As String
   Dim lorecord As Recordset
   Dim loReceipt As Recordset
   
   lsOldProc = "PrintAR"
   '''On Error GoTo errProc

   lsSQL = " Select b.sWarrntNo" & _
            " from MC_SO_Detail a" & _
            ", MC_Serial b " & _
            " WHERE a.sTransNox = " & strParm(oTrans.Master("sTransNox")) & _
            " AND a.sSerialId = b.sSerialId"
   Debug.Print lsSQL
   Set lorecord = New Recordset
   lorecord.Open lsSQL, oApp.Connection, adOpenKeyset, adLockOptimistic, adCmdText
   
   lsSQL = " Select a.nTranAmtx nTranTotl, a.sTransNox, a.nTranAmtx nAmtPaidX" & _
            " from Receipt_Master a" & _
            " WHERE a.sTransNox = " & strParm(oTrans.Master("sTransNox")) & _
            " AND a.sORNoxxxx = " & strParm(oTrans.Master("sORNoxxxx"))
            
   Debug.Print lsSQL
   Set loReceipt = New Recordset
   loReceipt.Open lsSQL, oApp.Connection, adOpenKeyset, adLockOptimistic, adCmdText
   
   
   With oTrans
      'Initialize the report
      Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\OfficialReceiptAR.rpt")
      oReport.DiscardSavedData
      oReport.FieldMappingType = crAutoFieldMapping
      'Disable parameter prompting
      oReport.EnableParameterPrompting = False
      
'she 2020-10-23 for AR
    'Set Header
      oReport.Sections("D").ReportObjects("txtCustName").SetText .Master("xFullName")
      oReport.Sections("D").ReportObjects("txtAddressx").SetText .Master("xAddressx")
      oReport.Sections("D").ReportObjects("txtDtIssued").SetText Format(.Master("dTransact"), "DD-MMM-YYYY")
      'Set details
      oReport.Sections("D").ReportObjects("txtAmtPaidx").SetText Format(.Master("nTranTotl"), "#,#0.00")
      oReport.Sections("D").ReportObjects("txtNum2text").SetText NumToText(.Master("nTranTotl"))
      oReport.ParameterFields.GetItemByName("nSaleTotl").AddCurrentValue lnTranAmtx
      
      oReport.PrintOutEx False, 1, , 1, 1
      

   End With
   PrintAR = True
   
endProc:
   Set oReport = Nothing
   Set lors = Nothing
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"
End Function

