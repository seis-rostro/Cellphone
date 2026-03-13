VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmCashierTrans 
   BorderStyle     =   0  'None
   Caption         =   "Office Collection"
   ClientHeight    =   7935
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10110
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   10110
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   7260
      Index           =   0
      Left            =   1590
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   8400
      _ExtentX        =   14817
      _ExtentY        =   12806
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   44
         Left            =   1665
         TabIndex        =   32
         Top             =   3975
         Width           =   6435
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   41
         Left            =   1665
         TabIndex        =   46
         Top             =   6765
         Width           =   6420
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   40
         Left            =   1665
         TabIndex        =   44
         Top             =   6450
         Width           =   6420
      End
      Begin VB.ComboBox cmbcRegisFrm 
         Height          =   315
         ItemData        =   "frmCashierTrans.frx":0000
         Left            =   5790
         List            =   "frmCashierTrans.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   2505
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   37
         Left            =   6675
         MaxLength       =   30
         TabIndex        =   70
         Top             =   5685
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
         Left            =   6675
         MaxLength       =   30
         TabIndex        =   69
         Top             =   5985
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
         Left            =   6675
         MaxLength       =   30
         TabIndex        =   68
         Top             =   5385
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
         Left            =   4080
         MaxLength       =   30
         TabIndex        =   67
         Top             =   5685
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
         Left            =   4080
         MaxLength       =   30
         TabIndex        =   66
         Top             =   5985
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
         Left            =   4080
         MaxLength       =   30
         TabIndex        =   65
         Top             =   5385
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
         Left            =   1500
         MaxLength       =   30
         TabIndex        =   64
         Top             =   5685
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
         Left            =   1500
         MaxLength       =   30
         TabIndex        =   62
         Top             =   5985
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
         Index           =   30
         Left            =   1500
         MaxLength       =   30
         TabIndex        =   61
         Top             =   5385
         Width           =   1425
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
         TabIndex        =   1
         Top             =   165
         Width           =   2310
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   9
         Left            =   1650
         MaxLength       =   50
         TabIndex        =   24
         Top             =   3180
         Width           =   2250
      End
      Begin xrControl.xrFrame xrFrame2 
         Height          =   975
         Index           =   0
         Left            =   1605
         Tag             =   "wt0;fb0"
         Top             =   2835
         Width           =   6600
         _ExtentX        =   11642
         _ExtentY        =   1720
         BorderStyle     =   4
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   10
            Left            =   45
            MaxLength       =   50
            TabIndex        =   26
            Top             =   645
            Width           =   2250
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   12
            Left            =   4230
            MaxLength       =   50
            TabIndex        =   30
            Top             =   645
            Width           =   2250
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   11
            Left            =   4230
            MaxLength       =   50
            TabIndex        =   28
            Top             =   345
            Width           =   2250
         End
         Begin VB.OptionButton OptField 
            Caption         =   "Solo"
            Height          =   195
            Index           =   1
            Left            =   60
            TabIndex        =   21
            Tag             =   "et0;fb0"
            Top             =   60
            Width           =   630
         End
         Begin VB.OptionButton OptField 
            Caption         =   "Tricycle"
            Height          =   195
            Index           =   0
            Left            =   885
            TabIndex        =   22
            Tag             =   "et0;fb0"
            Top             =   60
            Width           =   930
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Color"
            Height          =   285
            Index           =   14
            Left            =   3600
            TabIndex        =   29
            Top             =   690
            Width           =   405
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Model"
            Height          =   285
            Index           =   15
            Left            =   3600
            TabIndex        =   27
            Top             =   405
            Width           =   525
         End
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
         TabIndex        =   38
         Top             =   4890
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
         TabIndex        =   34
         Top             =   4290
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
         Index           =   14
         Left            =   1665
         MaxLength       =   30
         TabIndex        =   36
         Top             =   4590
         Width           =   1860
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   6
         Left            =   1665
         MaxLength       =   50
         TabIndex        =   17
         Top             =   2505
         Width           =   2250
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   5955
         MaxLength       =   6
         TabIndex        =   9
         Top             =   990
         Width           =   2295
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   17
         Left            =   1665
         TabIndex        =   5
         Top             =   975
         Width           =   2310
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   1665
         MaxLength       =   50
         TabIndex        =   3
         Top             =   675
         Width           =   2310
      End
      Begin VB.ComboBox cmbField 
         Height          =   315
         ItemData        =   "frmCashierTrans.frx":0004
         Left            =   5940
         List            =   "frmCashierTrans.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   675
         Width           =   2310
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   1665
         TabIndex        =   11
         Top             =   1425
         Width           =   6420
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
         Left            =   5670
         MaxLength       =   30
         TabIndex        =   42
         Tag             =   "ht0"
         Top             =   4800
         Width           =   2430
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   1665
         MaxLength       =   128
         TabIndex        =   15
         Top             =   2025
         Width           =   6420
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   1665
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1725
         Width           =   6420
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
         Left            =   315
         TabIndex        =   31
         Top             =   3990
         Width           =   900
      End
      Begin VB.Label lblOtherC 
         BackStyle       =   0  'Transparent
         Caption         =   "Co Owner #2"
         Height          =   285
         Index           =   1
         Left            =   315
         TabIndex        =   45
         Top             =   6795
         Width           =   1200
      End
      Begin VB.Label lblOtherC 
         BackStyle       =   0  'Transparent
         Caption         =   "Co Owner #1"
         Height          =   285
         Index           =   0
         Left            =   315
         TabIndex        =   43
         Top             =   6480
         Width           =   1200
      End
      Begin VB.Shape Shape2 
         Height          =   720
         Index           =   1
         Left            =   120
         Top             =   6390
         Width           =   8130
      End
      Begin VB.Label lblcRegisFrm 
         BackStyle       =   0  'Transparent
         Caption         =   "Regis. Form"
         Height          =   285
         Left            =   4755
         TabIndex        =   18
         Top             =   2505
         Width           =   960
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Down Payment"
         Height          =   285
         Index           =   20
         Left            =   315
         TabIndex        =   63
         Top             =   5685
         Width           =   1095
      End
      Begin VB.Shape Shape3 
         Height          =   1035
         Index           =   2
         Left            =   120
         Top             =   5310
         Width           =   8130
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Rebate"
         Height          =   285
         Index           =   27
         Left            =   5670
         TabIndex        =   60
         Top             =   5985
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount Due"
         Height          =   285
         Index           =   26
         Left            =   5670
         TabIndex        =   59
         Top             =   5685
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cash Balance"
         Height          =   285
         Index           =   25
         Left            =   3060
         TabIndex        =   58
         Top             =   5685
         Width           =   1005
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Term"
         Height          =   285
         Index           =   24
         Left            =   3060
         TabIndex        =   57
         Top             =   5985
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Balance"
         Height          =   285
         Index           =   23
         Left            =   5670
         TabIndex        =   56
         Top             =   5385
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Mon Amort"
         Height          =   285
         Index           =   22
         Left            =   3060
         TabIndex        =   55
         Top             =   5385
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Gross Amount"
         Height          =   285
         Index           =   21
         Left            =   315
         TabIndex        =   54
         Top             =   5985
         Width           =   1080
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "PN Value"
         Height          =   285
         Index           =   19
         Left            =   315
         TabIndex        =   53
         Top             =   5385
         Width           =   975
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   285
         Left            =   1755
         Tag             =   "et0;ht2"
         Top             =   270
         Width           =   2310
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
         Left            =   165
         TabIndex        =   0
         Top             =   -2460
         Width           =   1395
      End
      Begin VB.Shape Shape3 
         Height          =   1350
         Index           =   1
         Left            =   120
         Top             =   3915
         Width           =   8130
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Frame No"
         Height          =   285
         Index           =   7
         Left            =   615
         TabIndex        =   25
         Top             =   3540
         Width           =   795
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Registration Type"
         Height          =   285
         Index           =   18
         Left            =   315
         TabIndex        =   20
         Top             =   2850
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Amount"
         Height          =   195
         Index           =   17
         Left            =   315
         TabIndex        =   33
         Top             =   4335
         Width           =   540
      End
      Begin VB.Shape Shape3 
         Height          =   1425
         Index           =   0
         Left            =   120
         Top             =   2430
         Width           =   8130
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Discoun&t"
         Height          =   195
         Index           =   16
         Left            =   315
         TabIndex        =   35
         Top             =   4635
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Others"
         Height          =   195
         Index           =   5
         Left            =   315
         TabIndex        =   37
         Top             =   4935
         Width           =   465
      End
      Begin VB.Shape Shape2 
         Height          =   1035
         Index           =   0
         Left            =   120
         Top             =   1350
         Width           =   8130
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Engine No"
         Height          =   285
         Index           =   13
         Left            =   615
         TabIndex        =   23
         Top             =   3240
         Width           =   795
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Insurance Type"
         Height          =   285
         Index           =   0
         Left            =   315
         TabIndex        =   16
         Top             =   2505
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "O.R. No."
         Height          =   285
         Index           =   2
         Left            =   4905
         TabIndex        =   8
         Top             =   1020
         Width           =   930
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
         Left            =   5670
         TabIndex        =   40
         Tag             =   "ht0;ft0"
         Top             =   4290
         Width           =   2430
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   285
         Index           =   12
         Left            =   315
         TabIndex        =   14
         Top             =   2070
         Width           =   690
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Account No."
         Height          =   285
         Index           =   10
         Left            =   165
         TabIndex        =   4
         Top             =   1020
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
         Left            =   4350
         TabIndex        =   39
         Top             =   4410
         Width           =   720
      End
      Begin VB.Label Label1 
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
         Height          =   285
         Index           =   4
         Left            =   4395
         TabIndex        =   41
         Top             =   4920
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Trans. Type"
         Height          =   285
         Index           =   6
         Left            =   4905
         TabIndex        =   6
         Top             =   690
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cust. Address"
         Height          =   195
         Index           =   11
         Left            =   315
         TabIndex        =   12
         Top             =   1770
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Date"
         Height          =   285
         Index           =   1
         Left            =   165
         TabIndex        =   2
         Top             =   720
         Width           =   1245
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name"
         Height          =   285
         Index           =   3
         Left            =   315
         TabIndex        =   10
         Top             =   1455
         Width           =   1200
      End
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   600
      Index           =   3
      Left            =   90
      TabIndex        =   51
      Top             =   5715
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
      Picture         =   "frmCashierTrans.frx":0008
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   90
      TabIndex        =   47
      Top             =   3825
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
      Picture         =   "frmCashierTrans.frx":0782
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   90
      TabIndex        =   52
      Top             =   5715
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
      Picture         =   "frmCashierTrans.frx":0EFC
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   600
      Index           =   1
      Left            =   90
      TabIndex        =   48
      Top             =   4455
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
      Picture         =   "frmCashierTrans.frx":1676
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   90
      TabIndex        =   49
      Top             =   4455
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
      Picture         =   "frmCashierTrans.frx":1DF0
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   90
      TabIndex        =   50
      Top             =   5085
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
      Picture         =   "frmCashierTrans.frx":256A
   End
End
Attribute VB_Name = "frmCashierTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmCashierTrans"

Private WithEvents oTrans As clsORReceipt
Attribute oTrans.VB_VarHelpID = -1
Private oReceipt As ggcReceipt.Receipt
Private oSkin As clsFormSkin

Dim pnIndex As Integer, pnCtr As Integer
Dim pbMoveCombo As Boolean

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
         lblOtherC(0) = "Co Brwr #1"
         lblOtherC(1) = "Co Brwr #2"
      Case 4, 5, 6, 7, 8, 9
         lblOtherC(0) = "Co Owner #1"
         lblOtherC(1) = "Co Owner #2"
         lbEnabled = True
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
   
   If cmbField.ListIndex = 0 Then
      cmbField.ListIndex = 1
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

   lsOldProc = "cmdButton_Click"
   '''On Error GoTo errProc

   If Not pbMoveCombo Then txtField_LostFocus pnIndex
   Select Case Index
   Case 0
      If isEntryOk Then
         If oTrans.Master("nTranTotl") > 0# Then Call Receipt
         If oReceipt.Cancelled Then Exit Sub
         oTrans.Master("sSystemCd") = "GL"

         If oTrans.SaveTransaction = True Then
            MsgBox "Record Updated Successfully!!!", vbInformation, "Notice"
                        
            'Print Slip if in Real/Time Mode
            If IFNull(oApp.getConfiguration("cRealTime", oTrans.Branch), "0") = "1" Then
'               MsgBox "Please load a 1/4 sheet of short coupon(short)!"
'               Call PrintTransSlip(oTrans)
               
               lnRep = MsgBox("Print Transaction!!!", vbYesNo + vbQuestion, "Confirm")
               If lnRep = vbYes Then
                  MsgBox "Please mount the OR..."
                  'she 11-26-2020
                  'check if what is the trantype of transaction para alamkung ano ang tatawagin na print
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
               End If
            End If
                        
            Call cmdButton_Click(4) 'new
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
      If lblTotal.Caption > 0# And _
         txtField(16).Text > 0# Then
         Call Receipt
      End If
   Case 3
      lnRep = MsgBox("Transaction is in Update Mode!!!" & vbCrLf & _
                     "Do you want to Cancel Transaction!!!", vbYesNo + vbQuestion, "Confirm")

      If lnRep = vbYes Then
         ClearFields
         initButton xeModeReady
      Else
         txtField(pnIndex).SetFocus
      End If
   Case 4
      oTrans.NewTransaction
      initButton xeModeAddNew
      ClearFields
      
      txtField(1).SetFocus
   Case 5
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
   '''On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oTrans = New clsORReceipt
   Set oTrans.AppDriver = oApp
   
   oTrans.Branch = oApp.BranchCode
   oTrans.InitTransaction
   oTrans.NewTransaction

   Set oReceipt = New ggcReceipt.Receipt
   Set oReceipt.AppDriver = oApp
   oReceipt.InitReceipt
   oReceipt.CheckFieldEnabled = False

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransaction
   
   InitForm
   ClearFields
   cmbField.ListIndex = 9

   initButton xeModeAddNew
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
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
      ElseIf Trim(oTrans.Master(Index)) = "" Then
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
End Sub

Private Sub ClearFields()
   Dim loTxt As TextBox
   
   For Each loTxt In txtField
      pnCtr = loTxt.Index
      Select Case pnCtr
      Case 0
         txtField(pnCtr).Text = Format(oTrans.Master(pnCtr), IIf(Len(oApp.BranchCode) = 2, "@@@@-@@@@@@", "@@@@@@-@@@@@@"))
      Case 1
         txtField(pnCtr).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
'      Case 2
'         txtField(pnCtr).Text = oTrans.Master("sORNoxxxx")
      Case 7, 8, 18 To 29
      Case 13, 14, 15, 16, 30 To 38
         txtField(pnCtr).Text = Format(oTrans.Master(pnCtr), "#,##0.00")
      Case Else
         txtField(pnCtr).Text = oTrans.Master(pnCtr)
'         txtField(pnCtr).Text = ""
      End Select
   Next
   cmbField.Enabled = False
   cmbField.ListIndex = 9
   oTrans.Master("cTranType") = 9
   
   cmbcRegisFrm.Visible = False
   lblcRegisFrm.Visible = False
   cmbcRegisFrm.ListIndex = 0
   
   lblTotal.Caption = Format(oTrans.Master("nTranTotl"), "#,##0.00")
   xrFrame2(0).Enabled = False

   txtField(40).Enabled = False
   txtField(41).Enabled = False

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
   '''On Error GoTo errProc

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
   End Select
End Sub

Private Sub initButton(lnStat As Integer)
   Dim lbShow As Boolean

   lbShow = IIf(lnStat = 0, False, True)
   cmdButton(0).Visible = lbShow
   cmdButton(1).Visible = lbShow
   cmdButton(3).Visible = lbShow
   
'   cmbField.Enabled = lbShow
   
   cmdButton(4).Visible = Not lbShow
   cmdButton(5).Visible = Not lbShow

   xrFrame1(0).Enabled = lbShow

   If Not lbShow Then cmdButton(4).SetFocus
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
      .CheckAmount = oTrans.Checks("nAmountxx")
      If oTrans.Checks("nAmountxx") > 0 Then .CashAmount = oTrans.Master("nTranTotl") - oTrans.Checks("nAmountxx")
      .AmountPaid = oTrans.Master("nTranTotl")
      .Remarks = txtField(5).Text
      .SystemCd = "GL"
      .PRNo = oTrans.PRNumber

      If oTrans.Checks("nAmountxx") > 0# Then
         oReceipt.Checks("sCheckNox") = oTrans.Checks("sCheckNox")
         oReceipt.Checks("sAcctNoxx") = oTrans.Checks("sAcctNoxx")
         oReceipt.Checks("sBankIDxx") = oTrans.Checks("sBankIDxx")
         oReceipt.Checks("dCheckDte") = oTrans.Checks("dCheckDte")
         oReceipt.Checks("nAmountxx") = oTrans.Checks("nAmountxx")
         oReceipt.Checks("sPRNoxxxx") = oTrans.Checks("sPRNoxxxx")
         oTrans.PRNumber = .PRNo
      Else
         oReceipt.Checks("sCheckNox") = ""
         oReceipt.Checks("sAcctNoxx") = ""
         oReceipt.Checks("sBankIDxx") = ""
         oReceipt.Checks("dCheckDte") = oApp.ServerDate
         oReceipt.Checks("nAmountxx") = 0#
         oReceipt.Checks("sPRNoxxxx") = ""
         .CheckAmount = 0
      End If
      
      Set .GiftCoupon = oTrans.GiftCoupon.Clone
            
      .ShowReceipt
   End With

   If oTrans.EditMode <> xeModeAddNew Then Exit Sub
   
   If Not oReceipt.Cancelled Then
      With oTrans
         If oReceipt.CheckAmount > 0# Then
            .Checks("sCheckNox") = oReceipt.Checks("sCheckNox")
            .Checks("sAcctNoxx") = oReceipt.Checks("sAcctNoxx")
            .Checks("sBankIDxx") = oReceipt.Checks("sBankIDxx")
            .Checks("dCheckDte") = oReceipt.Checks("dCheckDte")
            .Checks("nAmountxx") = oReceipt.Checks("nAmountxx")
         End If
         
         .Master("sORNoxxxx") = oReceipt.ORNo
         .Master("sRemarksx") = oReceipt.Remarks
         .Checks("nAmountxx") = oReceipt.CheckAmount
         .PRNumber = oReceipt.PRNo
      End With
      
      Set oTrans.GiftCoupon = oReceipt.GiftCoupon
      
      txtField(2).Text = oReceipt.ORNo
      txtField(5).Text = oReceipt.Remarks
   End If
End Sub

Private Function isEntryOk() As Boolean
   Dim lsTranType As String

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
   

   If Not isTransValid(CDate(txtField(1)), "MCSc", Trim(txtField(2)), CDbl(txtField(13))) Then GoTo EntryNotOK
   
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
   '''On Error GoTo errProc

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
            If .Text > 9999999.99 Then
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
      End Select
      
      Select Case Index
      Case 13, 14, 15
         oTrans.Master(Index) = CDbl(.Text)
      Case 17
         If .Text <> "" Then oTrans.Master(Index) = .Text
      Case 16
      Case Else
         oTrans.Master(Index) = .Text
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
   Dim lsReleased As String
   Dim ls4Release As String

   lsOldProc = "PrintOR"
   '''On Error GoTo errProc

   With oTrans
      'Initialize the report
      Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\OfficialReceipt.rpt")
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
'      oReport.Sections("D").ReportObjects("txtCustName").SetText ""
'      oReport.Sections("D").ReportObjects("txtAddressx").SetText ""
'      oReport.Sections("D").ReportObjects("txtTinNo").SetText ""
'      oReport.Sections("D").ReportObjects("txtDtIssued").SetText ""
      
      'Set details
      oReport.Sections("D").ReportObjects("txtTranType").SetText cmbField.List(Val(.Master("cTranType")))
      Select Case .Master("cTranType")
      Case "0"
         lsRemarks1 = "Engine#: " & .Master("sEngineNo")
         
      Case "4", "5", "6", "7", "8"
         lsRemarks1 = "Engine#: " & .Master("sEngineNo")
      Case "1"
         'kalyptus - 2013.05.14
         'Remove Monthly Payment from the print out
         oReport.Sections("D").ReportObjects("txtTranType").SetText "Payment"
         lsRemarks1 = "AcctNmbr#: " & .Master("sAcctNmbr")
      Case "2", "3"
         lsRemarks1 = "AcctNmbr#: " & .Master("sAcctNmbr")
      End Select
      
      If oTrans.Master("cTranType") = 4 Or oTrans.Master("cTranType") = 5 Then
         lsRemarks1 = lsRemarks1 & " - " & cmbcRegisFrm.List(IFNull(.Master("cRegisFrm"), 0))
      End If
      
      oReport.Sections("D").ReportObjects("txtRemarks1").SetText lsRemarks1
      oReport.Sections("D").ReportObjects("txtAmtPaidx").SetText Format(.Master("nTranAmtx") + .Master("nDiscount") + .Master("nOthersxx"), "#,#0.00")
      oReport.Sections("D").ReportObjects("txtDiscount").SetText IIf(.Master("nDiscount") = 0, "-", Format(.Master("nDiscount"), "#,#0.00"))
      oReport.Sections("D").ReportObjects("nTranTotal").SetText Format(.Master("nTranAmtx") + .Master("nOthersxx"), "#,#0.00")
      oReport.Sections("D").ReportObjects("txtRemarks2").SetText .Master("sRemarksx")
      'kalyptus - 2013.05.14
      'Remove txtRemarks3 from the printing of the receipt
      If txtField(44).Text <> "" Then
         oReport.Sections("D").ReportObjects("txtRemarks3").SetText "Paid By: " + txtField(44).Text
      Else
         oReport.Sections("D").ReportObjects("txtRemarks3").SetText ""
      End If
      oReport.Sections("D").ReportObjects("txtValClt").SetText .Master("xFullName")
      oReport.Sections("D").ReportObjects("txtValDoc").SetText "OR# " + .Master("sORNoxxxx")
      oReport.Sections("D").ReportObjects("txtValAmt").SetText Format(.Master("nTranAmtx") + .Master("nOthersxx"), "#,#0.00")
      oReport.Sections("D").ReportObjects("txtValDate").SetText Format(.Master("dTransact"), "DD-MMM-YYYY")
      
      lnTranAmtx = .Master("nTranAmtx") + .Master("nOthersxx")
      oReport.ParameterFields.GetItemByName("nSaleTotl").AddCurrentValue lnTranAmtx
      
      'iMac 2017.06.03
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
