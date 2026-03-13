VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0"; "xrGridControl.ocx"
Begin VB.Form frmMPCreditApprovalNeo 
   BorderStyle     =   0  'None
   Caption         =   "Credit Application Approval/Disapproval"
   ClientHeight    =   7125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11985
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7125
   ScaleWidth      =   11985
   ShowInTaskbar   =   0   'False
   Begin xrGridEditor.GridEditor GridEditor1 
      Height          =   1260
      Left            =   1710
      TabIndex        =   102
      TabStop         =   0   'False
      Tag             =   "et0;eb0;et0;bc2"
      Top             =   5655
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   2223
      AllowBigSelection=   -1  'True
      AutoAdd         =   0   'False
      AutoNumber      =   -1  'True
      BACKCOLOR       =   -2147483643
      BACKCOLORBKG    =   8421504
      BACKCOLORFIXED  =   -2147483633
      BACKCOLORSEL    =   -2147483635
      BORDERSTYLE     =   1
      COLS            =   2
      FILLSTYLE       =   0
      FIXEDCOLS       =   1
      FIXEDROWS       =   1
      FOCUSRECT       =   1
      EDITORBACKCOLOR =   -2147483643
      EDITORFORECOLOR =   -2147483640
      FORECOLOR       =   -2147483640
      FORECOLORFIXED  =   -2147483630
      FORECOLORSEL    =   -2147483634
      FORMATSTRING    =   ""
      Object.HEIGHT          =   1260
      GRIDCOLOR       =   12632256
      GRIDCOLORFIXED  =   0
      BeginProperty GRIDFONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GRIDLINES       =   1
      GRIDLINESFIXED  =   2
      GRIDLINEWIDTH   =   1
      MOUSEICON       =   "frmCreditApprovalNeo.frx":0000
      MOUSEPOINTER    =   0
      REDRAW          =   -1  'True
      RIGHTTOLEFT     =   0   'False
      ROWS            =   3
      SCROLLBARS      =   3
      SCROLLTRACK     =   0   'False
      SELECTIONMODE   =   0
      Object.TOOLTIPTEXT     =   ""
      WORDWRAP        =   0   'False
   End
   Begin xrGridEditor.GridEditor GridEditor2 
      Height          =   2640
      Left            =   1710
      TabIndex        =   101
      TabStop         =   0   'False
      Tag             =   "et0;eb0;et0;bc2"
      Top             =   4320
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   4657
      AllowBigSelection=   -1  'True
      AutoAdd         =   -1  'True
      AutoNumber      =   -1  'True
      BACKCOLOR       =   -2147483643
      BACKCOLORBKG    =   8421504
      BACKCOLORFIXED  =   -2147483633
      BACKCOLORSEL    =   -2147483635
      BORDERSTYLE     =   1
      COLS            =   2
      FILLSTYLE       =   0
      FIXEDCOLS       =   1
      FIXEDROWS       =   1
      FOCUSRECT       =   1
      EDITORBACKCOLOR =   -2147483643
      EDITORFORECOLOR =   -2147483640
      FORECOLOR       =   -2147483640
      FORECOLORFIXED  =   -2147483630
      FORECOLORSEL    =   -2147483634
      FORMATSTRING    =   ""
      Object.HEIGHT          =   2640
      GRIDCOLOR       =   12632256
      GRIDCOLORFIXED  =   0
      BeginProperty GRIDFONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GRIDLINES       =   1
      GRIDLINESFIXED  =   2
      GRIDLINEWIDTH   =   1
      MOUSEICON       =   "frmCreditApprovalNeo.frx":001C
      MOUSEPOINTER    =   0
      REDRAW          =   -1  'True
      RIGHTTOLEFT     =   0   'False
      ROWS            =   2
      SCROLLBARS      =   3
      SCROLLTRACK     =   0   'False
      SELECTIONMODE   =   0
      Object.TOOLTIPTEXT     =   ""
      WORDWRAP        =   0   'False
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   2655
      Index           =   0
      Left            =   1620
      Tag             =   "wt0;fb0"
      Top             =   975
      Width           =   10245
      _ExtentX        =   18071
      _ExtentY        =   4683
      Begin xrControl.xrButton cmdInfo 
         Height          =   300
         Index           =   0
         Left            =   4335
         TabIndex        =   8
         Top             =   450
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   529
         Caption         =   "..."
         AccessKey       =   "..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   86
         Left            =   1065
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1395
         Width           =   3660
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   80
         Left            =   1065
         MaxLength       =   50
         TabIndex        =   7
         Top             =   465
         Width           =   3240
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   15
         Left            =   1065
         MaxLength       =   50
         TabIndex        =   10
         Text            =   "X"
         Top             =   765
         Width           =   3660
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   81
         Left            =   5970
         MaxLength       =   50
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   465
         Width           =   1590
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   82
         Left            =   5970
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   765
         Width           =   2925
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   16
         Left            =   5970
         MaxLength       =   50
         TabIndex        =   24
         Text            =   "X"
         Top             =   1065
         Width           =   2925
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   83
         Left            =   5970
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1365
         Width           =   4170
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   17
         Left            =   5970
         MaxLength       =   50
         TabIndex        =   28
         Text            =   "X"
         Top             =   1965
         Width           =   4170
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   84
         Left            =   5970
         MaxLength       =   50
         TabIndex        =   30
         Text            =   "X"
         Top             =   2250
         Width           =   4170
      End
      Begin VB.ComboBox cmbField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         ItemData        =   "frmCreditApprovalNeo.frx":0038
         Left            =   1065
         List            =   "frmCreditApprovalNeo.frx":003A
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1065
         Width           =   3660
      End
      Begin VB.TextBox txtOthers 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   8085
         MaxLength       =   50
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   465
         Width           =   810
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   1245
         MaxLength       =   50
         TabIndex        =   5
         Top             =   75
         Width           =   1710
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   25
         Left            =   1065
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   16
         Text            =   "frmCreditApprovalNeo.frx":003C
         Top             =   1695
         Width           =   3660
      End
      Begin xrControl.xrFrame xrFrame2 
         Height          =   1185
         Left            =   8955
         Top             =   60
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   2090
         Begin VB.Image imgField 
            Height          =   1095
            Left            =   30
            Picture         =   "frmCreditApprovalNeo.frx":0040
            Stretch         =   -1  'True
            Top             =   30
            Width           =   1095
         End
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contact No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   14
         Left            =   60
         TabIndex        =   13
         Top             =   1410
         Width           =   810
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   13
         Left            =   60
         TabIndex        =   6
         Top             =   510
         Width           =   420
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nickname"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   16
         Left            =   60
         TabIndex        =   9
         Top             =   810
         Width           =   720
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Birth Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   17
         Left            =   4875
         TabIndex        =   17
         Top             =   525
         Width           =   705
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Birth Place"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   18
         Left            =   4875
         TabIndex        =   21
         Top             =   810
         Width           =   765
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Length of Stay"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   19
         Left            =   4875
         TabIndex        =   23
         Top             =   1110
         Width           =   1035
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   20
         Left            =   4875
         TabIndex        =   25
         Top             =   1410
         Width           =   570
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Former Addrs."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   22
         Left            =   4875
         TabIndex        =   27
         Top             =   2010
         Width           =   975
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Town/City"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   23
         Left            =   4875
         TabIndex        =   29
         Top             =   2310
         Width           =   735
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Civil Status"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   24
         Left            =   60
         TabIndex        =   11
         Top             =   1125
         Width           =   780
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Age"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   81
         Left            =   7710
         TabIndex        =   19
         Top             =   525
         Width           =   285
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   285
         Left            =   1320
         Tag             =   "et0;ht2"
         Top             =   150
         Width           =   1710
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   4
         Top             =   120
         Width           =   1140
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   80
         Left            =   60
         TabIndex        =   15
         Top             =   1740
         Width           =   630
      End
      Begin VB.Shape Shape4 
         Height          =   330
         Index           =   0
         Left            =   6000
         Top             =   90
         Width           =   2865
      End
      Begin VB.Shape Shape3 
         Height          =   390
         Index           =   0
         Left            =   5970
         Top             =   60
         Width           =   2925
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
         Left            =   6000
         TabIndex        =   215
         Tag             =   "eb0;et0"
         Top             =   135
         Width           =   2850
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'Transparent
         FillStyle       =   0  'Solid
         Height          =   285
         Index           =   0
         Left            =   6030
         Tag             =   "et0;et0"
         Top             =   120
         Width           =   2820
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   450
      Index           =   1
      Left            =   1620
      Tag             =   "wt0;fb0"
      Top             =   510
      Width           =   10245
      _ExtentX        =   18071
      _ExtentY        =   794
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   1245
         TabIndex        =   1
         Top             =   60
         Width           =   1710
      End
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   4695
         TabIndex        =   3
         Top             =   60
         Width           =   5460
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Full Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   9
         Left            =   3825
         TabIndex        =   2
         Top             =   105
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   19
         Left            =   60
         TabIndex        =   0
         Top             =   105
         Width           =   1155
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3450
      Left            =   1650
      TabIndex        =   31
      Tag             =   "et0;wb0"
      Top             =   3600
      Width           =   10230
      _ExtentX        =   18045
      _ExtentY        =   6085
      _Version        =   393216
      Tabs            =   8
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "App&lication Information"
      TabPicture(0)   =   "frmCreditApprovalNeo.frx":4CD4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "xrFrame3(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Spouse/Co-Maker"
      TabPicture(1)   =   "frmCreditApprovalNeo.frx":4CF0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "xrFrame3(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&Personal References"
      TabPicture(2)   =   "frmCreditApprovalNeo.frx":4D0C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "&Employment Data (applicant)"
      TabPicture(3)   =   "frmCreditApprovalNeo.frx":4D28
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "xrFrame3(2)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Employment Data (sp&ouse)"
      TabPicture(4)   =   "frmCreditApprovalNeo.frx":4D44
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "xrFrame3(3)"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "&Monthly Income/Expenses"
      TabPicture(5)   =   "frmCreditApprovalNeo.frx":4D60
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "xrFrame3(4)"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "&Financer Information"
      TabPicture(6)   =   "frmCreditApprovalNeo.frx":4D7C
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "xrFrame3(5)"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "&Dependents"
      TabPicture(7)   =   "frmCreditApprovalNeo.frx":4D98
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "xrFrame3(6)"
      Tab(7).ControlCount=   1
      Begin xrControl.xrFrame xrFrame3 
         Height          =   2715
         Index           =   5
         Left            =   -74955
         Tag             =   "wt0;fb0"
         Top             =   645
         Width           =   10005
         _ExtentX        =   17648
         _ExtentY        =   4789
         Begin VB.ComboBox cmbRelation 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            ItemData        =   "frmCreditApprovalNeo.frx":4DB4
            Left            =   1200
            List            =   "frmCreditApprovalNeo.frx":4DB6
            Style           =   2  'Dropdown List
            TabIndex        =   204
            Top             =   1320
            Width           =   3120
         End
         Begin VB.TextBox txtWaysMn 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   91
            Left            =   1215
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   202
            TabStop         =   0   'False
            Top             =   1020
            Width           =   3120
         End
         Begin VB.TextBox txtWaysMn 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   87
            Left            =   1215
            Locked          =   -1  'True
            MaxLength       =   50
            MultiLine       =   -1  'True
            TabIndex        =   200
            TabStop         =   0   'False
            Top             =   435
            Width           =   3120
         End
         Begin VB.TextBox txtWaysMn 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   86
            Left            =   1215
            MaxLength       =   50
            TabIndex        =   198
            Top             =   120
            Width           =   2670
         End
         Begin VB.TextBox txtWaysMn 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Index           =   44
            Left            =   6660
            MaxLength       =   45
            MultiLine       =   -1  'True
            TabIndex        =   210
            Top             =   405
            Width           =   3120
         End
         Begin VB.TextBox txtWaysMn 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   90
            Left            =   1215
            MaxLength       =   45
            TabIndex        =   206
            Top             =   1860
            Width           =   3120
         End
         Begin VB.TextBox txtWaysMn 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   88
            Left            =   6660
            MaxLength       =   45
            TabIndex        =   208
            Top             =   105
            Width           =   3120
         End
         Begin VB.TextBox txtWaysMn 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   46
            Left            =   6660
            MaxLength       =   45
            TabIndex        =   212
            Top             =   1095
            Width           =   3120
         End
         Begin VB.TextBox txtWaysMn 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   47
            Left            =   6660
            MaxLength       =   45
            TabIndex        =   214
            Top             =   1395
            Width           =   3120
         End
         Begin xrControl.xrButton cmdInfo 
            Height          =   300
            Index           =   3
            Left            =   3930
            TabIndex        =   219
            Top             =   120
            Width           =   390
            _ExtentX        =   688
            _ExtentY        =   529
            Caption         =   "..."
            AccessKey       =   "..."
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
            Caption         =   "Contact No."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   26
            Left            =   180
            TabIndex        =   201
            Top             =   1050
            Width           =   855
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   5
            Left            =   165
            TabIndex        =   199
            Top             =   480
            Width           =   570
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Relation"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   99
            Left            =   180
            TabIndex        =   203
            Top             =   1410
            Width           =   585
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   98
            Left            =   165
            TabIndex        =   197
            Top             =   180
            Width           =   420
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Function "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   106
            Left            =   5490
            TabIndex        =   209
            Top             =   465
            Width           =   660
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Country"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   103
            Left            =   180
            TabIndex        =   205
            Top             =   1920
            Width           =   540
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Occupation"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   100
            Left            =   5475
            TabIndex        =   207
            Top             =   165
            Width           =   825
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Income"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   101
            Left            =   5460
            TabIndex        =   211
            Top             =   1155
            Width           =   525
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Support"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   102
            Left            =   5445
            TabIndex        =   213
            Top             =   1440
            Width           =   810
         End
      End
      Begin xrControl.xrFrame xrFrame3 
         Height          =   2715
         Index           =   4
         Left            =   -74955
         Tag             =   "wt0;fb0"
         Top             =   645
         Width           =   10005
         _ExtentX        =   17648
         _ExtentY        =   4789
         Begin VB.TextBox txtWaysMn 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   16
            Left            =   1260
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   173
            TabStop         =   0   'False
            Text            =   "0.00"
            Top             =   180
            Width           =   2130
         End
         Begin VB.TextBox txtWaysMn 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   17
            Left            =   1260
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   175
            TabStop         =   0   'False
            Text            =   "0.00"
            Top             =   480
            Width           =   2130
         End
         Begin VB.TextBox txtWaysMn 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   18
            Left            =   5010
            MaxLength       =   50
            TabIndex        =   180
            Text            =   "0.00"
            Top             =   180
            Width           =   2130
         End
         Begin VB.TextBox txtWaysMn 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   19
            Left            =   5010
            MaxLength       =   50
            TabIndex        =   182
            Text            =   "0.00"
            Top             =   480
            Width           =   2130
         End
         Begin VB.TextBox txtWaysMn 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   20
            Left            =   5010
            MaxLength       =   50
            TabIndex        =   184
            Text            =   "0.00"
            Top             =   780
            Width           =   2130
         End
         Begin VB.TextBox txtWaysMn 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   21
            Left            =   5010
            MaxLength       =   50
            TabIndex        =   186
            Text            =   "0.00"
            Top             =   1080
            Width           =   2130
         End
         Begin VB.TextBox txtWaysMn 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   22
            Left            =   5010
            MaxLength       =   50
            TabIndex        =   188
            Text            =   "0.00"
            Top             =   1380
            Width           =   2130
         End
         Begin VB.TextBox txtWaysMn 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   23
            Left            =   5010
            MaxLength       =   50
            TabIndex        =   190
            Text            =   "0.00"
            Top             =   1680
            Width           =   2130
         End
         Begin VB.TextBox txtWaysMn 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   24
            Left            =   5010
            MaxLength       =   50
            TabIndex        =   192
            Text            =   "0.00"
            Top             =   1980
            Width           =   2130
         End
         Begin VB.TextBox txtOthers 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   3
            Left            =   5010
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   194
            TabStop         =   0   'False
            Text            =   "0.00"
            Top             =   2280
            Width           =   2130
         End
         Begin VB.TextBox txtOthers 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   2
            Left            =   1260
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   177
            TabStop         =   0   'False
            Text            =   "0.00"
            Top             =   2280
            Width           =   2130
         End
         Begin VB.TextBox txtOthers 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Impact"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Index           =   4
            Left            =   7695
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   196
            TabStop         =   0   'False
            Text            =   "0.00"
            Top             =   2175
            Width           =   2130
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Gross Income"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   64
            Left            =   105
            TabIndex        =   172
            Top             =   225
            Width           =   975
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Other Income"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   65
            Left            =   105
            TabIndex        =   174
            Top             =   525
            Width           =   960
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cost of Living"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   66
            Left            =   3855
            TabIndex        =   179
            Top             =   225
            Width           =   960
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Education"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   67
            Left            =   3855
            TabIndex        =   181
            Top             =   525
            Width           =   720
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Transportation"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   68
            Left            =   3855
            TabIndex        =   183
            Top             =   825
            Width           =   1020
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rent Expense"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   69
            Left            =   3855
            TabIndex        =   185
            Top             =   1125
            Width           =   1005
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Utilities"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   70
            Left            =   3855
            TabIndex        =   187
            Top             =   1425
            Width           =   495
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Monthly Amort."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   71
            Left            =   3855
            TabIndex        =   189
            Top             =   1725
            Width           =   1050
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Other Expense"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   72
            Left            =   3855
            TabIndex        =   191
            Top             =   2025
            Width           =   1050
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Expenses"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   84
            Left            =   3855
            TabIndex        =   193
            Top             =   2340
            Width           =   1095
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Income"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   83
            Left            =   105
            TabIndex        =   176
            Top             =   2340
            Width           =   930
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Net Income"
            Height          =   195
            Index           =   85
            Left            =   7695
            TabIndex        =   195
            Top             =   1995
            Width           =   990
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Expenses"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   87
            Left            =   5010
            TabIndex        =   178
            Top             =   0
            Width           =   735
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Income"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   94
            Left            =   1260
            TabIndex        =   171
            Top             =   0
            Width           =   555
         End
      End
      Begin xrControl.xrFrame xrFrame3 
         Height          =   2715
         Index           =   3
         Left            =   -74955
         Tag             =   "wt0;fb0"
         Top             =   645
         Width           =   10005
         _ExtentX        =   17648
         _ExtentY        =   4789
         Begin VB.TextBox txtWaysMn 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   25
            Left            =   1725
            MaxLength       =   50
            TabIndex        =   139
            Top             =   240
            Width           =   3120
         End
         Begin VB.TextBox txtWaysMn 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   26
            Left            =   1725
            MaxLength       =   50
            TabIndex        =   141
            Top             =   540
            Width           =   3120
         End
         Begin VB.TextBox txtWaysMn 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   83
            Left            =   1725
            MaxLength       =   50
            TabIndex        =   143
            Top             =   840
            Width           =   3120
         End
         Begin VB.TextBox txtWaysMn 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   28
            Left            =   1725
            MaxLength       =   50
            TabIndex        =   145
            Top             =   1140
            Width           =   1455
         End
         Begin VB.TextBox txtWaysMn 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   29
            Left            =   4125
            MaxLength       =   50
            TabIndex        =   147
            Top             =   1140
            Width           =   720
         End
         Begin VB.TextBox txtWaysMn 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   85
            Left            =   1725
            MaxLength       =   50
            TabIndex        =   149
            Top             =   1440
            Width           =   3120
         End
         Begin VB.TextBox txtWaysMn 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   32
            Left            =   6690
            MaxLength       =   50
            TabIndex        =   153
            Text            =   "0.00"
            Top             =   240
            Width           =   1005
         End
         Begin VB.TextBox txtWaysMn 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   33
            Left            =   8835
            MaxLength       =   50
            TabIndex        =   155
            Text            =   "0.00"
            Top             =   225
            Width           =   1005
         End
         Begin VB.TextBox txtWaysMn 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   35
            Left            =   6690
            MaxLength       =   50
            TabIndex        =   160
            Top             =   1170
            Width           =   3165
         End
         Begin VB.TextBox txtWaysMn 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   36
            Left            =   6690
            MaxLength       =   50
            TabIndex        =   162
            Top             =   1470
            Width           =   3165
         End
         Begin VB.TextBox txtWaysMn 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   84
            Left            =   6690
            MaxLength       =   50
            TabIndex        =   164
            Top             =   1770
            Width           =   3165
         End
         Begin VB.TextBox txtWaysMn 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   38
            Left            =   6690
            MaxLength       =   50
            TabIndex        =   166
            Top             =   2070
            Width           =   1455
         End
         Begin VB.TextBox txtWaysMn 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   39
            Left            =   9135
            MaxLength       =   50
            TabIndex        =   168
            Top             =   2070
            Width           =   720
         End
         Begin VB.TextBox txtWaysMn 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   40
            Left            =   6690
            MaxLength       =   50
            TabIndex        =   170
            Text            =   "0.00"
            Top             =   2385
            Width           =   1005
         End
         Begin VB.ComboBox cmbField 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   4
            Left            =   6690
            Style           =   2  'Dropdown List
            TabIndex        =   157
            Top             =   540
            Width           =   3165
         End
         Begin VB.TextBox txtWaysMn 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   630
            Index           =   31
            Left            =   1725
            MaxLength       =   50
            TabIndex        =   151
            Top             =   1740
            Width           =   3120
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Employer/Co. Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   49
            Left            =   135
            TabIndex        =   138
            Top             =   285
            Width           =   1425
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   50
            Left            =   135
            TabIndex        =   140
            Top             =   585
            Width           =   570
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Town/City"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   51
            Left            =   135
            TabIndex        =   142
            Top             =   885
            Width           =   735
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Telephone No."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   52
            Left            =   135
            TabIndex        =   144
            Top             =   1185
            Width           =   1065
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Service Len"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   53
            Left            =   3225
            TabIndex        =   146
            Top             =   1215
            Width           =   855
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Occupation"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   54
            Left            =   135
            TabIndex        =   148
            Top             =   1485
            Width           =   825
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Estimated Mon. Salary"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   55
            Left            =   5010
            TabIndex        =   152
            Top             =   285
            Width           =   1575
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Other Income"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   56
            Left            =   7785
            TabIndex        =   154
            Top             =   285
            Width           =   960
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Employment Status"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   57
            Left            =   5010
            TabIndex        =   156
            Top             =   585
            Width           =   1350
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nature of Business"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   58
            Left            =   5040
            TabIndex        =   159
            Top             =   1215
            Width           =   1335
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   59
            Left            =   5040
            TabIndex        =   161
            Top             =   1515
            Width           =   570
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Town/City"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   60
            Left            =   5040
            TabIndex        =   163
            Top             =   1815
            Width           =   735
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Telephone No."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   61
            Left            =   5040
            TabIndex        =   165
            Top             =   2115
            Width           =   1065
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Yr in Busnes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   62
            Left            =   8205
            TabIndex        =   167
            Top             =   2115
            Width           =   885
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Estmted Mon. Income"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   63
            Left            =   5070
            TabIndex        =   169
            Top             =   2430
            Width           =   1545
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "If Employed"
            Height          =   195
            Index           =   88
            Left            =   1785
            TabIndex        =   137
            Top             =   15
            Width           =   1005
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "If Self-Employed"
            Height          =   195
            Index           =   89
            Left            =   6690
            TabIndex        =   158
            Top             =   945
            Width           =   1395
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Function"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   28
            Left            =   135
            TabIndex        =   150
            Top             =   1785
            Width           =   615
         End
      End
      Begin xrControl.xrFrame xrFrame3 
         Height          =   2715
         Index           =   2
         Left            =   -74955
         Tag             =   "wt0;fb0"
         Top             =   645
         Width           =   10005
         _ExtentX        =   17648
         _ExtentY        =   4789
         Begin VB.TextBox txtWaysMn 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   1590
            MaxLength       =   50
            TabIndex        =   105
            Top             =   255
            Width           =   3195
         End
         Begin VB.TextBox txtWaysMn 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   1590
            MaxLength       =   50
            TabIndex        =   107
            Top             =   555
            Width           =   3195
         End
         Begin VB.TextBox txtWaysMn 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   80
            Left            =   1590
            MaxLength       =   50
            TabIndex        =   109
            Top             =   855
            Width           =   3195
         End
         Begin VB.TextBox txtWaysMn 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   1590
            MaxLength       =   50
            TabIndex        =   111
            Top             =   1155
            Width           =   1455
         End
         Begin VB.TextBox txtWaysMn 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   4
            Left            =   4065
            MaxLength       =   50
            TabIndex        =   113
            Top             =   1155
            Width           =   720
         End
         Begin VB.TextBox txtWaysMn 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   82
            Left            =   1590
            MaxLength       =   50
            TabIndex        =   115
            Top             =   1455
            Width           =   3195
         End
         Begin VB.TextBox txtWaysMn 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   7
            Left            =   6750
            MaxLength       =   50
            TabIndex        =   119
            Text            =   "0.00"
            Top             =   210
            Width           =   1005
         End
         Begin VB.TextBox txtWaysMn 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   8
            Left            =   8880
            MaxLength       =   50
            TabIndex        =   121
            Text            =   "0.00"
            Top             =   210
            Width           =   1005
         End
         Begin VB.TextBox txtWaysMn 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   10
            Left            =   6780
            MaxLength       =   50
            TabIndex        =   126
            Top             =   1170
            Width           =   3105
         End
         Begin VB.TextBox txtWaysMn 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   11
            Left            =   6780
            MaxLength       =   50
            TabIndex        =   128
            Top             =   1470
            Width           =   3105
         End
         Begin VB.TextBox txtWaysMn 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   81
            Left            =   6780
            MaxLength       =   50
            TabIndex        =   130
            Top             =   1770
            Width           =   3105
         End
         Begin VB.TextBox txtWaysMn 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   13
            Left            =   6780
            MaxLength       =   50
            TabIndex        =   132
            Top             =   2070
            Width           =   1455
         End
         Begin VB.TextBox txtWaysMn 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   14
            Left            =   9165
            MaxLength       =   50
            TabIndex        =   134
            Top             =   2070
            Width           =   720
         End
         Begin VB.TextBox txtWaysMn 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   15
            Left            =   6780
            MaxLength       =   50
            TabIndex        =   136
            Text            =   "0.00"
            Top             =   2370
            Width           =   1005
         End
         Begin VB.ComboBox cmbField 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   6750
            Style           =   2  'Dropdown List
            TabIndex        =   123
            Top             =   510
            Width           =   3135
         End
         Begin VB.TextBox txtWaysMn 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   630
            Index           =   6
            Left            =   1590
            MaxLength       =   50
            MultiLine       =   -1  'True
            TabIndex        =   117
            Top             =   1755
            Width           =   3195
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Employer/Co. Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   34
            Left            =   135
            TabIndex        =   104
            Top             =   300
            Width           =   1425
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   35
            Left            =   135
            TabIndex        =   106
            Top             =   600
            Width           =   570
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Town/City"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   36
            Left            =   135
            TabIndex        =   108
            Top             =   900
            Width           =   735
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Telephone No."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   37
            Left            =   135
            TabIndex        =   110
            Top             =   1200
            Width           =   1065
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Service Len"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   38
            Left            =   3120
            TabIndex        =   112
            Top             =   1200
            Width           =   855
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Occupation"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   39
            Left            =   135
            TabIndex        =   114
            Top             =   1500
            Width           =   825
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Estimated Mon. Salary"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   40
            Left            =   5055
            TabIndex        =   118
            Top             =   255
            Width           =   1575
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Other Income"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   41
            Left            =   7875
            TabIndex        =   120
            Top             =   255
            Width           =   960
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Employment Status"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   42
            Left            =   60
            TabIndex        =   216
            Top             =   2775
            Width           =   1350
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nature of Business"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   43
            Left            =   5040
            TabIndex        =   125
            Top             =   1215
            Width           =   1335
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   44
            Left            =   5040
            TabIndex        =   127
            Top             =   1515
            Width           =   570
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Town/City"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   45
            Left            =   5040
            TabIndex        =   129
            Top             =   1815
            Width           =   735
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Telephone No."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   46
            Left            =   5040
            TabIndex        =   131
            Top             =   2115
            Width           =   1065
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Yr in Busnes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   47
            Left            =   8265
            TabIndex        =   133
            Top             =   2115
            Width           =   885
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Estimated Mon. Income"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   48
            Left            =   5040
            TabIndex        =   135
            Top             =   2415
            Width           =   1665
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "If Employed"
            Height          =   195
            Index           =   90
            Left            =   1785
            TabIndex        =   103
            Top             =   30
            Width           =   1005
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "If Self-Employed"
            Height          =   195
            Index           =   91
            Left            =   6735
            TabIndex        =   124
            Top             =   915
            Width           =   1395
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Employment Status"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   86
            Left            =   5055
            TabIndex        =   122
            Top             =   540
            Width           =   1350
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Function"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   92
            Left            =   135
            TabIndex        =   116
            Top             =   1800
            Width           =   615
         End
      End
      Begin xrControl.xrFrame xrFrame3 
         Height          =   2685
         Index           =   1
         Left            =   -74970
         Tag             =   "wt0;fb0"
         Top             =   645
         Width           =   10005
         _ExtentX        =   17648
         _ExtentY        =   4736
         Begin VB.ComboBox cmbRelation 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            ItemData        =   "frmCreditApprovalNeo.frx":4DB8
            Left            =   6150
            List            =   "frmCreditApprovalNeo.frx":4DBA
            Style           =   2  'Dropdown List
            TabIndex        =   83
            TabStop         =   0   'False
            Top             =   1410
            Width           =   3660
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            Index           =   90
            Left            =   1410
            Locked          =   -1  'True
            MaxLength       =   50
            MultiLine       =   -1  'True
            TabIndex        =   60
            TabStop         =   0   'False
            Top             =   525
            Width           =   3480
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   100
            Left            =   1410
            MaxLength       =   50
            TabIndex        =   74
            Text            =   "X"
            Top             =   2325
            Width           =   3480
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   21
            Left            =   1410
            MaxLength       =   50
            TabIndex        =   72
            Text            =   "X"
            Top             =   2025
            Width           =   3480
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   94
            Left            =   6165
            Locked          =   -1  'True
            MaxLength       =   50
            MultiLine       =   -1  'True
            TabIndex        =   79
            Top             =   510
            Width           =   3660
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   92
            Left            =   6165
            MaxLength       =   50
            TabIndex        =   76
            Top             =   210
            Width           =   3210
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   96
            Left            =   6165
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   81
            Top             =   1110
            Width           =   3660
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   91
            Left            =   3570
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   64
            Top             =   1110
            Width           =   1320
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   35
            Left            =   1410
            MaxLength       =   50
            TabIndex        =   62
            Text            =   "X"
            Top             =   1110
            Width           =   1350
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   89
            Left            =   1410
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   70
            Top             =   1725
            Width           =   3480
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   88
            Left            =   1410
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   66
            Top             =   1410
            Width           =   1350
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   87
            Left            =   1410
            MaxLength       =   50
            TabIndex        =   57
            Top             =   210
            Width           =   3030
         End
         Begin VB.TextBox txtOthers 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   3570
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   68
            Top             =   1410
            Width           =   1320
         End
         Begin xrControl.xrButton cmdInfo 
            Height          =   300
            Index           =   1
            Left            =   4485
            TabIndex        =   58
            Top             =   210
            Width           =   390
            _ExtentX        =   688
            _ExtentY        =   529
            Caption         =   "..."
            AccessKey       =   "..."
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin xrControl.xrButton cmdInfo 
            Height          =   300
            Index           =   2
            Left            =   9435
            TabIndex        =   77
            Top             =   180
            Width           =   390
            _ExtentX        =   688
            _ExtentY        =   529
            Caption         =   "..."
            AccessKey       =   "..."
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
            Caption         =   "Town/City"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   74
            Left            =   375
            TabIndex        =   73
            Top             =   2370
            Width           =   735
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Former Addrs."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   33
            Left            =   375
            TabIndex        =   71
            Top             =   2070
            Width           =   975
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            Caption         =   "Spouse Info"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   15
            Left            =   420
            TabIndex        =   217
            Tag             =   "wt0;fb0"
            Top             =   0
            Width           =   1020
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            Caption         =   "Co-Maker Info"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   104
            Left            =   5355
            TabIndex        =   218
            Tag             =   "wt0;fb0"
            Top             =   -15
            Width           =   1200
         End
         Begin VB.Shape Shape2 
            Height          =   2535
            Index           =   1
            Left            =   5070
            Top             =   90
            Width           =   4845
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   76
            Left            =   5160
            TabIndex        =   78
            Top             =   555
            Width           =   570
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   73
            Left            =   5160
            TabIndex        =   75
            Top             =   255
            Width           =   420
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Relation"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   78
            Left            =   5160
            TabIndex        =   82
            Top             =   1455
            Width           =   585
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contact No."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   21
            Left            =   5175
            TabIndex        =   80
            Top             =   1110
            Width           =   930
         End
         Begin VB.Shape Shape2 
            Height          =   2535
            Index           =   2
            Left            =   150
            Top             =   105
            Width           =   4845
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tel No."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   27
            Left            =   2880
            TabIndex        =   63
            Top             =   1125
            Width           =   525
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   97
            Left            =   375
            TabIndex        =   59
            Top             =   570
            Width           =   570
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Length Stay"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   32
            Left            =   375
            TabIndex        =   61
            Top             =   1110
            Width           =   855
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Birth Place"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   31
            Left            =   375
            TabIndex        =   69
            Top             =   1740
            Width           =   765
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Birth Date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   30
            Left            =   375
            TabIndex        =   65
            Top             =   1425
            Width           =   705
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   25
            Left            =   375
            TabIndex        =   56
            Top             =   240
            Width           =   420
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Age"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   82
            Left            =   3120
            TabIndex        =   67
            Top             =   1470
            Width           =   285
         End
      End
      Begin xrControl.xrFrame xrFrame3 
         Height          =   2715
         Index           =   0
         Left            =   45
         Tag             =   "wt0;fb0"
         Top             =   645
         Width           =   10005
         _ExtentX        =   17648
         _ExtentY        =   4789
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   1560
            MaxLength       =   50
            TabIndex        =   33
            Top             =   105
            Width           =   2805
         End
         Begin VB.ComboBox cmbField 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   645
            Width           =   2805
         End
         Begin VB.ComboBox cmbField 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   960
            Width           =   2805
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   7
            Left            =   1560
            MaxLength       =   50
            TabIndex        =   39
            Top             =   1290
            Width           =   2805
         End
         Begin VB.TextBox txtField 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   8
            Left            =   1560
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   41
            TabStop         =   0   'False
            Text            =   "0.00"
            Top             =   1590
            Width           =   2805
         End
         Begin VB.TextBox txtField 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   9
            Left            =   1560
            MaxLength       =   50
            TabIndex        =   43
            Text            =   "0"
            Top             =   1890
            Width           =   2805
         End
         Begin VB.TextBox txtField 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   10
            Left            =   1560
            MaxLength       =   50
            TabIndex        =   45
            Text            =   "0.00"
            Top             =   2190
            Width           =   2805
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   97
            Left            =   6870
            MaxLength       =   50
            TabIndex        =   47
            Top             =   645
            Width           =   2805
         End
         Begin VB.TextBox txtField 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   12
            Left            =   6870
            MaxLength       =   50
            TabIndex        =   51
            TabStop         =   0   'False
            Text            =   "0.00"
            Top             =   1245
            Width           =   2805
         End
         Begin VB.TextBox txtField 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   13
            Left            =   6870
            MaxLength       =   50
            TabIndex        =   49
            Top             =   945
            Width           =   1560
         End
         Begin VB.TextBox txtField 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   14
            Left            =   6870
            MaxLength       =   50
            TabIndex        =   55
            Top             =   1845
            Width           =   2805
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   99
            Left            =   6870
            TabIndex        =   53
            Text            =   "X"
            Top             =   1545
            Width           =   2805
         End
         Begin xrControl.xrButton xrButton1 
            Height          =   285
            Left            =   8475
            TabIndex        =   220
            TabStop         =   0   'False
            Top             =   945
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   503
            Caption         =   "A&uto-Compute"
            AccessKey       =   "u"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   6
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
            Caption         =   "Specify"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   540
            TabIndex        =   38
            Top             =   1320
            Width           =   525
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Application Date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   255
            TabIndex        =   32
            Top             =   120
            Width           =   1170
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Application Type"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   255
            TabIndex        =   34
            Top             =   705
            Width           =   1185
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Unit Applied"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   255
            TabIndex        =   36
            Top             =   1035
            Width           =   855
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Gross Price"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   6
            Left            =   285
            TabIndex        =   40
            Top             =   1635
            Width           =   810
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PN Value"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   7
            Left            =   285
            TabIndex        =   42
            Top             =   1935
            Width           =   675
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Down Payment"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   8
            Left            =   285
            TabIndex        =   44
            Top             =   2235
            Width           =   1080
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Model"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   9
            Left            =   5475
            TabIndex        =   46
            Top             =   645
            Width           =   435
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Monthly Amort."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   10
            Left            =   5475
            TabIndex        =   50
            Top             =   1245
            Width           =   1050
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Term"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   11
            Left            =   5475
            TabIndex        =   48
            Top             =   945
            Width           =   360
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quick Match No."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   12
            Left            =   5475
            TabIndex        =   54
            Top             =   1845
            Width           =   1215
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Credit Investigator"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   95
            Left            =   5475
            TabIndex        =   52
            Top             =   1545
            Width           =   1275
         End
      End
      Begin xrControl.xrFrame xrFrame3 
         Height          =   2715
         Index           =   6
         Left            =   -74970
         Tag             =   "wt0;fb0"
         Top             =   645
         Width           =   10005
         _ExtentX        =   17648
         _ExtentY        =   4789
         Begin VB.ComboBox cmbRelation 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            ItemData        =   "frmCreditApprovalNeo.frx":4DBC
            Left            =   1140
            List            =   "frmCreditApprovalNeo.frx":4DBE
            Style           =   2  'Dropdown List
            TabIndex        =   88
            TabStop         =   0   'False
            Top             =   405
            Width           =   3240
         End
         Begin VB.TextBox txtDependnt 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   4
            Left            =   5475
            MaxLength       =   30
            MultiLine       =   -1  'True
            TabIndex        =   98
            Text            =   "frmCreditApprovalNeo.frx":4DC0
            Top             =   735
            Width           =   3240
         End
         Begin VB.TextBox txtDependnt 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   5475
            MaxLength       =   50
            TabIndex        =   96
            Text            =   "X"
            Top             =   435
            Width           =   3240
         End
         Begin VB.TextBox txtDependnt 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   5475
            MaxLength       =   50
            TabIndex        =   94
            Text            =   "X"
            Top             =   135
            Width           =   3240
         End
         Begin VB.TextBox txtDependnt 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   3345
            TabIndex        =   92
            Text            =   "X"
            Top             =   735
            Width           =   1035
         End
         Begin VB.TextBox txtDependnt 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   82
            Left            =   1140
            Locked          =   -1  'True
            TabIndex        =   90
            TabStop         =   0   'False
            Top             =   735
            Width           =   1590
         End
         Begin VB.TextBox txtDependnt 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   80
            Left            =   1140
            TabIndex        =   85
            Top             =   105
            Width           =   2805
         End
         Begin xrControl.xrButton cmdDependent 
            Height          =   390
            Index           =   0
            Left            =   9090
            TabIndex        =   99
            Top             =   555
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   688
            Caption         =   "&Add"
            AccessKey       =   "A"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin xrControl.xrButton cmdDependent 
            Height          =   390
            Index           =   1
            Left            =   9090
            TabIndex        =   100
            Top             =   975
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   688
            Caption         =   "&Delete"
            AccessKey       =   "D"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin xrControl.xrButton cmdInfo 
            Height          =   300
            Index           =   4
            Left            =   3975
            TabIndex        =   86
            Top             =   105
            Width           =   390
            _ExtentX        =   688
            _ExtentY        =   529
            Caption         =   "..."
            AccessKey       =   "..."
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
            Caption         =   "Remarks"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   108
            Left            =   4470
            TabIndex        =   97
            Top             =   780
            Width           =   630
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Company"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   107
            Left            =   4470
            TabIndex        =   95
            Top             =   495
            Width           =   660
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "School"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   105
            Left            =   4470
            TabIndex        =   93
            Top             =   195
            Width           =   495
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Age"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   93
            Left            =   2895
            TabIndex        =   91
            Top             =   855
            Width           =   285
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Birth Date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   79
            Left            =   135
            TabIndex        =   89
            Top             =   795
            Width           =   705
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Relation"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   77
            Left            =   135
            TabIndex        =   87
            Top             =   465
            Width           =   585
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   75
            Left            =   135
            TabIndex        =   84
            Top             =   180
            Width           =   420
         End
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   90
      TabIndex        =   221
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
      Picture         =   "frmCreditApprovalNeo.frx":4DC4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   90
      TabIndex        =   222
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
      Picture         =   "frmCreditApprovalNeo.frx":553E
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   90
      TabIndex        =   223
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
      Picture         =   "frmCreditApprovalNeo.frx":5CB8
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   90
      TabIndex        =   224
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
      Picture         =   "frmCreditApprovalNeo.frx":6432
   End
End
Attribute VB_Name = "frmMPCreditApprovalNeo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmCreditAppRegNeo"

Private oSkin As clsFormSkin
Private WithEvents oTrans As ggcLRApplication.clsLRApplication
Attribute oTrans.VB_VarHelpID = -1

Dim psObjectNme As String

Dim pnCtr As Integer
Dim pnIndex As Integer
Dim panCmdHwnd(3) As Long

Dim pnRow As Integer

Private Sub cmbRelation_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
   Case 0
      oTrans.Master("sRelation") = CStr(cmbRelation(0).ListIndex)
   Case 1
      oTrans.Dependent("sRelatnID") = CStr(cmbRelation(1).ListIndex)
   Case 2
      oTrans.WaysMeans("sReltnCD2") = CStr(cmbRelation(2).ListIndex)
   End Select
   Debug.Print "Relation to Dependent: " & oTrans.Dependent("sRelatnID")
End Sub

Private Sub cmdDependent_Click(Index As Integer)
   Dim lnCtr As Integer
   Dim lotxt As TextBox

   Select Case Index
   Case 0 'Add row
      If oTrans.AddDependent Then
         'Add a row to the grid
         GridEditor1.Rows = GridEditor1.Rows + 1

         'Save the last row to pnrow
         pnRow = GridEditor1.Rows - 2

         'Load to the textboxes
         For Each lotxt In txtDependnt
            lotxt.Text = IFNull(oTrans.Dependent(lotxt.Index))
         Next
         cmbRelation(1).ListIndex = oTrans.Dependent("sRelatnID")

         txtDependnt(80).SetFocus

      End If
   Case 1
      'Check if we have a record to delete
      If pnRow < 0 Then
         MsgBox "Please load a record to delete!", vbInformation, "Warning"
         Exit Sub
      End If

      If oTrans.DeleteDependent() Then
         'Remove data from textboxes
         For Each lotxt In txtDependnt
            lotxt.Text = ""
         Next
         cmbRelation(1).ListIndex = 0

         With GridEditor1
            'Delete grid accordingly
            If .Rows = 2 Then
               For lnCtr = 1 To .Cols - 1
                  .TextMatrix(pnRow + 1, lnCtr) = ""
               Next
            Else
               .Row = pnRow + 1
               .deleteRow
            End If

            'Position the record pointer
            pnRow = pnRow + 1
            If pnRow > .Rows - 2 Then
               pnRow = .Rows - 2
            End If

            'load the depentdent
            Call oTrans.LoadDependent(pnRow)

            'Load to the textboxes
            For Each lotxt In txtDependnt
               lotxt.Text = IFNull(oTrans.Dependent(lotxt.Index))
            Next
            cmbRelation(1).ListIndex = IFNull(oTrans.Dependent("sRelatnID"), 0)
         End With
      End If
   End Select
End Sub

Private Sub cmdInfo_Click(Index As Integer)
   Select Case Index
   Case 0   'Applicant
      If txtField(80).Locked = False Then
         Call oTrans.SearchMaster(80, txtField(80).Text)
      Else
         Call oTrans.SearchMaster(2, oTrans.Master(2))
      End If
   Case 1   'Spouse
      If txtField(87).Locked = False Then
         Call oTrans.SearchMaster(87, txtField(87).Text)
      Else
         Call oTrans.SearchMaster(19, oTrans.Master(19))
      End If
   Case 2   'Comaker
      Call oTrans.SearchMaster(92, txtField(92).Text)
   Case 3   'Financer
      Call oTrans.SearchOnWays(86, txtWaysMn(86).Text)
   Case 4   'Dependent
      Call oTrans.SearchOnDependent(80, txtDependnt(80).Text)
   End Select
End Sub

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   GridEditor1.Refresh
   GridEditor2.Refresh
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String

   lsOldProc = "Form_Load"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   ''On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oTrans = New ggcLRApplication.clsLRApplication
   Set oTrans.AppDriver = oApp
   oTrans.UnitApplied = 3 'Financing MP
   oTrans.InitTransaction
   oTrans.LoadMode = 1
   
   oTrans.Filter = ""
   If oApp.ProductID <> "LRTrackr" Then oTrans.Filter = "a.sTransNox LIKE " & strParm(oApp.BranchCode & "%")
   'oTrans.Filter = oTrans.Filter & IIf(Trim(oTrans.Filter) = "", "", " AND ") & " a.cTranStat IN ('0','1', '3', '2') AND sApproved = ''"
   oTrans.Filter = oTrans.Filter & IIf(Trim(oTrans.Filter) = "", "", " AND ") & " (a.cTranStat IN ('0','1') OR (a.cTranStat IN ('3', '2') AND sApproved = ''))"
      
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualLeft

   InitGrid
   InitEntry
   InitValue
   initButton xeModeReady

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub cmbField_GotFocus(Index As Integer)
   psObjectNme = "cmbField"
End Sub

Private Sub cmbField_LostFocus(Index As Integer)
   If cmbField(Index).ListIndex < 0 Then cmbField(Index).ListIndex = -1
End Sub

Private Sub InitEntry()
   Dim lotxt As TextBox

   txtField(0).Enabled = False

   'set maximum lenth to txtField
   For Each lotxt In txtField
      Select Case lotxt.Index
      Case 15, 25, 16, 17, 7, 20, 21
         lotxt.MaxLength = oTrans.MasterMasFldSize(lotxt.Index)
      Case Else
         lotxt.MaxLength = 0
      End Select
   Next

   'set maximum lenth to txtWaysMn
   For Each lotxt In txtWaysMn
      Select Case lotxt.Index
      Case 0, 1, 3, 6, 10, 11, 13, 25, 26, 28, 31, 35, 36, 38, 44
         lotxt.MaxLength = oTrans.WayMeansMasFldSize(lotxt.Index)
      Case Else
         lotxt.MaxLength = 0
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

   cmbField(3).List(0) = "Regular"
   cmbField(3).List(1) = "Casual"
   cmbField(3).List(2) = "Provisionary"

   cmbField(4).List(0) = "Regular"
   cmbField(4).List(1) = "Casual"
   cmbField(4).List(2) = "Provisionary"

   'Comaker
   cmbRelation(0).List(0) = "Others"
   cmbRelation(0).List(1) = "Father"
   cmbRelation(0).List(2) = "Mother"
   cmbRelation(0).List(3) = "Child"
   cmbRelation(0).List(4) = "Sibling"
   cmbRelation(0).List(5) = "Spouse"

   'Dependent
   cmbRelation(1).List(0) = "Others"
   cmbRelation(1).List(1) = "Father"
   cmbRelation(1).List(2) = "Mother"
   cmbRelation(1).List(3) = "Child"
   cmbRelation(1).List(4) = "Sibling"
   cmbRelation(1).List(5) = "Spouse"

   'Financer
   cmbRelation(2).List(0) = "Others"
   cmbRelation(2).List(1) = "Father"
   cmbRelation(2).List(2) = "Mother"
   cmbRelation(2).List(3) = "Child"
   cmbRelation(2).List(4) = "Sibling"
   cmbRelation(2).List(5) = "Spouse"

   panCmdHwnd(0) = cmdButton(1).hwnd   ' Search
   panCmdHwnd(1) = cmdButton(2).hwnd   ' Delete
   panCmdHwnd(2) = cmdButton(3).hwnd   ' Cancel
   panCmdHwnd(3) = cmdButton(0).hwnd   ' Save
End Sub

Private Sub GridEditor1_DblClick()
   Dim lotxt As TextBox

   If oTrans.EditMode = xeModeReady Or oTrans.EditMode = xeModeUnknown Then Exit Sub

   'Load the double-click row to the text box
   If oTrans.LoadDependent(GridEditor1.Row - 1) Then
      For Each lotxt In txtDependnt
         lotxt.Text = IFNull(oTrans.Dependent(lotxt.Index))
      Next
      cmbRelation(1).ListIndex = IFNull(oTrans.Dependent("sRelatnID"), "0")

      'Lock/Unlock textbox
      If txtDependnt(80).Text <> Empty Then
         txtDependnt(80).Locked = True
         cmdInfo(4).Enabled = True
      Else
         txtDependnt(80).Locked = False
         cmdInfo(4).Enabled = False
      End If

      pnRow = GridEditor1.Row - 1
   End If
End Sub

Private Sub GridEditor1_GotFocus()
   psObjectNme = "GridEditor1"
End Sub

Private Sub GridEditor2_AddingRow(Cancel As Boolean)
   With GridEditor2
      If .TextMatrix(.Row, 1) = "" Then Cancel = True
      If .Rows > 12 Then
         .ColWidth(2) = 3117
         .ColWidth(3) = 3118
      End If
   End With
End Sub

Private Sub GridEditor2_GotFocus()
   SSTab1.Tab = 2
   psObjectNme = "GridEditor2"
End Sub

Private Sub GridEditor2_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String

   lsOldProc = "GridEditor2_KeyDown"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   ''On Error GoTo errProc

   If KeyCode = vbKeyF3 Then
      With GridEditor2
         Select Case .Col
         Case 3
            oTrans.SearchOnReference .Row - 1, .Col - 1, .TextMatrix(.Row, .Col)
            .SetFocus
            .Refresh
         End Select
      End With
      KeyCode = 0
   End If

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " _
                       & "  " & KeyCode _
                       & ", " & Shift _
                       & " )", True
End Sub

Private Sub GridEditor2_EditorValidate(Cancel As Boolean)
   Dim lsOldProc As String

   lsOldProc = "GridEditor2_EditorValidate"
   Debug.Print pxeMODULENAME & "." & lsOldProc

   With GridEditor2
      oTrans.Reference(.Row - 1, .Col - 1) = .TextMatrix(.Row, .Col)
      If .Col = 3 Then .TextMatrix(.Row, .Col) = oTrans.Reference(.Row - 1, .Col - 1)
   End With
End Sub

Private Sub oTrans_LoadData()
   LoadMaster
   LoadDetail
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
   Case 98
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
   ''On Error GoTo errProc
   
   Select Case Index
   Case 0
      If oTrans.SearchTransaction Then
         LoadMaster
         LoadDetail
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
   xrFrame1(1).Enabled = Not lbShow

   For pnCtr = 0 To xrFrame3.Count - 1
      xrFrame3(pnCtr).Enabled = lbShow
   Next

   With GridEditor2
      For pnCtr = 1 To .Cols - 1
         .ColEnabled(pnCtr) = lbShow
      Next
   End With

   'Disable entry for name, spouse, comaker, financer, and dependent if already entered
   'to update use the command button
   'Kalyptus - 2009.05.01
   If lnStat = xeModeUpdate Then
      'Applicant
      If txtField(80).Text <> Empty Then
         txtField(80).Locked = True
         cmdInfo(0).Enabled = True
      Else
         txtField(80).Locked = False
         cmdInfo(0).Enabled = False
      End If

      'Spouse
      If txtField(87).Text <> Empty Then
         txtField(87).Locked = True
         cmdInfo(1).Enabled = True
      Else
         txtField(87).Locked = False
         cmdInfo(1).Enabled = False
      End If

      'Comaker
      If txtField(92).Text <> Empty Then
         txtField(92).Locked = True
         cmdInfo(2).Enabled = True
      Else
         txtField(92).Locked = False
         cmdInfo(2).Enabled = False
      End If

      'Financer
      If txtWaysMn(86).Text <> Empty Then
         txtWaysMn(86).Locked = True
         cmdInfo(3).Enabled = True
      Else
         txtWaysMn(86).Locked = False
         cmdInfo(3).Enabled = False
      End If

      cmdInfo(4).Enabled = False
   Else
      cmdInfo(0).Enabled = False
      cmdInfo(1).Enabled = False
      cmdInfo(2).Enabled = False
      cmdInfo(3).Enabled = False
      cmdInfo(4).Enabled = False
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

Private Sub SSTab1_Click(PreviousTab As Integer)
   Dim lsOldProc As String

   lsOldProc = "SSTab1_Click"
   Debug.Print pxeMODULENAME & "." & lsOldProc
   ''On Error GoTo errProc

   If GridEditor1.Visible Then GridEditor1.Visible = False
   If GridEditor2.Visible Then GridEditor2.Visible = False

   With SSTab1
      Select Case .Tab
      Case 2
         GridEditor2.Visible = True
         GridEditor2.SetFocus
      Case 7
         GridEditor1.Visible = True
      End Select
   End With

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & PreviousTab & " )", True
End Sub

Private Sub InitGrid()
   Dim lnCtr As Integer

   With GridEditor1
      .Rows = 6
      .Cols = 6
      .Font = "MS Sans Serif"

      'column title
      .TextMatrix(0, 1) = "Full Name"
      .TextMatrix(0, 2) = "Age"
      .TextMatrix(0, 3) = "School/Address"
      .TextMatrix(0, 4) = "Company"
      .TextMatrix(0, 5) = "Remarks"
      .Row = 0

      'Set Column Alignment and Disable Grid Editing
      For pnCtr = 0 To .Cols - 1
         .Col = pnCtr
         .CellFontBold = True
         .CellAlignment = 1
         .ColEnabled(pnCtr) = False
      Next

      'Column Width
      .ColWidth(0) = 330
      .ColWidth(1) = 2450
      .ColWidth(2) = 430
      .ColWidth(3) = 2300
      .ColWidth(4) = 2300
      .ColWidth(5) = 1925

      .ColNumberOnly(2) = True
      .ColLimit(1) = 25
      .ColLimit(3) = 50
      .ColLimit(4) = 50
      .ColLimit(5) = 30

      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      .ColAlignment(4) = 1
      .ColAlignment(5) = 1

      .ScrollBars = flexScrollBarVertical
      .Row = 1
      .Col = 1
   End With

   With GridEditor2
      .Cols = 4
      .Font = "MS Sans Serif"

      'column title
      .TextMatrix(0, 1) = "Full Name"
      .TextMatrix(0, 2) = "Address"
      .TextMatrix(0, 3) = "Town/City"
      .Row = 0

       'Column Alignment
      For pnCtr = 0 To .Cols - 1
         .Col = pnCtr
         .CellFontBold = True
         .CellAlignment = 1
      Next

      'Column Width
      .ColWidth(0) = 330
      .ColWidth(1) = 3140

      .ColLimit(1) = 35
      .ColLimit(2) = 45
      .ColLimit(3) = 50

      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1

      .Row = 1
      .Col = 1
   End With
End Sub

Private Sub InitValue()
   Dim lotxt As TextBox
   Dim lnCtr As Integer

   'Load Value from the master table
   For Each lotxt In txtField
      Select Case lotxt.Index
      Case 8 To 10, 12
         lotxt.Text = ".00"
      Case Else
         lotxt.Text = ""
      End Select
   Next

   'Load Value from the ways and means
   For Each lotxt In txtWaysMn
      Select Case lotxt.Index
      Case 7, 8, 32, 33, 40, 15 To 24
         lotxt.Text = "0.00"
      Case Else
         lotxt.Text = oTrans.Master(lotxt.Index)
      End Select
   Next

   txtOthers(0).Text = "yrs"
   txtOthers(1).Text = "yrs"
   txtOthers(2).Text = "0.00"
   txtOthers(3).Text = "0.00"
   txtOthers(4).Text = "0.00"

   cmbField(0).ListIndex = 0
   cmbField(1).ListIndex = 0
   cmbField(2).ListIndex = 0
   cmbField(3).ListIndex = 0
   cmbField(4).ListIndex = 0

   Label2.Caption = "UNKNOWN"

   'Initialize the value of Grids
   With GridEditor1
      For pnCtr = 1 To .Rows - 1
         For lnCtr = 1 To .Cols - 1
            .TextMatrix(pnCtr, lnCtr) = ""
         Next
      Next
      .Col = 1
   End With

   With GridEditor2
      .Rows = 2
      .Col = 1

      For pnCtr = 1 To .Cols - 1
         .TextMatrix(1, pnCtr) = ""
      Next

      .ColWidth(2) = 3217
      .ColWidth(3) = 3218
   End With

   SSTab1.Tab = 0
   GridEditor1.Visible = False
   GridEditor2.Visible = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyReturn, vbKeyUp, vbKeyDown
      If psObjectNme = "cmbField" And KeyCode <> vbKeyReturn Then
         Exit Sub
      End If

      Select Case KeyCode
      Case vbKeyReturn, vbKeyDown
         If GetFocus = GridEditor1.hwnd Or GetFocus = GridEditor2.hwnd Then Exit Sub
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
   txtOthers(3).Text = Format(lnTotExpense, "#,##0.00")
   txtOthers(4).Text = Format(IFNull(oTrans.WaysMeans("nMonGross"), 0) + IFNull(oTrans.WaysMeans("nMonOther"), 0) - lnTotExpense, "#,##0.00")
End Sub




Private Sub txtOthers_GotFocus(Index As Integer)
   psObjectNme = "txtOthers"
End Sub

Private Sub LoadMaster()
   Dim lotxt As TextBox
   Dim lnCtr As Integer
   Dim lsOldProc As String

   lsOldProc = "LoadMaster()"
   Debug.Print pxeMODULENAME & "." & lsOldProc

   'Load Value from the master table
   For Each lotxt In txtField
      Select Case lotxt.Index
      Case 0
         lotxt = Format(oTrans.Master(lotxt.Index), IIf(Len(oApp.BranchCode) = 2, "@@@@-@@@@@@", "@@@@@@-@@@@@@"))
         txtSearch(0) = oTrans.Master(0)
         txtSearch(1) = oTrans.Master(80)
         txtSearch(0).Tag = oTrans.Master(0)
         txtSearch(1).Tag = oTrans.Master(80)
      Case 8 To 10, 12
         lotxt.Text = Format(oTrans.Master(lotxt.Index), "#,##0.00")
      Case 11
         lotxt.Text = CInt(oTrans.Master(lotxt.Index))
      Case 3, 81, 88
         lotxt.Text = Format(oTrans.Master(lotxt.Index), "Mmm DD, YYYY")
      Case Else
         lotxt.Text = IFNull(oTrans.Master(lotxt.Index))
      End Select
   Next

   'Load Value from the ways and means
   For Each lotxt In txtWaysMn
      Select Case lotxt.Index
      Case 7, 8, 32, 33, 40, 15 To 24, 46, 47
         lotxt.Text = Format(IFNull(oTrans.WaysMeans(lotxt.Index), 0), "#,##0.00")
      Case Else
         lotxt.Text = IFNull(oTrans.WaysMeans(lotxt.Index))
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
   txtOthers(2).Text = Format(IFNull(oTrans.WaysMeans("nMonGross"), 0) + IFNull(oTrans.WaysMeans("nMonOther"), 0), "#,##0.00")
   ComInEx

   cmbField(0).ListIndex = IFNull(oTrans.Master("cApplType"), 0)
   cmbField(1).ListIndex = IFNull(oTrans.Master("cUnitAppl") - 3, 0)
   cmbField(2).ListIndex = IIf(IsNull(oTrans.Master("cCvlStat1")), -1, oTrans.Master("cCvlStat1"))
   cmbField(3).ListIndex = IIf(IFNull(oTrans.WaysMeans("cEmpStatx")) = "", -1, oTrans.WaysMeans("cEmpStatx"))
   cmbField(4).ListIndex = IIf(IFNull(oTrans.WaysMeans("cEmpStat1")) = "", -1, oTrans.WaysMeans("cEmpStat1"))

   cmbRelation(0).ListIndex = CInt(IIf(IsNumeric(oTrans.Master("sRelation")) = False, -1, oTrans.Master("sRelation")))
   cmbRelation(1).ListIndex = CInt(IIf(IsNumeric(oTrans.Dependent("sRelatnID")) = False, -1, oTrans.Dependent("sRelatnID")))
   cmbRelation(2).ListIndex = CInt(IIf(IsNumeric(oTrans.WaysMeans("sReltnCD2")) = False, -1, oTrans.WaysMeans("sReltnCD2")))

   Label2.Caption = Format(ApplStat(oTrans.Master("cTranStat")), ">")

   SSTab1.Tab = 0

   GridEditor1.Visible = False
   GridEditor2.Visible = False

End Sub

Private Sub LoadDetail()
   Dim lnCtr As Integer
   Dim lotxt As TextBox
   Dim lsOldProc As String

   lsOldProc = "LoadDetail()"
   Debug.Print pxeMODULENAME & "." & lsOldProc

   With GridEditor1
      .Rows = oTrans.NoofDependent + 1
      For pnCtr = 1 To .Rows - 1
         'Load each recordset
         If oTrans.LoadDependent(pnCtr - 1) Then
            For lnCtr = 1 To .Cols - 1
               If lnCtr = 1 Then
                  .TextMatrix(pnCtr, lnCtr) = oTrans.Dependent("sFullName")
               Else
                  .TextMatrix(pnCtr, lnCtr) = oTrans.Dependent(lnCtr - 1)
               End If
            Next
         End If
      Next

      'Load to the text boxes the last loaded record
      For Each lotxt In txtDependnt
         lotxt.Text = IFNull(oTrans.Dependent(lotxt.Index))
      Next
      cmbRelation(1).ListIndex = IFNull(oTrans.Dependent("sRelatnID"), "0")

      'Save the record pointer of the last loaded record
      pnRow = GridEditor1.Rows - 2

   End With

   With GridEditor2
      .Rows = IIf(oTrans.NoofReference < 1, 1, oTrans.NoofReference) + 1
      For pnCtr = 1 To .Rows - 1
         For lnCtr = 1 To .Cols - 1
            .TextMatrix(pnCtr, lnCtr) = oTrans.Reference(pnCtr - 1, lnCtr - 1)
         Next
      Next
   End With
End Sub

Private Sub txtSearch_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String

   lsOldProc = "txtSearch_KeyDown"
   ''On Error GoTo errProc

   If KeyCode = vbKeyReturn Then
      If txtSearch(Index).Text <> txtSearch(Index).Tag Then
         If oTrans.SearchTransaction(IIf(Index = 0, CodeFormat(oApp.BranchCode, txtSearch(Index).Text) _
            , txtSearch(Index).Text) _
            , IIf(Index = 0, True, False)) Then
            LoadMaster
            LoadDetail
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
   ''On Error GoTo errProc

   If txtSearch(Index).Text <> "" Then
      If txtSearch(Index).Text <> txtSearch(Index).Tag Then
         If oTrans.SearchTransaction(IIf(Index = 0, CodeFormat(oApp.BranchCode, txtSearch(Index).Text) _
            , txtSearch(Index).Text) _
            , IIf(Index = 0, True, False)) Then
            LoadMaster
            LoadDetail
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

Function isEQDep() As Boolean
   Dim lotxt As TextBox
   With GridEditor1
      For Each lotxt In txtDependnt
         If lotxt.Index = 80 Then
            If lotxt.Text <> .TextMatrix(pnRow + 1, 1) Then
               Exit Function
            End If
         Else
            If lotxt.Text <> .TextMatrix(pnRow + 1, lotxt.Index) Then
               Exit Function
            End If
         End If
      Next
   End With
   isEQDep = True
End Function

Private Sub xrButton1_Click()
   If txtField(9).Locked = True Then
      Call oTrans.AutoCompute(oTrans.Master("sModelIDx"), oTrans.Master("nPNValueX"), oTrans.Master("nDownPaym"), oTrans.Master("nAcctTerm"))
   Else
      oTrans.Master(12) = oTrans.Master("nPNValueX") / oTrans.Master("nAcctTerm")
   End If
   
   oTrans.Master(8) = oTrans.Master("nPNValueX") + oTrans.Master("nDownPaym")
End Sub

