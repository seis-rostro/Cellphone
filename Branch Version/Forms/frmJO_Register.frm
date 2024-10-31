VERSION 5.00
Object = "{0A7B56A6-35D0-4533-91DA-1715D3A0DD3E}#1.1#0"; "xrGridControl.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmJobOrder_Register 
   BorderStyle     =   0  'None
   Caption         =   "Job Order Register"
   ClientHeight    =   5895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11205
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5895
   ScaleWidth      =   11205
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1920
      Index           =   4
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   3000
      Width           =   9540
      _ExtentX        =   16828
      _ExtentY        =   3387
      BackColor       =   14286077
      ClipControls    =   0   'False
      Begin xrGridEditor.GridEditor GridEditor1 
         Height          =   1785
         Left            =   45
         TabIndex        =   13
         Tag             =   "et0;eb0;et0;bc2"
         Top             =   60
         Width           =   9405
         _ExtentX        =   16589
         _ExtentY        =   3149
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
         Object.HEIGHT          =   1785
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
         MOUSEICON       =   "frmJO_Register.frx":0000
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
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   840
      Index           =   5
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   4965
      Width           =   9540
      _ExtentX        =   16828
      _ExtentY        =   1482
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtfield 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00D9FCFD&
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   8
         Left            =   7215
         TabIndex        =   45
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   405
         Width           =   2235
      End
      Begin VB.TextBox txtfield 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00D9FCFD&
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   6
         Left            =   4140
         TabIndex        =   20
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   405
         Width           =   1600
      End
      Begin VB.TextBox txtfield 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00D9FCFD&
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   5
         Left            =   4140
         TabIndex        =   18
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   105
         Width           =   1600
      End
      Begin VB.TextBox txtfield 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00D9FCFD&
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   4
         Left            =   1230
         TabIndex        =   17
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   405
         Width           =   1600
      End
      Begin VB.TextBox txtfield 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00D9FCFD&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   3
         Left            =   1230
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   105
         Width           =   1600
      End
      Begin VB.TextBox txtfield 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   7
         Left            =   7215
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Tag             =   "ht0;ft0"
         Text            =   "Total Amount"
         Top             =   105
         Width           =   2235
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Date Released"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   3
         Left            =   5760
         TabIndex        =   46
         Top             =   405
         Width           =   1230
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Spareparts"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   14
         Left            =   2985
         TabIndex        =   21
         Top             =   105
         Width           =   1035
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Labor"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   11
         Left            =   -45
         TabIndex        =   19
         Top             =   405
         Width           =   1155
      End
      Begin VB.Label lblField 
         BackStyle       =   0  'Transparent
         Caption         =   "Miscellaneous"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   10
         Left            =   120
         TabIndex        =   4
         Top             =   105
         Width           =   1170
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Amount Paid"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   9
         Left            =   2790
         TabIndex        =   16
         Top             =   405
         Width           =   1230
      End
      Begin VB.Label lblField 
         BackStyle       =   0  'Transparent
         Caption         =   " Total"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   100
         Left            =   6480
         TabIndex        =   15
         Top             =   105
         Width           =   660
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   480
      Index           =   0
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   9540
      _ExtentX        =   16828
      _ExtentY        =   847
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   4935
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   90
         Width           =   4455
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   0
         Left            =   1305
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   90
         Width           =   2010
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name"
         Height          =   195
         Index           =   8
         Left            =   3720
         TabIndex        =   2
         Top             =   90
         Width           =   1125
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Job Order No."
         Height          =   195
         Index           =   5
         Left            =   150
         TabIndex        =   0
         Top             =   90
         Width           =   990
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1200
      Index           =   1
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   1065
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   2117
      BackColor       =   12632256
      Enabled         =   0   'False
      ClipControls    =   0   'False
      BorderStyle     =   3
      Begin VB.Label lblFields 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Sept. 15, 2000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   1350
         TabIndex        =   34
         Tag             =   "tc0"
         Top             =   855
         Width           =   3975
      End
      Begin VB.Label lblFields 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   1350
         TabIndex        =   33
         Tag             =   "tc0"
         Top             =   600
         Width           =   3750
      End
      Begin VB.Label lblFields 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "01231456890123456789012345678901234567"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   1350
         TabIndex        =   32
         Tag             =   "tc0"
         Top             =   345
         Width           =   4005
      End
      Begin VB.Label lblFields 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXX"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   1350
         TabIndex        =   31
         Tag             =   "tc0"
         Top             =   90
         Width           =   1305
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   30
         Left            =   2280
         TabIndex        =   30
         Top             =   690
         Width           =   30
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Received :"
         Height          =   195
         Index           =   28
         Left            =   105
         TabIndex        =   25
         Top             =   855
         Width           =   1170
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contact No. :"
         Height          =   195
         Index           =   2
         Left            =   330
         TabIndex        =   24
         Top             =   600
         Width           =   945
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Job Order No.:"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   23
         Top             =   90
         Width           =   1035
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name :"
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   22
         Top             =   345
         Width           =   1215
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1200
      Index           =   2
      Left            =   5610
      Tag             =   "wt0;fb0"
      Top             =   1065
      Width           =   4020
      _ExtentX        =   7091
      _ExtentY        =   2117
      BackColor       =   12632256
      ClipControls    =   0   'False
      BorderStyle     =   3
      Begin VB.Label lblFields 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXXXXXXXXXXXX"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   1395
         TabIndex        =   40
         Tag             =   "tc0"
         Top             =   90
         Width           =   2490
      End
      Begin VB.Label lblFields 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Sept. 15, 2000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   1380
         TabIndex        =   39
         Tag             =   "tc0"
         Top             =   345
         Width           =   2415
      End
      Begin VB.Label lblFields 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXXXXXXXXXXXX"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   1395
         TabIndex        =   38
         Tag             =   "tc0"
         Top             =   855
         Width           =   2595
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status :"
         Height          =   195
         Index           =   6
         Left            =   735
         TabIndex        =   37
         Top             =   855
         Width           =   540
      End
      Begin VB.Label lblFields 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXXXXXXXXXXXX"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   1395
         TabIndex        =   36
         Tag             =   "tc0"
         Top             =   600
         Width           =   2490
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Purchase Date :"
         Height          =   195
         Index           =   27
         Left            =   120
         TabIndex        =   28
         Top             =   345
         Width           =   1155
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IMEI No. :"
         Height          =   195
         Index           =   7
         Left            =   555
         TabIndex        =   27
         Top             =   90
         Width           =   720
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Brand & Model :"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   26
         Top             =   600
         Width           =   1035
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   705
      Index           =   3
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   2250
      Width           =   9540
      _ExtentX        =   16828
      _ExtentY        =   1244
      BackColor       =   12632256
      Enabled         =   0   'False
      ClipControls    =   0   'False
      BorderStyle     =   3
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   2
         Left            =   1350
         MultiLine       =   -1  'True
         TabIndex        =   35
         Tag             =   "tc0;fb0"
         Text            =   "frmJO_Register.frx":001C
         Top             =   90
         Width           =   4050
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Description :"
         Height          =   195
         Index           =   15
         Left            =   5565
         TabIndex        =   44
         Top             =   345
         Width           =   1215
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Job Description :"
         Height          =   195
         Index           =   13
         Left            =   5610
         TabIndex        =   43
         Top             =   90
         Width           =   1185
      End
      Begin VB.Label lblFields 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXXXXXXXXXXXX"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   8
         Left            =   6885
         TabIndex        =   42
         Tag             =   "tc0"
         Top             =   90
         Width           =   2655
      End
      Begin VB.Label lblFields 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXXXXXXXXXXXX"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   9
         Left            =   6885
         TabIndex        =   41
         Tag             =   "tc0"
         Top             =   345
         Width           =   2595
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Complaint :"
         Height          =   195
         Index           =   12
         Left            =   495
         TabIndex        =   29
         Top             =   75
         Width           =   780
      End
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   405
      Index           =   3
      Left            =   9855
      TabIndex        =   8
      Top             =   1800
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
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
      Picture         =   "frmJO_Register.frx":002B
      CaptionAlign    =   0
      BackColor       =   14286077
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   4
      Left            =   9855
      TabIndex        =   10
      Top             =   2220
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
      Caption         =   "&Delete"
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
      Picture         =   "frmJO_Register.frx":07A5
      CaptionAlign    =   0
      BackColor       =   14286077
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   5
      Left            =   9855
      TabIndex        =   12
      Top             =   2640
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
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
      Picture         =   "frmJO_Register.frx":0F1F
      CaptionAlign    =   0
      BackColor       =   14286077
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   1
      Left            =   9855
      TabIndex        =   7
      Top             =   1800
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
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
      Picture         =   "frmJO_Register.frx":1699
      CaptionAlign    =   0
      BackColor       =   14286077
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   0
      Left            =   9855
      TabIndex        =   6
      Top             =   1380
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
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
      Picture         =   "frmJO_Register.frx":1E13
      CaptionAlign    =   0
      BackColor       =   14286077
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   6
      Left            =   9855
      TabIndex        =   11
      Top             =   2640
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
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
      Picture         =   "frmJO_Register.frx":258D
      CaptionAlign    =   0
      BackColor       =   14286077
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   2
      Left            =   9855
      TabIndex        =   9
      Top             =   2220
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
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
      Picture         =   "frmJO_Register.frx":2D07
      CaptionAlign    =   0
      BackColor       =   14286077
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
End
Attribute VB_Name = "frmJobOrder_Register"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents oDriver As FormDriver
Attribute oDriver.VB_VarHelpID = -1
Private oSkin As FormSkin
Private bLoaded As Boolean
Private oRS As New ADODB.Recordset

Dim txtfieldGotfocus As Boolean
Dim txtOthersGotfocus As Boolean
Dim pbnewitem As Boolean

Dim psSelected() As String

Dim pnindex As Integer
Dim Index As Integer
Dim pnCtr As Integer
Dim lnCtr As Integer

Private Sub cmdButton_Click(Index As Integer)
Dim Cancel As Boolean
   
   Select Case Index
      Case 0   'New
         InitTxtField
         EmptyGrid
         InitButton xeModeAddNew
         
      Case 1   'Save
         If lblFields(0).Caption = "" Then Exit Sub
         Cancel = Not UpdateCP_JO_Detail
         InitButton xeModeAddNew
         
      Case 2   'Update
         If txtField(0).Text <> "" Then
            InitButton xeModeReady
            GridEditor1.SetFocus
         Else
            txtField(0).SetFocus
         End If
         
      Case 3   'Search
         If pnindex >= 0 And pnindex <= 1 Then Search_JobOrder False
         
      Case 4 'Delete Row
            With GridEditor1
               If .Rows <> 2 Then
                  .DeleteRow
               End If
            End With
            
      Case 5 'Cancel
         InitButton xeModeAddNew
         
      Case 6
         Unload Me
   End Select

End Sub

Private Sub InitButton(lnStat As Integer)
   Dim lbShow As Boolean

   lbShow = IIf(lnStat = xeModeReady, False, True)
   cmdButton(2).Visible = lbShow
   cmdButton(3).Visible = lbShow
   cmdButton(6).Visible = lbShow
   cmdButton(0).Visible = lbShow
   
   cmdButton(1).Visible = Not lbShow
   cmdButton(4).Visible = Not lbShow
   cmdButton(5).Visible = Not lbShow
   
   xrFrame1(4).Enabled = Not lbShow
   xrFrame1(5).Enabled = Not lbShow
   
End Sub

Private Sub Form_Load()

CenterChildForm mdiMain, Me
bLoaded = False

Set oDriver = New FormDriver
Set oDriver.AppDriver = oApp
Set oDriver.MainForm = Me

InitButton xeModeAddNew

Set oSkin = New FormSkin
Set oSkin.AppDriver = oApp
Set oSkin.Form = Me
oSkin.ApplySkin xeFormTransMaintenance
   
InitTxtField
InitGrid
EmptyGrid
         
End Sub

Private Sub InitGrid()
Dim Index As Long

   With GridEditor1
      .Rows = 2
      .Cols = 8
      .Font = "MS Sans Serif"
      
      'column title
      .TextMatrix(0, 1) = "Particulars"
      .TextMatrix(0, 2) = "Parts Amount"
      .TextMatrix(0, 3) = "Labor Amount"
      .TextMatrix(0, 4) = "%"
      .TextMatrix(0, 5) = "% Amt."
      .TextMatrix(0, 6) = "Qty"
      .TextMatrix(0, 7) = "Sub Total"
               
      'column width
      .ColWidth(0) = 500
      .ColWidth(1) = 3400
      .ColWidth(2) = 1200
      .ColWidth(3) = 1200
      .ColWidth(4) = 500
      .ColWidth(5) = 800
      .ColWidth(6) = 500
      .ColWidth(7) = 1200
      .ColEnabled(7) = False
      
      For Index = 1 To 7
         Select Case Index
            Case 1
               .ColAlignment(Index) = 1
            Case 2, 3, 5, 7
               .ColFormat(Index) = "#,##0.00"
               .ColDefault(Index) = "0.00"
               .ColAlignment(Index) = 6
               If Index = 5 Then .ColEnabled(Index) = False
            Case 4, 6
               .ColAlignment(Index) = 6
               If Index = 4 Then
                  .ColMaxValue(Index) = 100
                  .ColDefault(Index) = 0
               ElseIf Index = 6 Then
                  .ColMaxValue(Index) = 999
                  .ColDefault(Index) = 1
               End If
         End Select
      Next
      .ColDefault(6) = 1
        
      .Row = 1
      .Rows = 2
      
   End With
   
End Sub
Private Sub InitTxtField()
Dim lnCtr As Integer
Dim Index As Integer

For lnCtr = 0 To 9
   lblFields(lnCtr).Caption = ""
Next

For Index = 0 To 8
   Select Case Index
      Case 3 To 7
         txtField(Index).Text = "0.00"
      Case 0 To 2, 8
         txtField(Index).Text = ""
   End Select
Next

txtField(8).Enabled = True

End Sub
Private Sub Search_JobOrder(ByVal SearchValue As Boolean)
Dim lsSearch As String
Dim lsSQL As String
Dim lnCtr As Integer
Dim lrs As ADODB.Recordset
Dim Index As Integer

   Set lrs = New ADODB.Recordset
   lsSQL = "SELECT" _
                  & " a.sJobOrdNo, " _
                  & " d.sLastName + ', ' + d.sFrstName + ' ' + d.sMiddName as xFullName, " _
                  & " a.sTelNoxxx, " _
                  & " a.dTransact, " _
                  & " a.sIMEINoxx, " _
                  & " a.dPurchase, " _
                  & " b.sBrandNme+' '+ c.sModelNme as BrandModel, " _
                  & " a.cTranstat, " _
                  & " a.cWarranty, " _
                  & " a.cCategory, " _
                  & " a.sCategory, " _
                  & " a.sComplent, " _
                  & " a.nTranTotl, " _
                  & " a.sBckJobNo, " _
                  & " a.dPaymentx, " _
                  & " a.sTransNox, " _
                  & " a.nMiscChrg, " _
                  & " a.nLaborTot, " _
                  & " a.nPartsTot, " _
                  & " a.nAmtPaidx " _
                  
   lsSQL = lsSQL _
               & " FROM CP_JobOrder_Master a " _
                  & " LEFT JOIN Brand b " _
                     & " ON a.sBrandIDx = b.sBrandIDx " _
                  & " LEFT JOIN Model c " _
                     & " ON a.sModelIDx = c.sModelIDx " _
                  & " LEFT JOIN Client_Master d " _
                     & " ON a.sClientID = d.sClientID " _

                     
   If pnindex = 0 Then
      If SearchValue Then
         lsSQL = lsSQL & " WHERE sJobOrdNo = '" & txtField(0).Text & "'"
      Else
         lsSQL = lsSQL & " WHERE sJobOrdNo LIKE '" & txtField(0).Text & "%' "
      End If
   ElseIf pnindex = 1 Then
      If SearchValue Then
         lsSQL = lsSQL & " WHERE sLastName + ', ' + sFrstName + ' ' + sMiddName = '" & txtField(1).Text & "'"
      Else
         lsSQL = lsSQL & " WHERE sLastName + ', ' + sFrstName + ' ' + sMiddName LIKE '" & txtField(1).Text & "%' "
      End If
   End If
                           
   lsSQL = lsSQL & " AND left(a.sTransNox,2) = '" & oApp.BranchCode & "'" _
                  & " ORDER BY sLastName + ', ' + sFrstName + ' ' + sMiddName"
   lrs.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
            
   If lrs.RecordCount = 1 Then
      For Index = 0 To 8
         Select Case Index
            Case 0, 1
               txtField(Index).Text = IIf(IsNull(lrs(Index)), "", lrs(Index))
               If Index = 0 Then txtField(Index).Tag = lrs(15)
            Case 2
               txtField(Index).Text = lrs(11)
            Case 3 To 6
               txtField(Index).Text = Format(lrs(Index + 13), "#,##0.00")
            Case 7
               txtField(Index).Text = Format(lrs(12), "#,##0.00")
            Case 8
               If lrs("cTranStat") = 1 Then
                  txtField(Index).Text = Format(lrs(14), "MMMM dd, yyyy")
               Else
                  txtField(Index).Text = ""
               End If
         End Select
      Next

      For lnCtr = 0 To 9
         With lblFields(lnCtr)
            Select Case lnCtr
               Case 0 To 2, 4, 6
                  .Caption = IIf(IsNull(lrs(lnCtr)), "", lrs(lnCtr))
               Case 3, 5
                  .Caption = Format(lrs(lnCtr), "MMMM dd, yyyy")
               Case 7
                  If lrs(lnCtr) = 0 Then
                     .Caption = "Pending"
                  Else
                     .Caption = "Claimed"
                  End If
               Case 8
                  '1-Void Warranty ; 2-Under Limited Warranty ; 3-Back Job
                  If lrs(lnCtr) = 1 Then
                     .Caption = "Void Warranty"
                  ElseIf lrs(lnCtr) = 2 Then
                     .Caption = "Limited Warranty"
                  ElseIf lrs(lnCtr) = 3 Then
                     .Caption = "Back Job" & " " & lrs(13)
                  End If
               Case 9
                  '0-New Unit; 1-Battery; 2-Sim Card; 3-Back Cover; 4-Charger; 5-Others
                  Select Case lrs(lnCtr)
                     Case 0
                        .Caption = "New Unit"
                     Case 1
                        .Caption = "Battery"
                     Case 2
                        .Caption = "Sim Card"
                     Case 3
                        .Caption = "Back Cover"
                     Case 4
                        .Caption = "Charger"
                     Case 5
                        .Caption = lrs(10)
                  End Select
            End Select
         End With
      Next
      ShowGrid
      InitButton xeModeAddNew
      
   ElseIf lrs.RecordCount > 1 Then
        lsSearch = KwikBrowse(oApp, lrs, _
                          "sJobOrdNo»xFullName»sIMEINoxx»BrandModel»dTransact", _
                          "J.O. No.»Customer Name»IMEI No.»Brand & Model»Tran. Date", _
                          "@»@»@»@»MMM dd, yyyy»MMM dd, yyyy")

        If lsSearch <> "" Then
            psSelected = Split(lsSearch, "»")
            ShowMoreRec
            ShowGrid
        End If
        InitButton xeModeAddNew
      
   Else
      MsgBox "No Record Found!!!", vbInformation, Me.Caption
   End If
   
   Set lrs = Nothing

End Sub

Private Sub ShowMoreRec()
Dim Index As Integer

For Index = 0 To 8
   Select Case Index
      Case 0, 1
         txtField(Index).Text = psSelected(Index)
         If Index = 0 Then txtField(Index).Tag = psSelected(15)
      Case 2
         txtField(Index).Text = psSelected(11)
      Case 3 To 6
         txtField(Index).Text = Format(psSelected(Index + 13), "#,##0.00")
      Case 7
         txtField(Index).Text = Format(psSelected(12), "#,##0.00")
      Case 8
         If psSelected(7) = 1 Then
            txtField(Index).Text = Format(psSelected(14), "MMMM dd, yyyy")
         Else
            txtField(Index).Text = ""
         End If
   End Select
Next
      For lnCtr = 0 To 9
         With lblFields(lnCtr)
            Select Case lnCtr
               Case 0 To 2, 4, 6
                  .Caption = psSelected(lnCtr)
               Case 3, 5
                  .Caption = Format(psSelected(lnCtr), "MMMM dd, yyyy")
               Case 7
                  If psSelected(lnCtr) = 0 Then
                     .Caption = "Pending"
                  Else
                     .Caption = "Claimed"
                  End If
               Case 8
                  '1-Void Warranty ; 2-Under Limited Warranty ; 3-Back Job
                  If psSelected(lnCtr) = 1 Then
                     .Caption = "Void Warranty"
                  ElseIf psSelected(lnCtr) = 2 Then
                     .Caption = "Limited Warranty"
                  ElseIf psSelected(lnCtr) = 3 Then
                     .Caption = "Back Job" & " " & psSelected(13)
                  End If
               Case 9
                  '0-New Unit; 1-Battery; 2-Sim Card; 3-Back Cover; 4-Charger; 5-Others
                  Select Case psSelected(lnCtr)
                     Case 0
                        .Caption = "New Unit"
                     Case 1
                        .Caption = "Battery"
                     Case 2
                        .Caption = "Sim Card"
                     Case 3
                        .Caption = "Back Cover"
                     Case 4
                        .Caption = "Charger"
                     Case 5
                        .Caption = psSelected(10)
                  End Select
            End Select
         End With
      Next

End Sub

Private Function UpdateCP_JO_Detail() As Boolean
Dim lsSQL As String
Dim lnrow As Long

UpdateCP_JO_Detail = True
oApp.Connection.BeginTrans

On Error GoTo errProc

   If oRS.State = adStateOpen Then oRS.Close

      oRS.Open "SELECT * From CP_JobOrder_Detail " _
               & "WHERE sTransNox = '" & txtField(0).Tag & "' ", _
                   oApp.Connection, adOpenDynamic, adLockReadOnly, adCmdText

   If oRS.RecordCount <> 0 Then
      lsSQL = "DELETE CP_JobOrder_Detail " _
               & "WHERE sTransNox = '" & txtField(0).Tag & "' "
      oApp.Connection.Execute lsSQL, lnrow, adCmdText
      If lnrow <> 0 Then oApp.RegisDelete lsSQL
   End If
      With GridEditor1
         For lnCtr = 1 To .Rows - 1
            lsSQL = "INSERT INTO CP_JobOrder_Detail " _
                     & "( sTransNox, " _
                     & "  nEntryNox, " _
                     & "  sDescript, " _
                     & "  nPartsAmt, " _
                     & "  nLaborAmt, " _
                     & "  nDiscount, " _
                     & "  nQuantity, " _
                     & "  dModified) " _
                         & "VALUES " _
                         & "('" & txtField(0).Tag & "', " _
                         & "'" & .TextMatrix(lnCtr, 0) & "', " _
                         & "'" & .TextMatrix(lnCtr, 1) & "', " _
                         & "'" & CDbl(.TextMatrix(lnCtr, 2)) & "', " _
                         & "'" & CDbl(.TextMatrix(lnCtr, 3)) & "', " _
                         & "'" & CLng(.TextMatrix(lnCtr, 4)) & "', " _
                         & "'" & CLng(.TextMatrix(lnCtr, 6)) & "', " _
                         & " getdate())"

            oApp.Connection.Execute lsSQL, lnrow, adCmdText

            lsSQL = "UPDATE CP_JobOrder_Master SET" _
                        & " nLaborTot = '" & CDbl(txtField(4).Text) & "', " _
                        & " nPartsTot = '" & CDbl(txtField(5).Text) & "', " _
                        & " nMiscChrg = '" & CDbl(txtField(3).Text) & "', " _
                        & " nTranTotl = '" & CDbl(txtField(7).Text) & "', " _
                        & " dModified = getdate() " _
                  & " WHERE sTransNox = '" & txtField(0).Tag & "' "
            oApp.Connection.Execute lsSQL, lnrow, adCmdText
            
            If txtField(8).Text <> "" Then
               lsSQL = "UPDATE CP_JobOrder_Master SET" _
                           & " dPaymentx = '" & CDate(txtField(8).Text) & "', " _
                           & " dModified = getdate() " _
                     & " WHERE sTransNox = '" & txtField(0).Tag & "' "
               oApp.Connection.Execute lsSQL, lnrow, adCmdText
            End If
            
         Next
         
            If lnrow <= 0 Then
               MsgBox "Unable to Update Job Order Master!!!" & vbCrLf & vbCrLf & _
               "Notify ROSALYN LAZO DESCALLAR for Assistance", vbCritical, "Warning"
               UpdateCP_JO_Detail = False
               GoTo endProc
            Else
               MsgBox "Record Successfully Updated!!!", vbInformation, Me.Caption
            End If
         
      End With

endProc:
   oApp.Connection.CommitTrans
   Exit Function
errProc:
   oApp.Connection.RollbackTrans
   UpdateCP_JO_Detail = False
   MsgBox Err.Description, vbCritical, "Warning"
End Function

Private Sub ShowGrid()
Dim lsSQL As String
Dim showdetail As New ADODB.Recordset
Dim Discount As Double
Dim Total As Double
Dim DiscAmount As Double

   lsSQL = "SELECT " _
               & " sTransNox, " _
               & " nEntryNox, " _
               & " sDescript, " _
               & " nPartsAmt, " _
               & " nLaborAmt, " _
               & " nDiscount, " _
               & " nQuantity, " _
               & " dModified  " _
         & " FROM CP_JobOrder_Detail " _
         & " WHERE sTransNox = '" & txtField(0).Tag & "'" _
         & " ORDER BY nEntryNox "
   
   Set showdetail = New ADODB.Recordset
   showdetail.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   
   With GridEditor1
      .Rows = showdetail.RecordCount + 1
      For lnCtr = 0 To showdetail.RecordCount - 1
         .TextMatrix(lnCtr + 1, 0) = showdetail("nEntryNox")
         .TextMatrix(lnCtr + 1, 1) = showdetail("sDescript")
         .TextMatrix(lnCtr + 1, 2) = Format(showdetail("nPartsAmt"), "#,##0.00")
         .TextMatrix(lnCtr + 1, 3) = Format(showdetail("nLaborAmt"), "#,##0.00")
         .TextMatrix(lnCtr + 1, 4) = showdetail("nDiscount")
         .TextMatrix(lnCtr + 1, 6) = Format(showdetail("nQuantity"), "#,##0")
         
         Discount = (showdetail("nDiscount") / 100)
         Total = (CDbl(showdetail("nPartsAmt")) + CDbl(showdetail("nLaborAmt"))) * _
                  showdetail("nQuantity")
         .TextMatrix(lnCtr + 1, 5) = Format((Total * Discount), "#,##0.00")
         .TextMatrix(lnCtr + 1, 7) = Format((Total - .TextMatrix(lnCtr + 1, 5)), "#,##0.00")
         
         showdetail.MoveNext
      Next
   End With
   
   Set showdetail = Nothing

End Sub
Private Sub EmptyGrid()
   With GridEditor1
      .Rows = 2
      For lnCtr = 1 To 7
         Select Case lnCtr
            Case 1
               .TextMatrix(1, lnCtr) = ""
            Case 2, 3, 5, 7
               .TextMatrix(1, lnCtr) = "0.00"
            Case 4, 6
               .TextMatrix(1, lnCtr) = 0
         End Select
      Next
   End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
   Case vbKeyReturn, vbKeyUp, vbKeyDown
      Select Case KeyCode
      Case vbKeyReturn, vbKeyDown
         If GetFocus = GridEditor1.hWnd Then Exit Sub
         SetNextFocus
      Case vbKeyUp
         SetPreviousFocus
      End Select
   End Select

End Sub

Private Sub GridEditor1_AddingRow(Cancel As Boolean)
   With GridEditor1
      If .TextMatrix(.Row, 1) = "" Then
         Cancel = True
      ElseIf .TextMatrix(.Row, 6) = "0.00" Or .TextMatrix(.Row, 6) = 0# Or _
         .TextMatrix(.Row, 6) = "" Then
         Cancel = True
      End If
   End With

End Sub

Private Sub GridEditor1_EditorValidate(Cancel As Boolean)
Dim Discount As Double
Dim Total As Double
Dim DiscAmount As Double
Dim Labor As Double
Dim Parts As Double

   With GridEditor1
      If .TextMatrix(.Row, 2) = "" Then
         .TextMatrix(.Row, 2) = "0.00"
      ElseIf .TextMatrix(.Row, 3) = "" Then
         .TextMatrix(.Row, 3) = "0.00"
      ElseIf .TextMatrix(.Row, 4) = "" Then
         .TextMatrix(.Row, 4) = 0
      ElseIf .TextMatrix(.Row, 6) = "" Then
         .TextMatrix(.Row, 6) = 1
      Else
         Discount = (.TextMatrix(.Row, 4) / 100)
         Total = (CDbl(.TextMatrix(.Row, 2)) + CDbl(.TextMatrix(.Row, 3))) * _
                  .TextMatrix(.Row, 6)
         .TextMatrix(.Row, 5) = Format((Total * Discount), "#,##0.00")
         .TextMatrix(.Row, 7) = Format((Total - .TextMatrix(.Row, 5)), "#,##0.00")
         
         For lnCtr = 1 To .Rows - 1
            If .TextMatrix(lnCtr, 3) <> "0.00" Then
               Labor = Labor + CDbl(.TextMatrix(lnCtr, 7))
            ElseIf .TextMatrix(lnCtr, 2) <> "0.00" Then
               Parts = Parts + CDbl(.TextMatrix(lnCtr, 7))
            End If
         Next
         txtField(4).Text = Format(Labor, "#,##0.00")
         txtField(5).Text = Format(Parts, "#,##0.00")
         txtField(7).Text = Format(CDbl(txtField(4).Text) + CDbl(txtField(5).Text) _
                              + CDbl(txtField(3).Text), "#,##0.00")
      End If
   End With
End Sub


Private Sub txtField_GotFocus(Index As Integer)
   txtfieldGotfocus = True
   pnindex = Index
   oDriver.ColumnIndex = Index
   txtField(Index).BackColor = &HE1FEFF

   If txtField(Index).Text <> "" Then
      txtField(Index).SelStart = 0
      txtField(Index).SelLength = Len(txtField(Index).Text)
   End If

End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

   If KeyCode = vbKeyF3 Or KeyCode = 13 Then
      If Index = 0 Or Index = 1 Then Search_JobOrder False
      If Index = 3 Then cmdButton(1).SetFocus
      If txtField(Index).Text <> "" Then SetNextFocus
      KeyCode = 0
   End If

End Sub

Private Sub txtField_LostFocus(Index As Integer)
   If Index = 3 Then
      If Not IsNumeric(txtField(Index).Text) Then txtField(Index).Text = 0#
         txtField(Index).Text = Format(txtField(Index).Text, "#,##0.00")
         txtField(7).Text = Format(CDbl(txtField(7).Text) _
                           + CDbl(txtField(Index).Text), "#,##0.00")
   
   End If
   txtField(Index).BackColor = &H80000005
End Sub

Private Sub txtfield_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 8
         If Not IsDate(txtField(Index).Text) Then
            txtField(Index).Text = Format(oApp.ServerDate, "MMMM dd,yyyy")
         Else
            txtField(Index).Text = Format(txtField(Index).Text, "MMMM dd,yyyy")
         End If
   End Select
   txtField(Index).Text = TitleCase(txtField(Index).Text)
End Sub
