VERSION 5.00
Object = "{0A7B56A6-35D0-4533-91DA-1715D3A0DD3E}#1.1#0"; "xrGridControl.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmChange_Unit 
   BorderStyle     =   0  'None
   Caption         =   "Change Unit"
   ClientHeight    =   6945
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12930
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6945
   ScaleWidth      =   12930
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame5 
      Height          =   720
      Left            =   1575
      Tag             =   "wt0;wb0"
      Top             =   5355
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   1270
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   525
         Index           =   5
         Left            =   1320
         MaxLength       =   30
         TabIndex        =   8
         Text            =   "Remarks"
         Top             =   75
         Width           =   4350
      End
      Begin VB.Label lblField 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   13
         Left            =   120
         TabIndex        =   7
         Top             =   90
         Width           =   1350
      End
   End
   Begin xrControl.xrFrame xrFrame4 
      Height          =   750
      Left            =   1575
      Tag             =   "wt0;wb0"
      Top             =   6090
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   1323
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtothers 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   510
         Index           =   1
         Left            =   1320
         TabIndex        =   16
         TabStop         =   0   'False
         Tag             =   "ht0;ft0"
         Text            =   "Cash Due"
         Top             =   105
         Width           =   3330
      End
      Begin VB.Image Image1 
         Height          =   525
         Left            =   4725
         Picture         =   "frmChange_Unit.frx":0000
         Stretch         =   -1  'True
         Top             =   120
         Width           =   930
      End
      Begin VB.Label lblField 
         BackStyle       =   0  'Transparent
         Caption         =   "Cash Due"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   12
         Left            =   105
         TabIndex        =   15
         Top             =   120
         Width           =   1620
      End
   End
   Begin xrControl.xrFrame xrFrame3 
      Height          =   1485
      Left            =   7380
      Tag             =   "wt0;wb0"
      Top             =   5355
      Width           =   5430
      _ExtentX        =   9578
      _ExtentY        =   2619
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtfield 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   510
         Index           =   3
         Left            =   1680
         TabIndex        =   10
         TabStop         =   0   'False
         Tag             =   "ht0;fb0"
         Text            =   "Total Amount"
         Top             =   90
         Width           =   3585
      End
      Begin VB.TextBox txtothers 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Index           =   0
         Left            =   1680
         TabIndex        =   14
         TabStop         =   0   'False
         Tag             =   "ht0;ft0"
         Text            =   "Change"
         Top             =   1005
         Width           =   3585
      End
      Begin VB.TextBox txtfield 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   4
         Left            =   1680
         TabIndex        =   12
         Tag             =   "ht0;fb0"
         Text            =   "Cash Given"
         Top             =   615
         Width           =   3585
      End
      Begin VB.Label lblField 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   285
         Index           =   100
         Left            =   165
         TabIndex        =   9
         Top             =   180
         Width           =   1725
      End
      Begin VB.Label lblField 
         BackStyle       =   0  'Transparent
         Caption         =   "Change Due"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   1
         Left            =   180
         TabIndex        =   13
         Top             =   1050
         Width           =   1620
      End
      Begin VB.Label lblField 
         BackStyle       =   0  'Transparent
         Caption         =   "Cash &Given"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   300
         Index           =   4
         Left            =   180
         TabIndex        =   11
         Top             =   630
         Width           =   1545
      End
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   2955
      Left            =   1575
      Tag             =   "wt0;fb0"
      Top             =   2385
      Width           =   11235
      _ExtentX        =   19817
      _ExtentY        =   5212
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin xrGridEditor.GridEditor GridEditor1 
         Height          =   2790
         Left            =   60
         TabIndex        =   6
         Tag             =   "et0;eb0;et0;bc2"
         Top             =   60
         Width           =   11085
         _ExtentX        =   19553
         _ExtentY        =   4921
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
         Object.HEIGHT          =   2790
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
         MOUSEICON       =   "frmChange_Unit.frx":076A
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
      Height          =   540
      Index           =   0
      Left            =   1575
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   11235
      _ExtentX        =   19817
      _ExtentY        =   953
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   0
         Left            =   8880
         TabIndex        =   5
         Text            =   "Transaction No."
         Top             =   120
         Width           =   2220
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   1
         Left            =   1005
         TabIndex        =   1
         Text            =   "Invoice No."
         Top             =   120
         Width           =   2220
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   2
         Left            =   4575
         TabIndex        =   3
         Text            =   "Transaction Date"
         Top             =   120
         Width           =   2220
      End
      Begin VB.Label lblField 
         BackStyle       =   0  'Transparent
         Caption         =   "Tran. No."
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   0
         Left            =   8100
         TabIndex        =   4
         Top             =   120
         Width           =   1035
      End
      Begin VB.Label lblField 
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice No."
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   7
         Left            =   105
         TabIndex        =   0
         Top             =   135
         Width           =   1350
      End
      Begin VB.Label lblField 
         BackStyle       =   0  'Transparent
         Caption         =   "Tran. Date"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   3
         Left            =   3615
         TabIndex        =   2
         Top             =   120
         Width           =   1005
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   3
      Left            =   90
      TabIndex        =   23
      Top             =   4935
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
      Picture         =   "frmChange_Unit.frx":0786
      CaptionAlign    =   0
      BackColor       =   14286077
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   405
      Index           =   2
      Left            =   90
      TabIndex        =   22
      Top             =   4095
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
      Picture         =   "frmChange_Unit.frx":0F00
      CaptionAlign    =   0
      BackColor       =   14286077
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   1
      Left            =   90
      TabIndex        =   20
      Top             =   4515
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
      Caption         =   "&Replace"
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
      Picture         =   "frmChange_Unit.frx":167A
      CaptionAlign    =   0
      BackColor       =   14286077
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   4
      Left            =   90
      TabIndex        =   21
      Top             =   4515
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
      Picture         =   "frmChange_Unit.frx":1DF4
      CaptionAlign    =   0
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1200
      Index           =   1
      Left            =   1575
      Tag             =   "wt0;fb0"
      Top             =   1125
      Width           =   11235
      _ExtentX        =   19817
      _ExtentY        =   2117
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice No. :"
         Height          =   195
         Index           =   11
         Left            =   7470
         TabIndex        =   41
         Top             =   75
         Width           =   915
      End
      Begin VB.Label lblFields 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXXXXXXXXXX"
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
         Left            =   8490
         TabIndex        =   40
         Tag             =   "tc0"
         Top             =   75
         Width           =   2340
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Total :"
         Height          =   195
         Index           =   10
         Left            =   7050
         TabIndex        =   39
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblFields 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "P0.00"
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
         Left            =   8490
         TabIndex        =   38
         Tag             =   "tc0"
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount Paid :"
         Height          =   195
         Index           =   9
         Left            =   7395
         TabIndex        =   37
         Top             =   585
         Width           =   1020
      End
      Begin VB.Label lblFields 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "P0.00"
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
         Left            =   8490
         TabIndex        =   36
         Tag             =   "tc0"
         Top             =   585
         Width           =   1575
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mode of Payment :"
         Height          =   195
         Index           =   6
         Left            =   7050
         TabIndex        =   35
         Top             =   330
         Width           =   1335
      End
      Begin VB.Label lblFields 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXXXXXXXXXX"
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
         Left            =   8490
         TabIndex        =   34
         Tag             =   "tc0"
         Top             =   330
         Width           =   1575
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction No :"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   33
         Top             =   75
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
         Index           =   0
         Left            =   1485
         TabIndex        =   32
         Tag             =   "tc0"
         Top             =   75
         Width           =   3975
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name :"
         Height          =   195
         Index           =   8
         Left            =   210
         TabIndex        =   31
         Top             =   330
         Width           =   1215
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Person :"
         Height          =   195
         Index           =   5
         Left            =   405
         TabIndex        =   30
         Top             =   840
         Width           =   1020
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Date :"
         Height          =   195
         Index           =   28
         Left            =   105
         TabIndex        =   29
         Top             =   585
         Width           =   1320
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   30
         Left            =   2280
         TabIndex        =   28
         Top             =   690
         Width           =   30
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
         Left            =   1500
         TabIndex        =   27
         Tag             =   "tc0"
         Top             =   330
         Width           =   4005
      End
      Begin VB.Label lblFields 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
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
         Height          =   270
         Index           =   2
         Left            =   1500
         TabIndex        =   26
         Tag             =   "tc0"
         Top             =   840
         Width           =   4185
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
         Index           =   3
         Left            =   1485
         TabIndex        =   25
         Tag             =   "tc0"
         Top             =   585
         Width           =   4455
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   5
      Left            =   90
      TabIndex        =   24
      Top             =   4935
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
      Picture         =   "frmChange_Unit.frx":256E
      CaptionAlign    =   0
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   7
      Left            =   90
      TabIndex        =   18
      ToolTipText     =   "Cheque Payment"
      Top             =   3675
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
      Caption         =   "Cheque"
      AccessKey       =   "Cheque"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmChange_Unit.frx":2CE8
      CaptionAlign    =   0
      BackColor       =   14286077
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   6
      Left            =   75
      TabIndex        =   17
      ToolTipText     =   "Log Out"
      Top             =   3255
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
      Caption         =   "Card"
      AccessKey       =   "Card"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmChange_Unit.frx":3462
      CaptionAlign    =   0
      BackColor       =   14286077
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   8
      Left            =   90
      TabIndex        =   19
      ToolTipText     =   "Credit Card Payment"
      Top             =   4095
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
      Caption         =   "Installment"
      AccessKey       =   "Installment"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmChange_Unit.frx":3BDC
      CaptionAlign    =   0
      BackColor       =   14286077
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
End
Attribute VB_Name = "frmChange_Unit"
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
Dim pbnewitem As Boolean
Dim psSelected() As String
Dim lsSQL As String
Dim pnindex As Integer
Dim pnctr As Integer
Dim lrs As New ADODB.Recordset
Dim lsSearch As String
Dim TranStat As Integer
Dim ClientID As String
Dim Cashier As String
Dim Time As String

Dim psPayment As String

Property Let Payment(Payment As String)
   psPayment = Payment
End Property

Private Sub cmdButton_Click(Index As Integer)
Dim temp As Integer

With GridEditor1
   Select Case Index
      Case 1 'Replace
            If txtfield(1).Text <> "" Then
               oDriver.RecordNew
               txtfield(3).Text = lblFields(7).Caption
               InitButton xeModeReady
               oDriver.ShowButton 6
               oDriver.ShowButton 7
               oDriver.ShowButton 8
            Else
               MsgBox "No Active Transaction!!!", vbInformation, "Information"
            End If
      Case 2 'Search
            SearchTrans
      Case 3 'Close
            Unload Me
      Case 4 'Save
         For pnctr = 1 To .Rows - 1
           If .TextMatrix(pnctr, 15) = "Yes" Then
              temp = temp + 1
           End If
         Next
         If temp = 0 Then
            MsgBox "No Item to be Replaced!!!" & vbCrLf & _
            "Please Mark the Item then Try Again." & vbCrLf & _
            "" & vbCrLf & _
            "Save Unsuccessful!!!", vbInformation, "Notice"
         Else
            oDriver.RecordSave
         End If
      Case 5 'Cancel
            ClearFields
            EmptyGrid
            InitButton xeModeAddNew
            oDriver.HideButton 6
            oDriver.HideButton 7
            oDriver.HideButton 8

      Case 6 'Card Transaction
         If txtothers(1) <> 0# Then
            frmCard_POS.txtfield(1) = Format(txtothers(1), "#,##0.00")
            frmCard_POS.Transaction = "Change"
            frmCard_POS.Show 1
         End If
      Case 7 'Cheque Transaction
         If txtothers(1) <> 0# Then
            frmCheque_POS.txtfield(1) = Format(txtothers(1), "#,##0.00")
            frmCheque_POS.Transaction = "Change"
            frmCheque_POS.Show 1
         End If
      Case 8 'Installment Transaction
         If txtothers(1) <> 0# Then
            frmInstallment_POS.txtfield(1) = Format(txtothers(1), "#,##0.00")
            frmInstallment_POS.Transaction = "Change"
            frmInstallment_POS.Show 1
         End If

      End Select
End With
End Sub

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0

   If bLoaded = False Then
      oDriver.RecordNew
      bLoaded = True
      txtfield(1).SetFocus
      oDriver.HideButton 6
      oDriver.HideButton 7
      oDriver.HideButton 8
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn, vbKeyDown
         If GetFocus = GridEditor1.hWnd Then Exit Sub
         SetNextFocus
      Case vbKeyUp
         SetPreviousFocus
      Case 17
         If txtothers(1).Text <> 0# Then
            txtfield(4).SetFocus
         End If
      Case 27
         Call Modified("CP_SO_Master", "sTransNox = '" & oDriver.FieldValue(0) & "' ")
   End Select
End Sub

Private Sub Form_Load()
Dim lsSQL As String
Dim lnctr As Integer
   
   CenterChildForm mdiMain, Me
   bLoaded = False
   
   Set oDriver = New FormDriver
   Set oDriver.AppDriver = oApp
   Set oDriver.MainForm = Me
   
   InitButton xeModeAddNew
   
   Set oSkin = New FormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransaction
   
   oDriver.RecQuery = "SELECT" _
                  & " sTransNox, " _
                  & " sSalesInv, " _
                  & " dTransact, " _
                  & " nTranTotl, " _
                  & " nAmtPaidx, " _
                  & " sRemarksx, " _
                  & " sClientID, " _
                  & " sCashierx, " _
                  & " nGiftCpnx, " _
                  & " cTranStat, " _
                  & " sModified, " _
                  & " dModified, " _
                  & " vTimeStmp  " _
            & " FROM CP_SO_Master " _

   oDriver.InitRecForm

   oDriver.FieldStart = 1
   oDriver.FieldFormat(2) = "MMMM DD, YYYY"

   InitGrid
   EmptyGrid
   ClearFields

End Sub

Private Sub InitButton(lnStat As Integer)
Dim lbShow As Boolean
   lbShow = IIf(lnStat = xeModeReady, False, True)
   cmdButton(5).Visible = Not lbShow
   cmdButton(4).Visible = Not lbShow
   xrFrame2.Enabled = Not lbShow
   xrFrame5.Enabled = Not lbShow
   xrFrame1(1).Enabled = Not lbShow
   
   cmdButton(2).Visible = lbShow
   cmdButton(1).Visible = lbShow
   cmdButton(3).Visible = lbShow
End Sub

Private Sub ClearFields()
Dim lnctr As Integer

For pnctr = 0 To 7
   lblFields(pnctr).Caption = ""
Next

For lnctr = 1 To 4
   Select Case lnctr
   Case 1 To 2
      txtfield(lnctr).Text = ""
      txtfield(lnctr).Enabled = False
      If lnctr = 1 Then txtfield(lnctr).Enabled = True
      If lnctr = 2 Then txtfield(lnctr).Text = Format(Date, "MMMM dd, yyyy")
   Case 3 To 4
      txtfield(lnctr).Text = "0.00"
   End Select
Next
txtfield(5).Text = ""
txtfield(5).Enabled = True
txtothers(0).Text = "0.00"
txtothers(1).Text = "0.00"
psPayment = ""
pbnewitem = True

End Sub

Private Sub InitGrid()

   With GridEditor1
      .Rows = 2
      .Cols = 16
      .Font = "Arial"
      
      'column title
      .TextMatrix(0, 1) = "Bar Code"
      .TextMatrix(0, 2) = "Description"
      .TextMatrix(0, 3) = "Unit Price"
      .TextMatrix(0, 4) = "Qty"
      .TextMatrix(0, 5) = "%"
      .TextMatrix(0, 6) = "Amt"
      .TextMatrix(0, 7) = "Stock ID"
      .TextMatrix(0, 8) = "IMEI No. / Cell #"
      .TextMatrix(0, 9) = "Serial ID. / Ref. #"
      .TextMatrix(0, 10) = "Sub Total"
      .TextMatrix(0, 11) = "Pur Price"
      .TextMatrix(0, 12) = "Category"
      .TextMatrix(0, 13) = "Old Serial"
      .TextMatrix(0, 14) = "Old StockID"
      .TextMatrix(0, 15) = "Replace"
      
      'column width
      .ColWidth(0) = 250
      .ColWidth(1) = 1700
      .ColWidth(2) = 2600
      .ColWidth(3) = 1100
      .ColWidth(4) = 400
      .ColWidth(5) = 350
      .ColWidth(6) = 700
      .ColWidth(7) = 0
      .ColWidth(8) = 2100
      .ColWidth(9) = 0
      .ColWidth(10) = 1100
      .ColWidth(11) = 0
      .ColWidth(12) = 0
      .ColWidth(13) = 0
      .ColWidth(14) = 0
      .ColWidth(15) = 680

      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 6
      .ColAlignment(4) = 6
      .ColAlignment(5) = 6
      .ColAlignment(6) = 6
      .ColAlignment(7) = 6
      .ColAlignment(8) = 1
      .ColAlignment(9) = 1
      .ColAlignment(10) = 6
      
      .ColEnabled(2) = False
      .ColEnabled(3) = False
      .ColEnabled(6) = False
      .ColEnabled(7) = False
      .ColEnabled(9) = False
      .ColEnabled(10) = False
      .ColEnabled(11) = False
      .ColEnabled(12) = False
      .ColEnabled(13) = False
      .ColEnabled(14) = False
      
      .ColNumberOnly(3) = True
      .ColNumberOnly(4) = True
      .ColNumberOnly(5) = True
      .ColNumberOnly(6) = True
      
      .ColFormat(3) = "#,##0.00"
      .ColFormat(6) = "#,##0.00"
      .ColFormat(10) = "#,##0.00"
      
      .Row = 1
   End With

End Sub

Private Sub EmptyGrid()
Dim lnctr As Integer

With GridEditor1
   .Rows = 2
   For lnctr = 1 To .Cols - 1
      .TextMatrix(1, lnctr) = ""
   Next
End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oDriver = Nothing
   Set oSkin = Nothing
End Sub

Private Sub GridEditor1_AddingRow(Cancel As Boolean)
   Cancel = True
End Sub

Private Sub GridEditor1_EditorValidate(Cancel As Boolean)
Dim temp As Double

With GridEditor1
   Select Case .Col
      Case 4
         If .TextMatrix(.Row, .Col) = 0 Or .TextMatrix(.Row, .Col) = "" Then
            MsgBox "Invalid Quantity!!!", vbCritical, "Warning"
            .TextMatrix(.Row, .Col) = 1
            .Col = .Col - 1
         End If
      Case 5
         If .TextMatrix(.Row, .Col) = "" Then
            MsgBox "Invalid Input!!!", vbCritical, "Warning"
            .TextMatrix(.Row, .Col) = 0#
            .Col = .Col - 1
         Else
            .TextMatrix(.Row, 6) = .TextMatrix(.Row, 3) * (.TextMatrix(.Row, .Col) / 100)
         End If
      Case 6
         If .TextMatrix(.Row, .Col) = "" Then
            MsgBox "Invalid Input!!!", vbCritical, "Warning"
            .TextMatrix(.Row, .Col) = 0#
            .Col = .Col - 1
         End If
      Case 10
         .TextMatrix(.Row, .Col) = (.TextMatrix(.Row, 3) * .TextMatrix(.Row, 4)) - .TextMatrix(.Row, 6)
      Case 15
         Select Case Trim(LCase(.TextMatrix(.Row, .Col)))
            Case "y", "ye", "yes"
               .TextMatrix(.Row, .Col) = "Yes"
            Case Else
               .TextMatrix(.Row, .Col) = ""
         End Select
      End Select
   Grand_Total
End With
End Sub

Private Sub GridEditor1_KeyDown(KeyCode As Integer, Shift As Integer)
   With GridEditor1
      If KeyCode = vbKeyF3 Or KeyCode = 13 Then
         If .Col = 1 Then
            SearchBarCode
         ElseIf .Col = 8 Then
            SearchIMEI False
         End If
      End If
   End With
End Sub
Private Sub SearchBarCode()
Dim lsSQL As String
Dim lsCondition As String
Dim lsSearch As String
      
   With GridEditor1
         lsSQL = "SELECT" _
                & " a.sBarrcode, " _
                & " a.sStockIDx, " _
                & " b.sBrandNme, " _
                & " c.sModelNme, " _
                & " a.sDescript, " _
                & " d.sColorNme, " _
                & " a.nSelPrice, " _
                & " a.cWdSerial, " _
                & " a.nPurPrice  " _
            & " FROM CP_Inventory a " _
                & " LEFT JOIN Brand b " _
                  & " ON a.sBrandIdx = b.sBrandIdx " _
                & " LEFT JOIN Model c " _
                  & " ON a.sModelIdx = c.sModelIdx " _
                & " LEFT JOIN Color d " _
                  & " ON a.sColorIDx = d.sColorIDx " _
            & " WHERE a.sBarrcode like  '" & .TextMatrix(.Row, 1) & "%' " _
               & " AND cCellLoad = 0 " _
               & " AND cWalletxx = 0 " _
               & " AND cCellCard = 0 "
            If oRS.State = adStateOpen Then oRS.Close
            oRS.Open lsSQL, oApp.Connection, adOpenKeyset, adLockOptimistic, adCmdText

         If Not oRS.EOF Then
            If oRS.RecordCount = 1 Then
               .TextMatrix(.Row, 1) = IIf(IsNull(oRS(0)), "", oRS(0))
               .TextMatrix(.Row, 2) = Trim(IIf(IsNull(oRS(2)), "", oRS(2)) & " " & _
                                       IIf(IsNull(oRS(3)), "", oRS(3)) & " " & _
                                       IIf(IsNull(oRS(4)), "", oRS(4)) & " " & _
                                       IIf(IsNull(oRS(5)), "", oRS(5)))
               .TextMatrix(.Row, 3) = Format(oRS(6), "#,##0.00")
               .TextMatrix(.Row, 4) = 1
               .TextMatrix(.Row, 5) = "0"
               .TextMatrix(.Row, 6) = "0.00"
               .TextMatrix(.Row, 7) = oRS(1)
               .TextMatrix(.Row, 8) = ""
               .TextMatrix(.Row, 9) = ""
               .TextMatrix(.Row, 10) = Format(oRS(6), "#,##0.00")
               .TextMatrix(.Row, 11) = oRS(8)
               .TextMatrix(.Row, 12) = oRS(7)
            Else
               lsSearch = KwikSearch(oApp, lsSQL, _
                          "sBarrcode»sBrandNme»sModelNme»sDescript»sColorNme", _
                          "Bar Code»Brand»Model»Description»Color")
               If lsSearch <> "" Then
                  psSelected = Split(lsSearch, "»")
                  .TextMatrix(.Row, 1) = IIf(IsNull(psSelected(0)), "", psSelected(0))
                  .TextMatrix(.Row, 2) = Trim(IIf(IsNull(psSelected(2)), "", psSelected(2)) & " " & _
                                          IIf(IsNull(psSelected(3)), "", psSelected(3)) & " " & _
                                          IIf(IsNull(psSelected(4)), "", psSelected(4)) & " " & _
                                          IIf(IsNull(psSelected(5)), "", psSelected(5)))
                  .TextMatrix(.Row, 3) = Format(psSelected(6), "#,##0.00")
                  .TextMatrix(.Row, 4) = 1
                  .TextMatrix(.Row, 5) = "0"
                  .TextMatrix(.Row, 6) = "0.00"
                  .TextMatrix(.Row, 7) = psSelected(1)
                  .TextMatrix(.Row, 8) = ""
                  .TextMatrix(.Row, 9) = ""
                  .TextMatrix(.Row, 10) = Format(psSelected(6), "#,##0.00")
                  .TextMatrix(.Row, 11) = psSelected(8)
                  .TextMatrix(.Row, 12) = psSelected(7)
                  .Col = 3
               End If
            End If
            .SetFocus
            .Refresh
         Else
            MsgBox "No Record Found!!!", vbCritical, "Warning"
         End If
      Set oRS = Nothing
   End With
   
End Sub

Private Sub SearchIMEI(ByVal SearchValue As Boolean)
Dim lsSearch As String
Dim lrs As ADODB.Recordset
Dim lsSQL As String
   
   
With GridEditor1
   Set lrs = New ADODB.Recordset
   lsSQL = "SELECT" _
            & " a.sIMEINoxx, " _
            & " a.sSerialID, " _
            & " a.sStockIDx, " _
            & " b.sBarrCode  " _
         & " FROM CP_Serial_Master a " _
            & " LEFT JOIN CP_Inventory b " _
               & " ON a.sStockIDx = b.sStockIDx " _
         & " WHERE a.sStockIDx = '" & .TextMatrix(.Row, 7) & "' " _
         & " AND cRecdStat = 1 " _
         & " AND cSoldStat = 0 " _
         & " AND cLocation = 1 "
   
   If SearchValue Then
      lsSQL = lsSQL & " AND a.sIMEINoxx = '" & .TextMatrix(.Row, 8) & "'"
   Else
      lsSQL = lsSQL & " AND a.sIMEINoxx LIKE '%" & .TextMatrix(.Row, 8) & "%' "
   End If
            
   lsSQL = lsSQL & " ORDER BY sIMEINoxx"
   lrs.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   
   If lrs.RecordCount = 1 Then
      .TextMatrix(.Row, 8) = lrs(0)
      .TextMatrix(.Row, 9) = lrs(1)
      
   ElseIf lrs.RecordCount > 1 Then
        lsSearch = KwikBrowse(oApp, lrs, _
                          "sIMEINoxx", _
                          "IMEI No.")
        If lsSearch <> "" Then
            psSelected = Split(lsSearch, "»")
            .TextMatrix(.Row, 8) = psSelected(0)
            .TextMatrix(.Row, 9) = psSelected(1)
        End If
   Else
      MsgBox "IMEI No. Not Existing!!!", vbCritical, "Warning"
   End If
   
   Set lrs = Nothing
End With

End Sub

Private Sub GridEditor1_RowColChange()
Dim lnctr As Integer

With GridEditor1
   If .TextMatrix(.Row, 12) = 2 Then
      For lnctr = 1 To .Cols - 1
         .ColEnabled(lnctr) = False
      Next
   Else
      .ColEnabled(1) = True
      .ColEnabled(4) = True
      .ColEnabled(5) = True
      .ColEnabled(6) = True
      .ColEnabled(8) = True
      If .TextMatrix(.Row, 12) = 1 Then
         .ColEnabled(4) = False
      End If
   End If
   .TextMatrix(.Row, 6) = CDbl(.TextMatrix(.Row, 3)) * CDbl((.TextMatrix(.Row, 5) / 100))
   .TextMatrix(.Row, 10) = (CDbl(.TextMatrix(.Row, 3)) * CDbl(.TextMatrix(.Row, 4))) - CDbl(.TextMatrix(.Row, 6))
   Grand_Total
End With
End Sub

Private Sub GridEditor1_Validate(Cancel As Boolean)
Dim lnctr As Integer

With GridEditor1
   Select Case .Col
      Case 4
         If .TextMatrix(.Row, .Col) = 0 Or .TextMatrix(.Row, .Col) = "" Then
            MsgBox "Invalid Quantity!!!", vbCritical, "Warning"
            .TextMatrix(.Row, .Col) = 1
            .Col = .Col - 1
         Else
         End If
      Case 5
         If .TextMatrix(.Row, .Col) = "" Then
            MsgBox "Invalid Input!!!", vbCritical, "Warning"
            .TextMatrix(.Row, .Col) = 0#
            .Col = .Col - 1
         End If
      Case 6
         If .TextMatrix(.Row, .Col) = "" Then
            MsgBox "Invalid Input!!!", vbCritical, "Warning"
            .TextMatrix(.Row, .Col) = 0#
            .Col = .Col - 1
         End If
      Case 8
         If .TextMatrix(.Row, 12) = 1 Then
            If .TextMatrix(.Row, .Col) = "" Then
               MsgBox "Invalid IMEI No!!!", vbCritical, "Warning"
            End If
         End If

      Case 15
         Select Case Trim(LCase(.TextMatrix(.Row, .Col)))
            Case "y", "ye", "yes"
               .TextMatrix(.Row, .Col) = "Yes"
            Case Else
               .TextMatrix(.Row, .Col) = ""
         End Select
   End Select
End With
End Sub

Private Sub oDriver_DisableOtherControl()
   oDriver.DisableTextbox 0
   xrFrame3.Enabled = False
   xrFrame4.Enabled = False
End Sub

Private Sub oDriver_EnableOtherControl()
   oDriver.DisableTextbox 0
   xrFrame3.Enabled = False
   xrFrame4.Enabled = False
End Sub

Private Sub oDriver_InitValue()
   oDriver.FieldReference(0) = True
   If Not oDriver.SetValue(0, getNextCode("CP_SO_Master", "sTransNox", True, oApp.Connection, True, oApp.BranchCode)) Then Exit Sub
   oDriver.DisableTextbox 0
   oDriver.FieldValue(2) = Date
   psPayment = ""
End Sub

Private Sub oDriver_SaveComplete()
   ClearFields
   InitButton xeModeAddNew
   EmptyGrid
   oDriver.HideButton 6
   oDriver.HideButton 7
   oDriver.HideButton 8
   MsgBox "Transaction Successfully Saved!!!", vbInformation, "Information"
End Sub

Private Sub oDriver_WillSave(Cancel As Boolean)

   If oDriver.FieldValue(0) = "" Then
      MsgBox "Invalid Transaction No. Detected!!!", vbCritical, "Warning"
      Cancel = True
   ElseIf oDriver.FieldValue(1) = "" Then
      MsgBox "Invalid Invoice No. Detected!!!", vbCritical, "Warning"
      txtfield(1).SetFocus
      Cancel = True
   ElseIf oDriver.FieldValue(2) = "" Then
      MsgBox "Invalid Date Detected!!!", vbCritical, "Warning"
      txtfield(2).SetFocus
      Cancel = True
   ElseIf oDriver.FieldValue(3) = "" Then
      MsgBox "Invalid Amount Detected!!!", vbCritical, "Warning"
      txtfield(3).SetFocus
      Cancel = True
   ElseIf oDriver.FieldValue(4) = "" Then
      MsgBox "Invalid Amount Detected!!!", vbCritical, "Warning"
      txtfield(4).SetFocus
      Cancel = True
   ElseIf CDbl(txtothers(1).Text) > CDbl(txtfield(4).Text) And psPayment = "" Then
      MsgBox "Amount Can't be Less than Cash Due!!!", vbCritical, "Warning"
      txtfield(4).SetFocus
      Cancel = True
   Else
      Time = Format(Now, "hh:nn:ss AM/PM")
      Cancel = Not CP_SODetail
         If Cancel Then Exit Sub
      Cancel = Not Serial_Returned
         If Cancel Then Exit Sub
      Cancel = Not Serial_Replacement
         If Cancel Then Exit Sub
      Cancel = Not Inventory_Returned
         If Cancel Then Exit Sub
      Cancel = Not Inventory_Replacement
         If Cancel Then Exit Sub
      
      oDriver.FieldValue(2) = CDate(txtfield(2).Text) & " " & Time
      oDriver.FieldValue(4) = CDbl(txtothers(1).Text)
      oDriver.FieldValue(6) = ClientID 'sClientID
      oDriver.FieldValue(7) = Cashier 'sCashierx
      oDriver.FieldValue(8) = 0 'nGiftCpnx
      
      'Mode of Payment
      Select Case psPayment
         Case "Credit"
            Cancel = Not SaveCP_SOCredit
               If Cancel Then Exit Sub
            oDriver.FieldValue(9) = 1
            oDriver.FieldValue(3) = CDbl(frmCard_POS.txtfield(2))
            Unload frmCard_POS
         Case "Cheque"
            Cancel = Not SaveCP_SOCheque
               If Cancel Then Exit Sub
            oDriver.FieldValue(9) = 2
            oDriver.FieldValue(3) = CDbl(frmCheque_POS.txtfield(2))
            Unload frmCheque_POS
         Case "Installment"
            Cancel = Not SaveCP_SOInstallment
               If Cancel Then Exit Sub
            oDriver.FieldValue(9) = 3
            oDriver.FieldValue(3) = CDbl(frmInstallment_POS.txtfield(5))
            Unload frmInstallment_POS
         Case Else
            oDriver.FieldValue(9) = 0
            oDriver.FieldValue(3) = CDbl(txtfield(3).Text)
      End Select
   End If
   
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   oDriver.ColumnIndex = Index
   txtfieldGotfocus = True
   pnindex = Index
   txtfield(Index).BackColor = &HE1FEFF
End Sub

Private Sub Grand_Total()
Dim Total As Double
Dim lnctr As Integer

With GridEditor1
   For lnctr = 1 To .Rows - 1
      Total = Total + CDbl(.TextMatrix(lnctr, 10))
   Next
   txtfield(3).Text = Format(CDbl(Total), "#,##0.00")
End With
txtothers(1).Text = Format((CDbl(txtfield(3).Text) - CDbl(lblFields(7).Caption)), "#,##0.00")
End Sub


Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Or KeyCode = vbKeyF3 Then
      Select Case Index
         Case 1  'Search Transaction
            If pbnewitem = True Then SearchTrans
      End Select
      If txtfield(Index).Text <> "" Then SetNextFocus
      KeyCode = 0
   End If
End Sub

Private Sub ShowGrid()
Dim lsSQL As String
Dim lnctr As Integer

   'Show Detail
   lsSQL = "SELECT" _
         & " Distinct " _
         & " a.nEntryNox, " _
         & " a.nQuantity, " _
         & " a.nUnitPrce, " _
         & " a.nPurPrice, " _
         & " a.nDiscount, " _
         & " a.nDiscAmnt, " _
         & " a.nSubTotal, " _
         & " b.sSerialID, " _
         & " c.sIMEINoxx, " _
         & " d.sStockIDx, " _
         & " d.sBarrCode, " _
         & " d.sDescript, " _
         & " d.cWdSerial, " _
         & " e.sPhoneNum, " _
         & " e.sReferNox, " _
         & " f.sBrandNme, " _
         & " g.sModelNme, " _
         & " h.sColorNme, " _
         & " d.cWalletxx, " _
         & " d.cCellLoad  "

   lsSQL = lsSQL _
         & " FROM CP_SO_Detail a " _
            & " LEFT JOIN CP_SO_Serial b " _
               & " ON a.sTransNox = b.sTransNox " _
               & " AND a.nEntryNox = b.nEntryNox " _
            & " LEFT JOIN CP_Serial_Master c " _
               & " ON b.sSerialID = c.sSerialID " _
            & " LEFT JOIN CP_Inventory d " _
               & " ON a.sStockIDx = d.sStockIDx " _
            & " LEFT join ELoad_Ledger e " _
               & " ON a.sTransnox = e.ssourceno " _
                  & " AND a.nEntryNox = e.sTransNox " _
            & " LEFT JOIN Brand f " _
               & " ON d.sBrandIdx = f.sBrandIDx " _
            & " LEFT JOIN Model g " _
               & " ON d.sModelIDx = g.sModelIDx " _
            & " LEFT JOIN Color h " _
               & " ON d.sColorIDx = h.sColorIDx " _
         & " WHERE a.sTransNox = '" & txtfield(0).Text & "' " _
         & " ORDER by a.nEntryNox " _

   If oRS.State = adStateOpen Then oRS.Close
   oRS.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   With GridEditor1
      If oRS.RecordCount <> 0 Then
         .Rows = oRS.RecordCount + 1
         For lnctr = 0 To oRS.RecordCount - 1
            .TextMatrix(lnctr + 1, 0) = oRS("nEntryNox")
            .TextMatrix(lnctr + 1, 1) = oRS("sBarrCode")
            .TextMatrix(lnctr + 1, 2) = Trim(IIf(IsNull(oRS("sBrandNme")), "", oRS("sBrandNme")) _
                                    & " " & IIf(IsNull(oRS("sModelNme")), "", oRS("sModelNme")) _
                                    & " " & IIf(IsNull(oRS("sDescript")), "", oRS("sDescript")) _
                                    & " " & IIf(IsNull(oRS("sColorNme")), "", oRS("sColorNme")))
            .TextMatrix(lnctr + 1, 3) = Format(oRS("nUnitPrce"), "#,##0.00")
            .TextMatrix(lnctr + 1, 4) = oRS("nQuantity")
            .TextMatrix(lnctr + 1, 5) = oRS("nDiscount")
            .TextMatrix(lnctr + 1, 6) = oRS("nDiscAmnt")
            .TextMatrix(lnctr + 1, 7) = oRS("sStockIDx")
            .TextMatrix(lnctr + 1, 10) = Format(oRS("nSubTotal"), "#,##0.00")
            .TextMatrix(lnctr + 1, 11) = Format(oRS("nPurPrice"), "#,##0.00")
            .TextMatrix(lnctr + 1, 12) = oRS("cWdSerial")
            .TextMatrix(lnctr + 1, 13) = oRS("sSerialID")
            .TextMatrix(lnctr + 1, 14) = oRS("sStockIDx")
            Select Case oRS("cWdSerial")
            Case 1                                                'Units and Microphone
               .TextMatrix(lnctr + 1, 8) = IIf(IsNull(oRS("sIMEINoxx")), "", oRS("sIMEINoxx"))
               .TextMatrix(lnctr + 1, 9) = IIf(IsNull(oRS("sSerialID")), "", oRS("sSerialID"))
               .TextMatrix(lnctr + 1, 13) = IIf(IsNull(oRS("sSerialID")), "", oRS("sSerialID"))
            Case 0
               If oRS("cWalletxx") = 1 Or oRS("cCellLoad") = 1 Then 'Load Retail/ Wallet
                  .TextMatrix(lnctr + 1, 8) = IIf(IsNull(oRS("sPhoneNum")), "", oRS("sPhoneNum"))
                  .TextMatrix(lnctr + 1, 9) = IIf(IsNull(oRS("sReferNox")), "", oRS("sReferNox"))
               Else                                               'Accessories
                  .TextMatrix(lnctr + 1, 8) = ""
                  .TextMatrix(lnctr + 1, 9) = ""
                  .TextMatrix(lnctr + 1, 13) = ""
               End If
            End Select
            oRS.MoveNext
         Next
      Else
         .Rows = 2
      End If
   End With
   Set oRS = Nothing

End Sub

Private Sub SearchTrans()

   'Show Master
   lsSQL = "SELECT" _
            & " a.sTransNox, " _
            & " a.sSalesInv, " _
            & " a.nTranTotl, " _
            & " a.nAmtPaidx, " _
            & " a.sCashierx, " _
            & " a.dTransact, " _
            & " b.sLastName + ' , ' + b.sFrstName + ' ' + b.sMiddName as xFullName," _
            & " a.cTranStat, " _
            & " c.sLastName + ' , ' + c.sFrstName + ' ' + c.sMiddName as xSalesPer," _
            & " b.sClientID, " _
            & " a.sRemarksx  " _
         & " FROM CP_SO_Master a " _
            & " LEFT JOIN Client_Master b " _
               & " ON a.sClientID = b.sClientID " _
            & " LEFT JOIN Sales_Person c " _
               & " ON a.sCashierx = c.sEmployID " _
         & " WHERE a.sSalesInv like '%" & txtfield(1).Text & "%' " _
         & " ORDER BY a.sSalesInv, a.dTransact"
         
   If lrs.State = adStateOpen Then lrs.Close
   lrs.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
            
      If lrs.RecordCount = 1 Then
         txtfield(0).Text = lrs("sTransNox")
         txtfield(1).Text = lrs("sSalesInv")
         txtfield(2).Text = Format(lrs("dTransact"), "MMMM dd, yyyy")
         txtfield(3).Text = Format(lrs("nTranTotl"), "#,##0.00")
         txtfield(4).Text = Format(lrs("nAmtPaidx"), "#,##0.00")
         txtfield(5).Text = IIf(IsNull(lrs("sRemarksx")), "", lrs("sRemarksx"))
         lblFields(0).Caption = lrs("sTransNox")
         lblFields(1).Caption = lrs("xFullName")
         lblFields(2).Caption = lrs("xSalesPer")
         lblFields(3).Caption = Format(lrs("dTransact"), "MMMM dd, yyyy")
         lblFields(4).Caption = lrs("sSalesInv")
         ClientID = lrs(9)
         Cashier = lrs(4)
         
         TranStat = lrs("cTranStat")
         TransactionType
         lblFields(6).Caption = Format(lrs("nAmtPaidx"), "#,##0.00")
         lblFields(7).Caption = Format(lrs("nTranTotl"), "#,##0.00")
         ShowGrid
         pbnewitem = False
         
      ElseIf lrs.RecordCount > 1 Then
         lsSearch = KwikBrowse(oApp, lrs, _
                        "sSalesInv»xFullName»dTransact»nTranTotl»nAmtPaidx", _
                        "Invoice»Customer Name»Date»Tran Total»Amount Paid", _
                        "@»@»MM/dd/yy»#,##0.00»#,##0.00")
         If lsSearch <> "" Then
            psSelected = Split(lsSearch, "»")
            txtfield(0).Text = psSelected(0)
            txtfield(1).Text = psSelected(1)
            txtfield(2).Text = Format(psSelected(5), "MMMM dd, yyyy")
            txtfield(3).Text = Format(psSelected(2), "#,##0.00")
            txtfield(4).Text = Format(psSelected(3), "#,##0.00")
            txtfield(5).Text = psSelected(10)
            lblFields(0).Caption = psSelected(0)
            lblFields(1).Caption = psSelected(6)
            lblFields(1).Tag = psSelected(9)
            lblFields(2).Caption = psSelected(8)
            lblFields(2).Tag = psSelected(4)
            lblFields(3).Caption = Format(psSelected(5), "MMMM dd, yyyy")
            lblFields(4).Caption = psSelected(1)
            ClientID = lrs(9)
            Cashier = lrs(4)
            TranStat = psSelected(7)
            TransactionType
            lblFields(6).Caption = Format(psSelected(3), "#,##0.00")
            lblFields(7).Caption = Format(psSelected(2), "#,##0.00")
            pbnewitem = False
         End If
      ShowGrid
      Else
         ClearFields
      End If
Set lrs = New ADODB.Recordset
End Sub
Private Sub TransactionType()
   Select Case TranStat
      Case 0
         lblFields(5).Caption = "Cash"
      Case 1
         lblFields(5).Caption = "Credit Card"
      Case 2
         lblFields(5).Caption = "Cheque"
      Case 3
         lblFields(5).Caption = "Installment"
      Case 4
         lblFields(5).Caption = "Cancelled"
   End Select
End Sub

'Add New CP_SO_Detail for the Replacement
Private Function CP_SODetail() As Boolean
Dim lnctr As Integer
Dim lnrow As Long

CP_SODetail = True
On Error GoTo errProc
   
   With GridEditor1
      For lnctr = 1 To .Rows - 1
         lsSQL = "INSERT INTO CP_SO_Detail " _
                  & "( sTransNox, " _
                  & "  nEntryNox, " _
                  & "  sStockIDx, " _
                  & "  nQuantity, " _
                  & "  nPurPrice, " _
                  & "  nUnitPrce, " _
                  & "  nDiscount, " _
                  & "  nDiscAmnt, " _
                  & "  nSubTotal, " _
                  & "  dModified) " _
                     & "VALUES " _
                        & "('" & oDriver.FieldValue(0) & "', " _
                        & "'" & .TextMatrix(lnctr, 0) & "', " _
                        & "'" & .TextMatrix(lnctr, 7) & "', " _
                        & "'" & CLng(.TextMatrix(lnctr, 4)) & "', " _
                        & "'" & CDbl(.TextMatrix(lnctr, 11)) & "', " _
                        & "'" & CDbl(.TextMatrix(lnctr, 3)) & "', " _
                        & "'" & CLng(.TextMatrix(lnctr, 5)) & "', " _
                        & "'" & CDbl(.TextMatrix(lnctr, 6)) & "', " _
                        & "'" & CDbl(.TextMatrix(lnctr, 10)) & "', " _
                        & " getdate())"
         
         oApp.Connection.Execute lsSQL, lnrow, adCmdText
         
         If lnrow <= 0 Then
            MsgBox "Unable to Save SO Detail!!!" & vbCrLf & vbCrLf & _
            "Notify ROSALYN LAZO DESCALLAR for Assistance", vbCritical, "Warning"
            CP_SODetail = False
            GoTo endProc
         End If
      
      Next
   End With

endProc:
   Exit Function
errProc:
   CP_SODetail = False
   MsgBox Err.Description, vbCritical, "Warning"
End Function

Private Function Serial_Replacement() As Boolean
Dim lnctr As Integer
Dim lnrow As Long
Dim lsSQL As String
Dim lnEntry As Integer

Serial_Replacement = True
On Error GoTo errProc
   
   With GridEditor1
      For lnctr = 1 To .Rows - 1
         If .TextMatrix(lnctr, 12) = 1 Then
                        
            'Get Last Entry No
            lnEntry = getIMEIEntry("'" & .TextMatrix(lnctr, 9) & "'")
                                    
            'Update Location, CP_Serial_Master
            lsSQL = "UPDATE CP_Serial_Master SET" _
                  & " cSoldStat = '1', " _
                  & " cLocation = '2', " _
                  & " sClientID = '" & ClientID & "', " _
                  & " sModified = '" & Encrypt(oApp.UserID) & "', " _
                  & " dModified = getdate() " _
            & " WHERE sSerialID = '" & .TextMatrix(lnctr, 9) & "' "
            oApp.Connection.Execute lsSQL, lnrow, adCmdText
            
            'CP_SO_Serial
            lsSQL = "INSERT INTO CP_SO_Serial" _
                        & "( sTransNox ," _
                        & "  nEntryNox ," _
                        & "  sSerialID ," _
                        & "  dModified) " _
                           & " VALUES " _
                              & "('" & oDriver.FieldValue(0) & "', " _
                              & " '" & .TextMatrix(lnctr, 0) & "', " _
                              & " '" & .TextMatrix(lnctr, 9) & "', " _
                              & " getdate())"
            oApp.Connection.Execute lsSQL, lnrow, adCmdText
            
            'CP_Serial_Ledger
            lsSQL = "INSERT INTO CP_Serial_Ledger" _
                        & "( sSerialID ," _
                        & "  sBranchcd ," _
                        & "  dTransact ," _
                        & "  nEntryNox ," _
                        & "  sSourceCd ," _
                        & "  sSourceNo ," _
                        & "  cSoldStat ," _
                        & "  cLocation ," _
                        & "  dModified) " _
                           & " VALUES " _
                              & "('" & .TextMatrix(lnctr, 9) & "', " _
                              & " '" & oApp.BranchCode & "', " _
                              & "'" & CDate(txtfield(2)) & " " & Time & "', " _
                              & " '" & lnEntry & "', " _
                              & " 'CPSp', " _
                              & " '" & oDriver.FieldValue(0) & "', " _
                              & " '1', " _
                              & " '2', " _
                              & " getdate())"
            oApp.Connection.Execute lsSQL, lnrow, adCmdText
                       
            If lnrow <= 0 Then
               MsgBox "Unable to Save Serial Replacement!!!" & vbCrLf & vbCrLf & _
               "Notify ROSALYN LAZO DESCALLAR for Assistance", vbCritical, "Warning"
               Serial_Replacement = False
               GoTo endProc
            End If
         End If
      Next
   End With

endProc:
   Exit Function
errProc:
   Serial_Replacement = False
   MsgBox Err.Description, vbCritical, "Warning"
End Function
Private Function Serial_Returned() As Boolean
Dim lnctr As Integer
Dim lnrow As Long
Dim lsSQL As String
Dim lnEntry As Integer

Serial_Returned = True
On Error GoTo errProc
   
  With GridEditor1
      For lnctr = 1 To .Rows - 1
         If .TextMatrix(lnctr, 12) = 1 And .TextMatrix(lnctr, 15) = "Yes" Then
                        
            'Get Last Entry No
            lnEntry = getIMEIEntry("'" & .TextMatrix(lnctr, 13) & "'")
            
            'CP_Serial_Ledger
            lsSQL = "INSERT INTO CP_Serial_Ledger" _
                        & "( sSerialID ," _
                        & "  sBranchcd ," _
                        & "  dTransact ," _
                        & "  nEntryNox ," _
                        & "  sSourceCd ," _
                        & "  sSourceNo ," _
                        & "  cSoldStat ," _
                        & "  cLocation ," _
                        & "  dModified) " _
                           & " VALUES " _
                              & "('" & .TextMatrix(lnctr, 13) & "', " _
                              & " '" & oApp.BranchCode & "', " _
                              & "'" & CDate(txtfield(2)) & " " & Time & "', " _
                              & " '" & lnEntry & "', " _
                              & " 'CPSR', " _
                              & " '" & oDriver.FieldValue(0) & "', " _
                              & " '0', " _
                              & " '1', " _
                              & " getdate())"
            oApp.Connection.Execute lsSQL, lnrow, adCmdText
            
            'Update Location, CP_Serial_Master
            lsSQL = "UPDATE CP_Serial_Master SET" _
                  & " cSoldStat = '0', " _
                  & " cLocation = '1', " _
                  & " sClientID = '" & oDriver.FieldValue(6) & "', " _
                  & " sModified = '" & Encrypt(oApp.UserID) & "', " _
                  & " dModified = getdate() " _
            & " WHERE sSerialID = '" & .TextMatrix(lnctr, 13) & "' "
            oApp.Connection.Execute lsSQL, lnrow, adCmdText
            
            If lnrow <= 0 Then
               MsgBox "Unable to Save Serial Returned!!!" & vbCrLf & vbCrLf & _
               "Notify ROSALYN LAZO DESCALLAR for Assistance", vbCritical, "Warning"
               Serial_Returned = False
               GoTo endProc
            End If
         
         End If
      Next
   End With

endProc:
   Exit Function
errProc:
   Serial_Returned = False
   MsgBox Err.Description, vbCritical, "Warning"
End Function
Private Function Inventory_Replacement() As Boolean
Dim lsSQL As String
Dim lnrow As Long
Dim lnEntry As Integer
Dim QOH As Integer
Dim lnctr As Integer
   
Inventory_Replacement = True
On Error GoTo errProc
   
   With GridEditor1
         
      For lnctr = 1 To .Rows - 1
         If Trim(.TextMatrix(lnctr, 12)) <> 2 And .TextMatrix(lnctr, 15) = "Yes" Then
            
            'Get Last Entry No
            lnEntry = getEntryNo("CP_Inventory_Ledger", "'" & .TextMatrix(lnctr, 7) & "'", _
                        "'" & oApp.BranchCode & "'")

            'Get QOH
            QOH = getQuantity("'" & .TextMatrix(lnctr, 7) & "'", "'" & oApp.BranchCode & "'") _
                     - .TextMatrix(lnctr, 4)
            
               'Add Record, CP_Inventory_Ledger
               lsSQL = "INSERT INTO CP_Inventory_Ledger " _
                     & "( sStockIDx, " _
                     & "  sBranchCd, " _
                     & "  sLocation, " _
                     & "  sSourceCd, " _
                     & "  sSourceNo, " _
                     & "  nQtyInxxx, " _
                     & "  nQtyOutxx, " _
                     & "  nQtyOnHnd, " _
                     & "  nEntryNox, " _
                     & "  dTransact, " _
                     & "  dModified) " _
                        & "VALUES " _
                           & "('" & .TextMatrix(lnctr, 7) & "', " _
                           & "'" & oApp.BranchCode & "', " _
                           & "'" & oApp.BranchCode & "', " _
                           & "'CPSp' , " _
                           & "'" & oDriver.FieldValue(0) & "', " _
                           & " 0, " _
                           & "'" & CLng(.TextMatrix(lnctr, 4)) & "', " _
                           & "'" & CLng(QOH) & "', " _
                           & "'" & lnEntry & "', " _
                           & "'" & CDate(txtfield(2)) & " " & Time & "', " _
                           & " getdate())"
               oApp.Connection.Execute lsSQL, lnrow, adCmdText
            
               'Update QOH, CP_Inventory_Master
               lsSQL = "UPDATE CP_Inventory_Master SET" _
                     & " nQtyOnHnd = '" & CLng(QOH) & "', " _
                     & " sModified = '" & Encrypt(oApp.UserID) & "', " _
                     & " dModified = getdate() " _
               & " WHERE sStockIDx = '" & .TextMatrix(lnctr, 7) & "' " _
                     & " And sBranchCd = '" & oApp.BranchCode & "' "
               oApp.Connection.Execute lsSQL, lnrow, adCmdText
            If lnrow <= 0 Then
               MsgBox "Unable to Update Replacement Inventory !!!" & vbCrLf & vbCrLf & _
               "Notify ROSALYN LAZO DESCALLAR for Assistance", vbCritical, "Warning"
               Inventory_Replacement = False
               GoTo endProc
            End If

         End If
      Next
   
   End With

endProc:
   Exit Function
errProc:
   Inventory_Replacement = False
   MsgBox Err.Description, vbCritical, "Warning"

End Function

Private Function Inventory_Returned() As Boolean
Dim lsSQL As String
Dim lnrow As Long
Dim lnEntry As Integer
Dim QOH As Integer
Dim lnctr As Integer
   
Inventory_Returned = True
On Error GoTo errProc
   
   With GridEditor1
         
      For lnctr = 1 To .Rows - 1
         If Trim(.TextMatrix(lnctr, 12)) <> 2 And .TextMatrix(lnctr, 15) = "Yes" Then
            
            'Get Last Entry No
            lnEntry = getEntryNo("CP_Inventory_Ledger", "'" & .TextMatrix(lnctr, 14) & "'", _
                        "'" & oApp.BranchCode & "'")

            'Get QOH
            QOH = getQuantity("'" & .TextMatrix(lnctr, 14) & "'", "'" & oApp.BranchCode & "'") _
                     + .TextMatrix(lnctr, 4)

               'Add Record, CP_Inventory_Ledger
               lsSQL = "INSERT INTO CP_Inventory_Ledger " _
                     & "( sStockIDx, " _
                     & "  sBranchCd, " _
                     & "  sLocation, " _
                     & "  sSourceCd, " _
                     & "  sSourceNo, " _
                     & "  nQtyInxxx, " _
                     & "  nQtyOutxx, " _
                     & "  nQtyOnHnd, " _
                     & "  nEntryNox, " _
                     & "  dTransact, " _
                     & "  dModified) " _
                        & "VALUES " _
                           & "('" & .TextMatrix(lnctr, 14) & "', " _
                           & "'" & oApp.BranchCode & " ', " _
                           & "'" & oApp.BranchCode & " ', " _
                           & "'CPSR' , " _
                           & "'" & oDriver.FieldValue(0) & "', " _
                           & "'" & CLng(.TextMatrix(lnctr, 4)) & "', " _
                           & " 0, " _
                           & "'" & CLng(QOH) & "', " _
                           & "'" & lnEntry & "', " _
                           & "'" & CDate(txtfield(2)) & " " & Time & "', " _
                           & " getdate())"
               oApp.Connection.Execute lsSQL, lnrow, adCmdText
            
               'Update QOH, CP_Inventory_Master
               lsSQL = "UPDATE CP_Inventory_Master SET" _
                     & " nQtyOnHnd = '" & CLng(QOH) & "', " _
                     & " sModified = '" & Encrypt(oApp.UserID) & "', " _
                     & " dModified = getdate() " _
               & " WHERE sStockIDx = '" & .TextMatrix(lnctr, 14) & "' " _
                     & " And sBranchCd = '" & oApp.BranchCode & "' "
               oApp.Connection.Execute lsSQL, lnrow, adCmdText
               
               'Update cTranStat, CP_SO_Master
               lsSQL = "UPDATE CP_SO_Master SET" _
                     & " cTranStat = '4', " _
                     & " sRemarksx = sRemarksx + '  Replaced Unit' ," _
                     & " sModified = '" & Encrypt(oApp.UserID) & "'," _
                     & " dModified = getdate() " _
               & " WHERE sTransNox = '" & lblFields(0).Caption & "' "
               
               oApp.Connection.Execute lsSQL, lnrow, adCmdText
            
            If lnrow <= 0 Then
               MsgBox "Unable to Update Returned Inventory !!!" & vbCrLf & vbCrLf & _
               "Notify ROSALYN LAZO DESCALLAR for Assistance", vbCritical, "Warning"
               Inventory_Returned = False
               GoTo endProc
            End If

         End If
      Next
   
   End With

endProc:
   Exit Function
errProc:
   Inventory_Returned = False
   MsgBox Err.Description, vbCritical, "Warning"
End Function

'Credit Card Transaction
Private Function SaveCP_SOCredit() As Boolean
Dim lnctr As Integer
Dim lnrow As Long
Dim lsSQL As String

SaveCP_SOCredit = True
On Error GoTo errProc
   
   With frmCard_POS
      lsSQL = "INSERT INTO CP_SO_Credit " _
               & "(   sTransNox, " _
                  & " dTransact, " _
                  & " sClientID, " _
                  & " sCreditID, " _
                  & " nTranTotl, " _
                  & " nCashAmnt, " _
                  & " nCardAmnt, " _
                  & " sSalesInv, " _
                  & " sAcctNmbr, " _
                  & " nPercentx, " _
                  & " dModified) " _
                     & "VALUES " _
                        & "('" & oDriver.FieldValue(0) & "', " _
                        & "'" & CDate(txtfield(2).Text) & " " & Time & "', " _
                        & "'" & ClientID & "', " _
                        & "'" & .txtfield(0).Tag & "', " _
                        & "'" & CDbl(txtothers(1)) & "', " _
                        & "'" & CDbl(.txtfield(2).Text) & "', " _
                        & "'" & CDbl(.txtfield(3).Text) & "', " _
                        & "'" & txtfield(1).Text & "', " _
                        & "'" & .txtfield(5).Text & "', " _
                        & "'" & .txtothers(0).Text & "', " _
                        & " getdate())"
      oApp.Connection.Execute lsSQL, lnrow, adCmdText
      
      If lnrow <= 0 Then
         MsgBox "Unable to Save Credit Card Details!!!", vbCritical, "Warning"
         SaveCP_SOCredit = False
         GoTo endProc
      End If
   End With
      
endProc:
   Exit Function
errProc:
   SaveCP_SOCredit = False
   MsgBox Err.Description, vbCritical, "Warning"

End Function

'Cheque Transaction
Private Function SaveCP_SOCheque() As Boolean
Dim lnrow As Long
Dim lsSQL As String

SaveCP_SOCheque = True
On Error GoTo errProc
         
   With frmCheque_POS
      lsSQL = "INSERT INTO CP_SO_Cheque " _
               & "(   sTransNox, " _
                  & " dTransact, " _
                  & " sClientID, " _
                  & " sBankIDxx, " _
                  & " nTranTotl, " _
                  & " nCashAmnt, " _
                  & " nCheqAmnt, " _
                  & " sAccntNum, " _
                  & " sSalesInv, " _
                  & " dModified) " _
                     & "VALUES " _
                        & "('" & oDriver.FieldValue(0) & "', " _
                        & "'" & CDate(txtfield(2).Text) & " " & Time & "', " _
                        & "'" & ClientID & "', " _
                        & "'" & .txtfield(0).Tag & "', " _
                        & "'" & CDbl(txtothers(1).Text) & "', " _
                        & "'" & CDbl(.txtfield(2).Text) & "', " _
                        & "'" & CDbl(.txtfield(3).Text) & "', " _
                        & "'" & .txtfield(5).Text & "', " _
                        & "'" & txtfield(1).Text & "', " _
                        & " getdate())"
      
      oApp.Connection.Execute lsSQL, lnrow, adCmdText
      
      If lnrow <= 0 Then
         MsgBox "Unable to Save Cheque Details!!!", vbCritical, "Warning"
         SaveCP_SOCheque = False
         GoTo endProc
      End If
   End With
endProc:
   Exit Function
errProc:
   SaveCP_SOCheque = False
   MsgBox Err.Description, vbCritical, "Warning"

End Function

'Installment Transaction
Private Function SaveCP_SOInstallment() As Boolean
Dim lnrow As Long
Dim lsSQL As String

SaveCP_SOInstallment = True
On Error GoTo errProc
         
   With frmInstallment_POS
      lsSQL = "INSERT INTO CP_SO_Installment " _
               & "(   sTransNox, " _
                  & " dTransact ," _
                  & " sClientID ," _
                  & " nTranTotl ," _
                  & " nDownPaym ," _
                  & " nBalancex ," _
                  & " nPaymTerm ," _
                  & " nMonthlyP ," _
                  & " sSalesInv ," _
                  & " dModified) " _
                     & "VALUES " _
                        & "('" & oDriver.FieldValue(0) & "', " _
                        & "'" & oApp.ServerDate & "', " _
                        & "'" & oDriver.FieldValue(4) & "', " _
                        & "'" & CDbl(txtfield(2).Text) & "', " _
                        & "'" & CDbl(.txtfield(1).Text) & "', " _
                        & "'" & CDbl(.txtfield(2).Text) & "', " _
                        & "'" & CLng(.txtfield(3).Text) & "', " _
                        & "'" & CDbl(.txtfield(4).Text) & "', " _
                        & "'" & txtfield(1).Text & "', " _
                        & " getdate())"
      oApp.Connection.Execute lsSQL, lnrow, adCmdText
      
      If lnrow <= 0 Then
         MsgBox "Unable to Save Installment Details!!!", vbCritical, "Warning"
         SaveCP_SOInstallment = False
         GoTo endProc
      End If
   End With
endProc:
   Exit Function
errProc:
   SaveCP_SOInstallment = False
   MsgBox Err.Description, vbCritical, "Warning"

End Function

Private Sub txtField_LostFocus(Index As Integer)
   txtfield(Index).BackColor = &HFFFFFF
   If Index = 4 Then
      txtothers(0).Text = CDbl(txtfield(Index).Text) - CDbl(txtothers(1).Text)
      txtothers(0).Text = Format(txtothers(0).Text, "#,##0.00")
   End If
End Sub

Private Sub txtfield_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 2
         If Not IsDate(txtfield(Index).Text) Then
            txtfield(Index).Text = Format(Date, "MMMM dd, yyyy")
         Else
            txtfield(Index).Text = Format(txtfield(Index).Text, "MMMM dd, yyyy")
         End If
      Case 3
         If Not IsNumeric(txtfield(Index).Text) Then txtfield(Index).Text = 0#
      Case 4
         If Not IsNumeric(txtfield(Index).Text) Then
            txtfield(Index).Text = "0.00"
         Else
            txtfield(Index).Text = Format(txtfield(Index).Text, "#,##0.00")
         End If
         If CDbl(txtfield(Index).Text) >= CDbl(txtothers(1).Text) Then
            txtothers(0).Text = CDbl(txtfield(Index).Text) - CDbl(txtothers(1).Text)
            txtothers(0).Text = Format(txtothers(0).Text, "#,##0.00")
         Else
            MsgBox "Amount Can't be Less than Cash Due!!!", vbCritical, "Warning"
            txtfield(Index).SetFocus
         End If
   End Select
   txtfield(Index).Text = TitleCase(txtfield(Index).Text)
   Cancel = Not oDriver.ValidateField(Index)
End Sub

Private Sub txtothers_Validate(Index As Integer, Cancel As Boolean)
   If Not IsNumeric(txtothers(Index).Text) Then
      txtothers(Index).Text = "0.00"
   Else
      txtothers(Index).Text = Format(txtothers(Index).Text, "#,##0.00")
   End If
End Sub

 

