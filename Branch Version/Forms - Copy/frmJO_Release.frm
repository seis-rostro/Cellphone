VERSION 5.00
Object = "{0A7B56A6-35D0-4533-91DA-1715D3A0DD3E}#1.1#0"; "xrGridControl.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmJobOrder_Release 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Job Order Release"
   ClientHeight    =   6600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11220
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6600
   ScaleWidth      =   11220
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1920
      Index           =   4
      Left            =   1575
      Tag             =   "wt0;fb0"
      Top             =   2910
      Width           =   9540
      _ExtentX        =   16828
      _ExtentY        =   3387
      BackColor       =   14286077
      Enabled         =   0   'False
      ClipControls    =   0   'False
      BorderStyle     =   1
      Begin xrGridEditor.GridEditor GridEditor1 
         Height          =   1785
         Left            =   60
         TabIndex        =   21
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
         MOUSEICON       =   "frmJO_Release.frx":0000
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
      Height          =   450
      Index           =   0
      Left            =   1575
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   9540
      _ExtentX        =   16828
      _ExtentY        =   794
      BackColor       =   7716603
      ClipControls    =   0   'False
      BorderStyle     =   1
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
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1200
      Index           =   1
      Left            =   1575
      Tag             =   "wt0;fb0"
      Top             =   1020
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   2117
      BackColor       =   7716603
      Enabled         =   0   'False
      ClipControls    =   0   'False
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name :"
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   30
         Top             =   345
         Width           =   1215
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Job Order No.:"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   29
         Top             =   90
         Width           =   1035
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contact No. :"
         Height          =   195
         Index           =   2
         Left            =   330
         TabIndex        =   28
         Top             =   600
         Width           =   945
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Received :"
         Height          =   195
         Index           =   28
         Left            =   105
         TabIndex        =   27
         Top             =   855
         Width           =   1170
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   30
         Left            =   2280
         TabIndex        =   26
         Top             =   690
         Width           =   30
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
         TabIndex        =   25
         Tag             =   "tc0"
         Top             =   90
         Width           =   2505
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
         TabIndex        =   24
         Tag             =   "tc0"
         Top             =   345
         Width           =   4005
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
         TabIndex        =   23
         Tag             =   "tc0"
         Top             =   600
         Width           =   4005
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
         Left            =   1350
         TabIndex        =   22
         Tag             =   "tc0"
         Top             =   855
         Width           =   3600
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1200
      Index           =   2
      Left            =   7095
      Tag             =   "wt0;fb0"
      Top             =   1020
      Width           =   4020
      _ExtentX        =   7091
      _ExtentY        =   2117
      BackColor       =   7716603
      Enabled         =   0   'False
      ClipControls    =   0   'False
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Brand & Model :"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   38
         Top             =   600
         Width           =   1035
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IMEI No. :"
         Height          =   195
         Index           =   7
         Left            =   555
         TabIndex        =   37
         Top             =   90
         Width           =   720
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Purchase Date :"
         Height          =   195
         Index           =   27
         Left            =   120
         TabIndex        =   36
         Top             =   345
         Width           =   1155
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
         TabIndex        =   35
         Tag             =   "tc0"
         Top             =   600
         Width           =   2490
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status :"
         Height          =   195
         Index           =   6
         Left            =   735
         TabIndex        =   34
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
         Index           =   7
         Left            =   1395
         TabIndex        =   33
         Tag             =   "tc0"
         Top             =   855
         Width           =   2595
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
         TabIndex        =   32
         Tag             =   "tc0"
         Top             =   345
         Width           =   2430
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
         Index           =   4
         Left            =   1395
         TabIndex        =   31
         Tag             =   "tc0"
         Top             =   90
         Width           =   2490
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   690
      Index           =   3
      Left            =   1575
      Tag             =   "wt0;fb0"
      Top             =   2205
      Width           =   9540
      _ExtentX        =   16828
      _ExtentY        =   1217
      BackColor       =   7716603
      Enabled         =   0   'False
      ClipControls    =   0   'False
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         BackColor       =   &H0075BEFB&
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
         TabIndex        =   39
         TabStop         =   0   'False
         Tag             =   "tc0;fb0"
         Text            =   "frmJO_Release.frx":001C
         Top             =   90
         Width           =   4050
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Complaint :"
         Height          =   195
         Index           =   12
         Left            =   495
         TabIndex        =   44
         Top             =   90
         Width           =   780
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
         TabIndex        =   43
         Tag             =   "tc0"
         Top             =   345
         Width           =   2595
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
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Job Description :"
         Height          =   195
         Index           =   13
         Left            =   5610
         TabIndex        =   41
         Top             =   90
         Width           =   1185
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Description :"
         Height          =   195
         Index           =   15
         Left            =   5565
         TabIndex        =   40
         Top             =   345
         Width           =   1215
      End
   End
   Begin xrControl.xrFrame xrFrame3 
      Height          =   1635
      Left            =   1575
      Tag             =   "wt0;fb0"
      Top             =   4860
      Width           =   9540
      _ExtentX        =   16828
      _ExtentY        =   2884
      BackColor       =   7716603
      ClipControls    =   0   'False
      BorderStyle     =   1
      Begin VB.TextBox txtfield 
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
         Index           =   9
         Left            =   1680
         TabIndex        =   47
         Text            =   "Text1"
         Top             =   200
         Width           =   3420
      End
      Begin VB.TextBox txtfield 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   465
         Index           =   6
         Left            =   6825
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Tag             =   "ht0;ft0"
         Text            =   "Total Amount"
         Top             =   180
         Width           =   2565
      End
      Begin VB.TextBox txtfield 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   3615
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   675
         Width           =   1500
      End
      Begin VB.TextBox txtfield 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   840
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   1035
         Width           =   1500
      End
      Begin VB.TextBox txtfield 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   3615
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   1065
         Width           =   1500
      End
      Begin VB.TextBox txtfield 
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
         Height          =   375
         Index           =   8
         Left            =   6825
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Tag             =   "ht0;ft0"
         Text            =   "Change"
         Top             =   1050
         Width           =   2565
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
         Index           =   7
         Left            =   6825
         TabIndex        =   14
         Text            =   "Cash Given"
         Top             =   660
         Width           =   2565
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Released"
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
         Index           =   17
         Left            =   240
         TabIndex        =   46
         Top             =   225
         Width           =   2715
      End
      Begin VB.Shape Shape6 
         Height          =   1410
         Left            =   75
         Top             =   90
         Width           =   5190
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Charges "
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
         Index           =   16
         Left            =   240
         TabIndex        =   4
         Top             =   735
         Width           =   765
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   210
         X2              =   5115
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Labor"
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
         Index           =   14
         Left            =   2550
         TabIndex        =   7
         Top             =   735
         Width           =   990
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Parts"
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
         Index           =   11
         Left            =   2580
         TabIndex        =   9
         Top             =   1065
         Width           =   945
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Misc."
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
         Index           =   10
         Left            =   240
         TabIndex        =   5
         Top             =   1095
         Width           =   465
      End
      Begin VB.Shape Shape1 
         Height          =   1395
         Left            =   5310
         Top             =   105
         Width           =   4140
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
         Left            =   5370
         TabIndex        =   11
         Top             =   165
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
         Index           =   9
         Left            =   5475
         TabIndex        =   15
         Top             =   1095
         Width           =   1620
      End
      Begin VB.Label lblField 
         BackStyle       =   0  'Transparent
         Caption         =   "Cash Given"
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
         Index           =   3
         Left            =   5550
         TabIndex        =   13
         Top             =   720
         Width           =   1545
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   1
      Left            =   90
      TabIndex        =   18
      Top             =   4140
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
      Caption         =   "Pre&view"
      AccessKey       =   "v"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmJO_Release.frx":002B
      CaptionAlign    =   0
      BackColor       =   14286077
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   0
      Left            =   90
      TabIndex        =   17
      Top             =   3720
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
      Picture         =   "frmJO_Release.frx":113D
      CaptionAlign    =   0
      BackColor       =   14286077
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   2
      Left            =   90
      TabIndex        =   20
      Top             =   4560
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
      Picture         =   "frmJO_Release.frx":18B7
      CaptionAlign    =   0
      BackColor       =   14286077
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   405
      Index           =   4
      Left            =   90
      TabIndex        =   19
      Top             =   4140
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
      Picture         =   "frmJO_Release.frx":2031
      CaptionAlign    =   0
      BackColor       =   14286077
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   5
      Left            =   90
      TabIndex        =   45
      Top             =   4560
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
      Picture         =   "frmJO_Release.frx":27AB
      CaptionAlign    =   0
   End
End
Attribute VB_Name = "frmJobOrder_Release"
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
Dim response As String

   Select Case Index
      Case 0   'save
         If lblFields(0).Caption <> "" Then
            If txtField(6).Text = 0# Then
               response = MsgBox("Unit Under Warranty?", vbYesNo, "Confirmation")
               If response <> vbYes Then
                  MsgBox "Verify Job Order Details!!!", vbCritical, "Warning"
                  Exit Sub
               Else
                  Cancel = Not UpdateCP_JobOrder_Master
                     If Cancel Then Exit Sub
                  Cancel = Not Update_Warranty
                     If Cancel Then Exit Sub
                  oDriver.HideButton 0
                  response = MsgBox("Print Job Order?", vbYesNo, "Confirmation")
                  If response <> vbYes Then Exit Sub
                  Print_JobOrder
               End If
            Else
               If CDbl(txtField(7).Text) < CDbl(txtField(6).Text) Then
                  MsgBox "Invalid Input!!!", vbCritical, "Warning"
                  txtField(7).SetFocus
               Else
                  If CDbl(txtField(Index).Text) < CDbl(txtField(6).Text) Then
                     MsgBox "Invalid Amount!!!", vbCritical, "Warning"
                     txtField(Index).SetFocus
                  Else
                     Cancel = Not UpdateCP_JobOrder_Master
                        If Cancel Then Exit Sub
                     Cancel = Not Update_Warranty
                        If Cancel Then Exit Sub
                     oDriver.HideButton 0
                     response = MsgBox("Print Job Order?", vbYesNo, "Confirmation")
                     If response <> vbYes Then Exit Sub
                     Print_JobOrder
                  End If
               End If
            End If
         Else
            Exit Sub
         End If
      Case 1   'Print
            Print_JobOrder
      Case 2   'Close
            Unload Me
      Case 4   'Browse
            If pnindex = 0 Or pnindex = 1 Then Search_JobOrder False
      Case 5   'New
            InitTxtField
            EmptyGrid
            InitButton xeModeAddNew
      End Select

End Sub

Private Sub Form_Activate()
Dim lnCtr As Integer
   Select Case lnCtr
   Case 2 To 6, 8
      oDriver.DisableTextbox lnCtr
   End Select

End Sub

Private Sub InitButton(lnStat As Integer)
   Dim lbShow As Boolean

   lbShow = IIf(lnStat = xeModeReady, False, True)
   cmdButton(2).Visible = lbShow
   cmdButton(4).Visible = lbShow
   
   cmdButton(1).Visible = Not lbShow
   cmdButton(5).Visible = Not lbShow
   
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
oSkin.ApplySkin xeFormTransaction

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
      
      For Index = 1 To 7
         Select Case Index
            Case 1
               .ColAlignment(Index) = 1
            Case 2, 3, 5, 7
               .ColFormat(Index) = "#,##0.00"
               .ColDefault(Index) = "0.00"
               .ColAlignment(Index) = 6
               If Index = 5 Then .ColEnabled(Index) = False
            Case 4
                  .ColMaxValue(Index) = 100
                  .ColDefault(Index) = 0
            Case 6
                  .ColMaxValue(Index) = 999
                  .ColDefault(Index) = 1
         End Select
      Next
        
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
      Case 3 To 8
         txtField(Index).Text = "0.00"
      Case 0 To 2
         txtField(Index).Text = ""
   End Select
Next
txtField(9).Text = Format(Date, "MMMM dd, yyyy")
txtField(9).Enabled = True
oDriver.ShowButton 0

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
                     & " WHERE a.cTranStat = 0 " _

   If pnindex = 0 Then
      If SearchValue Then
         lsSQL = lsSQL & " AND sJobOrdNo = '" & txtField(0).Text & "'"
      Else
         lsSQL = lsSQL & " AND sJobOrdNo LIKE '" & txtField(0).Text & "%' "
      End If
   ElseIf pnindex = 1 Then
      If SearchValue Then
         lsSQL = lsSQL & " AND sLastName + ', ' + sFrstName + ' ' + sMiddName = '" & txtField(1).Text & "'"
      Else
         lsSQL = lsSQL & " AND sLastName + ', ' + sFrstName + ' ' + sMiddName LIKE '" & txtField(1).Text & "%' "
      End If
   End If
   
   lsSQL = lsSQL & " ORDER BY sLastName + ', ' + sFrstName + ' ' + sMiddName"
   lrs.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
            
   If lrs.RecordCount = 1 Then
      For Index = 0 To 9
         Select Case Index
            Case 0, 1
               txtField(Index).Text = IIf(IsNull(lrs(Index)), "", lrs(Index))
               If Index = 0 Then txtField(Index).Tag = lrs(15)
            Case 2
               txtField(Index).Text = lrs(11)
            Case 3 To 6
               txtField(Index).Text = Format(lrs(Index + 13), "#,##0.00")
               If Index = 6 Then txtField(Index).Text = Format(lrs(12), "#,##0.00")
            Case 9
               txtField(Index).Text = Format(lrs(14), "MMMM dd, yyyy")
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
                     txtField(7).Text = "0.00"
                     txtField(8).Text = "0.00"
                     If xrFrame3.Enabled = False Then xrFrame3.Enabled = True
                  Else
                     .Caption = "Claimed"
                     txtField(7).Text = Format(lrs("nAmtPaidx"), "#,##0.00")
                     txtField(8).Text = Format(CDbl(lrs("nAmtPaidx") - lrs("nTranTotl")), "#,##0.00")
                     If xrFrame3.Enabled = True Then xrFrame3.Enabled = False
                  End If
               Case 8
                  '1-Void Warranty ; 2-Under Limited Warranty ; 3-Back Job
                  If lrs(lnCtr) = 1 Then
                     .Caption = "Void Warranty"
                  ElseIf lrs(lnCtr) = 2 Then
                     .Caption = "Under Limited Warranty"
                  ElseIf lrs(lnCtr) = 3 Then
                     .Caption = "Back Job" & " " & lrs(13)
                  End If
               Case 9
                  '0-New Unit; 1-Battery; 2-Sim Card; 3-Back Cover; 4-Charger; 5-Others
                  If lrs(lnCtr) = 0 Then
                     .Caption = "New Unit"
                  ElseIf lrs(lnCtr) = 1 Then
                     .Caption = "Battery"
                  ElseIf lrs(lnCtr) = 2 Then
                     .Caption = "Sim Card"
                  ElseIf lrs(lnCtr) = 3 Then
                     .Caption = "Back Cover"
                  ElseIf lrs(lnCtr) = 4 Then
                     .Caption = "Charger"
                  ElseIf lrs(lnCtr) = 5 Then
                     .Caption = lrs(10)
                  End If
            End Select
         End With
      Next
      ShowGrid
      InitButton xeModeReady
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
      InitButton xeModeReady
   Else
      MsgBox "No Record Found!!!", vbInformation, Me.Caption
   End If
   
   Set lrs = Nothing

End Sub

Private Sub ShowMoreRec()
Dim Index As Integer

For Index = 0 To 9
   Select Case Index
      Case 0, 1
         txtField(Index).Text = psSelected(Index)
         If Index = 0 Then txtField(Index).Tag = psSelected(15)
      Case 2
         txtField(Index).Text = psSelected(11)
      Case 3 To 6
         txtField(Index).Text = Format(psSelected(Index + 13), "#,##0.00")
         If Index = 6 Then txtField(Index).Text = Format(psSelected(12), "#,##0.00")
      Case 9
         txtField(Index).Text = Format(psSelected(14), "MMMM dd,yyyy")
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
                     txtField(7).Text = "0.00"
                     txtField(8).Text = "0.00"
                     If xrFrame3.Enabled = False Then xrFrame3.Enabled = True
                  Else
                     .Caption = "Claimed"
                     txtField(7).Text = Format(psSelected(19), "#,##0.00")
                     txtField(8).Text = Format(CDbl(psSelected(19) - psSelected(12)), "#,##0.00")
                     If xrFrame3.Enabled = True Then xrFrame3.Enabled = False
                  End If
               Case 8
                  '1-Void Warranty ; 2-Under Limited Warranty ; 3-Back Job
                  If psSelected(lnCtr) = 1 Then
                     .Caption = "Void Warranty"
                  ElseIf psSelected(lnCtr) = 2 Then
                     .Caption = "Under Limited Warranty"
                  ElseIf psSelected(lnCtr) = 3 Then
                     .Caption = "Back Job" & " " & psSelected(13)
                  End If
               Case 9
                  '0-New Unit; 1-Battery; 2-Sim Card; 3-Back Cover; 4-Charger; 5-Others
                  If psSelected(lnCtr) = 0 Then
                     .Caption = "New Unit"
                  ElseIf psSelected(lnCtr) = 1 Then
                     .Caption = "Battery"
                  ElseIf psSelected(lnCtr) = 2 Then
                     .Caption = "Sim Card"
                  ElseIf psSelected(lnCtr) = 3 Then
                     .Caption = "Back Cover"
                  ElseIf psSelected(lnCtr) = 4 Then
                     .Caption = "Charger"
                  ElseIf psSelected(lnCtr) = 5 Then
                     .Caption = psSelected(10)
                  End If
            End Select
         End With
      Next

End Sub

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

Private Sub oDriver_DisableOtherControl()
Dim lnCtr As Integer
   Select Case lnCtr
   Case 2 To 6, 8
      oDriver.DisableTextbox lnCtr
   End Select
End Sub

Private Sub oDriver_EnableOtherControl()
Dim lnCtr As Integer
   Select Case lnCtr
   Case 2 To 6, 8
      oDriver.DisableTextbox lnCtr
   End Select

End Sub

Private Sub txtField_GotFocus(Index As Integer)
   txtfieldGotfocus = True
   pnindex = Index
   oDriver.ColumnIndex = Index
   txtField(Index).BackColor = &HE1FEFF
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

   If KeyCode = vbKeyF3 Or KeyCode = 13 Then
      If Index = 0 Or Index = 1 Then Search_JobOrder False
      If txtField(Index).Text <> "" Then SetNextFocus
      KeyCode = 0
   End If

End Sub

Private Sub txtfield_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
   Case 7
      If CDbl(txtField(Index).Text) > CDbl(txtField(6).Text) Then
         txtField(8).Text = Format(CDbl(txtField(7).Text) _
                                 - CDbl(txtField(6).Text), "#,##0.00")
      End If
   Case 9
      If Not IsDate(txtField(Index).Text) Then
         txtField(Index).Text = Format(oApp.ServerDate, "MMMM dd,yyyy")
      Else
         txtField(Index).Text = Format(txtField(Index).Text, "MMMM dd,yyyy")
      End If
   End Select
   txtField(Index).Text = TitleCase(txtField(Index).Text)
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   Select Case Index
      Case 7
         If Not IsNumeric(txtField(Index).Text) Then
            MsgBox "Invalid Input!!!", vbCritical, "Warning"
            txtField(Index).Text = 0#
         Else
            txtField(Index).Text = Format(txtField(Index).Text, "#,##0.00")
         End If
   End Select
   txtField(Index).BackColor = &HFFFFFF
End Sub
Private Function Update_Warranty() As Boolean
Dim lsSQL As String
Dim lnrow As Long
   
Update_Warranty = True
On Error GoTo errProc

      'Update CP_JobOrder_Master
      lsSQL = "UPDATE CP_JobOrder_Master SET" _
            & " cTranStat = '1', " _
            & " sPaymRecv = '" & Encrypt(oApp.UserID) & "', " _
            & " sModified = '" & Encrypt(oApp.UserID) & "', " _
            & " dModified = getdate() " _
      & " WHERE sTransNox = '" & txtField(0).Text & "' "
      oApp.Connection.Execute lsSQL, lnrow, adCmdText
                   
'      If lnrow <= 0 Then
'         MsgBox "Unable to Update Job Order Master!!!" & vbCrLf & vbCrLf & _
'         "Notify ROSALYN LAZO DESCALLAR for Assistance", vbCritical, "Warning"
'         Update_Warranty = False
'         GoTo endProc
'      End If
   
endProc:
   Exit Function
errProc:
   Update_Warranty = False
   MsgBox Err.Description, vbCritical, "Warning"
End Function

Private Function UpdateCP_JobOrder_Master() As Boolean
Dim lsSQL As String
Dim lnrow As Long

UpdateCP_JobOrder_Master = True
On Error GoTo errProc

   With frmJobOrder
      'Update CP_JobOrder_Master
      lsSQL = "UPDATE CP_JobOrder_Master SET" _
            & " nAmtPaidx = '" & CDbl(txtField(6).Text) & "', " _
            & " dPaymentx = '" & CDate(txtField(9).Text) & "', " _
            & " sPaymRecv = '" & Encrypt(oApp.UserID) & "', " _
            & " cTranstat = '1' ," _
            & " sModified = '" & Encrypt(oApp.UserID) & "', " _
            & " dModified = getdate() " _
      & " WHERE sTransNox = '" & txtField(0).Tag & "' " _

      oApp.Connection.Execute lsSQL, lnrow, adCmdText

      If lnrow = 0 Then
         MsgBox "Unable to Update Job Order Master!!!" & vbCrLf & vbCrLf & _
         "Notify ROSALYN LAZO DESCALLAR for Assistance", vbCritical, "Warning"
         UpdateCP_JobOrder_Master = False
      Else
         MsgBox "Job Order Successfully Updated!!!", vbInformation, Me.Caption
      End If
   End With

endProc:
   Exit Function
errProc:
   UpdateCP_JobOrder_Master = False
   MsgBox Err.Description, vbCritical, "Warning"

End Function

Private Sub Print_JobOrder()
Dim lnCtr As Integer
Dim lrs As New ADODB.Recordset
Dim lrsReport As New ADODB.Recordset
Dim lsSQL As String

   Set lrs = New ADODB.Recordset
   Set lrsReport = New ADODB.Recordset

   lrs.Fields.Append "sField01", adVarChar, 150
   lrs.Fields.Append "sField02", adVarChar, 150
   lrs.Open

   'Job_Order
    lsSQL = "SELECT" _
               & " a.sTransNox, " _
               & " b.sAddressx + ' ' + c.sTownName as xAddressx " _

   lsSQL = lsSQL & " FROM CP_JobOrder_Master a " _
            & " LEFT JOIN Client_Master b " _
               & " ON a.sClientID = b.sClientID " _
            & " LEFT JOIN TownCity c " _
               & " ON b.sTownIDxx = c.sTownIDxx, " _
            & " Branch f " _
         & " WHERE a.sTransNox = '" & txtField(0).Tag & "' " _
            & " AND f.sBranchCd = '" & Left(txtField(0).Tag, 2) & "' " _

   If lrsReport.State = adStateOpen Then lrsReport.Close
   lrsReport.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   
   If lrsReport.EOF Then
      MsgBox "Save Transaction!!!" & vbCrLf & _
             "Then Try Again!!!", vbCritical, "Notice"
      Exit Sub
   End If
            lrs.AddNew
            lrs("sField01").Value = lrsReport("xAddressx")

         Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\CP_JobOrder_Release.rpt")
         oReport.DiscardSavedData
         oReport.FieldMappingType = crAutoFieldMapping
         oReport.Database.SetDataSource lrs
         
         With oReport
            .Sections("RH").ReportObjects("txtJONo").SetText txtField(0).Text
            .Sections("PH").ReportObjects("txtCustomer").SetText txtField(1).Text
            .Sections("PH").ReportObjects("txtTelephone").SetText lblFields(2).Caption
            .Sections("PH").ReportObjects("txtBrandModel").SetText lblFields(6).Caption
            .Sections("PH").ReportObjects("txtIMEINo").SetText lblFields(4).Caption
            .Sections("PH").ReportObjects("txtComplaint").SetText txtField(2).Text
            .Sections("PH").ReportObjects("txtReceived").SetText lblFields(3).Caption
            .Sections("PH").ReportObjects("txtLabor").SetText txtField(4).Text
            .Sections("PH").ReportObjects("txtParts").SetText txtField(5).Text
            .Sections("PH").ReportObjects("txtMisc").SetText txtField(3).Text
            .Sections("PH").ReportObjects("txtTotal").SetText txtField(6).Text
         End With
         
      Set lrs = Nothing
      Set lrsReport = Nothing
      
      frmViewer.CRViewer91.ReportSource = oReport
      frmViewer.CRViewer91.ViewReport
      frmViewer.Show

End Sub


