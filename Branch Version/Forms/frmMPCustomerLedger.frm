VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0"; "xrGridControl.ocx"
Begin VB.Form frmMPCustomerLedger 
   BorderStyle     =   0  'None
   Caption         =   "Customer Ledger (Active)"
   ClientHeight    =   7440
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11940
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   HelpContextID   =   10590
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   11940
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   8
      Left            =   10590
      TabIndex        =   50
      Top             =   1185
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&ReCalc"
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
      Picture         =   "frmMPCustomerLedger.frx":0000
   End
   Begin xrGridEditor.GridEditor GridEditor1 
      Height          =   3420
      Left            =   120
      TabIndex        =   40
      Tag             =   "et0;eb0;et0;bc2"
      Top             =   3420
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   6033
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
      Object.HEIGHT          =   3420
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
      MOUSEICON       =   "frmMPCustomerLedger.frx":072E
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
      Height          =   2295
      Index           =   0
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   1095
      Width           =   5490
      _ExtentX        =   9684
      _ExtentY        =   4048
      Begin VB.TextBox txtField 
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
         Height          =   390
         Index           =   2
         Left            =   1215
         MultiLine       =   -1  'True
         TabIndex        =   9
         Tag             =   "tc0;fb0"
         Text            =   "frmMPCustomerLedger.frx":074A
         Top             =   615
         Width           =   4185
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Co-Brwr #2:"
         Height          =   285
         Index           =   26
         Left            =   135
         TabIndex        =   63
         Top             =   1230
         Width           =   960
      End
      Begin VB.Label lblFields 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "012314568901234567890123456789012345679"
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
         Index           =   18
         Left            =   1200
         TabIndex        =   62
         Tag             =   "tc0"
         Top             =   1245
         Width           =   4185
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Co-Brwr #1:"
         Height          =   285
         Index           =   25
         Left            =   135
         TabIndex        =   61
         Top             =   1005
         Width           =   960
      End
      Begin VB.Label lblFields 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "012314568901234567890123456789012345679"
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
         Index           =   17
         Left            =   1215
         TabIndex        =   60
         Tag             =   "tc0"
         Top             =   1035
         Width           =   4185
      End
      Begin VB.Label lblFields 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "012314568901234567890123456789012345679"
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
         Index           =   16
         Left            =   1215
         TabIndex        =   13
         Tag             =   "tc0"
         Top             =   1725
         Width           =   4185
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Col. Branch:"
         Height          =   285
         Index           =   23
         Left            =   150
         TabIndex        =   12
         Top             =   1725
         Width           =   960
      End
      Begin VB.Label lblFields 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "012314568901234567890123456789012345679"
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
         Left            =   1215
         TabIndex        =   15
         Tag             =   "tc0"
         Top             =   1980
         Width           =   4185
      End
      Begin VB.Label lblFields 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "012314568901234567890123456789012345679"
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
         Left            =   1215
         TabIndex        =   11
         Tag             =   "tc0"
         Top             =   1485
         Width           =   4185
      End
      Begin VB.Label lblFields 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "012314568901234567890123456789012345679"
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
         Left            =   1215
         TabIndex        =   7
         Tag             =   "tc0"
         Top             =   375
         Width           =   4185
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
         Left            =   1215
         TabIndex        =   5
         Tag             =   "tc0"
         Top             =   150
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Account No.:"
         Height          =   285
         Index           =   0
         Left            =   150
         TabIndex        =   4
         Top             =   135
         Width           =   960
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Model:"
         Height          =   285
         Index           =   4
         Left            =   150
         TabIndex        =   10
         Top             =   1470
         Width           =   960
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Collector:"
         Height          =   285
         Index           =   7
         Left            =   150
         TabIndex        =   14
         Top             =   1965
         Width           =   960
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
         Height          =   285
         Index           =   3
         Left            =   150
         TabIndex        =   8
         Top             =   600
         Width           =   960
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   285
         Index           =   1
         Left            =   150
         TabIndex        =   6
         Top             =   360
         Width           =   960
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   525
      Index           =   1
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   10245
      _ExtentX        =   18071
      _ExtentY        =   926
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   1200
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   90
         Width           =   1305
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   3945
         MaxLength       =   50
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   90
         Width           =   6180
      End
      Begin xrControl.xrFrame xrFrame3 
         Height          =   525
         Index           =   0
         Left            =   5745
         Tag             =   "wt0;fb0"
         Top             =   31500
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   926
         BackColor       =   12632256
         ClipControls    =   0   'False
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Discount"
            Height          =   285
            Index           =   20
            Left            =   75
            TabIndex        =   59
            Top             =   165
            Width           =   675
         End
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Account No."
         Height          =   285
         Index           =   9
         Left            =   165
         TabIndex        =   0
         Top             =   135
         Width           =   930
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer &Name"
         Height          =   285
         Index           =   8
         Left            =   2715
         TabIndex        =   2
         Top             =   135
         Width           =   1125
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   10590
      TabIndex        =   56
      Top             =   3690
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
      Picture         =   "frmMPCustomerLedger.frx":0759
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   10590
      TabIndex        =   49
      Top             =   555
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
      Picture         =   "frmMPCustomerLedger.frx":0ED3
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   10590
      TabIndex        =   51
      Top             =   1815
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
      Picture         =   "frmMPCustomerLedger.frx":164D
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   10590
      TabIndex        =   52
      Top             =   555
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
      Picture         =   "frmMPCustomerLedger.frx":1DC7
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   10590
      TabIndex        =   57
      Top             =   4320
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
      Picture         =   "frmMPCustomerLedger.frx":2541
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   2295
      Index           =   2
      Left            =   5595
      Tag             =   "wt0;fb0"
      Top             =   1095
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   4048
      ClipControls    =   0   'False
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Maturity:"
         Height          =   285
         Index           =   17
         Left            =   60
         TabIndex        =   22
         Top             =   960
         Width           =   945
      End
      Begin VB.Label lblFields 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Active"
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
         Left            =   1020
         TabIndex        =   27
         Tag             =   "tc0"
         Top             =   1545
         Width           =   1290
      End
      Begin VB.Label lblFields 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "3,456.00"
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
         Left            =   1020
         TabIndex        =   25
         Tag             =   "tc0"
         Top             =   1260
         Width           =   1290
      End
      Begin VB.Label lblFields 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Oct 02, 2005"
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
         Left            =   1020
         TabIndex        =   23
         Tag             =   "tc0"
         Top             =   945
         Width           =   1290
      End
      Begin VB.Label lblFields 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "36 months"
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
         Left            =   1020
         TabIndex        =   21
         Tag             =   "tc0"
         Top             =   660
         Width           =   1290
      End
      Begin VB.Label lblFields 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Nov 02, 2005"
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
         Left            =   1020
         TabIndex        =   19
         Tag             =   "tc0"
         Top             =   420
         Width           =   1290
      End
      Begin VB.Label lblFields 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Sep 30, 2005"
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
         Left            =   1020
         TabIndex        =   17
         Tag             =   "tc0"
         Top             =   150
         Width           =   1290
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "1st Pay Date:"
         Height          =   285
         Index           =   11
         Left            =   60
         TabIndex        =   18
         Top             =   405
         Width           =   960
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Term:"
         Height          =   285
         Index           =   10
         Left            =   60
         TabIndex        =   20
         Top             =   690
         Width           =   945
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Status:"
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   60
         TabIndex        =   26
         Top             =   1530
         Width           =   945
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Mo. Instal:"
         Height          =   285
         Index           =   5
         Left            =   60
         TabIndex        =   24
         Top             =   1245
         Width           =   945
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Acct Date:"
         Height          =   285
         Index           =   2
         Left            =   60
         TabIndex        =   16
         Top             =   135
         Width           =   945
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   2295
      Index           =   3
      Left            =   8040
      Tag             =   "wt0;fb0"
      Top             =   1095
      Width           =   2310
      _ExtentX        =   4075
      _ExtentY        =   4048
      BackColor       =   -2147483644
      Enabled         =   0   'False
      ClipControls    =   0   'False
      Begin VB.Label lblFields 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   285
         Index           =   15
         Left            =   960
         TabIndex        =   39
         Tag             =   "tc0"
         Top             =   1545
         Width           =   1245
      End
      Begin VB.Label lblFields 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   285
         Index           =   11
         Left            =   960
         TabIndex        =   31
         Tag             =   "tc0"
         Top             =   405
         Width           =   1245
      End
      Begin VB.Label lblFields 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   285
         Index           =   12
         Left            =   960
         TabIndex        =   33
         Tag             =   "tc0"
         Top             =   660
         Width           =   1245
      End
      Begin VB.Label lblFields 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   285
         Index           =   13
         Left            =   960
         TabIndex        =   35
         Tag             =   "tc0"
         Top             =   945
         Width           =   1245
      End
      Begin VB.Label lblFields 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   285
         Index           =   14
         Left            =   960
         TabIndex        =   37
         Tag             =   "tc0"
         Top             =   1260
         Width           =   1245
      End
      Begin VB.Label lblFields 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   285
         Index           =   10
         Left            =   960
         TabIndex        =   29
         Tag             =   "tc0"
         Top             =   150
         Width           =   1245
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Reb. Guide:"
         Height          =   285
         Index           =   18
         Left            =   45
         TabIndex        =   38
         Top             =   1530
         Width           =   945
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Down Paym:"
         Height          =   285
         Index           =   16
         Left            =   45
         TabIndex        =   30
         Top             =   405
         Width           =   945
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cash Bal:"
         Height          =   285
         Index           =   15
         Left            =   45
         TabIndex        =   32
         Top             =   690
         Width           =   945
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pen. Guide:"
         Height          =   285
         Index           =   14
         Left            =   45
         TabIndex        =   36
         Top             =   1245
         Width           =   945
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "P/N Value:"
         Height          =   285
         Index           =   13
         Left            =   45
         TabIndex        =   34
         Top             =   960
         Width           =   945
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Gross Price:"
         Height          =   285
         Index           =   12
         Left            =   45
         TabIndex        =   28
         Top             =   135
         Width           =   945
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   10590
      TabIndex        =   53
      Top             =   1185
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Branch"
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
      Picture         =   "frmMPCustomerLedger.frx":2CBB
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   585
      Index           =   6
      Left            =   10590
      TabIndex        =   54
      Top             =   1815
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1032
      Caption         =   "C&ollectr"
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
      Picture         =   "frmMPCustomerLedger.frx":3435
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   7
      Left            =   10590
      TabIndex        =   58
      Top             =   3690
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
      Picture         =   "frmMPCustomerLedger.frx":3BAF
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   465
      Index           =   4
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   6870
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   820
      BackColor       =   12632256
      Enabled         =   0   'False
      ClipControls    =   0   'False
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Index           =   6
         Left            =   2925
         TabIndex        =   42
         Text            =   "Text1"
         Top             =   60
         Width           =   465
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "No of Paym"
         Height          =   285
         Index           =   22
         Left            =   1995
         TabIndex        =   41
         Top             =   120
         Width           =   840
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   465
      Index           =   5
      Left            =   3615
      Tag             =   "wt0;fb0"
      Top             =   6870
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   820
      BackColor       =   12632256
      Enabled         =   0   'False
      ClipControls    =   0   'False
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Index           =   3
         Left            =   810
         TabIndex        =   44
         Text            =   "Text1"
         Top             =   60
         Width           =   1305
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Payment"
         Height          =   285
         Index           =   21
         Left            =   60
         TabIndex        =   43
         Top             =   120
         Width           =   675
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   465
      Index           =   6
      Left            =   5850
      Tag             =   "wt0;fb0"
      Top             =   6870
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   820
      BackColor       =   12632256
      Enabled         =   0   'False
      ClipControls    =   0   'False
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Index           =   4
         Left            =   825
         TabIndex        =   46
         Text            =   "Text1"
         Top             =   60
         Width           =   1305
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Discount"
         Height          =   285
         Index           =   19
         Left            =   75
         TabIndex        =   45
         Top             =   120
         Width           =   675
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   465
      Index           =   7
      Left            =   8085
      Tag             =   "wt0;fb0"
      Top             =   6870
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   820
      BackColor       =   12632256
      Enabled         =   0   'False
      ClipControls    =   0   'False
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Index           =   5
         Left            =   660
         TabIndex        =   48
         Tag             =   "ht0;hb0"
         Text            =   "Text1"
         Top             =   60
         Width           =   1515
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   24
         Left            =   75
         TabIndex        =   47
         Top             =   105
         Width           =   570
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   10
      Left            =   10590
      TabIndex        =   55
      Top             =   3060
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Print"
      AccessKey       =   "P"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmMPCustomerLedger.frx":4329
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   9
      Left            =   10590
      TabIndex        =   64
      Top             =   2430
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Impound"
      AccessKey       =   "I"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmMPCustomerLedger.frx":4AA3
   End
End
Attribute VB_Name = "frmMPCustomerLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmCustomerLedger"
'Ok
Private oSkin As clsFormSkin
Private oFormImpounded As frmImpounded
Private oRSMaster As ADODB.Recordset
Private oRSDetail As ADODB.Recordset

Dim pnDebtTotl As Double, pnCredTotl As Double, pnDownTotl As Double, pnCashTotl As Double
Dim psMatrixValue As String, psSelected() As String, psAcctNo As String
Dim pnRebTotlx As Double, pnPenTotlx As Double, pnPaymTotl As Double
Dim pbGridFocus As Boolean, pbEditMode As Boolean
Dim pnItemCount As Integer, pnCtr As Integer

Private Sub cmdButton_Click(Index As Integer)
   Dim lnCtr As Integer, lnRep As Long
   Dim lbCancel As Boolean
   Dim lrs As ADODB.Recordset
   Dim lsOldProc As String

   lsOldProc = "cmdButton_Click"
   'On Error GoTo errProc

   With GridEditor1
      Select Case Index
      Case 0 'Save
         If pnItemCount + 1 < .Rows Then
            pnCtr = pnItemCount
            Do While pnCtr < .Rows
               If Trim(.TextMatrix(pnCtr, 2)) = "" Or _
                  Trim(.TextMatrix(pnCtr, 3)) = "" Or _
                  (CDbl(.TextMatrix(pnCtr, 5)) = 0# And _
                  CDbl(.TextMatrix(pnCtr, 7)) = 0# And _
                  Trim(.TextMatrix(pnCtr, 3)) <> "DP" And _
                  oRSMaster("cLoanType") <> 3) Then
                  .Row = pnCtr
                  oRSDetail("nEntryNox") = pnCtr
                  If DeleteDatail(.Row - 1) Then .deleteRow
               Else
                  pnCtr = pnCtr + 1
               End If
            Loop

            If pnItemCount + 1 > oRSDetail.RecordCount Then
               InitValue

               .Rows = .Rows + 1
               .Row = .Rows - 1

               .TextMatrix(.Row, 1) = Format(oRSDetail.Fields("dTransact"), "MM/DD/YYYY")
               .TextMatrix(.Row, 5) = Format(oRSDetail.Fields("nTranAmtx"), "MM/DD/YYYY")
               .TextMatrix(.Row, 6) = Format(oRSDetail.Fields("nRebatesx"), "MM/DD/YYYY")
               .TextMatrix(.Row, 7) = Format(oRSDetail.Fields("nOthersxx"), "MM/DD/YYYY")
               .TextMatrix(.Row, 8) = Format(oRSDetail.Fields("nABalance"), "MM/DD/YYYY")
               .TextMatrix(.Row, 9) = Format(oRSDetail.Fields("nMonDelay"), "MM/DD/YYYY")
            End If

            .ColWidth(2) = 2700
            If .Rows > 16 Then .ColWidth(2) = 2500
         End If

         If Not isEntryOk Then Exit Sub
         If SaveTransaction Then
            MsgBox "Transaction Save Successfully!!!", vbInformation, "Notice"
            initButton xeModeReady
            txtField(0).SetFocus
         Else
            MsgBox "Unable to Save Transaction!!!" & vbCrLf & _
                   "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
         End If
      Case 1 'Cancel
         lnRep = MsgBox("Transaction is in Update Mode!!!" & vbCrLf & _
                        "Do you want to Cancel Transaction!!!", vbYesNo + vbQuestion, "Confirm")
         If lnRep = vbYes Then
            If SearchTransaction(psAcctNo, True) Then
               LoadMaster
               LoadDetail
               initButton xeModeReady

               .Row = 1
               .Col = 1

               txtField(0).SetFocus
               pbEditMode = False
            End If
         Else
            .SetFocus
         End If
      Case 2 'Browse
         If SearchTransaction(, , True) Then
            LoadMaster
            LoadDetail
         End If
         .Refresh
      Case 3 'Update
         If lblFields(0).Caption = "" Then
            MsgBox "No Account is Loaded to modify!!!" & vbCrLf & _
                   "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
         Else
            initButton xeModeUpdate
            InitValue
            .Rows = .Rows + 1

            .Row = .Rows - 1
            .TextMatrix(.Row, 1) = Format(oRSDetail.Fields("dTransact"), "MM/DD/YYYY")
            .TextMatrix(.Row, 5) = Format(oRSDetail.Fields("nTranAmtx"), "#,##0.00")
            .TextMatrix(.Row, 6) = Format(oRSDetail.Fields("nRebatesx"), "#,##0.00")
            .TextMatrix(.Row, 7) = Format(oRSDetail.Fields("nOthersxx"), "#,##0.00")
            .TextMatrix(.Row, 8) = Format(oRSDetail.Fields("nABalance"), "#,##0.00")

            .TopRow = .Rows - 1
            .Col = 1
            .SetFocus
            .ColWidth(2) = 2700
            If .Rows > 16 Then .ColWidth(2) = 2500

            .TextMatrix(.Row, 9) = Format(getDelay, "#,##0.00")
            txtField(6).Text = oRSDetail.RecordCount
            pbEditMode = True
         End If
      Case 4 'Close
         Unload Me
      Case 5 'Branch
         If Not pbGridFocus Then Exit Sub
         If .Col = 2 Then
            .TextMatrix(.Row, .Col) = SearhBranch(False, "")
            .Refresh
         End If
      Case 6 'Collectr
         If Not pbGridFocus Then Exit Sub
         If .Col = 2 Then
            .TextMatrix(.Row, .Col) = SearchCollector(False, "")
            .Refresh
         End If
      Case 7 'Delete Row
         If .Row > pnItemCount Then
            If .Rows > 2 Then
               If DeleteDatail(.Row - 1) Then .deleteRow

               .ColWidth(2) = 2700
               If .Rows > 16 Then .ColWidth(2) = 2500
            End If
         End If
      Case 8 'ReCalculate
         If lblFields(0).Caption = "" Then
            MsgBox "No Account is Loaded to modify!!!" & vbCrLf & _
                   "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
            GoTo endProc
         End If

         If reCalc() Then
            MsgBox "Transaction Updated Successfully!!!", vbInformation, "Notice"
            If SearchTransaction(oRSMaster("sAcctNmbr"), True, False) Then
               LoadMaster
               LoadDetail
            End If
         End If
      Case 9 'Impounded
         If Trim(txtField(0).Text) = "" Then
            MsgBox "No Account is Loaded!!!" & vbCrLf & _
                   "Please Verify your Entry then Try Again!!!", vbInformation, "Notice"
            GoTo endProc
         End If

         Set lrs = New ADODB.Recordset
         lrs.Open "Select" _
                     & " sAcctNmbr" _
                  & " From Impound" _
                  & " Where sAcctNmbr = " & strParm(psSelected(0)) _
         , oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText

         If lrs.EOF Then
            MsgBox "No Record found!!!" & vbCrLf & _
                   "Customer not yet Impounded!!!", vbInformation, "Notice"
         Else
            With oFormImpounded
               .AcctNumber = lrs("sAcctNmbr")
               .CustName = lblFields(1).Caption
               .Show 1
            End With
         End If

         .Refresh
      Case 10 ' Print
         If lblFields(0).Caption = "" Then Exit Sub

         Dim loMCARPrinting As clsMCARPrinting

         Set loMCARPrinting = New clsMCARPrinting
         Set loMCARPrinting.AppDriver = oApp

         lnRep = MsgBox("Are you sure you want to print this ledger...", vbInformation + vbYesNo)
         If lnRep = vbNo Then Exit Sub

         loMCARPrinting.AcctNumber = oRSMaster("sAcctNmbr")
         loMCARPrinting.PrintTransaction

         Set loMCARPrinting = Nothing
      End Select
   End With

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   GridEditor1.Refresh
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String

   lsOldProc = "Form_Load"
   'On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oFormImpounded = New frmImpounded

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualRight

   InitGrid
   InitEntry
   initButton xeModeReady

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
   Set oRSMaster = Nothing
   Set oRSDetail = Nothing
   Set oFormImpounded = Nothing
End Sub

Private Sub GridEditor1_AddingRow(Cancel As Boolean)
   With GridEditor1
      If Trim(.TextMatrix(.Row, 2)) = "" Then
         Cancel = True
      ElseIf Trim(.TextMatrix(.Row, 3)) = "" Then
         Cancel = True
      ElseIf CDbl(.TextMatrix(.Row, 5)) = 0# And CDbl(.TextMatrix(.Row, 7)) = 0# Then
         Cancel = True
      End If

      If Not Cancel Then InitValue

      .ColWidth(2) = 2700
      If .Rows > 16 Then .ColWidth(2) = 2500
   End With
End Sub

Private Sub GridEditor1_EditorValidate(Cancel As Boolean)
   Dim lsOldProc As String

   lsOldProc = "GridEditor1_EditorValidate"
   'On Error GoTo errProc

   If pbEditMode Then
      With GridEditor1
         If .Row <= pnItemCount Then
            .TextMatrix(.Row, .Col) = psMatrixValue
            MsgBox "Unable to Update current row!!!" & vbCrLf & _
                   "Transaction already posted!!!", vbCritical, "Warning"
            Exit Sub
         End If

         Select Case .Col
         Case 1
            If Not IsDate(.TextMatrix(.Row, .Col)) Then
               .TextMatrix(.Row, .Col) = Format(oApp.ServerDate, "MM/DD/YYYY")
            Else
               .TextMatrix(.Row, .Col) = Format(.TextMatrix(.Row, .Col), "MM/DD/YYYY")
            End If

            oRSDetail(.Col) = CDate(.TextMatrix(.Row, .Col))
         Case 2
            If Trim(.TextMatrix(.Row, .Col)) = "" Then
               oRSDetail("sCollIDxx") = Null
            Else
               If IsNull(oRSDetail("sCollIDxx")) And IsNull(oRSDetail("sBranchCd")) Then
                  .TextMatrix(.Row, .Col) = SearchCollector(True, .TextMatrix(.Row, .Col))
               ElseIf .TextMatrix(.Row, .Col) <> .Tag Then
                  .TextMatrix(.Row, .Col) = SearchCollector(True, .TextMatrix(.Row, .Col))
               End If
            End If

            If IsNull(oRSDetail("sBranchCd")) Then oRSDetail("sBranchCd") = oRSMaster("sBranchCd")
            .Tag = .TextMatrix(.Row, .Col)
         Case 3
            Select Case LCase(Trim(.TextMatrix(.Row, .Col)))
            Case "p", "mp"
               .TextMatrix(.Row, .Col) = TranType("p")
               oRSDetail(.Col) = "p"
            Case "d", "dp"
               .TextMatrix(.Row, .Col) = TranType("d")
               oRSDetail(.Col) = "d"
            Case "m", "dm"
               .TextMatrix(.Row, .Col) = TranType("m")
               oRSDetail(.Col) = "m"
            Case "c", "cm"
               .TextMatrix(.Row, .Col) = TranType("c")
               oRSDetail(.Col) = "c"
            Case "b", "cb"
               .TextMatrix(.Row, .Col) = TranType("b")
               oRSDetail(.Col) = "b"
            Case Else
               .TextMatrix(.Row, .Col) = ""
               oRSDetail(.Col) = Empty
            End Select
         Case 5
            If Not IsNumeric(.TextMatrix(.Row, .Col)) Then
               .TextMatrix(.Row, .Col) = 0#
            Else
               .TextMatrix(.Row, .Col) = Format(.TextMatrix(.Row, .Col), "#,##0.00")
            End If

            oRSDetail(.Col) = 0#
            oRSDetail("nDebitAmt") = 0#

            If LCase(Trim(.TextMatrix(.Row, 3))) = "mp" Then computeAmount
            If oRSDetail("cTranType") = "m" Then
               oRSDetail("nDebitAmt") = CDbl(.TextMatrix(.Row, .Col))
            Else
               oRSDetail(.Col) = CDbl(.TextMatrix(.Row, .Col))
            End If
         Case 6
            If Not IsNumeric(.TextMatrix(.Row, .Col)) Then
               .TextMatrix(.Row, .Col) = 0#
            Else
               .TextMatrix(.Row, .Col) = Format(.TextMatrix(.Row, .Col), "#,##0.00")
            End If

            If LCase(Trim(.TextMatrix(.Row, 3))) = "mp" Then
               computeAmount
            Else
               .TextMatrix(.Row, .Col) = 0#
            End If

            oRSDetail(.Col) = CDbl(.TextMatrix(.Row, .Col))
         Case 7
            If Not IsNumeric(.TextMatrix(.Row, .Col)) Then
               .TextMatrix(.Row, .Col) = 0#
            Else
               .TextMatrix(.Row, .Col) = Format(.TextMatrix(.Row, .Col), "#,##0.00")
            End If

            oRSDetail(.Col) = CDbl(.TextMatrix(.Row, .Col))
         Case Else
            oRSDetail(.Col) = .TextMatrix(.Row, .Col)
         End Select

         If .Col > 4 Then
            .TextMatrix(.Row, 9) = Format(getDelay, "#,##0.00")
            .TextMatrix(.Row, 8) = Format(CDbl(.TextMatrix(.Row - 1, 8) - (CDbl(.TextMatrix(.Row, 5)) + CDbl(.TextMatrix(.Row, 6)))), "#,##0.00")
            oRSDetail(8) = CDbl(.TextMatrix(.Row, 8))
            oRSDetail(9) = CDbl(.TextMatrix(.Row, 9))
         End If
      End With
   End If

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Cancel & " )", True
End Sub

Private Sub GridEditor1_GotFocus()
   pbGridFocus = True
End Sub

Private Sub GridEditor1_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String

   lsOldProc = "GridEditor1_KeyDown"
   'On Error GoTo errProc

   If KeyCode = vbKeyF3 Then
      With GridEditor1
         If .Col = 2 Then
            .TextMatrix(.Row, .Col) = SearchCollector(False, .TextMatrix(.Row, .Col))

             .Tag = .TextMatrix(.Row, .Col)
             .Col = .Col + 1
            .Refresh
            .SetFocus
         End If
         KeyCode = 0
      End With
   ElseIf KeyCode = vbKeyF4 Then
      With GridEditor1
         If .Col = 2 Then
            .TextMatrix(.Row, .Col) = SearhBranch(False, .TextMatrix(.Row, .Col))

            .Tag = .TextMatrix(.Row, .Col)
            .Col = .Col + 1
            .Refresh
            .SetFocus
         End If
         KeyCode = 0
      End With
   End If

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " _
                       & "  " & KeyCode _
                       & ", " & Shift _
                       & " )", True
End Sub

Private Sub InitEntry()
   For pnCtr = 0 To txtField.Count - 1
      Select Case pnCtr
      Case 0, 1
         txtField(pnCtr).Text = ""
         txtField(pnCtr).Tag = ""
      Case 3 To 5, 7
         txtField(pnCtr).Text = "0.00"
      Case 6
         txtField(pnCtr).Text = 0
      Case Else
         txtField(pnCtr).Text = ""
      End Select
   Next

   For pnCtr = 0 To lblFields.Count - 1
      Select Case pnCtr
      Case 10 To 15
         lblFields(pnCtr).Caption = "0.00"
      Case Else
         lblFields(pnCtr).Caption = ""
      End Select
   Next

   With GridEditor1
      .Rows = 2
      .ColWidth(2) = 2700

      .TextMatrix(1, 1) = ""
      .TextMatrix(1, 2) = ""
      .TextMatrix(1, 3) = ""
      .TextMatrix(1, 4) = ""
      .TextMatrix(1, 5) = 0#
      .TextMatrix(1, 6) = 0#
      .TextMatrix(1, 7) = 0#
      .TextMatrix(1, 8) = 0#
      .TextMatrix(1, 9) = 0#
      .TextMatrix(1, 10) = ""
   End With

   psAcctNo = Empty
   pbEditMode = False
   pnItemCount = 0
End Sub

Private Sub InitGrid()
   With GridEditor1
      .Rows = 2
      .Cols = 11
      .Font = "MS Sans Serif"

      'column title
      .TextMatrix(0, 1) = "Date"
      .TextMatrix(0, 2) = "Collector"
      .TextMatrix(0, 3) = "TC"
      .TextMatrix(0, 4) = "OR NO."
      .TextMatrix(0, 5) = "Amount"
      .TextMatrix(0, 6) = "Rebate"
      .TextMatrix(0, 7) = "Penalty"
      .TextMatrix(0, 8) = "Balance"
      .TextMatrix(0, 9) = "Delay"
      .TextMatrix(0, 10) = "Remarks"
      .Row = 0

      'column alignment
      For pnCtr = 0 To .Cols - 1
         .Col = pnCtr
         .CellFontBold = True
         .CellAlignment = 3
      Next

      .ColEnabled(8) = False
      .ColEnabled(9) = False

      'column width
      .ColWidth(0) = 330
      .ColWidth(1) = 960
      .ColWidth(3) = 400
      .ColWidth(4) = 900
      .ColWidth(5) = 1100
      .ColWidth(6) = 900
      .ColWidth(7) = 1100
      .ColWidth(8) = 900
      .ColWidth(9) = 900
      .ColWidth(10) = 2500

      .ColFormat(1) = "MM/DD/YYYY"
      .ColFormat(3) = ">"
      .ColFormat(4) = ">"

      .ColAlignment(3) = 1
      .ColAlignment(4) = 1
      .ColAlignment(10) = 1

      .ColDefault(1) = Format(oApp.ServerDate, "MM/DD/YYYY")

      .ColLimit(3) = 2
      .ColLimit(4) = 8
      .ColLimit(10) = 30

      For pnCtr = 5 To 9
         .ColDefault(pnCtr) = "0.00"
         .ColFormat(pnCtr) = "#,##0.00"
         .ColNumberOnly(pnCtr) = True
      Next

      .ColMaxValue(5) = 999999.99
      .ColMaxValue(6) = 9999.99
      .ColMaxValue(7) = 999999.99
      .ColMaxValue(8) = 999999.99
      .ColMaxValue(9) = 999999.99

      .Row = 1
      .Col = 1
   End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyReturn, vbKeyUp, vbKeyDown
      Select Case KeyCode
      Case vbKeyReturn, vbKeyDown
         If GetFocus = GridEditor1.hwnd Or _
         GetFocus = txtField(0).hwnd Or _
         GetFocus = txtField(1).hwnd Then Exit Sub
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
   cmdButton(5).Visible = lbShow
   cmdButton(6).Visible = lbShow
   cmdButton(7).Visible = lbShow

   With GridEditor1
      For pnCtr = 1 To .Cols - 1
         Select Case pnCtr
         Case 1 To 7, 10
            .ColEnabled(pnCtr) = lbShow
         Case 8, 9
         End Select
      Next
   End With

   xrFrame1(1).Enabled = Not lbShow

   cmdButton(2).Visible = Not lbShow
   cmdButton(3).Visible = Not lbShow
   cmdButton(4).Visible = Not lbShow
   cmdButton(8).Visible = Not lbShow
End Sub

Private Sub GridEditor1_LostFocus()
   With GridEditor1
      .LeftCol = 1
      .Col = 1
   End With
End Sub

Private Sub GridEditor1_RowAdded()
   With GridEditor1
      .TextMatrix(.Row, 8) = .TextMatrix(.Rows - 1, 8)
      .TextMatrix(.Row, 9) = Format(getDelay, "#,##0.00")
      txtField(6).Text = oRSDetail.RecordCount
   End With
End Sub

Private Sub GridEditor1_RowColChange()
   Dim lnCol As Integer

   If Not pbEditMode Then Exit Sub
   With GridEditor1
      For lnCol = 1 To .Cols - 1
         Select Case lnCol
         Case 1 To 7, 10
            .ColEnabled(lnCol) = False
            If .Row > pnItemCount Then .ColEnabled(lnCol) = True
         Case 8, 9
         End Select
      Next

      psMatrixValue = .TextMatrix(.Row, .Col)
      If .Row > pnItemCount Then
         .Tag = .TextMatrix(.Row, 2)
         oRSDetail.Move .Row - 1, adBookmarkFirst
      End If
   End With
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
   pbGridFocus = False
End Sub

Private Function SearchTransaction(Optional sValue As Variant, Optional bByCode As Variant, Optional bSearch As Variant) As Boolean
   Dim lsOldProc As String
   Dim lsBrowse As String
   Dim lsSQL As String

   lsOldProc = "SearchTransaction"
   'On Error GoTo errProc
   SearchTransaction = False

   Set oRSMaster = New ADODB.Recordset

   lsSQL = "Select" _
               & "  a.sAcctNmbr" _
               & ", CONCAT(b.sLastName, ', ', b.sFrstName, ' ', b.sMiddName) xFullName" _
               & ", CONCAT(g.sBrandNme, ' ', f.sModelNme, ' ( ', RTrim(e.sEngineNo), ' )') xModelNme" _
               & ", CONCAT(i.sLastName, ', ', i.sFrstName, ' ', i.sMiddName) xCollectr" _
               & ", a.dPurchase" _
               & ", a.dFirstPay" _
               & ", a.nAcctTerm" _
               & ", a.dDueDatex" _
               & ", a.nMonAmort" _
               & ", a.cAcctStat" _
               & ", a.nGrossPrc" _
               & ", a.nDownPaym" _
               & ", a.nCashBalx" _
               & ", a.nPNValuex" _
               & ", a.nPenaltyx" _
               & ", a.nRebatesx" _
               & ", j.sBranchNm" _
               & ", a.nPaymTotl" _
               & ", a.nRebTotlx" _
               & ", a.nPenTotlx" _
               & ", CONCAT(b.sAddressx, ', ', c.sTownName, ', ', d.sProvName, ' ', c.sZippCode) xAddressx" _
               & ", h.sBranchCd" _
               & ", a.nABalance"
   lsSQL = lsSQL _
               & ", a.nDebtTotl" _
               & ", a.nCredTotl" _
               & ", a.nDownTotl" _
               & ", a.nCashTotl" _
               & ", a.cRatingxx" _
               & ", a.nLastPaym" _
               & ", a.dLastPaym" _
               & ", a.nAmtDuexx" _
               & ", a.nDelayAvg" _
               & ", a.nLedgerNo" _
               & ", a.sModified" _
               & ", a.dModified" _
               & ", a.dClosedxx" _
               & ", CONCAT(k.sLastName, ', ', k.sFrstName, ' ', k.sMiddName) xCoCltNm1" _
               & ", CONCAT(l.sLastName, ', ', l.sFrstName, ' ', l.sMiddName) xCoCltNm2" _
               & ", a.cLoanType" _
               & ", CONCAT(n.sLastName, ', ', n.sFrstName, ' ', n.sMiddName) xCoMakrNm" _
               & ", a.nLgrLinex" _
               & ", a.nPassLine" _
               & ", CONCAT(p.sLastName, ', ', p.sFrstName, ' ', p.sMiddName) zCollectr" _
               & ", a.cLoanType"

   lsSQL = lsSQL _
            & " From MC_AR_Master a" _
               & " LEFT JOIN MC_Serial e" _
                  & " On a.sSerialID = e.sSerialID" _
               & " Left Join MC_Model f" _
                  & " On e.sModelIDx = f.sModelIDx" _
               & " Left Join Brand g" _
                  & " On f.sBrandIDx = g.sBrandIDx" _
               & " Left Join Client_Master k" _
                  & " On a.sCoCltID1 = k.sClientID" _
               & " Left Join Client_Master l" _
                  & " On a.sCoCltID2 = l.sClientID" _
               & " Left Join MC_Credit_Application m" _
                  & " Left Join Client_Master n" _
                     & " On m.sCoMakrID = n.sClientID" _
                  & " On a.sApplicNo = m.sTransNox" _
                  & " And a.sClientID = m.sClientID" _
               & ", Client_Master b" _
                  & " Left Join TownCity c" _
                     & " On b.sTownIDxx = c.sTownIDxx" _
                  & " Left Join Province d" _
                     & " On c.sProvIDxx = d.sProvIDxx"
   lsSQL = lsSQL _
               & ", Route_Area h" _
                  & " LEFT JOIN Employee_Master i" _
                     & " ON h.sCollctID = i.sEmployID" _
                  & " LEFT JOIN Employee_Master001 o" _
                     & " LEFT JOIN Client_Master p" _
                        & " ON o.sEmployID = p.sClientID" _
                     & " ON h.sCollctID = o.sEmployID" _
               & ", Branch j"

   lsSQL = lsSQL _
            & " Where a.sClientID = b.sClientID" _
               & " And a.sRouteIDx = h.sRouteIDx" _
               & " And h.sBranchCd = j.sBranchCd" _
               & " And a.cAcctstat = '0'" _
               & " AND a.cLoanType = '4'"
      If Not IsMissing(sValue) Then
      If Not IsMissing(bByCode) Then
         If bByCode Then
            lsSQL = lsSQL & " And a.sAcctNmbr = " & strParm(sValue)
         Else
            lsSQL = lsSQL & " And CONCAT(b.sLastName, ', ', b.sFrstName, ' ', b.sMiddName) Like " & strParm(Trim(sValue) & "%")
         End If
      Else
         lsSQL = lsSQL & " And CONCAT(b.sLastName, ', ', b.sFrstName, ' ', b.sMiddName) = " & strParm(Trim(sValue))
      End If
   End If
   lsSQL = lsSQL & " Order By sAcctNmbr, xFullName"
Debug.Print lsSQL
   oRSMaster.Open lsSQL, oApp.Connection, adOpenStatic, adLockOptimistic, adCmdText

   If oRSMaster.EOF Then
      InitEntry
      GoTo endProc
   ElseIf oRSMaster.RecordCount = 1 Then
      ReDim psSelected(oRSMaster.Fields.Count - 1) As String

      For pnCtr = 0 To oRSMaster.Fields.Count - 1
         psSelected(pnCtr) = IIf(IsNull(oRSMaster(pnCtr)), "", oRSMaster(pnCtr))
      Next
      psAcctNo = psSelected(0)
   Else
      lsBrowse = KwikBrowse(oApp _
                              , oRSMaster _
                              , "sAcctNmbr»xFullName»xModelNme»xCollectr" _
                              , "Acct. No.»Customer Name»Model»Collector" _
                              , "@»@»@»@" _
                              , "a.sAcctNmbr»" _
                              & "CONCAT(b.sLastName, ', ', b.sLastName, ' ', b.sMiddName)»" _
                              & "CONCAT(g.sBrandNme, ' ', f.sModelNme)»" _
                              & "CONCAT(i.sLastName, ', ', i.sFrstName, ' ', i.sMiddName)")

      If lsBrowse = "" Then GoTo endProc
      psSelected = Split(lsBrowse, "»")
      psAcctNo = psSelected(0)
   End If
   SearchTransaction = True
   Set oRSMaster.ActiveConnection = Nothing

endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " _
                       & "  " & IFNull(sValue) _
                       & ", " & IFNull(bByCode) _
                       & ", " & IFNull(bSearch) _
                       & " )"
End Function

Private Sub LoadMaster()
   For pnCtr = 0 To 20
      Select Case pnCtr
      Case 0, 1
         lblFields(pnCtr).Caption = IIf(psSelected(pnCtr) = "", "", psSelected(pnCtr))
         txtField(pnCtr).Text = lblFields(pnCtr).Caption
         txtField(pnCtr).Tag = txtField(pnCtr).Text
      Case 3
         If psSelected(pnCtr) = "" Then
            lblFields(pnCtr).Caption = psSelected(42)
         Else
            lblFields(pnCtr).Caption = psSelected(pnCtr)
         End If
      Case 4, 5, 7
         lblFields(pnCtr).Caption = IIf(psSelected(pnCtr) = "", "", Format(psSelected(pnCtr), "MMM DD, YYYY"))
      Case 6
         lblFields(pnCtr).Caption = IIf(psSelected(pnCtr) = "", "", psSelected(pnCtr) & " months")
      Case 8, 10 To 15
         lblFields(pnCtr).Caption = IIf(psSelected(pnCtr) = "", "0.00", Format(psSelected(pnCtr), "#,##0.00"))
      Case 9
         If CDate(lblFields(7)) < oApp.ServerDate Then
            lblFields(pnCtr).Caption = "PastDue"
         Else
            lblFields(pnCtr).Caption = AcctStat(psSelected(pnCtr))
         End If
      Case 17, 18, 19
         txtField(pnCtr - 14).Text = IIf(psSelected(pnCtr) = "", "0.00", Format(psSelected(pnCtr), "#,##0.00"))
      Case 20
         txtField(2).Text = IIf(psSelected(pnCtr) = "", "", psSelected(pnCtr))
      Case Else
         lblFields(pnCtr).Caption = IIf(psSelected(pnCtr) = "", "", psSelected(pnCtr))
      End Select
   Next
   txtField(5) = Format(CDbl(txtField(3)) + CDbl(txtField(4)), "#,##0.00")

   lblFields(17).Caption = IIf(psSelected(35) = "", "N-O-N-E", psSelected(35))
   lblFields(18).Caption = IIf(psSelected(36) = "", "N-O-N-E", psSelected(36))

   pnPaymTotl = CDbl(psSelected(17))
   pnRebTotlx = CDbl(psSelected(18))
   pnPenTotlx = CDbl(psSelected(19))
   pnDebtTotl = CDbl(psSelected(23))
   pnCredTotl = CDbl(psSelected(24))
   pnDownTotl = CDbl(psSelected(25))
   pnCashTotl = CDbl(psSelected(26))
End Sub

Private Sub LoadDetail()
   Dim lsOldProc As String
   Dim lsSQL As String
   Dim lnCol As Integer

   lsOldProc = "LoadDetail"
   'On Error GoTo errProc

   Set oRSDetail = New ADODB.Recordset

   lsSQL = "Select" _
               & "  a.sAcctNmbr" _
               & ", a.dTransact" _
               & ", CONCAT(IFNULL(b.sFrstName, e.sFrstName), ' ', IFNULL(b.sLastName, e.sLastName)) xCollectr" _
               & ", a.cTrantype" _
               & ", a.sORNoxxxx" _
               & ", a.nTranAmtx" _
               & ", a.nRebatesx" _
               & ", a.nOthersxx" _
               & ", a.nABalance" _
               & ", a.nMonDelay" _
               & ", a.sRemarksx" _
               & ", c.sBranchNm" _
               & ", a.nEntryNox" _
               & ", a.sCollIDxx" _
               & ", a.nAmtDuexx" _
               & ", a.nDebitAmt" _
               & ", a.sBranchCd" _
               & ", a.cOffPaymx"
   lsSQL = lsSQL _
            & " From MC_AR_Ledger a" _
               & " Left Join Employee_Master b" _
                  & " On a.sCollIDxx = b.sEmployID" _
               & " LEFT JOIN Employee_Master001 d" _
                  & " LEFT JOIN Client_Master e" _
                     & " ON d.sEmployID = e.sClientID" _
                  & " ON a.sCollIDxx = d.sEmployID" _
               & " , Branch c" _
            & " Where a.sAcctNmbr = " & strParm(psAcctNo) _
               & " And a.sBranchCd = c.sBranchCd" _
            & " Order By a.dTransact, a.sORNoxxxx"

'   lsSQL = "SELECT * FROM (" & lsSQL & ") x"
Debug.Print lsSQL
   oRSDetail.Open lsSQL, oApp.Connection, adOpenStatic, adLockOptimistic, adCmdText
   If oRSDetail.EOF Then
      With GridEditor1
         .Rows = 2
         .ColWidth(2) = 2700

         .TextMatrix(1, 1) = ""
         .TextMatrix(1, 2) = ""
         .TextMatrix(1, 3) = ""
         .TextMatrix(1, 4) = ""
         .TextMatrix(1, 5) = 0#
         .TextMatrix(1, 6) = 0#
         .TextMatrix(1, 7) = 0#
         .TextMatrix(1, 8) = 0#
         .TextMatrix(1, 9) = 0#
         .TextMatrix(1, 10) = ""
      End With

      GoTo endProc
   End If

   pnItemCount = oRSDetail.RecordCount
   With GridEditor1
      .Rows = oRSDetail.RecordCount + 1

      txtField(6).Text = oRSDetail.RecordCount
      For pnCtr = 0 To oRSDetail.RecordCount - 1
         For lnCol = 1 To .Cols - 1
            Select Case lnCol
            Case 1
               .TextMatrix(pnCtr + 1, lnCol) = IIf(IsNull(oRSDetail(lnCol)), "", Format(oRSDetail(lnCol), "MM/DD/YYYY"))
            Case 2
               '.TextMatrix(pnCtr + 1, lnCol) = IIf(IsNull(oRSDetail(lnCol)), oRSDetail("sBranchNm"), IIf(oRSDetail(lnCol) = "", oRSDetail("sBranchNm"), oRSDetail(lnCol)))
               .TextMatrix(pnCtr + 1, lnCol) = IIf(IsNull(oRSDetail(lnCol)), oRSDetail("sBranchNm"), IIf(oRSDetail(lnCol) = "", oRSDetail("xCollectr"), oRSDetail(lnCol)))
            Case 3
               .TextMatrix(pnCtr + 1, lnCol) = IIf(IsNull(oRSDetail(lnCol)), "", TranType(oRSDetail(lnCol)))
            Case 5
               If oRSDetail("cTrantype") = "m" Then
                  .TextMatrix(pnCtr + 1, lnCol) = IIf(IsNull(oRSDetail("nDebitAmt")), 0#, Format(oRSDetail("nDebitAmt"), "#,##0.00"))
               Else
                  .TextMatrix(pnCtr + 1, lnCol) = IIf(IsNull(oRSDetail(lnCol)), 0#, Format(oRSDetail(lnCol), "#,##0.00"))
               End If
            Case 6, 7, 8, 9
               .TextMatrix(pnCtr + 1, lnCol) = IIf(IsNull(oRSDetail(lnCol)), 0#, Format(oRSDetail(lnCol), "#,##0.00"))
            Case Else
               .TextMatrix(pnCtr + 1, lnCol) = IIf(IsNull(oRSDetail(lnCol)), "", oRSDetail(lnCol))
            End Select
         Next
         oRSDetail.MoveNext
      Next

      .ColWidth(2) = 2700
      .TopRow = .Rows - 1
      If .Rows > 16 Then .ColWidth(2) = 2500
   End With

   Set oRSDetail.ActiveConnection = Nothing

endProc:

   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )"
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

Private Function AcctStat(ByVal sValue As String) As String
   Select Case sValue
   Case 0
      AcctStat = "Open"
   Case 1
      AcctStat = "Closed"
   Case 2
      AcctStat = "Dead"
   Case 3
      AcctStat = "Impounded"
   Case 4
      AcctStat = "Discarded"
   Case 5
      AcctStat = "Rejected"
   Case Else
      AcctStat = "Unknown"
   End Select
End Function

Private Function TranType(ByVal sValue As String) As String
   Select Case LCase(sValue)
   Case "p"
      TranType = "MP"
   Case "d"
      TranType = "DP"
   Case "m"
      TranType = "DM"
   Case "c"
      TranType = "CM"
   Case "b"
      TranType = "CB"
   Case "r"
      TranType = "Rl"
   Case Else
      TranType = ""
   End Select
End Function

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String

   lsOldProc = "txtField_KeyDown"
   'On Error GoTo errProc

   If KeyCode = vbKeyReturn Or KeyCode = vbKeyF3 Then
      If Trim(txtField(Index).Text) = "" Then
         InitEntry
         Exit Sub
      End If

      If txtField(Index).Text <> txtField(Index).Tag Then
         If SearchTransaction(txtField(Index).Text, IIf(Index = 0, True, False), False) Then
            LoadMaster
            LoadDetail
         Else
            InitEntry
         End If
      End If

      txtField(Index).SelStart = 0
      txtField(Index).SelLength = Len(txtField(Index).Text)
      GridEditor1.Refresh
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
   'On Error GoTo errProc

   With txtField(Index)
      If Trim(.Text) = "" Then
         InitEntry
         Exit Sub
      End If

      If .Text <> .Tag Then
         If SearchTransaction(.Text, IIf(Index = 0, True, False), False) Then
            LoadMaster
            LoadDetail
         Else
            InitEntry
         End If
      End If

      .Tag = .Text
   End With
   GridEditor1.Refresh

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " _
                       & "  " & Index _
                       & ", " & Cancel _
                       & " )", True
End Sub

Private Function SearhBranch(Optional bSearch As Variant, Optional sValue As Variant) As String
   Dim lrs As ADODB.Recordset
   Dim lsSelected() As String
   Dim lsOldProc As String
   Dim lsBrowse As String
   Dim lsSQL As String

   lsOldProc = "SearchBranch"
   'On Error GoTo errProc
   SearhBranch = ""

   Set lrs = New ADODB.Recordset

   lsSQL = "Select" _
               & "  a.sBranchCd" _
               & ", a.sBranchNm" _
               & ", b.sCompnyNm" _
            & " From Branch a" _
               & ", Company b" _
            & " Where a.sCompnyID = b.sCompnyID" _
               & " And a.cWareHous = '0'"

   If Not IsMissing(sValue) Then
      If Not IsMissing(bSearch) Then
         If bSearch Then
            lsSQL = lsSQL & " And a.sBranchNm = " & strParm(sValue)
         Else
            lsSQL = lsSQL & " And a.sBranchNm Like " & strParm(sValue & "%")
         End If
      End If
   End If

   lsSQL = lsSQL & " Order By a.sBranchNm,b.sCompnyNm"
   lrs.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText

   If lrs.EOF Then
      oRSDetail("sBranchCd") = Null
      GoTo endProc
   ElseIf lrs.RecordCount = 1 Then
      oRSDetail("sBranchCd") = lrs("sBranchCd")
      oRSDetail("sCollIDxx") = Null
      SearhBranch = lrs("sBranchNm")
   Else
      lsBrowse = KwikBrowse(oApp _
                              , lrs _
                              , "sBranchCd»sBranchNm»sCompnyNm" _
                              , "Code»Branch»Compny" _
                              , "@»@»@" _
                              , "a.sBrancCd»a.sBranchNm»b.sCompnyNm")

      If lsBrowse = "" Then GoTo endProc
      lsSelected = Split(lsBrowse, "»")
      oRSDetail("sBranchCd") = lsSelected(0)
      SearhBranch = lsSelected(1)
      oRSDetail("sCollIDxx") = Null
   End If

endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " _
                       & "  " & IFNull(bSearch) _
                       & ", " & IFNull(sValue) _
                       & " ) "
End Function

Private Function SearchCollector(Optional bSearch As Variant, Optional sValue As Variant) As String
   Dim lrs As ADODB.Recordset
   Dim lsSelected() As String
   Dim lsOldProc As String
   Dim lsBrowse As String
   Dim lsSQL As String

   lsOldProc = "SearchCollector"
   'On Error GoTo errProc
   SearchCollector = ""

   Set lrs = New ADODB.Recordset

   lsSQL = "Select" _
               & "  a.sEmployID" _
               & ", CONCAT(a.sLastName, ', ', a.sFrstName, ' ', a.sMiddName) xCollectr" _
               & ", b.sBranchNm" _
            & " From Employee_Master a" _
               & " Left Join Branch b" _
                  & " On a.sBranchCd = b.sBranchCd" _
            & " Where a.cCollectr = '1'"

   If Not IsMissing(sValue) Then
      If Not IsMissing(bSearch) Then
         If bSearch Then
            lsSQL = lsSQL & " And CONCAT(a.sLastName, ', ', a.sFrstName, ' ', a.sMiddName) = " & strParm(sValue)
         Else
            lsSQL = lsSQL & " And CONCAT(a.sLastName, ', ', a.sFrstName, ' ', a.sMiddName) Like " & strParm(sValue & "%")
         End If
      End If
   End If


   lsSQL = lsSQL _
            & " UNION Select" _
               & "  c.sEmployID" _
               & ", CONCAT(a.sLastName, ', ', a.sFrstName, ' ', a.sMiddName) xCollectr" _
               & ", b.sBranchNm" _
            & " From Employee_Master001 c" _
               & " Left Join Branch b" _
                  & " On c.sBranchCd = b.sBranchCd" _
               & ", Client_Master a" _
            & " Where c.cCollectr = '1'" _
               & " AND c.sEmployID = a.sClientID"

   If Not IsMissing(sValue) Then
      If Not IsMissing(bSearch) Then
         If bSearch Then
            lsSQL = lsSQL & " And CONCAT(a.sLastName, ', ', a.sFrstName, ' ', a.sMiddName) = " & strParm(sValue)
         Else
            lsSQL = lsSQL & " And CONCAT(a.sLastName, ', ', a.sFrstName, ' ', a.sMiddName) Like " & strParm(sValue & "%")
         End If
      End If
   End If

   lsSQL = lsSQL & " Order By xCollectr,sBranchNm"
   Debug.Print lsSQL
   lrs.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText

   If lrs.EOF Then
      oRSDetail("sCollIDxx") = Null
      GoTo endProc
   ElseIf lrs.RecordCount = 1 Then
      SearchCollector = lrs("xCollectr")
      oRSDetail("sCollIDxx") = lrs("sEmployID")
   Else
      lsBrowse = KwikBrowse(oApp _
                              , lrs _
                              , "sEmployID»xCollectr»sBranchNm" _
                              , "Code»Collector»Branch" _
                              , "@»@»@" _
                              , "a.sBrancCd»CONCAT(a.sLastName, ', ', a.sFrstName, ' ', a.sMiddName)»b.sBranchNm")

      If lsBrowse = "" Then GoTo endProc
      lsSelected = Split(lsBrowse, "»")
      oRSDetail("sCollIDxx") = lsSelected(0)
      SearchCollector = lsSelected(1)
   End If

   With GridEditor1
      .Tag = .TextMatrix(.Row, .Col)
   End With

endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " _
                       & "  " & IFNull(bSearch) _
                       & ", " & IFNull(sValue) _
                       & " ) "
End Function

Private Sub computeAmount()
   Dim lsOldProc As String
   Dim lnCtr As Integer

   lsOldProc = "computeAmount"
   'On Error GoTo errProc

   If pnItemCount = 0 Then pnItemCount = 1
   With GridEditor1
      For lnCtr = pnItemCount To .Rows - 1
         pnRebTotlx = .TextMatrix(lnCtr, 6) + oRSMaster("nRebTotlx")
         pnPenTotlx = .TextMatrix(lnCtr, 7) + oRSMaster("nPenTotlx")
         Select Case LCase(.TextMatrix(lnCtr, 3))
         Case "mp"
            pnPaymTotl = CDbl(.TextMatrix(lnCtr, 5)) + CDbl(.TextMatrix(lnCtr, 6)) + oRSMaster("nPaymTotl")
         Case "dm"
            pnDebtTotl = CDbl(.TextMatrix(lnCtr, 5)) + oRSMaster("nDebtTotl")
         Case "cm"
            pnCredTotl = CDbl(.TextMatrix(lnCtr, 5)) + oRSMaster("nCredTotl")
         Case "dp"
            pnDownTotl = CDbl(.TextMatrix(lnCtr, 5)) + oRSMaster("nDownTotl")
         Case "cb"
            pnCashTotl = CDbl(.TextMatrix(lnCtr, 5)) + oRSMaster("nCashTotl")
         End Select
      Next
   End With

   txtField(3).Text = Format(pnPaymTotl - pnRebTotlx, "#,##0.00")
   txtField(4).Text = Format(pnRebTotlx, "#,##0.00")
   txtField(5).Text = Format(pnPaymTotl, "#,##0.00")

endProc:

   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )"
End Sub

Private Function SaveTransaction() As Boolean
   Dim lsOldProc As String
   Dim lsSQL As String
   Dim lnCtr As Integer
   Dim lnRow As Integer
   Dim lnTerm As Integer
   Dim lnAmtDuex As Double, lnDelayxx As Double

   lsOldProc = "SaveTransaction"
   oApp.BeginTrans

   'On Error GoTo errProc
   SaveTransaction = False

   With GridEditor1
      oRSDetail.Move pnItemCount, adBookmarkFirst
      For lnCtr = pnItemCount To oRSDetail.RecordCount - 1
         lsSQL = "INSERT INTO MC_AR_Ledger" _
                     & " (sAcctNmbr" _
                     & ", sBranchCd" _
                     & ", nEntryNox" _
                     & ", dTransact" _
                     & ", cOffPaymx" _
                     & ", sCollIDxx" _
                     & ", sORNoxxxx" _
                     & ", cTrantype" _
                     & ", sRemarksx" _
                     & ", nTranAmtx" _
                     & ", nDebitAmt" _
                     & ", nOthersxx" _
                     & ", nRebatesx" _
                     & ", nAmtDuexx" _
                     & ", nABalance" _
                     & ", nMonDelay" _
                     & ", dModified" _
                     & ", cPostedxx)"

         With oRSDetail
            lsSQL = lsSQL & "VALUES" _
                     & "(" & strParm(.Fields("sAcctNmbr")) _
                     & "," & strParm(.Fields("sBranchCd")) _
                     & "," & CDbl(.Fields("nEntryNox")) _
                     & "," & dateParm(.Fields("dTransact")) _
                     & "," & strParm(IIf(IsNull(.Fields("sCollIDxx")), 1, 0)) _
                     & "," & strParm(IIf(IsNull(.Fields("sCollIDxx")), "", .Fields("sCollIDxx"))) _
                     & "," & strParm(.Fields("sORNoxxxx")) _
                     & "," & strParm(.Fields("cTrantype")) _
                     & "," & strParm(.Fields("sRemarksx")) _
                     & "," & CDbl(.Fields("nTranAmtx")) _
                     & "," & CDbl(.Fields("nDebitAmt")) _
                     & "," & CDbl(.Fields("nOthersxx")) _
                     & "," & CDbl(.Fields("nRebatesx")) _
                     & "," & CDbl(.Fields("nAmtDuexx")) _
                     & "," & CDbl(.Fields("nABalance")) _
                     & "," & CDbl(.Fields("nMonDelay")) _
                     & "," & dateParm(oApp.ServerDate) _
                     & "," & strParm(xeStateClosed) & ")"
         End With

         If oApp.Execute(lsSQL, "MC_AR_Ledger") = 0 Then
            MsgBox "Unable to Save MC_AR_Ledger!!!", vbCritical, "Warning"
            GoTo endWithRoll
         End If
         oRSDetail.MoveNext
      Next

      With oRSMaster
         'now compute the delay for the master table
         lnTerm = getMonthTerm(.Fields("dFirstPay"), Date)
         lnAmtDuex = lnTerm * .Fields("nMonAmort") + .Fields("nDownPaym") + .Fields("nCashBalx")
         lnAmtDuex = lnAmtDuex - pnPaymTotl - pnDownTotl - _
                     pnCashTotl + pnDebtTotl - pnCredTotl
         lnAmtDuex = IIf(lnAmtDuex < 0, 0, lnAmtDuex)

         If .Fields("nMonAmort") > 0# Then
            lnDelayxx = Round(lnAmtDuex / .Fields("nMonAmort"), 2)
         Else
            If .Fields("dDueDatex") < Date Then lnDelayxx = 1
         End If
      End With

      lsSQL = "UPDATE MC_AR_Master Set" _
                  & "  nLastPaym = " & CDbl(.TextMatrix(.Rows - 1, 5)) _
                  & ", dLastPaym = " & dateParm(.TextMatrix(.Rows - 1, 1)) _
                  & ", nPaymTotl = " & pnPaymTotl _
                  & ", nPenTotlx = " & pnPenTotlx _
                  & ", nRebTotlx = " & pnRebTotlx _
                  & ", nDebtTotl = " & pnDebtTotl _
                  & ", nCredTotl = " & pnCredTotl _
                  & ", nAmtDuexx = " & lnAmtDuex _
                  & ", nABalance = " & CDbl(.TextMatrix(.Rows - 1, 8)) _
                  & ", nDownTotl = " & pnDownTotl _
                  & ", nCashTotl = " & pnCashTotl _
                  & ", nDelayAvg = " & lnDelayxx _
                  & ", nLedgerNo = " & CDbl(.TextMatrix(.Rows - 1, 0))

      If CDbl(.TextMatrix(.Rows - 1, 8)) <= 0# Then
         lsSQL = lsSQL _
                  & ", cAcctstat = '1'" _
                  & ", dClosedxx = " & dateParm(.TextMatrix(.Rows - 1, 1)) _
                  & ", cRatingxx = " & strParm(getRating(lnDelayxx, psSelected(27), psSelected(6)))
      End If
      lsSQL = lsSQL _
               & " Where sAcctNmbr = " & strParm(oRSMaster("sAcctNmbr"))

      If oApp.Execute(lsSQL, "MC_AR_Master") = 0 Then
         MsgBox "Unable to Save MC_AR_Master!!!", vbCritical, "Warning"
         GoTo endWithRoll
      End If

      If CDbl(.TextMatrix(.Rows - 1, 8)) = 0# Then lblFields(9).Caption = AcctStat(xeActStatClosed)
   End With
   oApp.CommitTrans

   SaveTransaction = True

endProc:
   Exit Function
endWithRoll:
   oApp.RollbackTrans
   GoTo endProc
errProc:
   oApp.RollbackTrans
   ShowError lsOldProc & "( " & " )"
End Function

Private Sub InitValue()
   With oRSDetail
      .AddNew
      .Fields("sAcctNmbr") = oRSMaster("sAcctNmbr")
      .Fields("dTransact") = oApp.ServerDate
      .Fields("cTrantype") = ""
      .Fields("sORNoxxxx") = ""
      .Fields("nTranAmtx") = 0#
      .Fields("nRebatesx") = 0#
      .Fields("nOthersxx") = 0#
      .Fields("nABalance") = oRSDetail("nABalance")
      .Fields("nMonDelay") = 0#
      .Fields("sRemarksx") = ""
      .Fields("sBranchCd") = Null
      .Fields("nEntryNox") = pnItemCount + 1
      .Fields("nAmtDuexx") = 0#
   End With
End Sub

Private Function DeleteDatail(ByVal nRow As Integer) As Boolean
   Dim lsSQL As String
   Dim lsOldProc As String

   lsOldProc = "DeleteDatail"
   'On Error GoTo errProc
   DeleteDatail = False

   With oRSDetail
      .Move nRow, adBookmarkFirst

      .Delete adAffectCurrent
      'If Not .EOF Then .MoveNext
      txtField(6).Text = .RecordCount
   End With
   DeleteDatail = True

endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & nRow & " )"
End Function

Private Function isEntryOk() As Boolean
   With GridEditor1
      If Trim(.TextMatrix(pnItemCount + 1, 2)) = "" Then
         MsgBox "No Collector found!!!" & vbCrLf & _
                "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"

         .Col = 2
         .SetFocus
         GoTo EntryNotOK
      End If

      If Trim(.TextMatrix(pnItemCount + 1, 3)) = "" Then
         MsgBox "Invalis Transaction Type!!!" & vbCrLf & _
                "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"

         .Col = 3
         .SetFocus
         GoTo EntryNotOK
      End If

      If CDbl(.TextMatrix(pnItemCount + 1, 5)) = 0# Then
         MsgBox "Invalid Amount Detected!!!" & vbCrLf & _
                "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"

         .Col = 5
         .SetFocus
         GoTo EntryNotOK
      End If
   End With

EntryOK:
   isEntryOk = True
   Exit Function
EntryNotOK:
   isEntryOk = False
End Function

Private Function getDelay() As Double
   Dim lnTerm As Integer
   Dim lnAmtDuex As Double

   lnTerm = getMonthTerm(oRSMaster("dFirstPay"), oRSDetail("dTransact"))
   lnAmtDuex = lnTerm * oRSMaster("nMonAmort") + oRSMaster("nDownPaym") + oRSMaster("nCashBalx")
   lnAmtDuex = lnAmtDuex - oRSMaster("nPaymTotl") - oRSMaster("nDownTotl") - _
               oRSMaster("nCashTotl") + oRSMaster("nDebtTotl") - oRSMaster("nCredTotl")
   lnAmtDuex = IIf(lnAmtDuex < 0, 0, lnAmtDuex)

   If oRSMaster("nMonAmort") > 0# Then
      getDelay = Round(lnAmtDuex / oRSMaster("nMonAmort"), 2)
   Else
      If oRSMaster("dDueDatex") < oRSDetail("dTransact") Then getDelay = 1
   End If
End Function

Private Function reCalc() As Boolean
   Dim lsOldProc As String
   Dim lsSQL As String

   lsOldProc = "reCalc"
   'On Error GoTo errProc

   oApp.BeginTrans

   If Recalculate(oRSMaster, oApp) = False Then GoTo endProc
   Debug.Print oRSMaster("nLedgerNo")
   lsSQL = ADO2SQL(oRSMaster, "MC_AR_Master", _
                     "sAcctNmbr = " & strParm(oRSMaster("sAcctNmbr")), _
                     Encrypt(oApp.UserID), _
                     oApp.ServerDate, _
                     "xFullName»xModelNme»xCollectr»xAddressx»sBranchCd")
   Debug.Print lsSQL
   If lsSQL <> "" Then
      If oApp.Execute(lsSQL, "MC_AR_Master") = 0 Then
         MsgBox "Unable to Save Loan Receivable Master!!!", vbCritical, "Warning"
         GoTo endWithRoll
      End If
   End If

   oApp.CommitTrans
   reCalc = True

endProc:
   Exit Function
endWithRoll:
   oApp.RollbackTrans
   GoTo endProc
errProc:
   oApp.RollbackTrans
   ShowError lsOldProc & "( " & " )"
End Function

Private Function getAveDelay(ldClosedxx As Date) As Double
   Dim lsOldProc As String
   Dim lnDelayxxx As Double, lnTotDelay As Double
   Dim lnTranAmtx As Double
   Dim lanDayMon(11) As Integer, lnDayOfMon As Integer
   Dim lnCtr As Integer
   Dim ldTranDate As Date

   lsOldProc = "getAveDelay"
   'On Error GoTo errProc

   With oRSMaster
      ' Pastdue account's delay is based on the past due month
      If CLng(Format(ldClosedxx, "YYYYMM")) > _
                        CLng(Format(.Fields("dDueDatex"), "YYYYMM")) Then
         getAveDelay = DateDiff("m", .Fields("dDueDatex"), ldClosedxx)
         GoTo endProc
      End If

      If .Fields("nAcctTerm") = 0 Then GoTo endProc

      lanDayMon(0) = 31: lanDayMon(1) = 28: lanDayMon(2) = 31: lanDayMon(3) = 30
      lanDayMon(4) = 31: lanDayMon(5) = 30: lanDayMon(6) = 31: lanDayMon(7) = 31
      lanDayMon(8) = 30: lanDayMon(9) = 31: lanDayMon(10) = 30: lanDayMon(11) = 31

      lnTotDelay = 0#
      lnTranAmtx = 0#
      ldTranDate = .Fields("dFirstPay")
      lnDayOfMon = Day(.Fields("dFirstPay"))

      oRSDetail.MoveFirst
      For lnCtr = 1 To .Fields("nAcctTerm")
         If oRSDetail.EOF = False Then
            Do While DateDiff("d", oRSDetail("dTransact"), ldTranDate) >= 1
               lnTranAmtx = .Fields("nGrossPrc") - oRSDetail("nABalance")
               oRSDetail.MoveNext

            If oRSDetail.EOF Then Exit Do
            Loop
         End If
         lnDelayxxx = (lnCtr * .Fields("nMonAmort") + .Fields("nDownPaym") - lnTranAmtx) / _
                        .Fields("nMonAmort")
         lnTotDelay = lnTotDelay + lnDelayxxx

         ldTranDate = DateAdd("m", 1, ldTranDate)

         If lnDayOfMon > Day(ldTranDate) Then
            If lanDayMon(Month(ldTranDate) - 1) > Day(ldTranDate) Then
               ldTranDate = CDate(Month(ldTranDate) & "/" & _
                              lanDayMon(Month(ldTranDate) - 1) & "/" & _
                              Year(ldTranDate))
            End If
         End If
      Next

      If lnCtr > 1 Then lnCtr = lnCtr - 1

      lnTotDelay = Round(lnTotDelay / lnCtr, 2)
      If lnTotDelay < 0# Then
         getAveDelay = 0#
      Else
         getAveDelay = lnTotDelay * (-1)
      End If
   End With

endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & ldClosedxx & " )"
End Function

