VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0"; "xrGridControl.ocx"
Begin VB.Form frmMPCustomerLedgerClosed 
   BorderStyle     =   0  'None
   Caption         =   "Customer Ledger (Closed)"
   ClientHeight    =   7710
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
   ScaleHeight     =   7710
   ScaleWidth      =   11940
   ShowInTaskbar   =   0   'False
   Begin xrGridEditor.GridEditor GridEditor1 
      Height          =   3510
      Left            =   120
      TabIndex        =   50
      Tag             =   "et0;eb0;et0;bc2"
      Top             =   3645
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   6191
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
      Object.HEIGHT          =   3510
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
      MOUSEICON       =   "frmMPCustomerLedgerClosed.frx":0000
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
      Height          =   2505
      Index           =   0
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   1095
      Width           =   5490
      _ExtentX        =   9684
      _ExtentY        =   4419
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
         Text            =   "frmMPCustomerLedgerClosed.frx":001C
         Top             =   585
         Width           =   4185
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Co-Brwr #2:"
         Height          =   285
         Index           =   31
         Left            =   150
         TabIndex        =   69
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
         Index           =   23
         Left            =   1185
         TabIndex        =   68
         Tag             =   "tc0"
         Top             =   1230
         Width           =   4185
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Co-Brwr #1:"
         Height          =   285
         Index           =   30
         Left            =   150
         TabIndex        =   67
         Top             =   990
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
         Index           =   22
         Left            =   1185
         TabIndex        =   66
         Tag             =   "tc0"
         Top             =   990
         Width           =   4185
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Sell. Branch:"
         Height          =   285
         Index           =   29
         Left            =   45
         TabIndex        =   14
         Top             =   1950
         Width           =   1065
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
         Index           =   21
         Left            =   1215
         TabIndex        =   15
         Tag             =   "tc0"
         Top             =   1950
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
         TabIndex        =   11
         Tag             =   "tc0"
         Top             =   1470
         Width           =   4185
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Col. Branch:"
         Height          =   285
         Index           =   23
         Left            =   150
         TabIndex        =   10
         Top             =   1470
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
         TabIndex        =   13
         Tag             =   "tc0"
         Top             =   1710
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
         TabIndex        =   17
         Tag             =   "tc0"
         Top             =   2190
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
         Top             =   345
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
         Top             =   105
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
         Top             =   105
         Width           =   960
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Model:"
         Height          =   285
         Index           =   4
         Left            =   150
         TabIndex        =   16
         Top             =   2190
         Width           =   960
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Collector:"
         Height          =   285
         Index           =   7
         Left            =   150
         TabIndex        =   12
         Top             =   1710
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
         Top             =   585
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
         Top             =   345
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
            TabIndex        =   65
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
      Index           =   3
      Left            =   10590
      TabIndex        =   64
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
      Picture         =   "frmMPCustomerLedgerClosed.frx":002B
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   10590
      TabIndex        =   59
      Top             =   540
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
      Picture         =   "frmMPCustomerLedgerClosed.frx":07A5
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   2505
      Index           =   2
      Left            =   5580
      Tag             =   "wt0;fb0"
      Top             =   1095
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   4419
      Enabled         =   0   'False
      ClipControls    =   0   'False
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Status:"
         Enabled         =   0   'False
         Height          =   285
         Index           =   28
         Left            =   60
         TabIndex        =   32
         Top             =   1785
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
         Index           =   20
         Left            =   1020
         TabIndex        =   33
         Tag             =   "tc0"
         Top             =   1785
         Width           =   1290
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Mo. Instal:"
         Height          =   285
         Index           =   26
         Left            =   60
         TabIndex        =   26
         Top             =   1065
         Width           =   945
      End
      Begin VB.Label lblFields 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Apr 20, 2005"
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
         Left            =   1020
         TabIndex        =   27
         Tag             =   "tc0"
         Top             =   1305
         Width           =   1290
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Maturity:"
         Height          =   285
         Index           =   17
         Left            =   60
         TabIndex        =   24
         Top             =   825
         Width           =   945
      End
      Begin VB.Label lblFields 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Good"
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
         TabIndex        =   31
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
         TabIndex        =   29
         Tag             =   "tc0"
         Top             =   1065
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
         TabIndex        =   25
         Tag             =   "tc0"
         Top             =   825
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
         TabIndex        =   23
         Tag             =   "tc0"
         Top             =   585
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
         TabIndex        =   21
         Tag             =   "tc0"
         Top             =   345
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
         TabIndex        =   19
         Tag             =   "tc0"
         Top             =   105
         Width           =   1290
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "1st Pay Date:"
         Height          =   285
         Index           =   11
         Left            =   60
         TabIndex        =   20
         Top             =   345
         Width           =   960
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Term:"
         Height          =   285
         Index           =   10
         Left            =   60
         TabIndex        =   22
         Top             =   585
         Width           =   945
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Rating:"
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   60
         TabIndex        =   30
         Top             =   1545
         Width           =   945
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Date Closed:"
         Height          =   285
         Index           =   5
         Left            =   60
         TabIndex        =   28
         Top             =   1305
         Width           =   945
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Acct Date:"
         Height          =   285
         Index           =   2
         Left            =   60
         TabIndex        =   18
         Top             =   105
         Width           =   945
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   2505
      Index           =   3
      Left            =   8010
      Tag             =   "wt0;fb0"
      Top             =   1095
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   4419
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
         Index           =   19
         Left            =   960
         TabIndex        =   49
         Tag             =   "tc0"
         Top             =   1785
         Width           =   1245
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cash Total:"
         Height          =   285
         Index           =   27
         Left            =   45
         TabIndex        =   48
         Top             =   1785
         Width           =   945
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
         Index           =   17
         Left            =   960
         TabIndex        =   47
         Tag             =   "tc0"
         Top             =   1545
         Width           =   1245
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Delay Ave:"
         Height          =   285
         Index           =   25
         Left            =   45
         TabIndex        =   46
         Top             =   1545
         Width           =   945
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
         Index           =   15
         Left            =   960
         TabIndex        =   45
         Tag             =   "tc0"
         Top             =   1305
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
         TabIndex        =   37
         Tag             =   "tc0"
         Top             =   345
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
         TabIndex        =   39
         Tag             =   "tc0"
         Top             =   585
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
         TabIndex        =   41
         Tag             =   "tc0"
         Top             =   825
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
         TabIndex        =   43
         Tag             =   "tc0"
         Top             =   1065
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
         TabIndex        =   35
         Tag             =   "tc0"
         Top             =   105
         Width           =   1245
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Reb. Guide:"
         Height          =   285
         Index           =   18
         Left            =   45
         TabIndex        =   44
         Top             =   1305
         Width           =   945
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Down Paym:"
         Height          =   285
         Index           =   16
         Left            =   45
         TabIndex        =   36
         Top             =   345
         Width           =   945
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cash Bal:"
         Height          =   285
         Index           =   15
         Left            =   45
         TabIndex        =   38
         Top             =   585
         Width           =   945
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pen. Guide:"
         Height          =   285
         Index           =   14
         Left            =   45
         TabIndex        =   42
         Top             =   1065
         Width           =   945
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "P/N Value:"
         Height          =   285
         Index           =   13
         Left            =   45
         TabIndex        =   40
         Top             =   825
         Width           =   945
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Gross Price:"
         Height          =   285
         Index           =   12
         Left            =   45
         TabIndex        =   34
         Top             =   105
         Width           =   945
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   10590
      TabIndex        =   62
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
      Picture         =   "frmMPCustomerLedgerClosed.frx":0F1F
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   10590
      TabIndex        =   61
      Top             =   1800
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "Ra&ting"
      AccessKey       =   "t"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmMPCustomerLedgerClosed.frx":1699
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   465
      Index           =   4
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   7185
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
         TabIndex        =   52
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
         TabIndex        =   51
         Top             =   120
         Width           =   840
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   465
      Index           =   5
      Left            =   3615
      Tag             =   "wt0;fb0"
      Top             =   7185
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   820
      BackColor       =   12632256
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
         Index           =   3
         Left            =   810
         TabIndex        =   54
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
         TabIndex        =   53
         Top             =   120
         Width           =   675
      End
   End
   Begin xrControl.xrFrame xrFrame3 
      Height          =   465
      Index           =   1
      Left            =   5850
      Tag             =   "wt0;fb0"
      Top             =   7185
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
         TabIndex        =   56
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
         TabIndex        =   55
         Top             =   120
         Width           =   675
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   465
      Index           =   6
      Left            =   8085
      Tag             =   "wt0;fb0"
      Top             =   7185
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   820
      BackColor       =   12632256
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
         TabIndex        =   58
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
         TabIndex        =   57
         Top             =   105
         Width           =   570
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   10590
      TabIndex        =   60
      Top             =   1170
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
      Picture         =   "frmMPCustomerLedgerClosed.frx":1E13
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   10590
      TabIndex        =   63
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
      Picture         =   "frmMPCustomerLedgerClosed.frx":2541
   End
End
Attribute VB_Name = "frmMPCustomerLedgerClosed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmCustomerLedgerClosed"
'Ok
Private oSkin As clsFormSkin
Private oRSMaster As ADODB.Recordset
Private oRSDetail As ADODB.Recordset
Private oFormImpounded As frmImpounded
Private oFormRating As frmRatingStat

Dim psSelected() As String
Dim pnCtr As Integer

Private Sub cmdButton_Click(Index As Integer)
   Dim lnCtr As Integer
   Dim lbCancel As Boolean
   Dim lrs As ADODB.Recordset
   Dim lsOldProc As String
   Dim lnRep As Integer

   lsOldProc = "cmdButton_Click"
   'On Error GoTo errProc

   With GridEditor1
      Select Case Index
      Case 0 'Browse
         If SearchTransaction(, , True) Then
            LoadMaster
            LoadDetail
         End If
         .Refresh
      Case 1 'Rating
         If Trim(txtField(0).Text) = "" Then
            MsgBox "No Account is Loaded to modify!!!" & vbCrLf & _
                   "Please Verify your Entry then Try Again!!!", vbInformation, "Notice"
            GoTo endProc
         End If

         With oFormRating
            .RatingStat = psSelected(9)
            .Show 1

            If Not .Cancelled Then saveRating
         End With
         .Refresh
      Case 2 'Impound
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
      Case 3 'Close
         Unload Me
      Case 4 'ReCalculate
         If Trim(txtField(0).Text) = "" Then
            MsgBox "No Account is Loaded to modify!!!" & vbCrLf & _
                   "Please Verify your Entry then Try Again!!!", vbInformation, "Notice"
            GoTo endProc
         End If

         If reCalc() Then
            MsgBox "Transaction Updated Successfully!!!", vbInformation, "Notice"
            If SearchTransaction(oRSMaster("sAcctNmbr"), True, False) Then
               LoadMaster
               LoadDetail
            End If
         End If
      Case 5 'Print
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
   Set oFormRating = New frmRatingStat

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualRight

   InitGrid
   InitEntry

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
   Set oRSMaster = Nothing
   Set oRSDetail = Nothing
   Set oFormRating = Nothing
   Set oFormImpounded = Nothing
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
      Case 10 To 15, 17, 19
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
         .ColEnabled(pnCtr) = False
      Next

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

Private Sub GridEditor1_LostFocus()
   With GridEditor1
      .LeftCol = 1
      .Col = 1
   End With
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
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
               & ", IFNULL(CONCAT(i.sLastName, ', ', i.sFrstName, ' ', i.sMiddName), CONCAT(p.sLastName, ', ', p.sFrstName, ' ', p.sMiddName)) xCollectr" _
               & ", a.dPurchase" _
               & ", a.dFirstPay" _
               & ", a.nAcctTerm" _
               & ", a.dDueDatex" _
               & ", a.nMonAmort" _
               & ", a.cRatingxx" _
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

   lsSQL = lsSQL _
               & ", a.nDelayAvg" _
               & ", a.dClosedxx" _
               & ", a.nCashTotl" _
               & ", a.cAcctStat" _
               & ", k.sBranchNm xSelBrnch" _
               & ", a.nDebtTotl" _
               & ", a.nCredTotl" _
               & ", a.nDownTotl" _
               & ", h.sBranchCd" _
               & ", a.nABalance" _
               & ", a.nLastPaym" _
               & ", a.dLastPaym" _
               & ", a.nAmtDuexx" _
               & ", a.nDelayAvg" _
               & ", a.nLedgerNo" _
               & ", a.sModified" _
               & ", a.dModified" _
               & ", CONCAT(m.sLastName, ', ', m.sFrstName, ' ', m.sMiddName) xCoCltNm1" _
               & ", CONCAT(n.sLastName, ', ', n.sFrstName, ' ', n.sMiddName) xCoCltNm2" _
               & ", IFNULL(CONCAT(i.sLastName, ', ', i.sFrstName, ' ', i.sMiddName), CONCAT(p.sLastName, ', ', p.sFrstName, ' ', p.sMiddName)) zCollectr"

   lsSQL = lsSQL _
            & " From MC_AR_Master a" _
               & " LEFT JOIN MC_Serial e" _
                  & " On a.sSerialID = e.sSerialID" _
               & " Left Join MC_Model f" _
                  & " On e.sModelIDx = f.sModelIDx" _
               & " Left Join Brand g" _
                  & " On f.sBrandIDx = g.sBrandIDx" _
               & " LEFT JOIN Branch k" _
                  & " On Left(a.sAcctNmbr, " & Len(oApp.BranchCode) & ") = k.sBranchCd" _
               & " Left Join Client_Master m" _
                  & " On a.sCoCltID1 = m.sClientID" _
               & " Left Join Client_Master n" _
                  & " On a.sCoCltID2 = n.sClientID" _
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
               & " And a.cAcctstat <> '0'" _
               & " AND a.cLoanType = '4'"
               
   If Not IsMissing(sValue) Then
      If Not IsMissing(bByCode) Then
         If bByCode Then
            lsSQL = lsSQL & " And a.sAcctNmbr = " & strParm(sValue)
         Else
            lsSQL = lsSQL & " And CONCAT(b.sLastName, ', ', b.sFrstName, ' ', b.sMiddName) Like " & strParm(Trim(sValue) & "%")
         End If
      Else
         lsSQL = lsSQL & " And CONCAT(b.sLastName, ', ', b.sFrstName, ' ',  b.sMiddName) = " & strParm(Trim(sValue))
      End If
   End If
   lsSQL = lsSQL & " Order By a.sAcctNmbr, xFullName"
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

Private Sub LoadMaster()
   For pnCtr = 0 To 25
      Select Case pnCtr
      Case 0, 1
         lblFields(pnCtr).Caption = IIf(psSelected(pnCtr) = "", "", psSelected(pnCtr))
         txtField(pnCtr).Text = lblFields(pnCtr).Caption
         txtField(pnCtr).Tag = txtField(pnCtr).Text
      Case 4, 5, 7
         lblFields(pnCtr).Caption = IIf(psSelected(pnCtr) = "", "", Format(psSelected(pnCtr), "MMM DD, YYYY"))
      Case 6
         lblFields(pnCtr).Caption = IIf(psSelected(pnCtr) = "", "", psSelected(pnCtr) & " months")
      Case 8, 10 To 15
         lblFields(pnCtr).Caption = IIf(psSelected(pnCtr) = "", "0.00", Format(psSelected(pnCtr), "#,##0.00"))
      Case 9
         lblFields(pnCtr).Caption = RatingStat(psSelected(pnCtr))
      Case 17, 18, 19
         txtField(pnCtr - 14).Text = IIf(psSelected(pnCtr) = "", "0.00", Format(psSelected(pnCtr), "#,##0.00"))
      Case 20
         txtField(2).Text = IIf(psSelected(pnCtr) = "", "", psSelected(pnCtr))
      Case 21
         lblFields(17).Caption = Format(psSelected(pnCtr), "#,##0.00")
      Case 22
         lblFields(18).Caption = Format(psSelected(pnCtr), "MMM DD, YYYY")
      Case 23
         lblFields(19).Caption = Format(psSelected(pnCtr), "#,##0.00")
      Case 24
         lblFields(20).Caption = AcctStat(psSelected(pnCtr))
      Case 25
         lblFields(21).Caption = IIf(psSelected(pnCtr) = "", Left(psSelected(0), 4), psSelected(pnCtr))
      Case Else
         lblFields(pnCtr).Caption = IIf(psSelected(pnCtr) = "", "", psSelected(pnCtr))
      End Select
   Next
   txtField(5) = Format(CDbl(txtField(3)) + CDbl(txtField(4)), "#,##0.00")

   lblFields(22).Caption = IIf(psSelected(38) = "", "N-O-N-E", psSelected(38))
   lblFields(23).Caption = IIf(psSelected(39) = "", "N-O-N-E", psSelected(39))

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
               & ", a.sBranchCd"
   lsSQL = lsSQL _
            & " From MC_AR_Ledger a" _
               & " Left Join Employee_Master b" _
                  & " On a.sCollIDxx = b.sEmployID" _
               & " LEFT JOIN Employee_Master001 d" _
                  & " ON a.sCollIDxx = d.sEmployID" _
                  & " LEFT JOIN Client_Master e" _
                     & " ON d.sEmployID = e.sClientID" _
               & " , Branch c" _
            & " Where a.sAcctNmbr = " & strParm(psSelected(0)) _
               & " And a.sBranchCd = c.sBranchCd" _
            & " Order By a.dTransact, a.sORNoxxxx"

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

   With GridEditor1
      .Rows = oRSDetail.RecordCount + 1

      txtField(6).Text = oRSDetail.RecordCount
      For pnCtr = 0 To oRSDetail.RecordCount - 1
         .Row = pnCtr + 1
         For lnCol = 1 To .Cols - 1
            Select Case lnCol
            Case 1
               .TextMatrix(.Row, lnCol) = IIf(IsNull(oRSDetail(lnCol)), "", Format(oRSDetail(lnCol), "MM/DD/YYYY"))
            Case 2
               .TextMatrix(.Row, lnCol) = IIf(IsNull(oRSDetail(lnCol)), oRSDetail("sBranchNm"), IIf(oRSDetail(lnCol) = "", oRSDetail("sBranchNm"), oRSDetail(lnCol)))
            Case 3
               .TextMatrix(.Row, lnCol) = IIf(IsNull(oRSDetail(lnCol)), "", TranType(oRSDetail(lnCol)))
            Case 5
               If oRSDetail("cTrantype") = "m" Then
                  .TextMatrix(.Row, lnCol) = IIf(IsNull(oRSDetail("nDebitAmt")), 0#, Format(oRSDetail("nDebitAmt"), "#,##0.00"))
               Else
                  .TextMatrix(.Row, lnCol) = IIf(IsNull(oRSDetail(lnCol)), 0#, Format(oRSDetail(lnCol), "#,##0.00"))
               End If
            Case 6, 7, 8, 9
               .TextMatrix(.Row, lnCol) = IIf(IsNull(oRSDetail(lnCol)), 0#, Format(oRSDetail(lnCol), "#,##0.00"))
            Case Else
               .TextMatrix(.Row, lnCol) = IIf(IsNull(oRSDetail(lnCol)), "", oRSDetail(lnCol))
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

Private Sub saveRating()
   Dim lsOldProc As String, lsSQL As String

   lsOldProc = "saveRating"
   'On Error GoTo errProc

   If psSelected(9) = oFormRating.RatingStat Then GoTo endProc

   lsSQL = "Update MC_AR_Master Set" _
                              & "  cRatingxx = " & strParm(oFormRating.RatingStat) _
                              & ", dModified = " & dateParm(oApp.ServerDate()) _
                           & " Where sAcctNmbr = " & strParm(psSelected(0))
   If oApp.Execute(lsSQL, "MC_AR_Master") = 0 Then
      MsgBox "Unable to Update Rating!!!" & vbCrLf & lsSQL, vbCritical, "Warning"
      GoTo endProc
   End If
   lblFields(9).Caption = RatingStat(oFormRating.RatingStat)
   psSelected(9) = oFormRating.RatingStat

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )"
End Sub

Private Function reCalc() As Boolean
   Dim lsOldProc As String
   Dim lsSQL As String

   lsOldProc = "reCalc"
   'On Error GoTo errProc

   oApp.BeginTrans

   If Recalculate(oRSMaster, oApp) = False Then GoTo endProc

   lsSQL = ADO2SQL(oRSMaster, "MC_AR_Master", _
                     "sAcctNmbr = " & strParm(oRSMaster("sAcctNmbr")), _
                     Encrypt(oApp.UserID), _
                     oApp.ServerDate, _
                     "xFullName»xModelNme»xCollectr»xAddressx»sBranchCd")

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


