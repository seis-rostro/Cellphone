VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0"; "xrGridControl.ocx"
Begin VB.Form frmCP_JobOrderReg 
   BorderStyle     =   0  'None
   Caption         =   "Job Order"
   ClientHeight    =   7800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10980
   DrawWidth       =   18832
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7800
   ScaleWidth      =   10980
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   2190
      Left            =   105
      TabIndex        =   22
      Tag             =   "wt0;wb0"
      Top             =   3210
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   3863
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "UNIT INFO"
      TabPicture(0)   =   "frmCP_JobOrderReg.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "xrFrame3(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "UNIT CONDITION"
      TabPicture(1)   =   "frmCP_JobOrderReg.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "xrFrame3(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "UNIT STATUS"
      TabPicture(2)   =   "frmCP_JobOrderReg.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "xrFrame3(2)"
      Tab(2).ControlCount=   1
      Begin xrControl.xrFrame xrFrame3 
         Height          =   1860
         Index           =   0
         Left            =   15
         Tag             =   "wt0;fb0"
         Top             =   315
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   3281
         BackColor       =   12632256
         Enabled         =   0   'False
         ClipControls    =   0   'False
         BorderStyle     =   4
         Begin VB.CheckBox chkBackJob 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Back Job (J.O. #)"
            Height          =   195
            Left            =   6825
            TabIndex        =   31
            Tag             =   "et0;fb0"
            Top             =   195
            Width           =   1665
         End
         Begin VB.OptionButton chkServiceType 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Void Warranty"
            Height          =   195
            Index           =   0
            Left            =   4575
            TabIndex        =   29
            Tag             =   "et0;fb0"
            Top             =   180
            Width           =   1440
         End
         Begin VB.OptionButton chkServiceType 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Under Limited Warranty"
            Height          =   195
            Index           =   1
            Left            =   4575
            TabIndex        =   30
            Tag             =   "et0;fb0"
            Top             =   465
            Width           =   1965
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   13
            Left            =   6825
            TabIndex        =   32
            Top             =   435
            Width           =   1665
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   12
            Left            =   915
            MaxLength       =   25
            TabIndex        =   28
            Top             =   435
            Width           =   3240
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   10
            Left            =   915
            TabIndex        =   24
            Top             =   105
            Width           =   1440
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   11
            Left            =   2985
            TabIndex        =   26
            Top             =   105
            Width           =   1170
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   450
            Index           =   14
            Left            =   915
            MaxLength       =   512
            MultiLine       =   -1  'True
            TabIndex        =   34
            Top             =   825
            Width           =   8280
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   450
            Index           =   15
            Left            =   915
            MaxLength       =   512
            MultiLine       =   -1  'True
            TabIndex        =   36
            Top             =   1290
            Width           =   8280
         End
         Begin VB.Label lblField 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dealer"
            Height          =   195
            Index           =   13
            Left            =   360
            TabIndex        =   27
            Top             =   480
            Width           =   480
         End
         Begin VB.Label lblField 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DOP"
            Height          =   195
            Index           =   12
            Left            =   360
            TabIndex        =   23
            Top             =   150
            Width           =   480
         End
         Begin VB.Label lblField 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ref No."
            Height          =   195
            Index           =   14
            Left            =   2400
            TabIndex        =   25
            Top             =   195
            Width           =   555
         End
         Begin VB.Label lblField 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Accessory"
            Height          =   195
            Index           =   2
            Left            =   75
            TabIndex        =   33
            Top             =   810
            Width           =   750
         End
         Begin VB.Label lblField 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Remarks"
            Height          =   195
            Index           =   3
            Left            =   195
            TabIndex        =   35
            Top             =   1275
            Width           =   630
         End
      End
      Begin xrControl.xrFrame xrFrame3 
         Height          =   1860
         Index           =   1
         Left            =   -74985
         Tag             =   "wt0;fb0"
         Top             =   315
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   3281
         BackColor       =   12632256
         Enabled         =   0   'False
         ClipControls    =   0   'False
         BorderStyle     =   4
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   16
            Left            =   900
            TabIndex        =   42
            Top             =   1365
            Width           =   3285
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   5
            Left            =   900
            TabIndex        =   38
            Top             =   225
            Width           =   3300
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   450
            Index           =   6
            Left            =   900
            MaxLength       =   250
            MultiLine       =   -1  'True
            TabIndex        =   40
            Top             =   555
            Width           =   3300
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   21
            Left            =   5160
            TabIndex        =   50
            Top             =   1035
            Width           =   3405
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   23
            Left            =   5160
            TabIndex        =   53
            Top             =   1365
            Width           =   3405
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   17
            Left            =   5160
            TabIndex        =   44
            Top             =   210
            Width           =   3405
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   19
            Left            =   5160
            TabIndex        =   47
            Top             =   540
            Width           =   3405
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Index           =   18
            Left            =   8580
            TabIndex        =   45
            Top             =   210
            Width           =   600
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Index           =   20
            Left            =   8580
            TabIndex        =   48
            Top             =   540
            Width           =   600
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Index           =   22
            Left            =   8580
            TabIndex        =   51
            Top             =   1035
            Width           =   600
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Index           =   24
            Left            =   8580
            TabIndex        =   54
            Top             =   1365
            Width           =   600
         End
         Begin VB.Label lblField 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Technician"
            Height          =   195
            Index           =   21
            Left            =   60
            TabIndex        =   41
            Top             =   1395
            Width           =   795
         End
         Begin VB.Label lblField 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ASC Name"
            Height          =   195
            Index           =   4
            Left            =   60
            TabIndex        =   37
            Top             =   240
            Width           =   780
         End
         Begin VB.Label lblField 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            Height          =   195
            Index           =   6
            Left            =   270
            TabIndex        =   39
            Top             =   585
            Width           =   570
         End
         Begin VB.Label lblField 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Condition"
            Height          =   195
            Index           =   16
            Left            =   4470
            TabIndex        =   43
            Top             =   270
            Width           =   660
         End
         Begin VB.Label lblField 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Symptom"
            Height          =   195
            Index           =   17
            Left            =   4470
            TabIndex        =   46
            Top             =   615
            Width           =   660
         End
         Begin VB.Label lblField 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Defect"
            Height          =   195
            Index           =   19
            Left            =   4470
            TabIndex        =   49
            Top             =   1095
            Width           =   660
         End
         Begin VB.Label lblField 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Repair"
            Height          =   195
            Index           =   20
            Left            =   4665
            TabIndex        =   52
            Top             =   1425
            Width           =   465
         End
      End
      Begin xrControl.xrFrame xrFrame3 
         Height          =   1860
         Index           =   2
         Left            =   -74985
         Tag             =   "wt0;fb0"
         Top             =   315
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   3281
         BackColor       =   12632256
         Enabled         =   0   'False
         ClipControls    =   0   'False
         BorderStyle     =   4
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   29
            Left            =   7965
            TabIndex        =   62
            Top             =   645
            Width           =   1215
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   28
            Left            =   4530
            TabIndex        =   60
            Top             =   615
            Width           =   2115
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   33
            Left            =   7965
            TabIndex        =   70
            Top             =   1305
            Width           =   1215
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   32
            Left            =   4530
            TabIndex        =   68
            Top             =   1275
            Width           =   2115
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   31
            Left            =   7965
            TabIndex        =   66
            Top             =   975
            Width           =   1215
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   30
            Left            =   4530
            TabIndex        =   64
            Top             =   945
            Width           =   2115
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   27
            Left            =   1335
            TabIndex        =   58
            Top             =   555
            Width           =   1215
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   26
            Left            =   1335
            TabIndex        =   56
            Top             =   225
            Width           =   3225
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date Encoded"
            Height          =   195
            Index           =   30
            Left            =   6810
            TabIndex        =   61
            Top             =   705
            Width           =   1035
         End
         Begin VB.Label lblField 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Encoded By"
            Height          =   195
            Index           =   29
            Left            =   3510
            TabIndex        =   59
            Top             =   690
            Width           =   870
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Branch Location"
            Enabled         =   0   'False
            Height          =   195
            Index           =   28
            Left            =   105
            TabIndex        =   55
            Top             =   255
            Width           =   1170
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date Payment"
            Height          =   195
            Index           =   27
            Left            =   6810
            TabIndex        =   69
            Top             =   1365
            Width           =   1005
         End
         Begin VB.Label lblField 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Paym. Recv. By"
            Height          =   195
            Index           =   26
            Left            =   3240
            TabIndex        =   67
            Top             =   1350
            Width           =   1140
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date Released"
            Height          =   195
            Index           =   25
            Left            =   6810
            TabIndex        =   65
            Top             =   1035
            Width           =   1065
         End
         Begin VB.Label lblField 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Released By"
            Height          =   195
            Index           =   24
            Left            =   3480
            TabIndex        =   63
            Top             =   1020
            Width           =   900
         End
         Begin VB.Label lblField 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date Repaired"
            Enabled         =   0   'False
            Height          =   195
            Index           =   8
            Left            =   105
            TabIndex        =   57
            Top             =   645
            Width           =   1035
         End
      End
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   2265
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   5415
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   3995
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   25
         Left            =   4770
         TabIndex        =   73
         Text            =   "0,000.00"
         Top             =   1605
         Width           =   810
      End
      Begin xrGridEditor.GridEditor GridEditor1 
         Height          =   1500
         Left            =   75
         TabIndex        =   71
         Tag             =   "et0;eb0;et0;bc2"
         Top             =   75
         Width           =   9120
         _ExtentX        =   16087
         _ExtentY        =   2646
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
         Object.HEIGHT          =   1500
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
         MOUSEICON       =   "frmCP_JobOrderReg.frx":0054
         MOUSEPOINTER    =   0
         REDRAW          =   -1  'True
         RIGHTTOLEFT     =   0   'False
         ROWS            =   6
         SCROLLBARS      =   3
         SCROLLTRACK     =   0   'False
         SELECTIONMODE   =   0
         Object.TOOLTIPTEXT     =   ""
         WORDWRAP        =   0   'False
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OTHER CHARGES"
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
         Index           =   23
         Left            =   3120
         TabIndex        =   72
         Top             =   1650
         Width           =   1605
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL PARTS"
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
         Index           =   22
         Left            =   5610
         TabIndex        =   76
         Top             =   1920
         Width           =   1290
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL LABOR"
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
         Index           =   9
         Left            =   5610
         TabIndex        =   74
         Top             =   1650
         Width           =   1290
      End
      Begin VB.Label lblTotalParts 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,000.00"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   6930
         TabIndex        =   77
         Tag             =   "wt0"
         Top             =   1890
         Width           =   765
      End
      Begin VB.Label lblTotalLabor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,000.00"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   6930
         TabIndex        =   75
         Tag             =   "wt0"
         Top             =   1605
         Width           =   765
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,000.00"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   7710
         TabIndex        =   78
         Tag             =   "ht0;hb0"
         Top             =   1605
         Width           =   1485
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   2115
      Index           =   0
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   1080
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   3731
      BackColor       =   12632256
      Enabled         =   0   'False
      BorderStyle     =   1
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
         Height          =   315
         Index           =   0
         Left            =   990
         TabIndex        =   7
         Top             =   150
         Width           =   1620
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   3210
         MaxLength       =   10
         TabIndex        =   11
         Top             =   630
         Width           =   1080
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   990
         TabIndex        =   9
         Text            =   "DEC-01-2010"
         Top             =   630
         Width           =   1620
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   8
         Left            =   990
         TabIndex        =   15
         Top             =   1290
         Width           =   1620
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   7
         Left            =   990
         MaxLength       =   25
         TabIndex        =   13
         Top             =   960
         Width           =   3300
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   450
         Index           =   4
         Left            =   5235
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   21
         Top             =   960
         Width           =   3945
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   3
         Left            =   5235
         TabIndex        =   19
         Top             =   630
         Width           =   3945
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   9
         Left            =   990
         TabIndex        =   17
         Top             =   1620
         Width           =   1620
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trans. #"
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
         Index           =   0
         Left            =   90
         TabIndex        =   6
         Top             =   210
         Width           =   735
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trans. Date"
         Height          =   195
         Index           =   1
         Left            =   105
         TabIndex        =   8
         Top             =   675
         Width           =   840
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Brand"
         Height          =   195
         Index           =   5
         Left            =   450
         TabIndex        =   14
         Top             =   1350
         Width           =   480
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "J.O.#"
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
         Index           =   18
         Left            =   2700
         TabIndex        =   10
         Top             =   690
         Width           =   480
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   315
         Left            =   1080
         Tag             =   "et0;ht2"
         Top             =   240
         Width           =   1620
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IMEI"
         Height          =   195
         Index           =   7
         Left            =   450
         TabIndex        =   12
         Top             =   1005
         Width           =   480
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cust. Add."
         Height          =   195
         Index           =   10
         Left            =   4395
         TabIndex        =   20
         Top             =   990
         Width           =   735
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
         Height          =   195
         Index           =   11
         Left            =   4485
         TabIndex        =   18
         Top             =   645
         Width           =   645
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
         Height          =   195
         Index           =   15
         Left            =   450
         TabIndex        =   16
         Top             =   1665
         Width           =   480
      End
      Begin VB.Shape Shape3 
         Height          =   390
         Index           =   0
         Left            =   6675
         Top             =   120
         Width           =   2505
      End
      Begin VB.Shape Shape4 
         Height          =   330
         Index           =   0
         Left            =   6705
         Top             =   150
         Width           =   2445
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "FORWARDED"
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
         Height          =   240
         Left            =   6735
         TabIndex        =   82
         Tag             =   "eb0;et0"
         Top             =   195
         Width           =   2385
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H8000000F&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   285
         Index           =   2
         Left            =   6735
         Tag             =   "et0;et0"
         Top             =   180
         Width           =   2400
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   525
      Index           =   1
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   926
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   36
         Left            =   6570
         TabIndex        =   5
         Top             =   90
         Width           =   2625
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   35
         Left            =   4995
         TabIndex        =   3
         Top             =   90
         Width           =   1065
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   34
         Left            =   990
         TabIndex        =   1
         Top             =   90
         Width           =   3300
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "IMEI"
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
         Index           =   1
         Left            =   6120
         TabIndex        =   4
         Top             =   135
         Width           =   510
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "J.0. #"
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
         Index           =   8
         Left            =   4395
         TabIndex        =   2
         Top             =   135
         Width           =   510
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Branch"
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
         Left            =   135
         TabIndex        =   0
         Top             =   135
         Width           =   600
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   9630
      TabIndex        =   81
      Top             =   1800
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
      Picture         =   "frmCP_JobOrderReg.frx":0070
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   9630
      TabIndex        =   79
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
      Picture         =   "frmCP_JobOrderReg.frx":07EA
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   9630
      TabIndex        =   80
      Top             =   1170
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Ledger"
      AccessKey       =   "L"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCP_JobOrderReg.frx":0F64
   End
End
Attribute VB_Name = "frmCP_JobOrderReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmCP_JobOrderReg"

Private WithEvents oTrans As clsJobOrder
Attribute oTrans.VB_VarHelpID = -1
Private oBranch As ggcParameter.clsBranch
Private oSkin As clsFormSkin

Dim pnCtr As Integer, pnIndex As Integer
Dim pbIsSrvcCenter As Boolean
Dim psBranhCd As String

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String

   lsOldProc = "cmdButton_Click"
   ''On Error GoTo errProc

   With GridEditor1
      Select Case Index
      Case 0
         If Trim(txtField(34).Text) = "" Then Exit Sub
         If oTrans.SearchOrigin(psBranhCd) Then
            Call LoadMaster
            Call LoadDetail
         End If
      Case 1
         With frmCP_JOMovementLedger
            .txtField(0) = txtField(0)
            .txtField(1) = txtField(7)
            .txtField(2) = txtField(8)
            .txtField(3) = txtField(9)
            
            .TransNox = oTrans.Master("sTransNox")
            .Show 1
         End With
      Case 2
         Unload Me
      End Select
   End With

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyReturn, vbKeyUp, vbKeyDown
      Select Case KeyCode
      Case vbKeyReturn, vbKeyDown
         If GetFocus = GridEditor1.hwnd Then Exit Sub
         SetNextFocus
      Case vbKeyUp
         SetPreviousFocus
      End Select
   End Select
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String

   lsOldProc = "Form_Load"
   ''On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oTrans = New clsJobOrder
   Set oTrans.AppDriver = oApp

   oTrans.InitTransaction
   oTrans.NewTransaction

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransMaintenance
   
   Set oBranch = New ggcParameter.clsBranch
   Set oBranch.AppDriver = oApp
   
   oBranch.Filter = "cAutomate = " & strParm(xeYes)
   oBranch.InitRecord
   oBranch.NewRecord
                        
   InitGrid
   ClearFields
   pbIsSrvcCenter = BranchStatus(oApp.BranchCode, "cSrvcCntr = " & strParm(xeYes))
   txtField(34).Enabled = (oApp.isMainOffice Or pbIsSrvcCenter)
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oTrans = Nothing
   Set oSkin = Nothing
End Sub

Private Sub GridEditor1_GotFocus()
   With GridEditor1
      .EditorBackColor = oApp.getColor("HT1")
   End With
End Sub

Private Sub GridEditor1_LostFocus()
   With GridEditor1
      .EditorBackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub oTrans_DetailRetrieved(ByVal Index As Integer)
   With GridEditor1
      Select Case Index
      Case 6
         .TextMatrix(.Row, Index) = IIf(IsNull(oTrans.Detail(.Row - 1, Index)), "0%", IIf(oTrans.Detail(.Row - 1, Index) = "", "0%", oTrans.Detail(.Row - 1, Index) & "%"))
      Case Else
         .TextMatrix(.Row, Index) = oTrans.Detail(.Row - 1, Index)
      End Select
   End With
End Sub

Private Sub oTrans_MasterRetrieved(ByVal Index As Integer)
   With oTrans
      If Index = 23 Then
         lblTotalLabor.Caption = Format(oTrans.Master("nLaborAmt"), "#,##0.00")
         DisplayComputation
      End If
      txtField(Index).Text = IFNull(.Master(Index), "")
   End With
End Sub

Private Sub ClearFields()
   For pnCtr = 0 To txtField.Count - 1
      Select Case pnCtr
      Case 0
         txtField(pnCtr).Text = Format(oTrans.Master(pnCtr), "@@@@-@@@@@@")
      Case 1, 10
         txtField(pnCtr).Text = Format(oApp.ServerDate, "MMM-DD-YYYY")
      Case 2
         txtField(pnCtr).Text = oTrans.Master(pnCtr)
      Case 25
         txtField(pnCtr).Text = "0.00"
      Case 34
         If txtField(pnCtr).Text = "" Then txtField(pnCtr).Text = oApp.BranchName
         txtField(pnCtr).Tag = txtField(pnCtr).Text
         psBranhCd = oApp.BranchCode
      Case Else
         txtField(pnCtr).Text = ""
      End Select
   Next
   
   With GridEditor1
      .Rows = 2
      .ColWidth(2) = 3100
      
      'empty row
      .TextMatrix(1, 1) = ""
      .TextMatrix(1, 2) = ""
      .TextMatrix(1, 3) = 0
      .TextMatrix(1, 4) = 0#
      .TextMatrix(1, 5) = 0
      .TextMatrix(1, 6) = "0" & "%"
      .TextMatrix(1, 7) = 0#
   End With
   
   lblTotalLabor.Caption = "0.00"
   lblTotalParts.Caption = "0.00"
   lblTotal.Caption = "0.00"
   Label2.Caption = "OPEN"
   chkServiceType(oTrans.Master("cJOTypexx")).Value = True
   chkBackJob.Value = IFNull(oTrans.Master("cBackJobx"), Unchecked)
   SSTab1.Tab = 0
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = oApp.getColor("HT1")
   End With

   pnIndex = Index
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String
   Dim lsValue As String

   lsOldProc = "txtField_KeyDown"
   ''On Error GoTo errProc

   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
      With txtField(Index)
         Select Case Index
         Case 3, 5, 7, 12, 16, 17, 19, 21, 23
            If KeyCode = vbKeyF3 Then
               oTrans.SearchMaster Index, .Text
               If .Text <> "" Then SetNextFocus
            Else
               If .Text <> "" Then oTrans.SearchMaster Index, .Text
            End If
            
            If Index = 7 Then
               If oTrans.Master("sClientID") <> "" Then txtField(3).SetFocus
            End If
         End Select
      End With
      KeyCode = 0
   End If

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index _
                       & ", " & KeyCode _
                       & ", " & Shift & " )", True
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub LoadMaster()
   For pnCtr = 0 To txtField.Count - 1
      Select Case pnCtr
      Case 0
         txtField(pnCtr).Text = Format(oTrans.Master(pnCtr), "@@@@-@@@@@@")
      Case 1, 10
         txtField(pnCtr).Text = Format(oTrans.Master(pnCtr), "MMM-DD-YYYY")
      Case 2, 35
         txtField(pnCtr).Text = oTrans.Master("sReferNox")
      Case 7, 36
         txtField(pnCtr).Text = oTrans.Master("sSerialNo")
      Case 26
         txtField(pnCtr).Text = oTrans.Master("sBranchNm")
      Case 27
         txtField(pnCtr).Text = IFNull(Format(oTrans.Master("dRepaired"), "MMM-DD-YYYY"), "")
      Case 28
         txtField(pnCtr).Text = oApp.getUserName(Decrypt(oTrans.Master("sModified")))
      Case 29
         txtField(pnCtr).Text = IFNull(Format(oTrans.Master("dModified"), "MMM-DD-YYYY"), "")
      Case 30
         txtField(pnCtr).Text = oApp.getUserName(IFNull(oTrans.Master("sReleased"), ""))
      Case 31
         txtField(pnCtr).Text = IFNull(Format(oTrans.Master("dReleased"), "MMM-DD-YYYY"), "")
      Case 32
         txtField(pnCtr).Text = oApp.getUserName(IFNull(oTrans.Master("sPaymRecv"), ""))
      Case 33
         txtField(pnCtr).Text = IFNull(Format(oTrans.Master("dPaymRecv"), "MMM-DD-YYYY"), "")
      Case 34
      Case Else
         txtField(pnCtr).Text = IIf(IsNull(oTrans.Master(pnCtr)), "", oTrans.Master(pnCtr))
      End Select
   Next
   
   lblTotal.Caption = Format(oTrans.Master("nTranTotl"), "#,##0.00")
   lblTotalParts.Caption = Format(oTrans.Master("nPartsAmt"), "#,##0.00")
   lblTotalLabor.Caption = Format(oTrans.Master("nLaborAmt"), "#,##0.00")
   Label2.Caption = JobOrderStatus(oTrans.Master("cTranStat"))
   chkServiceType(oTrans.Master("cJOTypexx")).Value = True
   chkBackJob.Value = IFNull(oTrans.Master("cBackJobx"), Unchecked)
End Sub

Private Sub LoadDetail()
   With GridEditor1
      .Rows = IIf(oTrans.ItemCount = 0, 2, oTrans.ItemCount + 1)
      
      .ColWidth(2) = 3100
      If .Rows > 6 Then .ColWidth(2) = 2900
      
      For pnCtr = 1 To .Rows - 1
         If Not IsNull(oTrans.Detail(0, 1)) Then
            .TextMatrix(pnCtr, 1) = IIf(oTrans.Detail(pnCtr - 1, "sBarrCode") = "", "", oTrans.Detail(pnCtr - 1, "sBarrCode"))
            .TextMatrix(pnCtr, 2) = IIf(oTrans.Detail(pnCtr - 1, "sDescript") = "", "", oTrans.Detail(pnCtr - 1, "sDescript"))
            .TextMatrix(pnCtr, 3) = IFNull(oTrans.Detail(pnCtr - 1, "nQtyOnHnd"), 0)
            .TextMatrix(pnCtr, 5) = IFNull(oTrans.Detail(pnCtr - 1, "nQuantity"), "0.00")
            .TextMatrix(pnCtr, 4) = IFNull(oTrans.Detail(pnCtr - 1, "nUnitPrce"), 0)
            .TextMatrix(pnCtr, 6) = IFNull(oTrans.Detail(pnCtr - 1, "nDiscount"), 0) & "%"
            .TextMatrix(pnCtr, 7) = Format(TotalUnitPrice(.TextMatrix(pnCtr, 5), .TextMatrix(pnCtr, 4), Left(.TextMatrix(pnCtr, 6) _
            , Len(.TextMatrix(pnCtr, 6)) - 1)), "#,##0.00")
         Else
            .TextMatrix(pnCtr, 1) = ""
            .TextMatrix(pnCtr, 2) = ""
            .TextMatrix(pnCtr, 3) = 0
            .TextMatrix(pnCtr, 4) = 0#
            .TextMatrix(pnCtr, 5) = 0
            .TextMatrix(pnCtr, 6) = 0 & "%"
            .TextMatrix(pnCtr, 7) = 0#
         End If
      Next
   End With
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   With txtField(Index)
      .Text = TitleCase(.Text)
      Select Case Index
      Case 1, 10
         If Not IsDate(.Text) Then .Text = Date
         .Text = Format(.Text, "MMM-DD-YYYY")
      Case 2, 7, 9, 11, 13
         .Text = UCase(.Text)
      Case 25
         If Not IsNumeric(.Text) Then .Text = 0#
         .Text = Format(.Text, "#,##0.00")
      Case 34
         If Trim(.Text) = "" Then
            ClearFields
            Exit Sub
         End If
         
         If Trim(LCase(.Text)) <> Trim(LCase(.Tag)) Then
            If oBranch.SearchRecord(.Text, False) Then
               oTrans.Branch = oBranch.Master("sBranchCd")
               oTrans.InitTransaction
               oTrans.NewTransaction
               ClearFields
               
               .Text = oBranch.Master("sBranchNm")
            Else
               If Trim(.Tag) <> "" Then
                  .Text = .Tag
                  Exit Sub
               End If
               
               ClearFields
               .SetFocus
            End If
         End If
         
         .Tag = .Text
      Case 35, 36
         If .Text = "" Then
            ClearFields
            Exit Sub
         End If
         
         If Trim(LCase(.Text)) <> Trim(LCase(.Tag)) Then
            If oTrans.SearchOrigin( _
               psBranhCd, .Text, IIf(Index = 35, True, False)) Then
               LoadMaster
               LoadDetail
            Else
               ClearFields
               .SetFocus
            End If
         End If
      End Select
      
      If Index = 25 Then
         oTrans.Master("nMiscChrg") = CDbl(.Text)
         Call DisplayComputation
      Else
         If Index < 25 Then oTrans.Master(Index) = .Text
      End If
   End With
End Sub

Private Sub InitGrid()
   With GridEditor1
      .Rows = 2
      .Cols = 9
      .Font = "MS Sans Serif"

      'column title
      .TextMatrix(0, 1) = "Part #"
      .TextMatrix(0, 2) = "Description"
      .TextMatrix(0, 3) = "QOH."
      .TextMatrix(0, 4) = "Unit Price"
      .TextMatrix(0, 5) = "Qty."
      .TextMatrix(0, 6) = "Disc."
      .TextMatrix(0, 7) = "Total"
      
      .Row = 0

      'column alignment
      For pnCtr = 0 To .Cols - 1
         .Col = pnCtr
         .CellFontBold = True
         .CellAlignment = 1
         .ColEnabled(pnCtr) = False
      Next

      'column width
      .ColWidth(0) = 330
      .ColWidth(1) = 1600
      .ColWidth(3) = 700
      .ColWidth(4) = 1000
      .ColWidth(5) = 700
      .ColWidth(6) = 600
      .ColWidth(7) = 1000
      .ColWidth(8) = 0

      .ColDefault(3) = 0
      .ColDefault(4) = "0.00"
      .ColDefault(5) = 0
      .ColDefault(6) = "0" & "%"
      .ColDefault(7) = "0.00"
      .ColDefault(8) = 0
      
      .ColAlignment(1) = 1
      
      .ColNumberOnly(7) = True
      .ColFormat(4) = "#,##0.00"
      .ColFormat(7) = "#,##0.00"
      
      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ScrollBars = flexScrollBarVertical
      
      .EditorBackColor = oApp.getColor("HT1")
      
      .Row = 1
      .Col = 1
   End With
End Sub

Private Sub DisplayComputation()
   With GridEditor1
      If Trim(.TextMatrix(.Row, 4)) = "" Then .TextMatrix(.Row, 4) = 0#
      If Trim(.TextMatrix(.Row, 5)) = "" Then .TextMatrix(.Row, 5) = 0
      
      .TextMatrix(.Row, 7) = Format(TotalUnitPrice(.TextMatrix(.Row, 5), .TextMatrix(.Row, 4), Left(.TextMatrix(.Row, 6) _
      , Len(.TextMatrix(.Row, 6)) - 1)), "#,##0.00")
      lblTotal.Caption = Format(GrandTotal + oTrans.Master("nLaborAmt") + oTrans.Master("nMiscChrg"), "#,##0.00")
      
      oTrans.Master("nTranTotl") = CDbl(lblTotal.Caption)
   End With
End Sub

Private Function TotalUnitPrice(lnQuantity As Double _
   , lnUnitPrice As Double, lnDiscount As Double) As Double
   
   Dim lnUnitTotal As Double

   lnUnitTotal = 0#
      
   lnUnitTotal = CDbl(lnQuantity) * CDbl(lnUnitPrice) * _
                  (100 - CDbl(lnDiscount)) / 100
   
   TotalUnitPrice = lnUnitTotal
End Function

Private Function GrandTotal() As Double
   Dim lnCtr As Integer
   Dim lnSum As Double
   
   lnSum = 0#
   With GridEditor1
      For lnCtr = 1 To .Rows - 1
         lnSum = .TextMatrix(lnCtr, 7) + lnSum
      Next
   End With
   
   oTrans.Master("nPartsAmt") = lnSum
   lblTotalParts.Caption = Format(lnSum, "#,##0.00")
   GrandTotal = lnSum
End Function

Private Function inputDate(ByVal sLabelCap As String, _
                              ByRef dDateEntry As Date) As Boolean
   Dim loFormDate As frmDateCriteria
   
   Set loFormDate = New frmDateCriteria
   Set loFormDate.AppDriver = oApp
   
   loFormDate.Label1(0).Caption = sLabelCap
   loFormDate.Show 1
   inputDate = Not loFormDate.Cancelled
   dDateEntry = loFormDate.DateEntry
End Function

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

