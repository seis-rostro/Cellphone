VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmCP_Inventory 
   BorderStyle     =   0  'None
   Caption         =   "CP Inventory"
   ClientHeight    =   6420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11940
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6420
   ScaleWidth      =   11940
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame2 
      Height          =   4410
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   1065
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   7779
      BorderStyle     =   1
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   14
         Left            =   1095
         Locked          =   -1  'True
         TabIndex        =   83
         Top             =   1725
         Width           =   1890
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   14
         Left            =   10065
         TabIndex        =   81
         Top             =   3525
         Width           =   1425
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   16
         Left            =   10065
         TabIndex        =   79
         Top             =   2760
         Width           =   1425
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   15
         Left            =   10065
         TabIndex        =   77
         Top             =   3150
         Width           =   1425
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   13
         Left            =   7290
         TabIndex        =   76
         Top             =   3510
         Width           =   1425
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   12
         Left            =   7290
         TabIndex        =   75
         Top             =   3135
         Width           =   1425
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   11
         Left            =   7290
         TabIndex        =   74
         Top             =   2760
         Width           =   1425
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   10
         Left            =   1560
         TabIndex        =   25
         Top             =   3420
         Width           =   1890
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   9
         Left            =   1560
         TabIndex        =   23
         Top             =   3120
         Width           =   1890
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   8
         Left            =   1560
         TabIndex        =   21
         Top             =   2820
         Width           =   1890
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   7
         Left            =   1560
         TabIndex        =   19
         Top             =   2520
         Width           =   1890
      End
      Begin VB.TextBox txtOthers 
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
         Height          =   285
         Index           =   5
         Left            =   10065
         TabIndex        =   46
         Top             =   1200
         Width           =   1425
      End
      Begin VB.TextBox txtOthers 
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
         Height          =   285
         Index           =   4
         Left            =   6990
         TabIndex        =   44
         Top             =   1200
         Width           =   1425
      End
      Begin VB.TextBox txtOthers 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   2
         Left            =   6990
         TabIndex        =   40
         Text            =   "0,000"
         Top             =   705
         Width           =   1425
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   19
         Left            =   1095
         TabIndex        =   33
         Top             =   3885
         Width           =   2820
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   17
         Left            =   4455
         TabIndex        =   29
         Top             =   2010
         Width           =   1245
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   18
         Left            =   4455
         TabIndex        =   31
         Top             =   2310
         Width           =   1245
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   20
         Left            =   4455
         TabIndex        =   27
         Top             =   1710
         Width           =   1245
      End
      Begin VB.TextBox txtOthers 
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
         Height          =   285
         Index           =   11
         Left            =   10065
         Locked          =   -1  'True
         TabIndex        =   58
         Top             =   2250
         Width           =   1425
      End
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   8
         Left            =   7290
         TabIndex        =   52
         Top             =   2250
         Width           =   1350
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   1095
         TabIndex        =   5
         Top             =   240
         Width           =   2820
      End
      Begin VB.CheckBox chkHsSerial 
         Caption         =   "w/ Serial"
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
         Left            =   4605
         TabIndex        =   34
         Tag             =   "wt0;fb0"
         Top             =   3870
         Width           =   1095
      End
      Begin VB.TextBox txtOthers 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   10
         Left            =   10065
         Locked          =   -1  'True
         TabIndex        =   56
         Top             =   1950
         Width           =   1425
      End
      Begin VB.TextBox txtOthers 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   7
         Left            =   7290
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   1950
         Width           =   1350
      End
      Begin VB.TextBox txtOthers 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   9
         Left            =   10065
         Locked          =   -1  'True
         TabIndex        =   54
         Top             =   1650
         Width           =   1425
      End
      Begin VB.TextBox txtOthers 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   7290
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   1650
         Width           =   1350
      End
      Begin VB.TextBox txtOthers 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   3
         Left            =   10065
         TabIndex        =   42
         Text            =   "0,000"
         Top             =   705
         Width           =   1425
      End
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   9795
         TabIndex        =   38
         Top             =   240
         Width           =   1680
      End
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   6735
         TabIndex        =   36
         Top             =   240
         Width           =   1680
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   6
         Left            =   1095
         TabIndex        =   17
         Top             =   2040
         Width           =   1890
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   3810
         TabIndex        =   15
         Top             =   1410
         Width           =   1890
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   3810
         TabIndex        =   13
         Top             =   1110
         Width           =   1890
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   1095
         TabIndex        =   11
         Top             =   1410
         Width           =   1890
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   1095
         TabIndex        =   9
         Top             =   1110
         Width           =   1890
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   1095
         TabIndex        =   7
         Top             =   810
         Width           =   4605
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model Code"
         Height          =   195
         Index           =   34
         Left            =   195
         TabIndex        =   84
         Top             =   1725
         Width           =   855
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sel Price"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   12
         Left            =   8940
         TabIndex        =   82
         Top             =   3555
         Width           =   1080
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Price as of"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   33
         Left            =   8730
         TabIndex        =   80
         Top             =   2790
         Width           =   1290
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Price"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   32
         Left            =   8805
         TabIndex        =   78
         Top             =   3180
         Width           =   1215
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   31
         Left            =   6870
         TabIndex        =   73
         Top             =   3540
         Width           =   315
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   30
         Left            =   7020
         TabIndex        =   72
         Top             =   3165
         Width           =   165
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3 Months"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   27
         Left            =   6045
         TabIndex        =   71
         Top             =   2790
         Width           =   1125
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cat 04"
         Height          =   195
         Index           =   29
         Left            =   645
         TabIndex        =   24
         Top             =   3480
         Width           =   465
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cat 03"
         Height          =   195
         Index           =   28
         Left            =   645
         TabIndex        =   22
         Top             =   3150
         Width           =   465
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cat 02"
         Height          =   195
         Index           =   26
         Left            =   645
         TabIndex        =   20
         Top             =   2865
         Width           =   465
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cat 01"
         Height          =   195
         Index           =   19
         Left            =   645
         TabIndex        =   18
         Top             =   2565
         Width           =   465
      End
      Begin VB.Line Line3 
         Index           =   3
         X1              =   1260
         X2              =   2130
         Y1              =   3585
         Y2              =   3585
      End
      Begin VB.Line Line3 
         Index           =   2
         X1              =   1260
         X2              =   2130
         Y1              =   2670
         Y2              =   2670
      End
      Begin VB.Line Line3 
         Index           =   1
         X1              =   1260
         X2              =   2130
         Y1              =   3255
         Y2              =   3255
      End
      Begin VB.Line Line3 
         Index           =   0
         X1              =   1260
         X2              =   2130
         Y1              =   2970
         Y2              =   2970
      End
      Begin VB.Line Line2 
         X1              =   1260
         X2              =   1260
         Y1              =   2340
         Y2              =   3600
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DmBeg Bal."
         Height          =   195
         Index           =   22
         Left            =   6045
         TabIndex        =   43
         Top             =   1230
         Width           =   840
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dm. On Hand"
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
         Left            =   8850
         TabIndex        =   45
         Top             =   1230
         Width           =   1155
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Part No"
         Height          =   195
         Index           =   17
         Left            =   195
         TabIndex        =   32
         Top             =   3915
         Width           =   795
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reserve Order"
         Height          =   195
         Index           =   16
         Left            =   8985
         TabIndex        =   57
         Top             =   2310
         Width           =   1035
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Max. Level"
         Height          =   195
         Index           =   25
         Left            =   6060
         TabIndex        =   49
         Top             =   1995
         Width           =   780
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Back Order"
         Height          =   195
         Index           =   14
         Left            =   9210
         TabIndex        =   55
         Top             =   1995
         Width           =   810
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Discount 3"
         Height          =   195
         Index           =   11
         Left            =   3600
         TabIndex        =   30
         Top             =   2370
         Width           =   765
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Discount 2"
         Height          =   195
         Index           =   10
         Left            =   3600
         TabIndex        =   28
         Top             =   2070
         Width           =   765
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Discount 1"
         Height          =   195
         Index           =   9
         Left            =   3600
         TabIndex        =   26
         Top             =   1755
         Width           =   765
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Level"
         Height          =   195
         Index           =   3
         Left            =   9120
         TabIndex        =   37
         Top             =   270
         Width           =   390
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Section"
         Height          =   195
         Index           =   2
         Left            =   6030
         TabIndex        =   35
         Top             =   270
         Width           =   540
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Beg. Inv. Date"
         Height          =   195
         Index           =   13
         Left            =   6060
         TabIndex        =   51
         Top             =   2310
         Width           =   1035
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   6000
         X2              =   11490
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   6000
         X2              =   11490
         Y1              =   615
         Y2              =   615
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Min. Level"
         Height          =   195
         Index           =   24
         Left            =   6060
         TabIndex        =   47
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ReOrder Level"
         Height          =   210
         Index           =   23
         Left            =   8970
         TabIndex        =   53
         Top             =   1680
         Width           =   1035
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qty. On Hand"
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
         Index           =   21
         Left            =   8655
         TabIndex        =   41
         Top             =   795
         Width           =   1380
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Beg. Bal."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   20
         Left            =   6045
         TabIndex        =   39
         Top             =   780
         Width           =   810
      End
      Begin VB.Shape Shape4 
         Height          =   1590
         Left            =   5895
         Top             =   2685
         Width           =   5715
      End
      Begin VB.Shape Shape3 
         Height          =   2535
         Left            =   5895
         Top             =   120
         Width           =   5715
      End
      Begin VB.Shape Shape2 
         Height          =   4140
         Left            =   90
         Top             =   120
         Width           =   5700
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
         Height          =   195
         Index           =   8
         Left            =   195
         TabIndex        =   16
         Top             =   2055
         Width           =   630
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Color"
         Height          =   195
         Index           =   7
         Left            =   3315
         TabIndex        =   14
         Top             =   1425
         Width           =   360
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Size"
         Height          =   195
         Index           =   6
         Left            =   3315
         TabIndex        =   12
         Top             =   1125
         Width           =   300
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model Desc"
         Height          =   195
         Index           =   5
         Left            =   195
         TabIndex        =   10
         Top             =   1410
         Width           =   855
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Brand"
         Height          =   315
         Index           =   4
         Left            =   195
         TabIndex        =   8
         Top             =   1125
         Width           =   420
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   285
         Left            =   1200
         Tag             =   "et0;ht2"
         Top             =   360
         Width           =   2820
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Barcode"
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   4
         Top             =   270
         Width           =   600
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   1
         Left            =   195
         TabIndex        =   6
         Top             =   825
         Width           =   795
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   0
      Left            =   11085
      TabIndex        =   66
      Top             =   5625
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmCP_Inventory.frx":0000
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   1
      Left            =   10305
      TabIndex        =   65
      Top             =   5625
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmCP_Inventory.frx":077A
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   2
      Left            =   7185
      TabIndex        =   59
      Top             =   5625
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmCP_Inventory.frx":0EF4
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   4
      Left            =   7185
      TabIndex        =   61
      Top             =   5625
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmCP_Inventory.frx":166E
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   5
      Left            =   11085
      TabIndex        =   67
      Top             =   5625
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmCP_Inventory.frx":1DE8
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   705
      Index           =   7
      Left            =   7965
      TabIndex        =   60
      Top             =   5625
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmCP_Inventory.frx":2562
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   3
      Left            =   7965
      TabIndex        =   62
      Top             =   5625
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmCP_Inventory.frx":2CDC
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   6
      Left            =   7965
      TabIndex        =   69
      Top             =   5625
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmCP_Inventory.frx":3456
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   8
      Left            =   9525
      TabIndex        =   64
      Top             =   5625
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmCP_Inventory.frx":3BD0
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   9
      Left            =   7965
      TabIndex        =   68
      Top             =   5625
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmCP_Inventory.frx":434A
      PicturePos      =   1
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   480
      Index           =   1
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   847
      Begin VB.TextBox txtOthers 
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
         Index           =   13
         Left            =   5910
         TabIndex        =   3
         Top             =   90
         Width           =   5700
      End
      Begin VB.TextBox txtOthers 
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
         IMEMode         =   3  'DISABLE
         Index           =   12
         Left            =   1095
         MaxLength       =   50
         TabIndex        =   1
         Top             =   90
         Width           =   2820
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Description"
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
         Left            =   4890
         TabIndex        =   2
         Top             =   135
         Width           =   960
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Barc&ode"
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
         Left            =   195
         TabIndex        =   0
         Top             =   135
         Width           =   885
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   10
      Left            =   8745
      TabIndex        =   63
      Top             =   5625
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
      Caption         =   "Serial"
      AccessKey       =   "Serial"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCP_Inventory.frx":4AC4
      PicturePos      =   1
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cat 1"
      Height          =   195
      Index           =   15
      Left            =   720
      TabIndex        =   70
      Top             =   3375
      Width           =   375
   End
End
Attribute VB_Name = "frmCP_Inventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmCP_Inventory"

Private WithEvents oDriver As clsFormDriver
Attribute oDriver.VB_VarHelpID = -1
Private oFormSerialNew As frmCPSerial
Private oSkin As clsFormSkin
Private bLoaded As Boolean
Private oRS As New ADODB.Recordset

Dim pbtxtOthers As Boolean
Dim pnCtr As Integer, pnIndex As Integer

Dim psCPInventory As String
Dim pbEnblButtons As Boolean
Dim pbNewInvntory As Boolean
Dim psPriceCode As String
Dim psConcatDescx As String

Private Sub chkHsSerial_Click()
   chkHsSerial.Value = oDriver.FieldValue(21)
End Sub

Private Sub cmdButton_Click(Index As Integer)
   Dim lsSearch As String
   Dim lnRep As Integer
   Dim lsOldProc As String
   
   lsOldProc = "cmdButton_Click"
   '''On Error GoTo errProc
   
   If Index = 2 Then
      If pbtxtOthers Then
         Call txtOthers_Validate(pnIndex, False)
      Else
         Call txtField_Validate(pnIndex, False)
      End If
   End If
   
   Select Case Index
   Case 0 'cancel
      oDriver.RecordCancelUpdate
      pbEnblButtons = False
   Case 1 'browse
      oDriver.BrowseRecord
   Case 2 'save
      oDriver.RecordSave
   Case 3 'update
'      If Not IsNumeric(txtField(11).Text) Then
'         txtField(11).Text = Format(Code2Price(txtField(11).Text, psPriceCode), "#,##0.00")
'         txtField(12).Text = Format(Code2Price(txtField(12).Text, psPriceCode), "#,##0.00")
'         txtField(13).Text = Format(Code2Price(txtField(13).Text, psPriceCode), "#,##0.00")
'         txtField(14).Text = Format(Code2Price(txtField(14).Text, psPriceCode), "#,##0.00")
'         txtField(15).Text = Format(Code2Price(txtField(15).Text, psPriceCode), "#,##0.00")
'      End If
      oDriver.RecordUpdate
   Case 4 'new
      oDriver.RecordNew
   Case 5 'close
      Unload Me
   Case 6 'delete
      oDriver.RecordDelete
   Case 7 'search
      If pbtxtOthers Then
         If pnIndex = 2 And oDriver.FieldValue(2) <> "" Then
            oDriver.LookupQuery(pnIndex) = AddCondition(oDriver.LookupQuery(pnIndex), "sBrandIDx = " & strParm(oDriver.FieldValue(2)))
         End If
      
         oDriver.RecordSearch
         txtField(pnIndex).SetFocus
      Else
         SearchOthers pnIndex, Empty, False
         txtOthers(pnIndex).SetFocus
      End If
   Case 8 'ledger
      If Not pbNewInvntory Then
         With frmCP_InventoryLedger
            .txtField(0) = txtField(0)
            .txtField(1) = txtField(1)
            .txtField(2) = txtField(2)
            .txtField(3) = txtField(3)
            .cmbField.ListIndex = 1
            
            .StockID = oDriver.FieldValue(22)
            .Show 1
         End With
      Else
         MsgBox "Unable to Load Inventory Ledger!!!" & vbCrLf & _
                "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
      End If
   Case 10
      If oDriver.FieldValue(21) = xeYes Then
         With oFormSerialNew
            .StockID = oDriver.FieldValue(22)
            .Barcode = oDriver.FieldValue(0)
            .Description = oDriver.FieldValue(1)
            .Brand = txtField(2).Text
            .Model = txtField(3).Text
            .Color = txtField(5).Text
            .Category = txtField(6).Text
            .Branch = oApp.BranchCode
            
            .Show 1
         End With
      End If
   End Select

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Sub Form_Activate()
   Dim lsOldProc As String
   
   lsOldProc = "Form_Activate"
   '''On Error GoTo errProc
   
   oApp.MenuName = Me.Tag
   Me.ZOrder 0
   
   If bLoaded = False Then
      oDriver.RecordCancelUpdate
      oDriver_InitValue
      bLoaded = True
      txtOthers(12).SetFocus
   End If
   mdiMain.StatusBar1.Panels(1).Text = "Press F9 to encrypt selling price!!!"
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Load()
   Dim lsSQL As String
   Dim lsOldProc As String
   
   lsOldProc = "Form_Load"
   '''On Error GoTo errProc

   CenterChildForm mdiMain, Me
   
   bLoaded = False
   
   Set oFormSerialNew = New frmCPSerial
   Set oRS = New ADODB.Recordset
   
   Set oDriver = New clsFormDriver
   Set oDriver.AppDriver = oApp
   Set oDriver.MainForm = Me
   
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin
                           
   oDriver.RecQuery = "SELECT" _
                           & "  sBarrcode" _
                           & ", sDescript" _
                           & ", sBrandIDx" _
                           & ", sModelIDx" _
                           & ", sSizeIDxx" _
                           & ", sColorIDx" _
                           & ", sCategID1" _
                           & ", sCategID2" _
                           & ", sCategID3" _
                           & ", sCategID4" _
                           & ", sCategID5" _
                           & ", nSelPrce4" _
                           & ", nSelPrce3" _
                           & ", nSelPrce2" _
                           & ", nSelPrice" _
                           & ", nLastPrce" _
                           & ", dPrceAsOf" _
                           & ", nMaxDisc1" _
                           & ", nMaxDisc2" _
                           & ", nMaxDisc3" _
                           & ", sPartNoxx" _
                           & ", cHsSerial" _
                           & ", sStockIDx" _
                           & ", cRecdStat"
   oDriver.RecQuery = oDriver.RecQuery _
                           & ", sModified" _
                           & ", dModified" _
                        & " FROM CP_Inventory"
                        
   'iMac 2016.02.09
   '  Bawasan ko lang ah, masyado kasing mahaba :)
'   psConcatDescx = "CONCAT(a.sDescript, ' '" _
'                           & ", IF(b.sBrandNme IS NULL, '', b.sBrandNme), ' '" _
'                           & ", IF(c.sModelNme IS NULL, '', c.sModelNme), ' '" _
'                           & ", IF(c.sModelCde IS NULL, '', c.sModelCde), ' '" _
'                           & ", IF(d.sColorNme IS NULL, '', d.sColorNme), ' '" _
'                           & ", IF(f.sSizeName IS NULL, '', f.sSizeName))"

   psConcatDescx = "CONCAT(a.sDescript, ' ', IF(c.sModelNme IS NULL, '', c.sModelNme))" _
   
   oDriver.BrowseQuery = "SELECT" _
                              & "  a.sBarrcode" _
                              & ", " & psConcatDescx & " xDescript" _
                              & ", b.sBrandNme" _
                              & ", c.sModelCde" _
                              & ", c.sModelNme" _
                              & ", f.sSizeName" _
                              & ", d.sColorNme" _
                              & ", e.sCategrNm" _
                              & ", g.nQtyonHnd" _
                           & " FROM CP_Inventory a" _
                              & " LEFT JOIN CP_Brand b" _
                                 & " ON a.sBrandIDx = b.sBrandIDx" _
                              & " LEFT JOIN CP_Model c" _
                                 & " ON a.sModelIDx = c.sModelIDx" _
                              & " LEFT JOIN Color d" _
                                 & " ON a.sColorIDx = d.sColorIDx" _
                              & " LEFT JOIN Size f" _
                                 & " ON a.sSizeIDxx = f.sSizeIDxx" _
                              & " LEFT JOIN Category e" _
                                 & " ON a.sCategID1 = e.sCategrID" _
                              & ", CP_Inventory_Master g" _
                           & " WHERE a.sStockIDx = g.sStockIDx" _
                              & " AND g.sBranchCd = " & strParm(oApp.BranchCode)
   Debug.Print oDriver.BrowseQuery
   oDriver.InitRecForm
      
   oDriver.BrowseColumn(0) = "sBarrcode"
   oDriver.BrowseColumn(1) = "xDescript"
   oDriver.BrowseColumn(2) = "sBrandNme"
   oDriver.BrowseColumn(3) = "sModelCde"
   oDriver.BrowseColumn(4) = "sModelNme"
   oDriver.BrowseColumn(5) = "sColorNme"
   oDriver.BrowseColumn(6) = "sCategrNm"
   oDriver.BrowseColumn(7) = "nQtyOnHnd"
   
   oDriver.BrowseFTitle(0) = "Barcode"
   oDriver.BrowseFTitle(1) = "Description"
   oDriver.BrowseFTitle(2) = "Brand"
   oDriver.BrowseFTitle(3) = "Model Code"
   oDriver.BrowseFTitle(4) = "Model"
   oDriver.BrowseFTitle(5) = "Color"
   oDriver.BrowseFTitle(6) = "Category"
   oDriver.BrowseFTitle(7) = "QOH"

   oDriver.LookupQuery(2) = "SELECT" _
                              & "  sBrandIDx" _
                              & ", sBrandNme" _
                           & " FROM CP_Brand" _
                           & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
                           & " ORDER BY sBrandNme"
   oDriver.LookupReference(2) = "sBrandIDx»sBrandNme"
   oDriver.LookupColumn(2) = "sBrandIDx»sBrandNme"
   oDriver.LookupTitle(2) = "Code»Brand"
   
   oDriver.LookupQuery(3) = "SELECT" _
                              & "  a.sModelIDx" _
                              & ", a.sModelNme" _
                              & ", IFNull(a.sModelCde, '') sModelCde" _
                              & ", b.sBrandNme" _
                           & " FROM CP_Model a" _
                              & ", CP_Brand b" _
                           & " WHERE a.sBrandIDx = b.sBrandIDx" _
                              & " AND a.cRecdStat = " & strParm(xeRecStateActive) _
                           & " ORDER BY sModelNme"
   oDriver.LookupReference(3) = "a.sModelIDx»a.sModelNme»a.sModelCde»b.sBrandNme"
   oDriver.LookupColumn(3) = "sModelIDx»sModelNme»sModelCde»sBrandNme"
   oDriver.LookupTitle(3) = "ID»Model Name»Model Code»Brand"

   oDriver.LookupQuery(4) = "SELECT" _
                                 & "  sSizeIDxx" _
                                 & ", sSizeName" _
                           & " FROM Size" _
                           & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
                           & " ORDER BY sSizeName"
   oDriver.LookupReference(4) = "sSizeIDxx»sSizeName"
   oDriver.LookupColumn(4) = "sSizeName"
   oDriver.LookupTitle(4) = "Size Name"
   
    oDriver.LookupQuery(5) = "SELECT" _
                              & "  sColorIDx" _
                              & ", sColorNme" _
                           & " FROM Color" _
                           & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
                              & " AND sColorIDx LIKE 'C0%'" _
                           & " ORDER BY sColorNme"
   oDriver.LookupReference(5) = "sColorIDx»sColorNme"
   oDriver.LookupColumn(5) = "sColorNme"
   oDriver.LookupTitle(5) = "Color Name"
    
   oDriver.LookupQuery(6) = "SELECT" _
                              & "  sCategrID" _
                              & ", sCategrNm" _
                           & " FROM Category" _
                           & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
                           & " AND cLevelxxx = '1' " _
                           & " ORDER BY sCategrNm"
   oDriver.LookupReference(6) = "sCategrID»sCategrNm"
   oDriver.LookupColumn(6) = "sCategrNm"
   oDriver.LookupTitle(6) = "Category Name"
   
   oDriver.LookupQuery(7) = "SELECT" _
                              & "  sCategrID" _
                              & ", sCategrNm" _
                              & ", cSerialze" _
                           & " FROM Category" _
                           & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
                           & " AND cLevelxxx = '2' " _
                           & " ORDER BY sCategrNm"
   oDriver.LookupReference(7) = "sCategrID»sCategrNm»cSerialze"
   oDriver.LookupColumn(7) = "sCategrNm»cSerialze"
   oDriver.LookupTitle(7) = "Category Name»Serialize"
   
   oDriver.LookupQuery(8) = "SELECT" _
                              & "  sCategrID" _
                              & ", sCategrNm" _
                              & ", cSerialze" _
                           & " FROM Category" _
                           & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
                           & " AND cLevelxxx = '3' " _
                           & " ORDER BY sCategrNm"
   oDriver.LookupReference(8) = "sCategrID»sCategrNm»cSerialze"
   oDriver.LookupColumn(8) = "sCategrNm»cSerialze"
   oDriver.LookupTitle(8) = "Category Name»Serialize"
   
   oDriver.LookupQuery(9) = "SELECT" _
                              & "  sCategrID" _
                              & ", sCategrNm" _
                              & ", cSerialze" _
                           & " FROM Category" _
                           & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
                           & " AND cLevelxxx = '4' " _
                           & " ORDER BY sCategrNm"
   oDriver.LookupReference(9) = "sCategrID»sCategrNm»cSerialze"
   oDriver.LookupColumn(9) = "sCategrNm»cSerialze"
   oDriver.LookupTitle(9) = "Category Name»Serialize"
   
   oDriver.LookupQuery(10) = "SELECT" _
                              & "  sCategrID" _
                              & ", sCategrNm" _
                           & " FROM Category" _
                           & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
                           & " AND cLevelxxx = '5' " _
                           & " ORDER BY sCategrNm"
   oDriver.LookupReference(10) = "sCategrID»sCategrNm"
   oDriver.LookupColumn(10) = "sCategrNm"
   oDriver.LookupTitle(10) = "Category Name"
                        
   psCPInventory = "SELECT" _
                     & "  sSectnIDx" _
                     & ", sLevelIDx" _
                     & ", nBegQtyxx" _
                     & ", nQtyOnHnd" _
                     & ", nDmoQtyxx" _
                     & ", nDemoUnit" _
                     & ", nMinLevel" _
                     & ", nMaxLevel" _
                     & ", dBegInvxx" _
                     & ", nReorderx" _
                     & ", nBackOrdr" _
                     & ", nResvOrdr" _
                     & ", nFloatQty" _
                     & ", nLedgerNo" _
                     & ", dLastTran" _
                     & ", cRecdStat" _
                     & ", sStockIDx" _
                     & ", sBranchCd" _
                     & ", sModified" _
                     & ", dModified" _
                  & " FROM CP_Inventory_Master" _
                  & " ORDER BY sBranchCd"
   
   Set oRS = New ADODB.Recordset
   lsSQL = AddCondition(psCPInventory, "0 = 1")
   oRS.Open lsSQL, oApp.Connection, adOpenStatic, adLockPessimistic, adCmdText
   Set oRS.ActiveConnection = Nothing
   
   oDriver.FieldStart = 0
   oDriver.FieldFormat(0) = ">"
   oDriver.FieldFormat(16) = ">"
   
   oDriver.FieldFormat(11) = "#,##0.00"
   oDriver.FieldFormat(12) = "#,##0.00"
   
   psPriceCode = "PATRONIZEX"
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oDriver = Nothing
   Set oSkin = Nothing
   Set oRS = Nothing
   Set oFormSerialNew = Nothing

   mdiMain.StatusBar1.Panels(1).Text = ""
End Sub

Private Sub oDriver_DisableOtherControl()
   For pnCtr = 0 To txtOthers.Count - 1
      txtOthers(pnCtr).Enabled = False
   Next
   
   txtOthers(12).Enabled = True
   txtOthers(13).Enabled = True
   
   oDriver.hideButton 6
   oDriver.hideButton 9
End Sub

Private Sub oDriver_EnableOtherControl()
   For pnCtr = 0 To txtOthers.Count - 1
      Select Case pnCtr
      Case 12, 13
         txtOthers(pnCtr).Enabled = False
      Case Else
         txtOthers(pnCtr).Enabled = True
      End Select
   Next

   If oDriver.EditMode = xeModeUpdate Then
      For pnCtr = 0 To txtField.Count - 1
         Select Case pnCtr
         Case 14, 15
            If chkHsSerial.Value <> Checked Then 'she 2019-11-12
               txtField(pnCtr).Locked = IIf(oApp.IsWarehouse, False, True)
            End If
         Case Else
            txtField(pnCtr).Locked = IIf(oApp.IsWarehouse, False, True)
         End Select
      Next

      txtOthers(2).Enabled = pbEnblButtons
      txtOthers(3).Enabled = pbEnblButtons
      txtOthers(4).Enabled = pbEnblButtons
      txtOthers(5).Enabled = pbEnblButtons
      txtOthers(8).Enabled = pbEnblButtons
   Else
      For pnCtr = 0 To txtField.Count - 1
         txtField(pnCtr).Locked = False
      Next
   End If
End Sub

Private Sub InitOthers()
   For pnCtr = 0 To txtOthers.Count - 3
      Select Case pnCtr
      Case 2 To 7, 9, 10, 11
         oRS(pnCtr) = 0
         txtOthers(pnCtr).Text = Format(oRS(pnCtr), "#,##0")
      Case 8
         oRS(pnCtr) = oApp.ServerDate
         txtOthers(pnCtr).Text = Format(oRS(pnCtr), "MMM DD, YYYY")
      Case 12
         txtOthers(pnCtr).Text = ""
      Case Else
         oRS(pnCtr) = Empty
         txtOthers(pnCtr).Text = oRS(pnCtr)
      End Select
   Next

   oRS("cRecdStat") = xeRecStateActive
   oRS("sStockIDx") = oDriver.FieldValue(22)
   oRS("sBranchCd") = oApp.BranchCode
   oRS("nFloatQty") = 0
   oRS("nLedgerNo") = 0
End Sub

Private Sub oDriver_InitValue()
   Dim lsOldProc As String
   
   lsOldProc = "oDriver_InitValue"
   '''On Error GoTo errProc

   oDriver.FieldReference(0) = True
   oDriver.FieldValue(0) = NewBarrCode
   txtField(0).Text = oDriver.FieldValue(0)
   oDriver.FieldValue(2) = ""
   oDriver.FieldValue(3) = ""
   oDriver.FieldValue(4) = ""
   oDriver.FieldValue(5) = ""
   oDriver.FieldValue(22) = GetNextCode("CP_Inventory", "sStockIDx", True, oApp.Connection, True, oApp.BranchCode)
   oDriver.FieldValue(23) = xeRecStateActive
   
   For pnCtr = 0 To txtOthers.Count - 1
      Select Case pnCtr
      Case 2 To 7, 9, 10, 11
         txtOthers(pnCtr).Text = 0
      Case 8
         txtOthers(pnCtr).Text = Format(oApp.ServerDate, "MMM DD, YYYY")
      Case Else
         txtOthers(pnCtr).Text = ""
      End Select
      txtOthers(pnCtr).Tag = ""
   Next
   
   chkHsSerial.Value = 0
   oDriver.FieldValue(21) = chkHsSerial.Value
   
   txtOthers(2).Locked = False
   txtOthers(3).Locked = False
   txtOthers(4).Locked = False
   txtOthers(5).Locked = False
   txtOthers(8).Locked = False
'   txtOthers(6).Locked = False
'   txtOthers(7).Locked = False
   
   oRS.AddNew
   InitOthers
   pbEnblButtons = True
   pbNewInvntory = True
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )"
End Sub

Private Sub oDriver_LoadOtherData()
   Dim lsOldProc As String
   Dim lsSQL As String
   
   lsOldProc = "oDriver_LoadOtherData"
   '''On Error GoTo errProc
   
   Set oRS = New ADODB.Recordset
   lsSQL = AddCondition(psCPInventory, "sStockIDx = " & strParm(oDriver.FieldValue(22)) _
                                    & " AND sBranchCd = " & strParm(oApp.BranchCode)) _
                                    & " AND cRecdStat = " & strParm(xeRecStateActive)
   oRS.Open lsSQL, oApp.Connection, adOpenStatic, adLockOptimistic, adCmdText
   Set oRS.ActiveConnection = Nothing
         
   If oRS.EOF Then
      oRS.AddNew
      InitOthers
   Else
      For pnCtr = 0 To txtOthers.Count - 1
         Select Case pnCtr
         Case 0
            If Not IsNull(oRS("sSectnIDx")) Then SearchOthers pnCtr, oRS("sSectnIDx"), True
         Case 1
            If Not IsNull(oRS("sLevelIDx")) Then SearchOthers pnCtr, oRS("sLevelIDx"), True
         Case 2 To 7, 9, 10, 11
            txtOthers(pnCtr).Text = Format(IIf(IsNull(oRS(pnCtr)), 0, oRS(pnCtr)), "#,##0")
         Case 14
            txtOthers(pnCtr).Text = getModelCode(oDriver.FieldValue(3))
         Case 8
            txtOthers(pnCtr).Text = IIf(IsNull(oRS(pnCtr)), "", Format(oRS(pnCtr), "MMM DD, YYYY"))
         End Select
      Next
         chkHsSerial.Value = IIf(oDriver.FieldValue(21) = "", Unchecked, oDriver.FieldValue(21))
      pbNewInvntory = False
   End If

   txtOthers(12).Text = oDriver.FieldValue(0)
   txtOthers(12).Tag = txtOthers(12).Text

   txtOthers(13).Text = oDriver.FieldValue(1)
   txtOthers(13).Tag = txtOthers(13).Text
   
   If oApp.IsWarehouse And oApp.UserLevel >= xeManager Then 'oApp.UserLevel > xeSupervisor
      txtField(11).Text = Format(IIf(IsNull(oDriver.FieldValue(11)), 0, oDriver.FieldValue(11)), "#,##0.00")
      txtField(12).Text = Format(IIf(IsNull(oDriver.FieldValue(12)), 0, oDriver.FieldValue(12)), "#,##0.00")
      txtField(13).Text = Format(IIf(IsNull(oDriver.FieldValue(13)), 0, oDriver.FieldValue(13)), "#,##0.00")
      txtField(14).Text = Format(IIf(IsNull(oDriver.FieldValue(14)), 0, oDriver.FieldValue(14)), "#,##0.00")
      txtField(15).Text = Format(IIf(IsNull(oDriver.FieldValue(15)), 0, oDriver.FieldValue(15)), "#,##0.00")
   Else
      txtField(11).Text = Format(Price2Code(IIf(IsNull(oDriver.FieldValue(11)), 0, oDriver.FieldValue(11)), psPriceCode), "#,##0.00")
      txtField(12).Text = Format(Price2Code(IIf(IsNull(oDriver.FieldValue(12)), 0, oDriver.FieldValue(12)), psPriceCode), "#,##0.00")
      txtField(13).Text = Format(Price2Code(IIf(IsNull(oDriver.FieldValue(13)), 0, oDriver.FieldValue(13)), psPriceCode), "#,##0.00")
      txtField(14).Text = Format(Price2Code(IIf(IsNull(oDriver.FieldValue(14)), 0, oDriver.FieldValue(14)), psPriceCode), "#,##0.00")
      txtField(15).Text = Format(Price2Code(IIf(IsNull(oDriver.FieldValue(15)), 0, oDriver.FieldValue(15)), psPriceCode), "#,##0.00")
   End If

'   txtField(14).Text = Format(IIf(IsNull(oDriver.FieldValue(14)), 0, oDriver.FieldValue(14)), "#,##0.00")
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )"
End Sub

Private Sub oDriver_Save(Saved As Boolean)
   Saved = False
End Sub

Private Sub oDriver_WillSave(Cancel As Boolean)
   Dim lsOldProc As String
   Dim lnCtr As Integer

   lsOldProc = "oDriver_WillSave"
   '''On Error GoTo errProc

   If oDriver.FieldValue(0) = "" Then
      MsgBox "Invalid BarrCode detected!!!", vbCritical, "Warning"
      txtField(0).SetFocus
      Cancel = True
   ElseIf oDriver.FieldValue(1) = "" Then
      MsgBox "Invalid Description detected!!!", vbCritical, "Warning"
      txtField(1).SetFocus
      Cancel = True
   ElseIf txtOthers(0).Text = "" Then
      MsgBox "Invalid Section detected!!!", vbCritical, "Warning"
      txtOthers(0).SetFocus
      Cancel = True
   ElseIf txtOthers(1).Text = "" Then
      MsgBox "Invalid Level detected!!!", vbCritical, "Warning"
      txtOthers(1).SetFocus
      Cancel = True
'   ElseIf CDbl(txtField(13).Text) = 0# Then
'      MsgBox "Invalid Selling Price detected!!!", vbCritical, "Warning"
'      txtField(13).SetFocus
'      Cancel = True
   ElseIf oDriver.FieldValue(6) = "" Then
      MsgBox "Invalid Category Detected!!!", vbCritical, "Warning"
      txtField(6).SetFocus
      Cancel = True
   ElseIf oDriver.FieldValue(7) = "" Then
      MsgBox "Invalid Sub Category Detected!!!", vbCritical, "Warning"
      txtField(7).SetFocus
      Cancel = True
   ElseIf oDriver.FieldValue(22) = "" Then
      MsgBox "Invalid Stock ID Detected!!!" & vbCrLf & _
               "Please contact GMC_SEG for assistant!!!", vbCritical, "Warning"
      Cancel = True
   Else
      Cancel = Not UpdateCPInventory
'      If pbNewInvntory Then Cancel = Not SaveCPInventoryLedger
   End If
   

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Cancel & " )"
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = oApp.getColor("HT1")
   End With

   oDriver.ColumnIndex = Index
   pbtxtOthers = False
   pnIndex = Index
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String
   Dim lsOldQuery As String
   
   lsOldProc = "txtField_KeyDown"
   '''On Error GoTo errProc
   
   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
      With txtField(Index)
         If KeyCode = vbKeyF3 Then
            oDriver.RecordSearch .Text
            If .Text <> "" Then
               If Index = 3 Then txtOthers(14) = getModelCode(oDriver.FieldValue(3))
               SetNextFocus
            End If
         Else
            If Index = 3 And oDriver.FieldValue(2) <> "" Then
               lsOldQuery = oDriver.LookupQuery(Index)
               oDriver.LookupQuery(Index) = AddCondition(oDriver.LookupQuery(Index), "a.sBrandIDx = " & strParm(oDriver.FieldValue(2)))
            End If
         
            If .Text <> "" Then oDriver.RecordSearch .Text
            
            If .Text <> "" Then
               If Index = 3 Then
                  txtOthers(14) = getModelCode(oDriver.FieldValue(3))
                  oDriver.LookupQuery(Index) = lsOldQuery
               End If
               SetNextFocus
            End If
            
         End If
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

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   Dim lsOldProc As String
   Dim lsOldQuery As String
   
   lsOldProc = "txtField_Validate"
   '''On Error GoTo errProc
   
   With txtField(Index)
      Select Case Index
      Case 0
         .Text = UCase(.Text)
      Case 3
         lsOldQuery = oDriver.LookupQuery(Index)
         If oDriver.FieldValue(2) <> "" Then
            oDriver.LookupQuery(Index) = AddCondition(oDriver.LookupQuery(Index), "a.sBrandIDx = " & strParm(oDriver.FieldValue(2)))
         End If
      Case 11, 12, 13, 15, 16
         GoTo endProc
      Case Else
         .Text = TitleCase(.Text)
      End Select
      Cancel = Not oDriver.ValidateField(Index)
      
      If Index = 3 Then oDriver.LookupQuery(Index) = lsOldQuery
   End With

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index _
                       & ", " & Cancel & " )", True
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtOthers_GotFocus(Index As Integer)
   With txtOthers(Index)
      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = oApp.getColor("HT1")
   End With
   
   pbtxtOthers = True
   pnIndex = Index
End Sub

Private Sub txtOthers_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lsSearch() As String
   Dim lnCtr As Integer
   Dim lsSQL As String
   Dim lsOldProc As String
   
   lsOldProc = "txtOthers_KeyDown"
   '''On Error GoTo errProc

   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
      With txtOthers(Index)
         Select Case Index
         Case 0, 1
            If KeyCode = vbKeyF3 Then
               SearchOthers Index, .Text, False
               If .Text <> "" Then SetNextFocus
            Else
               If .Text <> "" Then
                  If Trim(LCase(.Text)) <> Trim(LCase(.Tag)) Then SearchOthers Index, .Text, False
               End If
            End If
            .Tag = .Text
         Case 12, 13
            Call txtOthers_Validate(Index, False)
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

Private Sub txtOthers_LostFocus(Index As Integer)
   With txtOthers(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyReturn, vbKeyUp, vbKeyDown
      Select Case KeyCode
      Case vbKeyReturn, vbKeyDown
         SetNextFocus
      Case vbKeyUp
         SetPreviousFocus
      End Select
   Case vbKeyF9
      If Not IsNumeric(txtField(11).Text) Then
         txtField(11).Text = Format(Code2Price(txtField(11).Text, psPriceCode), "#,##0.00")
         txtField(12).Text = Format(Code2Price(txtField(12).Text, psPriceCode), "#,##0.00")
         txtField(13).Text = Format(Code2Price(txtField(13).Text, psPriceCode), "#,##0.00")
         txtField(14).Text = Format(Code2Price(txtField(14).Text, psPriceCode), "#,##0.00")
      Else
         txtField(11).Text = Price2Code(txtField(11).Text, psPriceCode)
         txtField(12).Text = Price2Code(txtField(12).Text, psPriceCode)
         txtField(13).Text = Price2Code(txtField(13).Text, psPriceCode)
         txtField(14).Text = Price2Code(txtField(14).Text, psPriceCode)
      End If
   End Select
End Sub

Private Sub SearchOthers(ByVal lnIndex As Integer, _
                         ByVal lsValue As String, _
                         ByVal lbByCode As Boolean)
                         
   Dim lsSQL As String
   Dim lsBrowse As String
   Dim lrs As ADODB.Recordset
   Dim lsOldProc As String
   Dim lsSelected() As String
   Dim lsFieldIDx As String
   Dim lsFieldNmx As String
   
   lsOldProc = "SearchOthers"
   '''On Error GoTo errProc
   
   Select Case lnIndex
   Case 0
      lsSQL = "SELECT" _
                  & "  sSectnIDx" _
                  & ", sSectnNme" _
               & " FROM Section" _
               & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
                  & IIf(lbByCode, "AND sSectnIDx = " & strParm(lsValue), "AND sSectnNme LIKE " & strParm(lsValue & "%")) _
               & " ORDER BY sSectnNme"
      lsFieldIDx = "sSectnIDx"
      lsFieldNmx = "sSectnNme"
   Case 1
      lsSQL = "SELECT" _
                  & "  sBinIDxxx" _
                  & ", sBinNamex" _
               & " FROM Bin" _
               & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
                  & IIf(lbByCode, "AND sBinIDxxx = " & strParm(lsValue), "AND sBinNamex LIKE " & strParm(lsValue & "%")) _
               & " ORDER BY sBinNamex"
      lsFieldIDx = "sLevelIDx"
      lsFieldNmx = "sBinNamex"
   Case 14
'      lsSQL = "SELECT" _
'                  & "  sModelIdx" _
'                  & ", sModelCde" _
'                  & ", sModelNme" _
'               & " FROM CP_Model" _
'               & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
'                  & IIf(lbByCode, "AND sModelIdx = " & strParm(lsValue), "AND sModelNme LIKE " & strParm(lsValue & "%")) _
'               & " ORDER BY sModelNme"
               
      lsSQL = "SELECT" _
                  & "  a.sModelIDx" _
                  & ", a.sModelCde" _
                  & ", a.sModelNme" _
               & " FROM CP_Model a" _
                  & ", CP_Inventory b" _
               & " WHERE a.cRecdStat = " & strParm(xeRecStateActive) _
                  & " AND a.sModelIDx = b.sModelIDx " _
                  & IIf(lbByCode, "AND b.sStockIDx = " & strParm(lsValue), "AND a.sModelNme LIKE " & strParm(lsValue & "%")) _
               & " ORDER BY a.sModelNme"
      
      lsFieldIDx = "sModelIDx"
      lsFieldNmx = "sModelNme"
   End Select
   
   Set lrs = New ADODB.Recordset
   lrs.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText

   With txtOthers(lnIndex)
      If lrs.RecordCount = 1 Then
         If lnIndex <> 14 Then oRS(lsFieldIDx) = lrs(0)
         .Text = IFNull(lrs(1))
      ElseIf lrs.RecordCount > 1 Then
         Select Case lnIndex
         Case 1
            lsBrowse = KwikBrowse(oApp, lrs _
                        , "sBinIDxxx" & "»" & lsFieldNmx _
                        , "Code»Description")
         Case 14
            lsBrowse = KwikBrowse(oApp, lrs _
                        , "sModelIDx»sModelCde»sModelNme" _
                        , "ID»Code»Description")
         Case Else
            lsBrowse = KwikBrowse(oApp, lrs _
                        , lsFieldIDx & "»" & lsFieldNmx _
                        , "Code»Description")
         End Select
         
         If lsBrowse <> "" Then
            lsSelected = Split(lsBrowse, "»")
            If lnIndex <> 14 Then oRS(lsFieldIDx) = lsSelected(0)
            .Text = IFNull(lsSelected(1))
         Else
            If lnIndex <> 14 Then If oRS(lsFieldIDx) <> "" Then .Text = .Tag
         End If
      Else
         .Text = ""
         If lnIndex <> 14 Then oRS(lsFieldIDx) = ""
      End If
      
      .Tag = .Text
   End With

endProc:
   Set lrs = Nothing
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & lsFieldIDx _
                       & ", " & lsFieldNmx _
                       & ", " & lnIndex _
                       & ", " & lsValue & " )"
End Sub

Private Sub oDriver_Delete(Deleted As Boolean)
   Deleted = True
End Sub

Private Sub oDriver_DeleteComplete()
   For pnCtr = 0 To txtOthers.Count - 1
      txtOthers(pnCtr).Text = ""
   Next
   
   chkHsSerial.Value = 0
End Sub

Private Sub txtOthers_Validate(Index As Integer, Cancel As Boolean)
   Dim lsOldProc As String
   Dim lnCtr As Integer
   
   lsOldProc = "txtOthers_Validate"
   '''On Error GoTo errProc
   
   With txtOthers(Index)
      
      Select Case Index
      Case 0, 1
         If .Text <> "" Then
            If Trim(LCase(.Text)) <> Trim(LCase(.Tag)) Then SearchOthers Index, .Text, False
         End If

         .Tag = .Text
      Case 2 To 7, 9, 10, 11
         If Not IsNumeric(.Text) Then .Text = 0
         .Text = Format(.Text, "#,##0")
         oRS(Index) = CDbl(.Text)
      Case 8
         If Not IsDate(.Text) Then .Text = oApp.ServerDate
         .Text = Format(.Text, "MMMM DD, YYYY")
         oRS(Index) = .Text
      Case 12, 13
         If Trim(.Text) = "" Then
            If oRS.EOF Then Exit Sub
            InitOthers
            txtOthers(IIf(Index = 12, 13, 12)).Text = ""
            txtOthers(IIf(Index = 12, 13, 12)).Tag = ""
            
            For lnCtr = 0 To txtField.Count - 1
               Select Case lnCtr
               Case 0
               Case 11, 12
                  txtField(lnCtr).Text = "0.00"
               Case 13, 14, 15
                  txtField(lnCtr).Text = 0
               Case Else
                  txtField(lnCtr).Text = ""
                  txtField(lnCtr).Tag = txtField(lnCtr).Text
               End Select
            Next
            Exit Sub
         End If
         
         If .Tag <> .Text Then SearchBarCode .Text, IIf(Index = 12, True, False)
         .Tag = .Text
      End Select
   End With

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Cancel & " )", True
End Sub

Function NewBarrCode() As String
   Dim lrs As Recordset
   Dim lsSQL As String
   Dim lnCtr As Long
   Dim lsOldProc As String
   Dim lrsBranch As Recordset
   Dim lsCode As String
   
   lsOldProc = "NewBarrCode"
   '''On Error GoTo errProc
   
   Set lrsBranch = New ADODB.Recordset
   lrsBranch.Open "SELECT" _
                     & "  a.sCompnyCd" _
                  & " FROM Company a" _
                     & ", Branch b" _
                  & " WHERE a.sCompnyID = b.sCompnyID" _
                     & " AND b.sBranchCd = " & strParm(oApp.BranchCode) _
                  , oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
                   
   lsCode = "GMC"
   If Not lrsBranch.EOF Then lsCode = lrsBranch("sCompnyCd")

   lsSQL = "SELECT" & _
               " sBarrCode" & _
            " FROM CP_Inventory" & _
            " WHERE sBarrCode LIKE " & strParm(Format(Date, "yy") & "-" & lsCode & "-%") & _
            " ORDER BY sBarrCode DESC" & _
            " LIMIT 1"
   Set lrs = New Recordset
   lrs.Open lsSQL, oApp.Connection, , , adCmdText
   
   If lrs.EOF Then
      lnCtr = 1
   Else
      If Left(lrs("sBarrCode"), 2) = Format(Date, "yy") Then
         lnCtr = CLng(Right(lrs("sBarrCode"), 6)) + 1
      Else
         lnCtr = 1
      End If
   End If
   NewBarrCode = Format(Date, "yy") & "-" & lsCode & "-" & Format(lnCtr, "000000")
   
   Set lrs = Nothing
   Set lrsBranch = Nothing

endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"
End Function

Private Sub SearchBarCode(ByVal lsValue As String, ByVal lbByCode As Boolean)
   Dim lrsCellphone As ADODB.Recordset
   Dim lrsInventoryx As ADODB.Recordset
   Dim lsSelected() As String
   Dim lsBrowse As String
   Dim lsOldProc As String
   Dim lsStockIDx As String
   Dim lsSQL As String
   
   lsOldProc = "SearchSpareParts"
   '''On Error GoTo errProc
   
   lsSQL = "SELECT" _
               & "  a.sStockIDx" _
               & ", a.sBarrcode" _
               & ", " & psConcatDescx & " xDescript" _
               & ", b.sBrandNme" _
               & ", c.sModelCde" _
               & ", c.sModelNme" _
               & ", d.sColorNme" _
               & ", f.sSizeName" _
               & ", e.sCategrNm" _
               & ", g.sCategrNm sSubCtgID "
      lsSQL = lsSQL _
            & " FROM CP_Inventory a" _
               & " LEFT JOIN CP_Brand b" _
                  & " ON a.sBrandIDx = b.sBrandIDx" _
               & " LEFT JOIN CP_Model c" _
                  & " ON a.sModelIDx = c.sModelIDx" _
               & " LEFT JOIN Color d" _
                  & " ON a.sColorIDx = d.sColorIDx" _
               & " LEFT JOIN Category e" _
                  & " ON a.sCategID1 = e.sCategrID" _
               & " LEFT JOIN Category g" _
                  & " ON a.sCategID1 = g.sCategrID" _
               & " LEFT JOIN Size f" _
                  & " ON a.sSizeIDxx = f.sSizeIDxx" _
            & " ORDER BY a.sBarrCode" _
               & ", xDescript"
                  
   lsSQL = AddCondition(lsSQL, IIf(lbByCode, "a.sBarrcode LIKE " & strParm(lsValue & "%") _
                        , psConcatDescx & " LIKE " & strParm(lsValue & "%")))
   Debug.Print lsSQL
   Set lrsCellphone = New ADODB.Recordset
   With lrsCellphone
      .Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
      If .EOF Then
         oDriver_InitValue
         For pnCtr = 0 To txtField.Count - 1
            txtField(pnCtr).Text = Empty
         Next
         GoTo endProc
      End If
      
      If .RecordCount = 1 Then
         lsStockIDx = .Fields("sStockIDx")
         GoTo LoadRecord
      Else
         With txtOthers(pnIndex)
            .BackColor = oApp.getColor("EB")
            lsBrowse = KwikBrowse(oApp, lrsCellphone _
                                    , "sBarrcode»xDescript»sBrandNme»sModelCde»sModelNme»sColorNme»sSizeName»sCategrNm" _
                                    , "BarrCode»Description»Brand»Model Code»Model Name»Color»Size»Category" _
                                    , "@»@»@»@»@»@»@»@" _
                                    , "a.sBarrcode»xDescript»b.sBrandNme»c.sModelCde»c.sModelNme»d.sColorNme»f.sSizeName»e.sCategrNm")
                                    
            If lsBrowse <> "" Then
               lsSelected = Split(lsBrowse, "»")
               lsStockIDx = lsSelected(0)
               GoTo LoadRecord
            End If
            .BackColor = oApp.getColor("HT1")
            .SelStart = 0
            .SelLength = Len(.Text)
         End With
      End If
      GoTo endProc
   End With
      
LoadRecord:
   lsSQL = "SELECT" _
               & "  a.sStockIDx" _
               & ", a.sBranchCd" _
               & ", a.cRecdStat" _
               & ", b.sBarrcode" _
               & ", b.cHsSerial" _
            & " FROM CP_Inventory_Master a" _
               & ", CP_Inventory b" _
            & " WHERE a.sStockIDx = b.sStockIDx" _
            & " ORDER BY a.sBranchCd DESC"
   lsSQL = AddCondition(lsSQL, "a.sStockIDx = " & strParm(lsStockIDx))
   
   Set lrsInventoryx = New ADODB.Recordset
   
   pbEnblButtons = True
   pbNewInvntory = False
   With lrsInventoryx
      .Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
      If Not .EOF Then
         .Find "sBranchCd = " & strParm(oApp.BranchCode), 0, adSearchForward
        
         If Not lrsInventoryx.EOF Then

            If .Fields("cRecdStat") = xeRecStateActive Then
               oDriver.LookupValue(0) = .Fields("sBarrcode")
               oDriver.LoadRecord
            
               pbEnblButtons = False
            Else
               MsgBox "CP Inventory Status is Deactivated!!!" & vbCrLf & _
                        "Please Save the record to activate!!!", vbInformation, "Notice"
               
               oDriver.LookupValue(0) = .Fields("sBarrcode")
               oDriver.LoadRecord
               oDriver.RecordUpdate
               
               txtOthers(0).SetFocus
               oRS.Fields("nBegQtyxx") = 0
               oRS.Fields("nQtyOnHnd") = 0
               'ask mac, para saan? wala pa sa form 2013/03/07 `she`
'               oRS.Fields("nWrtQtyxx") = 0
'               oRS.Fields("nWrntyUnt") = 0
               oRS.Fields("nDmoQtyxx") = 0
               oRS.Fields("nDemoUnit") = 0
               oRS.Fields("cRecdStat") = xeRecStateActive
               
               txtOthers(2).Text = 0
               txtOthers(3).Text = 0
            End If
         Else
            MsgBox "No Inventory found in your warehouse!!!" & vbCrLf & _
                     "Plese Save the record to create!!!", vbInformation, "Notice"
            
            .MoveFirst

            oDriver.LookupValue(0) = .Fields("sBarrcode")
            oDriver.LoadRecord
            oDriver.RecordUpdate
            txtOthers(0).SetFocus
   
            pbNewInvntory = True
         End If
      Else
         'no record at all
         pbNewInvntory = True
         If Not lrsCellphone.EOF Then
            MsgBox "No Inventory found in your warehouse!!!" & vbCrLf & _
                     "Plese Save the record to create!!!", vbInformation, "Notice"
            
            oDriver.LookupValue(0) = lrsCellphone("sBarrcode")
            oDriver.LoadRecord
            oDriver.RecordUpdate
            txtOthers(0).SetFocus
            
            pbNewInvntory = True
         End If
      End If
      .Close
   End With
endProc:
   Set lrsCellphone = Nothing
   Set lrsInventoryx = Nothing
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & oDriver.FieldValue(2) & "»" & oDriver.FieldValue(3) & " )", True
End Sub

Private Function UpdateCPInventory() As Boolean
   Dim lsOldProc As String
   Dim lrs As Recordset
   Dim lrs1 As Recordset
   Dim lsSQL As String
   Dim lnRow As Integer
   
   lsOldProc = "UpdateCPInventory"
   '''On Error GoTo errProc

   lsSQL = AddCondition(psCPInventory, "sStockIDx = " & strParm(oDriver.FieldValue(22)) _
                  & " AND sBranchCd = " & strParm(oApp.BranchCode))

   Set lrs = New Recordset
   lrs.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   
   If lrs.EOF Then
      lsSQL = ADO2SQL(oRS, "CP_Inventory_Master", "", oApp.UserID, oApp.ServerDate, "")
   Else
      lsSQL = ADO2SQL(oRS, "CP_Inventory_Master", "sStockIDx = " & strParm(oDriver.FieldValue(22)) & " AND sBranchCd = " & strParm(oApp.BranchCode), oApp.UserID, oApp.ServerDate, "")
   End If
   
   If lsSQL <> "" Then
      lnRow = oApp.Execute(lsSQL, "CP_Inventory_Master", oApp.BranchCode, "")
      If lnRow <= 0 Then
         MsgBox "Unable to Save Inventory" & vbCrLf & _
                  lsSQL, vbCritical, "Warning"
         GoTo endProc
      End If
   End If
   Set lrs1 = New Recordset
   If oDriver.FieldValue(6) <> "" Then
      lrs1.Open "SELECT * FROM Category WHERE sCategrID = " & strParm(oDriver.FieldValue(7)), oApp.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
      If Not lrs1.EOF Then oDriver.FieldValue(21) = lrs1("cSerialze")
   Else
'      MsgBox "Invalid Model Name Detected!!!" & vbCrLf & _
'               "Please verify your entry then try again!!!", vbCritical, "WARNING"
'      txtField(3).SetFocus
'      GoTo endProc
   End If
   
   UpdateCPInventory = True

endProc:
   Set lrs = Nothing
   Set lrs1 = Nothing
   Exit Function
errProc:
   ShowError lsOldProc & "( " & UpdateCPInventory & " )", True
End Function

Private Function SaveCPInventoryLedger() As Boolean
   Dim lsSQL As String
   Dim lnRow As Long
   Dim lsOldProc As String

   lsOldProc = "SaveSPInventoryLedger"
   '''On Error GoTo errProc

   lsSQL = "INSERT INTO CP_Inventory_Ledger SET" _
               & "  sStockIDx = " & strParm(oDriver.FieldValue(22)) _
               & ", sBranchCd = " & strParm(oApp.BranchCode) _
               & ", sSourceCd = 'CPAd'" _
               & ", sSourceNo = '9900000001'" _
               & ", nQtyInxxx = " & CLng(txtOthers(3).Text) _
               & ", nQtyOutxx = '0'" _
               & ", nQtyOrder = '0'" _
               & ", nQtyIssue = '0'" _
               & ", nLedgerNo = '000001'" _
               & ", nQtyOnHnd = '0'" _
               & ", cUnitType = '1'" _
               & ", dTransact = " & dateParm(oApp.ServerDate) _
               & ", dModified = " & dateParm(oApp.ServerDate)
   
   lnRow = oApp.Execute(lsSQL, "CP_Inventory_Ledger", oApp.BranchCode)
   If lnRow <= 0 Then
      MsgBox "Unable to Save CP_Inventory_Ledger!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   SaveCPInventoryLedger = True

endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"
End Function

Private Sub ComputePricing()
   Dim lnPurPrice As Double

   lnPurPrice = oDriver.FieldValue(17)
                                                                          
   If lnPurPrice >= CDbl(txtField(11).Text) Then
      txtField(11).Text = "0.00"
      
      oDriver.FieldValue(11) = txtField(11).Text
      Exit Sub
   End If
                                                                                                   
   If lnPurPrice > Round(CDbl(txtField(11).Text) * (100 - oDriver.FieldValue(11)) / 100, 2) Then
      txtOthers(11).Text = "0.00"
      Exit Sub
   End If
                                                                                                   
'   If lnPurPrice > Round(CDbl(txtOthers(10).Text) * (100 - CDbl(txtOthers(12).Text)) / 100, 2) Then
'      txtOthers(12).Text = "0.00"
'      Exit Sub
'   End If
'
'   If CDbl(txtOthers(11).Text) > CDbl(txtOthers(12).Text) Then txtOthers(12).Text = "0.00"
End Sub

Private Function getModelCode(ByVal lsModelID As String) As String
   Dim lors As Recordset
   Dim lsSQL As String
   
   lsSQL = "SELECT sModelCde FROM CP_Model WHERE sModelIDx = " & strParm(lsModelID)
   Set lors = New Recordset
   
   lors.Open lsSQL, oApp.Connection, , , adCmdText
   Set lors.ActiveConnection = Nothing
   
   If lors.EOF Or lors.RecordCount <> 1 Then GoTo endProc
   
   getModelCode = IFNull(lors(0))
endProc:
   Set lors = Nothing
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
