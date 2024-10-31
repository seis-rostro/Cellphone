VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmCP_Inventory_Branch 
   BorderStyle     =   0  'None
   Caption         =   "Cellphone Maintenance"
   ClientHeight    =   6210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11940
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6210
   ScaleWidth      =   11940
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame2 
      Height          =   4110
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   1065
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   7250
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   10
         Left            =   1560
         TabIndex        =   27
         Top             =   3135
         Width           =   1890
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   9
         Left            =   1560
         TabIndex        =   25
         Top             =   2835
         Width           =   1890
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   8
         Left            =   1560
         TabIndex        =   23
         Top             =   2535
         Width           =   1890
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   7
         Left            =   1560
         TabIndex        =   21
         Top             =   2235
         Width           =   1890
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   16
         Left            =   1095
         TabIndex        =   35
         Top             =   3570
         Width           =   2820
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   14
         Left            =   4470
         TabIndex        =   31
         Top             =   1905
         Width           =   1245
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   15
         Left            =   4470
         TabIndex        =   33
         Top             =   2205
         Width           =   1245
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   13
         Left            =   4470
         TabIndex        =   29
         Top             =   1605
         Width           =   1245
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   12
         Left            =   10065
         TabIndex        =   43
         Top             =   1110
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
         Index           =   9
         Left            =   10065
         TabIndex        =   59
         Top             =   2730
         Width           =   1425
      End
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   6
         Left            =   7290
         TabIndex        =   53
         Top             =   2730
         Width           =   1350
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   1095
         TabIndex        =   7
         Top             =   240
         Width           =   2820
      End
      Begin VB.CheckBox chkHsSerial 
         Caption         =   "w/ Serial"
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
         Left            =   4575
         TabIndex        =   36
         Tag             =   "wt0;fb0"
         Top             =   3525
         Width           =   1095
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   11
         Left            =   9240
         TabIndex        =   61
         Text            =   "0,000.00"
         Top             =   3315
         Width           =   2250
      End
      Begin VB.TextBox txtOthers 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   8
         Left            =   10065
         TabIndex        =   57
         Top             =   2430
         Width           =   1425
      End
      Begin VB.TextBox txtOthers 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   7290
         TabIndex        =   51
         Top             =   2430
         Width           =   1350
      End
      Begin VB.TextBox txtOthers 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   7
         Left            =   10065
         TabIndex        =   55
         Top             =   2130
         Width           =   1425
      End
      Begin VB.TextBox txtOthers 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   7290
         TabIndex        =   49
         Top             =   2130
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
         Height          =   420
         Index           =   3
         Left            =   10065
         TabIndex        =   47
         Text            =   "0,000"
         Top             =   1545
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
         Height          =   420
         Index           =   2
         Left            =   7290
         TabIndex        =   45
         Top             =   1545
         Width           =   1350
      End
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   6870
         TabIndex        =   41
         Top             =   810
         Width           =   4620
      End
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   6870
         TabIndex        =   39
         Top             =   510
         Width           =   4620
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   6
         Left            =   1095
         TabIndex        =   19
         Top             =   1785
         Width           =   1890
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   3825
         TabIndex        =   17
         Top             =   1305
         Width           =   1890
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   3825
         TabIndex        =   15
         Top             =   1005
         Width           =   1890
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   1095
         TabIndex        =   13
         Top             =   1305
         Width           =   1890
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   1095
         TabIndex        =   11
         Top             =   1005
         Width           =   1890
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   1095
         TabIndex        =   9
         Top             =   705
         Width           =   4620
      End
      Begin VB.Line Line2 
         X1              =   1320
         X2              =   1320
         Y1              =   2055
         Y2              =   3315
      End
      Begin VB.Line Line3 
         Index           =   0
         X1              =   1320
         X2              =   2190
         Y1              =   2685
         Y2              =   2685
      End
      Begin VB.Line Line3 
         Index           =   1
         X1              =   1320
         X2              =   2190
         Y1              =   2970
         Y2              =   2970
      End
      Begin VB.Line Line3 
         Index           =   2
         X1              =   1320
         X2              =   2190
         Y1              =   2385
         Y2              =   2385
      End
      Begin VB.Line Line3 
         Index           =   3
         X1              =   1320
         X2              =   2190
         Y1              =   3300
         Y2              =   3300
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cat 04"
         Height          =   195
         Index           =   26
         Left            =   810
         TabIndex        =   26
         Top             =   3195
         Width           =   465
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cat 03"
         Height          =   195
         Index           =   22
         Left            =   810
         TabIndex        =   24
         Top             =   2880
         Width           =   465
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cat 02"
         Height          =   195
         Index           =   19
         Left            =   810
         TabIndex        =   22
         Top             =   2580
         Width           =   465
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cat 01"
         Height          =   195
         Index           =   18
         Left            =   810
         TabIndex        =   20
         Top             =   2265
         Width           =   465
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Part Nmbr"
         Height          =   195
         Index           =   17
         Left            =   210
         TabIndex        =   34
         Top             =   3615
         Width           =   705
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reserve Order"
         Height          =   195
         Index           =   16
         Left            =   8985
         TabIndex        =   58
         Top             =   2790
         Width           =   1035
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Max. Level"
         Height          =   195
         Index           =   25
         Left            =   6060
         TabIndex        =   50
         Top             =   2475
         Width           =   780
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Back Order"
         Height          =   195
         Index           =   14
         Left            =   9210
         TabIndex        =   56
         Top             =   2475
         Width           =   810
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Discount 3"
         Height          =   195
         Index           =   11
         Left            =   3570
         TabIndex        =   32
         Top             =   2250
         Width           =   765
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Discount 2"
         Height          =   195
         Index           =   10
         Left            =   3570
         TabIndex        =   30
         Top             =   1935
         Width           =   765
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Discount 1"
         Height          =   195
         Index           =   9
         Left            =   3570
         TabIndex        =   28
         Top             =   1635
         Width           =   765
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Level"
         Height          =   195
         Index           =   3
         Left            =   6015
         TabIndex        =   40
         Top             =   855
         Width           =   390
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Section"
         Height          =   195
         Index           =   2
         Left            =   6015
         TabIndex        =   38
         Top             =   540
         Width           =   540
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Purc. Price"
         Height          =   195
         Index           =   12
         Left            =   8865
         TabIndex        =   42
         Top             =   1155
         Width           =   1125
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Branch Detail Info"
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
         Index           =   15
         Left            =   6015
         TabIndex        =   37
         Top             =   195
         Width           =   1560
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Beg. Inv. Date"
         Height          =   195
         Index           =   13
         Left            =   6060
         TabIndex        =   52
         Top             =   2790
         Width           =   1035
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Selling Price"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   27
         Left            =   7290
         TabIndex        =   60
         Top             =   3420
         Width           =   1785
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   6000
         X2              =   11490
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   6000
         X2              =   11490
         Y1              =   1455
         Y2              =   1455
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Min. Level"
         Height          =   195
         Index           =   24
         Left            =   6060
         TabIndex        =   48
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ReOrder Level"
         Height          =   210
         Index           =   23
         Left            =   8970
         TabIndex        =   54
         Top             =   2160
         Width           =   1035
      End
      Begin VB.Label lblField 
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
         TabIndex        =   46
         Top             =   1635
         Width           =   1380
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Beg. Balance"
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
         Left            =   6060
         TabIndex        =   44
         Top             =   1620
         Width           =   1215
      End
      Begin VB.Shape Shape4 
         Height          =   750
         Left            =   5895
         Top             =   3210
         Width           =   5715
      End
      Begin VB.Shape Shape3 
         Height          =   3045
         Left            =   5895
         Top             =   120
         Width           =   5715
      End
      Begin VB.Shape Shape2 
         Height          =   3840
         Left            =   105
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
         TabIndex        =   18
         Top             =   1800
         Width           =   630
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Color"
         Height          =   195
         Index           =   7
         Left            =   3135
         TabIndex        =   16
         Top             =   1305
         Width           =   360
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Size"
         Height          =   195
         Index           =   6
         Left            =   3135
         TabIndex        =   14
         Top             =   1020
         Width           =   300
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
         Height          =   195
         Index           =   5
         Left            =   195
         TabIndex        =   12
         Top             =   1305
         Width           =   435
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Brand"
         Height          =   315
         Index           =   4
         Left            =   195
         TabIndex        =   10
         Top             =   1020
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
         Caption         =   "Barr Code"
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   6
         Top             =   270
         Width           =   705
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   1
         Left            =   195
         TabIndex        =   8
         Top             =   720
         Width           =   795
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   0
      Left            =   11085
      TabIndex        =   69
      Top             =   5400
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
      Picture         =   "frmCP_Inventory_Branch.frx":0000
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   1
      Left            =   10305
      TabIndex        =   68
      Top             =   5400
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
      Picture         =   "frmCP_Inventory_Branch.frx":077A
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   2
      Left            =   7185
      TabIndex        =   62
      Top             =   5400
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
      Picture         =   "frmCP_Inventory_Branch.frx":0EF4
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   4
      Left            =   7185
      TabIndex        =   64
      Top             =   5400
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
      Picture         =   "frmCP_Inventory_Branch.frx":166E
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   5
      Left            =   11085
      TabIndex        =   70
      Top             =   5400
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
      Picture         =   "frmCP_Inventory_Branch.frx":1DE8
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   705
      Index           =   7
      Left            =   7965
      TabIndex        =   63
      Top             =   5400
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
      Picture         =   "frmCP_Inventory_Branch.frx":2562
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   3
      Left            =   7965
      TabIndex        =   65
      Top             =   5400
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
      Picture         =   "frmCP_Inventory_Branch.frx":2CDC
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   6
      Left            =   7965
      TabIndex        =   72
      Top             =   5400
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
      Picture         =   "frmCP_Inventory_Branch.frx":3456
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   8
      Left            =   9525
      TabIndex        =   67
      Top             =   5400
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
      Picture         =   "frmCP_Inventory_Branch.frx":3BD0
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   9
      Left            =   7965
      TabIndex        =   71
      Top             =   5400
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
      Picture         =   "frmCP_Inventory_Branch.frx":434A
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
         Height          =   300
         Index           =   12
         Left            =   1095
         TabIndex        =   1
         Top             =   90
         Width           =   3345
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
         Height          =   300
         Index           =   11
         Left            =   8670
         TabIndex        =   5
         Top             =   90
         Width           =   2940
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
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   10
         Left            =   5415
         MaxLength       =   50
         TabIndex        =   3
         Top             =   90
         Width           =   2205
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
         Height          =   300
         Index           =   1
         Left            =   150
         TabIndex        =   0
         Top             =   135
         Width           =   960
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
         Left            =   7680
         TabIndex        =   4
         Top             =   135
         Width           =   960
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Barr C&ode"
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
         Left            =   4515
         TabIndex        =   2
         Top             =   135
         Width           =   885
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   10
      Left            =   8745
      TabIndex        =   66
      Top             =   5400
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
      Picture         =   "frmCP_Inventory_Branch.frx":4AC4
      PicturePos      =   1
   End
End
Attribute VB_Name = "frmCP_Inventory_Branch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Private Const pxeMODULENAME = "frmCP_Inventory"
'
'Private WithEvents oDriver As clsFormDriver
'Private oFormSerialNew As frmCPSerial
'Private oSkin As clsFormSkin
'Private oBranch As clsBranch
'Private bLoaded As Boolean
'Private oRS As New ADODB.Recordset
'
'Dim pbtxtOthers As Boolean
'Dim psBranchCd As String
'Dim pnCtr As Integer, pnIndex As Integer
'
'Dim psCPInventory As String
'Dim pbEnblButtons As Boolean
'Dim pbNewInvntory As Boolean
'Dim psPriceCode As String
'Dim pnCmdClosehwd As Long
'
'Private Sub cmdButton_Click(Index As Integer)
'   Dim lsSearch As String
'   Dim lnRep As Integer
'   Dim lsOldProc As String
'
'   lsOldProc = "cmdButton_Click"
'   'On Error GoTo errProc
'
'   If Index = 2 Then
'      If pbtxtOthers Then
'         Call txtOthers_Validate(pnIndex, False)
'      Else
'         Call txtField_Validate(pnIndex, False)
'      End If
'   End If
'
'   Select Case Index
'   Case 0 'cancel
'      oDriver.RecordCancelUpdate
'      pbEnblButtons = False
'   Case 1 'browse
'      oDriver.BrowseRecord
'   Case 2 'save
'      oDriver.RecordSave
'   Case 3 'update
'      If Not IsNumeric(txtField(11).Text) Then txtField(11).Text = Format(Code2Price(txtField(11).Text, psPriceCode), "#,##0.00")
'      oDriver.RecordUpdate
'   Case 4 'new
'      oDriver.RecordNew
'   Case 5 'close
'      Unload Me
'   Case 6 'delete
'      oDriver.RecordDelete
'   Case 7 'search
'      If pbtxtOthers Then
'         oDriver.RecordSearch
'         txtField(pnIndex).SetFocus
'      Else
'         SearchOthers pnIndex, Empty, False
'         txtOthers(pnIndex).SetFocus
'      End If
'   Case 8 'ledger
'      If Not pbNewInvntory Then
'         With frmCP_InventoryLedger
'            .txtField(0) = txtField(0)
'            .txtField(1) = txtField(1)
'            .txtField(2) = txtField(2)
'            .txtField(3) = txtField(3)
'
'            .StockID = oDriver.FieldValue(18)
'            .Show 1
'         End With
'      Else
'         MsgBox "Unable to Load Inventory Ledger!!!" & vbCrLf & _
'                "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
'      End If
'   Case 10
'      If oDriver.FieldValue(17) = xeYes Then
'         With oFormSerialNew
'            .StockID = oDriver.FieldValue(18)
'            .Barcode = oDriver.FieldValue(0)
'            .Description = oDriver.FieldValue(1)
'            .Brand = txtField(2).Text
'            .Model = txtField(3).Text
'            .Color = txtField(5).Text
'            .Category = txtField(6).Text
'            .Branch = psBranchCd
'
'            .Show 1
'         End With
'      End If
'   End Select
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " & Index & " )", True
'End Sub
'
'Private Sub Form_Activate()
'   Dim lsOldProc As String
'
'   lsOldProc = "Form_Activate"
'   'On Error GoTo errProc
'
'   oApp.MenuName = Me.Tag
'   Me.ZOrder 0
'
'   If bLoaded = False Then
'      oDriver.RecordCancelUpdate
'      oDriver_InitValue
'      bLoaded = True
'      txtOthers(12).SetFocus
'   End If
'   mdiMain.StatusBar1.Panels(1).Text = "Press F9 to encrypt selling price!!!"
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " & " )", True
'End Sub
'
'Private Sub Form_Load()
'   Dim lsSQL As String
'   Dim lsOldProc As String
'
'   lsOldProc = "Form_Load"
'   'On Error GoTo errProc
'
'   CenterChildForm mdiMain, Me
'
'   bLoaded = False
'
'   Set oFormSerialNew = New frmCPSerial
'
'   Set oRS = New ADODB.Recordset
'
'   Set oDriver = New clsFormDriver
'   Set oDriver.AppDriver = oApp
'   Set oDriver.MainForm = Me
'
'   Set oSkin = New clsFormSkin
'   Set oSkin.AppDriver = oApp
'   Set oSkin.Form = Me
'   oSkin.ApplySkin
'
'   Set oBranch = New clsBranch
'   Set oBranch.AppDriver = oApp
'   oBranch.Filter = "sBranchCd <> " & strParm(oApp.BranchCode)
'   oBranch.InitRecord
'   oBranch.NewRecord
'
'   oDriver.RecQuery = "SELECT" _
'                           & "  sBarrcode" _
'                           & ", sDescript" _
'                           & ", sBrandIDx" _
'                           & ", sModelIDx" _
'                           & ", sMadeIDxx" _
'                           & ", sColorIDx" _
'                           & ", sCategID1" _
'                           & ", sCategID2" _
'                           & ", sCategID3" _
'                           & ", sCategID4" _
'                           & ", sCategID5" _
'                           & ", nSelPrice" _
'                           & ", nLastPrce" _
'                           & ", nMaxDisc1" _
'                           & ", nMaxDisc2" _
'                           & ", nMaxDisc3" _
'                           & ", sPartNoxx" _
'                           & ", cHsSerial" _
'                           & ", sStockIDx" _
'                           & ", cRecdStat" _
'                           & ", sModified" _
'                           & ", dModified" _
'                        & " FROM CP_Inventory"
'
'   oDriver.BrowseQuery = "SELECT" _
'                              & "  a.sBarrcode" _
'                              & ", a.sDescript" _
'                              & ", b.sBrandNme" _
'                              & ", c.sModelNme" _
'                              & ", e.sMadeName" _
'                              & ", d.sColorNme" _
'                              & ", f.sCategrNm" _
'                              & ", g.nQtyonHnd" _
'                           & " FROM CP_Inventory a" _
'                              & " LEFT JOIN CP_Brand b" _
'                                 & " ON a.sBrandIDx = b.sBrandIDx" _
'                              & " LEFT JOIN CP_Model c" _
'                                 & " ON a.sModelIDx = c.sModelIDx" _
'                              & " LEFT JOIN Color d" _
'                                 & " ON a.sColorIDx = d.sColorIDx" _
'                              & " LEFT JOIN Made e" _
'                                 & " ON a.sMadeIDxx = e.sMadeIDxx" _
'                              & " LEFT JOIN Category f" _
'                                 & " ON a.sCategID1 = f.sCategrID" _
'                              & ", CP_Inventory_Master g" _
'                           & " WHERE a.sStockIDx = g.sStockIDx" _
'                              & " AND g.sBranchCd = " & strParm(psBranchCd)
'   oDriver.InitRecForm
'
'   oDriver.BrowseColumn(0) = "sBarrcode"
'   oDriver.BrowseColumn(1) = "sDescript"
'   oDriver.BrowseColumn(2) = "sModelNme"
'   oDriver.BrowseColumn(3) = "nQtyonHnd"
'   oDriver.BrowseColumn(4) = "sColorNme"
'   oDriver.BrowseColumn(5) = "sMadeName"
'   oDriver.BrowseColumn(6) = "sCategrNm"
'
'   oDriver.BrowseFTitle(0) = "Barcode"
'   oDriver.BrowseFTitle(1) = "Description"
'   oDriver.BrowseFTitle(2) = "Model"
'   oDriver.BrowseFTitle(3) = "QOH"
'   oDriver.BrowseFTitle(4) = "Color"
'   oDriver.BrowseFTitle(5) = "Made"
'   oDriver.BrowseFTitle(6) = "Category"
'
'   oDriver.LookupQuery(2) = "SELECT" _
'                              & "  sBrandIDx" _
'                              & ", sBrandNme" _
'                           & " FROM CP_Brand" _
'                           & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
'                           & " ORDER BY sBrandNme"
'   oDriver.LookupReference(2) = "sBrandIDx»sBrandNme"
'   oDriver.LookupColumn(2) = "sBrandNme"
'   oDriver.LookupTitle(2) = "Brand Name"
'
'
'   oDriver.LookupQuery(3) = "SELECT" _
'                              & " sModelIDx" _
'                              & ",sModelNme " _
'                           & "FROM CP_Model " _
'                           & "WHERE cRecdStat = 1 " _
'                           & "ORDER BY sModelNme"
'   oDriver.LookupReference(3) = "sModelIDx»sModelNme"
'   oDriver.LookupColumn(3) = "sModelNme"
'   oDriver.LookupTitle(3) = "Model Name"
'
'   oDriver.LookupQuery(4) = "SELECT" _
'                                 & "  sMadeIDxx" _
'                                 & ", sMadeName" _
'                           & " FROM Made" _
'                           & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
'                           & " ORDER BY sMadeName"
'   oDriver.LookupReference(4) = "sMadeIDxx»sMadeName"
'   oDriver.LookupColumn(4) = "sMadeName"
'   oDriver.LookupTitle(4) = "Made Name"
'
'   oDriver.LookupQuery(5) = "SELECT" _
'                              & "  sColorIDx" _
'                              & ", sColorNme" _
'                           & " FROM Color" _
'                           & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
'                           & " ORDER BY sColorNme"
'   oDriver.LookupReference(5) = "sColorIDx»sColorNme"
'   oDriver.LookupColumn(5) = "sColorNme"
'   oDriver.LookupTitle(5) = "Color Name"
'
'   oDriver.LookupQuery(6) = "SELECT" _
'                              & "  sCategrID" _
'                              & ", sCategrNm" _
'                           & " FROM Category" _
'                           & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
'                           & " ORDER BY sCategrNm"
'   oDriver.LookupReference(6) = "sCategrID»sCategrNm"
'   oDriver.LookupColumn(6) = "sCategrNm"
'   oDriver.LookupTitle(6) = "Category Name"
'
'   oDriver.LookupQuery(7) = "SELECT" _
'                              & "  sCategrID" _
'                              & ", sCategrNm" _
'                           & " FROM Category" _
'                           & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
'                           & " ORDER BY sCategrNm"
'   oDriver.LookupReference(7) = "sCategrID»sCategrNm"
'   oDriver.LookupColumn(7) = "sCategrNm"
'   oDriver.LookupTitle(7) = "Category Name"
'
'   oDriver.LookupQuery(8) = "SELECT" _
'                              & "  sCategrID" _
'                              & ", sCategrNm" _
'                           & " FROM Category" _
'                           & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
'                           & " ORDER BY sCategrNm"
'   oDriver.LookupReference(8) = "sCategrID»sCategrNm"
'   oDriver.LookupColumn(8) = "sCategrNm"
'   oDriver.LookupTitle(8) = "Category Name"
'
'   oDriver.LookupQuery(9) = "SELECT" _
'                              & "  sCategrID" _
'                              & ", sCategrNm" _
'                           & " FROM Category" _
'                           & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
'                           & " ORDER BY sCategrNm"
'   oDriver.LookupReference(9) = "sCategrID»sCategrNm"
'   oDriver.LookupColumn(9) = "sCategrNm"
'   oDriver.LookupTitle(9) = "Category Name"
'
'   oDriver.LookupQuery(10) = "SELECT" _
'                              & "  sCategrID" _
'                              & ", sCategrNm" _
'                           & " FROM Category" _
'                           & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
'                           & " ORDER BY sCategrNm"
'   oDriver.LookupReference(10) = "sCategrID»sCategrNm"
'   oDriver.LookupColumn(10) = "sCategrNm"
'   oDriver.LookupTitle(10) = "Category Name"
'
'   psCPInventory = "SELECT" _
'                     & "  sSectnIDx" _
'                     & ", sBinIDxxx" _
'                     & ", nBegQtyxx" _
'                     & ", nQtyOnHnd" _
'                     & ", nMinLevel" _
'                     & ", nMaxLevel" _
'                     & ", dBegInvxx" _
'                     & ", nReorderx" _
'                     & ", nBackOrdr" _
'                     & ", nResvOrdr" _
'                     & ", nFloatQty" _
'                     & ", nLedgerNo" _
'                     & ", dLastTran" _
'                     & ", cRecdStat" _
'                     & ", sStockIDx" _
'                     & ", sBranchCd" _
'                     & ", sModified" _
'                     & ", dModified" _
'                  & " FROM CP_Inventory_Master" _
'                  & " ORDER BY sBranchCd"
'
'   Set oRS = New ADODB.Recordset
'   lsSQL = AddCondition(psCPInventory, "0 = 1")
'   oRS.Open lsSQL, oApp.Connection, adOpenStatic, adLockPessimistic, adCmdText
'   Set oRS.ActiveConnection = Nothing
'
'   oDriver.FieldStart = 0
'
'   oDriver.FieldFormat(0) = ">"
'   oDriver.FieldFormat(16) = ">"
'
'   oDriver.FieldFormat(11) = "#,##0.00"
'   oDriver.FieldFormat(12) = "#,##0.00"
'
'   txtOthers(12).Text = ""
'   txtOthers(12).Tag = ""
'
'   psPriceCode = "PATRONIZEX"
'   psBranchCd = oApp.BranchCode
'   pnCmdClosehwd = cmdButton(5).hwnd
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " & " )", True
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'   Set oDriver = Nothing
'   Set oSkin = Nothing
'   Set oRS = Nothing
'   Set oFormSerialNew = Nothing
'   Set oBranch = Nothing
'
'   mdiMain.StatusBar1.Panels(1).Text = ""
'End Sub
'
'Private Sub oDriver_DisableOtherControl()
'   For pnCtr = 0 To txtOthers.Count - 1
'      txtOthers(pnCtr).Enabled = False
'   Next
'
'   txtOthers(10).Enabled = True
'   txtOthers(11).Enabled = True
'   txtOthers(12).Enabled = True
'
'   chkHsSerial.Enabled = False
'
'   oDriver.hideButton 6
'   oDriver.hideButton 9
'End Sub
'
'Private Sub oDriver_EnableOtherControl()
'   For pnCtr = 0 To txtOthers.Count - 1
'      Select Case pnCtr
'      Case 10, 11, 12
'         txtOthers(pnCtr).Enabled = False
'      Case Else
'         txtOthers(pnCtr).Enabled = True
'      End Select
'   Next
'
'   chkHsSerial.Enabled = True
'
'   If oDriver.EditMode = xeModeUpdate Then
'      For pnCtr = 0 To txtField.Count - 1
'         txtField(pnCtr).Locked = IIf(oApp.IsWarehouse, False, True)
'      Next
'
'      txtOthers(2).Enabled = pbEnblButtons
'      txtOthers(3).Enabled = pbEnblButtons
'      txtOthers(6).Enabled = pbEnblButtons
'   Else
'      For pnCtr = 0 To txtField.Count - 1
'         txtField(pnCtr).Locked = False
'      Next
'   End If
'End Sub
'
'Private Sub InitOthers()
'   For pnCtr = 0 To txtOthers.Count - 3
'      Select Case pnCtr
'      Case 2 To 5, 7, 8, 9
'         oRS(pnCtr) = 0
'         txtOthers(pnCtr).Text = Format(oRS(pnCtr), "#,##0")
'      Case 6
'         oRS(pnCtr) = oApp.ServerDate
'         txtOthers(pnCtr).Text = Format(oRS(pnCtr), "MMM DD, YYYY")
'      Case Else
'         oRS(pnCtr) = Empty
'         txtOthers(pnCtr).Text = oRS(pnCtr)
'      End Select
'   Next
'
'   txtOthers(10).Text = ""
'   txtOthers(11).Text = ""
'
'
'   oRS("cRecdStat") = xeRecStateActive
'   oRS("sStockIDx") = oDriver.FieldValue(18)
'   oRS("sBranchCd") = psBranchCd
'   oRS("nFloatQty") = 0
'   oRS("nLedgerNo") = 0
'End Sub
'
'Private Sub oDriver_InitValue()
'   Dim lsOldProc As String
'
'   lsOldProc = "oDriver_InitValue"
'   'On Error GoTo errProc
'
'   oDriver.FieldReference(0) = True
'   oDriver.FieldValue(0) = NewBarrCode
'   txtField(0).Text = oDriver.FieldValue(0)
'   oDriver.FieldValue(18) = GetNextCode("CP_Inventory", "sStockIDx", True, oApp.Connection, True, psBranchCd)
'   oDriver.FieldValue(19) = xeRecStateActive
'
'   For pnCtr = 0 To txtOthers.Count - 1
'      Select Case pnCtr
'      Case 2 To 5, 7, 8, 9
'         txtOthers(pnCtr).Text = 0
'      Case 6
'         txtOthers(pnCtr).Text = Format(oApp.ServerDate, "MMM DD, YYYY")
'      Case Else
'         txtOthers(pnCtr).Text = ""
'      End Select
'      txtOthers(pnCtr).Tag = ""
'   Next
'
'   chkHsSerial.Value = 0
'   oDriver.FieldValue(17) = chkHsSerial.Value
'
'   txtOthers(2).Locked = False
'   txtOthers(3).Locked = False
'
'   oRS.AddNew
'   InitOthers
'   pbEnblButtons = True
'   pbNewInvntory = True
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " & " )"
'End Sub
'
'Private Sub oDriver_LoadOtherData()
'   Dim lsOldProc As String
'   Dim lsSQL As String
'
'   lsOldProc = "oDriver_LoadOtherData"
'   'On Error GoTo errProc
'
'   Set oRS = New ADODB.Recordset
'   lsSQL = AddCondition(psCPInventory, "sStockIDx = " & strParm(oDriver.FieldValue(18)) _
'                                    & " AND sBranchCd = " & strParm(psBranchCd)) _
'                                    & " AND cRecdStat = " & strParm(xeRecStateActive)
'   oRS.Open lsSQL, oApp.Connection, adOpenStatic, adLockOptimistic, adCmdText
'   Set oRS.ActiveConnection = Nothing
'
'   If oRS.EOF Then
'      oRS.AddNew
'      InitOthers
'   Else
'      For pnCtr = 0 To txtOthers.Count - 1
'         Select Case pnCtr
'         Case 0
'            If Not IsNull(oRS("sSectnIDx")) Then SearchOthers pnCtr, oRS("sSectnIDx"), True
'         Case 1
'            If Not IsNull(oRS("sBinIDxxx")) Then SearchOthers pnCtr, oRS("sBinIDxxx"), True
'         Case 2 To 5, 7, 8, 9
'            txtOthers(pnCtr).Text = Format(IIf(IsNull(oRS(pnCtr)), 0, oRS(pnCtr)), "#,##0")
'         Case 6
'            txtOthers(pnCtr).Text = IIf(IsNull(oRS(pnCtr)), "", Format(oRS(pnCtr), "MMM DD, YYYY"))
'         End Select
'      Next
'      chkHsSerial.Value = oDriver.FieldValue(17)
'      pbNewInvntory = False
'   End If
'
'   txtOthers(10).Text = oDriver.FieldValue(0)
'   txtOthers(10).Tag = txtOthers(10).Text
'
'   txtOthers(11).Text = oDriver.FieldValue(1)
'   txtOthers(11).Tag = txtOthers(11).Text
'
'   txtOthers(12).Text = oBranch.Master("sBranchNm")
'   psBranchCd = oBranch.Master("sBranchCd")
'
'   If oApp.UserLevel > xeSupervisor Then
'      txtField(11).Text = Format(IIf(IsNull(oDriver.FieldValue(11)), 0, oDriver.FieldValue(11)), "#,##0.00")
'   Else
'      txtField(11).Text = Format(Price2Code(IIf(IsNull(oDriver.FieldValue(11)), 0, oDriver.FieldValue(11)), psPriceCode), "#,##0.00")
'   End If
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " & " )"
'End Sub
'
'Private Sub oDriver_Save(Saved As Boolean)
'   Saved = False
'End Sub
'
'Private Sub oDriver_WillSave(Cancel As Boolean)
'   Dim lsOldProc As String
'   Dim lnCtr As Integer
'
'   lsOldProc = "oDriver_WillSave"
'   'On Error GoTo errProc
'
'   If oDriver.FieldValue(0) = "" Then
'      MsgBox "Invalid BarrCode detected!!!", vbCritical, "Warning"
'      txtField(0).SetFocus
'      Cancel = True
'   ElseIf oDriver.FieldValue(1) = "" Then
'      MsgBox "Invalid Description detected!!!", vbCritical, "Warning"
'      txtField(1).SetFocus
'      Cancel = True
'   ElseIf txtOthers(0).Text = "" Then
'      MsgBox "Invalid Section detected!!!", vbCritical, "Warning"
'      txtOthers(0).SetFocus
'      Cancel = True
'   ElseIf txtOthers(1).Text = "" Then
'      MsgBox "Invalid Level detected!!!", vbCritical, "Warning"
'      txtOthers(1).SetFocus
'      Cancel = True
''   ElseIf CDbl(txtField(13).Text) = 0# Then
''      MsgBox "Invalid Selling Price detected!!!", vbCritical, "Warning"
''      txtField(13).SetFocus
''      Cancel = True
'   ElseIf oDriver.FieldValue(18) = "" Then
'      MsgBox "Invalid Stock ID Detected!!!" & vbCrLf & _
'               "Please contact GMC_SEG for assistant!!!", vbCritical, "Warning"
'      Cancel = True
'   Else
'      Cancel = Not UpdateCPInventory
''      If pbNewInvntory Then Cancel = Not SaveCPInventoryLedger
'   End If
'
'   oDriver.FieldValue(17) = chkHsSerial.Value
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " & Cancel & " )"
'End Sub
'
'Private Sub txtField_GotFocus(Index As Integer)
'   With txtField(Index)
'      .SelStart = 0
'      .SelLength = Len(.Text)
'      .BackColor = oApp.getColor("HT1")
'   End With
'
'   oDriver.ColumnIndex = Index
'   pbtxtOthers = False
'   pnIndex = Index
'End Sub
'
'Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'   Dim lsOldProc As String
'
'   lsOldProc = "txtField_KeyDown"
'   'On Error GoTo errProc
'
'   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
'      With txtField(Index)
'         If KeyCode = vbKeyF3 Then
'            oDriver.RecordSearch .Text
'            If .Text <> "" Then SetNextFocus
'         Else
'            If .Text <> "" Then oDriver.RecordSearch .Text
'         End If
'      End With
'      KeyCode = 0
'   End If
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " & Index _
'                       & ", " & KeyCode _
'                       & ", " & Shift & " )", True
'End Sub
'
'Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
'   Dim lsOldProc As String
'
'   lsOldProc = "txtField_Validate"
'   'On Error GoTo errProc
'
'   With txtField(Index)
'      Select Case Index
'      Case 0
'         .Text = UCase(.Text)
'      Case Else
'         .Text = TitleCase(.Text)
'      End Select
'      Cancel = Not oDriver.ValidateField(Index)
'   End With
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " & Index _
'                       & ", " & Cancel & " )", True
'End Sub
'
'Private Sub txtField_LostFocus(Index As Integer)
'   With txtField(Index)
'      .BackColor = oApp.getColor("EB")
'   End With
'End Sub
'
'Private Sub txtOthers_GotFocus(Index As Integer)
'   With txtOthers(Index)
'      .SelStart = 0
'      .SelLength = Len(.Text)
'      .BackColor = oApp.getColor("HT1")
'   End With
'
'   pbtxtOthers = True
'   pnIndex = Index
'End Sub
'
'Private Sub txtOthers_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'   Dim lsSearch() As String
'   Dim lnCtr As Integer
'   Dim lsSQL As String
'   Dim lsOldProc As String
'
'   lsOldProc = "txtOthers_KeyDown"
'   'On Error GoTo errProc
'
'   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
'      With txtOthers(Index)
'         Select Case Index
'         Case 0, 1
'            If KeyCode = vbKeyF3 Then
'               SearchOthers Index, .Text, False
'               If .Text <> "" Then SetNextFocus
'            Else
'               If .Text <> "" Then
'                  If Trim(LCase(.Text)) <> Trim(LCase(.Tag)) Then SearchOthers Index, .Text, False
'               End If
'            End If
'            .Tag = .Text
'         Case 10, 11
'            Call txtField_Validate(Index, False)
'         End Select
'      End With
'      KeyCode = 0
'   End If
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " & Index _
'                       & ", " & KeyCode _
'                       & ", " & Shift & " )", True
'End Sub
'
'Private Sub txtOthers_LostFocus(Index As Integer)
'   With txtOthers(Index)
'      .BackColor = oApp.getColor("EB")
'   End With
'End Sub
'
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'   Select Case KeyCode
'   Case vbKeyReturn, vbKeyUp, vbKeyDown
'      Select Case KeyCode
'      Case vbKeyReturn, vbKeyDown
'         SetNextFocus
'      Case vbKeyUp
'         SetPreviousFocus
'      End Select
'   Case vbKeyF9
'      If Not IsNumeric(txtField(11).Text) Then
'         txtField(11).Text = Format(Code2Price(txtField(11).Text, psPriceCode), "#,##0.00")
'      Else
'         txtField(11).Text = Price2Code(txtField(11).Text, psPriceCode)
'      End If
'   End Select
'End Sub
'
'Private Sub SearchOthers(ByVal lnIndex As Integer, _
'                         ByVal lsValue As String, _
'                         ByVal lbByCode As Boolean)
'
'   Dim lsSQL As String
'   Dim lsBrowse As String
'   Dim lrs As ADODB.Recordset
'   Dim lsOldProc As String
'   Dim lsSelected() As String
'   Dim lsFieldIDx As String
'   Dim lsFieldNmx As String
'
'   lsOldProc = "SearchOthers"
'   'On Error GoTo errProc
'
'   Select Case lnIndex
'   Case 0
'      lsSQL = "SELECT" _
'                  & "  sSectnIDx" _
'                  & ", sSectnNme" _
'               & " FROM Section" _
'               & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
'                  & IIf(lbByCode, "AND sSectnIDx = " & strParm(lsValue), "AND sSectnNme LIKE " & strParm(lsValue & "%")) _
'               & " ORDER BY sSectnNme"
'      lsFieldIDx = "sSectnIDx"
'      lsFieldNmx = "sSectnNme"
'   Case 1
'      lsSQL = "SELECT" _
'                  & "  sBinIDxxx" _
'                  & ", sBinNamex" _
'               & " FROM Bin" _
'               & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
'                  & IIf(lbByCode, "AND sBinIDxxx = " & strParm(lsValue), "AND sBinNamex LIKE " & strParm(lsValue & "%")) _
'               & " ORDER BY sBinNamex"
'      lsFieldIDx = "sLevelIDx"
'      lsFieldNmx = "sLevelNme"
'   End Select
'
'   Set lrs = New ADODB.Recordset
'   lrs.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
'
'   With txtOthers(lnIndex)
'      If lrs.RecordCount = 1 Then
'         oRS(lsFieldIDx) = lrs(0)
'         .Text = lrs(1)
'      ElseIf lrs.RecordCount > 1 Then
'         lsBrowse = KwikBrowse(oApp, lrs _
'                        , lsFieldIDx & "»" & lsFieldNmx _
'                        , "Code»Description")
'
'         If lsBrowse <> "" Then
'            lsSelected = Split(lsBrowse, "»")
'            oRS(lsFieldIDx) = lsSelected(0)
'            .Text = lsSelected(1)
'         Else
'            If oRS(lsFieldIDx) <> "" Then .Text = .Tag
'         End If
'      Else
'         .Text = ""
'         oRS(lsFieldIDx) = ""
'      End If
'
'      .Tag = .Text
'   End With
'   Set lrs = Nothing
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " & lsFieldIDx _
'                       & ", " & lsFieldNmx _
'                       & ", " & lnIndex _
'                       & ", " & lsValue & " )"
'End Sub
'
'Private Sub oDriver_Delete(Deleted As Boolean)
'   Deleted = True
'End Sub
'
'Private Sub oDriver_DeleteComplete()
'   For pnCtr = 0 To txtOthers.Count - 1
'      txtOthers(pnCtr).Text = ""
'   Next
'
'   chkHsSerial.Value = 0
'End Sub
'
'Private Sub txtOthers_Validate(Index As Integer, Cancel As Boolean)
'   Dim lsOldProc As String
'   Dim lnCtr As Integer
'
'   lsOldProc = "txtOthers_Validate"
'   'On Error GoTo errProc
'
'   With txtOthers(Index)
'      .Text = TitleCase(.Text)
'
'      Select Case Index
'      Case 0, 1
'         If .Text <> "" Then
'            If Trim(LCase(.Text)) <> Trim(LCase(.Tag)) Then SearchOthers Index, .Text, False
'         End If
'
'         .Tag = .Text
'      Case 2, 3, 4, 5, 6, 7, 8, 9
'         If Not IsNumeric(.Text) Then .Text = 0
'         .Text = Format(.Text, "#,##0")
'         oRS(Index) = CDbl(.Text)
'      Case 10, 11
'         If Trim(.Text) = "" Then
'            If oRS.EOF Then Exit Sub
'            InitOthers
'            txtOthers(IIf(Index = 10, 11, 10)).Text = ""
'            txtOthers(IIf(Index = 10, 11, 10)).Tag = ""
'
'            For lnCtr = 0 To txtField.Count - 1
'               Select Case lnCtr
'               Case 0
'               Case 11, 12
'                  txtField(lnCtr).Text = "0.00"
'               Case 13, 14, 15
'                  txtField(lnCtr).Text = 0
'               Case Else
'                  txtField(lnCtr).Text = ""
'                  txtField(lnCtr).Tag = txtField(lnCtr).Text
'               End Select
'            Next
'            Exit Sub
'         End If
'
'         If .Tag <> .Text Then SearchBarCode .Text, IIf(Index = 10, True, False)
'         .Tag = .Text
'      Case 12
'         If .Text = "" Then
'            Cancel = True
'            Exit Sub
'         End If
'
'         If oBranch.SearchRecord(.Text, False) Then
'            psBranchCd = oBranch.Master("sBranchCd")
'            .Text = oBranch.Master("sBranchNm")
'         Else
'            If Trim(.Tag) <> "" Then
'               .Text = .Tag
'               Exit Sub
'            End If
'
'            .SetFocus
'         End If
'      End Select
'   End With
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " & Cancel & " )", True
'End Sub
'
'Function NewBarrCode() As String
'   Dim lrs As Recordset
'   Dim lsSQL As String
'   Dim lnCtr As Long
'   Dim lsOldProc As String
'   Dim lrsBranch As Recordset
'   Dim lsCode As String
'
'   lsOldProc = "NewBarrCode"
'   'On Error GoTo errProc
'
'   Set lrsBranch = New ADODB.Recordset
'   lrsBranch.Open "SELECT" _
'                     & "  a.sCompnyCd" _
'                  & " FROM Company a" _
'                     & ", Branch b" _
'                  & " WHERE a.sCompnyID = b.sCompnyID" _
'                     & " AND b.sBranchCd = " & strParm(psBranchCd) _
'                  , oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
'
'   lsCode = "GMC"
'   If Not lrsBranch.EOF Then lsCode = lrsBranch("sCompnyCd")
'
'   lsSQL = "SELECT" & _
'               " sBarrCode" & _
'            " FROM CP_Inventory" & _
'            " WHERE sBarrCode LIKE " & strParm(Format(Date, "yy") & "-" & lsCode & "-%") & _
'            " ORDER BY sBarrCode DESC" & _
'            " LIMIT 1"
'   Set lrs = New Recordset
'   lrs.Open lsSQL, oApp.Connection, , , adCmdText
'
'   If lrs.EOF Then
'      lnCtr = 1
'   Else
'      If Left(lrs("sBarrCode"), 2) = Format(Date, "yy") Then
'         lnCtr = CLng(Right(lrs("sBarrCode"), 6)) + 1
'      Else
'         lnCtr = 1
'      End If
'   End If
'   NewBarrCode = Format(Date, "yy") & "-" & lsCode & "-" & Format(lnCtr, "000000")
'
'   Set lrs = Nothing
'   Set lrsBranch = Nothing
'
'endProc:
'   Exit Function
'errProc:
'   ShowError lsOldProc & "( " & " )"
'End Function
'
'Private Sub SearchBarCode(ByVal lsValue As String, ByVal lbByCode As Boolean)
'   Dim lrsCellphone As ADODB.Recordset
'   Dim lrsInventoryx As ADODB.Recordset
'   Dim lsSelected() As String
'   Dim lsBrowse As String
'   Dim lsOldProc As String
'   Dim lsStockIDx As String
'   Dim lsSQL As String
'
'   lsOldProc = "SearchSpareParts"
'   'On Error GoTo errProc
'
'   lsSQL = "SELECT" _
'               & "  a.sStockIDx" _
'               & ", a.sBarrcode" _
'               & ", a.sDescript" _
'               & ", b.sBrandNme" _
'               & ", c.sModelNme" _
'            & " FROM CP_Inventory a" _
'               & " LEFT JOIN CP_Brand b" _
'                  & " ON a.sBrandIDx = b.sBrandIDx" _
'               & " LEFT JOIN CP_Model c" _
'                  & " ON a.sModelIDx = c.sModelIDx" _
'            & " ORDER BY a.sBarrCode" _
'               & ", a.sDescript"
'
'   lsSQL = AddCondition(lsSQL, IIf(lbByCode, "a.sBarrcode LIKE " & strParm(lsValue & "%"), "a.sDescript LIKE " & strParm(lsValue & "%")))
'   Set lrsCellphone = New ADODB.Recordset
'   With lrsCellphone
'      .Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
'      If .EOF Then
'         oDriver_InitValue
'         For pnCtr = 0 To txtField.Count - 1
'            txtField(pnCtr).Text = Empty
'         Next
'         GoTo endProc
'      End If
'
'      If .RecordCount = 1 Then
'         lsStockIDx = .Fields("sStockIDx")
'         GoTo LoadRecord
'      Else
'         With txtOthers(pnIndex)
'            .BackColor = oApp.getColor("EB")
'            lsBrowse = KwikBrowse(oApp, lrsCellphone _
'                                    , "sBarrcode»sDescript»sBrandNme»sModelNme" _
'                                    , "BarrCode»Description»Brand»Model" _
'                                    , "@»@»@»@" _
'                                    , "a.sBarrcode»a.sDescript»b.sBrandNme»c.sModelNme")
'
'            If lsBrowse <> "" Then
'               lsSelected = Split(lsBrowse, "»")
'               lsStockIDx = lsSelected(0)
'               GoTo LoadRecord
'            End If
'            .BackColor = oApp.getColor("HT1")
'            .SelStart = 0
'            .SelLength = Len(.Text)
'         End With
'      End If
'      GoTo endProc
'   End With
'
'LoadRecord:
'   lsSQL = "SELECT" _
'               & "  a.sStockIDx" _
'               & ", a.sBranchCd" _
'               & ", a.cRecdStat" _
'               & ", b.sBarrcode" _
'            & " FROM CP_Inventory_Master a" _
'               & ", CP_Inventory b" _
'            & " WHERE a.sStockIDx = b.sStockIDx" _
'            & " ORDER BY a.sBranchCd DESC"
'   lsSQL = AddCondition(lsSQL, "a.sStockIDx = " & strParm(lsStockIDx))
'   Set lrsInventoryx = New ADODB.Recordset
'
'   pbEnblButtons = True
'   pbNewInvntory = False
'   With lrsInventoryx
'      .Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
'      If Not .EOF Then
'         .Find "sBranchCd = " & strParm(psBranchCd), 0, adSearchForward
'         If Not lrsInventoryx.EOF Then
'            If .Fields("cRecdStat") = xeRecStateActive Then
'               oDriver.LookupValue(0) = .Fields("sBarrcode")
'               oDriver.LoadRecord
'
'               pbEnblButtons = False
'            Else
'               MsgBox "CP Inventory Status is Deactivated!!!" & vbCrLf & _
'                        "Please Save the record to activate!!!", vbInformation, "Notice"
'
'               oDriver.LookupValue(0) = .Fields("sBarrcode")
'               oDriver.LoadRecord
'               oDriver.RecordUpdate
'
'               txtOthers(0).SetFocus
'               oRS.Fields("nBegQtyxx") = 0
'               oRS.Fields("nQtyOnHnd") = 0
'               oRS.Fields("cRecdStat") = xeRecStateActive
'
'               txtOthers(2).Text = 0
'               txtOthers(3).Text = 0
'            End If
'         Else
'            MsgBox "No Inventory found in your warehouse!!!" & vbCrLf & _
'                     "Plese Save the record to create!!!", vbInformation, "Notice"
'
'            .MoveFirst
'            oDriver.LookupValue(0) = .Fields("sBarrcode")
'            oDriver.LoadRecord
'            oDriver.RecordUpdate
'            txtOthers(0).SetFocus
'
'            pbNewInvntory = True
'         End If
'      Else
'         'no record at all
'         pbNewInvntory = True
'         If Not lrsCellphone.EOF Then
'            MsgBox "No Inventory found in your warehouse!!!" & vbCrLf & _
'                     "Plese Save the record to create!!!", vbInformation, "Notice"
'
'            oDriver.LookupValue(0) = lrsCellphone("sBarrcode")
'            oDriver.LoadRecord
'            oDriver.RecordUpdate
'            txtOthers(0).SetFocus
'
'            pbNewInvntory = True
'         End If
'      End If
'      .Close
'   End With
'endProc:
'   Set lrsCellphone = Nothing
'   Set lrsInventoryx = Nothing
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " & oDriver.FieldValue(2) & "»" & oDriver.FieldValue(3) & " )", True
'End Sub
'
'Private Function UpdateCPInventory() As Boolean
'   Dim lsOldProc As String
'   Dim lrs As ADODB.Recordset
'   Dim lsSQL As String
'   Dim lnRow As Integer
'
'   lsOldProc = "UpdateCPInventory"
'   'On Error GoTo errProc
'
'   lsSQL = AddCondition(psCPInventory, "sStockIDx = " & strParm(oDriver.FieldValue(18)) _
'                  & " AND sBranchCd = " & strParm(psBranchCd))
'
'   Set lrs = New ADODB.Recordset
'   lrs.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
'
'   If lrs.EOF Then
'      lsSQL = ADO2SQL(oRS, "CP_Inventory_Master", "", oApp.UserID, oApp.ServerDate, "")
'   Else
'      lsSQL = ADO2SQL(oRS, "CP_Inventory_Master", "sStockIDx = " & strParm(oDriver.FieldValue(18)) & " AND sBranchCd = " & strParm(psBranchCd), oApp.UserID, oApp.ServerDate, "")
'   End If
'
'   If lsSQL <> "" Then
'      lnRow = oApp.Execute(lsSQL, "CP_Inventory_Master", psBranchCd, "")
'      If lnRow <= 0 Then
'         MsgBox "Unable to Save Inventory" & vbCrLf & _
'                  lsSQL, vbCritical, "Warning"
'         GoTo endProc
'      End If
'   End If
'
'   UpdateCPInventory = True
'
'endProc:
'   Set lrs = Nothing
'   Exit Function
'errProc:
'   ShowError lsOldProc & "( " & UpdateCPInventory & " )", True
'End Function
'
'Private Function SaveCPInventoryLedger() As Boolean
'   Dim lsSQL As String
'   Dim lnRow As Long
'   Dim lsOldProc As String
'
'   lsOldProc = "SaveSPInventoryLedger"
'   'On Error GoTo errProc
'
'   lsSQL = "INSERT INTO CP_Inventory_Ledger SET" _
'               & "  sStockIDx = " & strParm(oDriver.FieldValue(18)) _
'               & ", sBranchCd = " & strParm(psBranchCd) _
'               & ", sSourceCd = 'CPAd'" _
'               & ", sSourceNo = '9900000001'" _
'               & ", nQtyInxxx = " & CLng(txtOthers(3).Text) _
'               & ", nQtyOutxx = '0'" _
'               & ", nQtyOrder = '0'" _
'               & ", nQtyIssue = '0'" _
'               & ", nLedgerNo = '000001'" _
'               & ", nQtyOnHnd = '0'" _
'               & ", dTransact = " & dateParm(oApp.ServerDate) _
'               & ", dModified = " & dateParm(oApp.ServerDate)
'
'   lnRow = oApp.Execute(lsSQL, "CP_Inventory_Ledger", psBranchCd)
'   If lnRow <= 0 Then
'      MsgBox "Unable to Save SP_Inventory_Ledger!!!", vbCritical, "Warning"
'      GoTo endProc
'   End If
'   SaveCPInventoryLedger = True
'
'endProc:
'   Exit Function
'errProc:
'   ShowError lsOldProc & "( " & " )"
'End Function
'
'Private Sub ComputePricing()
'   Dim lnPurPrice As Double
'
'   lnPurPrice = oDriver.FieldValue(17)
'
'   If lnPurPrice >= CDbl(txtField(7).Text) Then
'      txtField(7).Text = "0.00"
'
'      oDriver.FieldValue(7) = txtField(7).Text
'      Exit Sub
'   End If
'
'   If lnPurPrice > Round(CDbl(txtField(7).Text) * (100 - oDriver.FieldValue(7)) / 100, 2) Then
'      txtOthers(11).Text = "0.00"
'      Exit Sub
'   End If
'
''   If lnPurPrice > Round(CDbl(txtOthers(10).Text) * (100 - CDbl(txtOthers(12).Text)) / 100, 2) Then
''      txtOthers(12).Text = "0.00"
''      Exit Sub
''   End If
''
''   If CDbl(txtOthers(11).Text) > CDbl(txtOthers(12).Text) Then txtOthers(12).Text = "0.00"
'End Sub
'
'Private Sub ShowError(ByVal lsProcName As String, Optional bEnd As Boolean = False)
'   With oApp
'      .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
'      If bEnd Then
'         .xShowError
'         End
'      Else
'         With Err
'            .Raise .Number, .Source, .Description
'         End With
'      End If
'   End With
'End Sub
