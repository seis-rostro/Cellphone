VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmExtremePromoCat 
   BorderStyle     =   0  'None
   Caption         =   "North Point Extreme Appliances Promo Rate"
   ClientHeight    =   8220
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   13680
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2640
      Left            =   7515
      TabIndex        =   49
      Top             =   5475
      Width           =   6030
      _ExtentX        =   10636
      _ExtentY        =   4657
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      FocusRect       =   0
      SelectionMode   =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   120
      TabIndex        =   44
      Top             =   1815
      Width           =   1200
      _ExtentX        =   2117
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
      Picture         =   "frmExtremePromoCat.frx":0000
      CaptionAlign    =   0
   End
   Begin xrControl.xrFrame xrFrame 
      Height          =   4860
      Index           =   0
      Left            =   1605
      Tag             =   "wt0;fb0"
      Top             =   585
      Width           =   5865
      _ExtentX        =   10345
      _ExtentY        =   8573
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmExtremePromoCat.frx":077A
         Left            =   1215
         List            =   "frmExtremePromoCat.frx":0790
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   780
         Width           =   2040
      End
      Begin VB.Frame Frame1 
         Caption         =   "TERM"
         Height          =   1185
         Index           =   1
         Left            =   150
         TabIndex        =   20
         Tag             =   "wt0;fb0"
         Top             =   3555
         Width           =   5550
         Begin VB.TextBox txtField 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   12
            Left            =   3075
            TabIndex        =   28
            Text            =   "0.00"
            Top             =   750
            Width           =   495
         End
         Begin VB.TextBox txtField 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   11
            Left            =   3075
            TabIndex        =   26
            Text            =   "0.00"
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox txtField 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   10
            Left            =   1365
            TabIndex        =   24
            Text            =   "0.00"
            Top             =   750
            Width           =   495
         End
         Begin VB.TextBox txtField 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   9
            Left            =   1365
            TabIndex        =   22
            Text            =   "0.00"
            Top             =   360
            Width           =   495
         End
         Begin VB.CheckBox chkZeroInt 
            Caption         =   "ZERO INT."
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
            Left            =   4200
            TabIndex        =   29
            Tag             =   "wt0;fb0"
            Top             =   360
            Width           =   1320
         End
         Begin VB.CheckBox chkTerm 
            Caption         =   "12 Months"
            Height          =   195
            Index           =   3
            Left            =   1980
            TabIndex        =   27
            Tag             =   "wt0;fb0"
            Top             =   810
            Width           =   1185
         End
         Begin VB.CheckBox chkTerm 
            Caption         =   "9 Months"
            Height          =   195
            Index           =   2
            Left            =   1980
            TabIndex        =   25
            Tag             =   "wt0;fb0"
            Top             =   420
            Width           =   1185
         End
         Begin VB.CheckBox chkTerm 
            Caption         =   "6 Months"
            Height          =   195
            Index           =   1
            Left            =   300
            TabIndex        =   23
            Tag             =   "wt0;fb0"
            Top             =   810
            Width           =   1185
         End
         Begin VB.CheckBox chkTerm 
            Caption         =   "3 Months"
            Height          =   195
            Index           =   0
            Left            =   285
            TabIndex        =   21
            Tag             =   "wt0;fb0"
            Top             =   420
            Width           =   1185
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "SC"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   9
            Left            =   3210
            TabIndex        =   51
            Top             =   90
            Width           =   225
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "SC"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   1515
            TabIndex        =   50
            Top             =   90
            Width           =   225
         End
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   4845
         TabIndex        =   19
         Text            =   "0.00"
         Top             =   2730
         Width           =   795
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   4845
         TabIndex        =   17
         Text            =   "250.00"
         Top             =   2325
         Width           =   795
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   2805
         TabIndex        =   15
         Text            =   "1,250.00"
         Top             =   3150
         Width           =   795
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   2805
         TabIndex        =   13
         Text            =   "30.00"
         Top             =   2745
         Width           =   795
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   2805
         TabIndex        =   11
         Text            =   "30.00"
         Top             =   2340
         Width           =   795
      End
      Begin VB.TextBox txtField 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   1560
         TabIndex        =   9
         Text            =   "December 30, 2000"
         Top             =   1935
         Width           =   2040
      End
      Begin VB.TextBox txtField 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   1560
         TabIndex        =   7
         Text            =   "December 30, 2000"
         Top             =   1530
         Width           =   2040
      End
      Begin VB.TextBox txtField 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1215
         TabIndex        =   5
         Text            =   "Text9"
         Top             =   1125
         Width           =   4440
      End
      Begin VB.TextBox txtField 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   1245
         TabIndex        =   1
         Text            =   "Text9"
         Top             =   120
         Width           =   2280
      End
      Begin xrControl.xrFrame xrFrame 
         Height          =   3495
         Index           =   3
         Left            =   10080
         Tag             =   "wt0;fb0"
         Top             =   3240
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   6165
         BackColor       =   12632256
         ClipControls    =   0   'False
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   10
         Left            =   225
         TabIndex        =   2
         Top             =   817
         Width           =   630
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "End Mrtg"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   3285
         TabIndex        =   18
         Top             =   2775
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Rebate"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   3300
         TabIndex        =   16
         Top             =   2385
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Min MP Payment"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   1245
         TabIndex        =   14
         Top             =   3195
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Max. Down %"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   1245
         TabIndex        =   12
         Top             =   2820
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Min. Down %"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   1245
         TabIndex        =   10
         Top             =   2370
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Date Thru"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   -30
         TabIndex        =   8
         Top             =   2010
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Date From"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   -15
         TabIndex        =   6
         Top             =   1590
         Width           =   1455
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   375
         Left            =   1395
         Tag             =   "et0;ht2"
         Top             =   240
         Width           =   2250
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Category ID"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   180
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   225
         TabIndex        =   4
         Top             =   1185
         Width           =   1455
      End
   End
   Begin xrControl.xrFrame xrFrame 
      Height          =   2670
      Index           =   1
      Left            =   1605
      Tag             =   "wt0;fb0"
      Top             =   5460
      Width           =   5865
      _ExtentX        =   10345
      _ExtentY        =   4710
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.Frame Frame1 
         Caption         =   "Model"
         Height          =   2415
         Index           =   0
         Left            =   120
         TabIndex        =   30
         Tag             =   "wt0;fb0"
         Top             =   105
         Width           =   5580
         Begin VB.TextBox txtDetail 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   375
            Index           =   3
            Left            =   3180
            TabIndex        =   38
            Text            =   "0.00"
            Top             =   1635
            Width           =   1035
         End
         Begin VB.TextBox txtDetail 
            Enabled         =   0   'False
            Height          =   375
            Index           =   2
            Left            =   1320
            TabIndex        =   36
            Text            =   "Text1"
            Top             =   1230
            Width           =   2895
         End
         Begin VB.TextBox txtDetail 
            Height          =   375
            Index           =   1
            Left            =   1320
            TabIndex        =   34
            Text            =   "Text1"
            Top             =   825
            Width           =   2895
         End
         Begin VB.TextBox txtDetail 
            Height          =   375
            Index           =   0
            Left            =   1320
            TabIndex        =   32
            Text            =   "Text1"
            Top             =   420
            Width           =   2895
         End
         Begin xrControl.xrButton cmdButton 
            Height          =   300
            Index           =   2
            Left            =   4485
            TabIndex        =   40
            Top             =   750
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   529
            Caption         =   "&Del Row"
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
            CaptionAlign    =   0
         End
         Begin xrControl.xrButton cmdButton 
            Height          =   300
            Index           =   8
            Left            =   4485
            TabIndex        =   39
            Top             =   420
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   529
            Caption         =   "&Add Row"
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
            CaptionAlign    =   0
         End
         Begin xrControl.xrButton cmdButton 
            CausesValidation=   0   'False
            Height          =   300
            Index           =   1
            Left            =   4485
            TabIndex        =   41
            Top             =   1080
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   529
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
            CaptionAlign    =   0
         End
         Begin VB.Shape Shape3 
            Height          =   1965
            Index           =   0
            Left            =   150
            Top             =   255
            Width           =   5235
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Model Code"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   285
            TabIndex        =   35
            Top             =   1275
            Width           =   945
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Selling Price"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   1695
            TabIndex        =   37
            Top             =   1680
            Width           =   1395
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Model"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   480
            TabIndex        =   33
            Top             =   870
            Width           =   735
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Brand"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   720
            TabIndex        =   31
            Top             =   465
            Width           =   495
         End
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   120
      TabIndex        =   43
      Top             =   1200
      Width           =   1200
      _ExtentX        =   2117
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
      Picture         =   "frmExtremePromoCat.frx":07F5
      CaptionAlign    =   0
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   120
      TabIndex        =   46
      Top             =   585
      Width           =   1200
      _ExtentX        =   2117
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
      Picture         =   "frmExtremePromoCat.frx":0F6F
      CaptionAlign    =   0
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   6
      Left            =   120
      TabIndex        =   45
      Top             =   2430
      Width           =   1200
      _ExtentX        =   2117
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
      Picture         =   "frmExtremePromoCat.frx":16E9
      CaptionAlign    =   0
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   600
      Index           =   3
      Left            =   120
      TabIndex        =   47
      Top             =   1200
      Width           =   1200
      _ExtentX        =   2117
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
      Picture         =   "frmExtremePromoCat.frx":1E63
      CaptionAlign    =   0
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Height          =   4830
      Left            =   7515
      TabIndex        =   48
      Top             =   600
      Width           =   6030
      _ExtentX        =   10636
      _ExtentY        =   8520
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      AllowBigSelection=   0   'False
      FocusRect       =   0
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   7
      Left            =   120
      TabIndex        =   42
      Top             =   585
      Width           =   1200
      _ExtentX        =   2117
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
      Picture         =   "frmExtremePromoCat.frx":25DD
   End
End
Attribute VB_Name = "frmExtremePromoCat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pxeMODULENAME = "frmMPPromoCat"
Private WithEvents oTrans As clsCPPromo
Attribute oTrans.VB_VarHelpID = -1

Private oSkin As clsFormSkin

Private pMIndex As Integer
Private pnActiveRow As Integer
Private pDIndex As Integer
Private pnCtr As Integer

Private pbMasterGotFocus As Boolean
Private pbDetailGotFocus As Boolean
Private pbFormLoad As Boolean

Private Sub InitGrid()
    Dim lnCtr As Integer
    
    With MSFlexGrid2
        .Cols = 5
        .Rows = 2
        .Clear
        
        .TextMatrix(0, 0) = "No."
        .TextMatrix(0, 1) = "ID"
        .TextMatrix(0, 2) = "Description"
        .TextMatrix(0, 3) = "FROM"
        .TextMatrix(0, 4) = "THRU"
        
        .Row = 0
        .ColWidth(0) = 600
        .ColWidth(1) = 0
        .ColWidth(2) = 3100
        .ColWidth(3) = 1000
        .ColWidth(4) = 1000
        
        For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = 3
      Next

      .Row = 1
      .ColAlignment(0) = flexAlignCenterCenter
      .ColAlignment(1) = flexAlignLeftCenter
      .ColAlignment(2) = flexAlignLeftCenter
      .ColAlignment(3) = flexAlignLeftCenter
      .ColAlignment(4) = flexAlignLeftCenter
      
      .Col = 1
      .Row = 1
      .ColSel = .Cols - 1
   End With
    
    With MSFlexGrid1
        .Cols = 4
        .Rows = 2
        .Clear
        
        .TextMatrix(0, 0) = "No."
        .TextMatrix(0, 1) = "Model"
        .TextMatrix(0, 2) = "Model Code"
        .TextMatrix(0, 3) = "SRP"
        
        .Row = 0
        .ColWidth(0) = 500
        .ColWidth(1) = 2500
        .ColWidth(2) = 1900
        .ColWidth(3) = 1000
        
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = 3
      Next

      .Row = 1
      .ColAlignment(0) = flexAlignCenterCenter
      .ColAlignment(1) = flexAlignLeftCenter
      .ColAlignment(2) = flexAlignLeftCenter
      .ColAlignment(3) = flexAlignRightCenter
      
      .Col = 1
      .Row = 1
      .ColSel = .Cols - 1
   End With
End Sub

Private Sub chkTerm_Click(Index As Integer)
   Select Case Index
   Case 0
      oTrans.Master("n3monthsx") = chkTerm(Index).Value
   Case 1
      oTrans.Master("n6monthsx") = chkTerm(Index).Value
   Case 2
      oTrans.Master("n9monthsx") = chkTerm(Index).Value
   Case 3
      oTrans.Master("n12months") = chkTerm(Index).Value
   End Select
End Sub

Private Sub chkZeroInt_Click()
   oTrans.Master("nZeroIntx") = chkZeroInt.Value
End Sub

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   Dim lnMsg As String
   
   lsOldProc = "cmdButton_Click"
   ' 'On Error GoTo errProc
   
   With MSFlexGrid1
      Select Case Index
      Case 0 ' Save
         If .Rows > 2 Then
            pnCtr = 1
            Do While pnCtr < .Rows
               If Trim(.TextMatrix(pnCtr, 1)) = "" Then
                  .Row = pnCtr
                  If oTrans.deleteDetail(.Row - 1) Then Call LoadDetail
               Else
                  pnCtr = pnCtr + 1
               End If
            Loop
         End If
   
         If isEntryOk Then
            If oTrans.SaveTransaction() = True Then
               MsgBox "Transaction saved successfuly.", vbInformation, pxeMODULENAME
               Call initButton(oTrans.EditMode)
               Call LoadRecord
               Call ClearFields
            Else
               MsgBox "Unable to save transaction.", vbCritical, pxeMODULENAME
            End If
         End If
      Case 1 ' Search
         If pbDetailGotFocus Then
            If pDIndex = 0 Or pDIndex = 1 Then
               Call oTrans.searchDetail(pnActiveRow - 1, pDIndex)
            End If
         End If
      Case 2 ' Del. Row
         lnMsg = MsgBox("Do you want to delete this item?", vbYesNo + vbQuestion, "Confirm")
         If lnMsg = vbYes Then
            If oTrans.deleteDetail(pnActiveRow - 1) Then
               If oTrans.ItemCount = 0 Then oTrans.addDetail
               LoadDetail
               
               .Row = .Rows - 1
               .Col = 1
               .ColSel = .Cols - 1
               
               txtDetail(0) = oTrans.Detail(pnActiveRow - 1, "sBrandNme")
               txtDetail(1) = oTrans.Detail(pnActiveRow - 1, "sModelNme")
               txtDetail(2) = oTrans.Detail(pnActiveRow - 1, "sModelCde")
               txtDetail(3) = Format(oTrans.Detail(pnActiveRow - 1, "nSelPrice"), "#,##0.00")
               
               txtDetail(1).SetFocus
            End If
         End If
      Case 3 ' Cancel
         lnMsg = MsgBox("Transaction is in Update Mode!!!" & vbCrLf & _
                           "Do you want to Cancel Transaction!!!", vbYesNo + vbQuestion, "Confirm")
         If lnMsg = vbYes Then
            oTrans.InitTransaction
            ClearFields
            initButton oTrans.EditMode
         End If
      Case 4 ' Browse
         If oTrans.SearchTransaction Then
            LoadMaster
            LoadDetail
         End If
      Case 5 ' Update
         If oTrans.Master(27) <> 3 Then
            If txtField(1) <> "" And oTrans.Master(1) <> "" Then
               If oTrans.UpdateTransaction Then
                  Call initButton(oTrans.EditMode)
                  txtField(6).SetFocus
               End If
                  Else
                  lnMsg = MsgBox("Please select transaction before update!", vbCritical, pxeMODULENAME)
               End If
         Else
            lnMsg = MsgBox("Cannot update cancelled transaction!", vbCritical, pxeMODULENAME)
         End If
      Case 6 ' Close
         Unload Me
      Case 7 ' New
         If oTrans.NewTransaction Then
            ClearFields
            LoadMaster
            initButton oTrans.EditMode
            txtField(1).SetFocus
         End If
      Case 8 ' Add Detail
         If oTrans.Detail(pnActiveRow - 1, "sModelIDx") <> "" Then
            If oTrans.addDetail Then
               Call LoadDetail
               .Row = .Rows - 1
               .Col = 1
               .ColSel = .Cols - 1
               
               txtDetail(0) = oTrans.Detail(pnActiveRow - 1, "sBrandNme")
               txtDetail(1) = oTrans.Detail(pnActiveRow - 1, "sModelNme")
               txtDetail(2) = oTrans.Detail(pnActiveRow - 1, "sModelCde")
               txtDetail(3) = Format(oTrans.Detail(pnActiveRow - 1, "nSelPrice"), "#,##0.00")
               
               txtDetail(1).SetFocus
            End If
         End If
      End Select
   End With
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Function isEntryOk() As Boolean
   Dim lnCtr As Integer
   Dim lbWithTerm As Boolean
   
   If txtField(1) = "" Then
      MsgBox "Invalid Description Detected!" & vbCrLf & _
               "Pls Verify Entry Then Try Again!!!", vbCritical, "WARNING"
      txtField(1).SetFocus
      GoTo EntryNotOK
   End If
   
   If CDbl(txtField(4)) <= 0# Then
      MsgBox "Invalid Minimum Down Detected!" & vbCrLf & _
               "Pls Verify Entry Then Try Again!!!", vbCritical, "WARNING"
      txtField(4).SetFocus
      GoTo EntryNotOK
   End If
   
   If CDbl(txtField(5)) <= 0# Then
      MsgBox "Invalid Maximum Down Detected!" & vbCrLf & _
               "Pls Verify Entry Then Try Again!!!", vbCritical, "WARNING"
      txtField(5).SetFocus
      GoTo EntryNotOK
   End If
   
   lbWithTerm = False
   For lnCtr = 0 To 3
      If chkTerm(lnCtr).Value = Checked Then
         lbWithTerm = True
         Exit For
      End If
   Next
   
   If Not lbWithTerm Then
      MsgBox "Invalid Term Name Detected!" & vbCrLf & _
               "Pls Verify Entry Then Try Again!!!", vbCritical, "WARNING"
      GoTo EntryNotOK
   End If
   
   With MSFlexGrid1
      If .TextMatrix(1, 1) = "" Then
         MsgBox "No Item Entry Detected!" & vbCrLf & _
                    "Pls Verify Entry Then Try Again!!!", vbCritical, "WARNING"
         GoTo EntryNotOK
      End If
   End With

EntryOK:
   isEntryOk = True
   Exit Function
EntryNotOK:
   isEntryOk = False
End Function

Private Sub ClearFields()
   Dim loTxt As TextBox
   
   For Each loTxt In txtField
      loTxt = ""
   Next
   
   For Each loTxt In txtDetail
      loTxt = ""
   Next
   
   Combo1.ListIndex = 0
   
   pMIndex = -1
   pDIndex = -1
   
   With MSFlexGrid1
      .Rows = 2
      .TextMatrix(1, 0) = 1
      .TextMatrix(1, 1) = ""
      .TextMatrix(1, 2) = ""
      .TextMatrix(1, 3) = "0.00"
      
      .Row = 1
      .Col = 1
      .ColSel = .Cols - 1
      
      pnActiveRow = .Row
   End With
   chkZeroInt.Value = False
   chkTerm(0).Value = False
   chkTerm(1).Value = False
   chkTerm(2).Value = False
   chkTerm(3).Value = False
End Sub

Private Sub initButton(lnStat As Integer)
   Dim lbShow As Boolean

   lbShow = IIf(lnStat = 0, False, True)
   cmdButton(4).Visible = Not lbShow
   cmdButton(5).Visible = Not lbShow
   cmdButton(6).Visible = Not lbShow
   cmdButton(7).Visible = Not lbShow

   cmdButton(0).Visible = lbShow
   cmdButton(3).Visible = lbShow
   
   txtField(1).Enabled = lbShow
   xrFrame(0).Enabled = lbShow
   xrFrame(1).Enabled = lbShow
'   If lbShow Then cmdButton(0).SetFocus
End Sub


Private Sub Combo1_Click()
   oTrans.Master("cProdctTp") = CStr(Combo1.ListIndex) + 1
End Sub

Private Sub Form_Activate()
   Dim lsOldProc As String

   lsOldProc = "Form_Activate"
   ''' 'On Error GoTo errProc

   oApp.MenuName = Me.Tag
   Me.ZOrder 0
   
   If Not pbFormLoad Then
      pbFormLoad = True
    End If
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
   End Select
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String

   lsOldProc = "Form_Load"
   '' 'On Error GoTo errProc

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualLeft
   
   CenterChildForm mdiMain, Me

   Set oTrans = New clsCPPromo
   Set oTrans.AppDriver = oApp
   
   Call InitGrid
   oTrans.InitTransaction
   oTrans.NewTransaction
   oTrans.ProductType = "1"
   
   ClearFields
   LoadRecord
   initButton oTrans.EditMode

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc
End Sub

Private Sub MSFlexGrid1_RowColChange()
   If Not pbFormLoad Then Exit Sub
   pnActiveRow = MSFlexGrid1.Row
   Call showdetail
End Sub

Private Sub showdetail()
   With MSFlexGrid1
      txtDetail(0) = oTrans.Detail(pnActiveRow - 1, "sBrandNme")
      txtDetail(1) = oTrans.Detail(pnActiveRow - 1, "sModelNme")
      txtDetail(2) = oTrans.Detail(pnActiveRow - 1, "sModelCde")
      txtDetail(3) = Format(oTrans.Detail(pnActiveRow - 1, "nSelPrice"), "#,##0.00")
   End With
End Sub

Private Sub MSFlexGrid2_DblClick()
   Dim lbContinue As Boolean
   
   With MSFlexGrid2
      lbContinue = True
      If oTrans.EditMode = xeModeAddNew Or _
         oTrans.EditMode = xeModeAddNew Then
         
         lbContinue = MsgBox("Loading other record will disregard changes made." & vbCrLf & vbCrLf & _
                              "Do you want to continue?", vbQuestion + vbYesNo, "Confirm")
      End If
      
      If lbContinue Then
         If oTrans.OpenTransaction(.TextMatrix(.Row, 1)) Then
            LoadMaster
            LoadDetail
            initButton oTrans.EditMode
         End If
      End If
   End With
End Sub

Private Sub oTrans_DetailRetrieved(ByVal Row As Integer, ByVal Index As Integer)
   With MSFlexGrid1
      Select Case Index
      Case 0
         txtDetail(Index).Text = oTrans.Detail(Row, Index)
      Case 3
         txtDetail(Index).Text = Format(oTrans.Detail(Row, Index), "#,##0.00")
         .TextMatrix(pnActiveRow, Index) = txtDetail(Index).Text
      Case Else
         txtDetail(Index).Text = oTrans.Detail(Row, Index)
         .TextMatrix(pnActiveRow, Index) = txtDetail(Index).Text
      End Select
   End With
End Sub

Private Sub oTrans_MasterRetrieved(ByVal Index As Integer)
   Select Case Index
   Case 13
      chkTerm(0).Value = oTrans.Master(Index)
   Case 14
      chkTerm(1).Value = oTrans.Master(Index)
   Case 15
      chkTerm(2).Value = oTrans.Master(Index)
   Case 16
      chkTerm(3).Value = oTrans.Master(Index)
   Case 17
      chkZeroInt.Value = oTrans.Master(Index)
   End Select
End Sub

Private Sub txtDetail_GotFocus(Index As Integer)
   With txtDetail(Index)
      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = oApp.getColor("HT1")
   End With
   
   pbDetailGotFocus = True
   pbMasterGotFocus = False
   pDIndex = Index
End Sub

Private Sub txtDetail_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String

   lsOldProc = "txtDetail_KeyDown"
   '''On Error GoTo errProc

   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
      With txtDetail(Index)
         If KeyCode = vbKeyF3 Then
            oTrans.searchDetail pnActiveRow - 1, Index, .Text
            If .Text <> "" Then SetNextFocus
         Else
            If .Text <> "" Then oTrans.searchDetail pnActiveRow - 1, Index, .Text
         End If
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

Private Sub txtDetail_LostFocus(Index As Integer)
   Call HighlightOff(Me.txtDetail(Index))
End Sub

Private Sub txtDetail_Validate(Index As Integer, Cancel As Boolean)
   With txtDetail(Index)
      oTrans.Detail(pnActiveRow - 1, Index) = .Text
   End With
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
       .BackColor = oApp.getColor("HT1")
       .SelStart = 0
       .SelLength = Len(.Text)
   End With
   
   pMIndex = Index
   pbMasterGotFocus = True
   pbDetailGotFocus = False
End Sub

Private Sub LoadRecord()
   Dim lnCtr As Integer
   Dim lors As Recordset
   
   Set lors = oTrans.oRSMaster
   With MSFlexGrid2
      .Rows = lors.RecordCount + 1
      For lnCtr = 0 To oTrans.oRSMaster.RecordCount - 1
         .TextMatrix(lnCtr + 1, 0) = lnCtr + 1
         .TextMatrix(lnCtr + 1, 1) = lors("sMPCatIDx")
         .TextMatrix(lnCtr + 1, 2) = lors("sMPCatNme")
         .TextMatrix(lnCtr + 1, 3) = Format(lors("dDateFrom"), "YYYY/MM/DD")
         .TextMatrix(lnCtr + 1, 4) = Format(lors("dDateThru"), "YYYY/MM/DD")
         lors.MoveNext
      Next
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oTrans = Nothing
   Set oSkin = Nothing

   pbFormLoad = False
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

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
   
   pbMasterGotFocus = False
End Sub

Private Sub LoadMaster()
   Dim lnCtr As Integer
   
   With oTrans
      For lnCtr = 0 To txtField.Count - 1
         Select Case lnCtr
         Case 2, 3
            txtField(lnCtr) = Format(.Master(lnCtr), "MMMM DD, YYYY")
         Case 4, 5
            txtField(lnCtr) = Format(.Master(lnCtr), "0.00")
         Case 6 To 12
            txtField(lnCtr) = Format(.Master(lnCtr), "#,##0.00")
         Case Else
            txtField(lnCtr) = .Master(lnCtr)
         End Select
      Next
      
      chkTerm(0).Value = oTrans.Master("n3Monthsx")
      chkTerm(1).Value = oTrans.Master("n6Monthsx")
      chkTerm(2).Value = oTrans.Master("n9Monthsx")
      chkTerm(3).Value = oTrans.Master("n12Months")
      chkZeroInt = oTrans.Master("nZeroIntx")
      
      Combo1.ListIndex = CInt(oTrans.Master("cProdctTp")) - 1
   End With
End Sub

Private Sub LoadDetail()
   Dim lnRow As Integer
   Dim lnCtr As Integer

   With MSFlexGrid1
      .Rows = oTrans.ItemCount + 1
      For lnCtr = 0 To oTrans.ItemCount - 1
         .TextMatrix(lnCtr + 1, 0) = lnCtr + 1
         .TextMatrix(lnCtr + 1, 1) = oTrans.Detail(lnCtr, "sModelNme")
         .TextMatrix(lnCtr + 1, 2) = oTrans.Detail(lnCtr, "sModelCde")
         .TextMatrix(lnCtr + 1, 3) = Format(oTrans.Detail(lnCtr, "nSelPrice"), "#,##0.00")
      Next
      
      .Row = 1
      .Col = 1
      .ColSel = .Cols - 1
      
      pnActiveRow = .Row
      showdetail
   End With
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   With txtField(Index)
      Select Case Index
      Case 2, 3
         If Not IsDate(.Text) Then .Text = oApp.ServerDate
         If Not oTrans.Master("sMPCatIDx") = "C001000001" Then
            If CDate(.Text) < oApp.ServerDate Then .Text = oApp.ServerDate
         End If
         
         oTrans.Master(Index) = CDate(.Text)
         
         .Text = Format(oTrans.Master(Index), "MMMM DD, YYYY")
      Case 4 To 9
         If Not IsNumeric(.Text) Then .Text = 0#
         oTrans.Master(Index) = CDbl(.Text)
         
         .Text = Format(oTrans.Master(Index), "#,##0.00")
      Case Else
         oTrans.Master(Index) = .Text
      End Select
   End With
End Sub
