VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0"; "xrGridControl.ocx"
Begin VB.Form frmCP_JobOrder 
   BorderStyle     =   0  'None
   Caption         =   "Job Order"
   ClientHeight    =   7800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10980
   DrawWidth       =   18832
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7800
   ScaleWidth      =   10980
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrButton cmdButton 
      Height          =   540
      Index           =   17
      Left            =   75
      TabIndex        =   69
      Top             =   4530
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   953
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
      Picture         =   "frmCP_JobOrder.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   540
      Index           =   5
      Left            =   75
      TabIndex        =   73
      Top             =   6240
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   953
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
      Picture         =   "frmCP_JobOrder.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   540
      Index           =   14
      Left            =   75
      TabIndex        =   57
      Top             =   540
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   953
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
      Picture         =   "frmCP_JobOrder.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   540
      Index           =   15
      Left            =   75
      TabIndex        =   66
      Top             =   2820
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   953
      Caption         =   "U. Labor"
      AccessKey       =   "U. Labor"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCP_JobOrder.frx":166E
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   540
      Index           =   11
      Left            =   75
      TabIndex        =   68
      Top             =   3960
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   953
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
      Picture         =   "frmCP_JobOrder.frx":1DE8
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   540
      Index           =   6
      Left            =   75
      TabIndex        =   64
      Top             =   1680
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   953
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
      Picture         =   "frmCP_JobOrder.frx":2562
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   540
      Index           =   10
      Left            =   75
      TabIndex        =   65
      Top             =   2250
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   953
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
      Picture         =   "frmCP_JobOrder.frx":2CDC
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   540
      Index           =   7
      Left            =   75
      TabIndex        =   67
      Top             =   3390
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   953
      Caption         =   "&Parts"
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
      Picture         =   "frmCP_JobOrder.frx":3456
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   540
      Index           =   8
      Left            =   75
      TabIndex        =   72
      Top             =   5670
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   953
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
      Picture         =   "frmCP_JobOrder.frx":3BD0
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   540
      Index           =   4
      Left            =   75
      TabIndex        =   74
      Top             =   6810
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   953
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
      Picture         =   "frmCP_JobOrder.frx":434A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   540
      Index           =   3
      Left            =   75
      TabIndex        =   70
      Top             =   5100
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   953
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
      Picture         =   "frmCP_JobOrder.frx":4AC4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   540
      Index           =   2
      Left            =   75
      TabIndex        =   63
      Top             =   3960
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   953
      Caption         =   "Pa&y"
      AccessKey       =   "y"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCP_JobOrder.frx":523E
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   540
      Index           =   0
      Left            =   75
      TabIndex        =   58
      Top             =   1110
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   953
      Caption         =   "Repaired"
      AccessKey       =   "Repaired"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCP_JobOrder.frx":6670
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   540
      Index           =   1
      Left            =   75
      TabIndex        =   59
      Top             =   1680
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   953
      Caption         =   "Released"
      AccessKey       =   "Released"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCP_JobOrder.frx":6DEA
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   4860
      Left            =   1575
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   8573
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.CheckBox chkBackJob 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Back Job (J.O. #)"
         Height          =   195
         Left            =   7455
         TabIndex        =   31
         Tag             =   "et0;fb0"
         Top             =   1830
         Width           =   1665
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   16
         Left            =   5160
         TabIndex        =   28
         Top             =   1425
         Width           =   4020
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   24
         Left            =   8580
         TabIndex        =   44
         Top             =   3435
         Width           =   600
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   22
         Left            =   8580
         TabIndex        =   41
         Top             =   3105
         Width           =   600
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   20
         Left            =   8580
         TabIndex        =   38
         Top             =   2745
         Width           =   600
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   18
         Left            =   8580
         TabIndex        =   35
         Top             =   2415
         Width           =   600
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   19
         Left            =   5160
         TabIndex        =   37
         Top             =   2745
         Width           =   3405
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   17
         Left            =   5160
         TabIndex        =   34
         Top             =   2415
         Width           =   3405
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   23
         Left            =   5160
         TabIndex        =   43
         Top             =   3435
         Width           =   3405
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   21
         Left            =   5160
         TabIndex        =   40
         Top             =   3105
         Width           =   3405
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   450
         Index           =   6
         Left            =   1005
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   1290
         Width           =   3300
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   5
         Left            =   1005
         TabIndex        =   7
         Top             =   960
         Width           =   3300
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   9
         Left            =   1005
         TabIndex        =   16
         Top             =   2685
         Width           =   1440
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   13
         Left            =   7470
         TabIndex        =   32
         Top             =   2040
         Width           =   1665
      End
      Begin VB.OptionButton chkServiceType 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Under Limited Warranty"
         Height          =   195
         Index           =   1
         Left            =   5220
         TabIndex        =   30
         Tag             =   "et0;fb0"
         Top             =   2115
         Width           =   1965
      End
      Begin VB.OptionButton chkServiceType 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Void Warranty"
         Height          =   195
         Index           =   0
         Left            =   5220
         TabIndex        =   29
         Tag             =   "et0;fb0"
         Top             =   1830
         Width           =   1440
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   450
         Index           =   15
         Left            =   1005
         MaxLength       =   512
         MultiLine       =   -1  'True
         TabIndex        =   48
         Top             =   4275
         Width           =   8175
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   11
         Left            =   3075
         TabIndex        =   20
         Top             =   3015
         Width           =   1170
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   10
         Left            =   1005
         TabIndex        =   18
         Top             =   3015
         Width           =   1440
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   12
         Left            =   1005
         MaxLength       =   25
         TabIndex        =   22
         Top             =   3345
         Width           =   3240
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   3
         Left            =   5160
         TabIndex        =   24
         Top             =   630
         Width           =   4020
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   450
         Index           =   4
         Left            =   5160
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   26
         Top             =   960
         Width           =   4020
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   7
         Left            =   1005
         MaxLength       =   25
         TabIndex        =   12
         Top             =   2025
         Width           =   3240
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   8
         Left            =   1005
         TabIndex        =   14
         Top             =   2355
         Width           =   1440
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   1005
         TabIndex        =   3
         Text            =   "DEC-01-2010"
         Top             =   630
         Width           =   1620
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   3225
         MaxLength       =   10
         TabIndex        =   5
         Top             =   630
         Width           =   1080
      End
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
         TabIndex        =   1
         Top             =   150
         Width           =   1620
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   450
         Index           =   14
         Left            =   1005
         MaxLength       =   512
         MultiLine       =   -1  'True
         TabIndex        =   46
         Top             =   3810
         Width           =   8175
      End
      Begin VB.Shape Shape6 
         Height          =   645
         Left            =   5160
         Top             =   1755
         Width           =   4020
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "UNIT INFORMATION"
         Height          =   195
         Index           =   8
         Left            =   240
         TabIndex        =   10
         Tag             =   "et0;fb0"
         Top             =   1770
         Width           =   1560
      End
      Begin VB.Shape Shape5 
         Height          =   1920
         Left            =   150
         Top             =   1830
         Width           =   4155
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
         TabIndex        =   75
         Tag             =   "eb0;et0"
         Top             =   195
         Width           =   2385
      End
      Begin VB.Shape Shape4 
         Height          =   330
         Index           =   0
         Left            =   6705
         Top             =   150
         Width           =   2445
      End
      Begin VB.Shape Shape3 
         Height          =   390
         Index           =   0
         Left            =   6675
         Top             =   120
         Width           =   2505
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Technician"
         Height          =   195
         Index           =   21
         Left            =   4320
         TabIndex        =   27
         Top             =   1455
         Width           =   795
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Repair"
         Height          =   195
         Index           =   20
         Left            =   4665
         TabIndex        =   42
         Top             =   3495
         Width           =   465
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Defect"
         Height          =   195
         Index           =   19
         Left            =   4470
         TabIndex        =   39
         Top             =   3165
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
         TabIndex        =   36
         Top             =   2820
         Width           =   660
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Condition"
         Height          =   195
         Index           =   16
         Left            =   4470
         TabIndex        =   33
         Top             =   2475
         Width           =   660
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   5160
         X2              =   9165
         Y1              =   3075
         Y2              =   3075
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   195
         Index           =   6
         Left            =   375
         TabIndex        =   8
         Top             =   1320
         Width           =   570
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ASC Name"
         Height          =   195
         Index           =   4
         Left            =   165
         TabIndex        =   6
         Top             =   975
         Width           =   780
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
         Height          =   195
         Index           =   15
         Left            =   450
         TabIndex        =   15
         Top             =   2730
         Width           =   480
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   195
         Index           =   3
         Left            =   285
         TabIndex        =   47
         Top             =   4260
         Width           =   630
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ref No."
         Height          =   195
         Index           =   14
         Left            =   2490
         TabIndex        =   19
         Top             =   3105
         Width           =   555
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DOP"
         Height          =   195
         Index           =   12
         Left            =   450
         TabIndex        =   17
         Top             =   3060
         Width           =   480
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dealer"
         Height          =   195
         Index           =   13
         Left            =   450
         TabIndex        =   21
         Top             =   3390
         Width           =   480
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
         Height          =   195
         Index           =   11
         Left            =   4485
         TabIndex        =   23
         Top             =   645
         Width           =   645
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cust. Add."
         Height          =   195
         Index           =   10
         Left            =   4395
         TabIndex        =   25
         Top             =   990
         Width           =   735
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IMEI"
         Height          =   195
         Index           =   7
         Left            =   450
         TabIndex        =   11
         Top             =   2070
         Width           =   480
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Accessory"
         Height          =   195
         Index           =   2
         Left            =   165
         TabIndex        =   45
         Top             =   3795
         Width           =   750
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
         TabIndex        =   4
         Top             =   690
         Width           =   480
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Brand"
         Height          =   195
         Index           =   5
         Left            =   450
         TabIndex        =   13
         Top             =   2415
         Width           =   480
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trans. Date"
         Height          =   195
         Index           =   1
         Left            =   105
         TabIndex        =   2
         Top             =   675
         Width           =   840
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
         TabIndex        =   0
         Top             =   210
         Width           =   735
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
   Begin xrControl.xrFrame xrFrame2 
      Height          =   2265
      Left            =   1575
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
         TabIndex        =   51
         Text            =   "0,000.00"
         Top             =   1605
         Width           =   810
      End
      Begin xrGridEditor.GridEditor GridEditor1 
         Height          =   1500
         Left            =   75
         TabIndex        =   49
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
         MOUSEICON       =   "frmCP_JobOrder.frx":7564
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
         TabIndex        =   50
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
         TabIndex        =   54
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
         TabIndex        =   52
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
         TabIndex        =   55
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
         TabIndex        =   53
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
         TabIndex        =   56
         Tag             =   "ht0;hb0"
         Top             =   1605
         Width           =   1485
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   540
      Index           =   9
      Left            =   75
      TabIndex        =   60
      Top             =   2250
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   953
      Caption         =   "Forward"
      AccessKey       =   "Forward"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCP_JobOrder.frx":7580
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   540
      Index           =   12
      Left            =   75
      TabIndex        =   71
      Top             =   5670
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   953
      Caption         =   "Back&Out"
      AccessKey       =   "O"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCP_JobOrder.frx":7CFA
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   540
      Index           =   13
      Left            =   75
      TabIndex        =   61
      Top             =   2820
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   953
      Caption         =   "Received"
      AccessKey       =   "Received"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCP_JobOrder.frx":8474
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   540
      Index           =   16
      Left            =   75
      TabIndex        =   62
      Top             =   3390
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   953
      Caption         =   "Replaced"
      AccessKey       =   "Replaced"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCP_JobOrder.frx":8BEE
   End
End
Attribute VB_Name = "frmCP_JobOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmCP_JobOrder"

Private WithEvents oTrans As clsJobOrder
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin

Dim pbLoadRecord As Boolean
Dim pbGridFocus As Boolean, pbSearch As Boolean
Dim pnCtr As Integer, pnIndex As Integer
Dim pbIsSrvcCenter As Boolean

Private Sub chkBackJob_Click()
   With chkBackJob
      If .Value = Unchecked Then
         txtField(13).Text = ""
         oTrans.Master("sPrevJONo") = ""
      End If
   End With
   oTrans.Master("cBackJobx") = chkBackJob.Value
End Sub

Private Sub chkServiceType_Click(Index As Integer)
   Dim lsAppvID As String, lsAppvName As String
   Dim lnAppvRights As Integer

   If oTrans.EditMode = xeModeUpdate Then
      If GetApproval(oApp, _
                     lnAppvRights, _
                     lsAppvID, _
                     lsAppvName, _
                     oApp.MenuName) = False Then Exit Sub
                     
      If lnAppvRights < xeManager Then
            MsgBox "User is not allowed to update JobOrder Type!!!" & vbCrLf & _
                     "Request can not be granted!!!", vbCritical, "Warning"
            Exit Sub
      End If
   End If
   
   oTrans.Master("cJOTypexx") = Index
End Sub

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   Dim lnRep As Integer
   Dim ldDate As Date

   lsOldProc = "cmdButton_Click"
   ''On Error GoTo errProc

   txtField_LostFocus pnIndex
   With GridEditor1
      Select Case Index
      Case 0 ' Repaired
         If Not pbLoadRecord Then Exit Sub
         If oTrans.Master("cTranStat") = xeJOStateJobOrder Or _
            oTrans.Master("cTranStat") = xeJOStateOpen Or _
            oTrans.Master("cTranStat") = xeJOStateForwarded Then
            lnRep = MsgBox("Are you sure you want repair transaction!!!", vbYesNo + vbQuestion)
            If lnRep = vbNo Then Exit Sub
            If inputDate("Date Repaired", ldDate) = False Then Exit Sub
            
            If oTrans.Repaired(ldDate) Then
               MsgBox "Transaction Updated Successfully!!!", vbInformation, "Confirm"
               Call oTrans.OpenTransaction(oTrans.Master("sTransNox"))
               Call LoadMaster
            Else
               MsgBox "Unable to update Transaction!!!", vbCritical, "Warning"
            End If
         End If
      Case 1 ' Released
         If Not pbLoadRecord Then Exit Sub
         If oTrans.Master("cTranStat") = xeJOStateJobOrder Or _
            oTrans.Master("cTranStat") = xeJOStateOpen Or _
            oTrans.Master("cTranStat") = xeJOStateForwarded Or _
            oTrans.Master("cTranStat") = xeJOStateRepaired Then
            lnRep = MsgBox("Are you sure you want release transaction!!!", vbYesNo + vbQuestion)
            If lnRep = vbNo Then Exit Sub
            If oTrans.Master("sReleased") <> "" And oTrans.Master("cTranStat") <> xeJOStateJobOrder Then
               MsgBox "Transaction Already Released!!!", vbInformation, "Confirm"
               GoTo endProc
            End If
            
            If inputDate("Date Released", ldDate) = False Then Exit Sub
            If oTrans.Released(oTrans.Master("sTransNox"), ldDate, oApp.UserID, oTrans.Master("sClientID")) Then
               MsgBox "Transaction Updated Successfully!!!", vbInformation, "Confirm"
               ClearFields
            Else
               MsgBox "Unable to Released Transaction!!!", vbCritical, "Warning"
            End If
         End If
      Case 2 ' Pay Transaction
         If Not pbLoadRecord Then Exit Sub
         If oTrans.Master("cTranStat") = xeJOStateJobOrder Or _
            oTrans.Master("cTranStat") = xeJOStateForwarded Or _
            oTrans.Master("cTranStat") = xeJOStateRepaired Then
            lnRep = MsgBox("Are you sure you want pay transaction!!!", vbYesNo + vbQuestion)
            If lnRep = vbNo Then Exit Sub
            
            Dim loFormPayment As frmJOPayment
   
            Set loFormPayment = New frmJOPayment
            Set loFormPayment.AppDriver = oApp
            
            With loFormPayment
               .txtField(0).Text = lblTotalLabor.Caption
               .txtField(1).Text = lblTotalParts.Caption
               .txtField(2).Text = txtField(25).Text
               .txtField(3).Text = ""
               .txtField(4).Text = Format(oApp.ServerDate, "MMM-DD-YYYY")
               .lblTotal.Caption = lblTotal.Caption
               .Show 1
               
               If Not .Cancelled Then
                  lblTotalLabor.Caption = .Labor
                  lblTotalParts.Caption = .Parts
                  txtField(25).Text = .Others
                  lblTotal.Caption = .GrandTotal
                  
                  If oTrans.PayTransaction(oTrans.Master("sTransNox"), _
                                             .DatePayment, _
                                             .SalesInvoice, _
                                             .Labor, _
                                             .Parts, _
                                             .Others, _
                                             .GrandTotal) Then
                     'for S.I printing
                     If lnRep = vbYes Then
                        Call PrintTrans
                        Call UpdateSI
                     End If
                     ClearFields
                  Else
                     MsgBox "Unable to Pay Transaction!!!", vbCritical, "Warning"
                  End If
               End If
            End With
            
            Set loFormPayment = Nothing
         End If
      Case 3 ' Browse
         If oTrans.SearchTransaction() Then
            LoadMaster
            LoadDetail
   
            If cmdButton(4).Visible = False Then
               initButton xeModeReady
               cmdButton(4).SetFocus
            End If
         End If
      Case 4 ' Closed
         Unload Me
      Case 5 ' Update
         If Not pbLoadRecord Then Exit Sub
         If oTrans.Master("cTranStat") = xeJOStateJobOrder Or _
            oTrans.Master("cTranStat") = xeJOStateForwarded Or _
            oTrans.Master("cTranStat") = xeJOStateOpen Then
            If oTrans.UpdateTransaction Then
               initButton xeModeAddNew
               txtField(1).SetFocus
            Else
               MsgBox "Unable to Update Transaction!!!", vbCritical, "Warning"
            End If
         End If
      Case 6 'Save
         If .Rows > 2 Then
            pnCtr = 0
            Do While pnCtr < .Rows
               If Trim(.TextMatrix(pnCtr, 1)) = "" Then
                  .Row = pnCtr
                  If oTrans.deleteDetail(.Row - 1) Then .deleteRow
               Else
                  pnCtr = pnCtr + 1
               End If
            Loop
   
            .ColWidth(2) = 3100
            If .Rows > 6 Then .ColWidth(2) = 2900
         End If
      
         If isEntryOk Then
            If oTrans.SaveTransaction Then
               MsgBox "Record Updated Successfully!!!", vbInformation, "Notice"
               pbLoadRecord = True
               initButton xeModeReady
            Else
               MsgBox "Unable to Update Record!!!", vbCritical, "Warning"
            End If
         End If
      Case 7 ' Search Detail
         If pbGridFocus = False Then Exit Sub
         Select Case .Col
         Case 1, 2
            If oTrans.searchDetail(.Row - 1, .Col) Then
               DisplayComputation
               .Col = 4
            Else
               .Col = 1
            End If
            .SetFocus
            .Refresh
         End Select
      Case 8 ' Cancel Update
         lnRep = MsgBox("Transaction is in Update Mode!!!" & vbCrLf & _
                        "Do you want to Cancel Update!!!", vbYesNo + vbQuestion, "Confirm")
         If lnRep = vbYes Then
            initButton xeModeReady
            If pbLoadRecord Then
               oTrans.OpenTransaction oTrans.Master("sTransNox")
               LoadMaster
               LoadDetail
            Else
               ClearFields
            End If
            cmdButton(4).SetFocus
         Else
            txtField(pnIndex).SetFocus
         End If
      Case 9 ' Forwarded
         If Not pbLoadRecord Then Exit Sub
         If oTrans.Master("cTranStat") = xeJOStateOpen Or _
            oTrans.Master("cTranStat") = xeJOStateOpen Or _
            oTrans.Master("cTranStat") = xeJOStateJobOrder Then
            lnRep = MsgBox("Are you sure you want forward transaction!!!", vbYesNo + vbQuestion)
            If lnRep = vbNo Then Exit Sub
            If inputDate("Date Forwarded", ldDate) = False Then Exit Sub
            If oTrans.Forwarded(ldDate) Then
               MsgBox "Transaction Updated Successfully!!!", vbInformation, "Innformation"
               Call oTrans.OpenTransaction(oTrans.Master("sTransNox"))
               Call LoadMaster
            End If
         End If
      Case 10 ' Search Master
         If Not pbGridFocus Then
            oTrans.SearchMaster pnIndex, txtField(pnIndex).Text
            txtField(pnIndex).SetFocus
         End If
         .Refresh
      Case 11 ' Deleta Row
         If pbIsSrvcCenter = False Then GoTo endProc
         If .Rows = 2 Then
            If oTrans.deleteDetail(.Row - 1) Then
               .TextMatrix(1, 1) = ""
               .TextMatrix(1, 2) = ""
               .TextMatrix(1, 3) = 0
               .TextMatrix(1, 4) = 0#
               .TextMatrix(1, 5) = 0
               .TextMatrix(1, 6) = 0 & "%"
               .TextMatrix(1, 7) = 0#
            End If
         Else
            If oTrans.deleteDetail(.Row - 1) Then .deleteRow
         End If
         
         .ColWidth(2) = 3100
         If .Rows > 6 Then .ColWidth(2) = 2900
         DisplayComputation
      Case 12 ' Backout
         If pbLoadRecord Then
            If oTrans.Master("cTranStat") = xeJOStateForwarded Or _
               oTrans.Master("cTranStat") = xeJOStateJobOrder Or _
               oTrans.Master("cTranStat") = xeJOStateOpen Then
               lnRep = MsgBox("Are you sure you want backout transaction!!!", vbYesNo + vbQuestion)
               If lnRep = vbNo Then Exit Sub
               If oTrans.CancelTransaction Then
                  oTrans.NewTransaction
                  initButton xeModeAddNew
                  ClearFields
                  txtField(1).SetFocus
               Else
                  MsgBox "Unable to BackOut Transaction!!!", vbCritical, "Warning"
               End If
            End If
         Else
            MsgBox "Unable to BackOut Transaction!!!" & vbCrLf & _
                   "No Transaction is Loaded!!!", vbCritical, "Warning"
         End If
      Case 13 ' Received
         If Not pbLoadRecord Then Exit Sub
         If (oTrans.Master("cTranStat") = xeJOStateForwarded Or _
            oTrans.Master("cTranStat") = xeJOStateRepaired) And _
            oTrans.Master("sLocation") <> "" Then
            lnRep = MsgBox("Are you sure you want receive transaction!!!", vbYesNo + vbQuestion)
            If lnRep = vbNo Then Exit Sub
            If inputDate("Date Received", ldDate) = False Then Exit Sub
            If oTrans.Received(ldDate) Then
               MsgBox "Transaction Updated Successfully!!!", vbInformation, "Confirm"
               Call oTrans.OpenTransaction(oTrans.Master("sTransNox"))
               Call LoadMaster
            Else
               MsgBox "Unable to update Transaction!!!", vbCritical, "Warning"
            End If
         End If
      Case 14 ' New
         oTrans.NewTransaction
         initButton xeModeAddNew
         ClearFields
   
         txtField(1).SetFocus
      Case 15 ' Update Labor
         Dim loFormLabor As frmLaborPrice
         
         Set loFormLabor = New frmLaborPrice
         Set loFormLabor.AppDriver = oApp
         
         loFormLabor.txtField(0).Text = oTrans.Master("sLaborIDx")
         loFormLabor.txtField(1).Text = oTrans.Master("sLaborNme")
         loFormLabor.txtField(2).Text = oTrans.Master("sLaborCde")
         loFormLabor.txtField(3).Text = oTrans.Master("xLborBrnd")
         loFormLabor.txtField(4).Text = Format(oTrans.Master("nLaborAmt"), "#,##0.00")
         
         loFormLabor.Show 1
         If loFormLabor.Cancelled = False Then
            oTrans.Master("nLaborAmt") = CDbl(loFormLabor.txtField(4).Text)
            lblTotalLabor.Caption = loFormLabor.txtField(4).Text
            Call DisplayComputation
         End If
         
         Set loFormLabor = Nothing
      Case 16 ' Replaced
         If Not pbLoadRecord Then Exit Sub
         If oTrans.Master("cTranStat") = xeJOStateJobOrder Or _
            oTrans.Master("cTranStat") = xeJOStateForwarded Or _
            oTrans.Master("cTranStat") = xeJOStateRepaired Or _
            oTrans.Master("cTranStat") = xeJOStateOpen Then
            
            lnRep = MsgBox("Are you sure you want replace transaction!!!", vbYesNo + vbQuestion)
            If lnRep = vbNo Then Exit Sub
            If inputDate("Date Replaced", ldDate) = False Then Exit Sub
            If oTrans.Replaced(ldDate) Then
               MsgBox "Transaction Updated Successfully!!!", vbInformation, "Confirm"
               Call oTrans.OpenTransaction(oTrans.Master("sTransNox"))
               Call LoadMaster
            Else
               MsgBox "Unable to update Transaction!!!", vbCritical, "Warning"
            End If
         End If
      Case 17 ' Ledger
         If Not pbLoadRecord Then Exit Sub
         With frmCP_JOMovementLedger
            .txtField(0) = txtField(0)
            .txtField(1) = txtField(7)
            .txtField(2) = txtField(8)
            .txtField(3) = txtField(9)
            
            .TransNox = oTrans.Master("sTransNox")
            .Show 1
         End With
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

   oTrans.JOStatus = 10560
   oTrans.InitTransaction
   oTrans.NewTransaction

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualLeft
   
   txtField(25).Enabled = oApp.isMainOffice

   InitGrid
   ClearFields
   initButton xeModeReady
   
   pbIsSrvcCenter = BranchStatus(oApp.BranchCode, "cSrvcCntr = " & strParm(xeYes))
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oTrans = Nothing
   Set oSkin = Nothing
End Sub

Private Sub initButton(lnStat As Integer)
   Dim lbShow As Boolean

   lbShow = IIf(lnStat = 0, False, True)
   cmdButton(0).Visible = Not lbShow
   cmdButton(1).Visible = Not lbShow
   cmdButton(2).Visible = Not lbShow
   cmdButton(4).Visible = Not lbShow
   cmdButton(5).Visible = Not lbShow
   cmdButton(9).Visible = Not lbShow
   cmdButton(12).Visible = Not lbShow
   cmdButton(13).Visible = Not lbShow
   cmdButton(14).Visible = Not lbShow
   cmdButton(16).Visible = Not lbShow
   
   cmdButton(6).Visible = lbShow
   cmdButton(7).Visible = lbShow
   cmdButton(8).Visible = lbShow
   cmdButton(10).Visible = lbShow
   cmdButton(11).Visible = lbShow
   cmdButton(15).Visible = lbShow
   
   xrFrame1.Enabled = lbShow
   
   With GridEditor1
      If pbIsSrvcCenter Then
         .ColEnabled(1) = lbShow
         .ColEnabled(2) = lbShow
         .ColEnabled(4) = lbShow
         .ColEnabled(5) = lbShow
         .ColEnabled(6) = lbShow
      
         
         If lbShow Then .SetFocus
      Else
         .ColEnabled(1) = lbShow
         .ColEnabled(2) = lbShow
         .ColEnabled(4) = lbShow
         .ColEnabled(5) = lbShow
         .ColEnabled(6) = lbShow
      End If
   End With
End Sub

Private Sub GridEditor1_AddingRow(Cancel As Boolean)
   With GridEditor1
      If .TextMatrix(.Row, 1) = "" Then
         Cancel = True
      ElseIf .TextMatrix(.Row, 2) = "" Then
         Cancel = True
      ElseIf .TextMatrix(.Row, 5) = "0" Then
         Cancel = True
      End If
      If Not Cancel Then oTrans.addDetail

'      If .Rows > 12 Then .ColWidth(2) = 3940
      DisplayComputation
   End With
End Sub

Private Sub GridEditor1_EditorValidate(Cancel As Boolean)
   Dim lnPercent As Integer
   Dim lnDiscount As Variant
   Dim lsRep As String

   With GridEditor1
      Select Case .Col
      Case 4
         If Not IsNumeric(.TextMatrix(.Row, .Col)) Then .TextMatrix(.Row, .Col) = 0
      Case 5
         If oTrans.Detail(.Row - 1, "nQtyOnHnd") <= 0 Then
            If .TextMatrix(.Row, 1) <> "" Then
               If .TextMatrix(.Row, .Col) > 0 Then
                  MsgBox "No Stock is Currently Availble!!!", vbCritical, "Confirm"
                  .TextMatrix(.Row, .Col) = 0
               End If
            End If
         Else
            If CDbl(.TextMatrix(.Row, .Col)) > CDbl(.TextMatrix(.Row, 3)) Then .TextMatrix(.Row, .Col) = 0
         End If
         
         oTrans.Detail(.Row, "nQuantity") = CDbl(.TextMatrix(.Row, .Col))
      Case 6
         If Not IsNumeric(lnDiscount) Then
            .TextMatrix(.Row, .Col) = 0
         Else
            lnDiscount = .TextMatrix(.Row, .Col)
            lnPercent = InStr(lnDiscount, "%")
            If lnPercent > 0 Then lnDiscount = Left(lnDiscount, lnPercent - 1)

            If lnDiscount > 99 Then lnDiscount = 0
         End If
         .TextMatrix(.Row, .Col) = lnDiscount & "%"
      End Select

      If .Col = 6 Then
         oTrans.Detail(.Row - 1, .Col) = CDbl(lnDiscount)
      Else
         If .Col = 1 Or .Col = 2 Then
            If pbSearch = False Then oTrans.Detail(.Row - 1, .Col) = .TextMatrix(.Row, .Col)
         Else
            oTrans.Detail(.Row - 1, .Col) = CDbl(.TextMatrix(.Row, .Col))
         End If
      End If
      DisplayComputation
   End With
End Sub

Private Sub GridEditor1_GotFocus()
   With GridEditor1
      .EditorBackColor = oApp.getColor("HT1")
   End With
   pbGridFocus = True
End Sub

Private Sub GridEditor1_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim lsRep As String
   Dim lsOldProc As String
   
   lsOldProc = "GridEditor1_KeyDown"
   ''On Error GoTo errProc

   If cmdButton(6).Visible Then
      If KeyCode = vbKeyF3 Then
         With GridEditor1
            Select Case .Col
            Case 1, 2
               pbSearch = True
               If oTrans.searchDetail(.Row - 1, .Col, .TextMatrix(.Row, .Col)) Then
                  If oTrans.Detail(.Row - 1, 3) <= 0 Then
                     MsgBox "No Stock is Currently Availble!!!", vbCritical, "Warning"
                     .TextMatrix(.Row, .Col) = ""
                  End If
                  
                  If .TextMatrix(.Row, .Col) <> "" Then .Col = 4
               End If
               pbSearch = False
            End Select
            
            KeyCode = 0
            .SetFocus
            .Refresh
            
            DisplayComputation
         End With
      End If
   End If

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " _
                       & "  " & KeyCode _
                       & ", " & Shift _
                       & " )", True
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
      ElseIf Index = 13 Then
         If .Master(Index) = "" Then chkBackJob.Value = Unchecked
         chkServiceType(oTrans.Master("cJOTypexx")).Value = True
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
   chkBackJob.Value = oTrans.Master("cBackJobx")
   txtField(13).Enabled = oTrans.Master("cBackJobx") = xeYes
   If pbIsSrvcCenter Or oApp.isMainOffice Then oTrans.Master("cTranStat") = xeJOStateJobOrder
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = oApp.getColor("HT1")
   End With

   pnIndex = Index
   pbGridFocus = False
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

Private Function isEntryOk() As Boolean
   If Trim(txtField(7).Text) = "" Then
      MsgBox "Serial not found!!!", vbCritical, "Warning"
      txtField(7).SetFocus
      GoTo EntryNotOK
   End If

EntryOK:
   isEntryOk = True
   Exit Function
EntryNotOK:
   isEntryOk = False
End Function

Private Sub LoadMaster()
   For pnCtr = 0 To txtField.Count - 1
      Select Case pnCtr
      Case 0
         txtField(pnCtr).Text = Format(oTrans.Master(pnCtr), "@@@@-@@@@@@")
      Case 1, 10
         txtField(pnCtr).Text = Format(oTrans.Master(pnCtr), "MMM-DD-YYYY")
      Case Else
         txtField(pnCtr).Text = IIf(IsNull(oTrans.Master(pnCtr)), "", oTrans.Master(pnCtr))
      End Select
   Next
   
   lblTotal.Caption = Format(oTrans.Master("nTranTotl"), "#,##0.00")
   lblTotalParts.Caption = Format(oTrans.Master("nPartsAmt"), "#,##0.00")
   lblTotalLabor.Caption = Format(oTrans.Master("nLaborAmt"), "#,##0.00")
   Label2.Caption = JobOrderStatus(oTrans.Master("cTranStat"))
   pbLoadRecord = True
   
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
      End Select
      
      If Index = 25 Then
         oTrans.Master("nMiscChrg") = CDbl(.Text)
         Call DisplayComputation
      Else
         oTrans.Master(Index) = .Text
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

      .ColEnabled(3) = False
      .ColEnabled(7) = False
      .ColDefault(3) = 0
      .ColDefault(4) = "0.00"
      .ColDefault(5) = 0
      .ColDefault(6) = "0" & "%"
      .ColDefault(7) = "0.00"
      .ColDefault(8) = 0
      
      .ColAlignment(1) = 1
      
      For pnCtr = 3 To 5
         .ColNumberOnly(pnCtr) = True
      Next
      
      .ColEnabled(8) = False
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

Function PrintTrans() As Boolean
   Dim lrs As New ADODB.Recordset
   Dim loRecd As Recordset
   Dim lnCtr As Integer
   Dim lsOldProc As String
   Dim lsSQL As String
   
   lsOldProc = "printTrans"
   ''On Error GoTo errProc
   
   PrintTrans = False
   
   Set lrs = New ADODB.Recordset
   
   lrs.Fields.Append "nField01", adInteger, 3
   lrs.Fields.Append "sField01", adVarChar, 50
   lrs.Fields.Append "sField02", adVarChar, 60
   lrs.Fields.Append "sField03", adVarChar, 20
   lrs.Fields.Append "lField01", adCurrency
   lrs.Fields.Append "lField02", adCurrency
   lrs.Fields.Append "lField03", adCurrency
   lrs.Open
      
   lsSQL = "SELECT a.sTransNox" & _
            ", c.sBarrCode" & _
            ", d.sModelNme" & _
            ", b.nUnitPrce" & _
            ", b.nQuantity" & _
            ", c.sDescript" & _
            ", a.dTransact" & _
         " FROM CP_SO_Master a" & _
            " LEFT JOIN CP_SO_Detail b" & _
               " ON a.sTransNox = b.sTransNox" & _
            " LEFT JOIN CP_Inventory c" & _
               " ON b.sStockIDx = c.sStockIdx" & _
            " LEFT JOIN CP_Model d" & _
               " ON c.sModelIDx = d.sModelIdx" & _
            " LEFT JOIN CP_Brand e" & _
               " ON c.sBrandIDx = e.sBrandIDx" & _
         " WHERE a.sTransNox = " & strParm(oTrans.Master("sSalesTrn"))
         
   Set loRecd = New Recordset
   loRecd.Open lsSQL, oApp.Connection, , , adCmdText
   
      lrs.AddNew
         lrs("lField01").Value = oTrans.Master("nLaborAmt")
         lrs("lField02").Value = oTrans.Master("nPartsAmt")
         lrs("lField03").Value = oTrans.Master("nMiscChrg")
         
   ' assign important info to the report
   Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\JO_SI.rpt")
   oReport.DiscardSavedData
   oReport.FieldMappingType = crAutoFieldMapping
   oReport.Database.SetDataSource lrs
   
   With oTrans
      oReport.Sections("PH").ReportObjects("txtCustomer").SetText oTrans.Master("xFullName")
      oReport.Sections("PH").ReportObjects("txtDate").SetText Format(loRecd("dTransact"), "MMM-DD-YYYY")
      oReport.Sections("PH").ReportObjects("txtAddress").SetText oTrans.Master("xAddressx")
      oReport.Sections("PH").ReportObjects("txtTIN").SetText ""
      oReport.Sections("PH").ReportObjects("txtBusiness").SetText ""
      oReport.Sections("PH").ReportObjects("txtTerm").SetText ""
      oReport.Sections("PH").ReportObjects("txtPrepared").SetText oApp.UserName
      oReport.Sections("RF").ReportObjects("Labor").SetText "Labor"
      oReport.Sections("RF").ReportObjects("Parts").SetText "Parts"
      oReport.Sections("RF").ReportObjects("misc").SetText "Other Charges"
   End With
   
   oReport.PrintOutEx False, 1
   lrs.Close
   PrintTrans = True

endProc:
   Set oReport = Nothing
   Set lrs = Nothing
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"
End Function

Private Sub UpdateSI()
   Dim lorSI As Recordset
   Dim lsSQL As String
   Dim lnRow As Integer
   
   lsSQL = "SELECT sTransNox" & _
               ", cTranStat" & _
            " FROM CP_SO_Master" & _
            " WHERE sTransNox = " & strParm(oTrans.Master("sSalesTrn")) & _
            " AND cTranStat = " & strParm(xeStateOpen)
   
   Set lorSI = New Recordset
   lorSI.Open lsSQL, oApp.Connection, , , adCmdText
   
   If Not lorSI.EOF Then
      lsSQL = "UPDATE CP_SO_Master SET cTranStat = " & strParm(xeStateClosed) & _
               "WHERE sTransNox = " & strParm(lorSI("sTransNox"))
   End If
   
   lnRow = oApp.Execute(lsSQL, "CP_SO_Master", oApp.BranchCode)
   If lnRow <= 0 Then
      MsgBox "Unable to update Si..."
      Exit Sub
   End If
End Sub
