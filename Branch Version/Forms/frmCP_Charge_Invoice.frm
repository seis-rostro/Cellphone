VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmCP_Charge_Invoice 
   BorderStyle     =   0  'None
   Caption         =   "Charge Invoice"
   ClientHeight    =   9555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11745
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   142.825
   ScaleMode       =   0  'User
   ScaleWidth      =   100.902
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      Height          =   1440
      Left            =   1590
      ScaleHeight     =   1380
      ScaleWidth      =   9975
      TabIndex        =   47
      TabStop         =   0   'False
      Tag             =   "wt0;fb0"
      Top             =   4425
      Width           =   10035
      Begin VB.TextBox txtDetail 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   8175
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   945
         Width           =   1755
      End
      Begin VB.TextBox txtDetail 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   8175
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   555
         Width           =   1755
      End
      Begin VB.TextBox txtDetail 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   1
         Left            =   8175
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   75
         Width           =   1755
      End
      Begin VB.TextBox txtDetail 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   0
         Left            =   60
         TabIndex        =   29
         Top             =   75
         Width           =   6510
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Disc. Amt. (F11)"
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
         Left            =   6750
         TabIndex        =   35
         Top             =   1020
         Width           =   1380
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity  (F10)"
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
         Left            =   6855
         TabIndex        =   33
         Top             =   660
         Width           =   1275
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "&IMEI / BARCODE"
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
         Index           =   15
         Left            =   60
         TabIndex        =   30
         Top             =   570
         Width           =   6900
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&UNIT PRICE (F9)"
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
         Index           =   8
         Left            =   6660
         TabIndex        =   31
         Top             =   225
         Width           =   1470
      End
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   3555
      Left            =   1575
      Tag             =   "wt0;fb0"
      Top             =   5880
      Width           =   10065
      _ExtentX        =   17754
      _ExtentY        =   6271
      BackColor       =   14737632
      ClipControls    =   0   'False
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   13
         Left            =   6705
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   2955
         Width           =   3240
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   2820
         Left            =   90
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   90
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   4974
         _Version        =   393216
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&AMOUNT PAID (F12)"
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
         Index           =   16
         Left            =   4515
         TabIndex        =   38
         Top             =   3075
         Width           =   2160
      End
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   600
      Index           =   3
      Left            =   90
      TabIndex        =   43
      Top             =   2445
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
      Picture         =   "frmCP_Charge_Invoice.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   90
      TabIndex        =   40
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
      Picture         =   "frmCP_Charge_Invoice.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   90
      TabIndex        =   46
      Top             =   1815
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
      Picture         =   "frmCP_Charge_Invoice.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   90
      TabIndex        =   42
      Top             =   1815
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
      Picture         =   "frmCP_Charge_Invoice.frx":166E
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   90
      TabIndex        =   41
      Top             =   1185
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
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
      Picture         =   "frmCP_Charge_Invoice.frx":1DE8
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   90
      TabIndex        =   44
      Top             =   555
      Width           =   1245
      _ExtentX        =   2196
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
      Picture         =   "frmCP_Charge_Invoice.frx":2562
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   3855
      Left            =   1575
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   10065
      _ExtentX        =   17754
      _ExtentY        =   6800
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   7830
         TabIndex        =   20
         Top             =   2085
         Width           =   1995
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   11
         Left            =   7830
         TabIndex        =   26
         Top             =   3030
         Width           =   1995
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   10
         Left            =   7830
         TabIndex        =   24
         Top             =   2715
         Width           =   1995
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   12
         Left            =   7830
         TabIndex        =   28
         Top             =   3345
         Width           =   1995
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   8
         Left            =   1365
         TabIndex        =   10
         Top             =   2400
         Width           =   4950
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   9
         Left            =   7830
         TabIndex        =   22
         Top             =   2400
         Width           =   1995
      End
      Begin VB.CheckBox chkClientTp 
         Caption         =   "Company / Institution"
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
         Left            =   1365
         TabIndex        =   4
         Tag             =   "wt0;fb0"
         Top             =   1230
         Width           =   2415
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   7830
         TabIndex        =   18
         Top             =   1770
         Width           =   1995
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   4
         Left            =   1365
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1770
         Width           =   4950
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
         Height          =   285
         Index           =   0
         Left            =   1365
         TabIndex        =   1
         Top             =   255
         Width           =   2310
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   1365
         TabIndex        =   6
         Top             =   1455
         Width           =   4950
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Index           =   7
         Left            =   1365
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   2715
         Width           =   4950
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   1365
         MaxLength       =   50
         TabIndex        =   3
         Top             =   735
         Width           =   2310
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   7830
         TabIndex        =   16
         Top             =   1455
         Width           =   1995
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Amount"
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
         Index           =   13
         Left            =   6585
         TabIndex        =   13
         Top             =   255
         Width           =   2070
      End
      Begin VB.Label lblTrantotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Caption         =   "999,000.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   630
         Left            =   6585
         TabIndex        =   14
         Top             =   495
         Width           =   3240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Due Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   6570
         TabIndex        =   19
         Top             =   2130
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Limit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   6570
         TabIndex        =   25
         Top             =   3075
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Discount Amt"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   6570
         TabIndex        =   23
         Top             =   2760
         Width           =   945
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*PIC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   750
         TabIndex        =   9
         Top             =   2445
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Balance"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   6570
         TabIndex        =   27
         Top             =   3390
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Discount Rate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   6570
         TabIndex        =   21
         Top             =   2445
         Width           =   1020
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   495
         TabIndex        =   7
         Top             =   1815
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Term"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   14
         Left            =   6570
         TabIndex        =   17
         Top             =   1815
         Width           =   360
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   285
         Left            =   1455
         Tag             =   "et0;ht2"
         Top             =   375
         Width           =   2325
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trans. No."
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
         Left            =   150
         TabIndex        =   0
         Top             =   300
         Width           =   915
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   12
         Left            =   435
         TabIndex        =   11
         Top             =   3075
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   6570
         TabIndex        =   15
         Top             =   1500
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transact. Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   2
         Top             =   780
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   600
         TabIndex        =   5
         Top             =   1500
         Width           =   465
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   6
      Left            =   90
      TabIndex        =   45
      Top             =   1185
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
      Picture         =   "frmCP_Charge_Invoice.frx":2CDC
   End
End
Attribute VB_Name = "frmCP_Charge_Invoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmCP_Charge_Invoice"
Private Const pxeAPPNAME = "CP Charge Invoice"
Private WithEvents oTrans As ggcCPSales.clsCPChargeInvoice
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin

Dim psTransNox As String
Dim pnIndex As Integer
Dim pnCtr As Integer

Dim pbLoaded As Boolean
Dim pbSave As Boolean
Dim bLoadRecord As Boolean
Dim pbClosedTrans As Boolean
Dim pnRow As Integer

Private Sub chkClientTp_Click()
   oTrans.Master("cClientTp") = chkClientTp.Value
End Sub

Private Sub cmdButton_Click(Index As Integer)
Dim lnRep As Long
Dim lsOldProc As String
      
lsOldProc = pxeMODULENAME & ".cmdButton_Click"
''On Error GoTo errProc
      
Select Case Index
   Case 0 'save
      If Not isEntryOk Then Exit Sub
         If oTrans.SaveTransaction Then
            lnRep = MsgBox("Do you want to print this transaction?", vbQuestion & vbYesNo, pxeAPPNAME)
            If lnRep = vbYes Then
               If PrinTrans Then
                  If oTrans.CloseTransaction(psTransNox) Then MsgBox "Printing..."
               End If
            If MsgBox("Reprint?", vbQuestion & vbYesNo, pxeAPPNAME) = vbYes Then PrinTrans
         End If
         
               're-open the previous trasaction made for printing.
         If Not oTrans.OpenTransaction(psTransNox) Then
            MsgBox "Unable to open transaction.", vbCritical, pxeAPPNAME
         Else
            bLoadRecord = True
         End If
                  
         Call initButton
         cmdButton(6).SetFocus
      Else
         MsgBox "Unable to save transaction.", vbCritical, pxeAPPNAME
      End If
   Case 1 'search
      Select Case pnIndex
         Case 3, 5
         Call txtField_KeyDown(pnIndex, vbKeyF3, 0)
      End Select
   Case 2 'delrow
      If oTrans.deleteDetail(pnRow) Then
         Call refreshGrid
      End If
   Case 3 'cancel
      If oTrans.InitTransaction Then
         Call InitForm
         Call InitGrid
         cmdButton(4).SetFocus
         bLoadRecord = False
      End If
   Case 4 'new
      'Call oTrans.InitTransaction
      If oTrans.NewTransaction Then
         Call InitGrid
         Call InitForm
         Call InitEntry
      End If
   Case 5 'close
      Unload Me
   Case 6 'print
      If Not bLoadRecord Then
         MsgBox "Unable to Print Transaction.", vbCritical, pxeAPPNAME
         Exit Sub
      End If

      lnRep = MsgBox("Do you want to print this transaction?", vbQuestion & vbYesNo, pxeAPPNAME)
      If lnRep = vbYes Then
         If PrinTrans Then
            If oTrans.CloseTransaction(psTransNox) Then MsgBox "Printing..."
         End If
   
         If MsgBox("Reprint?", vbQuestion & vbYesNo, pxeAPPNAME) = vbYes Then PrinTrans
      End If
End Select

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Activate()
If Not pbLoaded Then pbLoaded = True
   
oApp.MenuName = Me.Tag
Me.ZOrder 0

bLoadRecord = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn, vbKeyUp, vbKeyDown
         Select Case KeyCode
            Case vbKeyReturn, vbKeyDown
               If GetFocus = txtDetail(0).hwnd Then Exit Sub
               SetNextFocus
            Case vbKeyUp
               SetPreviousFocus
         End Select
      Case vbKeyF9
         txtDetail(1).Enabled = True
         txtDetail(1).SetFocus
      Case vbKeyF10
         txtDetail(2).Enabled = True
         txtDetail(2).SetFocus
      Case vbKeyF11
         txtDetail(3).Enabled = True
         txtDetail(3).SetFocus
      Case vbKeyF12
      txtField(13).Enabled = True
      txtField(13).SetFocus
   End Select
End Sub

Private Sub Form_Load()
Dim lsOldProc As String

lsOldProc = pxeMODULENAME & ".Form_Load"
''On Error GoTo errProc

CenterChildForm mdiMain, Me

Set oSkin = New clsFormSkin
Set oSkin.AppDriver = oApp
Set oSkin.Form = Me
oSkin.ApplySkin xeFormTransEqualLeft
   
Set oTrans = New ggcCPSales.clsCPChargeInvoice
Set oTrans.AppDriver = oApp
oTrans.Branch = oApp.BranchCode
oTrans.InitTransaction
   
If oTrans.NewTransaction Then
   Call InitForm
   Call InitEntry
   Call InitGrid
End If
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set oSkin = Nothing
Set oTrans = Nothing
   
pbLoaded = False
End Sub

Private Sub InitEntry()
Dim lsOldProc As String
   
lsOldProc = pxeMODULENAME & ".InitEntry"
''On Error GoTo errProc
   
With oTrans
   psTransNox = .Master("sTransNox")
   txtField(0) = Format(.Master("sTransNox"), "@@@@@@-@@@@@@")
   txtField(1) = Format(.Master("dTransact"), "MMMM DD, YYYY")
   txtField(2) = .Master("sChrgeInv")
   txtField(3) = ""
   txtField(4) = ""
   txtField(5) = .Master("sTermName")
   txtField(6) = Format(.Master("dDueDatex"), "MMMM DD, YYYY")
   txtField(7) = .Master("sRemarksx")
   txtField(8) = .Master("sCPerson1")
   txtField(9) = Format(.Master("nDiscRate"), "##0.00 %")
   txtField(10) = Format(.Master("nDiscAmtx"), "#,##0.00")
   txtField(11) = Format(.Master("nCredLimt"), "#,##0.00")
   txtField(12) = Format(.Master("nABalance"), "#,##0.00")
   txtField(13) = Format(.Master("nAmtPaidx"), "#,##0.00")
      
   txtDetail(0) = ""
   txtDetail(1) = "0.00"
   txtDetail(2) = "0"
   txtDetail(3) = "0.00"
   txtDetail(1).Enabled = False
   txtDetail(2).Enabled = False
   txtDetail(3).Enabled = False
      
   lblTrantotal = Format(.Master("nTranTotl"), "#,##0.00")
   chkClientTp.Value = .Master("cClientTp")

   pnRow = 0
End With
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub InitForm()
Dim lsOldProc As String
   
lsOldProc = pxeMODULENAME & ".InitForm"
''On Error GoTo errProc
   
txtField(0) = ""
txtField(1) = ""
txtField(2) = ""
txtField(3) = ""
txtField(4) = ""
txtField(5) = ""
txtField(6) = ""
txtField(7) = ""
txtField(8) = ""
txtField(9) = Format(0#, "##0.00 %")
txtField(10) = Format(0#, "#,##0.00")
txtField(11) = Format(0#, "#,##0.00")
txtField(12) = Format(0#, "#,##0.00")
txtField(13) = Format(0#, "#,##0.00")
   
txtDetail(0) = ""
txtDetail(1) = "0.00"
txtDetail(2) = "0"
txtDetail(3) = "0.00"
txtDetail(1).Enabled = False
txtDetail(2).Enabled = False
txtDetail(3).Enabled = False
   
lblTrantotal = Format(0#, "#,##0.00")
chkClientTp.Value = 0
   
Call initButton
   
pnRow = 0
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub initButton()
Dim lsOldProc As String
   
lsOldProc = pxeMODULENAME & ".InitButton"
''On Error GoTo errProc

With oTrans
   cmdButton(0).Visible = .EditMode = xeModeAddNew
   cmdButton(1).Visible = .EditMode = xeModeAddNew
   cmdButton(2).Visible = .EditMode = xeModeAddNew
   cmdButton(3).Visible = .EditMode = xeModeAddNew
   cmdButton(4).Visible = .EditMode = xeModeReady
   cmdButton(5).Visible = .EditMode = xeModeReady
   cmdButton(6).Visible = .EditMode = xeModeReady
      
   xrFrame1.Enabled = .EditMode = xeModeAddNew
   xrFrame2.Enabled = .EditMode = xeModeAddNew
   Picture1.Enabled = .EditMode = xeModeAddNew
End With
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub InitGrid()
Dim lsOldProc As String
Dim lnCtr As Integer

lsOldProc = pxeMODULENAME & ".initGrid"
''On Error GoTo errProc
   
With MSFlexGrid1
   .Clear
   .Cols = 8
   .Rows = 2
      
   .TextMatrix(0, 0) = ""
   .TextMatrix(0, 1) = "IMEI/Barcode"
   .TextMatrix(0, 2) = "Description"
   .TextMatrix(0, 3) = "Qty."
   .TextMatrix(0, 4) = "Sel. Price"
   .TextMatrix(0, 5) = "Disc."
   .TextMatrix(0, 6) = "Dsc. Amt."
   .TextMatrix(0, 7) = "Total"
      
   .Row = 0
      
      'column alignment
   For lnCtr = 0 To .Cols - 1
      .Col = lnCtr
      .CellFontBold = True
      .CellAlignment = flexAlignCenterCenter
   Next
         
   .Row = 1
   .ColWidth(0) = "450"
   .ColWidth(1) = "1600"
   .ColWidth(2) = "2950"
   
   .Col = 0
   .ColSel = .Cols - 1
End With
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub MSFlexGrid1_RowColChange()
With MSFlexGrid1
   .Col = 0
   .ColSel = .Cols - 1
      
   If .Row >= 1 Then
      txtDetail(1) = .TextMatrix(.Row, 4)
      txtDetail(2) = .TextMatrix(.Row, 3)
      txtDetail(3) = .TextMatrix(.Row, 6)
      pnRow = .Row - 1
   Else
      pnRow = 0
   End If
End With
End Sub

Private Sub oTrans_DetailRetrieved(ByVal Index As Integer)
Dim lsOldProc As String
   
lsOldProc = pxeMODULENAME & ".oTrans_DetailRetreived"
''On Error GoTo errProc
   
With MSFlexGrid1
   Select Case Index
   Case 1, 2
      .TextMatrix(pnRow + 1, Index) = oTrans.Detail(pnRow, Index)
   Case 7
      .TextMatrix(pnRow + 1, 3) = oTrans.Detail(pnRow, Index)
      txtDetail(2) = .TextMatrix(pnRow + 1, 3)
      GoTo endGrandTotal
   Case 8
      .TextMatrix(pnRow + 1, 4) = Format(oTrans.Detail(pnRow, Index), "#,##0.00")
      txtDetail(1) = .TextMatrix(pnRow + 1, 4)
      GoTo endGrandTotal
'   Case 9
'      .TextMatrix(pnRow + 1, 5) = Format(oTrans.Detail(pnRow, Index), "##0.00") & "%"
'      txtDetail(2) = .TextMatrix(pnRow + 1, 5)
'      GoTo endGrandTotal
   Case 10
      .TextMatrix(pnRow + 1, 6) = Format(oTrans.Detail(pnRow, Index), "#,##0.00")
      txtDetail(3) = .TextMatrix(pnRow + 1, 6)
      GoTo endGrandTotal
   End Select
End With
   
endProc:
   Exit Sub
endGrandTotal:
With MSFlexGrid1
   If .TextMatrix(pnRow + 1, 6) <> "" Then .TextMatrix(pnRow + 1, 7) = Format(CDbl(.TextMatrix(pnRow + 1, 3)) * CDbl(.TextMatrix(pnRow + 1, 4)) - CDbl(.TextMatrix(pnRow + 1, 6)), "#,##0.00")
   Call GrandTotal
End With
   
GoTo endProc
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub oTrans_MasterRetrieved(ByVal Index As Integer)
Select Case Index
      Case 3
   txtField(Index) = oTrans.Master(Index)
Case 6
   txtField(Index) = Format(txtField(Index), "MMMM DD, YYYY")
Case 9
   txtField(Index) = oTrans.Master(Index) & "%"
Case Else
   txtField(Index) = oTrans.Master(Index)
End Select
End Sub

Private Sub txtDetail_GotFocus(Index As Integer)
With txtDetail(Index)
   If Index = 2 Then .Text = Format(.Text, "0")
      
   .SelStart = 0
   .SelLength = Len(.Text)
End With

pnIndex = Index
End Sub

Private Sub txtDetail_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lsValue As String
   Dim lsBarrCode As String
   Dim lsQty As String
   Dim lnCtr As Integer
   Dim lnQty As Integer
   Dim lbDuplicate As Boolean

   With txtDetail(Index)
      Select Case Index
            Case 0 'barcode/imei
         If txtDetail(Index) = "" Then Exit Sub
         Select Case KeyCode
               Case vbKeyReturn
            lsValue = Trim(Left(.Text, 4))
            lsBarrCode = .Text
            lnQty = 1
      
            For lnCtr = 1 To Len(lsValue)
               If LCase(Left(Right(lsValue, lnCtr), 1)) = "x" Then
                  lsQty = Left(lsValue, Len(Trim(lsValue)) - lnCtr)
                  If IsNumeric(lsQty) Then
                     lnQty = lsQty
                     If Right(.Text, 1) = "x" Then
                        lnQty = 1
                     Else
                        lsBarrCode = Right(.Text, Len(.Text) - (Len(lsQty) + 1))
                     End If
                  Else
                     lnQty = 1
                     lsBarrCode = .Text
                  End If
               End If
            Next
         
            With MSFlexGrid1
               For lnCtr = 1 To .Rows - 1
                  If Trim(LCase(lsBarrCode)) = Trim(LCase(.TextMatrix(lnCtr, 1))) Then
                     .TextMatrix(lnCtr, 3) = CDbl(.TextMatrix(lnCtr, 3)) + lnQty
                     .TextMatrix(.Row, 7) = Format(CDbl(.TextMatrix(.Row, 3)) * CDbl(.TextMatrix(.Row, 4)) - CDbl(.TextMatrix(.Row, 6)), "#,##0.00")
                     oTrans.Detail(lnCtr - 1, "nQuantity") = CDbl(.TextMatrix(lnCtr, 3))
                     Call GrandTotal
                     lbDuplicate = True
                     End If
               Next
            End With
         
            If Not lbDuplicate Then
               If Trim(.Text) <> "" Then Call InsertDetail(lnQty, lsBarrCode)
            End If
                     
            .Text = ""
         
            .SetFocus
            Case vbKeyF3
         End Select
      Case 1
      Case 2
      Case 3
   End Select
End With
End Sub

Private Sub txtDetail_LostFocus(Index As Integer)
With txtDetail(Index)
   .BackColor = oApp.getColor("EB")
      
   Select Case Index
         Case 1, 2, 3
      txtDetail(Index).Enabled = False
   End Select
End With
End Sub

Private Sub txtDetail_Validate(Index As Integer, Cancel As Boolean)
With txtDetail(Index)
   Select Case Index
         Case 1
      If Not IsNumeric(.Text) Then .Text = 0#
      oTrans.Detail(pnRow, 8) = CDbl(.Text)
      Call GrandTotal
   Case 2
      If Not IsNumeric(.Text) Then .Text = 0#
      oTrans.Detail(pnRow, 7) = .Text
   Case 3
      If Not IsNumeric(.Text) Then .Text = 0#
      oTrans.Detail(pnRow, 10) = CDbl(.Text)
   End Select
End With
End Sub

Private Sub txtField_GotFocus(Index As Integer)
With txtField(Index)
   Select Case Index
         Case 1
      .Text = Format(.Text, "MM/DD/YYYY")
   Case 9
      .Text = Format(.Text, "##0.00")
   End Select
      
   .SelStart = 0
   .SelLength = Len(.Text)
   .BackColor = oApp.getColor("HT1")
End With

pnIndex = Index
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim lsOldProc As String
   
lsOldProc = "txtField_KeyDown"
''On Error GoTo errProc
   
If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
   With txtField(Index)
      If KeyCode = vbKeyF3 Then
         oTrans.SearchMaster Index, .Text
            
         If .Text <> "" Then SetNextFocus
      Else
         If .Text <> "" Then oTrans.SearchMaster Index, .Text
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

Private Sub txtField_LostFocus(Index As Integer)
With txtField(Index)
   .BackColor = oApp.getColor("EB")
End With
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
With txtField(Index)
   Select Case Index
         Case 1
      If Not IsDate(.Text) Then .Text = oApp.ServerDate
      .Text = Format(.Text, "MMMM DD, YYYY")
         
      oTrans.Master(Index) = CDate(.Text)
   Case 7
      .Text = TitleCase(.Text)
         
      oTrans.Master(Index) = .Text
   Case 9
      If Not IsNumeric(.Text) Then .Text = 0#
      .Text = Format(.Text, "##0.00") & "%"
         
      oTrans.Master(Index) = CDbl(Format(.Text, "##0.00"))
   Case 10, 13
      If Not IsNumeric(.Text) Then .Text = 0#
      .Text = Format(.Text, "#,##0.00")
         
      oTrans.Master(Index) = CDbl(.Text)
   Case Else
      oTrans.Master(Index) = .Text
   End Select
End With
End Sub

Private Sub GrandTotal()
Dim lsOldProc As String
Dim lnCtr As Integer
Dim lnTotal As Currency

lsOldProc = pxeMODULENAME & ".GrandTotal"
''On Error GoTo errProc
   
With MSFlexGrid1
   lnTotal = 0#
   
   For lnCtr = 1 To .Rows - 1

      If .TextMatrix(lnCtr, 4) = "" Then .TextMatrix(lnCtr, 4) = 0#
       If .TextMatrix(lnCtr, 6) = "" Then .TextMatrix(lnCtr, 6) = 0#
      .TextMatrix(lnCtr, 7) = Format(CDbl(.TextMatrix(lnCtr, 3)) * CDbl(.TextMatrix(lnCtr, 4)) - CDbl(.TextMatrix(lnCtr, 6)), "#,##0.00")
      lnTotal = lnTotal + CDbl(IIf(.TextMatrix(lnCtr, 7) = "", 0, .TextMatrix(lnCtr, 7)))
   Next
   
End With
lblTrantotal.Caption = Format(lnTotal, "#,##0.00")
oTrans.Master("nTranTotl") = CDbl(lnTotal)
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True

End Sub

Private Sub InsertDetail(ByVal Quantity As Integer, ByVal Value As String)
Dim lsOldProc As String
   
lsOldProc = pxeMODULENAME & ".InsertDetail"
''On Error GoTo errProc
   
With MSFlexGrid1
   If .Rows = 2 Then
      If .TextMatrix(.Row, 1) <> "" Then
         If oTrans.ItemCount <> .Row Then
            oTrans.addDetail
            oTrans.Detail(.Rows - 1, "xReferNox") = Value
            If oTrans.Detail(.Rows - 1, "xReferNox") <> "" Then
               .Rows = .Rows + 1
               .Row = .Rows - 1
               .TextMatrix(.Row, 1) = Value
               .TextMatrix(.Row, 0) = .Row
            Else
               oTrans.deleteDetail .Row
               Exit Sub
            End If
         Else
            oTrans.addDetail
            oTrans.Detail(.Row, "xReferNox") = Value
            If oTrans.Detail(.Row, "xReferNox") <> "" Then
               .Rows = .Rows + 1
               .Row = .Rows - 1
               .TextMatrix(.Row, 1) = Value
               .TextMatrix(.Row, 0) = .Row
            Else
               oTrans.deleteDetail .Row
               Exit Sub
            End If
         End If
      Else
         oTrans.Detail(.Row - 1, "xReferNox") = Value
         If oTrans.Detail(.Row - 1, "xReferNox") <> "" Then .TextMatrix(.Row, 1) = Value
         .TextMatrix(.Row, 0) = .Row
      End If
   Else
      If oTrans.ItemCount <> .Row Then
         oTrans.addDetail
         oTrans.Detail(.Rows - 1, "xReferNox") = Value
         If oTrans.Detail(.Rows - 1, "xReferNox") <> "" Then
            .Rows = .Rows + 1
            .Row = .Rows - 1
            .TextMatrix(.Row, 1) = Value
            .TextMatrix(.Row, 0) = .Row
         Else
            oTrans.deleteDetail .Rows
            Exit Sub
         End If
      Else
         oTrans.addDetail
         oTrans.Detail(.Row, "xReferNox") = Value
         If oTrans.Detail(.Row, "xReferNox") <> "" Then
            .Rows = .Rows + 1
            .Row = .Rows - 1
            .TextMatrix(.Row, 1) = Value
            .TextMatrix(.Row, 0) = .Row
         Else
            oTrans.deleteDetail .Row
            Exit Sub
         End If
      End If
   End If
   .ColSel = .Cols - 1

   .ColWidth(2) = 2950
   If .Rows > 21 Then .ColWidth(2) = 2700

   oTrans.Detail(.Row - 1, "nQuantity") = Quantity
   .TextMatrix(.Row, 1) = oTrans.Detail(.Row - 1, "xReferNox")
   .TextMatrix(.Row, 2) = oTrans.Detail(.Row - 1, "sDescript")
   .TextMatrix(.Row, 3) = oTrans.Detail(.Row - 1, "nQuantity")
   .TextMatrix(.Row, 4) = Format(oTrans.Detail(.Row - 1, "nUnitPrce"), "#,##0.00")
   .TextMatrix(.Row, 5) = Format(oTrans.Detail(.Row - 1, "nDiscRate"), "##0.00") & "%"
   .TextMatrix(.Row, 6) = Format(oTrans.Detail(.Row - 1, "nDiscAmtx"), "#,##0.00")
   If .TextMatrix(.Row, 6) <> "" Then .TextMatrix(.Row, 7) = Format(CDbl(.TextMatrix(.Row, 3)) * CDbl(.TextMatrix(.Row, 4)) - CDbl(.TextMatrix(.Row, 6)), "#,##0.00")

   Call GrandTotal
      
   pnRow = .Rows - 2
End With
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub refreshGrid()
Dim lnCtr As Integer
Dim lsOldProc As String
   
lsOldProc = pxeMODULENAME & ".refreshGrid"
''On Error GoTo errProc
   
Call InitGrid

With MSFlexGrid1
   .Rows = IIf(oTrans.ItemCount = 0, 2, oTrans.ItemCount + 1)
   For lnCtr = 1 To .Rows - 1
      .TextMatrix(lnCtr, 0) = lnCtr
      .TextMatrix(lnCtr, 1) = oTrans.Detail(lnCtr - 1, "xReferNox")
      .TextMatrix(lnCtr, 2) = oTrans.Detail(lnCtr - 1, "sDescript")
      .TextMatrix(lnCtr, 3) = oTrans.Detail(lnCtr - 1, "nQuantity")
      .TextMatrix(lnCtr, 4) = Format(oTrans.Detail(lnCtr - 1, "nUnitPrce"), "#,##0.00")
      .TextMatrix(lnCtr, 5) = Format(oTrans.Detail(lnCtr - 1, "nDiscRate"), "##0.00") & "%"
      .TextMatrix(lnCtr, 6) = Format(oTrans.Detail(lnCtr - 1, "nDiscAmtx"), "#,##0.00")
      If .TextMatrix(lnCtr, 3) <> 0 Then .TextMatrix(lnCtr, 7) = Format(CDbl(.TextMatrix(lnCtr, 3)) * CDbl(.TextMatrix(lnCtr, 4)) - CDbl(.TextMatrix(lnCtr, 6)), "#,##0.00")
   Next
      
   .Row = .Rows - 1
   .ColSel = .Cols - 1

   .ColWidth(2) = 2950
   If .Rows > 21 Then .ColWidth(2) = 2700
      
   pnRow = 0
   Call GrandTotal
End With
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Function isEntryOk() As Boolean
   Dim lsOldProc As String
   Dim lnCtr As Integer
   Dim lnUserRights As Integer
   Dim lsUserID As String
   Dim lsUserName As String
   
   lsOldProc = pxeMODULENAME & ".isEntryOK"
   ''On Error GoTo errProc
   With MSFlexGrid1
      If .TextMatrix(0, 1) = "" Then
         MsgBox "Unable to Save Transaction." _
                        & vbCrLf _
                        & vbCrLf _
                        & "Please verify your entry.", vbCritical, pxeAPPNAME
         GoTo endProc
      End If
         
      With oTrans
         If .Master("sTermIDxx") = "" Then
            MsgBox "Invalid Term Detected.", vbCritical, pxeAPPNAME
            GoTo endProc
         End If
         
'         If .Master("nAmtPaidx") = .Master("nTranTotl") Then
'            If GetApproval(oApp, lnUserRights, lsUserID, lsUserName, oApp.MenuName) = False Then GoTo endProc
'
'            If lnUserRights < xeManager Then
'               MsgBox "Approving Officer Has no Right to Save this transaction!!!" & vbCrLf & _
'                        "Request can not be granted!!!" & vbCrLf & _
'                        "Pls Issue Sales Invoice for Full Payment", vbCritical, "Warning"
'               GoTo endProc
'            End If
'         End If
         
         For lnCtr = 0 To .ItemCount - 1
            If .Detail(lnCtr, "nQuantity") = 0 Then
               MsgBox "There is an item with no quantity." & vbCrLf & vbCrLf & _
                  "Please verify your entry.", vbCritical, pxeAPPNAME
               GoTo endProc
            End If
         Next
      End With
   End With
   
   isEntryOk = True

endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )", True
End Function

Private Function PrinTrans()
    Dim lrs As Recordset
   Dim lors As Recordset
   Dim lnCtr As Integer
   Dim lsOldProc As String
   Dim lnTotalAmt As Currency
   Dim lsSQL As String

   lsOldProc = "PrinTrans"
   ''On Error GoTo errProc
   
   PrinTrans = True

      
   lsSQL = "SELECT " & _
            " a.sTransNox" & _
            ", b.sSerialNo" & _
            ", c.sBarrCode" & _
            ", c.sDescript" & _
            ", a.nQuantity" & _
            ", a.nUnitPrce" & _
            ", a.nDiscRate" & _
            ", a.nDiscAmtx" & _
            ", d.sModelNme" & _
            ", e.sBrandNme" & _
            " FROM CP_CO_Detail a" & _
               " LEFT JOIN CP_Inventory_Serial b" & _
                  " ON a.sSerialID = b.sSerialID" & _
               " LEFT JOIN CP_Inventory c" & _
                  " ON a.sStockIDx = c.sStockIDx" & _
               " LEFT JOIN CP_Model d ON c.sModelIDx = d.sModelIDx" & _
               " LEFT JOIN CP_Brand e ON c.sBrandIDx = e.sBrandIDx" & _
            " WHERE a.sTransNox = " & strParm(oTrans.Master("sTransNox"))
            
   Set lors = New Recordset
   lors.Open lsSQL, oApp.Connection, adOpenDynamic, adLockReadOnly, adCmdText
   

   Set lrs = New ADODB.Recordset

   lrs.Fields.Append "nField01", adInteger, 10
   lrs.Fields.Append "sField01", adVarChar, 200
   lrs.Fields.Append "sField02", adVarChar, 200
   lrs.Fields.Append "sField03", adVarChar, 200
   lrs.Fields.Append "sField04", adVarChar, 200
   lrs.Fields.Append "lField01", adCurrency
   lrs.Fields.Append "lField02", adCurrency

   lrs.Open

   With lors
      lnTotalAmt = 0

      For lnCtr = 0 To .RecordCount - 1
         lrs.AddNew
         lrs.Fields("nField01") = lors("nQuantity")
         lrs.Fields("sField01") = oTrans.Master(3)
         lrs.Fields("sField02") = oTrans.Master(4)
         lrs.Fields("sField03") = lors("sBarrCode") & " " & lors("sBrandNme") & " " & lors("sModelNme") & " " & IFNull(lors("sSerialNo"), "")
'         lrs.Fields("sField03") = lors("sBarrCode") & " " & lors("sDescript") & " " & IFNull(lors("sSerialNo"), "")
         lrs.Fields("sField04") = Format(oTrans.Master("dTransact"), "MMM DD, YYYY")
         lrs.Fields("lField01") = lors("nUnitPrce")
         lrs.Fields("lField02") = lors("nQuantity") * lors("nUnitPrce")
         
         lnTotalAmt = Format(lnTotalAmt + lors("nQuantity") * lors("nUnitPrce"), "#,##0.00")
         lors.MoveNext
      Next

      lrs.Sort = "sField01 DESC"
   End With

   'assign important info to the report
   Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\CP_CI.rpt")
   oReport.DiscardSavedData
   oReport.FieldMappingType = crAutoFieldMapping
   oReport.Database.SetDataSource lrs

   With oReport
      oReport.Sections("PF").ReportObjects("txtTotalAmt").SetText Format(lnTotalAmt, "#,##0.00")
      oReport.Sections("PF").ReportObjects("txtCustomer").SetText oTrans.Master(3)
      oReport.Sections("PF").ReportObjects("txtRemarks").SetText oTrans.Master("sRemarksx")
   End With
   
   oReport.PrintOutEx False, 1
   lrs.Close
   PrinTrans = True

endProc:
   Set oReport = Nothing
   Set lrs = Nothing
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"


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
