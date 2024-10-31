VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash9f.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCP_POS1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Guanzon Telecom POS(Branch Version)"
   ClientHeight    =   10245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15330
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10245
   ScaleWidth      =   15330
   ShowInTaskbar   =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H0075BEFB&
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   1
      Left            =   8295
      ScaleHeight     =   375
      ScaleWidth      =   6540
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   7485
      Width           =   6570
      Begin VB.TextBox txtothers 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   9
         Left            =   4515
         TabIndex        =   71
         TabStop         =   0   'False
         Text            =   "Withholding Tax"
         Top             =   45
         Width           =   1875
      End
      Begin VB.TextBox txtothers 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   7
         Left            =   885
         TabIndex        =   69
         TabStop         =   0   'False
         Text            =   "Withholding Tax"
         Top             =   45
         Width           =   1875
      End
      Begin VB.Label lblField 
         BackStyle       =   0  'Transparent
         Caption         =   "Withholding Tax"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   210
         Index           =   19
         Left            =   3000
         TabIndex        =   72
         Top             =   90
         Width           =   1620
      End
      Begin VB.Label lblField 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   165
         Index           =   13
         Left            =   135
         TabIndex        =   70
         Top             =   90
         Width           =   1620
      End
   End
   Begin VB.PictureBox Picture15 
      Appearance      =   0  'Flat
      BackColor       =   &H0023A0FE&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   30
      Left            =   75
      ScaleHeight     =   30
      ScaleWidth      =   15120
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   10065
      Width           =   15120
   End
   Begin VB.PictureBox Picture14 
      Appearance      =   0  'Flat
      BackColor       =   &H0023A0FE&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   30
      Left            =   75
      ScaleHeight     =   30
      ScaleWidth      =   15120
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   75
      Width           =   15120
   End
   Begin VB.PictureBox Picture13 
      Appearance      =   0  'Flat
      BackColor       =   &H0023A0FE&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9990
      Left            =   15165
      ScaleHeight     =   9990
      ScaleWidth      =   30
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   75
      Width           =   30
   End
   Begin VB.PictureBox Picture10 
      Appearance      =   0  'Flat
      BackColor       =   &H0023A0FE&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9990
      Left            =   75
      ScaleHeight     =   9990
      ScaleWidth      =   30
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   90
      Width           =   30
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0023A0FE&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4650
      Index           =   0
      Left            =   240
      ScaleHeight     =   4650
      ScaleWidth      =   14775
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   930
      Width           =   14775
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   1785
         Left            =   90
         ScaleHeight     =   1755
         ScaleWidth      =   14550
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   90
         Width           =   14580
         Begin VB.PictureBox Picture4 
            Appearance      =   0  'Flat
            BackColor       =   &H0075BEFB&
            ForeColor       =   &H00000000&
            Height          =   1695
            Left            =   30
            ScaleHeight     =   1665
            ScaleWidth      =   14460
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   30
            Width           =   14490
            Begin VB.TextBox txtothers 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   8
               Left            =   7185
               TabIndex        =   21
               Text            =   "Discount"
               Top             =   1125
               Width           =   1065
            End
            Begin VB.TextBox txtothers 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   0
               Left            =   1140
               TabIndex        =   10
               Text            =   "Bar Code"
               ToolTipText     =   "Esc"
               Top             =   300
               Width           =   2910
            End
            Begin VB.TextBox txtothers 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               Height          =   510
               Index           =   2
               Left            =   1140
               TabIndex        =   14
               Text            =   "Unit Price"
               Top             =   870
               Width           =   2910
            End
            Begin VB.TextBox txtothers 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   510
               Index           =   3
               Left            =   5175
               TabIndex        =   16
               Text            =   "Quantity"
               Top             =   855
               Width           =   630
            End
            Begin VB.TextBox txtothers 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   4
               Left            =   7185
               TabIndex        =   18
               Text            =   "Discount"
               Top             =   870
               Width           =   750
            End
            Begin VB.TextBox txtothers 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   18
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   510
               Index           =   5
               Left            =   10545
               Locked          =   -1  'True
               TabIndex        =   23
               Text            =   "Sub Total"
               Top             =   885
               Width           =   3465
            End
            Begin VB.TextBox txtothers 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   1
               Left            =   7185
               TabIndex        =   12
               TabStop         =   0   'False
               Text            =   "Description"
               Top             =   315
               Width           =   6825
            End
            Begin VB.Label lblField 
               BackStyle       =   0  'Transparent
               Caption         =   "%"
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
               Index           =   16
               Left            =   8040
               TabIndex        =   19
               Top             =   900
               Width           =   210
            End
            Begin VB.Label lblField 
               BackStyle       =   0  'Transparent
               Caption         =   "Disc. Amt."
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
               Index           =   15
               Left            =   6180
               TabIndex        =   20
               Top             =   1110
               Width           =   720
            End
            Begin VB.Shape Shape2 
               Height          =   750
               Index           =   0
               Left            =   150
               Top             =   765
               Width           =   8865
            End
            Begin VB.Label lblField 
               BackStyle       =   0  'Transparent
               Caption         =   "&Bar Code"
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
               Index           =   12
               Left            =   285
               TabIndex        =   9
               ToolTipText     =   "Esc"
               Top             =   330
               Width           =   1185
            End
            Begin VB.Label lblField 
               BackStyle       =   0  'Transparent
               Caption         =   "Unit Price"
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
               Height          =   315
               Index           =   21
               Left            =   270
               TabIndex        =   13
               Top             =   855
               Width           =   870
            End
            Begin VB.Label lblField 
               BackStyle       =   0  'Transparent
               Caption         =   "&Quantity"
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
               Index           =   30
               Left            =   4455
               TabIndex        =   15
               Top             =   855
               Width           =   615
            End
            Begin VB.Label lblField 
               BackStyle       =   0  'Transparent
               Caption         =   "Discount"
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
               Index           =   3
               Left            =   6165
               TabIndex        =   17
               Top             =   855
               Width           =   795
            End
            Begin VB.Label lblField 
               BackStyle       =   0  'Transparent
               Caption         =   "Sub Total"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   270
               Index           =   8
               Left            =   9225
               TabIndex        =   22
               Top             =   915
               Width           =   1005
            End
            Begin VB.Label lblField 
               BackStyle       =   0  'Transparent
               Caption         =   "Description"
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
               Height          =   225
               Index           =   9
               Left            =   6180
               TabIndex        =   11
               Top             =   330
               Width           =   900
            End
            Begin VB.Shape Shape2 
               Height          =   585
               Index           =   1
               Left            =   180
               Top             =   45
               Width           =   14160
            End
            Begin VB.Shape Shape2 
               Height          =   750
               Index           =   3
               Left            =   9045
               Top             =   765
               Width           =   5265
            End
         End
      End
      Begin VB.PictureBox Picture11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         ForeColor       =   &H80000008&
         Height          =   2685
         Left            =   90
         ScaleHeight     =   2655
         ScaleWidth      =   14550
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   1890
         Width           =   14580
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H00D9FCFD&
            ForeColor       =   &H80000008&
            Height          =   2595
            Index           =   2
            Left            =   30
            ScaleHeight     =   2565
            ScaleWidth      =   14460
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   30
            Width           =   14490
            Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
               Height          =   2505
               Left            =   30
               TabIndex        =   26
               TabStop         =   0   'False
               Top             =   30
               Width           =   14400
               _ExtentX        =   25400
               _ExtentY        =   4419
               _Version        =   393216
               ForeColor       =   4194304
               BackColorSel    =   12648447
               ForeColorSel    =   4194304
               FocusRect       =   0
               SelectionMode   =   1
               Appearance      =   0
            End
         End
      End
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H0023A0FE&
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   225
      ScaleHeight     =   885
      ScaleWidth      =   14775
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   9030
      Width           =   14805
      Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
         Height          =   780
         Left            =   45
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   45
         Width           =   14670
         _cx             =   25876
         _cy             =   1376
         FlashVars       =   ""
         Movie           =   ""
         Src             =   ""
         WMode           =   "Window"
         Play            =   -1  'True
         Loop            =   -1  'True
         Quality         =   "High"
         SAlign          =   ""
         Menu            =   -1  'True
         Base            =   ""
         AllowScriptAccess=   ""
         Scale           =   "ShowAll"
         DeviceFont      =   0   'False
         EmbedMovie      =   0   'False
         BGColor         =   ""
         SWRemote        =   ""
         MovieData       =   ""
         SeamlessTabbing =   -1  'True
         Profile         =   0   'False
         ProfileAddress  =   ""
         ProfilePort     =   0
         AllowNetworking =   "all"
         AllowFullScreen =   "false"
      End
   End
   Begin VB.PictureBox Picture8 
      Appearance      =   0  'Flat
      BackColor       =   &H0023A0FE&
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   225
      ScaleHeight     =   570
      ScaleWidth      =   14775
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   225
      Width           =   14805
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   60
         ScaleHeight     =   450
         ScaleWidth      =   14655
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   45
         Width           =   14685
         Begin VB.Label lblField 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Time"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   1
            Left            =   12885
            TabIndex        =   5
            Top             =   90
            Width           =   1650
         End
         Begin VB.Label lblField 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   0
            Left            =   7530
            TabIndex        =   4
            Top             =   105
            Width           =   5280
         End
         Begin VB.Label lblField 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Cashier"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   17
            Left            =   2190
            TabIndex        =   3
            Top             =   105
            Width           =   5190
         End
         Begin VB.Label lblField 
            BackStyle       =   0  'Transparent
            Caption         =   "Cashier-In-Charge "
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00D9FCFD&
            Height          =   300
            Index           =   18
            Left            =   60
            TabIndex        =   2
            Top             =   105
            Width           =   2040
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0023A0FE&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3270
      Index           =   3
      Left            =   240
      ScaleHeight     =   3270
      ScaleWidth      =   14775
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   5640
      Width           =   14775
      Begin VB.PictureBox Picture9 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   840
         Left            =   90
         ScaleHeight     =   810
         ScaleWidth      =   14550
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   2340
         Width           =   14580
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H0075BEFB&
            ForeColor       =   &H80000008&
            Height          =   750
            Index           =   5
            Left            =   30
            ScaleHeight     =   720
            ScaleWidth      =   14460
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   30
            Width           =   14490
            Begin VB.TextBox txtfield 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   240
               Index           =   6
               Left            =   1425
               TabIndex        =   42
               Text            =   "Remarks"
               Top             =   360
               Width           =   4785
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
               Left            =   7620
               MaxLength       =   15
               TabIndex        =   46
               Text            =   "Invoice No."
               Top             =   360
               Width           =   2055
            End
            Begin VB.TextBox txtfield 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   240
               Index           =   0
               Left            =   1425
               TabIndex        =   40
               TabStop         =   0   'False
               Text            =   "Transaction Number"
               Top             =   105
               Width           =   2025
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
               Index           =   4
               Left            =   7620
               TabIndex        =   44
               Text            =   "Customer Name"
               Top             =   105
               Width           =   6690
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
               Index           =   5
               Left            =   10905
               TabIndex        =   48
               Text            =   "Sales Person"
               Top             =   360
               Width           =   3405
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
               Index           =   2
               Left            =   180
               TabIndex        =   41
               Top             =   360
               Width           =   1515
            End
            Begin VB.Label lblField 
               BackStyle       =   0  'Transparent
               Caption         =   "Invoice Number"
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
               Left            =   6315
               TabIndex        =   45
               Top             =   360
               Width           =   1350
            End
            Begin VB.Label lblField 
               BackStyle       =   0  'Transparent
               Caption         =   "Transaction No."
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
               Index           =   10
               Left            =   165
               TabIndex        =   39
               Top             =   105
               Width           =   1515
            End
            Begin VB.Label lblField 
               BackStyle       =   0  'Transparent
               Caption         =   "Customer Name"
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
               Index           =   6
               Left            =   6315
               TabIndex        =   43
               Top             =   105
               Width           =   1350
            End
            Begin VB.Label lblField 
               BackStyle       =   0  'Transparent
               Caption         =   "Sales Person"
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
               Index           =   14
               Left            =   9780
               TabIndex        =   47
               Top             =   360
               Width           =   1350
            End
         End
      End
      Begin VB.PictureBox Picture12 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   2220
         Left            =   90
         ScaleHeight     =   2190
         ScaleWidth      =   14550
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   90
         Width           =   14580
         Begin VB.PictureBox Picture6 
            Appearance      =   0  'Flat
            BackColor       =   &H0075BEFB&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   0
            Left            =   30
            ScaleHeight     =   120
            ScaleWidth      =   7875
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   1995
            Width           =   7905
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H0075BEFB&
            ForeColor       =   &H80000008&
            Height          =   1710
            Index           =   4
            Left            =   7950
            ScaleHeight     =   1680
            ScaleWidth      =   6540
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   30
            Width           =   6570
            Begin VB.TextBox txtfield 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
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
               Height          =   450
               Index           =   2
               Left            =   2595
               TabIndex        =   32
               TabStop         =   0   'False
               Text            =   "Total Amount"
               Top             =   195
               Width           =   3585
            End
            Begin VB.TextBox txtothers 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   360
               Index           =   6
               Left            =   2595
               TabIndex        =   36
               TabStop         =   0   'False
               Text            =   "Change"
               Top             =   1155
               Width           =   3585
            End
            Begin VB.TextBox txtfield 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   360
               Index           =   3
               Left            =   2595
               TabIndex        =   34
               Text            =   "Cash Given"
               Top             =   675
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
               Left            =   465
               TabIndex        =   31
               Top             =   285
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
               Index           =   5
               Left            =   465
               TabIndex        =   35
               Top             =   1215
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
               Left            =   465
               TabIndex        =   33
               Top             =   675
               Width           =   1545
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00000000&
               X1              =   465
               X2              =   6165
               Y1              =   1080
               Y2              =   1080
            End
            Begin VB.Shape Shape2 
               Height          =   1485
               Index           =   2
               Left            =   135
               Top             =   105
               Width           =   6270
            End
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H0075BEFB&
            ForeColor       =   &H80000008&
            Height          =   1935
            Left            =   30
            ScaleHeight     =   1905
            ScaleWidth      =   7875
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   30
            Width           =   7905
            Begin xrControl.xrButton cmdButton 
               Height          =   540
               Index           =   4
               Left            =   1425
               TabIndex        =   52
               ToolTipText     =   "Save Transaction"
               Top             =   150
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   953
               Caption         =   "F2-Save"
               AccessKey       =   "F2-Save"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Picture         =   "frmCP_POS1.frx":0000
               PicturePos      =   3
               CaptionAlign    =   0
               BackColor       =   14286077
               BackColorDown   =   8775418
               BorderColorFocus=   8775418
               BorderColorHover=   8775418
            End
            Begin xrControl.xrButton cmdButton 
               Height          =   540
               Index           =   3
               Left            =   1425
               TabIndex        =   55
               ToolTipText     =   "Cheque Payment"
               Top             =   690
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   953
               Caption         =   "F5-Cheq"
               AccessKey       =   "F5-Cheq"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Picture         =   "frmCP_POS1.frx":077A
               PicturePos      =   3
               CaptionAlign    =   0
               BackColor       =   14286077
               BackColorDown   =   8775418
               BorderColorFocus=   8775418
               BorderColorHover=   8775418
            End
            Begin xrControl.xrButton cmdButton 
               Height          =   540
               Index           =   2
               Left            =   2730
               TabIndex        =   56
               ToolTipText     =   "Log Out"
               Top             =   690
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   953
               Caption         =   "F6-Card"
               AccessKey       =   "F6-Card"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Picture         =   "frmCP_POS1.frx":0EF4
               PicturePos      =   3
               CaptionAlign    =   0
               BackColor       =   14286077
               BackColorDown   =   8775418
               BorderColorFocus=   8775418
               BorderColorHover=   8775418
            End
            Begin xrControl.xrButton cmdButton 
               Height          =   540
               Index           =   8
               Left            =   4035
               TabIndex        =   61
               ToolTipText     =   "Exit"
               Top             =   1230
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   953
               Caption         =   "F11-Exit"
               AccessKey       =   "F11-Exit"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Picture         =   "frmCP_POS1.frx":166E
               PicturePos      =   3
               CaptionAlign    =   0
               BackColor       =   14286077
               BackColorDown   =   8775418
               BorderColorFocus=   8775418
               BorderColorHover=   8775418
            End
            Begin xrControl.xrButton cmdButton 
               Height          =   540
               Index           =   0
               Left            =   120
               TabIndex        =   51
               ToolTipText     =   "Discount"
               Top             =   150
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   953
               Caption         =   "F1-Disc"
               AccessKey       =   "F1-Disc"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Picture         =   "frmCP_POS1.frx":1DE8
               PicturePos      =   3
               CaptionAlign    =   0
               BackColor       =   14286077
               BackColorDown   =   8775418
               BorderColorFocus=   8775418
               BorderColorHover=   8775418
            End
            Begin xrControl.xrButton cmdButton 
               Height          =   540
               Index           =   1
               Left            =   120
               TabIndex        =   54
               ToolTipText     =   "Void Transaction"
               Top             =   690
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   953
               Caption         =   "F4-Void"
               AccessKey       =   "F4-Void"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Picture         =   "frmCP_POS1.frx":2562
               PicturePos      =   3
               CaptionAlign    =   0
               BackColor       =   14286077
               BackColorDown   =   8775418
               BorderColorFocus=   8775418
               BorderColorHover=   8775418
            End
            Begin xrControl.xrButton cmdButton 
               Height          =   540
               Index           =   5
               Left            =   4035
               TabIndex        =   57
               ToolTipText     =   "Credit Card Payment"
               Top             =   690
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   953
               Caption         =   "F7-Inst."
               AccessKey       =   "F7-Inst."
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Picture         =   "frmCP_POS1.frx":2CDC
               PicturePos      =   3
               CaptionAlign    =   0
               BackColor       =   14286077
               BackColorDown   =   8775418
               BorderColorFocus=   8775418
               BorderColorHover=   8775418
            End
            Begin xrControl.xrButton cmdButton 
               Height          =   540
               Index           =   9
               Left            =   2730
               TabIndex        =   53
               ToolTipText     =   "Job Order"
               Top             =   150
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   953
               Caption         =   "F3-Find"
               AccessKey       =   "F3-Find"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Picture         =   "frmCP_POS1.frx":3456
               PicturePos      =   3
               CaptionAlign    =   0
               BackColor       =   14286077
               BackColorDown   =   8775418
               BorderColorFocus=   8775418
               BorderColorHover=   8775418
            End
            Begin xrControl.xrButton cmdButton 
               Height          =   540
               Index           =   6
               Left            =   120
               TabIndex        =   58
               ToolTipText     =   "Credit Card Payment"
               Top             =   1230
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   953
               Caption         =   "F8-Regr"
               AccessKey       =   "F8-Regr"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Picture         =   "frmCP_POS1.frx":3BD0
               PicturePos      =   3
               CaptionAlign    =   0
               BackColor       =   14286077
               BackColorDown   =   8775418
               BorderColorFocus=   8775418
               BorderColorHover=   8775418
            End
            Begin xrControl.xrButton cmdButton 
               Height          =   540
               Index           =   7
               Left            =   1425
               TabIndex        =   59
               ToolTipText     =   "Credit Card Payment"
               Top             =   1230
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   953
               Caption         =   "F9-Repl"
               AccessKey       =   "F9-Repl"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Picture         =   "frmCP_POS1.frx":434A
               PicturePos      =   3
               CaptionAlign    =   0
               BackColor       =   14286077
               BackColorDown   =   8775418
               BorderColorFocus=   8775418
               BorderColorHover=   8775418
            End
            Begin xrControl.xrButton cmdButton 
               Height          =   540
               Index           =   10
               Left            =   2730
               TabIndex        =   60
               ToolTipText     =   "Credit Card Payment"
               Top             =   1230
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   953
               Caption         =   "F10-JO"
               AccessKey       =   "F10-JO"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Picture         =   "frmCP_POS1.frx":4AC4
               PicturePos      =   3
               CaptionAlign    =   0
               BackColor       =   14286077
               BackColorDown   =   8775418
               BorderColorFocus=   8775418
               BorderColorHover=   8775418
            End
            Begin VB.Image Image1 
               Height          =   510
               Left            =   4050
               Picture         =   "frmCP_POS1.frx":523E
               Stretch         =   -1  'True
               Top             =   150
               Width           =   1275
            End
            Begin VB.Image Image2 
               Height          =   1560
               Left            =   5265
               Picture         =   "frmCP_POS1.frx":62F6
               Stretch         =   -1  'True
               Top             =   165
               Width           =   2565
            End
         End
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   255
      Top             =   0
   End
   Begin VB.Label lblField 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   11
      Left            =   360
      TabIndex        =   66
      Top             =   405
      Width           =   3525
   End
End
Attribute VB_Name = "frmCP_POS1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
'Private WithEvents oDriver As FormDriver
'Private oSkin As FormSkin
'Private bLoaded As Boolean
'Private oRS As New ADODB.Recordset
'
'Dim txtfieldGotfocus As Boolean
'Dim txtOthersGotfocus As Boolean
'Dim pbnewitem As Boolean
'Dim psSelected() As String
'Dim lsSQL As String
'Dim pnIndex As Integer
'Dim pnCtr As Integer
'Dim void As Boolean
'
'Dim pnUserRights As Integer
'Dim psUserID As String
'Dim psUserName As String
'
'Dim psPayment As String
'Dim psClientID As String
'
'Property Let ClientID(ClientID As String)
'   psClientID = ClientID
'End Property
'
'
'Property Let Payment(Payment As String)
'   psPayment = Payment
'End Property
'
'Private Sub cmdButton_Click(Index As Integer)
'
'   Select Case Index
'         Case 0 'Discount Approval
'            If txtfield(2).Text = 0# Then Exit Sub
'            Disc_Approval
'         Case 1 'Void
'            If txtfield(2).Text = 0# Then Exit Sub
'            Void_Approval
'         Case 2 'Card
'            If txtfield(2).Text = 0# Then Exit Sub
'            frmCard_POS.txtfield(1) = Format(txtfield(2), "#,##0.00")
'            frmCard_POS.Transaction = "POS"
'            frmCard_POS.Show 1
'         Case 3 'Check
'            If txtfield(2).Text = 0# Then Exit Sub
'            frmCheque_POS.txtfield(1) = Format(txtfield(2), "#,##0.00")
'            frmCheque_POS.Transaction = "POS"
'            frmCheque_POS.Show 1
'         Case 4 'save
'            If txtfield(2).Text = 0# Then Exit Sub
'            oDriver.RecordSave
'         Case 5 'Installment
'            If txtfield(2).Text = 0# Then Exit Sub
'            Installment_Approval
'         Case 6 'Register
'            frmPOS_Register.Show
'         Case 7 'Change Unit
'            frmChange_Unit.Show
'         Case 8 'Exit
'            Unload Me
'         Case 9 'Search Bar Code
'            ClearFields
'            SearchBarCode False
'         Case 10 'JO
'            frmJOMenu.Show
'   End Select
'
'End Sub
'
'Private Sub Installment_Approval()
'Dim lsApproval As Integer
'
'If oApp.UserLevel = xeEncoder Then
'   lsApproval = MsgBox("User Not Allowed to allow Installment!!!" & vbCrLf & _
'               "Seek for Approval?", vbQuestion + vbYesNo, "Confirm")
'   If lsApproval = vbYes Then
'      If Not GetApproval(oApp, pnUserRights, psUserID, psUserName) Then Exit Sub
'      If pnUserRights < xeManager Then
'         MsgBox "Approving User is Not Authorized!!!", vbCritical, "Warning"
'         Exit Sub
'      Else
'         frmInstallment_POS.Transaction = "POS"
'         frmInstallment_POS.txtfield(0) = Format(txtfield(2).Text, "#,##0.00")
'         frmInstallment_POS.Show 1
'      End If
'   Else
'      txtothers(0).SetFocus
'      Exit Sub
'   End If
'Else
'   frmInstallment_POS.Transaction = "POS"
'   frmInstallment_POS.txtfield(0) = Format(txtfield(2).Text, "#,##0.00")
'   frmInstallment_POS.Show 1
'End If
'
'End Sub
'
'Private Sub Disc_Approval()
'Dim lsApproval As Integer
'
''If oApp.UserLevel = xeEncoder Then
''   lsApproval = MsgBox("User Not Allowed to give Discount!!!" & vbCrLf & _
''               "Seek for Approval?", vbQuestion + vbYesNo, "Confirm")
''   If lsApproval = vbYes Then
''      If Not GetApproval(oApp, pnUserRights, psUserID, psUserName) Then Exit Sub
''         If pnUserRights < xeManager Then
''            MsgBox "Approving User is Not Authorized!!!", vbCritical, "Warning"
''            Exit Sub
''         Else
''            txtothers(4).Enabled = True
''            txtothers(8).Enabled = True
''            txtothers(4).SetFocus
''         End If
''   Else
''      txtothers(0).SetFocus
''   End If
''Else
'   txtothers(4).Enabled = True
'   txtothers(8).Enabled = True
'   txtothers(4).SetFocus
''End If
'
'End Sub
'
'Private Sub Void_Approval()
'
'With MSFlexGrid1
'   void = True
'   If .TextMatrix(1, 1) <> "" Then
'      .SetFocus
'      .Col = 1
'      .Row = 1
'   End If
'End With
'
'End Sub
'
'Private Sub Form_Activate()
'   oApp.MenuName = Me.Tag
'   Me.ZOrder 0
'
'   If bLoaded = False Then
'      oDriver.RecordNew
'      oDriver.DisableTextbox 0
'      oDriver.DisableTextbox 2
'      bLoaded = True
'      If txtothers(0).Enabled = True Then txtothers(0).SetFocus
'   End If
'
'End Sub
'
'Private Sub oDriver_DisableOtherControl()
'   oDriver.DisableTextbox 0
'   oDriver.DisableTextbox 2
'   txtothers(2).Enabled = False
'   txtothers(5).Enabled = False
'   txtothers(6).Enabled = False
'End Sub
'
'Private Sub oDriver_EnableOtherControl()
'   oDriver.DisableTextbox 0
'   oDriver.DisableTextbox 2
'   txtothers(2).Enabled = False
'   txtothers(5).Enabled = False
'   txtothers(6).Enabled = False
'End Sub
'
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'   Select Case KeyCode
'      Case vbKeyReturn, vbKeyDown
'         If GetFocus = MSFlexGrid1.hWnd Then Exit Sub
'         SetNextFocus
'      Case vbKeyUp
'         SetPreviousFocus
'      Case vbKeyF1   'Discount Approval
'         If MSFlexGrid1.TextMatrix(1, 1) = "" Then Exit Sub
'         Disc_Approval
'      Case vbKeyF2   'Record Save
'         If MSFlexGrid1.TextMatrix(1, 1) = "" Then Exit Sub
'         oDriver.RecordSave
'      Case vbKeyF4   'Void Transaction
'         If MSFlexGrid1.TextMatrix(1, 1) = "" Then Exit Sub
'         Void_Approval
'      Case vbKeyF5   'Cheque Transaction
'         If MSFlexGrid1.TextMatrix(1, 1) = "" Then Exit Sub
'         If txtfield(3).Text <> "0.00" Then txtfield(3).Text = "0.00"
'         frmCheque_POS.txtfield(1) = Format(txtfield(2), "#,##0.00")
'         frmCheque_POS.Transaction = "POS"
'         frmCheque_POS.Show 1
'      Case vbKeyF6   'Credit Card Transaction
'         If MSFlexGrid1.TextMatrix(1, 1) = "" Then Exit Sub
'         If txtfield(3).Text <> "0.00" Then txtfield(3).Text = "0.00"
'         frmCard_POS.txtfield(1) = Format(txtfield(2), "#,##0.00")
'         frmCard_POS.Transaction = "POS"
'         frmCard_POS.Show 1
'      Case vbKeyF7   'Installment Transaction
'         If MSFlexGrid1.TextMatrix(1, 1) = "" Then Exit Sub
'         Installment_Approval
'      Case vbKeyF8   'Sales Register
'         frmPOS_Register.Show
'      Case vbKeyF9   'Change Unit
'         frmChange_Unit.Show
'      Case vbKeyF10   'Job Order
'         frmJOMenu.Show
'      Case vbKeyF11   'Exit
'         Unload Me
'         Unload mdiMain
'      Case 27        'Bar Code, Esc
'         txtothers(0).SetFocus
'      Case 34        'Invoice
'         If MSFlexGrid1.TextMatrix(1, 1) = "" Then Exit Sub
'         txtfield(1).SetFocus
'      Case 33        'Client Info
'         If MSFlexGrid1.TextMatrix(1, 1) = "" Then Exit Sub
'         txtothers(8).SetFocus
'      Case 17        'Cash Given
'         If MSFlexGrid1.TextMatrix(1, 1) = "" Then Exit Sub
'         txtfield(3).SetFocus
'   End Select
'End Sub
'
'Private Sub Form_Load()
'
'   CenterChildForm mdiMain, Me
'   bLoaded = False
'
'   Set oDriver = New FormDriver
'   Set oDriver.AppDriver = oApp
'   Set oDriver.MainForm = Me
'
'   InitGrid
'
'   oDriver.RecQuery = "SELECT" _
'                     & " sTransNox, " _
'                     & " sSalesInv, " _
'                     & " nTranTotl, " _
'                     & " nAmtPaidx, " _
'                     & " sClientID, " _
'                     & " sCashierx, " _
'                     & " sRemarksx, " _
'                     & " dTransact, " _
'                     & " nGiftCpnx, " _
'                     & " cTranStat, " _
'                     & " sModified, " _
'                     & " dModified, " _
'                     & " vTimeStmp  " _
'               & " FROM CP_SO_Master " _
'
'   oDriver.InitRecForm
'
'   'Customer
'   oDriver.LookupQuery(4) = "SELECT" _
'                  & " a.sClientID, " _
'                  & " a.sLastName + ', ' + a.sFrstName + ' ' + a.sMiddName as xFullName, " _
'                  & " a.sAddressx + ', ' + b.sTownName as xAddressx " _
'               & " FROM Client_Master a " _
'                  & " LEFT JOIN TownCity b " _
'                     & " ON a.sTownIDxx = b.sTownIDxx " _
'               & " ORDER BY slastname, sfrstname, smiddname "
'
'   oDriver.LookupReference(4) = "xFullName»xAddressx"
'   oDriver.LookupColumn(4) = "xFullName»xAddressx"
'   oDriver.LookupTitle(4) = "Customer Name»Address"
'
'   'Sales Person
'   oDriver.LookupQuery(5) = "SELECT" _
'                  & " sEmployID, " _
'                  & " sLastName + ', ' + sFrstName + ' ' + sMiddName xFullName," _
'               & " FROM Sales_Person " _
'               & " ORDER BY slastname, sfrstname, smiddname "
'
'   oDriver.LookupReference(5) = "sEmployID»xFullName"
'   oDriver.LookupColumn(5) = "sEmployID»xFullName"
'   oDriver.LookupTitle(5) = "Emp ID»Sales Person"
'
'   oDriver.FieldStart = 1
'   oDriver.FieldFormat(0) = "@@-@@@@@@@@"
'   oDriver.FieldSize(0) = Len(oDriver.FieldFormat(0))
'
'   oDriver.FieldFormat(2) = "#,##0.00"
'   oDriver.FieldFormat(3) = "#,##0.00"
'
'   EmptyGrid
'   ShockwaveFlash1.Movie = App.Path & "\images\CP_POS.swf"
'   ShockwaveFlash1.Play
'
'End Sub
'
'Private Sub InitGrid()
'
'   With MSFlexGrid1
'      .Rows = 2
'      .Cols = 14
'      .Font = "MS Sans Serif"
'
'      'column title
'      .TextMatrix(0, 1) = "Bar Code"
'      .TextMatrix(0, 2) = "Description"
'      .TextMatrix(0, 3) = "Unit Price"
'      .TextMatrix(0, 4) = "Qty"
'      .TextMatrix(0, 5) = "%"
'      .TextMatrix(0, 6) = "% Amnt"
'      .TextMatrix(0, 7) = "Particulars"  'Load
'      .TextMatrix(0, 8) = "Stock ID"
'      .TextMatrix(0, 9) = "Sub Total"
'      .TextMatrix(0, 10) = "sSerialID"    'CP
'      .TextMatrix(0, 11) = "Load Pur Price"
'      .TextMatrix(0, 12) = "Category"
'      .TextMatrix(0, 13) = "CP No."
'      .Row = 0
'
'      'column alignment
'      For pnCtr = 0 To .Cols - 1
'      .Col = pnCtr
'      .CellFontBold = True
'      .CellAlignment = 1
'      Next
'
'      'column width
'      .ColWidth(0) = 400
'      .ColWidth(1) = 1800
'      .ColWidth(2) = 4450
'      .ColWidth(3) = 1400
'      .ColWidth(4) = 700
'      .ColWidth(5) = 700
'      .ColWidth(6) = 1000
'      .ColWidth(7) = 2000
'      .ColWidth(8) = 0
'      .ColWidth(9) = 1890
'      .ColWidth(10) = 0
'      .ColWidth(11) = 0
'      .ColWidth(12) = 0
'      .ColWidth(13) = 0
'
'      .ColAlignment(1) = 1
'      .ColAlignment(2) = 1
'      .ColAlignment(3) = 6
'      .ColAlignment(4) = 6
'      .ColAlignment(5) = 6
'      .ColAlignment(6) = 6
'      .ColAlignment(9) = 6
'      .ColAlignment(13) = 1
'
'      .Row = 0
'      .ColSel = .Cols - 1
'
'   End With
'
'End Sub
'
'Private Sub EmptyGrid()
'
'With MSFlexGrid1
'   .Rows = 2
'   For pnCtr = 1 To .Cols - 1
'      .TextMatrix(1, pnCtr) = ""
'   Next
'End With
'
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'   Set oDriver = Nothing
'   Set oSkin = Nothing
'End Sub
'
'Private Sub MSFlexGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
'Dim Entry As Integer
'
'If KeyCode = 13 Then
'   With MSFlexGrid1
'      If .TextMatrix(.Row, 1) = "" Then Exit Sub
'
'      'Clear Fields
'      txtothers(0).Text = ""
'      ClearFields
'
'      If .Rows = 3 And .Row = 1 Then 'Delete first row
'         .Rows = .Rows - 2
'         'Add 1 row
'         .Rows = .Rows + 1
'      ElseIf .Row = .Rows - 2 Then 'Delete Last Row
'         .Rows = .Rows - 2
'         'Add 1 row
'         .Rows = .Rows + 1
'         .Row = .Row + 1
'      ElseIf .Rows > 3 Then  'Adjust Rows
'         For Entry = .Row To .Rows - 2
'            .TextMatrix(Entry, 1) = .TextMatrix(Entry + 1, 1)
'            .TextMatrix(Entry, 2) = .TextMatrix(Entry + 1, 2)
'            .TextMatrix(Entry, 3) = .TextMatrix(Entry + 1, 3)
'            .TextMatrix(Entry, 4) = .TextMatrix(Entry + 1, 4)
'            .TextMatrix(Entry, 5) = .TextMatrix(Entry + 1, 5)
'            .TextMatrix(Entry, 6) = .TextMatrix(Entry + 1, 6)
'            .TextMatrix(Entry, 7) = .TextMatrix(Entry + 1, 7)
'            .TextMatrix(Entry, 8) = .TextMatrix(Entry + 1, 8)
'            .TextMatrix(Entry, 9) = .TextMatrix(Entry + 1, 9)
'            .TextMatrix(Entry, 10) = .TextMatrix(Entry + 1, 10)
'            .TextMatrix(Entry, 11) = .TextMatrix(Entry + 1, 11)
'            .TextMatrix(Entry, 12) = .TextMatrix(Entry + 1, 12)
'            'Entry No
'            If .Rows = 4 Then
'               .TextMatrix(Entry, 0) = 1
'            ElseIf .TextMatrix(Entry + 1, 0) <> "" Then
'                  .TextMatrix(Entry, 0) = .TextMatrix(Entry + 1, 0) - 1
'            End If
'         Next
'         'Delete Last row
'         .Rows = .Rows - 2
'         'Add 1 row
'         .Rows = .Rows + 1
'         .Row = .Row + 1
'      Else
'         For pnCtr = 0 To .Cols - 1
'            .TextMatrix(.Row, pnCtr) = ""
'         Next
'         .Rows = .Rows - 1
'      End If
'
'   .BackColorSel = &H80000005
'   End With
'
'   Grand_Total
'   txtothers(0).SetFocus
'End If
'
'End Sub
'Private Sub ShowGrid()
'Dim lrs As ADODB.Recordset
'
'If txtothers(0).Text <> "" Then
'   With MSFlexGrid1
'      .Rows = .Rows + 1
'
'      If .Row = 1 Then
'         .TextMatrix(.Row, 0) = 1
'      Else
'         .TextMatrix(.Row, 0) = .TextMatrix(.Row - 1, 0) + 1
'      End If
'
'      If .Rows > 9 Then
'         .ColWidth(2) = 4250
'         .ColWidth(3) = 1300
'      Else
'         .ColWidth(2) = 4450
'         .ColWidth(3) = 1400
'      End If
'
'      .TextMatrix(.Row, 1) = txtothers(0).Text
'      .TextMatrix(.Row, 2) = txtothers(1).Text
'      .TextMatrix(.Row, 3) = txtothers(2).Text
'      .TextMatrix(.Row, 4) = txtothers(3).Text
'      .TextMatrix(.Row, 5) = txtothers(4).Text
'      .TextMatrix(.Row, 6) = txtothers(8).Text
'      .TextMatrix(.Row, 7) = ""
'      .TextMatrix(.Row, 8) = txtothers(0).Tag
'      .TextMatrix(.Row, 9) = txtothers(5).Text
'      .TextMatrix(.Row, 10) = ""
'      .TextMatrix(.Row, 11) = 0
'      .TextMatrix(.Row, 12) = txtothers(1).Tag
'      .TextMatrix(.Row, 13) = ""
'
'      'Get Purchase Price
'      lsSQL = "SELECT" _
'                  & " sStockIDx, " _
'                  & " nPurPrice  " _
'          & " FROM CP_Inventory " _
'          & " WHERE sStockIDx = '" & txtothers(0).Tag & "' "
'      Set lrs = New ADODB.Recordset
'      If lrs.State = adStateOpen Then lrs.Close
'      lrs.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
'
'      If Not lrs.EOF Then
'         .TextMatrix(.Row, 11) = lrs("nPurPrice")
'      End If
'      .Row = .Rows - 1
'
'      Grand_Total
'   End With
'End If
'
'End Sub
'
'Private Sub UpdateGrid()
'Dim temp As Integer
'
'With MSFlexGrid1
'If .Row = 0 Then Exit Sub
'   If .Row - 1 = 0 Then Exit Sub
'   MsgBox .Row - 1
'   .TextMatrix(.Row - 1, 5) = txtothers(4).Text
'   .TextMatrix(.Row - 1, 6) = Format(txtothers(8).Text, "#,##0.00")
'   .TextMatrix(.Row - 1, 9) = txtothers(5).Text
'   Grand_Total
'End With
'
'End Sub
'
'Private Sub Grand_Total()
'Dim Total As Double
'
'With MSFlexGrid1
'   If .Rows <= 3 Then
'      If .Rows = 2 Then
'         txtfield(2).Text = "0.00"
'      Else
'         txtfield(2).Text = Format(.TextMatrix(1, 9), "#,##0.00")
'         Total = txtfield(2).Text
'      End If
'   Else
'      For pnCtr = 1 To .Rows - 2
'         If .TextMatrix(pnCtr, 9) <> "" Then
'            Total = Total + CDbl(.TextMatrix(pnCtr, 9))
'         End If
'      Next
'      txtfield(2).Text = Format(CDbl(Total), "#,##0.00")
'   End If
'End With
'txtothers(7).Text = Format((txtfield(2).Text) / (1.12), "#,##0.00")
'txtothers(9).Text = Format((txtfield(2).Text - txtothers(7).Text), "#,##0.00")
'End Sub
'Function Invoice_No() As String
'   Dim lrs As Recordset
'   Dim lsSQL As String
'   Dim lnCtr As Long
'
'   lsSQL = "SELECT TOP 1" & _
'            " sSalesInv" & _
'            " FROM CP_SO_Master " & _
'            " WHERE sSalesInv LIKE " & strParm(oApp.BranchCode & "SI-%") & _
'            " ORDER BY sSalesInv DESC"
'
'   Set lrs = New Recordset
'   lrs.Open lsSQL, oApp.Connection, , , adCmdText
'
'   If lrs.EOF Then
'      lnCtr = 1
'   Else
'      If Left(lrs("sSalesInv"), 2) = oApp.BranchCode Then
'         lnCtr = CLng(Right(lrs("sSalesInv"), 8)) + 1
'      Else
'         lnCtr = 1
'      End If
'   End If
'
'   Invoice_No = oApp.BranchCode & "SI-" & Format(Date, "yy") & Format(lnCtr, "00000000")
'   Set lrs = Nothing
'End Function
'
'Private Sub oDriver_InitValue()
'Dim Index As Integer
'
'oDriver.FieldReference(0) = True
'oDriver.FieldValue(0) = GetNextCode("CP_SO_Master", "sTransNox", True, oApp.Connection, True, oApp.BranchCode)
'
'lblField(1).Caption = Time()
'lblField(0).Caption = Format(oApp.ServerDate, "MMMM dd,yyyy")
'
'txtfield(0).Text = oDriver.FieldValue(0)
'If txtfield(0).Enabled = True Then txtfield(0).Enabled = False
'txtfield(1).Text = Invoice_No
'
'txtfield(2).Text = "0.00"
'txtfield(3).Text = "0.00"
'
'txtothers(0).Tag = "" 'sStockIDx
'txtothers(0).Text = ""
'txtothers(1).Tag = ""               'Category 1,2,3
'txtothers(4).Text = 0               'Discount %
'txtothers(5).Text = "0.00"          'Subtotal
'txtothers(5).Locked = True
'txtothers(6).Text = "0.00"          'Change Due
'txtothers(6).Locked = True
'txtothers(7).Text = "0.00"          'Tax
'txtothers(7).Locked = True
'txtothers(9).Text = "0.00"          'Tax
'txtothers(9).Locked = True
'
'
'lblField(17).Caption = oApp.UserName   'User
'txtothers(8).Text = "0.00"          'Discount Amt
'
'psPayment = ""
'ClearFields
'EmptyGrid
'
''Tag Values
''txtothers(0).Tag = sStockIDx
''txtothers(1).Tag = sCategIDx
'
'pbnewitem = True
'
'End Sub
'
'Private Sub oDriver_SaveComplete()
'    txtothers(0).SetFocus
'    oDriver.FieldValue(4) = ""
'End Sub
'
'Private Sub Timer1_Timer()
'   lblField(1).Caption = Time()
'End Sub
'
'Private Sub SearchBarCode(ByVal SearchValue As Boolean)
'   Dim lsSearch As String
'   Dim lrs As ADODB.Recordset
'
'   Set lrs = New ADODB.Recordset
'   lsSQL = "SELECT" _
'                & " b.sStockIDx, " _
'                & " b.sBarrcode, " _
'                & " b.sCategIDx, " _
'                & " c.sBrandNme, " _
'                & " d.sModelNme, " _
'                & " b.sDescript, " _
'                & " b.nSelPrice, " _
'                & " b.cWdSerial, " _
'                & " b.cCellLoad, " _
'                & " b.cWalletxx, " _
'                & " e.sColorNme, " _
'                & " f.nQtyOnHnd  " _
'
'   lsSQL = lsSQL _
'         & " FROM CP_Inventory b" _
'               & " LEFT JOIN Brand c " _
'                  & " ON b.sBrandIDx = c.sBrandIDx " _
'               & " LEFT JOIN Model d " _
'                  & " ON b.sModelIDx = d.sModelIDx " _
'               & " LEFT JOIN Color e" _
'                  & " ON b.sColorIDx = e.sColorIDx " _
'               & " LEFT JOIN CP_Inventory_Master f " _
'                  & " ON b.sStockIDx = f.sStockIdx " _
'         & " WHERE b.cRecdStat = '" & xeRecStateActive & "' " _
'               & " AND f.nQtyOnHnd > 0 " _
'               & " AND f.sBranchCd = '" & oApp.BranchCode & "' " _
'
'   If SearchValue Then
'      lsSQL = lsSQL & " AND sBarrCode = '" & Trim(txtothers(0).Text) & "'"
'   Else
'      lsSQL = lsSQL & " AND sBarrCode LIKE '%" & Trim(txtothers(0).Text) & "%' "
'   End If
'   lsSQL = lsSQL & " ORDER BY sBarrCode"
'   lrs.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
'      If lrs.RecordCount = 1 Then
'         txtothers(0).Tag = lrs("sStockIDx")
'         txtothers(0).Text = lrs("sBarrCode")
'         Select Case lrs("cWdSerial")
'            Case 1   'w/ Serial
'               txtothers(1).Text = Trim(IIf(IsNull(lrs("sBrandNme")), "", lrs("sBrandNme")) _
'                              & " " & IIf(IsNull(lrs("sModelNme")), "", lrs("sModelNme")) _
'                              & " " & IIf(IsNull(lrs("sDescript")), "", lrs("sDescript")) _
'                              & " " & IIf(IsNull(lrs("sColorNme")), "", lrs("sColorNme")))
'               txtothers(2).Text = Format(lrs("nSelPrice"), "#,##0.00")
'               txtothers(3).Text = 1
'               txtothers(3).Enabled = False
'               txtothers(5).Text = Format(lrs("nSelPrice"), "#,##0.00")
'               txtothers(1).Tag = 1
'               frmCP_Serial.Show 1
'            Case 0   'No Serial
'               If lrs("cCellLoad") = 1 Then  '  Load Retail Transaction
'                  txtothers(1).Text = lrs("sDescript")
'                  txtothers(3).Text = 1
'                  txtothers(3).Enabled = False
'                  txtothers(1).Tag = 2
'                  frmLoadRetail_POS.oStock = txtothers(0).Tag
'                  frmLoadRetail_POS.txtfield(0).Tag = lrs("nQtyOnHnd")
'                  frmLoadRetail_POS.Show 1
'               ElseIf lrs("cWalletxx") = 1 Then 'Load Wallet Transaction
'                  txtothers(1).Text = lrs("sDescript")
'                  txtothers(3).Text = 1
'                  txtothers(3).Enabled = False
'                  txtothers(1).Tag = 2
'                  frmLoadWallet_POS.oStock = txtothers(0).Tag
'                  frmLoadWallet_POS.Show 1
'               Else                             'Accessories
'                  txtothers(1).Text = Trim(IIf(IsNull(lrs("sBrandNme")), "", lrs("sBrandNme")) _
'                                 & " " & IIf(IsNull(lrs("sModelNme")), "", lrs("sModelNme")) _
'                                 & " " & IIf(IsNull(lrs("sDescript")), "", lrs("sDescript")) _
'                                 & " " & IIf(IsNull(lrs("sColorNme")), "", lrs("sColorNme")))
'                  txtothers(2).Text = Format(lrs("nSelPrice"), "#,##0.00")
'                  txtothers(3).Text = 1
'                  txtothers(1).Tag = 3
'                  txtothers(3).Enabled = True
'               End If
'         End Select
'
'      ElseIf lrs.RecordCount > 1 Then
'         lsSearch = KwikBrowse(oApp, lrs, _
'                        "sBarrcode»sBrandNme»sModelNme»sDescript»nSelPrice»sColorNme»nQtyOnHnd", _
'                        "Bar Code»Brand»Model»Description»Unit Price»Color»Qty", _
'                        "@»@»@»@»#,##0.00»@»#,##0")
'
'         If lsSearch <> "" Then
'            psSelected = Split(lsSearch, "»")
'            txtothers(0).Tag = psSelected(0)
'            txtothers(0).Text = psSelected(1)
'            txtothers(1).Tag = psSelected(2)
'            Select Case psSelected(7)
'               Case 1   'w/ Serial
'                  txtothers(1).Text = Trim(IIf(IsNull(psSelected(3)), "", psSelected(3)) _
'                              & " " & IIf(IsNull(psSelected(4)), "", psSelected(4)) _
'                              & " " & IIf(IsNull(psSelected(5)), "", psSelected(5)) _
'                              & " " & IIf(IsNull(psSelected(7)), "", psSelected(7)))
'                  txtothers(2).Text = Format(psSelected(6), "#,##0.00")
'                  txtothers(3).Text = 1
'                  txtothers(3).Enabled = False
'                  txtothers(5).Text = Format(psSelected(6), "#,##0.00")
'                  txtothers(1).Tag = 1
'                  frmCP_Serial.Show 1
'               Case 0   'No Serial
'                  If psSelected(8) = 1 Then     'Load Retail Transaction
'                     txtothers(1).Text = psSelected(5)
'                     txtothers(3).Text = 1
'                     txtothers(3).Enabled = False
'                     txtothers(1).Tag = 2
'                     frmLoadRetail_POS.oStock = txtothers(0).Tag
'                     frmLoadRetail_POS.txtfield(0).Tag = psSelected(11)
'                     frmLoadRetail_POS.Show 1
'                  ElseIf psSelected(9) = 1 Then 'Load Wallet Transaction
'                     txtothers(1).Text = psSelected(5)
'                     txtothers(3).Text = 1
'                     txtothers(3).Enabled = False
'                     txtothers(1).Tag = 2
'                     frmLoadWallet_POS.Show 1
'                  Else                          'Accessories
'                     txtothers(1).Text = Trim(IIf(IsNull(psSelected(3)), "", psSelected(3)) _
'                                 & " " & IIf(IsNull(psSelected(4)), "", psSelected(4)) _
'                                 & " " & IIf(IsNull(psSelected(5)), "", psSelected(5)) _
'                                 & " " & IIf(IsNull(psSelected(7)), "", psSelected(7)))
'                     txtothers(2).Text = Format(psSelected(6), "#,##0.00")
'                     txtothers(3).Text = 1
'                     txtothers(1).Tag = 3
'                     txtothers(3).Enabled = True
'                  End If
'            End Select
'         End If
'   Else
'      ClearFields
'      txtothers(0).Text = ""
'   End If
'
'   Set lrs = Nothing
'
'End Sub
'Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Dim orig As String
'Dim lsCondition As String
'Dim lsSQL As String
'
'   If KeyCode = 13 Or KeyCode = vbKeyF3 Then
'      Select Case Index
'         Case 4
'            SearchClient False
'            If txtfield(Index).Text <> "" Then SetNextFocus
'         Case 5
'            SearchSales False
'            If txtfield(Index).Text <> "" Then SetNextFocus
'      End Select
'   End If
'End Sub
'
'Private Sub txtField_GotFocus(Index As Integer)
'   oDriver.ColumnIndex = Index
'   txtfieldGotfocus = True
'   pnIndex = Index
'   txtfield(Index).BackColor = &HE1FEFF
'End Sub
'
'Private Sub txtField_LostFocus(Index As Integer)
'   txtfieldGotfocus = False
'   txtfield(Index).BackColor = &H80000005
'   If Index = 3 Then txtfield(Index).BackColor = &HC0FFFF
'   If Index = 4 Then oDriver.FieldValue(4) = psClientID
'End Sub
'
'Private Sub txtOthers_GotFocus(Index As Integer)
'   txtothers(Index).SelStart = 0
'   txtothers(Index).SelLength = Len(txtothers(Index))
'
'   txtOthersGotfocus = True
'   txtfieldGotfocus = False
'   pnIndex = Index
'   txtothers(Index).BackColor = &HE1FEFF
'End Sub
'
'Private Sub txtOthers_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'
'If KeyCode = vbKeyF3 Or KeyCode = 13 Then
'   Select Case Index
'      Case 0
'         ClearFields
'         SearchBarCode False
'      Case 3
'         If Not IsNumeric(txtothers(Index).Text) Then txtothers(Index).Text = 1
'         If txtothers(3).Text = 0# Then
'            MsgBox "Quantity can't be zero!!!", vbCritical, "Warning"
'            txtothers(3).SetFocus
'         Else
'            txtothers(Index).Text = Format(txtothers(Index).Text, "#,##0")
'            txtothers(5).Text = Format((txtothers(2).Text * txtothers(3).Text), "#,##0.00")
'            ShowGrid
'         End If
'   End Select
'   KeyCode = 0
'End If
'
'End Sub
'
'Private Sub txtOthers_LostFocus(Index As Integer)
'Dim temp As Double
'Dim wdisc As Double
'
'   Select Case Index
'
'   Case 3
'      If Not IsNumeric(txtothers(Index).Text) Then txtothers(Index).Text = 1
'      If txtothers(3).Text = 0# Then
'         MsgBox "Quantity can't be zero!!!", vbCritical, "Warning"
'         txtothers(3).SetFocus
'      Else
'         txtothers(Index).Text = Format(txtothers(Index).Text, "#,##0")
'         txtothers(5).Text = Format((txtothers(2).Text * txtothers(3).Text), "#,##0.00")
'         txtothers(0).SetFocus
'      End If
'
'   Case 4
'      If Not IsNumeric(txtothers(Index).Text) Then txtothers(Index).Text = 0#
'         If txtothers(Index).Text <> 0 Then
'            txtothers(5).Text = Format(CDbl((txtothers(2).Text * txtothers(3).Text) - _
'                                 txtothers(8).Text), "#,##0.00")
'         ElseIf txtothers(4).Text <> 0 Then
'         wdisc = CDbl(txtothers(2).Text - txtothers(2).Text * (txtothers(Index).Text / 100))
'            txtothers(8).Text = wdisc
'            txtothers(5).Text = Format(CDbl(txtothers(3).Text * (txtothers(2).Text - wdisc)), _
'                              "#,##0.00")
'         End If
'         txtothers(8).Text = Format(txtothers(8).Text, "#,##0.00")
'         UpdateGrid
'
'   Case 5
'      ShowGrid
'
'   Case 8
'      If Not IsNumeric(txtothers(Index).Text) Then txtothers(Index).Text = 0#
'         If txtothers(Index).Text <> 0 Then
'            txtothers(5).Text = Format(CDbl((txtothers(2).Text * txtothers(3).Text) - _
'                                 txtothers(8).Text), "#,##0.00")
'         ElseIf txtothers(4).Text <> 0 Then
'            wdisc = CDbl(txtothers(2).Text * (txtothers(4).Text / 100))
'            txtothers(8).Text = wdisc
'            txtothers(5).Text = Format(CDbl(txtothers(3).Text * (txtothers(2).Text - wdisc)), _
'                              "#,##0.00")
'         End If
'         txtothers(8).Text = Format(txtothers(8).Text, "#,##0.00")
'         UpdateGrid
'
'   End Select
'   txtothers(Index).BackColor = &H80000005
'   txtOthersGotfocus = False
'
'End Sub
'
'Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
'Dim temp As Double
'   Select Case Index
'      Case 3
'      If Not IsNumeric(txtfield(Index).Text) Then txtfield(Index).Text = 0#
'         txtfield(Index).Text = Format(txtfield(Index).Text, "#,##0.00")
'         If txtfield(Index).Text <> 0 Then
'            If CDbl(txtfield(2).Text) > CDbl(txtfield(3).Text) _
'               And cmdButton(2).Tag = "" Then
'               MsgBox "Amount Greater than Cash Given", vbCritical, "Warning"
'               txtfield(Index).SetFocus
'            Else
'               temp = CDbl(txtfield(Index).Text) - CDbl(txtfield(2).Text)
'               txtothers(6).Text = Format(temp, "#,##0.00")
'               cmdButton(4).SetFocus
'            End If
'         End If
'   End Select
'   txtfield(Index).BackColor = &HFFFFFF
'   txtfield(Index).Text = TitleCase(txtfield(Index).Text)
'End Sub
'
'Private Sub ClearFields()
'   txtothers(1).Text = ""
'   txtothers(1).Tag = ""
'   txtothers(2).Text = "0.00"
'   txtothers(3).Text = "1"
'   txtothers(4).Text = "0"
'   txtothers(5).Text = "0.00"
'   txtothers(6).Text = "0.00"
'   txtothers(7).Text = "0.00"
'   txtothers(8).Text = "0.00"
'   txtothers(9).Text = "0.00"
'End Sub
'
'Private Sub MSFlexGrid1_GotFocus()
'Dim lnCtr As Integer
'
'With MSFlexGrid1
'   For lnCtr = 1 To .Cols - 1
'      .ColSel = lnCtr
'   Next
'   .BackColorSel = &HC0FFFF
'End With
'
'End Sub
'
'Private Sub MSFlexGrid1_LostFocus()
'
'With MSFlexGrid1
'   .BackColorSel = &H80000005
'   void = False
'End With
'
'End Sub
'
'Private Sub SearchClient(ByVal SearchValue As Boolean)
'Dim lsSearch As String
'Dim lrs As ADODB.Recordset
'Dim lsSQL As String
'
'   Set lrs = New ADODB.Recordset
'
'   lsSQL = "SELECT" _
'               & " a.sClientID, " _
'               & " a.sLastName + ', ' + a.sFrstName + ' ' + a.sMiddName as xFullName," _
'               & " a.sAddressx + ', ' + b.sTownName as xAddressx" _
'            & " FROM Client_Master a " _
'               & " LEFT JOIN TownCity b " _
'                  & " ON a.sTownIDxx = b.sTownIDxx " _
'
'   If SearchValue Then
'      lsSQL = lsSQL & " WHERE sLastName + ', ' + sFrstName + ' ' + sMiddName = '" & txtfield(4).Text & "'"
'   Else
'      lsSQL = lsSQL & " WHERE sLastName + ', ' + sFrstName + ' ' + sMiddName LIKE '" & txtfield(4).Text & "%' "
'   End If
'   lsSQL = lsSQL & " ORDER BY sLastName + ', ' + sFrstName + ' ' + sMiddName"
'   lrs.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
'
'
'   If lrs.RecordCount = 1 Then
'      psClientID = lrs("sClientID")
'      txtfield(4).Text = lrs("xFullName")
'
'   ElseIf lrs.RecordCount > 1 Then
'        lsSearch = KwikBrowse(oApp, lrs, _
'                          "sClientID»xFullName»xAddressx", _
'                          "Client ID»Name»Address")
'
'        If lsSearch <> "" Then
'            psSelected = Split(lsSearch, "»")
'            psClientID = psSelected(0)
'            txtfield(4).Text = psSelected(1)
'        Else
'            psClientID = ""
'            txtfield(4).Text = ""
'            txtfield(4).SetFocus
'        End If
'   Else
'      frmCustomer.Client = txtfield(4).Text
'      frmCustomer.Show 1
'   End If
'
'   Set lrs = Nothing
'
'End Sub
'
'Private Sub SearchSales(ByVal SearchValue As Boolean)
'Dim lsSearch As String
'Dim lrs As ADODB.Recordset
'Dim lsSQL As String
'
'   Set lrs = New ADODB.Recordset
'
'   lsSQL = "SELECT" _
'               & " sEmployID, " _
'               & " sFrstName + ' ' +  sLastName as xFullName " _
'            & " FROM Sales_Person " _
'
'   If SearchValue Then
'      lsSQL = lsSQL & " WHERE sFrstName = '" & txtfield(5).Text & "'"
'   Else
'      lsSQL = lsSQL & " WHERE sFrstName LIKE '%" & txtfield(5).Text & "%' "
'   End If
'   lsSQL = lsSQL & " ORDER BY xFullName "
'   lrs.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
'
'   If lrs.RecordCount = 1 Then
'      oDriver.FieldValue(5) = lrs("sEmployID")
'      txtfield(5).Text = lrs("xFullName")
'
'   ElseIf lrs.RecordCount > 1 Then
'        lsSearch = KwikBrowse(oApp, lrs, _
'                          "sEmployID»xFullName", _
'                          "Employee ID»Name")
'        If lsSearch <> "" Then
'            psSelected = Split(lsSearch, "»")
'            oDriver.FieldValue(5) = psSelected(0)
'            txtfield(5).Text = psSelected(1)
'        End If
'   ElseIf lrs.RecordCount = 0 Then
'      oDriver.FieldValue(5) = ""
'      txtfield(5).Text = ""
'   End If
'
'   Set lrs = Nothing
'
'End Sub
'
'Private Sub oDriver_WillSave(Cancel As Boolean)
'   If (CDbl(txtfield(2).Text) > CDbl(txtfield(3).Text)) _
'         And psPayment = "" Then
'            MsgBox "Amount Greater than Cash Given!!!", vbCritical, "Warning"
'            txtfield(3).SetFocus
'            Cancel = True
'   ElseIf oDriver.FieldValue(4) = "" _
'         And (psPayment <> "") Then
'            MsgBox "Invalid Client Detected!!!", vbCritical, "Warning"
'            txtfield(4).SetFocus
'            Cancel = True
'   ElseIf txtfield(2).Text = 0# And txtothers(8).Text = 0# Then
'            MsgBox "Invalid Amount Detected!!!", vbCritical, "Warning"
'            txtfield(2).SetFocus
'            Cancel = True
'   Else
'      Cancel = Not SaveCP_SODetail
'         If Cancel Then Exit Sub
'      Cancel = Not UpdateCP_Serial
'         If Cancel Then Exit Sub
'      Cancel = Not SaveCPInventory
'         If Cancel Then Exit Sub
'      Cancel = Not SaveLoadLedger
'         If Cancel Then Exit Sub
'
'      oDriver.FieldValue(1) = txtfield(1).Text
'      oDriver.FieldValue(2) = CDbl(txtfield(2).Text)
'      oDriver.FieldValue(6) = txtfield(6).Text
'      oDriver.FieldValue(7) = Format(oApp.ServerDate, "MM/dd/yyyy")
'      oDriver.FieldValue(8) = 0
'      If txtfield(4).Text = "" Then oDriver.FieldValue(4) = ""
'
'      'Mode of Payment
'      Select Case psPayment
'         Case "Credit"
'            If frmCard_POS.txtfield(0).Tag = "" Or frmCard_POS.txtfield(5).Text = "" Then
'               Cancel = True
'            Else
'               Cancel = Not SaveCP_SOCredit
'                  If Cancel Then Exit Sub
'               oDriver.FieldValue(9) = 1
'               oDriver.FieldValue(3) = CDbl(frmCard_POS.txtfield(2).Text)
'               Unload frmCard_POS
'            End If
'         Case "Cheque"
'            Cancel = Not SaveCP_SOCheque
'               If Cancel Then Exit Sub
'            oDriver.FieldValue(9) = 2
'            oDriver.FieldValue(3) = CDbl(frmCheque_POS.txtfield(2).Text)
'            Unload frmCheque_POS
'         Case "Installment"
'            Cancel = Not SaveCP_SOInstallment
'               If Cancel Then Exit Sub
'            oDriver.FieldValue(9) = 3
'            oDriver.FieldValue(3) = CDbl(frmInstallment_POS.txtfield(5).Text)
'            Unload frmInstallment_POS
'         Case Else
'            oDriver.FieldValue(9) = 0
'            oDriver.FieldValue(3) = CDbl(txtfield(2).Text)
'      End Select
'   End If
'
'End Sub
'
''Add Record CP_POSDetail
'Private Function SaveCP_SODetail() As Boolean
'Dim lnCtr As Integer
'Dim lnRow As Long
'
'SaveCP_SODetail = True
On Error Goto errProc
'
'   With MSFlexGrid1
'      For lnCtr = 1 To .Rows - 2
'         lsSQL = "INSERT INTO CP_SO_Detail " _
'                  & "( sTransNox, " _
'                  & "  nEntryNox, " _
'                  & "  sStockIDx, " _
'                  & "  nQuantity, " _
'                  & "  nPurPrice, " _
'                  & "  nUnitPrce, " _
'                  & "  nDiscount, " _
'                  & "  nDiscAmnt, " _
'                  & "  nSubTotal, " _
'                  & "  dModified) " _
'                     & "VALUES " _
'                        & "('" & oDriver.FieldValue(0) & "', " _
'                        & "'" & .TextMatrix(lnCtr, 0) & "', " _
'                        & "'" & .TextMatrix(lnCtr, 8) & "', " _
'                        & "'" & CLng(.TextMatrix(lnCtr, 4)) & "', " _
'                        & "'" & CDbl(.TextMatrix(lnCtr, 11)) & "', " _
'                        & "'" & CDbl(.TextMatrix(lnCtr, 3)) & "', " _
'                        & "'" & CLng(.TextMatrix(lnCtr, 5)) & "', " _
'                        & "'" & CDbl(.TextMatrix(lnCtr, 6)) & "', " _
'                        & "'" & CDbl(.TextMatrix(lnCtr, 9)) & "', " _
'                        & " getdate())"
'
'         oApp.Connection.Execute lsSQL, lnRow, adCmdText
'
'         If lnRow <= 0 Then
'            MsgBox "Unable to Save CP_SODetail!!!", vbCritical, "Warning"
'            SaveCP_SODetail = False
'            GoTo endProc
'         End If
'      Next
'   End With
'
'endProc:
'   Exit Function
'errProc:
'   SaveCP_SODetail = False
'   MsgBox Err.Description, vbCritical, "Warning"
'End Function
'
'Private Function UpdateCP_Serial() As Boolean
'Dim lnCtr As Integer
'Dim lnRow As Long
'Dim lnEntry As Integer
'
'UpdateCP_Serial = True
On Error Goto errProc
'
'   With MSFlexGrid1
'      For lnCtr = 1 To .Rows - 2
'         If .TextMatrix(lnCtr, 12) = 1 Then
'
'            'Get Last Entry No
'            lnEntry = getIMEIEntry("'" & .TextMatrix(lnCtr, 10) & "'")
'
'            'CP_Serial_Ledger
'            lsSQL = "INSERT INTO CP_Serial_Ledger" _
'                        & "( sSerialID ," _
'                        & "  sBranchcd ," _
'                        & "  dTransact ," _
'                        & "  nEntryNox ," _
'                        & "  sSourceCd ," _
'                        & "  sSourceNo ," _
'                        & "  cSoldStat ," _
'                        & "  cLocation ," _
'                        & "  dModified) " _
'                           & " VALUES " _
'                              & "('" & .TextMatrix(lnCtr, 10) & "', " _
'                              & " '" & oApp.BranchCode & "', " _
'                              & " '" & oApp.ServerDate & "', " _
'                              & " '" & lnEntry & "', " _
'                              & " 'CPSl', " _
'                              & " '" & oDriver.FieldValue(0) & "', " _
'                              & " '1', " _
'                              & " '2', " _
'                              & " getdate())"
'            oApp.Connection.Execute lsSQL, lnRow, adCmdText
'
'            'CP_SO_Serial
'            lsSQL = "INSERT INTO CP_SO_Serial" _
'                        & "( sTransNox ," _
'                        & "  nEntryNox ," _
'                        & "  sSerialID ," _
'                        & "  dModified) " _
'                           & " VALUES " _
'                              & "('" & oDriver.FieldValue(0) & "', " _
'                              & " '" & .TextMatrix(lnCtr, 0) & "', " _
'                              & " '" & .TextMatrix(lnCtr, 10) & "', " _
'                              & " getdate())"
'            oApp.Connection.Execute lsSQL, lnRow, adCmdText
'
'            'Update Location, CP_Serial_Master
'            lsSQL = "UPDATE CP_Serial_Master SET" _
'                  & " cSoldStat = '1', " _
'                  & " cLocation = '2', " _
'                  & " sClientID = '" & oDriver.FieldValue(4) & "' , " _
'                  & " sModified = '" & Encrypt(oApp.UserID) & "', " _
'                  & " dModified = getdate() " _
'            & " WHERE sSerialID = '" & .TextMatrix(lnCtr, 10) & "' "
'            oApp.Connection.Execute lsSQL, lnRow, adCmdText
'
'            If lnRow <= 0 Then
'               MsgBox "Unable to Save CP_Serial!!!", vbCritical, "Warning"
'               UpdateCP_Serial = False
'               GoTo endProc
'            End If
'         End If
'      Next
'      Set oRS = Nothing
'   End With
'
'endProc:
'   Exit Function
'errProc:
'   UpdateCP_Serial = False
'   MsgBox Err.Description, vbCritical, "Warning"
'End Function
'
''For CP Trans Only
'Private Function SaveCPInventory() As Boolean
'Dim lnRow As Long
'Dim lnEntry As Integer
'Dim QOH As Integer
'Dim lnCtr As Integer
'
'SaveCPInventory = True
On Error Goto errProc
'
'   With MSFlexGrid1
'
'      For lnCtr = 1 To .Rows - 2
'         If Trim(.TextMatrix(lnCtr, 12)) <> 2 Then
'
'         'Search sSourceNo
'         lsSQL = "SELECT" _
'                  & " sStockIDx, " _
'                  & " sSourceNo  " _
'               & " FROM CP_Inventory_Ledger " _
'               & " WHERE sStockIdx = '" & .TextMatrix(lnCtr, 8) & "'" _
'                  & " AND sSourceNo = '" & txtfield(0).Text & "'" _
'                  & " AND sSourceCd = 'CPSl' " _
'                  & " AND sBranchCd = '" & oApp.BranchCode & "'"
'         If oRS.State = adStateOpen Then oRS.Close
'         oRS.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
'
'         'Get QOH
'         QOH = getQuantity("'" & .TextMatrix(lnCtr, 8) & "'", "'" & oApp.BranchCode & "'") _
'                  - .TextMatrix(lnCtr, 4)
'
'            If oRS.EOF = False Then
'               'Update Record, CP_Inventory_Ledger
'               lsSQL = "UPDATE CP_Inventory_Ledger SET" _
'                        & " nQtyOutxx = nQtyOutxx + '" & CLng(.TextMatrix(lnCtr, 4)) & "', " _
'                        & " nQtyOnHnd = '" & CLng(QOH) & "'," _
'                        & " dmodified = getdate() " _
'                  & " WHERE sStockIdx = '" & .TextMatrix(lnCtr, 8) & "'" _
'                     & " AND sSourceNo = '" & txtfield(0).Text & "'" _
'                     & " AND sSourceCd = 'CPSl' " _
'                     & " AND sBranchCd = '" & oApp.BranchCode & "'"
'               oApp.Connection.Execute lsSQL, lnRow, adCmdText
'
'            Else
'               'Get last Entry No.
'               lnEntry = getEntryNo("CP_Inventory_Ledger", "'" & .TextMatrix(lnCtr, 8) & "'", "'" & oApp.BranchCode & "'")
'
'               'Add Record, CP_Inventory_Ledger
'               lsSQL = "INSERT INTO CP_Inventory_Ledger " _
'                     & "( sStockIDx, " _
'                     & "  sBranchCd, " _
'                     & "  sLocation, " _
'                     & "  sSourceCd, " _
'                     & "  sSourceNo, " _
'                     & "  nQtyInxxx, " _
'                     & "  nQtyOutxx, " _
'                     & "  nQtyOnHnd, " _
'                     & "  nEntryNox, " _
'                     & "  dTransact, " _
'                     & "  dModified) " _
'               & "VALUES " _
'                     & "('" & .TextMatrix(lnCtr, 8) & "', " _
'                     & "'" & oApp.BranchCode & " ', " _
'                     & "'" & oApp.BranchCode & " ', " _
'                     & "'CPSl' , " _
'                     & "'" & txtfield(0).Text & "', " _
'                     & "'0', " _
'                     & "'" & CLng(.TextMatrix(lnCtr, 4)) & "', " _
'                     & "'" & CLng(QOH) & "', " _
'                     & "'" & lnEntry & "', " _
'                     & "'" & oApp.ServerDate & "', " _
'                     & " getdate())"
'               oApp.Connection.Execute lsSQL, lnRow, adCmdText
'
'            End If
'
'            'Update QOH, CP_Inventory_Master
'            lsSQL = "UPDATE CP_Inventory_Master SET" _
'                  & " nQtyOnHnd = '" & CLng(QOH) & "', " _
'                  & " sModified = '" & Encrypt(oApp.UserID) & "', " _
'                  & " dModified = getdate() " _
'            & " WHERE sStockIDx = '" & .TextMatrix(lnCtr, 8) & "' " _
'                  & " And sBranchCd = '" & oApp.BranchCode & "' "
'            oApp.Connection.Execute lsSQL, lnRow, adCmdText
'
'            If lnRow <= 0 Then
'               MsgBox "Unable to Update CP_Inventory_Master!!!", vbCritical, "Warning"
'               SaveCPInventory = False
'               GoTo endProc
'            End If
'
'         End If
'      Next
'   End With
'
'endProc:
'   Exit Function
'errProc:
'   SaveCPInventory = False
'   MsgBox Err.Description, vbCritical, "Warning"
'
'End Function
'
''Add Record to ELoad_Ledger
'Private Function SaveLoadLedger() As Boolean
'Dim lnRow As Long
'Dim QOH As Double
'Dim lnCtr As Integer
'Dim lnEntry As Integer
'
'SaveLoadLedger = True
On Error Goto errProc
'
'   With MSFlexGrid1
'
'      'Get Last Entry No
'      lnEntry = getEntryNo("ELoad_Ledger", "'" & .TextMatrix(1, 8) & "'", _
'                  "'" & oApp.BranchCode & "'")
'
'      For lnCtr = 1 To .Rows - 2
'         If Trim(.TextMatrix(lnCtr, 12)) = 2 Then
'
'            'Get QOH
'            QOH = getQuantity("'" & .TextMatrix(lnCtr, 8) & "'", "'" & oApp.BranchCode & "'") _
'                     - .TextMatrix(lnCtr, 11)
'
'            'Add Record, ELoad_Ledger
'            lsSQL = "INSERT INTO ELoad_Ledger " _
'                        & "( sStockIDx, " _
'                        & "  sBranchCd, " _
'                        & "  sLocation, " _
'                        & "  dTransact, " _
'                        & "  sReferNox, " _
'                        & "  sPhoneNum, " _
'                        & "  sSourceCd, " _
'                        & "  sSourceNo, " _
'                        & "  sTransNox, " _
'                        & "  nQtyInxxx, " _
'                        & "  nQtyOutxx, " _
'                        & "  nEntryNox, " _
'                        & "  nQtyOnHnd, " _
'                        & "  sModified, " _
'                        & "  dModified) "
'            lsSQL = lsSQL _
'                     & "VALUES " _
'                        & "('" & .TextMatrix(lnCtr, 8) & "' ," _
'                        & "'" & oApp.BranchCode & "', " _
'                        & "'" & oApp.BranchCode & "', " _
'                        & "'" & oApp.ServerDate & "', " _
'                        & "'" & .TextMatrix(lnCtr, 7) & "', " _
'                        & "'" & .TextMatrix(lnCtr, 13) & "', " _
'                        & " 'CPSl', " _
'                        & "'" & oDriver.FieldValue(0) & "', " _
'                        & "'" & .TextMatrix(lnCtr, 0) & "', " _
'                        & "'0', " _
'                        & "'" & CDbl(.TextMatrix(lnCtr, 11)) & "', " _
'                        & "'" & lnEntry & "', " _
'                        & "'" & CDbl(QOH) & "', " _
'                        & "'" & Encrypt(oApp.UserID) & "'," _
'                        & " getdate())"
'            oApp.Connection.Execute lsSQL, lnRow, adCmdText
'
'            'Update QOH, CP_Inventory_Master
'            lsSQL = "UPDATE CP_Inventory_Master SET" _
'                  & " nQtyOnHnd = '" & CDbl(QOH) & "', " _
'                  & " sModified = '" & Encrypt(oApp.UserID) & "', " _
'                  & " dModified = getdate() " _
'            & " WHERE sStockIDx = '" & .TextMatrix(lnCtr, 8) & "' " _
'                  & " And sBranchCd = '" & oApp.BranchCode & "' "
'            oApp.Connection.Execute lsSQL, lnRow, adCmdText
'
'            If lnRow <= 0 Then
'               MsgBox "Unable to Save ELoad_Ledger!!!", vbCritical, "Warning"
'               SaveLoadLedger = False
'               GoTo endProc
'            End If
'
'         End If
'      Next
'   End With
'   Unload frmLoadRetail_POS
'   Unload frmLoadWallet_POS
'
'endProc:
'   Exit Function
'errProc:
'   SaveLoadLedger = False
'   MsgBox Err.Description, vbCritical, "Warning"
'End Function
''Credit Card Transaction
'Private Function SaveCP_SOCredit() As Boolean
'Dim lnCtr As Integer
'Dim lnRow As Long
'
'SaveCP_SOCredit = True
On Error Goto errProc
'
'   With frmCard_POS
'      lsSQL = "INSERT INTO CP_SO_Credit " _
'               & "(   sTransNox, " _
'                  & " dTransact, " _
'                  & " sClientID, " _
'                  & " sCreditID, " _
'                  & " nCashTotl, " _
'                  & " nTranTotl, " _
'                  & " nCashAmnt, " _
'                  & " nCardAmnt, " _
'                  & " sSalesInv, " _
'                  & " sAcctNmbr, " _
'                  & " nPercentx, " _
'                  & " dModified) "
'      lsSQL = lsSQL _
'                     & "VALUES " _
'                        & "('" & oDriver.FieldValue(0) & "', " _
'                        & "'" & oApp.ServerDate & "', " _
'                        & "'" & oDriver.FieldValue(4) & "', " _
'                        & "'" & .txtfield(0).Tag & "', " _
'                        & "'" & CDbl(MSFlexGrid1.TextMatrix(1, 3)) & "', " _
'                        & "'" & CDbl(txtfield(2).Text) & "', " _
'                        & "'" & CDbl(.txtfield(2).Text) & "', " _
'                        & "'" & CDbl(.txtfield(3).Text) & "', " _
'                        & "'" & txtfield(1).Text & "', " _
'                        & "'" & .txtfield(5).Text & "', " _
'                        & "'" & .txtothers(0).Text & "', " _
'                        & " getdate())"
'      oApp.Connection.Execute lsSQL, lnRow, adCmdText
'
'      If lnRow <= 0 Then
'         SaveCP_SOCredit = False
'      End If
'   End With
'
'endProc:
'   Exit Function
'errProc:
'   SaveCP_SOCredit = False
'   MsgBox Err.Description, vbCritical, "Warning"
'
'End Function
'
''Cheque Transaction
'Private Function SaveCP_SOCheque() As Boolean
'Dim lnRow As Long
'
'SaveCP_SOCheque = True
On Error Goto errProc
'
'   With frmCheque_POS
'      lsSQL = "INSERT INTO CP_SO_Cheque " _
'               & "(   sTransNox, " _
'                  & " dTransact ," _
'                  & " sClientID ," _
'                  & " sBankIDxx ," _
'                  & " nTranTotl ," _
'                  & " nCashAmnt ," _
'                  & " nCheqAmnt ," _
'                  & " sAccntNum ," _
'                  & " sSalesInv ," _
'                  & " dModified) " _
'                     & "VALUES " _
'                        & "('" & oDriver.FieldValue(0) & "', " _
'                        & "'" & oApp.ServerDate & "', " _
'                        & "'" & oDriver.FieldValue(4) & "', " _
'                        & "'" & .txtfield(0).Tag & "', " _
'                        & "'" & CDbl(txtfield(2).Text) & "', " _
'                        & "'" & CDbl(.txtfield(2).Text) & "', " _
'                        & "'" & CDbl(.txtfield(3).Text) & "', " _
'                        & "'" & .txtfield(5).Text & "', " _
'                        & "'" & txtfield(1).Text & "', " _
'                        & " getdate())"
'      oApp.Connection.Execute lsSQL, lnRow, adCmdText
'
'      If lnRow <= 0 Then
'   '      MsgBox "Unable to Save CP_SO_Cheque!!!", vbCritical, "Warning"
'         SaveCP_SOCheque = False
'      End If
'   End With
'endProc:
'   Exit Function
'errProc:
'   SaveCP_SOCheque = False
'   MsgBox Err.Description, vbCritical, "Warning"
'
'
'End Function
'
''Installment Transaction
'Private Function SaveCP_SOInstallment() As Boolean
'Dim lnRow As Long
'
'SaveCP_SOInstallment = True
On Error Goto errProc
'
'   With frmInstallment_POS
'      lsSQL = "INSERT INTO CP_SO_Installment " _
'               & "(   sTransNox, " _
'                  & " dTransact ," _
'                  & " sClientID ," _
'                  & " nTranTotl ," _
'                  & " nDownPaym ," _
'                  & " nBalancex ," _
'                  & " nPaymTerm ," _
'                  & " nMonthlyP ," _
'                  & " sSalesInv ," _
'                  & " dModified) " _
'                     & "VALUES " _
'                        & "('" & oDriver.FieldValue(0) & "', " _
'                        & "'" & oApp.ServerDate & "', " _
'                        & "'" & oDriver.FieldValue(4) & "', " _
'                        & "'" & CDbl(txtfield(2).Text) & "', " _
'                        & "'" & CDbl(.txtfield(1).Text) & "', " _
'                        & "'" & CDbl(.txtfield(2).Text) & "', " _
'                        & "'" & CLng(.txtfield(3).Text) & "', " _
'                        & "'" & CDbl(.txtfield(4).Text) & "', " _
'                        & "'" & txtfield(1).Text & "', " _
'                        & " getdate())"
'      oApp.Connection.Execute lsSQL, lnRow, adCmdText
'
'      If lnRow <= 0 Then
'         SaveCP_SOInstallment = False
'      End If
'   End With
'endProc:
'   Exit Function
'errProc:
'   SaveCP_SOInstallment = False
'   MsgBox Err.Description, vbCritical, "Warning"
'
'End Function
'
'
'
Private Sub cmdButton_Click(Index As Integer)

End Sub
