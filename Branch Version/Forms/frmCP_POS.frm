VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmCP_POS 
   BackColor       =   &H000080FF&
   BorderStyle     =   0  'None
   Caption         =   "frmDASerial"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCP_POS.frx":0000
   ScaleHeight     =   768
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtField 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Index           =   8
      Left            =   3555
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   9855
      Width           =   5475
   End
   Begin VB.TextBox txtField 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   7
      Left            =   3555
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   9405
      Width           =   5475
   End
   Begin VB.TextBox txtField 
      Alignment       =   1  'Right Justify
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
      Height          =   420
      Index           =   6
      Left            =   13710
      TabIndex        =   14
      Text            =   "0.00"
      Top             =   9570
      Width           =   1395
   End
   Begin VB.TextBox txtField 
      Alignment       =   1  'Right Justify
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
      Height          =   420
      Index           =   5
      Left            =   11085
      TabIndex        =   12
      Text            =   "00.00%"
      Top             =   9570
      Width           =   990
   End
   Begin VB.PictureBox Picture2 
      Height          =   2115
      Left            =   135
      ScaleHeight     =   2055
      ScaleWidth      =   2025
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   9255
      Width           =   2085
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Image Here"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Left            =   180
         TabIndex        =   34
         Top             =   330
         Width           =   1710
      End
   End
   Begin VB.TextBox txtField 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   3
      Left            =   3540
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   10605
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   30
      Top             =   30
   End
   Begin VB.TextBox txtField 
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
      Height          =   450
      Index           =   0
      Left            =   375
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "0000-000000"
      Top             =   1530
      Width           =   1875
   End
   Begin VB.TextBox txtField 
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
      Height          =   450
      Index           =   1
      Left            =   375
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   2265
      Width           =   1875
   End
   Begin VB.TextBox txtField 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   4
      Left            =   7095
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   10605
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Height          =   915
      Left            =   2550
      ScaleHeight     =   915
      ScaleWidth      =   12450
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1290
      Width           =   12450
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Index           =   2
         Left            =   3105
         TabIndex        =   8
         Text            =   "000000"
         Top             =   75
         Width           =   9270
      End
      Begin VB.Label lblUnitPrice 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0013B8FD&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00,000.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   29.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   75
         TabIndex        =   5
         Top             =   75
         Width           =   2970
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GridEditor1 
      Height          =   6285
      Left            =   2460
      TabIndex        =   6
      TabStop         =   0   'False
      Tag             =   "et0;eb0;et0;bc2"
      Top             =   2910
      Width           =   12645
      _ExtentX        =   22304
      _ExtentY        =   11086
      _Version        =   393216
      BackColor       =   16777215
      BackColorSel    =   9554414
      ForeColorSel    =   -2147483631
      FocusRect       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "C. Add.:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Index           =   8
      Left            =   2550
      TabIndex        =   41
      Top             =   9960
      Width           =   960
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "C. Name:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Index           =   4
      Left            =   2550
      TabIndex        =   40
      Top             =   9450
      Width           =   960
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackColor       =   &H001778E7&
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Person:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Index           =   3
      Left            =   5505
      TabIndex        =   37
      Top             =   10680
      Width           =   1485
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cashier:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Index           =   1
      Left            =   2550
      TabIndex        =   36
      Top             =   10665
      Width           =   960
   End
   Begin VB.Label lblField 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Price"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   2
      Left            =   2625
      TabIndex        =   4
      Top             =   2310
      Width           =   2895
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H001778E7&
      BorderWidth     =   2
      Height          =   525
      Index           =   2
      Left            =   150
      Top             =   4605
      Width           =   2070
   End
   Begin VB.Label lblButton 
      BackStyle       =   0  'Transparent
      Caption         =   "F4-"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Index           =   3
      Left            =   285
      TabIndex        =   35
      Top             =   4725
      Width           =   1830
   End
   Begin VB.Label lblTotalAmount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000,000.00"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   645
      Left            =   12105
      TabIndex        =   16
      Top             =   10515
      Width           =   3000
   End
   Begin VB.Label lblField 
      BackColor       =   &H001778E7&
      BackStyle       =   0  'Transparent
      Caption         =   "Disc. Amount:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   345
      Index           =   6
      Left            =   12180
      TabIndex        =   13
      Top             =   9660
      Width           =   1470
   End
   Begin VB.Label lblField 
      BackColor       =   &H001778E7&
      BackStyle       =   0  'Transparent
      Caption         =   "Discount Rate:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   345
      Index           =   5
      Left            =   9510
      TabIndex        =   11
      Top             =   9630
      Width           =   1515
   End
   Begin VB.Line Line7 
      BorderColor     =   &H0013B8FD&
      BorderWidth     =   4
      Index           =   1
      X1              =   619
      X2              =   619
      Y1              =   619
      Y2              =   755
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0013B8FD&
      BorderWidth     =   4
      Index           =   1
      X1              =   618
      X2              =   1014
      Y1              =   620
      Y2              =   620
   End
   Begin VB.Label lblField 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "&Barcode"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   24
      Left            =   5640
      TabIndex        =   7
      Top             =   2310
      Width           =   9360
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0013B8FD&
      BorderWidth     =   2
      FillColor       =   &H000080FF&
      Height          =   1620
      Index           =   1
      Left            =   285
      Top             =   1200
      Width           =   2085
   End
   Begin VB.Label lblButton 
      BackStyle       =   0  'Transparent
      Caption         =   "ESC"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   11
      Left            =   285
      TabIndex        =   28
      Top             =   8775
      Width           =   1830
   End
   Begin VB.Shape Shape11 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Height          =   525
      Left            =   150
      Top             =   8685
      Width           =   2070
   End
   Begin VB.Shape Shape10 
      BorderColor     =   &H001778E7&
      BorderWidth     =   2
      Height          =   525
      Index           =   1
      Left            =   150
      Top             =   8175
      Width           =   2070
   End
   Begin VB.Shape Shape9 
      BorderColor     =   &H001778E7&
      BorderWidth     =   2
      Height          =   525
      Index           =   1
      Left            =   150
      Top             =   7665
      Width           =   2070
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H001778E7&
      BorderWidth     =   2
      Height          =   525
      Index           =   1
      Left            =   150
      Top             =   7155
      Width           =   2070
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H001778E7&
      BorderWidth     =   2
      Height          =   525
      Index           =   1
      Left            =   150
      Top             =   6645
      Width           =   2070
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H001778E7&
      BorderWidth     =   2
      Height          =   525
      Index           =   1
      Left            =   150
      Top             =   6135
      Width           =   2070
   End
   Begin VB.Label lblButton 
      BackStyle       =   0  'Transparent
      Caption         =   "F11-Browse"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   10
      Left            =   285
      TabIndex        =   27
      Top             =   8265
      Width           =   1830
   End
   Begin VB.Label lblButton 
      BackStyle       =   0  'Transparent
      Caption         =   "F10-Eload"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   9
      Left            =   285
      TabIndex        =   26
      Top             =   7770
      Width           =   1830
   End
   Begin VB.Label lblButton 
      BackStyle       =   0  'Transparent
      Caption         =   "F9-"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   390
      Index           =   8
      Left            =   285
      TabIndex        =   25
      Top             =   7215
      Width           =   1830
   End
   Begin VB.Label lblButton 
      BackStyle       =   0  'Transparent
      Caption         =   "F8-Delete"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Index           =   7
      Left            =   285
      TabIndex        =   24
      Top             =   6735
      Width           =   1830
   End
   Begin VB.Label lblButton 
      BackStyle       =   0  'Transparent
      Caption         =   "F6-"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Index           =   5
      Left            =   285
      TabIndex        =   22
      Top             =   5715
      Width           =   1830
   End
   Begin VB.Label lblButton 
      BackStyle       =   0  'Transparent
      Caption         =   "F5-Save"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Index           =   4
      Left            =   285
      TabIndex        =   21
      Top             =   5250
      Width           =   1830
   End
   Begin VB.Label lblButton 
      BackStyle       =   0  'Transparent
      Caption         =   "F3-Search"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Index           =   2
      Left            =   285
      TabIndex        =   20
      Top             =   4215
      Width           =   1830
   End
   Begin VB.Label lblButton 
      BackStyle       =   0  'Transparent
      Caption         =   "F2-Disc."
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   330
      Index           =   1
      Left            =   285
      TabIndex        =   19
      Top             =   3705
      Width           =   1830
   End
   Begin VB.Label lblButton 
      BackStyle       =   0  'Transparent
      Caption         =   "F7-"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Index           =   6
      Left            =   285
      TabIndex        =   23
      Top             =   6225
      Width           =   1830
   End
   Begin VB.Shape Shape10 
      BorderColor     =   &H001778E7&
      BorderWidth     =   2
      Height          =   525
      Index           =   0
      Left            =   150
      Top             =   5625
      Width           =   2070
   End
   Begin VB.Shape Shape9 
      BorderColor     =   &H001778E7&
      BorderWidth     =   2
      Height          =   525
      Index           =   0
      Left            =   150
      Top             =   5115
      Width           =   2070
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H001778E7&
      BorderWidth     =   2
      Height          =   525
      Index           =   0
      Left            =   150
      Top             =   4095
      Width           =   2070
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H001778E7&
      BorderWidth     =   2
      Height          =   525
      Index           =   0
      Left            =   150
      Top             =   3585
      Width           =   2070
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H001778E7&
      BorderWidth     =   2
      Height          =   525
      Index           =   0
      Left            =   150
      Top             =   3075
      Width           =   2070
   End
   Begin VB.Label lblButton 
      BackStyle       =   0  'Transparent
      Caption         =   "F1-Help"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Index           =   0
      Left            =   285
      TabIndex        =   18
      Top             =   3180
      Width           =   1830
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   465
      Left            =   11850
      TabIndex        =   32
      Top             =   225
      Width           =   3000
   End
   Begin VB.Label lblDays 
      BackStyle       =   0  'Transparent
      Caption         =   "Wednesday"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   405
      Left            =   11880
      TabIndex        =   31
      Top             =   750
      Width           =   3075
   End
   Begin VB.Label lblTime 
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   525
      Left            =   8805
      TabIndex        =   30
      Top             =   405
      Width           =   2760
   End
   Begin VB.Line Line7 
      BorderColor     =   &H0013B8FD&
      BorderWidth     =   4
      Index           =   0
      X1              =   155
      X2              =   156
      Y1              =   194
      Y2              =   756
   End
   Begin VB.Line Line6 
      BorderColor     =   &H0013B8FD&
      BorderWidth     =   4
      X1              =   10
      X2              =   155
      Y1              =   194
      Y2              =   195
   End
   Begin VB.Line Line5 
      BorderColor     =   &H0013B8FD&
      BorderWidth     =   4
      X1              =   11
      X2              =   10
      Y1              =   11
      Y2              =   195
   End
   Begin VB.Line Line4 
      BorderColor     =   &H0013B8FD&
      BorderWidth     =   4
      X1              =   1014
      X2              =   10
      Y1              =   12
      Y2              =   12
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0013B8FD&
      BorderWidth     =   4
      X1              =   1013
      X2              =   1014
      Y1              =   619
      Y2              =   11
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0013B8FD&
      BorderWidth     =   4
      Index           =   0
      X1              =   156
      X2              =   619
      Y1              =   755
      Y2              =   756
   End
   Begin VB.Label lblField 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "&Trans. No."
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
      Height          =   270
      Index           =   11
      Left            =   375
      TabIndex        =   0
      Top             =   1275
      Width           =   1080
   End
   Begin VB.Label lblField 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "&Invoice No."
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
      Index           =   0
      Left            =   375
      TabIndex        =   2
      Top             =   2010
      Width           =   1185
   End
   Begin VB.Label lblField 
      BackColor       =   &H001778E7&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   555
      Index           =   7
      Left            =   9510
      TabIndex        =   15
      Top             =   10605
      Width           =   2460
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      Height          =   1935
      Left            =   9405
      Top             =   9420
      Width           =   5820
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      Height          =   1890
      Left            =   2490
      Top             =   9285
      Width           =   6660
   End
   Begin VB.Label lblField 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "GUANZON Mobile Shop"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   615
      Index           =   23
      Left            =   1350
      TabIndex        =   29
      Top             =   375
      Width           =   5940
   End
   Begin VB.Image Image1 
      Height          =   720
      Index           =   0
      Left            =   345
      Picture         =   "frmCP_POS.frx":629B
      Stretch         =   -1  'True
      Top             =   345
      Width           =   765
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0013B8FD&
      BorderWidth     =   2
      FillColor       =   &H000080FF&
      Height          =   1620
      Index           =   0
      Left            =   2460
      Top             =   1200
      Width           =   12645
   End
End
Attribute VB_Name = "frmCP_POS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
'Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
'Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
'Private Declare Function ReleaseCapture Lib "user32" () As Long
'
'Private Const pxeMODULENAME = "frmCP_POS"
'Private oFormSerialNo As frmSOSerialNo
'
'Private WithEvents oTrans As clsCPSales
'Private oReceipt As Receipt
''Private oSkin As clsFormSkin
'
'Dim psUserID As String
'Dim pnIndex As Integer
'Dim pnCtr As Integer
'Dim pbHsSerial As Boolean
'
'Private Sub Form_Activate()
'   oApp.MenuName = Me.Tag
'   Me.ZOrder 0
'End Sub
'
'Private Sub Form_Load()
'   Dim lsOldProc As String
'
'   lsOldProc = "Form_Load"
'   On Error GoTo errProc
'
'   Set oFormSerialNo = New frmSOSerialNo
'   Set oTrans = New clsCPSales
'   Set oTrans.AppDriver = oApp
'
'   oTrans.InitTransaction
'   oTrans.NewTransaction
'
'   Set oReceipt = New Receipt
'   Set oReceipt.AppDriver = oApp
''   Set oSkin = New clsFormSkin
''   Set oSkin.AppDriver = oApp
''   Set oSkin.Form = Me
''   oSkin.ApplySkin xeFormTransaction
'
'   psUserID = oApp.UserID
'   InitGrid
'   ClearFields
'
'   lblDate.Caption = Format(oApp.ServerDate, "MMMM DD, YYYY")
'   lblDays.Caption = Format(oApp.ServerDate, "DDDD")
'
''   ShowCursor 0
'   mdiMain.Hide
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " & " )", True
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'   Set oTrans = Nothing
'   Set oReceipt = Nothing
'   Set oFormSerialNo = Nothing
''   Set oSkin = Nothing
''   ShowCursor 1
''   ReleaseCapture
'   mdiMain.Show
'End Sub
'
'Private Sub GridEditor1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
''   SetCapture Me.hwnd
'End Sub
'
'Private Sub GridEditor1_RowColChange()
'   With GridEditor1
'      txtField(5).Text = IIf(.TextMatrix(.Row, 5) = "", "0.00%", .TextMatrix(.Row, 5))
'      txtField(6).Text = IIf(.TextMatrix(.Row, 6) = "", "0.00", .TextMatrix(.Row, 6))
'
'      lblUnitPrice.Caption = Format(.TextMatrix(.Row, 7), "#,##0.00")
'   End With
'End Sub
'
'Private Sub lblTime_Click()
'   lblTime.Caption = Format(oApp.ServerDate, "HH:MM:SS AM/PM")
'End Sub
'
'Private Sub InitGrid()
'   With GridEditor1
'      .Rows = 2
'      .Cols = 8
'      .Font = "MS Sans Serif"
'
'      'column title
'      .TextMatrix(0, 1) = "BarrCode"
'      .TextMatrix(0, 2) = "Description"
'      .TextMatrix(0, 3) = "Qty"
'      .TextMatrix(0, 4) = "Unit Price"
'      .TextMatrix(0, 5) = "Dsc Rte"
'      .TextMatrix(0, 6) = "Dsc Amt"
'      .TextMatrix(0, 7) = "Total"
'
'      .Row = 0
'
'      'column alignment
'      For pnCtr = 0 To .Cols - 1
'         .Col = pnCtr
'         .CellFontBold = True
'         .CellAlignment = 3
'         .CellFontSize = 12
'      Next
'
'      'row height
'      .RowHeight(0) = 400
'
'      'column width
'      .ColWidth(0) = 330
'      .ColWidth(1) = 2030
'      .ColWidth(3) = 800
'      .ColWidth(4) = 1300
'      .ColWidth(5) = 1100
'      .ColWidth(6) = 1100
'      .ColWidth(7) = 1100
'
'      .ColAlignment(1) = 1
'      .ColAlignment(2) = 1
'      .ColAlignment(3) = 6
'      .ColAlignment(4) = 6
'      .ColAlignment(5) = 6
'      .ColAlignment(6) = 6
'      .ColAlignment(7) = 6
'
'      .Row = 1
'      .Col = 1
'      .ColSel = .Cols - 1
'   End With
'End Sub
'
'Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
''   SetCapture Me.hwnd
'End Sub
'
'Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
''   SetCapture Me.hwnd
'End Sub
'
'Private Sub Timer1_Timer()
'   lblTime.Caption = Format(oApp.ServerDate, "HH:MM:SS AM/PM")
'End Sub
'
'Private Sub txtField_GotFocus(Index As Integer)
'   With txtField(Index)
'      Select Case Index
'      Case 2
'         txtField(5).Enabled = False
'         txtField(6).Enabled = False
'      Case 5
'         .Text = Replace(.Text, "%", "")
'      End Select
'
'      .SelStart = 0
'      .SelLength = Len(.Text)
'
''      .BackColor = oApp.getColor("HT1")
'   End With
'
''   pbGridFocus = False
'   pnIndex = Index
'End Sub
'
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'   With GridEditor1
'      Select Case KeyCode
'      Case vbKeyReturn, vbKeyUp, vbKeyDown
'         Select Case KeyCode
'         Case vbKeyReturn, vbKeyDown
'            If GetFocus = .hwnd Then Exit Sub
'            If GetFocus <> txtField(2).hwnd Then SetNextFocus
'         Case vbKeyUp
'            If GetFocus <> txtField(2).hwnd Then SetPreviousFocus
'         End Select
'      Case vbKeyF2
'         If Trim(.TextMatrix(.Row, 1)) <> "" Then
'            txtField(5).Enabled = True
'            txtField(6).Enabled = True
'            txtField(5).SetFocus
'         End If
'      Case vbKeyF3
'         txtField(2).Text = oTrans.SearchReferNo("xReferNox", txtField(2).Text)
'      Case vbKeyF4
'      Case vbKeyF5
'         If Not isEntryOK Then Exit Sub
'         If pbHsSerial Then
'            If Not withSerialNo Then Exit Sub
'         End If
'         If Receipt Then
'            If oTrans.SaveTransaction Then
'               oTrans.NewTransaction
'               ClearFields
'               txtField(2).SetFocus
'            Else
'               MsgBox "Unable to save transaction!!!" & vbCrLf & _
'                        "Please contact GMC/GGC SEG for assistance!!!", vbCritical, "Warning"
'            End If
'         End If
'      Case vbKeyF8
'         If Trim(.TextMatrix(.Row, 1)) <> "" Then
'            If .Rows = 2 Then
'               If oTrans.DeleteDetail(.Row - 1) Then
'                  oTrans.AddDetail
'                  .TextMatrix(1, 0) = "1"
'                  .TextMatrix(1, 1) = oTrans.Detail(0, "xReferNox")
'                  .TextMatrix(1, 2) = oTrans.Detail(0, "sDescript")
'                  .TextMatrix(1, 3) = Format(oTrans.Detail(0, "nQuantity"), "#,##0")
'                  .TextMatrix(1, 4) = Format(oTrans.Detail(0, "nUnitPrce"), "#,##0.00")
'                  .TextMatrix(1, 5) = Format(oTrans.Detail(0, "nDiscRate"), "##0.00") & "%"
'                  .TextMatrix(1, 6) = Format(oTrans.Detail(0, "nDiscAmtx"), "#,##0.00")
'                  .TextMatrix(1, 7) = "0.00"
'               End If
'            Else
'               If oTrans.DeleteDetail(.Row - 1) Then Call DeleteDetail
'            End If
'         End If
'      Case vbKeyF10
'         frmEload.Show 1
'      Case vbKeyEscape
'         Unload Me
'      End Select
'   End With
'End Sub
'
'Private Sub ClearFields()
'   For pnCtr = 0 To txtField.Count - 1
'      Select Case pnCtr
'      Case 0
'         txtField(pnCtr).Text = Format(oTrans.Master("sTransNox"), "@@@@-@@@@@@")
'      Case 1
'         txtField(pnCtr).Text = Format(oTrans.Master("sSalesInv"), ">")
'      Case 3
'         txtField(pnCtr).Text = Format(oApp.getLogName(oApp.UserID), ">")
'      Case 4
'         txtField(pnCtr).Text = Format(oApp.getLogName(psUserID), ">")
'      Case 5
'         txtField(pnCtr).Text = "0.00%"
'      Case 6
'         txtField(pnCtr).Text = "0.00"
'      Case Else
'         txtField(pnCtr).Text = ""
'      End Select
'   Next
'
'   lblTotalAmount.Caption = "0.00"
'   lblUnitPrice.Caption = "0.00"
'
'   With GridEditor1
'      .Rows = 2
'      .ColWidth(2) = 4800
'
'      'empty row
'      .TextMatrix(1, 0) = "1"
'      .TextMatrix(1, 1) = ""
'      .TextMatrix(1, 2) = ""
'      .TextMatrix(1, 3) = "0"
'      .TextMatrix(1, 4) = "0.00"
'      .TextMatrix(1, 5) = "0.00%"
'      .TextMatrix(1, 6) = "0.00"
'      .TextMatrix(1, 7) = "0.00"
'   End With
'   oReceipt.InitReceipt
'   psUserID = oApp.UserID
'   pbHsSerial = False
'
'   oTrans.Master("sCashierx") = psUserID
'   oTrans.Master("sSalesman") = psUserID
'End Sub
'
'Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'   Dim lsValue As String
'   Dim lsBarrCode As String
'   Dim lsQty As String
'   Dim lnCtr As Integer
'   Dim lnQty As Integer
'
'   Select Case KeyCode
'   Case vbKeyReturn
'      With txtField(Index)
'         If Index = 2 Then
'            lsValue = Trim(Left(.Text, 4))
'            lsBarrCode = .Text
'            lnQty = 1
'
'            For lnCtr = 1 To Len(lsValue)
'               If LCase(Left(Right(lsValue, lnCtr), 1)) = "x" Then
'                  lsQty = Left(lsValue, Len(Trim(lsValue)) - lnCtr)
'                  If IsNumeric(lsQty) Then
'                     lnQty = lsQty
'                     If Right(.Text, 1) = "x" Then
'                        lnQty = 1
'                     Else
'                        lsBarrCode = Right(.Text, Len(.Text) - (Len(lsQty) + 1))
'                     End If
'                  Else
'                     lnQty = 1
'                     lsBarrCode = .Text
'                  End If
'               End If
'            Next
'
'            If Trim(.Text) <> "" Then Call InsertDetail(lnQty, lsBarrCode)
'            .Text = ""
'
'            .SetFocus
'         End If
'      End With
'   Case vbKeyDown
'      If GridEditor1.Row = GridEditor1.Rows - 1 Then Exit Sub
'      GridEditor1.Row = GridEditor1.Row + 1
'      GridEditor1.ColSel = GridEditor1.Cols - 1
'   Case vbKeyUp
'      If GridEditor1.Row = 1 Then Exit Sub
'      GridEditor1.Row = GridEditor1.Row - 1
'      GridEditor1.ColSel = GridEditor1.Cols - 1
'   End Select
'End Sub
'
'Private Sub txtField_LostFocus(Index As Integer)
'   With txtField(Index)
'      Select Case Index
'      Case 6
'         txtField(5).Enabled = False
'         txtField(6).Enabled = False
'      End Select
'      .BackColor = oApp.getColor("EB")
'   End With
'End Sub
'
'Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
'   With txtField(Index)
'      Select Case Index
'      Case 5, 6
'         If Not IsNumeric(.Text) Then .Text = 0#
'         If Index = 5 Then
'            If CDbl(.Text) > 99.99 Then .Text = 0#
'            oTrans.Detail(GridEditor1.Row - 1, "nDiscRate") = CDbl(.Text)
'            .Text = Format(.Text, "##0.00") & "%"
'         Else
'            oTrans.Detail(GridEditor1.Row - 1, "nDiscAmtx") = CDbl(.Text)
'            .Text = Format(.Text, "##0.00")
'         End If
'
'         With GridEditor1
'            .TextMatrix(.Row, 5) = Format(oTrans.Detail(.Row - 1, "nDiscRate"), "##0.00") & "%"
'            .TextMatrix(.Row, 6) = Format(oTrans.Detail(.Row - 1, "nDiscAmtx"), "#,##0.00")
'            .TextMatrix(.Row, 7) = Format(CDbl(.TextMatrix(.Row, 3)) * CDbl(.TextMatrix(.Row, 4)) * _
'                                    (100 - CDbl(Replace(.TextMatrix(.Row, 5), "%", ""))) / 100 - CDbl(.TextMatrix(.Row, 6)), "#,##0.00")
'            lblUnitPrice.Caption = .TextMatrix(.Row, 7)
'         End With
'         Call GrandTotal
'      End Select
'   End With
'End Sub
'
'Private Function isEntryOK() As Boolean
'   If CDbl(lblTotalAmount.Caption) <= 0 Then
'      MsgBox "Unable to Pay Transaction!!!" & vbCrLf & _
'                "Total Amount is invalid!!!" & vbCrLf & _
'                "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
'      txtField(2).SetFocus
'      GoTo EntryNotOK
'   End If
'
'   With GridEditor1
'      If Trim(.TextMatrix(1, 1)) = "" Or .TextMatrix(1, 4) = 0 Then
'         MsgBox "Detail is required!!!" & vbCrLf & _
'                "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
'         .SetFocus
'         .Row = 1
'         .Col = 1
'         txtField(2).SetFocus
'         GoTo EntryNotOK
'      End If
'   End With
'
'EntryOK:
'   isEntryOK = True
'   Exit Function
'EntryNotOK:
'   isEntryOK = False
'End Function
'
'Private Sub DeleteDetail()
'   Dim lnCtr As Integer
'
'   With GridEditor1
'      .Rows = oTrans.ItemCount + 1
'      pbHsSerial = False
'      For lnCtr = 1 To .Rows - 1
'         .TextMatrix(lnCtr, 0) = lnCtr
'         .TextMatrix(lnCtr, 1) = oTrans.Detail(lnCtr - 1, "xReferNox")
'         .TextMatrix(lnCtr, 2) = oTrans.Detail(lnCtr - 1, "sDescript")
'         .TextMatrix(lnCtr, 3) = oTrans.Detail(lnCtr - 1, "nQuantity")
'         .TextMatrix(lnCtr, 4) = Format(oTrans.Detail(lnCtr - 1, "nUnitPrce"), "#,##0.00")
'         .TextMatrix(lnCtr, 5) = Format(oTrans.Detail(lnCtr - 1, "nDiscRate"), "##0.00") & "%"
'         .TextMatrix(lnCtr, 6) = Format(oTrans.Detail(lnCtr - 1, "nDiscAmtx"), "#,##0.00")
'         .TextMatrix(lnCtr, 7) = Format(CDbl(.TextMatrix(lnCtr, 3)) * CDbl(.TextMatrix(lnCtr, 4)) * _
'                                       (100 - CDbl(Replace(.TextMatrix(lnCtr, 5), "%", ""))) / 100 - CDbl(.TextMatrix(lnCtr, 6)), "#,##0.00")
'
'         If Not pbHsSerial Then oTrans.Detail(lnCtr - 1, "cHsSerial") = xeYes
'      Next
'      .Row = .Rows - 1
'      .ColSel = .Cols - 1
'
'      .ColWidth(2) = 4800
'      If .Rows > 21 Then .ColWidth(2) = 4550
'
'      lblUnitPrice.Caption = Format(.TextMatrix(.Row, 7), "#,##0.00")
'      Call GrandTotal
'   End With
'End Sub
'
'Private Sub InsertDetail(ByVal Quantity As Integer, ByVal Value As String)
'   With GridEditor1
'      If .Rows = 2 Then
'         If .TextMatrix(.Row, 1) <> "" Then
'            If oTrans.ItemCount <> .Row Then
'               oTrans.AddDetail
'               oTrans.Detail(.Rows - 1, "xReferNox") = Value
'               If oTrans.Detail(.Rows - 1, "xReferNox") <> "" Then
'                  .Rows = .Rows + 1
'                  .Row = .Rows - 1
'                  .TextMatrix(.Row, 1) = Value
'                  .TextMatrix(.Row, 0) = .Row
'               Else
'                  oTrans.DeleteDetail .Row
'                  Exit Sub
'               End If
'            Else
'               oTrans.AddDetail
'               oTrans.Detail(.Row, "xReferNox") = Value
'               If oTrans.Detail(.Row, "xReferNox") <> "" Then
'                  .Rows = .Rows + 1
'                  .Row = .Rows - 1
'                  .TextMatrix(.Row, 1) = Value
'                  .TextMatrix(.Row, 0) = .Row
'               Else
'                  oTrans.DeleteDetail .Row
'                  Exit Sub
'               End If
'            End If
'         Else
'            oTrans.Detail(.Row - 1, "xReferNox") = Value
'            If oTrans.Detail(.Row - 1, "xReferNox") <> "" Then .TextMatrix(.Row, 1) = Value
'            .TextMatrix(.Row, 0) = .Row
'         End If
'      Else
'         If oTrans.ItemCount <> .Row Then
'            oTrans.AddDetail
'            oTrans.Detail(.Rows - 1, "xReferNox") = Value
'            If oTrans.Detail(.Rows - 1, "xReferNox") <> "" Then
'               .Rows = .Rows + 1
'               .Row = .Rows - 1
'               .TextMatrix(.Row, 1) = Value
'               .TextMatrix(.Row, 0) = .Row
'            Else
'               oTrans.DeleteDetail .Rows
'               Exit Sub
'            End If
'         Else
'            oTrans.AddDetail
'            oTrans.Detail(.Row, "xReferNox") = Value
'            If oTrans.Detail(.Row, "xReferNox") <> "" Then
'               .Rows = .Rows + 1
'               .Row = .Rows - 1
'               .TextMatrix(.Row, 1) = Value
'               .TextMatrix(.Row, 0) = .Row
'            Else
'               oTrans.DeleteDetail .Row
'               Exit Sub
'            End If
'         End If
'      End If
'      .ColSel = .Cols - 1
'
'      .ColWidth(2) = 4800
'      If .Rows > 21 Then .ColWidth(2) = 4550
'
'      oTrans.Detail(.Row - 1, "nQuantity") = Quantity
'
'      .TextMatrix(.Row, 2) = oTrans.Detail(.Row - 1, "sDescript")
'      .TextMatrix(.Row, 3) = oTrans.Detail(.Row - 1, "nQuantity")
'      .TextMatrix(.Row, 4) = Format(oTrans.Detail(.Row - 1, "nUnitPrce"), "#,##0.00")
'      .TextMatrix(.Row, 5) = Format(oTrans.Detail(.Row - 1, "nDiscRate"), "##0.00") & "%"
'      .TextMatrix(.Row, 6) = Format(oTrans.Detail(.Row - 1, "nDiscAmtx"), "#,##0.00")
'      .TextMatrix(.Row, 7) = Format(CDbl(.TextMatrix(.Row, 3)) * CDbl(.TextMatrix(.Row, 4)) * _
'                                    (100 - CDbl(Replace(.TextMatrix(.Row, 5), "%", ""))) / 100 - CDbl(.TextMatrix(.Row, 6)), "#,##0.00")
'
'      If Not pbHsSerial Then pbHsSerial = oTrans.Detail(.Row - 1, "cHsSerial") = xeYes
'      lblUnitPrice.Caption = Format(.TextMatrix(.Row, 7), "#,##0.00")
'      Call GrandTotal
'   End With
'End Sub
'
'Private Sub GrandTotal()
'   Dim lnCtr As Integer
'   Dim lnTotal As Currency
'
'   With GridEditor1
'      lnTotal = 0#
'      For lnCtr = 1 To .Rows - 1
'         lnTotal = lnTotal + CDbl(.TextMatrix(lnCtr, 7))
'      Next
'   End With
'   lblTotalAmount.Caption = Format(lnTotal, "#,##0.00")
'   oTrans.Master("nTranTotl") = CDbl(lnTotal)
'End Sub
'
'Private Function Receipt() As Boolean
'   Dim lnCheckAmt As Currency
'   Dim lnCashAmtx As Currency
'   Dim lnCardAmtx As Currency
'   Dim lnTotalAmt As Currency
'   Dim lsOldProc As String
'
'   lsOldProc = "Receipt"
'   On Error GoTo errProc
'
'   With oReceipt
'      lnCheckAmt = oTrans.Receipt("nCheckAmt")
'      lnCardAmtx = oTrans.Receipt("nCardAmtx")
'
'      .CashAmount = oTrans.Master("nCashAmtx")
'      .AmountPaid = oTrans.Master("nTranTotl")
'      .TranDate = oTrans.Master("dTransact")
''      .CustomerName = oTrans.Master("xFullName")
'      .UserID = oTrans.Master("sSalesman")
'      .Remarks = oTrans.Master("sRemarksx")
'
'      'Customer Info
''      .ClientID = oTrans.Customer("sClientID")
''      .LastName = oTrans.Customer("sLastName")
''      .FirstName = oTrans.Customer("sFrstName")
''      .MiddleName = oTrans.Customer("sMiddName")
''      .Address = oTrans.Customer("sAddressx")
''      .TownName = oTrans.Customer("sTownName")
''      .BirthDte = oTrans.Customer("dBirthDte")
''      .PhoneNo = oTrans.Customer("sPhoneNox")
''      .MobileNo = oTrans.Customer("sMobileNo")
''      .EmailAdd = oTrans.Customer("sEmailAdd")
''      .GenderCode = oTrans.Customer("cGenderCd")
''      .CivilStatus = oTrans.Customer("cCivlStat")
''      .TownID = oTrans.Customer("sTownIDxx")
'      .InvoiceNo = oTrans.Master("sSalesInv")
'
'      If lnCheckAmt > 0 Then
'         .Checks("sBankIDxx") = oTrans.Checks("sBankIDxx")
'         .Checks("sCheckNox") = oTrans.Checks("sCheckNox")
'         .Checks("dCheckDte") = oTrans.Checks("dTransact")
'         .Checks("nCheckAmt") = oTrans.Checks("nCheckAmt")
'      Else
'         .Checks("sBankIDxx") = ""
'         .Checks("sCheckNox") = ""
'         .Checks("dCheckDte") = oTrans.Master("dTransact")
'         .Checks("nCheckAmt") = 0#
'      End If
'
'      If lnCardAmtx > 0 Then
'         .Cards("sBankIDxx") = oTrans.Cards("sBankIDxx")
'         .Cards("sCardIDxx") = oTrans.Cards("sCardIDxx")
'         .Cards("sCardNoxx") = oTrans.Cards("sCardNoxx")
'         .Cards("sApproval") = oTrans.Cards("sApproval")
'         .Cards("nCardAmtx") = oTrans.Cards("nCardAmtx")
'      Else
'         .Cards("sBankIDxx") = ""
'         .Cards("sCardIDxx") = ""
'         .Cards("sCardNoxx") = ""
'         .Cards("sApproval") = ""
'         .Cards("nCardAmtx") = 0#
'      End If
'      .HasSerial = pbHsSerial
'      .ShowReceipt
'
'      If Not .Cancelled Then
'         txtField(4).Text = Format(oApp.getLogName(.UserID), ">")
'
'         If .CheckAmount > 0 Then
'            oTrans.Checks("sBankIDxx") = .Checks("sBankIDxx")
'            oTrans.Checks("sCheckNox") = .Checks("sCheckNox")
'            oTrans.Checks("dCheckDte") = .Checks("dCheckDte")
'            oTrans.Checks("nCheckAmt") = .Checks("nCheckAmt")
'         End If
'
'         If .CardAmount > 0 Then
'            oTrans.Cards("sBankIDxx") = .Cards("sBankIDxx")
'            oTrans.Cards("sCardIDxx") = .Cards("sCardIDxx")
'            oTrans.Cards("sCardNoxx") = .Cards("sCardNoxx")
'            oTrans.Cards("sApproval") = .Cards("sApproval")
'            oTrans.Cards("nCardAmtx") = .Cards("nCardAmtx")
'         End If
'
'         oTrans.Master("sSalesman") = .UserID
''         oTrans.Master("sClientID") = .ClientID
'         oTrans.Master("sRemarksx") = .Remarks
'         oTrans.Master("nCashAmtx") = .CashAmount
'         oTrans.Master("nAmtPaidx") = .AmountPaid
'
'         'Customer Info
''         oTrans.Customer("sClientID") = .ClientID
''         oTrans.Customer("sLastName") = .LastName
''         oTrans.Customer("sFrstName") = .FirstName
''         oTrans.Customer("sMiddName") = .MiddleName
''         oTrans.Customer("sAddressx") = .Address
''         oTrans.Customer("sTownName") = .TownName
''         oTrans.Customer("dBirthDte") = .BirthDte
''         oTrans.Customer("sPhoneNox") = .PhoneNo
''         oTrans.Customer("sMobileNo") = .MobileNo
''         oTrans.Customer("sEmailAdd") = .EmailAdd
''         oTrans.Customer("cGenderCd") = .GenderCode
''         oTrans.Customer("cCivlStat") = .CivilStatus
''         oTrans.Customer("sTownIDxx") = .TownID
'         oTrans.Master("sSalesInv") = .InvoiceNo
'
'         txtField(1).Text = oTrans.Master("sSalesInv")
'         Receipt = True
'      End If
'   End With
'
'endProc:
'   Exit Function
'errProc:
'   ShowError lsOldProc & "( " & " )", True
'End Function
'
'Private Function withSerialNo() As Boolean
'   Dim lsOldProc As String
'
'   lsOldProc = "withSerialNo"
'   On Error GoTo errProc
'
'   With oFormSerialNo
'      Set .SerialTrans = oTrans
'      .InitGrid1
'      .Show 1
'
'      If .Cancelled Then Exit Function
'   End With
'
'   withSerialNo = True
'
'endProc:
'   Exit Function
'errProc:
'   ShowError lsOldProc & "( " & " )", True
'End Function
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
