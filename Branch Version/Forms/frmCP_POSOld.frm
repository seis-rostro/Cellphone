VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmCP_POSOld 
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
   Picture         =   "frmCP_POSOld.frx":0000
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
      Height          =   420
      Index           =   7
      Left            =   3690
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   10515
      Width           =   5295
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
      TabIndex        =   16
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
      TabIndex        =   14
      Text            =   "00.00%"
      Top             =   9570
      Width           =   990
   End
   Begin VB.PictureBox Picture2 
      Height          =   2115
      Left            =   135
      ScaleHeight     =   2055
      ScaleWidth      =   2025
      TabIndex        =   25
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
         TabIndex        =   26
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
      Left            =   3690
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   9375
      Width           =   5295
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
      Left            =   3690
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   9840
      Width           =   5295
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Height          =   915
      Left            =   2550
      ScaleHeight     =   915
      ScaleWidth      =   12450
      TabIndex        =   19
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
         TabIndex        =   9
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
      Height          =   6270
      Left            =   2460
      TabIndex        =   6
      TabStop         =   0   'False
      Tag             =   "et0;eb0;et0;bc2"
      Top             =   2895
      Width           =   12645
      _ExtentX        =   22304
      _ExtentY        =   11060
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
      BackColor       =   &H001778E7&
      BackStyle       =   0  'Transparent
      Caption         =   "C.Name:"
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
      Left            =   2580
      TabIndex        =   39
      Top             =   10605
      Width           =   1110
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
      Caption         =   "F7-Print"
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
      Left            =   270
      TabIndex        =   37
      Top             =   6225
      Width           =   1800
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
      Left            =   270
      TabIndex        =   36
      Top             =   3705
      Width           =   1800
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
      Left            =   270
      TabIndex        =   35
      Top             =   4215
      Width           =   1800
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
      Left            =   270
      TabIndex        =   34
      Top             =   5250
      Width           =   1800
   End
   Begin VB.Label lblButton 
      BackStyle       =   0  'Transparent
      Caption         =   "F6-Update"
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
      Left            =   270
      TabIndex        =   33
      Top             =   5715
      Width           =   1800
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
      Left            =   270
      TabIndex        =   32
      Top             =   6735
      Width           =   1800
   End
   Begin VB.Label lblButton 
      BackStyle       =   0  'Transparent
      Caption         =   "F9-Void"
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
      Left            =   270
      TabIndex        =   31
      Top             =   7215
      Width           =   1800
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
      Left            =   270
      TabIndex        =   30
      Top             =   7770
      Width           =   1800
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
      Left            =   270
      TabIndex        =   29
      Top             =   8265
      Width           =   1800
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
      Left            =   270
      TabIndex        =   28
      Top             =   8775
      Width           =   1800
   End
   Begin VB.Label lblButton 
      BackStyle       =   0  'Transparent
      Caption         =   "F4-New"
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
      Left            =   270
      TabIndex        =   27
      Top             =   4725
      Width           =   1800
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
      TabIndex        =   18
      Top             =   10515
      Width           =   3000
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
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
      Left            =   12165
      TabIndex        =   15
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
      TabIndex        =   13
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
      Caption         =   "IMEI No - &Barcode"
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
   Begin VB.Shape Shape11 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Height          =   525
      Left            =   150
      Top             =   8685
      Width           =   2070
   End
   Begin VB.Label lblField 
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
      Left            =   2580
      TabIndex        =   8
      Top             =   9405
      Width           =   1155
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
      Left            =   270
      TabIndex        =   20
      Top             =   3180
      Width           =   1800
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
      Left            =   11010
      TabIndex        =   24
      Top             =   225
      Width           =   4005
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
      Left            =   11040
      TabIndex        =   23
      Top             =   750
      Width           =   4080
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
      Left            =   8205
      TabIndex        =   22
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
      TabIndex        =   17
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
   Begin VB.Label lblField 
      BackColor       =   &H001778E7&
      BackStyle       =   0  'Transparent
      Caption         =   "Salesman:"
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
      Left            =   2565
      TabIndex        =   11
      Top             =   9900
      Width           =   1110
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      Height          =   1875
      Left            =   2490
      Top             =   9255
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
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   615
      Index           =   23
      Left            =   1350
      TabIndex        =   21
      Top             =   375
      Width           =   5940
   End
   Begin VB.Image Image1 
      Height          =   720
      Index           =   0
      Left            =   345
      Picture         =   "frmCP_POSOld.frx":629B
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
Attribute VB_Name = "frmCP_POSOld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShowCursor Lib "USER32" (ByVal bShow As Long) As Long
Private Declare Function SetCapture Lib "USER32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "USER32" () As Long

Private Const pxeMODULENAME = "frmCP_POS"
Private oFormSerialNewNo As frmSOSerialNo

Private WithEvents oTrans As clsCPSales
Attribute oTrans.VB_VarHelpID = -1
Private oReceipt As ggcCPSales.Receipt
'Private oSkin As clsFormSkin

Dim pnIndex As Integer
Dim pnCtr As Integer
Dim pbHsSerial As Boolean
Dim psUserName As String
Dim psUserIDxx As String
Dim pnTtlAdj As Currency

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String

   lsOldProc = "Form_Load"
   ''On Error GoTo errProc

   Set oFormSerialNewNo = New frmSOSerialNo
   Set oTrans = New clsCPSales
   Set oTrans.AppDriver = oApp

   oTrans.InitTransaction
   oTrans.NewTransaction

   Set oReceipt = New ggcCPSales.Receipt
   Set oReceipt.AppDriver = oApp

'   Set oSkin = New clsFormSkin
'   Set oSkin.AppDriver = oApp
'   Set oSkin.Form = Me
'   oSkin.ApplySkin xeFormTransaction

   psUserIDxx = ""
   psUserName = ""
   InitGrid
   initButton xeModeAddNew
   clearFields

   lblDate.Caption = Format(oApp.ServerDate, "MMMM DD, YYYY")
   lblDays.Caption = Format(oApp.ServerDate, "DDDD")

'   ShowCursor 0
   mdiMain.Hide

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oTrans = Nothing
   Set oReceipt = Nothing
   Set oFormSerialNewNo = Nothing
'   Set oSkin = Nothing
'   ShowCursor 1
'   ReleaseCapture
   mdiMain.Show
End Sub

Private Sub GridEditor1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   SetCapture Me.hwnd
End Sub

Private Sub GridEditor1_RowColChange()
   With GridEditor1
      txtField(5).Text = IIf(.TextMatrix(.Row, 5) = "", "0.00%", .TextMatrix(.Row, 5))
      txtField(6).Text = IIf(.TextMatrix(.Row, 6) = "", "0.00", .TextMatrix(.Row, 6))

      lblUnitPrice.Caption = Format(.TextMatrix(.Row, 7), "#,##0.00")
   End With
End Sub

Private Sub lblTime_Click()
   lblTime.Caption = Format(oApp.ServerDate, "HH:MM:SS AM/PM")
End Sub

Private Sub initButton(lnStat As Integer)
   Dim lbShow As Boolean

   lbShow = IIf(lnStat = 0, False, True)
   If lbShow Then
      lblButton(0).Caption = "F1-Help"
      lblButton(1).Caption = "F2-Disc."
      lblButton(2).Caption = "F3-Search"
      lblButton(3).Caption = "F4-"
      lblButton(4).Caption = "F5-Save"
      lblButton(5).Caption = "F6-Price"
      lblButton(6).Caption = "F7-"
      lblButton(7).Caption = "F8-Delete"
      lblButton(8).Caption = "F9-Wallet"
      lblButton(9).Caption = "F10-Eload"
   Else
      lblButton(0).Caption = "F1-Help"
      lblButton(1).Caption = "F2-"
      lblButton(2).Caption = "F3-"
      lblButton(3).Caption = "F4-New"
      lblButton(4).Caption = "F5-"
      lblButton(5).Caption = "F6-"
      lblButton(6).Caption = "F7-Access"
      lblButton(7).Caption = "F8-GAways"
      lblButton(8).Caption = "F9-"
      lblButton(9).Caption = "F10-"
   End If
End Sub

Private Sub InitGrid()
   With GridEditor1
      .Rows = 2
      .Cols = 8
      .Font = "MS Sans Serif"

      'column title
      .TextMatrix(0, 1) = "BarrCode"
      .TextMatrix(0, 2) = "Description"
      .TextMatrix(0, 3) = "Qty"
      .TextMatrix(0, 4) = "Unit Price"
      .TextMatrix(0, 5) = "Dsc Rte"
      .TextMatrix(0, 6) = "Dsc Amt"
      .TextMatrix(0, 7) = "Total"

      .Row = 0

      'column alignment
      For pnCtr = 0 To .Cols - 1
         .Col = pnCtr
         .CellFontBold = True
         .CellAlignment = 3
         .CellFontSize = 12
      Next

      'row height
      .RowHeight(0) = 400

      'column width
      .ColWidth(0) = 330
      .ColWidth(1) = 2030
      .ColWidth(3) = 800
      .ColWidth(4) = 1300
      .ColWidth(5) = 1100
      .ColWidth(6) = 1100
      .ColWidth(7) = 1100

      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 6
      .ColAlignment(4) = 6
      .ColAlignment(5) = 6
      .ColAlignment(6) = 6
      .ColAlignment(7) = 6

      .Row = 1
      .Col = 1
      .ColSel = .Cols - 1
   End With
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   SetCapture Me.hwnd
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   SetCapture Me.hwnd
End Sub

Private Sub Timer1_Timer()
   lblTime.Caption = Format(oApp.ServerDate, "HH:MM:SS AM/PM")
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      Select Case Index
      Case 2
         txtField(5).Enabled = False
         txtField(6).Enabled = False
      Case 5
         .Text = Replace(.Text, "%", "")
      End Select

      .SelStart = 0
      .SelLength = Len(.Text)

'      .BackColor = oApp.getColor("HT1")
   End With

'   pbGridFocus = False
   pnIndex = Index
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim lsUserID As String
   Dim lsUserName As String
   Dim lnUserRights As Integer
   Dim lasRights() As String
   Dim lnRep As Long

   With GridEditor1
      Select Case KeyCode
      Case vbKeyReturn, vbKeyUp, vbKeyDown
         Select Case KeyCode
         Case vbKeyReturn, vbKeyDown
            If GetFocus = .hwnd Then Exit Sub
            If GetFocus <> txtField(2).hwnd Then SetNextFocus
         Case vbKeyUp
            If GetFocus <> txtField(2).hwnd Then SetPreviousFocus
         End Select
      Case vbKeyF2
         If Trim(.TextMatrix(.Row, 1)) <> "" Then
            txtField(5).Enabled = True
            txtField(6).Enabled = True
            txtField(5).SetFocus
         End If
      Case vbKeyF3
         If oTrans.EditMode = xeModeAddNew Then
            txtField(2).Text = oTrans.SearchReferNo("xReferNox", txtField(2).Text)
         End If
      Case vbKeyF4
         If oTrans.EditMode <> xeModeAddNew Then
            oTrans.NewTransaction
            initButton xeModeAddNew
            clearFields

            txtField(2).SetFocus
         End If
      Case vbKeyF5
         If oTrans.EditMode = xeModeAddNew Then
            If Not isEntryOk Then Exit Sub
            If pbHsSerial Then
               If Not withSerialNo Then Exit Sub
            End If
            If Receipt Then
               If oTrans.SaveTransaction Then
                  lnRep = MsgBox("Do you want to print transaciton?", vbQuestion + vbYesNo, "Confirm")
                  
                  If lnRep = vbYes Then
                     MsgBox "Please mount the SI.", vbInformation, "Print"
                  
                     If Not PrintTrans Then MsgBox "Unable to Print Transaction!!!", vbCritical, "Warning"
                  End If
               
                  initButton xeModeReady
               Else
                  MsgBox "Unable to save transaction!!!" & vbCrLf & _
                           "Please contact GMC/GGC SEG for assistance!!!", vbCritical, "Warning"
               End If
            End If
         End If
      Case vbKeyF6
      'she 2022-10-25
      'temprary disable for mobile fiesta
      If oApp.BranchCode <> "C0M2" Then
         MsgBox "Updating of price was disallowed." & vbCrLf & vbCrLf & _
                  "Please input the price difference as discount rate or amount.", vbInformation, "Notice"
         Exit Sub
      End If
      
         If oTrans.EditMode = xeModeAddNew Then
            If oTrans.Detail(GridEditor1.Row - 1, "sStockIDx") = "" Then Exit Sub
            With frmSOParts
               .StockID = oTrans.Detail(GridEditor1.Row - 1, "sStockIDx")
               .Show 1

               If Not .Cancelled Then
                  With GridEditor1
                     oTrans.Detail(.Row - 1, "nUnitPrce") = frmSOParts.UnitPrice
                     .TextMatrix(.Row, 4) = Format(oTrans.Detail(.Row - 1, "nUnitPrce"), "#,##0.00")
                     .TextMatrix(.Row, 7) = Format(CDbl(.TextMatrix(.Row, 3)) * CDbl(.TextMatrix(.Row, 4)) * _
                                       (100 - CDbl(Replace(.TextMatrix(.Row, 5), "%", ""))) / 100 - CDbl(.TextMatrix(.Row, 6)), "#,##0.00")
                     Call GrandTotal
                  End With
               End If
            End With
         End If

'         Select Case oTrans.EditMode
'         Case xeModeAddNew
'            ClearFields
'         Case xeModeReady
'            oTrans.UpdateTransaction
'            InitButton xeModeUpdate
'         Case xeModeUpdate
'            If oTrans.OpenTransaction(oTrans.Master("sTransNox")) Then
'               LoadMaster
'               LoadDetail
'               Call oTrans.UpdateTransaction
'            End If
'         End Select
      Case vbKeyF7
      Case vbKeyF8
         If oTrans.EditMode = xeModeAddNew Then
            If Trim(.TextMatrix(.Row, 1)) <> "" Then
               If .Rows = 2 Then
                  If oTrans.deleteDetail(.Row - 1) Then
                     oTrans.addDetail
                     .TextMatrix(1, 0) = "1"
                     .TextMatrix(1, 1) = oTrans.Detail(0, "xReferNox")
                     .TextMatrix(1, 2) = oTrans.Detail(0, "sDescript")
                     .TextMatrix(1, 3) = Format(oTrans.Detail(0, "nQuantity"), "#,##0")
                     .TextMatrix(1, 4) = Format(oTrans.Detail(0, "nUnitPrce"), "#,##0.00")
                     .TextMatrix(1, 5) = Format(oTrans.Detail(0, "nDiscRate"), "##0.00") & "%"
                     .TextMatrix(1, 6) = Format(oTrans.Detail(0, "nDiscAmtx"), "#,##0.00")
                     .TextMatrix(1, 7) = "0.00"
                  End If
               Else
                  If oTrans.deleteDetail(.Row - 1) Then Call deleteDetail
               End If
               lblUnitPrice.Caption = .TextMatrix(.Row, 7)
               Call GrandTotal
            End If
         End If
      Case vbKeyF9
         Select Case oTrans.EditMode
         Case xeModeReady
            lasRights = Split(oApp.mdiMain.Controls(oApp.MenuName).Tag, "»")
            If GetApproval(oApp, lnUserRights, lsUserID, lsUserName, oApp.MenuName) = False Then Exit Sub

            If (lnUserRights And (xeSupervisor + xeSysAdmin)) = 0 Then
               MsgBox "Approving Officer Has No Right to Void this transaction!!!" & vbCrLf & _
                  "Request can not be granted!!!", vbCritical, "Warning"
               Exit Sub
            End If

            If oTrans.CancelTransaction Then
               MsgBox "Transaction Cancelled Successfully!!!", vbInformation, "Notice"
            Else
               MsgBox "Unable to Cancel Transaction!!!", vbCritical, "Warning"
            End If
            clearFields
         Case xeModeAddNew
            clearFields
            frmLoadWallet.Show 1
         Case xeModeUpdate
            If oTrans.OpenTransaction(oTrans.Master("sTransNox")) Then
               LoadMaster
               LoadDetail
            End If
            initButton xeModeReady
         End Select
      Case vbKeyF10
         clearFields
         frmEload.Show 1
      Case vbKeyF11
         If oTrans.SearchTransaction() Then
            LoadMaster
            LoadDetail
            initButton xeModeReady
         End If
         GridEditor1.Refresh
      Case vbKeyEscape
         Unload Me
      End Select
   End With
End Sub

Private Sub clearFields()
   For pnCtr = 0 To txtField.Count - 1
      Select Case pnCtr
      Case 0
         txtField(pnCtr).Text = Format(oTrans.Master("sTransNox"), "@@@@-@@@@@@")
      Case 1
         txtField(pnCtr).Text = Format(oTrans.Master("sSalesInv"), ">")
      Case 3
         txtField(pnCtr).Text = Format(oApp.getUserName(oApp.UserID), ">")
      Case 4
         txtField(pnCtr).Text = psUserName
      Case 5
         txtField(pnCtr).Text = "0.00%"
      Case 6
         txtField(pnCtr).Text = "0.00"
      Case Else
         txtField(pnCtr).Text = ""
      End Select
   Next

   lblTotalAmount.Caption = "0.00"
   lblUnitPrice.Caption = "0.00"

   With GridEditor1
      .Rows = 2
      .ColWidth(2) = 4800

      'empty row
      .TextMatrix(1, 0) = "1"
      .TextMatrix(1, 1) = ""
      .TextMatrix(1, 2) = ""
      .TextMatrix(1, 3) = "0"
      .TextMatrix(1, 4) = "0.00"
      .TextMatrix(1, 5) = "0.00%"
      .TextMatrix(1, 6) = "0.00"
      .TextMatrix(1, 7) = "0.00"
   End With
   oReceipt.InitReceipt
   pbHsSerial = False

   oTrans.Master("sCashierx") = oApp.UserID
   oTrans.Master("sSalesman") = psUserIDxx
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lsValue As String
   Dim lsBarrCode As String
   Dim lsQty As String
   Dim lnCtr As Integer
   Dim lnQty As Integer
   Dim lbDuplicate As Boolean

   Select Case KeyCode
   Case vbKeyReturn
      With txtField(Index)
         If Index = 2 Then
            lsValue = Trim(Left(.Text, 4))
            lsBarrCode = .Text
            lnQty = 1

'           2019-07-01 10:54 AM
'           jeff
'           temporary disable this option for equinox entry
'           instead update the qty manually
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

            With GridEditor1
               For lnCtr = 1 To .Rows - 1
                  If Trim(LCase(lsBarrCode)) = Trim(LCase(.TextMatrix(lnCtr, 1))) Then
                     .TextMatrix(lnCtr, 3) = CDbl(.TextMatrix(lnCtr, 3)) + lnQty
                     .TextMatrix(.Row, 7) = Format(CDbl(.TextMatrix(.Row, 3)) * CDbl(.TextMatrix(.Row, 4)) * _
                                    (100 - CDbl(Replace(.TextMatrix(.Row, 5), "%", ""))) / 100 - CDbl(.TextMatrix(.Row, 6)), "#,##0.00")
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
         End If
      End With
   Case vbKeyDown
      If GridEditor1.Row = GridEditor1.Rows - 1 Then Exit Sub
      GridEditor1.Row = GridEditor1.Row + 1
      GridEditor1.ColSel = GridEditor1.Cols - 1
   Case vbKeyUp
      If GridEditor1.Row = 1 Then Exit Sub
      GridEditor1.Row = GridEditor1.Row - 1
      GridEditor1.ColSel = GridEditor1.Cols - 1
   End Select
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      Select Case Index
      Case 6
         txtField(5).Enabled = False
         txtField(6).Enabled = False
      End Select
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   With txtField(Index)
      Select Case Index
      Case 5, 6
         If Not IsNumeric(.Text) Then .Text = 0#
         If Index = 5 Then
            If CDbl(.Text) > 99.99 Then .Text = 0#
            oTrans.Detail(GridEditor1.Row - 1, "nDiscRate") = CDbl(.Text)
            .Text = Format(.Text, "##0.00") & "%"
         Else
            oTrans.Detail(GridEditor1.Row - 1, "nDiscAmtx") = CDbl(.Text)
            .Text = Format(.Text, "##0.00")
         End If

         With GridEditor1
            .TextMatrix(.Row, 5) = Format(oTrans.Detail(.Row - 1, "nDiscRate"), "##0.00") & "%"
            .TextMatrix(.Row, 6) = Format(oTrans.Detail(.Row - 1, "nDiscAmtx"), "#,##0.00")
            .TextMatrix(.Row, 7) = Format(CDbl(.TextMatrix(.Row, 3)) * CDbl(.TextMatrix(.Row, 4)) * _
                                    (100 - CDbl(Replace(.TextMatrix(.Row, 5), "%", ""))) / 100 - CDbl(.TextMatrix(.Row, 6)), "#,##0.00")
            lblUnitPrice.Caption = .TextMatrix(.Row, 7)
         End With
         Call GrandTotal
      End Select
   End With
End Sub

Private Function isEntryOk() As Boolean
   If CDbl(lblTotalAmount.Caption) <= 0 Then
      MsgBox "Unable to Pay Transaction!!!" & vbCrLf & _
                "Total Amount is invalid!!!" & vbCrLf & _
                "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
      txtField(2).SetFocus
      GoTo EntryNotOK
   End If

   With GridEditor1
      If Trim(.TextMatrix(1, 1)) = "" Or .TextMatrix(1, 4) = 0 Then
         MsgBox "Detail is required!!!" & vbCrLf & _
                "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
         .SetFocus
         .Row = 1
         .Col = 1
         txtField(2).SetFocus
         GoTo EntryNotOK
      End If
   End With

EntryOK:
   isEntryOk = True
   Exit Function
EntryNotOK:
   isEntryOk = False
End Function

Private Sub deleteDetail()
   Dim lnCtr As Integer

   With GridEditor1
      .Rows = oTrans.ItemCount + 1
      pbHsSerial = False
      For lnCtr = 1 To .Rows - 1
         .TextMatrix(lnCtr, 0) = lnCtr
         .TextMatrix(lnCtr, 1) = oTrans.Detail(lnCtr - 1, "xReferNox")
         .TextMatrix(lnCtr, 2) = oTrans.Detail(lnCtr - 1, "sDescript")
         .TextMatrix(lnCtr, 3) = oTrans.Detail(lnCtr - 1, "nQuantity")
         .TextMatrix(lnCtr, 4) = Format(oTrans.Detail(lnCtr - 1, "nUnitPrce"), "#,##0.00")
         .TextMatrix(lnCtr, 5) = Format(oTrans.Detail(lnCtr - 1, "nDiscRate"), "##0.00") & "%"
         .TextMatrix(lnCtr, 6) = Format(oTrans.Detail(lnCtr - 1, "nDiscAmtx"), "#,##0.00")
         .TextMatrix(lnCtr, 7) = Format(CDbl(.TextMatrix(lnCtr, 3)) * CDbl(.TextMatrix(lnCtr, 4)) * _
                                       (100 - CDbl(Replace(.TextMatrix(lnCtr, 5), "%", ""))) / 100 - CDbl(.TextMatrix(lnCtr, 6)), "#,##0.00")

         If Not pbHsSerial Then oTrans.Detail(lnCtr - 1, "cHsSerial") = xeYes
      Next
      .Row = .Rows - 1
      .ColSel = .Cols - 1

      .ColWidth(2) = 4800
      If .Rows > 21 Then .ColWidth(2) = 4550

      lblUnitPrice.Caption = Format(.TextMatrix(.Row, 7), "#,##0.00")
      Call GrandTotal
   End With
End Sub

Private Sub InsertDetail(ByVal Quantity As Integer, ByVal Value As String)
   With GridEditor1
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

      .ColWidth(2) = 4800
      If .Rows > 21 Then .ColWidth(2) = 4550

      oTrans.Detail(.Row - 1, "nQuantity") = Quantity
      .TextMatrix(.Row, 2) = oTrans.Detail(.Row - 1, "sDescript")
      .TextMatrix(.Row, 3) = oTrans.Detail(.Row - 1, "nQuantity")
      .TextMatrix(.Row, 4) = Format(oTrans.Detail(.Row - 1, "nUnitPrce"), "#,##0.00")
      .TextMatrix(.Row, 5) = Format(oTrans.Detail(.Row - 1, "nDiscRate"), "##0.00") & "%"
      .TextMatrix(.Row, 6) = Format(oTrans.Detail(.Row - 1, "nDiscAmtx"), "#,##0.00")
      .TextMatrix(.Row, 7) = Format(CDbl(.TextMatrix(.Row, 3)) * CDbl(.TextMatrix(.Row, 4)) * _
                                    (100 - CDbl(Replace(.TextMatrix(.Row, 5), "%", ""))) / 100 - CDbl(.TextMatrix(.Row, 6)), "#,##0.00")
      
      If Not pbHsSerial Then pbHsSerial = oTrans.Detail(.Row - 1, "cHsSerial") = xeYes
      lblUnitPrice.Caption = Format(.TextMatrix(.Row, 7), "#,##0.00")
      Call GrandTotal
   End With
End Sub

Private Sub GrandTotal()
   Dim lnCtr As Integer
   Dim lnTotal As Currency

   With GridEditor1
      lnTotal = 0#
      For lnCtr = 1 To .Rows - 1
         lnTotal = lnTotal + CDbl(.TextMatrix(lnCtr, 7))
      Next
   End With
'   lblTotalAmount.Caption = Format(oTrans.Master("nTranTotl"), "#,##0.00")
   lblTotalAmount.Caption = Format(lnTotal, "#,##0.00")
   'she 2016-03-30 what if payment is from credit card with card rate?
   oTrans.Master("nTranTotl") = CDbl(lnTotal)
End Sub

Private Function Receipt() As Boolean
   Dim lnCheckAmt As Currency
   Dim lnCashAmtx As Currency
   Dim lnCardAmtx As Currency
   Dim lnTotalAmt As Currency
   Dim lsOldProc As String

   lsOldProc = "Receipt"
   ''On Error GoTo errProc

   With oReceipt
      lnCheckAmt = oTrans.Receipt("nCheckAmt")
      lnCardAmtx = oTrans.Receipt("nCardAmtx")
      .Client = oTrans.Client
      .Sales = oTrans
      .UserName = psUserName
      .UserID = psUserIDxx
      .CashAmount = oTrans.Master("nCashAmtx")
      .AmountPaid = Format(lblTotalAmount.Caption, "#,##0.00")
      .InvoiceDate = oTrans.Master("dTransact")
      .UserID = oTrans.Master("sSalesman")
      .Remarks = oTrans.Master("sRemarksx")
      .InvoiceNo = oTrans.Master("sSalesInv")
      .ORNo = oTrans.Master("sORNoxxxx")

      If lnCheckAmt > 0 Then
         .Checks("sBankIDxx") = oTrans.Checks("sBankIDxx")
         .Checks("sCheckNox") = oTrans.Checks("sCheckNox")
         .Checks("dCheckDte") = oTrans.Checks("dTransact")
         .Checks("nAmountxx") = oTrans.Checks("nAmountxx")
         .Checks("sAcctNoxx") = oTrans.Checks("sAcctNoxx")
      Else
         .Checks("sBankIDxx") = ""
         .Checks("sCheckNox") = ""
         .Checks("dCheckDte") = oTrans.Master("dTransact")
         .Checks("nAmountxx") = 0#
         .Checks("sAcctNoxx") = ""
      End If

      If lnCardAmtx > 0 Then
         .Cards("sBankIDxx") = oTrans.Card(0, "sBankIDxx")
         .Cards("sCrCardID") = oTrans.Card(0, "sCrCardID")
         .Cards("sCrCardNo") = oTrans.Card(0, "sCrCardNo")
         .Cards("sApprovNo") = oTrans.Card(0, "sApprovNo")
         .Cards("nCardAmtx") = oTrans.Card(0, "nTranTotl")
      Else
         .Cards("sBankIDxx") = ""
         .Cards("sCardIDxx") = ""
         .Cards("sCardNoxx") = ""
         .Cards("sApproval") = ""
         .Cards("nCardAmtx") = 0#
      End If
      .HasSerial = pbHsSerial
      .ShowReceipt

      If Not .Cancelled Then
         txtField(4).Text = .UserName
         oTrans.Receipt("nCheckAmt") = 0#
'         oTrans.Receipt("nCardAmtx") = 0#

         If .CheckAmount > 0 Then
            oTrans.Checks("sBankIDxx") = .Checks("sBankIDxx")
            oTrans.Checks("sCheckNox") = .Checks("sCheckNox")
            oTrans.Checks("dCheckDte") = .Checks("dCheckDte")
            oTrans.Checks("nAmountxx") = .Checks("nAmountxx")
            oTrans.Checks("sAcctNoxx") = .Checks("sAcctNoxx")
            
            oTrans.Receipt("nCheckAmt") = oTrans.Checks("nAmountxx")
         End If
         
         oTrans.Receipt("nTranTotl") = oTrans.Master("nTranTotl")
         oTrans.Receipt("nCashAmtx") = oTrans.Master("nCashAmtx")
         oTrans.Receipt("sRemarksx") = oTrans.Master("sRemarksx")
         oTrans.Master("sSalesman") = .UserID
         oTrans.Master("sRemarksx") = .Remarks
         oTrans.Master("nCashAmtx") = .CashAmount
         oTrans.Master("nAmtPaidx") = oTrans.Master("nAmtPaidX")
         oTrans.Master("sSalesInv") = .InvoiceNo
         oTrans.Master("dTransact") = .InvoiceDate
         oTrans.Master("sORNoxxxx") = .ORNo

         txtField(1).Text = oTrans.Master("sSalesInv")
         txtField(7).Text = oReceipt.Client.Master("sCompnyNm")

         oTrans.Client = oReceipt.Client
         oTrans.Master("sClientID") = oReceipt.Client.Master("sClientID")
         oTrans.Master("xFullName") = oReceipt.Client.Master("sCompnyNm")
         oTrans.Master("xAddressx") = oReceipt.Client.Master("sAddressx") & ", " & oReceipt.Client.Master("sTownName") & ", " & oReceipt.Client.Master("sProvName") & " " & oReceipt.Client.Master("sZippCode")

         psUserName = .UserName
         psUserIDxx = .UserID
         Receipt = True
      End If
   End With

endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )", True
End Function

Function PrintTrans() As Boolean
   'C00109000695
   Dim lrs As New ADODB.Recordset
   Dim lnCtr As Integer
   Dim lsIMEI As String
   Dim lReplAmt As String
   Dim lsOldProc As String
   Dim loModel As Recordset
   
   Dim lsSQL As String
   Dim lrsCOInv As Recordset
   Dim lnFinAmt As Currency
   Dim lsFinTrans As String
   Dim lnTotalDisc As Currency 'she temporary for mobile fest
   Dim lnSelPrice As Double
   
   lsOldProc = "printTrans"
   ''On Error GoTo errProc
   
   PrintTrans = False
   
   Set lrs = New ADODB.Recordset
   
   lrs.Fields.Append "nField01", adInteger, 3
   lrs.Fields.Append "sField01", adVarChar, 50
   lrs.Fields.Append "sField02", adVarChar, 60
   lrs.Fields.Append "sField03", adVarChar, 30
   lrs.Fields.Append "lField01", adCurrency
   lrs.Fields.Append "lField02", adCurrency
   lrs.Fields.Append "lField03", adCurrency
   lrs.Fields.Append "lField04", adCurrency
   lrs.Fields.Append "lField05", adCurrency
   lrs.Fields.Append "lField06", adCurrency
   lrs.Open
      
   lsSQL = "SELECT a.sTransNox" & _
            ", b.sBarrCode" & _
            ", c.sModelNme" & _
            ", c.sModelCde" & _
            ", d.sColorNme" & _
            ", a.nUnitPrce" & _
            ", a.nDiscRate" & _
            ", a.nQuantity" & _
            ", b.sDescript" & _
            ", b.cHsSerial" & _
            ", a.sSerialID" & _
            ", e.sSerialNo" & _
            ", f.nReplAmtx" & _
            ", f.nAmtPaidX" & _
            ", g.sTransNox `sFinTrans`" & _
            ", h.sCompnyNm" & _
            ", g.nFinAmtxx" & _
            ", a.nDiscAmtx"
            
    lsSQL = lsSQL & _
         " FROM CP_SO_Detail a" & _
               " LEFT JOIN CP_Inventory_Serial e" & _
                  " ON a.sSerialID = e.sSerialID" & _
            ", CP_Inventory b" & _
               " LEFT JOIN Color d" & _
                  " ON b.sColorIDx = d.sColorIDx" & _
            ", CP_Model c" & _
            ", CP_SO_Master f" & _
               " LEFT JOIN CP_SO_Finance g" & _
                  " ON f.sTransNox = g.sTransNox" & _
               " LEFT JOIN Client_Master h" & _
                  " ON g.sClientID = h.sClientID" & _
         " WHERE a.sTransNox = " & strParm(oTrans.Master("sTransNox")) & _
            " AND a.sTransNox = f.sTransNox" & _
            " AND a.sStockIDx = b.sStockIDx" & _
            " AND b.sModelIDx = c.sModelIDx" & _
         " ORDER BY a.nEntryNox"
   
   Set loModel = New Recordset
   loModel.Open lsSQL, oApp.Connection, , , adCmdText
   
   lsIMEI = "Unit IMEI: "
   lReplAmt = "PR Amt: "
   lnFinAmt = 0#
   lsFinTrans = ""
   lnTotalDisc = 0#
'   If IFNull(loModel("sFinTrans"), "") <> "" Then
'      lnFinAmt = loModel("nFinAmtxx")
'   End If
   
   With loModel
      .MoveFirst
      
      Do Until .EOF
         lrs.AddNew
         lrs("nField01").Value = loModel("nQuantity")
         
         'she 2015 - 4 - 10
         'print barrcode if <> serialize
         If loModel("cHsSerial") = xeYes Then
            lrs("sField01").Value = IFNull(loModel("sModelCde"), loModel("sModelNme"))
            lrs("sField02").Value = loModel("sBarrCode") & " " & IFNull(loModel("sColorNme")) & ";"
         Else
            lrs("sField01").Value = loModel("sBarrCode")
            lrs("sField02").Value = loModel("sDescript")
         End If
         
         If IFNull(loModel("nReplAmtx"), 0#) <> 0# Then
            lrs("sField03").Value = lReplAmt
            lrs("lField03").Value = loModel("nReplAmtx")
         Else
            lrs("lField03").Value = 0#
         End If
         
         If IFNull(loModel("sFinTrans"), "") <> "" Then
            If lsFinTrans = loModel("sFinTrans") Then
               lnFinAmt = lnFinAmt - loModel("nFinAmtxx")
            Else
               lnFinAmt = loModel("nFinAmtxx")
            End If
         End If
         
         'mac 2022-03-24
         lnSelPrice = loModel("nUnitPrce")
         If CDbl(loModel("nDiscRate")) > 0# Then
            lnSelPrice = loModel("nQuantity") * CDbl(loModel("nUnitPrce")) + CDbl(loModel("nDiscAmtx"))
            lnSelPrice = (lnSelPrice * 100) / (100 - CDbl(loModel("nDiscRate")))
         End If
         'mac 2022-03-24

         If IFNull(loModel("sFinTrans"), "") <> "" Then
            lrs("lField01").Value = CDbl(lnSelPrice) - (loModel("nQuantity") * CDbl(lnSelPrice * CDbl(loModel("nDiscRate") / 100))) - CDbl(loModel("nDiscAmtx"))
'            lrs("lField02").Value = CDbl(lrs("lField01").Value) * CDbl(loModel("nQuantity"))
            lrs("lField02").Value = CDbl(lrs("lField01").Value) * CDbl(loModel("nQuantity"))
            
            'she 2022-03-23 to get the total discount amount
            lnTotalDisc = lnTotalDisc + (loModel("nQuantity") * CDbl(lnSelPrice * CDbl(loModel("nDiscRate") / 100))) + CDbl(loModel("nDiscAmtx"))
'            lrs("lField05").Value = (loModel("nQuantity") * CDbl(loModel("nUnitPrce") * CDbl(loModel("nDiscRate") / 100))) + CDbl(loModel("nDiscAmtx"))
'            lrs("lField06").Value = CDbl(lrs("lField02").Value - lrs("lField05").Value)
            
            lnFinAmt = lnFinAmt - lrs("lField05")
            'lnFinAmt = lnFinAmt - (CDbl(loModel("nUnitPrce")) * (100 - (CDbl(loModel("nDiscRate")))) / 100)
         Else
            lrs("lField01").Value = CDbl(lnSelPrice) - (loModel("nQuantity") * CDbl(lnSelPrice * CDbl(loModel("nDiscRate") / 100))) - CDbl(loModel("nDiscAmtx"))
'            lrs("lField02").Value = CDbl(lrs("lField01").Value) * loModel("nQuantity")
            lrs("lField02").Value = CDbl(lrs("lField01").Value) * CDbl(loModel("nQuantity"))
            
            'she 2022-03-23 to get the total discount amount
            lnTotalDisc = lnTotalDisc + (loModel("nQuantity") * CDbl(lnSelPrice * CDbl(loModel("nDiscRate") / 100))) + CDbl(loModel("nDiscAmtx"))

'            lrs("lField05").Value = (loModel("nQuantity") * CDbl(loModel("nUnitPrce") * CDbl(loModel("nDiscRate") / 100))) + CDbl(loModel("nDiscAmtx"))
'            lrs("lField06").Value = CDbl(lrs("lField02").Value - lrs("lField05").Value)
         End If
         
         If loModel("cHsSerial") = xeYes Then
            lsIMEI = lsIMEI & ";" & IFNull(loModel("sSerialNo"))
         End If
         
         lsFinTrans = IFNull(loModel("sFinTrans"))

         .MoveNext
      Loop
   End With
   
   pnTtlAdj = 0#
   
   lsSQL = "SELECT" & _
            " b.sTransNox" & _
            " FROM CP_CO_Master a" & _
            ", AR_Payment_Detail b" & _
            " WHERE a.sTransNox = b.sReferNox" & _
            " AND a.sReferNox = " & strParm(oTrans.Master("sTransNox")) & _
            " AND a.cTranStat = '4' "
          
   Debug.Print lsSQL
   Set lrsCOInv = New Recordset
   lrsCOInv.Open lsSQL, oApp.Connection, , , adCmdText
   
   If Not lrsCOInv.EOF Then
      Call ComputeAdjustment(lrsCOInv("sTransNox"))
      lrs("lField04").Value = Format(pnTtlAdj, "#,##0.00")
   Else
      lrs("lField04").Value = 0#
   End If

   ' assign important info to the report
   Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\CP_SI.rpt")
   oReport.DiscardSavedData
   oReport.FieldMappingType = crAutoFieldMapping
   oReport.Database.SetDataSource lrs

   loModel.Requery
   With oReceipt
      oReport.Sections("PH").ReportObjects("txtCustomer").SetText oTrans.Master("xFullName")
      oReport.Sections("PH").ReportObjects("txtDate").SetText Format(oTrans.Master("dTransact"), "MMM-DD-YYYY")
      oReport.Sections("PH").ReportObjects("txtAddress").SetText oTrans.Master("xAddressx")
      oReport.Sections("PH").ReportObjects("txtTIN").SetText ""
      oReport.Sections("PH").ReportObjects("txtBusiness").SetText ""
      oReport.Sections("PH").ReportObjects("txtTerm").SetText oTrans.Master("sTermName")
      oReport.Sections("PH").ReportObjects("txtPrepared").SetText Trim(txtField(4))
      oReport.Sections("RF").ReportObjects("txtIMEI").SetText lsIMEI
'      oReport.Sections("RF").ReportObjects("txtDiscount").SetText "Total Discount:" & Format(lnTotalDisc, "#,##0.00")
      oReport.Sections("RF").ReportObjects("txtAccessories").SetText getAccesories
'      oReport.Sections("RF").ReportObjects("txtGiveaways").SetText getGiveAways
      If IFNull(loModel("sFinTrans"), "") <> "" Then
         oReport.Sections("RF").ReportObjects("txtRemarks").SetText loModel("sCompnyNm") & " " & Format(loModel("nFinAmtxx"), "#,##0.00")
      Else
         oReport.Sections("RF").ReportObjects("txtRemarks").SetText oTrans.Master("sRemarksx")
      End If
   End With
   
   oReport.PrintOutEx False, 1
   lrs.Close
   PrintTrans = True

endProc:
   oTrans.CloseTransaction oTrans.Master(0)
   Set oReport = Nothing
   Set lrs = Nothing
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"
End Function

Private Function getGiveAways() As String
   Dim lrs As Recordset
   Dim lsSQL As String
   
   lsSQL = "SELECT" & _
               "  a.sStockIDx" & _
               ", b.sDescript" & _
               ", a.nQuantity" & _
               ", a.nGivenxxx" & _
            " FROM CP_SO_GiveAways a" & _
               " LEFT JOIN CP_Inventory b" & _
                  " ON a.sStockIDx = b.sStockIDx" & _
            " WHERE a.cGAwyStat <> '2'" & _
               " AND a.sTransNox = " & strParm(oTrans.Master("sTransNox")) & _
            " ORDER BY a.nGivenxxx DESC"
   
   Set lrs = New Recordset
   With lrs
      .Open lsSQL, oApp.Connection, adOpenStatic, adLockOptimistic, adCmdText
      Set .ActiveConnection = Nothing
      
      lsSQL = ""
      
      If .EOF Then GoTo endProc
      
      lsSQL = "Giveaways: "
      .MoveFirst
      Do Until .EOF
         If lrs("nGivenxxx") = 1 Then
            lsSQL = lsSQL & lrs("sDescript") & "; "
         End If
         .MoveNext
      Loop
      
      lsSQL = "To Follow: "
      .MoveFirst
      Do Until .EOF
         If lrs("nGivenxxx") = 0 Then
            lsSQL = lsSQL & lrs("sDescript") & "; "
         End If
         .MoveNext
      Loop
   End With
endProc:
   getGiveAways = lsSQL
End Function

Private Function getAccesories() As String
   Dim lrs As Recordset
   Dim lsSQL As String
   
   
   lsSQL = "SELECT" & _
               "  b.sDescript" & _
               ", a.sSerialNo" & _
            " FROM CP_SO_Accessories a" & _
               " LEFT JOIN CP_Accessories b" & _
                  " ON a.sAccessID = b.sAccessID" & _
            " WHERE a.sTransNox = " & strParm(oTrans.Master("sTransNox"))
   
   Set lrs = New Recordset
   lrs.Open lsSQL, oApp.Connection, adOpenStatic, adLockOptimistic, adCmdText
   Set lrs.ActiveConnection = Nothing
   
   lsSQL = ""
   
   If lrs.EOF Then GoTo endProc
   
   lsSQL = "Accessories: "
   Do Until lrs.EOF
      lsSQL = lsSQL & lrs("sDescript") & " - " + lrs("sSerialNo") & "; "
               
      lrs.MoveNext
   Loop
   
endProc:
   getAccesories = lsSQL
End Function

Private Function withSerialNo() As Boolean
   Dim lsOldProc As String

   lsOldProc = "withSerialNo"
   ''On Error GoTo errProc

   With oFormSerialNewNo
      Set .SerialTrans = oTrans
      If oTrans.EditMode <> xeModeAddNew Then .GridEditor1.ColEnabled(1) = False
      .InitGrid1
      .Show 1

      If .Cancelled Then Exit Function
   End With

   withSerialNo = True

endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )", True
End Function

Private Sub LoadMaster()
   For pnCtr = 0 To txtField.Count - 1
      Select Case pnCtr
      Case 0
         txtField(pnCtr).Text = Format(oTrans.Master(pnCtr), "@@@@-@@@@@@")
      Case 1
         txtField(pnCtr).Text = Format(oTrans.Master("sSalesInv"), ">")
      Case 2
         txtField(pnCtr).Text = ""
         txtField(pnCtr).Tag = txtField(pnCtr).Text
      Case 3
         txtField(pnCtr).Tag = Format(oApp.getUserName(oTrans.Master("sCashierx")), ">")
      Case 4
         txtField(pnCtr).Tag = Format(oApp.getLogName(oTrans.Master("sSalesman")), ">")
      Case 5
         txtField(pnCtr).Tag = "0.00 %"
      Case 6
         txtField(pnCtr).Text = "0.00"
      Case Else
         txtField(pnCtr).Text = IIf(IsNull(oTrans.Master(pnCtr)), "", oTrans.Master(pnCtr))
      End Select
   Next

   lblTotalAmount.Caption = Format(oTrans.Master("nTranTotl") - oTrans.Master("nReplAmtx"), "#,##0.00")
End Sub

Private Sub LoadDetail()
   Dim lnCtr As Integer

   With GridEditor1
      .Rows = IIf(oTrans.ItemCount = 0, 2, oTrans.ItemCount + 1)

'      .ColWidth(3) = 3100
'      If .Rows > 16 Then .ColWidth(3) = 2900

      For pnCtr = 0 To oTrans.ItemCount - 1
         For lnCtr = 1 To .Cols - 1
            .TextMatrix(pnCtr + 1, lnCtr) = oTrans.Detail(pnCtr, lnCtr)
         Next
      Next
   End With
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

Private Sub ComputeAdjustment(lsTransNox As String)
   Dim lsSQL As String
   Dim lorec As Recordset
   
   lsSQL = "SELECT" & _
            " SUM(nCredtAmt) `nCredtAmt`" & _
            " FROM AR_Payment_Detail" & _
            " WHERE sTransNox = " & strParm(lsTransNox) & _
            " AND sSourceCD = 'CPCm' " & _
            " GROUP BY sTransNox"
   Set lorec = New Recordset
   lorec.Open lsSQL, oApp.Connection, , , adCmdText
   
   pnTtlAdj = lorec("nCredtAmt")
End Sub

