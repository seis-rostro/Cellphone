VERSION 5.00
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0#1.0#0"; "xrGridControl.ocx"
Begin VB.Form Form1 
   BackColor       =   &H000080FF&
   BorderStyle     =   0  'None
   Caption         =   " "
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   768
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text14 
      Height          =   450
      Left            =   3120
      TabIndex        =   42
      Text            =   "Text14"
      Top             =   7920
      Width           =   5070
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   90
      Top             =   45
   End
   Begin VB.TextBox Text13 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   11985
      TabIndex        =   27
      Text            =   "Text13"
      Top             =   2505
      Width           =   2865
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   1665
      TabIndex        =   24
      Text            =   "Text2"
      Top             =   2550
      Width           =   4200
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   5970
      TabIndex        =   23
      Text            =   "Text3"
      Top             =   2550
      Width           =   4200
   End
   Begin xrGridEditor.GridEditor GridEditor2 
      Height          =   3495
      Left            =   2700
      TabIndex        =   22
      Top             =   3855
      Width           =   12330
      _ExtentX        =   21749
      _ExtentY        =   6165
      AllowBigSelection=   -1  'True
      AutoAdd         =   0   'False
      AutoNumber      =   0   'False
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
      Object.HEIGHT          =   3495
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
      MOUSEICON       =   "Form1.frx":629B
      MOUSEPOINTER    =   0
      REDRAW          =   -1  'True
      RIGHTTOLEFT     =   0   'False
      ROWS            =   14
      SCROLLBARS      =   3
      SCROLLTRACK     =   0   'False
      SELECTIONMODE   =   0
      Object.TOOLTIPTEXT     =   ""
      WORDWRAP        =   0   'False
   End
   Begin VB.TextBox Text7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """Php""#,##0.000;(""Php""#,##0.000)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   13321
         SubFormatType   =   2
      EndProperty
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   10875
      TabIndex        =   18
      Text            =   "100,000.75"
      Top             =   7695
      Width           =   4005
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   10875
      TabIndex        =   17
      Text            =   "Text8"
      Top             =   9015
      Width           =   4005
   End
   Begin VB.TextBox Text9 
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
      ForeColor       =   &H000000FF&
      Height          =   750
      Left            =   10875
      TabIndex        =   16
      Text            =   "Text9"
      Top             =   10215
      Width           =   4005
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   3120
      TabIndex        =   12
      Text            =   "Text4"
      Top             =   8805
      Width           =   5070
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   3120
      TabIndex        =   11
      Text            =   "Text5"
      Top             =   9645
      Width           =   5070
   End
   Begin VB.TextBox Text6 
      Height          =   405
      Left            =   3120
      TabIndex        =   10
      Text            =   "Text6"
      Top             =   10410
      Width           =   5070
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   540
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   2550
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   540
      Left            =   495
      ScaleHeight     =   480
      ScaleWidth      =   14325
      TabIndex        =   0
      Top             =   1305
      Width           =   14385
      Begin VB.TextBox Text12 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   5790
         TabIndex        =   9
         Text            =   "Text12"
         Top             =   0
         Width           =   8490
      End
      Begin VB.TextBox Text11 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   945
         TabIndex        =   8
         Text            =   "Text11"
         Top             =   0
         Width           =   4770
      End
      Begin VB.TextBox Text10 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   15
         TabIndex        =   7
         Text            =   "Text10"
         Top             =   0
         Width           =   870
      End
      Begin VB.Line Line1 
         X1              =   870
         X2              =   870
         Y1              =   15
         Y2              =   390
      End
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "F11-EXIT"
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
      Left            =   1155
      TabIndex        =   44
      Top             =   9360
      Width           =   1230
   End
   Begin VB.Image Image12 
      Height          =   435
      Left            =   240
      Picture         =   "Form1.frx":62B7
      Stretch         =   -1  'True
      Top             =   9300
      Width           =   450
   End
   Begin VB.Shape Shape11 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Height          =   525
      Left            =   150
      Top             =   9270
      Width           =   2250
   End
   Begin VB.Label Label25 
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
      Left            =   2805
      TabIndex        =   43
      Top             =   7620
      Width           =   1155
   End
   Begin VB.Shape Shape10 
      BorderColor     =   &H001778E7&
      BorderWidth     =   2
      Height          =   525
      Index           =   1
      Left            =   150
      Top             =   8760
      Width           =   2250
   End
   Begin VB.Shape Shape9 
      BorderColor     =   &H001778E7&
      BorderWidth     =   2
      Height          =   525
      Index           =   1
      Left            =   150
      Top             =   8250
      Width           =   2250
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H001778E7&
      BorderWidth     =   2
      Height          =   525
      Index           =   1
      Left            =   150
      Top             =   7740
      Width           =   2250
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H001778E7&
      BorderWidth     =   2
      Height          =   525
      Index           =   1
      Left            =   150
      Top             =   7230
      Width           =   2250
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H001778E7&
      BorderWidth     =   2
      Height          =   525
      Index           =   1
      Left            =   150
      Top             =   6720
      Width           =   2250
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "F10-J.O."
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
      Left            =   1140
      TabIndex        =   41
      Top             =   8850
      Width           =   1050
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "F9-Repl"
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
      Left            =   1140
      TabIndex        =   40
      Top             =   8355
      Width           =   1050
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "F8-Regr."
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
      Left            =   1140
      TabIndex        =   39
      Top             =   7800
      Width           =   1050
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "F7-Inst."
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
      Left            =   1140
      TabIndex        =   38
      Top             =   7320
      Width           =   1050
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "F5-Cheq"
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
      Left            =   1140
      TabIndex        =   37
      Top             =   6300
      Width           =   1050
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "F4-Void"
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
      Left            =   1140
      TabIndex        =   36
      Top             =   5835
      Width           =   1050
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "F3-Find"
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
      Left            =   1140
      TabIndex        =   35
      Top             =   5310
      Width           =   1050
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "F2-Save"
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
      Left            =   1140
      TabIndex        =   34
      Top             =   4800
      Width           =   1050
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "F6-Card"
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
      Left            =   1140
      TabIndex        =   33
      Top             =   6810
      Width           =   1050
   End
   Begin VB.Shape Shape10 
      BorderColor     =   &H001778E7&
      BorderWidth     =   2
      Height          =   525
      Index           =   0
      Left            =   150
      Top             =   6210
      Width           =   2250
   End
   Begin VB.Shape Shape9 
      BorderColor     =   &H001778E7&
      BorderWidth     =   2
      Height          =   525
      Index           =   0
      Left            =   150
      Top             =   5700
      Width           =   2250
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H001778E7&
      BorderWidth     =   2
      Height          =   525
      Index           =   0
      Left            =   150
      Top             =   5190
      Width           =   2250
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H001778E7&
      BorderWidth     =   2
      Height          =   525
      Index           =   0
      Left            =   150
      Top             =   4680
      Width           =   2250
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H001778E7&
      BorderWidth     =   2
      Height          =   525
      Index           =   0
      Left            =   150
      Top             =   4170
      Width           =   2250
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "F1-Disc"
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
      Left            =   1140
      TabIndex        =   32
      Top             =   4275
      Width           =   1050
   End
   Begin VB.Image Image11 
      Height          =   465
      Left            =   255
      Picture         =   "Form1.frx":7F81
      Stretch         =   -1  'True
      Top             =   8790
      Width           =   480
   End
   Begin VB.Image Image10 
      Height          =   465
      Left            =   225
      Picture         =   "Form1.frx":86EB
      Stretch         =   -1  'True
      Top             =   8280
      Width           =   480
   End
   Begin VB.Image Image9 
      Height          =   465
      Left            =   240
      Picture         =   "Form1.frx":8E55
      Stretch         =   -1  'True
      Top             =   7770
      Width           =   480
   End
   Begin VB.Image Image8 
      Height          =   465
      Left            =   240
      Picture         =   "Form1.frx":95BF
      Stretch         =   -1  'True
      Top             =   7245
      Width           =   480
   End
   Begin VB.Image Image7 
      Height          =   465
      Left            =   240
      Picture         =   "Form1.frx":9D29
      Stretch         =   -1  'True
      Top             =   6735
      Width           =   480
   End
   Begin VB.Image Image6 
      Height          =   465
      Left            =   195
      Picture         =   "Form1.frx":A493
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   480
   End
   Begin VB.Image Image5 
      Height          =   465
      Left            =   225
      Picture         =   "Form1.frx":ABFD
      Stretch         =   -1  'True
      Top             =   5715
      Width           =   480
   End
   Begin VB.Image Image4 
      Height          =   465
      Left            =   225
      Picture         =   "Form1.frx":B367
      Stretch         =   -1  'True
      Top             =   5250
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   465
      Left            =   225
      Picture         =   "Form1.frx":BAD1
      Stretch         =   -1  'True
      Top             =   4695
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   465
      Left            =   210
      Picture         =   "Form1.frx":C23B
      Stretch         =   -1  'True
      Top             =   4170
      Width           =   480
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
      TabIndex        =   31
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
      Index           =   0
      Left            =   11880
      TabIndex        =   30
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
      TabIndex        =   29
      Top             =   405
      Width           =   2760
   End
   Begin VB.Line Line7 
      BorderColor     =   &H0013B8FD&
      BorderWidth     =   4
      X1              =   170
      X2              =   170
      Y1              =   260
      Y2              =   752
   End
   Begin VB.Line Line6 
      BorderColor     =   &H0013B8FD&
      BorderWidth     =   4
      X1              =   11
      X2              =   169
      Y1              =   260
      Y2              =   259
   End
   Begin VB.Line Line5 
      BorderColor     =   &H0013B8FD&
      BorderWidth     =   4
      X1              =   12
      X2              =   11
      Y1              =   12
      Y2              =   259
   End
   Begin VB.Line Line4 
      BorderColor     =   &H0013B8FD&
      BorderWidth     =   4
      X1              =   1008
      X2              =   11
      Y1              =   12
      Y2              =   12
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0013B8FD&
      BorderWidth     =   4
      X1              =   1008
      X2              =   1008
      Y1              =   752
      Y2              =   11
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0013B8FD&
      BorderWidth     =   4
      X1              =   170
      X2              =   1007
      Y1              =   752
      Y2              =   752
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Total:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   12705
      TabIndex        =   28
      Top             =   3165
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction #:"
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
      Height          =   390
      Left            =   3015
      TabIndex        =   26
      Top             =   3150
      Width           =   1560
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice #:"
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
      Left            =   7740
      TabIndex        =   25
      Top             =   3150
      Width           =   1080
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H001778E7&
      BorderWidth     =   2
      Height          =   1260
      Left            =   11685
      Top             =   2415
      Width           =   3330
   End
   Begin VB.Label Label11 
      BackColor       =   &H001778E7&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount:"
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
      Left            =   9405
      TabIndex        =   21
      Top             =   7905
      Width           =   1485
   End
   Begin VB.Label Label12 
      BackColor       =   &H001778E7&
      BackStyle       =   0  'Transparent
      Caption         =   "Amount Rendered:"
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
      Height          =   510
      Left            =   9720
      TabIndex        =   20
      Top             =   9135
      Width           =   1200
   End
   Begin VB.Label Label13 
      BackColor       =   &H001778E7&
      BackStyle       =   0  'Transparent
      Caption         =   "Change Due:"
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
      Height          =   420
      Left            =   9495
      TabIndex        =   19
      Top             =   10485
      Width           =   1380
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      Height          =   3570
      Left            =   9270
      Top             =   7545
      Width           =   5745
   End
   Begin VB.Label Label8 
      BackColor       =   &H001778E7&
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name:"
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
      Height          =   330
      Left            =   2775
      TabIndex        =   15
      Top             =   8475
      Width           =   1770
   End
   Begin VB.Label Label9 
      BackColor       =   &H001778E7&
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Person"
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
      Left            =   2790
      TabIndex        =   14
      Top             =   9330
      Width           =   1485
   End
   Begin VB.Label Label10 
      BackColor       =   &H001778E7&
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks:"
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
      Height          =   285
      Left            =   2820
      TabIndex        =   13
      Top             =   10125
      Width           =   1020
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      Height          =   3570
      Left            =   2685
      Top             =   7545
      Width           =   5640
   End
   Begin VB.Label Label5 
      BackColor       =   &H002B337D&
      BackStyle       =   0  'Transparent
      Caption         =   "Discounts:"
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
      Left            =   450
      TabIndex        =   5
      Top             =   3150
      Width           =   1080
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000080FF&
      Height          =   1260
      Left            =   330
      Top             =   2415
      Width           =   10020
   End
   Begin VB.Label Label16 
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
      Left            =   1350
      TabIndex        =   6
      Top             =   375
      Width           =   5940
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   345
      Picture         =   "Form1.frx":C9A5
      Stretch         =   -1  'True
      Top             =   345
      Width           =   765
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
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
      Height          =   300
      Left            =   10005
      TabIndex        =   3
      Top             =   1905
      Width           =   1275
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Bar-Code"
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
      Left            =   3300
      TabIndex        =   2
      Top             =   1920
      Width           =   1050
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity"
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
      Height          =   285
      Left            =   570
      TabIndex        =   1
      Top             =   1920
      Width           =   945
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0013B8FD&
      BorderWidth     =   2
      FillColor       =   &H000080FF&
      Height          =   1110
      Left            =   345
      Top             =   1200
      Width           =   14670
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub xrButton5_Click()
End Sub

Private Sub Form_Load()
   lblDate = Format(Date, "MMMM DD, YYYY")
   lblDay = Format(Date, "DDDD")

End Sub

Private Sub Timer1_Timer()
   lblTime.Caption = Format(Time, "HH:MM:SS AM/PM")
End Sub
