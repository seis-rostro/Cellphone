VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0"; "xrGridControl.ocx"
Begin VB.Form frmCP_Charge_Invoice_Reg 
   BorderStyle     =   0  'None
   Caption         =   "Charge Invoice"
   ClientHeight    =   8325
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8325
   ScaleWidth      =   11760
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame2 
      Height          =   3585
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   4620
      Width           =   10080
      _ExtentX        =   17780
      _ExtentY        =   6324
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
         Left            =   6735
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   2985
         Width           =   3240
      End
      Begin xrGridEditor.GridEditor GridEditor1 
         Height          =   2895
         Left            =   60
         TabIndex        =   42
         Top             =   75
         Width           =   9945
         _ExtentX        =   17542
         _ExtentY        =   5106
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
         Object.HEIGHT          =   2895
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
         MOUSEICON       =   "frmCP_Charge_Invoice_Reg.frx":0000
         MOUSEPOINTER    =   0
         REDRAW          =   -1  'True
         RIGHTTOLEFT     =   0   'False
         ROWS            =   2
         SCROLLBARS      =   3
         SCROLLTRACK     =   0   'False
         SELECTIONMODE   =   0
         Object.TOOLTIPTEXT     =   ""
         WORDWRAP        =   0   'False
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&AMOUNT PAID(F12)"
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
         Left            =   4590
         TabIndex        =   40
         Top             =   3105
         Width           =   2100
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   3480
      Index           =   0
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   1110
      Width           =   10080
      _ExtentX        =   17780
      _ExtentY        =   6138
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   7770
         TabIndex        =   27
         Top             =   1215
         Width           =   1995
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   7770
         TabIndex        =   29
         Top             =   1530
         Width           =   1995
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   9
         Left            =   7770
         TabIndex        =   33
         Top             =   2160
         Width           =   1995
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   12
         Left            =   7770
         TabIndex        =   39
         Top             =   3105
         Width           =   1995
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   10
         Left            =   7770
         TabIndex        =   35
         Top             =   2475
         Width           =   1995
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   11
         Left            =   7770
         TabIndex        =   37
         Top             =   2790
         Width           =   1995
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   7770
         TabIndex        =   31
         Top             =   1845
         Width           =   1995
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   1365
         MaxLength       =   50
         TabIndex        =   14
         Top             =   435
         Width           =   2310
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   915
         Index           =   7
         Left            =   1365
         MultiLine       =   -1  'True
         TabIndex        =   23
         Top             =   2265
         Width           =   4950
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   1365
         TabIndex        =   17
         Top             =   1005
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
         TabIndex        =   12
         Top             =   60
         Width           =   2310
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   600
         Index           =   4
         Left            =   1365
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1320
         Width           =   4950
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
         TabIndex        =   15
         Tag             =   "wt0;fb0"
         Top             =   780
         Width           =   2415
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   1365
         TabIndex        =   21
         Top             =   1950
         Width           =   4950
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "unknown"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   240
         Left            =   7275
         TabIndex        =   43
         Tag             =   "eb0;et0"
         Top             =   135
         Width           =   2385
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H80000007&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   285
         Index           =   0
         Left            =   7230
         Tag             =   "et0;et0"
         Top             =   120
         Width           =   2490
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice No."
         Height          =   195
         Index           =   2
         Left            =   6510
         TabIndex        =   26
         Top             =   1260
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Term"
         Height          =   195
         Index           =   14
         Left            =   6510
         TabIndex        =   28
         Top             =   1575
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Discount Rate"
         Height          =   195
         Index           =   0
         Left            =   6510
         TabIndex        =   32
         Top             =   2205
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Balance"
         Height          =   195
         Index           =   4
         Left            =   6510
         TabIndex        =   38
         Top             =   3150
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Discount Amt"
         Height          =   195
         Index           =   6
         Left            =   6510
         TabIndex        =   34
         Top             =   2520
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Limit"
         Height          =   195
         Index           =   7
         Left            =   6510
         TabIndex        =   36
         Top             =   2835
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Due Date"
         Height          =   195
         Index           =   10
         Left            =   6510
         TabIndex        =   30
         Top             =   1890
         Width           =   690
      End
      Begin VB.Shape Shape2 
         Height          =   330
         Index           =   1
         Left            =   7200
         Tag             =   "et0;et0"
         Top             =   90
         Width           =   2535
      End
      Begin VB.Shape Shape3 
         Height          =   390
         Index           =   0
         Left            =   7170
         Top             =   60
         Width           =   2595
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Name"
         Height          =   195
         Index           =   3
         Left            =   600
         TabIndex        =   16
         Top             =   1050
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transact. Date"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   13
         Top             =   480
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   195
         Index           =   12
         Left            =   435
         TabIndex        =   22
         Top             =   2625
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trans. No."
         Height          =   195
         Index           =   9
         Left            =   150
         TabIndex        =   11
         Top             =   105
         Width           =   750
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   285
         Left            =   1425
         Tag             =   "et0;ht2"
         Top             =   135
         Width           =   2325
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   195
         Index           =   11
         Left            =   495
         TabIndex        =   18
         Top             =   1365
         Width           =   570
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*PIC"
         Height          =   195
         Index           =   5
         Left            =   750
         TabIndex        =   20
         Top             =   1995
         Width           =   315
      End
      Begin VB.Label lblTrantotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Caption         =   "999,000.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   435
         Left            =   7170
         TabIndex        =   25
         Top             =   720
         Width           =   2595
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
         Left            =   7170
         TabIndex        =   24
         Top             =   465
         Width           =   2070
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   10410
      TabIndex        =   5
      Top             =   3045
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
      Picture         =   "frmCP_Charge_Invoice_Reg.frx":001C
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   10410
      TabIndex        =   1
      Top             =   525
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
      Picture         =   "frmCP_Charge_Invoice_Reg.frx":0796
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   10410
      TabIndex        =   3
      Top             =   1785
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Void"
      AccessKey       =   "V"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCP_Charge_Invoice_Reg.frx":0F10
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   10410
      TabIndex        =   4
      Top             =   2415
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
      Picture         =   "frmCP_Charge_Invoice_Reg.frx":168A
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   525
      Index           =   1
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   10080
      _ExtentX        =   17780
      _ExtentY        =   926
      Begin VB.TextBox txtField 
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
         Index           =   15
         Left            =   4560
         MaxLength       =   50
         TabIndex        =   9
         Top             =   105
         Width           =   5220
      End
      Begin VB.TextBox txtField 
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
         Index           =   14
         Left            =   1275
         TabIndex        =   8
         Top             =   105
         Width           =   1950
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Custo&mer"
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
         Left            =   3600
         TabIndex        =   10
         Top             =   135
         Width           =   840
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice No"
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
         Left            =   195
         TabIndex        =   0
         Top             =   135
         Width           =   1065
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   10410
      TabIndex        =   2
      Top             =   1155
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
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
      Picture         =   "frmCP_Charge_Invoice_Reg.frx":1E04
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   10410
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   525
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&OK"
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
      Picture         =   "frmCP_Charge_Invoice_Reg.frx":3236
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   6
      Left            =   10410
      TabIndex        =   7
      Top             =   1155
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
      Picture         =   "frmCP_Charge_Invoice_Reg.frx":39B0
   End
End
Attribute VB_Name = "frmCP_Charge_Invoice_Reg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pxeMODULENAME = "frmCP_Charge_Invoice_Reg"
Private Const pxeAPPNAME = "Charge Invoice History"
Private oTrans As ggcCPSales.clsCPChargeInvoice
Attribute oTrans.VB_VarHelpID = -1

Private oSkin As clsFormSkin

Dim pbClosedTrans As Boolean
Dim pbLoad As Boolean

Private Sub cmdButton_Click(Index As Integer)
   Dim lnRep As Long
   
   Select Case Index
   Case 0 'browse
      If oTrans.SearchTransaction() Then
         LoadMaster
         LoadDetail
      End If
   Case 1 'Void
      If txtField(0).Text <> "" Then
         lnRep = MsgBox("Do you want to void this transaction?", vbQuestion & vbYesNo, pxeAPPNAME)
         If lnRep = vbYes Then oTrans.CancelTransaction
         MsgBox "Transaction Save Successfully!!!", vbInformation, "INFO"
      Else
         MsgBox "No Transaction to Update!!!"
      End If
   Case 2 'printing
      If oTrans.Master("cTranStat") = 0 Then
         If oTrans.CloseTransaction(oTrans.Master("sTransNox")) = True Then
            lnRep = MsgBox("Do you want to print this transaction?", vbQuestion & vbYesNo, pxeAPPNAME)
               If lnRep = vbYes Then
                  PrinTrans
               Else
                  lnRep = MsgBox("Unable to print transaction", vbInformation)
               End If
         End If
      Else
         lnRep = MsgBox("Do you want to re-print Transaction???", vbQuestion & vbYesNo, pxeAPPNAME)
            If lnRep = vbYes Then PrinTrans
      End If
   Case 3 'close
      Unload Me
   Case 4 'pay
      If txtField(0).Text <> "" Then
         Call initButton(0)
      Else
         MsgBox "No Transaction to Update!!!"
      End If
   Case 5 'OK
      If oTrans.Master("cTranStat") = xeStatePosted Then
         If oTrans.Master("nAmtPaidx") > 0# Then
            If oTrans.PayTransaction(oTrans.Master("sTransNox")) Then
               MsgBox "Transaction Save Successfully!!!", vbInformation, "INFO"
               Call InitForm
               Call initButton(1)
               Call LoadMaster
               Call LoadDetail
            Else
               MsgBox "Unable to Pay Transaction!!!" & vbCrLf & _
                        "Please contact GGC SSG/SEG for assistance!!!", vbExclamation, "INFO"
            End If
         Else
            MsgBox "Unable to Pay Transaction!!!" & vbCrLf & _
                        "Please check Amount Paid!!!", vbExclamation, "INFO"
         End If
      ElseIf oTrans.Master("cTranstat") = xeStateOpen Then
         MsgBox "Charge Invoice not yet Printed!!!" & vbCrLf & _
                     "Print Transaction then try again!", vbExclamation, "INFO"
      ElseIf oTrans.Master("cTranstat") = xeStateClosed Then
         MsgBox "Unable to Pay Transaction!!!" & vbCrLf & _
                  "Please contact Finance Department to confirm Invoice!!!", vbExclamation, "INFO"
      ElseIf oTrans.Master("cTranstat") = xeStateUnknown Then
         MsgBox "Transaction Paid Already!!!" & vbCrLf & _
                  "Please Select other Invoice!!!", vbExclamation, "INFO"
      End If
   Case 6 'Cancel
      Call initButton(1)
      Call LoadMaster
      Call LoadDetail
   End Select
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
               SetNextFocus
            Case vbKeyUp
               SetPreviousFocus
         End Select
      Case vbKeyF12
         If pbLoad = False Then
            txtField(13).Enabled = True
            txtField(13).SetFocus
         End If
   End Select
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String

   lsOldProc = "Form_Load"
    ''On Error GoTo errProc

   CenterChildForm mdiMain, Me
   
   Set oTrans = New ggcCPSales.clsCPChargeInvoice
   Set oTrans.AppDriver = oApp
      
   oTrans.InitTransaction
   
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualRight
      
   InitGrid
   InitForm
   initButton (1)

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
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

Private Sub oTrans_MasterRetrieved(ByVal Index As Integer)
   Select Case Index
   Case 13
      txtField(Index) = Format(oTrans.Master(Index), "#,##0.00")
   Case Else
      txtField(Index) = oTrans.Master(Index)
   End Select
End Sub

Private Sub LoadDetail()
   Dim lnRow As Integer
   Dim lnCol As Integer
   Dim lnSubTotal As Currency
   Dim lsSQL As String
   Dim lors As Recordset
   
   lsSQL = "SELECT " & _
            " a.sTransNox" & _
            ", b.sSerialNo" & _
            ", c.sBarrCode" & _
            ", c.sDescript" & _
            ", a.nQuantity" & _
            ", a.nUnitPrce" & _
            ", a.nDiscRate" & _
            ", a.nDiscAmtx" & _
            " FROM CP_CO_Detail a" & _
               " LEFT JOIN CP_Inventory_Serial b" & _
                  " ON a.sSerialID = b.sSerialID" & _
               " LEFT JOIN CP_Inventory c" & _
                  " ON a.sStockIDx = c.sStockIDx" & _
            " WHERE a.sTransNox = " & strParm(oTrans.Master("sTransNox"))
            
   Set lors = New Recordset
   lors.Open lsSQL, oApp.Connection, adOpenDynamic, adLockReadOnly, adCmdText
         
    With GridEditor1
      .Rows = lors.RecordCount + 1

      lnRow = 1
      Do
         For lnCol = 0 To .Cols - 1
            If lnCol = 0 Then 'No
               .TextMatrix(lnRow, lnCol) = lnRow
            ElseIf lnCol = 1 Then 'imei
               .CellAlignment = flexAlignRightCenter
               .TextMatrix(lnRow, lnCol) = IFNull(lors("sSerialNo"), lors("sBarrCode"))
            ElseIf lnCol = 2 Then 'desc
               .TextMatrix(lnRow, lnCol) = lors("sDescript")
            ElseIf lnCol = 3 Then 'qty
               .TextMatrix(lnRow, lnCol) = lors("nQuantity")
            ElseIf lnCol = 4 Then 'sel price
               .TextMatrix(lnRow, lnCol) = Format(lors("nUnitPrce"), "#,##,0.00")
            ElseIf lnCol = 5 Then 'disc
               .TextMatrix(lnRow, lnCol) = lors("nDiscRate")
            ElseIf lnCol = 6 Then 'disc amt
               .TextMatrix(lnRow, lnCol) = Format(lors("nDiscAmtx"), "#,##,0.00")
            ElseIf lnCol = 7 Then 'total
               .TextMatrix(lnRow, lnCol) = Format(lors("nQuantity") * lors("nUnitPrce"), "#,##,0.00")
            End If
         Next
         lnRow = lnRow + 1
         lors.MoveNext
      Loop Until lors.EOF
   End With
End Sub

Private Sub InitGrid()
   Dim lsOldProc As String
   Dim lnCtr As Integer

   lsOldProc = pxeMODULENAME & ".initGrid"
   ''On Error GoTo errProc
   
   With GridEditor1
      .Refresh
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
      
      .ColEnabled(1) = False
      .ColEnabled(2) = False
      .ColEnabled(3) = False
      .ColEnabled(4) = False
      .ColEnabled(5) = False
      .ColEnabled(6) = False
      .ColEnabled(7) = False
     
      .Col = 0
   End With
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub LoadMaster()
   Dim pnCtr As Integer
   
   For pnCtr = 0 To txtField.Count - 1
      If pnCtr = 14 Then
         txtField(pnCtr) = IFNull(oTrans.Master(2), "")
      ElseIf pnCtr = 15 Then
         txtField(pnCtr) = IFNull(oTrans.Master(3), "")
      ElseIf pnCtr = 9 Then
      ElseIf pnCtr = 10 Then
      ElseIf pnCtr = 11 Then
      ElseIf pnCtr = 13 Then
         txtField(pnCtr) = Format(oTrans.Master("nTranTotl"), "#,##0.00")
      Else
         txtField(pnCtr) = IFNull(oTrans.Master(pnCtr), "")
      End If
   Next
   
   lblTrantotal = Format(oTrans.Master("nTranTotl"), "#,##0.00")
   Label2.Caption = Format(COTransStat(oTrans.Master("cTranStat")), ">")
End Sub

Private Function COTransStat(nStat As Integer) As String
   Select Case nStat
   Case 0
      COTransStat = "OPEN"
   Case 1
      COTransStat = "CLOSED"
   Case 2
      COTransStat = "POSTED"
   Case 3
      COTransStat = "CANCELLED"
   Case 4
      COTransStat = "PAID"
   End Select
End Function

Private Sub Form_Unload(Cancel As Integer)
   Set oTrans = Nothing
End Sub

Private Sub GridEditor1_EditorValidate(Cancel As Boolean)
   With GridEditor1
      If .Col = 5 Then
         .TextMatrix(.Row, 7) = (.TextMatrix(.Row, 3) * ((100 - .TextMatrix(.Row, 5)) / 100) * .TextMatrix(.Row, 4)) - .TextMatrix(.Row, 6)
         .TextMatrix(.Row, 7) = Format(.TextMatrix(.Row, 7), "#,##0.00")
         oTrans.Detail(.Row - 1, "nDiscRate") = .TextMatrix(.Row, 5)
         .TextMatrix(.Row, .Col) = Format(.TextMatrix(.Row, .Col), "##0.00")
      ElseIf .Col = 6 Then
         .TextMatrix(.Row, 7) = (.TextMatrix(.Row, 3) * ((100 - .TextMatrix(.Row, 5)) / 100) * .TextMatrix(.Row, 4)) - .TextMatrix(.Row, 6)
         .TextMatrix(.Row, 7) = Format(.TextMatrix(.Row, 7), "#,##0.00")
         oTrans.Detail(.Row - 1, "nDiscAmtx") = .TextMatrix(.Row, 6)
        .TextMatrix(.Row, .Col) = Format(.TextMatrix(.Row, .Col), "#,##0.00")
      End If
   End With
   
   Call GrandTotal
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = oApp.getColor("HT1")
   End With
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   With txtField(Index)
      If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
         Select Case Index
         Case 14
            If oTrans.SearchTransaction(.Text, True) Then
               LoadMaster
               LoadDetail
            End If
         Case 15
            If oTrans.SearchTransaction(.Text, False) Then
               LoadMaster
               LoadDetail
            End If
         End Select
      End If
   End With
End Sub

'To do
Private Function PrintSI()
   '
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
   
   lblTrantotal = Format(0#, "#,##0.00")
   chkClientTp.Value = 0
   
   oTrans.Master("nAmtPaidx") = CDbl(txtField(13).Text)
      
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   With txtField(Index)
      If Index = 13 Then
         If Not IsNumeric(.Text) Then .Text = 0#
         oTrans.Master("nAmtPaidx") = CDbl(.Text)
         .Text = Format(.Text, "#,##0.00")
      End If
   End With
End Sub

Private Sub initButton(lnStat As Integer)
   Dim lbShow As Boolean

   lbShow = IIf(lnStat = 0, False, True)
   cmdButton(5).Visible = Not lbShow
   cmdButton(6).Visible = Not lbShow

   cmdButton(0).Visible = lbShow
   cmdButton(1).Visible = lbShow
   cmdButton(2).Visible = lbShow
   cmdButton(3).Visible = lbShow
   cmdButton(4).Visible = lbShow
   
   With GridEditor1
      .ColEnabled(5) = Not lbShow
      .ColEnabled(6) = Not lbShow
      If Not lbShow Then .Col = 5
   End With
   
   If lbShow = True Then
      pbLoad = True
   Else
      pbLoad = False
   End If
End Sub

Private Sub GrandTotal()
   Dim lsOldProc As String
   Dim lnCtr As Integer
   Dim lnTotal As Currency

   lsOldProc = pxeMODULENAME & ".GrandTotal"
   ''On Error GoTo errProc
   
   With GridEditor1
      lnTotal = 0#
      For lnCtr = 1 To .Rows - 1
         lnTotal = lnTotal + CDbl(IIf(.TextMatrix(lnCtr, 7) = "", 0, .TextMatrix(lnCtr, 7)))
      Next
   End With
   lblTrantotal.Caption = Format(lnTotal, "#,##0.00")
   txtField(13).Text = Format(lnTotal, "#,##0.00")
   
   oTrans.Master("nTranTotl") = CDbl(lblTrantotal)
   oTrans.Master("nAmtPaidx") = 0#  'oTrans.Master("nTranTotl")
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub
