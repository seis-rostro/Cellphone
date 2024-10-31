VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmCPInvCount 
   BorderStyle     =   0  'None
   Caption         =   "SP Inventory Count"
   ClientHeight    =   6750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16050
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6750
   ScaleWidth      =   16050
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   7
      Left            =   75
      TabIndex        =   16
      Top             =   1230
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   1058
      Caption         =   "&New/Edit"
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
      Picture         =   "frmCPInvCount.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   9
      Left            =   75
      TabIndex        =   21
      Top             =   4380
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   1058
      Caption         =   "C&lose"
      AccessKey       =   "l"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCPInvCount.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   8
      Left            =   75
      TabIndex        =   17
      Top             =   1860
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   1058
      Caption         =   "&Register"
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
      Picture         =   "frmCPInvCount.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   75
      TabIndex        =   10
      Top             =   600
      Width           =   1320
      _ExtentX        =   2328
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
      Picture         =   "frmCPInvCount.frx":15EE
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   11
      Left            =   75
      TabIndex        =   19
      Top             =   3120
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   1058
      Caption         =   "A&pproved"
      AccessKey       =   "p"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCPInvCount.frx":1D68
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   10
      Left            =   75
      TabIndex        =   18
      Top             =   2490
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   1058
      Caption         =   "&Verify"
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
      Picture         =   "frmCPInvCount.frx":24E2
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   600
      Index           =   12
      Left            =   75
      TabIndex        =   20
      Top             =   3750
      Width           =   1320
      _ExtentX        =   2328
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
      Picture         =   "frmCPInvCount.frx":2C5C
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   600
      Index           =   6
      Left            =   75
      TabIndex        =   15
      Top             =   3750
      Width           =   1320
      _ExtentX        =   2328
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
      Picture         =   "frmCPInvCount.frx":33D6
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   75
      TabIndex        =   14
      Top             =   3120
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   1058
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
      Picture         =   "frmCPInvCount.frx":3B50
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   75
      TabIndex        =   11
      Top             =   1230
      Width           =   1320
      _ExtentX        =   2328
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
      Picture         =   "frmCPInvCount.frx":42CA
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   600
      Index           =   3
      Left            =   75
      TabIndex        =   12
      Top             =   1860
      Width           =   1320
      _ExtentX        =   2328
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
      Picture         =   "frmCPInvCount.frx":4A44
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   75
      TabIndex        =   13
      Top             =   2490
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   1058
      Caption         =   "&Add"
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
      Picture         =   "frmCPInvCount.frx":51BE
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   5985
      Left            =   1590
      Tag             =   "wt0;fb0"
      Top             =   585
      Width           =   14325
      _ExtentX        =   25268
      _ExtentY        =   10557
      BackColor       =   12632256
      BorderStyle     =   1
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   5685
         Left            =   5430
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   90
         Width           =   8760
         _ExtentX        =   15452
         _ExtentY        =   10028
         _Version        =   393216
         Rows            =   23
         Cols            =   5
         AllowUserResizing=   3
      End
      Begin xrControl.xrFrame xrFrame2 
         Height          =   2205
         Left            =   75
         Tag             =   "wt0;fb0"
         Top             =   90
         Width           =   5310
         _ExtentX        =   9366
         _ExtentY        =   3889
         BackColor       =   12632256
         ClipControls    =   0   'False
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Index           =   0
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   24
            TabStop         =   0   'False
            Text            =   "C001-01-123456"
            Top             =   135
            Width           =   1530
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   1095
            Index           =   2
            Left            =   1080
            TabIndex        =   3
            Text            =   "Text1"
            Top             =   990
            Width           =   4080
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   1
            Left            =   1080
            TabIndex        =   1
            Text            =   "September 12, 2017"
            Top             =   645
            Width           =   1620
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "APPROVED"
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
            Left            =   3345
            TabIndex        =   23
            Tag             =   "eb0;et0"
            Top             =   180
            Width           =   1680
         End
         Begin VB.Label Label1 
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
            Index           =   3
            Left            =   120
            TabIndex        =   22
            Top             =   225
            Width           =   1110
         End
         Begin VB.Shape Shape4 
            Height          =   330
            Index           =   0
            Left            =   3300
            Top             =   135
            Width           =   1785
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Remarks"
            Height          =   300
            Index           =   1
            Left            =   120
            TabIndex        =   2
            Top             =   990
            Width           =   915
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Date "
            Height          =   300
            Index           =   0
            Left            =   120
            TabIndex        =   0
            Top             =   645
            Width           =   915
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   360
            Left            =   1095
            Tag             =   "et0;ht2"
            Top             =   180
            Width           =   1590
         End
      End
      Begin xrControl.xrFrame xrFrame3 
         Height          =   3465
         Left            =   75
         Tag             =   "wt0;fb0"
         Top             =   2310
         Width           =   5310
         _ExtentX        =   9366
         _ExtentY        =   6112
         BackColor       =   12632256
         ClipControls    =   0   'False
         Begin VB.TextBox txtOthers 
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   3
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   40
            TabStop         =   0   'False
            Text            =   "Text1"
            Top             =   2100
            Width           =   2460
         End
         Begin VB.TextBox txtOthers 
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   2
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   39
            TabStop         =   0   'False
            Text            =   "G4 Metallic/Non-Metallic"
            Top             =   1755
            Width           =   2460
         End
         Begin VB.TextBox txtOthers 
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   1
            Left            =   1080
            TabIndex        =   34
            TabStop         =   0   'False
            Text            =   "Text1"
            Top             =   1410
            Width           =   2160
         End
         Begin VB.TextBox txtOthers 
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   5
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   33
            TabStop         =   0   'False
            Text            =   "Text1"
            Top             =   2445
            Width           =   4095
         End
         Begin VB.TextBox txtOthers 
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   4
            Left            =   3555
            TabIndex        =   32
            TabStop         =   0   'False
            Text            =   "Midnight Black"
            Top             =   1755
            Width           =   1605
         End
         Begin VB.TextBox txtOthers 
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   81
            Left            =   1080
            TabIndex        =   31
            Text            =   "Text1"
            Top             =   435
            Width           =   960
         End
         Begin VB.TextBox txtOthers 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Index           =   9
            Left            =   4200
            Locked          =   -1  'True
            TabIndex        =   30
            TabStop         =   0   'False
            Text            =   "205"
            Top             =   435
            Width           =   960
         End
         Begin VB.TextBox txtOthers 
            Appearance      =   0  'Flat
            Height          =   615
            Index           =   13
            Left            =   1080
            TabIndex        =   27
            Text            =   "Text1"
            Top             =   780
            Width           =   4080
         End
         Begin VB.TextBox txtOthers 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Index           =   8
            Left            =   4200
            Locked          =   -1  'True
            TabIndex        =   25
            TabStop         =   0   'False
            Text            =   "205"
            Top             =   90
            Width           =   960
         End
         Begin VB.TextBox txtOthers 
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   80
            Left            =   1080
            TabIndex        =   5
            Text            =   "Text1"
            Top             =   90
            Width           =   960
         End
         Begin xrControl.xrButton cmdButton 
            Height          =   495
            Index           =   1
            Left            =   2610
            TabIndex        =   28
            Top             =   2835
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   873
            Caption         =   "&Next"
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
            Picture         =   "frmCPInvCount.frx":6250
         End
         Begin xrControl.xrButton cmdButton 
            CausesValidation=   0   'False
            Height          =   495
            Index           =   13
            Left            =   3915
            TabIndex        =   29
            Top             =   2835
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   873
            Caption         =   "Seria&l"
            AccessKey       =   "l"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Picture         =   "frmCPInvCount.frx":A162
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Model"
            Height          =   195
            Index           =   1
            Left            =   135
            TabIndex        =   38
            Top             =   1830
            Width           =   435
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Brand"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   37
            Top             =   2175
            Width           =   420
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Barcode"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   36
            Top             =   1485
            Width           =   600
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   35
            Top             =   2520
            Width           =   795
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Serial Count"
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
            Index           =   5
            Left            =   3105
            TabIndex        =   26
            Top             =   150
            Width           =   1050
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Entry No"
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   4
            Top             =   150
            Width           =   615
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "QOH"
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
            Index           =   4
            Left            =   3120
            TabIndex        =   8
            Top             =   465
            Width           =   420
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Actual Count"
            Height          =   195
            Index           =   9
            Left            =   120
            TabIndex        =   6
            Top             =   465
            Width           =   915
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Notes"
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   7
            Top             =   765
            Width           =   420
         End
      End
   End
End
Attribute VB_Name = "frmCPInvCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Private Const pxeMODULENAME = "frmCPInvCount"
'
'Private Const pxeMaxVisible1 As Integer = 23
'Private Const pxeMaxVisible2 As Integer = 11
'
'Private WithEvents oTrans As clsCPInvCount
'Private oSkin As clsFormSkin
'
'Dim pnIndex As Integer
'Dim pbtxtFldFocus As Boolean
'Dim pnCtr As Integer
'Dim pnRow As Integer
'
'Dim pnCntLvl As Integer
'Dim pbLoaded As Boolean
'
'Property Let CountLevel(Value As Integer)
'   pnCntLvl = Value
'End Property
'
'Private Sub cmdButton_Click(Index As Integer)
'   Dim lsProcName As String
'   Dim lsRep As String
'
'   lsProcName = "cmdButton_Click"
'   ''On Error GoTo errProc
'
'   If pnIndex >= 0 Then
'      If pbtxtFldFocus Then
'         txtField_LostFocus pnIndex
'      Else
'         txtOthers_LostFocus pnIndex
'      End If
'   End If
'
'   With MSFlexGrid1
'      Select Case Index
'      Case 0 'Print
'         If oTrans.EditMode = xeModeReady Then
'            Call PrintTrans
'         ElseIf oTrans.EditMode = xeModeUpdate Or oTrans.EditMode = xeModeAddNew Then
'            If MsgBox("This will save this record!!" & vbCrLf & _
'                     "Do you want to continue?", vbYesNo + vbQuestion, "Confirm") = vbYes Then
'               If oTrans.SaveTransaction = True Then
'                  ClearFields
'                  LoadDetail
'                  Call PrintTrans
'               End If
'            End If
'         End If
'      Case 1 'Next
'         pnRow = pnRow + 1
'         If pnRow > 0 And pnRow < oTrans.ItemCount + 1 Then
'
'            If Not MSFlexGrid1.RowIsVisible(pnRow) Then MSFlexGrid1.TopRow = pnRow
'            MSFlexGrid1.Row = pnRow
'
'            txtOthers(1) = oTrans.Detail(pnRow - 1, "sBarrCode")
'            txtOthers(2) = oTrans.Detail(pnRow - 1, "sModelNme")
'            txtOthers(3) = IFNull(oTrans.Detail(pnRow - 1, "sBrandNme"))
'            txtOthers(4) = IFNull(oTrans.Detail(pnRow - 1, "sColorNme"))
'            txtOthers(5) = IFNull(oTrans.Detail(pnRow - 1, "sDescript"))
'            txtOthers(8) = oTrans.Detail(pnRow - 1, "nQtyOnHnd")
'            txtOthers(9) = oTrans.Detail(pnRow - 1, "nSerialCt")
'            txtOthers(13) = oTrans.Detail(pnRow - 1, "sRemarksx")
'            txtOthers(80) = oTrans.Detail(pnRow - 1, "nEntryNox")
'
'            Select Case pnCntLvl
'            Case 1
'               txtOthers(81) = oTrans.Detail(pnRow - 1, "nActCtr01")
'            Case 2
'               txtOthers(81) = IIf(oTrans.Detail(pnRow - 1, "nActCtr02") < 0, "", oTrans.Detail(pnRow - 1, "nActCtr02"))
'            Case 3
'               txtOthers(81) = IIf(oTrans.Detail(pnRow - 1, "nActCtr03") < 0, "", oTrans.Detail(pnRow - 1, "nActCtr03"))
'            End Select
'
'            If xrFrame3.Enabled Then txtOthers(81).SetFocus
'         Else
'            txtOthers(1) = ""
'            txtOthers(2) = ""
'            txtOthers(3) = ""
'            txtOthers(4) = ""
'            txtOthers(5) = ""
'
'            txtOthers(13) = ""
'
'            txtOthers(8) = 0
'            txtOthers(9) = 0
'
'            txtOthers(80) = 0
'            txtOthers(81) = 0
'
'            pnRow = .Rows
'         End If
'      Case 2 'Save
'         If oTrans.SaveTransaction = True Then
'            MsgBox "Transaction Saved Successfully!!!", vbInformation, "Notice"
'            initButton xeModeReady
'         End If
''      Case 3 'Search Row
''         If oTrans.EditMode = xeModeUpdate Then
''            If oTrans.Detail(MSFlexGrid1.Row - 1, "nQtyOnHnd") = 0 Then
''               If (pnIndex = 1 Or pnIndex = 2) And pbtxtFldFocus = False Then
''                  Call oTrans.searchDetail(MSFlexGrid1.Row - 1, 1, txtOthers(pnIndex))
''               End If
''            End If
''         End If
'      Case 4 'Add Row
'         If oTrans.EditMode = xeModeUpdate Then
'            If .TextMatrix(.Rows - 1, .Col) <> "" Then
'               .TopRow = .Rows - 1
'               If oTrans.addDetail Then
'                  .AddItem ""
'
'                  If .Rows > 26 Then .ColWidth(2) = 2780
'
'                  .TextMatrix(.Rows - 1, 0) = oTrans.ItemCount
'                  txtOthers(80) = oTrans.ItemCount
'                  txtOthers(1) = ""
'                  txtOthers(2) = ""
'                  txtOthers(3) = ""
'                  txtOthers(4) = ""
'                  txtOthers(5) = ""
'                  txtOthers(8) = 0
'                  txtOthers(9) = 0
'                  txtOthers(13) = ""
'
'                  txtOthers(1).SetFocus
'               End If
'            End If
'         End If
''      Case 5 'Delete Row
''         If oTrans.EditMode = xeModeUpdate Then
''            If .Rows > 2 Then
''               If oTrans.deleteDetail(.Row - 1) Then .RemoveItem .Row
''
''               If .Rows > 26 Then
''                  .ColWidth(2) = 2780
''               Else
''                  .ColWidth(2) = 3080
''               End If
''            End If
''         End If
'      Case 6 'Cancel Update
'         lsRep = MsgBox("Transaction is in Update Mode!!!" & vbCrLf & _
'                        "Do you want to Cancel Transaction!!!", vbYesNo + vbQuestion, "Confirm")
'
'         If lsRep = vbYes Then
'            If oTrans.EditMode = xeModeAddNew Then
'               If oTrans.InitTransaction Then
'                  ClearFields
'                  initButton xeModeReady
'               Else
'                  MsgBox "Cannot cancel the update of this CP Inventory Count!!!", vbInformation, "Notice"
'               End If
'            Else
'               If oTrans.NewTransaction Then
'                  ClearFields
'                  LoadDetail
'                  initButton xeModeReady
'               Else
'                  MsgBox "Cannot cancel the update of this CP Inventory Count!!!", vbInformation, "Notice"
'               End If
'            End If
'         Else
'            txtField(pnIndex).SetFocus
'         End If
'      Case 7 'New
'         If oTrans.EditMode = xeModeUnknown Then
'            If oTrans.NewTransaction Then
'               ClearFields
'               LoadDetail
'               If oTrans.Master("cTranStat") = xeStateOpen Then
'                  initButton xeModeAddNew
'                  txtField(1).SetFocus
'               End If
'            Else
'               MsgBox "Cannot create new CP Inventory Count!!!", vbInformation, "Notice"
'            End If
'         Else
'            If oTrans.UpdateTransaction Then
'               initButton xeModeAddNew
'               txtField(1).SetFocus
'            Else
'               MsgBox "Cannot update CP Inventory Count!!!", vbInformation, "Notice"
'            End If
'         End If
'      Case 8 'Register
'         'frmSPCountReg.Tag = ""
'         'frmSPCountReg.Show
'      Case 9 'Close
'         Unload Me
'      Case 10 'Verify
'         If oTrans.EditMode = xeModeReady Then
'            If oTrans.CloseTransaction(oTrans.Master("sTransNox")) Then
'               MsgBox "CP Inventory Count was verified successfully!!!", vbInformation, "Notice"
'            Else
'               MsgBox "Unable to verify CP Inventory Count!!!", vbInformation, "Notice"
'            End If
'         End If
'      Case 11 'Approve
'         If oTrans.EditMode = xeModeReady Then
'            If oTrans.PostTransaction(oTrans.Master("sTransNox")) Then
'               MsgBox "CP Inventory Count was posted/approved successfully!!!", vbInformation, "Notice"
'            Else
'               MsgBox "Unable to post/approved CP Inventory Count!!!", vbInformation, "Notice"
'            End If
'         End If
'      Case 12 'Disapprove/Cancel
'         If oTrans.EditMode = xeModeReady Then
'            If oTrans.CancelTransaction() Then
'               MsgBox "CP Inventory Count was cancelled successfully!!!", vbInformation, "Notice"
'            Else
'               MsgBox "Unable to cancel CP Inventory Count!!!", vbInformation, "Notice"
'            End If
'         End If
'      Case 13 ' Serial
'         Call LoadSerial
'      Case 14 ' Next Serial
'      Case 15 ' Cancel Serial Update
'         Call LoadSerial
'      Case 16 ' Confirm Serial Update
''         Call processSerial
'      End Select
'   End With
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsProcName & "( " & Index & " )", True
'End Sub
'
'Private Sub cmdDetail_Click(Index As Integer)
'   Select Case Index
'   Case 0 ' Ok
'      Call setSerial
'   Case 1 ' Next
'      With MSFlexGrid2
'         If .Row = .Rows - 1 Then
'            .Rows = .Rows + 1
'         End If
'         .Row = .Row + 1
'         Call showSerial(.Row - 1)
'      End With
'   Case 2 ' Cancel
'      Call clearSerial
'   End Select
'End Sub
'
'Private Sub Form_Activate()
'   Dim lbValid As Boolean
'   If Not pbLoaded Then
'      InitForm
'      ClearFields
'
'      lbValid = False
'      If pnCntLvl > 1 Then
'         If oTrans.NewTransaction Then
'            If oTrans.EditMode = xeModeReady Then
'               lbValid = True
'            End If
'         End If
'      End If
'
'      If lbValid Then
'         LoadDetail
'         If oTrans.Master("cTranStat") = xeStateOpen Then
'            oTrans.UpdateTransaction
'            initButton xeModeUpdate
'            txtOthers(80).SetFocus
'            Debug.Print oTrans.EditMode
'         Else
'            initButton xeModeReady
'            cmdButton(7).SetFocus
'         End If
'      Else
'         initButton xeModeReady
'         cmdButton(7).SetFocus
'      End If
'
'      pbLoaded = True
'      pnIndex = -1
'   End If
'
'   oApp.MenuName = Me.Tag
'   Me.ZOrder 0
'   MSFlexGrid1.Refresh
'End Sub
'
'Private Sub Form_Load()
'   Dim lsProcName As String
'
'   lsProcName = "Form_Load"
'   ''On Error GoTo errProc
'
'   CenterChildForm mdiMain, Me
'
'   Set oTrans = New clsCPInvCount
'   Set oTrans.AppDriver = oApp
'
'   oTrans.InitTransaction
'
'   Set oSkin = New clsFormSkin
'   Set oSkin.AppDriver = oApp
'   Set oSkin.Form = Me
'   oSkin.ApplySkin xeFormTransaction
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsProcName & "( " & " )", True
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'   Set oTrans = Nothing
'   Set oSkin = Nothing
'   pbLoaded = False
'End Sub
'
'Private Sub oTrans_DetailRetrieved(ByVal Index As Integer)
'   With MSFlexGrid1
'      Select Case Index
'      Case 7, 8, 9
'         If oTrans.Detail(.Row - 1, Index) < 0 Then
'            txtOthers(81) = ""
'            .TextMatrix(.Row, Index) = ""
'         Else
'            txtOthers(81) = oTrans.Detail(.Row - 1, Index)
'            .TextMatrix(.Row, Index) = oTrans.Detail(.Row - 1, Index)
'         End If
'      Case Else
'         txtOthers(Index) = oTrans.Detail(.Row - 1, Index)
'         .TextMatrix(.Row, Index) = oTrans.Detail(.Row - 1, Index)
'      End Select
'   End With
'End Sub
'
'Private Sub oTrans_MasterRetrieved(ByVal Index As Integer)
'   If Index = 1 Then
'      txtField(Index).Text = Format(oTrans.Master(Index), "MMMM DD, YYYY")
'   Else
'      txtField(Index).Text = oTrans.Master(Index)
'   End If
'End Sub
'
'Private Sub txtField_GotFocus(Index As Integer)
'   With txtField(Index)
'      If Index = 1 Then .Text = Format(.Text, "MM/DD/YYYY")
'
'      .SelStart = 0
'      .SelLength = Len(.Text)
'      .BackColor = oApp.getColor("HT1")
'   End With
'
'   pbtxtFldFocus = True
'   pnIndex = Index
'End Sub
'
'Private Sub initButton(lnStat As Integer)
'   Dim lbShow As Boolean
'
'   lbShow = IIf(lnStat = 0, False, True)
'   If oTrans.EditMode = xeModeReady Then
'      cmdButton(7).Caption = "&Update"
'   Else
'      cmdButton(7).Caption = "&New"
'   End If
'   cmdButton(7).Visible = Not lbShow
'   cmdButton(8).Visible = Not lbShow
'   cmdButton(9).Visible = Not lbShow
'   cmdButton(10).Visible = Not lbShow
'   cmdButton(11).Visible = Not lbShow
'   cmdButton(12).Visible = Not lbShow
'
'   cmdButton(2).Visible = lbShow
'   cmdButton(3).Visible = lbShow
'   cmdButton(4).Visible = lbShow
'   cmdButton(5).Visible = lbShow
'   cmdButton(6).Visible = lbShow
'
'   xrFrame1.Enabled = lbShow
'   xrFrame2.Enabled = lbShow
'   xrFrame3.Enabled = lbShow
'End Sub
'
'Private Sub InitForm()
'   Dim lnCtr As Integer
'
'   With MSFlexGrid1
'      .Cols = 14
'      .Rows = 2
'      .Font = "MS San Serif"
'
'      'Column Title
'      .TextMatrix(0, 1) = "Barcode"
'      .TextMatrix(0, 2) = "Model"
'      .TextMatrix(0, 3) = "Brand"
'      .TextMatrix(0, 4) = "Color"
'      .TextMatrix(0, 5) = "Description"
'      .TextMatrix(0, 6) = "Warehouse"
'      .TextMatrix(0, 7) = "Section"
'      .TextMatrix(0, 8) = "Srl Ct"
'      .TextMatrix(0, 9) = "QOH"
'      .TextMatrix(0, 10) = "Cnt 1"
'      .TextMatrix(0, 11) = "Cnt 2"
'      .TextMatrix(0, 12) = "Cnt 3"
'      .TextMatrix(0, 13) = "Notes"
'
'      .Row = 0
'
'      'Column Alignment
'      For lnCtr = 0 To .Cols - 1
'         .Col = lnCtr
'         .CellFontBold = True
'         .ColAlignment(lnCtr) = 1
'         .CellAlignment = 1
'      Next
'
'      'Column Width
'      .ColWidth(0) = 450
'      .ColWidth(1) = 1920
'      .ColWidth(2) = 2350
'      .ColWidth(3) = 0
'      .ColWidth(4) = 2000
'      .ColWidth(5) = 0
'      .ColWidth(6) = 0
'      .ColWidth(7) = 0
'      .ColWidth(8) = 650
'      .ColWidth(9) = 650
'      .ColWidth(10) = 650
'      .ColWidth(11) = 0
'      .ColWidth(12) = 0
'      .ColWidth(13) = 0
'
'      .ColAlignment(8) = 7
'      .ColAlignment(9) = 7
'
'      .Col = 1
'      .Row = 1
'   End With
'
'   With MSFlexGrid2
'      .Cols = 13
'      .Rows = 2
'      .Font = "MS San Serif"
'
'      'Column Title
'      .TextMatrix(0, 1) = "Serial ID"
'      .TextMatrix(0, 2) = "IMEI"
'      .TextMatrix(0, 3) = "Old-Loc"
'      .TextMatrix(0, 4) = "Old-Stat"
'      .TextMatrix(0, 5) = "Old-Branch"
'      .TextMatrix(0, 6) = "Old-Stock"
'      .TextMatrix(0, 7) = "Location"
'      .TextMatrix(0, 8) = "Status"
'      .TextMatrix(0, 9) = "Branch"
'      .TextMatrix(0, 10) = "Stock ID"
'      .TextMatrix(0, 11) = "Branch Cd"
'      .TextMatrix(0, 12) = ""
'
'      .Row = 0
'
'      'Column Alignment
'      For lnCtr = 0 To .Cols - 1
'         .Col = lnCtr
'         .CellFontBold = True
'         .ColAlignment(lnCtr) = 1
'         .CellAlignment = 1
'      Next
'
'      'Column Width
'      .ColWidth(0) = 450
'      .ColWidth(1) = 0 '1920
'      .ColWidth(2) = 1920 ' 2250
'      .ColWidth(3) = 0
'      .ColWidth(4) = 0 '2000
'      .ColWidth(5) = 0
'      .ColWidth(6) = 0
'      .ColWidth(7) = 1300 '0
'      .ColWidth(8) = 900 '650
'      .ColWidth(9) = 3350 '650
'      .ColWidth(10) = 0 '650
'      .ColWidth(11) = 0
'      .ColWidth(12) = 0
'
'      .Col = 1
'      .Row = 1
'   End With
'End Sub
'
'Private Sub ClearFields()
'   Dim loTxt As TextBox
'   If oTrans.Master("sTransNox") = "" Then
'      For Each loTxt In txtField
'         loTxt.Text = ""
'      Next
'
'      Label2 = "UNKNOWN"
'
'   Else
'      txtField(0) = Format(oTrans.Master(0), "000000-000000")
'      txtField(1) = Format(oTrans.Master(1), "MMMM DD, YYYY")
'      txtField(2) = oTrans.Master(2)
'
'      Select Case oTrans.Master("cTranStat")
'      Case xeStateOpen
'         Label2 = "OPEN"
'      Case xeStateClosed
'         Label2 = "VERIFIED"
'      Case xeStatePosted
'         Label2 = "APPROVED"
'      Case xeStateCancelled
'         Label2 = "CANCELLED"
'      Case Else
'         Label2 = "UNKNOWN"
'      End Select
'
'   End If
'
'   For Each loTxt In txtOthers
'      loTxt.Text = ""
'   Next
'
'   txtOthers(81) = 0
'   txtOthers(8) = 0
'   txtOthers(9) = 0
'
'   With MSFlexGrid1
'      .Rows = 2
'      .ColWidth(2) = 2350
'
'      .TextMatrix(1, 1) = ""
'      .TextMatrix(1, 2) = ""
'      .TextMatrix(1, 3) = ""
'      .TextMatrix(1, 4) = ""
'      .TextMatrix(1, 5) = ""
'
'      .TextMatrix(1, 6) = ""
'      .TextMatrix(1, 7) = ""
'      .TextMatrix(1, 8) = 0
'      .TextMatrix(1, 9) = 0
'      .TextMatrix(1, 10) = 0
'      .TextMatrix(1, 11) = 0
'      .TextMatrix(1, 12) = 0
'      .TextMatrix(1, 13) = ""
'
'      .Col = 1
'      .Row = 1
'   End With
'End Sub
'
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'   Select Case KeyCode
'   Case vbKeyReturn, vbKeyUp, vbKeyDown
'      Select Case KeyCode
'      Case vbKeyReturn, vbKeyDown
'         If GetFocus = MSFlexGrid1.hwnd Then Exit Sub
'         SetNextFocus
'      Case vbKeyUp
'         SetPreviousFocus
'      End Select
'   End Select
'End Sub
'
'Private Sub txtField_LostFocus(Index As Integer)
'   With txtField(Index)
'      .BackColor = oApp.getColor("EB")
'   End With
'End Sub
'
'Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
'   With txtField(Index)
'      .Text = TitleCase(.Text)
'      Select Case Index
'      Case 1
'         If Not IsDate(.Text) Then .Text = oApp.ServerDate
'         .Text = Format(.Text, "MMMM DD, YYYY")
'         oTrans.Master(Index) = .Text
'      Case Else
'         oTrans.Master(Index) = .Text
'      End Select
'   End With
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
'
'Private Function PrintTrans() As Boolean
'   Dim lrs As New ADODB.Recordset
'   Dim lnCtr As Integer
'   Dim lsProcName As String
'   Dim loreport As frmRepViewer
'   Dim loDataDiff As Recordset
'   Dim lsSQL As String
'
'   lsProcName = "PrintTrans"
'   ''On Error GoTo errProc
'
'   PrintTrans = True
'
'   Set lrs = New ADODB.Recordset
'
'   lrs.Fields.Append "sField01", adVarChar, 120
'   lrs.Fields.Append "sField02", adVarChar, 30
'   lrs.Fields.Append "sField03", adVarChar, 120
'   lrs.Fields.Append "nField01", adInteger, 10
'   lrs.Fields.Append "nField02", adInteger, 10
'   lrs.Fields.Append "nField03", adInteger, 10
'   lrs.Fields.Append "nField04", adInteger, 10
'   lrs.Fields.Append "nField05", adInteger, 10
'   lrs.Open
'
'   With oTrans
'      For lnCtr = 0 To .ItemCount - 1
'         If pnCntLvl = 1 Then
'            lrs.AddNew
'            lrs("sField01").Value = .Detail(lnCtr, "sWHouseNm") & "" & .Detail(lnCtr, "sSectnNme")
'            lrs("sField02").Value = .Detail(lnCtr, "sBarrCode")
'            lrs("sField03").Value = .Detail(lnCtr, "sDescript")
'            If oApp.BranchCode <> "M001" Then
'               lrs("nField01").Value = .Detail(lnCtr, "nQtyOnHnd")
'            End If
'            lrs("nField02").Value = .Detail(lnCtr, "nEntryNox")
'            lrs("nField03").Value = .Detail(lnCtr, "nActCtr01")
'         ElseIf pnCntLvl = 2 Then
'            If .Detail(lnCtr, "nQtyOnHnd") <> .Detail(lnCtr, "nActCtr01") Then
'               lrs.AddNew
'               lrs("sField01").Value = .Detail(lnCtr, "sWHouseNm") & "" & .Detail(lnCtr, "sSectnNme")
'               lrs("sField02").Value = .Detail(lnCtr, "sBarrCode")
'               lrs("sField03").Value = .Detail(lnCtr, "sDescript")
'               lrs("nField01").Value = .Detail(lnCtr, "nQtyOnHnd")
'               lrs("nField02").Value = .Detail(lnCtr, "nEntryNox")
'               lrs("nField03").Value = .Detail(lnCtr, "nActCtr01")
'               lrs("nField04").Value = .Detail(lnCtr, "nActCtr02")
'            End If
'         ElseIf pnCntLvl = 3 Then
'            If .Detail(lnCtr, "nQtyOnHnd") <> .Detail(lnCtr, "nActCtr01") Then
'               lrs.AddNew
'               lrs("sField01").Value = .Detail(lnCtr, "sWHouseNm") & "" & .Detail(lnCtr, "sSectnNme")
'               lrs("sField02").Value = .Detail(lnCtr, "sBarrCode")
'               lrs("sField03").Value = .Detail(lnCtr, "sDescript")
'               lrs("nField01").Value = .Detail(lnCtr, "nQtyOnHnd")
'               lrs("nField02").Value = .Detail(lnCtr, "nEntryNox")
'               lrs("nField03").Value = .Detail(lnCtr, "nActCtr01")
'               lrs("nField04").Value = .Detail(lnCtr, "nActCtr02")
'               lrs("nField05").Value = .Detail(lnCtr, "nActCtr03")
'            End If
'         End If
'      Next
'   End With
'
'   ' assign important info to the report
'   Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\SPInventoryCount.rpt")
'   oReport.DiscardSavedData
'   oReport.FieldMappingType = crAutoFieldMapping
'   oReport.Database.SetDataSource lrs
'
'   oReport.Sections("RH").ReportObjects("txtCompany").SetText oApp.BranchName
'   oReport.Sections("RH").ReportObjects("txtAddress").SetText oApp.Address
'   oReport.Sections("PH").ReportObjects("txtHeading1").SetText "CP Inventory Count - " & Format(oTrans.Master("dTransact"), "DD MMM YYYY")
'   oReport.Sections("PH").ReportObjects("txtHeading2").SetText "As of " & Format(oApp.ServerDate, "DD MMM YYYY")
'   oReport.Sections("PF").ReportObjects("txtRptUser").SetText oApp.UserName
'
'   Set loreport = New frmRepViewer
'   Set loreport.ReportSource = oReport
'   loreport.Show
''   oReport.PrintOutEx True, 1
'   lrs.Close
'
'endProc:
'   Set oReport = Nothing
'   Exit Function
'errProc:
'   PrintTrans = False
'   ShowError lsProcName & "( " & " )"
'End Function
'
'Private Sub LoadDetail()
'   Dim lnCtr As Integer
'   Dim lnLen As Integer
'
'   With MSFlexGrid1
'      .Rows = IIf(oTrans.ItemCount = 0, 2, oTrans.ItemCount + 1)
'      If .Rows > pxeMaxVisible1 Then
'         .ColWidth(2) = 2100
'      Else
'         .ColWidth(2) = 2350
'      End If
'      For pnCtr = 0 To oTrans.ItemCount - 1
'         DoEvents
'         .TextMatrix(pnCtr + 1, 0) = pnCtr + 1
'         .TextMatrix(pnCtr + 1, 1) = oTrans.Detail(pnCtr, "sBarrCode")
'         .TextMatrix(pnCtr + 1, 2) = oTrans.Detail(pnCtr, "sModelNme")
'         .TextMatrix(pnCtr + 1, 3) = IFNull(oTrans.Detail(pnCtr, "sBrandNme"), "")
'         .TextMatrix(pnCtr + 1, 4) = IFNull(oTrans.Detail(pnCtr, "sColorNme"), "")
'         .TextMatrix(pnCtr + 1, 5) = oTrans.Detail(pnCtr, "sDescript")
'         .TextMatrix(pnCtr + 1, 6) = IFNull(oTrans.Detail(pnCtr, "sWHouseNm"), "")
'         .TextMatrix(pnCtr + 1, 7) = IFNull(oTrans.Detail(pnCtr, "sSectnNme"), "")
'         .TextMatrix(pnCtr + 1, 8) = oTrans.Detail(pnCtr, "nQtyOnHnd")
'         .TextMatrix(pnCtr + 1, 9) = oTrans.Detail(pnCtr, "nSerialCt")
'         .TextMatrix(pnCtr + 1, 10) = oTrans.Detail(pnCtr, "nActCtr01")
'         .TextMatrix(pnCtr + 1, 11) = IIf(oTrans.Detail(pnCtr, "nActCtr02") < 0, "", oTrans.Detail(pnCtr, "nActCtr02"))
'         .TextMatrix(pnCtr + 1, 12) = IIf(oTrans.Detail(pnCtr, "nActCtr03") < 0, "", oTrans.Detail(pnCtr, "nActCtr03"))
'         .TextMatrix(pnCtr + 1, 13) = oTrans.Detail(pnCtr, "sRemarksx")
'      Next
'   End With
'End Sub
'
'Private Function LoadSerial() As Boolean
'   Dim lsProcName As String
'   Dim lors As Recordset
'   Dim loForm As frmCPInvSerial
'
'   lsProcName = "LoadSerial()"
'   'On Error GoTo errProc
'
'   If pnRow = MSFlexGrid1.Rows Then GoTo endProc
'
'   Set lors = oTrans.LoadSerial(pnRow)
'   If TypeName(lors) = "Nothing" Then GoTo endProc
'
'   Set loForm = New frmCPInvSerial
'   Set loForm.Serial = lors
'   loForm.Show 1
'
'
'   With MSFlexGrid2
'      .Rows = IIf(oTrans.ItemCount = 0, 2, oTrans.ItemCount + 2)
'
'      If .Rows > pxeMaxVisible2 Then
'         .ColWidth(2) = 3350
'      Else
'         .ColWidth(2) = 3050
'      End If
'
'      lors.MoveFirst
'      For pnCtr = 0 To oTrans.ItemCount - 1
'         DoEvents
'
'         .TextMatrix(pnCtr + 1, 0) = pnCtr + 1
'         .TextMatrix(pnCtr + 1, 1) = lors("sSerialID")
'         .TextMatrix(pnCtr + 1, 2) = lors("sSerialNo")
'         .TextMatrix(pnCtr + 1, 3) = lors("cLocation")
'         .TextMatrix(pnCtr + 1, 4) = lors("cSoldStat")
'         .TextMatrix(pnCtr + 1, 5) = lors("sBranchCd")
'         .TextMatrix(pnCtr + 1, 6) = lors("sStockIDx")
'         .TextMatrix(pnCtr + 1, 7) = lors("cAdjLocxx")
'         .TextMatrix(pnCtr + 1, 8) = lors("cAdjStatx")
'         .TextMatrix(pnCtr + 1, 9) = lors("sBranchNm")
'         .TextMatrix(pnCtr + 1, 10) = lors("sNewStock")
'         .TextMatrix(pnCtr + 1, 11) = lors("sNewBrnch")
'
'         If oTrans.Detail(pnRow, "sStockIDx") <> lors("sStockIDx") Or _
'               lors("cLocation") <> lors("cAdjLocxx") Or _
'               lors("cSoldStat") <> lors("cAdjStatx") Or _
'               lors("sBranchCd") <> lors("sNewBrnch") Then
'            .TextMatrix(pnCtr + 1, 12) = 1
'         Else
'            .TextMatrix(pnCtr + 1, 12) = 0
'         End If
'
'         lors.MoveNext
'      Next
'   End With
'
'endProc:
'   Exit Function
'errProc:
'   ShowError lsProcName & "( " & " )"
'End Function
'
'Private Sub showSerial(ByVal lnRow As Integer)
'   With MSFlexGrid2
'      If lnRow >= .Rows Then GoTo endProc
'
'      If .TextMatrix(lnRow, 1) = "" Then
'         ' new record, allow IMEI to be modified
'         txtSerial(0) = ""
'         txtSerial(1) = ""
'         txtSerial(3) = ""
'         cmbSerial(0).ListIndex = -1
'         cmbSerial(1).ListIndex = -1
'      Else
'         txtSerial(0) = .TextMatrix(lnRow, 1)
'         txtSerial(1) = .TextMatrix(lnRow, 2)
'         txtSerial(2) = .TextMatrix(lnRow, 9)
'         cmbSerial(0).ListIndex = CInt(.TextMatrix(lnRow, 7))
'         cmbSerial(1).ListIndex = CInt(.TextMatrix(lnRow, 8))
'      End If
'
'      If .TextMatrix(lnRow, 3) <> .TextMatrix(lnRow, 7) Or _
'            .TextMatrix(lnRow, 4) <> .TextMatrix(lnRow, 8) Or _
'            .TextMatrix(lnRow, 5) <> .TextMatrix(lnRow, 11) Or _
'            .TextMatrix(lnRow, 6) <> .TextMatrix(lnRow, 10) Then
'         txtSerial(1).Locked = False
'      Else
'         txtSerial(1).Locked = True
'      End If
'   End With
'
'endProc:
'   Exit Sub
'End Sub
'
'Private Sub txtOthers_GotFocus(Index As Integer)
'   With txtOthers(Index)
'      .SelStart = 0
'      .SelLength = Len(.Text)
'      .BackColor = oApp.getColor("HT1")
'   End With
'
'   pbtxtFldFocus = False
'   pnIndex = Index
'End Sub
'
'Private Sub txtOthers_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'   Select Case Index
'   Case 1, 2
'      If KeyCode = vbKeyReturn Or KeyCode = vbKeyF3 Then
'         Call oTrans.searchDetail(MSFlexGrid1.Row - 1, 1, txtOthers(Index))
'      End If
'   End Select
'End Sub
'
'Private Sub txtOthers_LostFocus(Index As Integer)
'   With txtOthers(Index)
'      .BackColor = oApp.getColor("EB")
'   End With
'End Sub
'
'Private Sub txtOthers_Validate(Index As Integer, Cancel As Boolean)
'   Dim lsField As String
'
'   With MSFlexGrid1
'      Select Case Index
'      Case 13
'         If txtOthers(1) <> "" Then
'            oTrans.Detail(.Row - 1, "sRemarksx") = txtOthers(Index)
'            .TextMatrix(.Row, 13) = txtOthers(Index)
'         End If
'      Case 80
'         pnRow = Val(txtOthers(80))
'         If pnRow > 0 And pnRow < oTrans.ItemCount + 1 Then
'            If Not .RowIsVisible(pnRow) Then
'               If .TopRow > pnRow Then
'                  .TopRow = pnRow
'               Else
'                  .TopRow = pnRow - pxeMaxVisible1
'               End If
'            End If
'            .Row = pnRow
'
'            txtOthers(1) = .TextMatrix(.Row, 1)
'            txtOthers(2) = .TextMatrix(.Row, 2)
'            txtOthers(3) = .TextMatrix(.Row, 3)
'            txtOthers(4) = .TextMatrix(.Row, 4)
'            txtOthers(5) = .TextMatrix(.Row, 5)
'            txtOthers(8) = .TextMatrix(.Row, 8)
'            txtOthers(9) = .TextMatrix(.Row, 9)
'            txtOthers(13) = .TextMatrix(.Row, 13)
'
'            Select Case pnCntLvl
'            Case 1
'               txtOthers(81) = .TextMatrix(.Row, 10)
'            Case 2
'               txtOthers(81) = IIf(.TextMatrix(.Row, 11) < 0, "", oTrans.Detail(pnRow - 1, "nActCtr02"))
'            Case 3
'               txtOthers(81) = IIf(oTrans.Detail(pnRow - 1, "nActCtr03") < 0, "", oTrans.Detail(pnRow - 1, "nActCtr03"))
'            End Select
'         Else
'            txtOthers(1) = ""
'            txtOthers(2) = ""
'            txtOthers(3) = ""
'            txtOthers(4) = ""
'            txtOthers(5) = ""
'            txtOthers(8) = 0
'            txtOthers(9) = 0
'            txtOthers(81) = 0
'
'            pnRow = .Rows
'         End If
'      Case 81
'         If txtOthers(1) <> "" Then
'            If Not IsNumeric(txtOthers(Index)) Then txtOthers(Index).Text = txtOthers(Index).Tag
'
'            Debug.Print CLng(txtOthers(Index).Text); oTrans.Detail(.Row - 1, "nSerialCt")
'            If CLng(txtOthers(Index).Text) <> oTrans.Detail(.Row - 1, "nSerialCt") Then
'               If Not oTrans.LoadSerial(.Row - 1) Then
'                  txtOthers(Index).Text = txtOthers(Index).Tag
'                  Cancel = False
'               End If
'            End If
'
'            Select Case pnCntLvl
'            Case 1
'               oTrans.Detail(.Row - 1, "nActCtr01") = txtOthers(Index)
'               .TextMatrix(.Row, 10) = txtOthers(Index)
'            Case 2
'               oTrans.Detail(.Row - 1, "nActCtr02") = txtOthers(Index)
'               .TextMatrix(.Row, 11) = txtOthers(Index)
'            Case 3
'               oTrans.Detail(.Row - 1, "nActCtr03") = txtOthers(Index)
'               .TextMatrix(.Row, 12) = txtOthers(Index)
'            End Select
'
'            Call LoadSerial
'         End If
'      End Select
'   End With
'End Sub
