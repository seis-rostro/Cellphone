VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmCP_Price_Protection_Reg 
   BorderStyle     =   0  'None
   Caption         =   "Price Protection"
   ClientHeight    =   8160
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9435
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   121.973
   ScaleMode       =   0  'User
   ScaleWidth      =   81.057
   ShowInTaskbar   =   0   'False
   Tag             =   "ht0"
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4515
      Left            =   2790
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   3420
      Width           =   4950
      _ExtentX        =   8731
      _ExtentY        =   7964
      _Version        =   393216
      FocusRect       =   0
      SelectionMode   =   1
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
   Begin xrControl.xrFrame xrFrame2 
      Height          =   4815
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   3255
      Width           =   7770
      _ExtentX        =   13705
      _ExtentY        =   8493
      BackColor       =   14737632
      ClipControls    =   0   'False
      Begin VB.TextBox txtDetail 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Index           =   3
         Left            =   1305
         TabIndex        =   23
         Text            =   "0.00"
         Top             =   990
         Width           =   1260
      End
      Begin VB.TextBox txtDetail 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Index           =   2
         Left            =   1305
         TabIndex        =   21
         Text            =   "0.00"
         Top             =   600
         Width           =   1260
      End
      Begin VB.TextBox txtDetail 
         Appearance      =   0  'Flat
         Height          =   345
         Index           =   1
         Left            =   855
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   150
         Width           =   1710
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   1050
         X2              =   2235
         Y1              =   3795
         Y2              =   3795
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   1050
         X2              =   2235
         Y1              =   3765
         Y2              =   3765
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Index           =   15
         Left            =   1800
         TabIndex        =   35
         Tag             =   "ht0"
         Top             =   3360
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL AMT:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   13
         Left            =   195
         TabIndex        =   34
         Top             =   2715
         Width           =   1530
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   990
         X2              =   2175
         Y1              =   2175
         Y2              =   2175
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   990
         X2              =   2175
         Y1              =   2145
         Y2              =   2145
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   8
         Left            =   1785
         TabIndex        =   32
         Tag             =   "ht0"
         Top             =   1815
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DIFF:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   180
         TabIndex        =   31
         Top             =   1935
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "New SRP"
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
         Index           =   7
         Left            =   135
         TabIndex        =   22
         Top             =   1065
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Old SRP"
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
         Index           =   6
         Left            =   330
         TabIndex        =   20
         Top             =   690
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
         Height          =   255
         Index           =   5
         Left            =   330
         TabIndex        =   18
         Top             =   210
         Width           =   435
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   2160
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   1050
      Width           =   7770
      _ExtentX        =   13705
      _ExtentY        =   3810
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   7
         Left            =   6180
         TabIndex        =   17
         Top             =   1530
         Width           =   1455
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   6
         Left            =   6180
         TabIndex        =   15
         Top             =   1230
         Width           =   1455
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   600
         Index           =   3
         Left            =   870
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   420
         Width           =   4080
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   870
         TabIndex        =   5
         Top             =   120
         Width           =   4080
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   915
         Index           =   4
         Left            =   870
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   1035
         Width           =   4080
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   6180
         MaxLength       =   50
         TabIndex        =   11
         Top             =   630
         Width           =   1455
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   6180
         TabIndex        =   13
         Top             =   930
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Label2"
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
         Left            =   5190
         TabIndex        =   27
         Top             =   165
         Width           =   2385
      End
      Begin VB.Shape Shape4 
         Height          =   330
         Index           =   0
         Left            =   5145
         Top             =   135
         Width           =   2460
      End
      Begin VB.Shape Shape3 
         Height          =   390
         Index           =   0
         Left            =   5115
         Top             =   105
         Width           =   2520
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H8000000F&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   285
         Index           =   0
         Left            =   5250
         Tag             =   "et0;et0"
         Top             =   165
         Width           =   2265
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reference No."
         Height          =   195
         Index           =   10
         Left            =   5070
         TabIndex        =   16
         Top             =   1545
         Width           =   1050
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   195
         Index           =   11
         Left            =   150
         TabIndex        =   6
         Top             =   480
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Thru"
         Height          =   195
         Index           =   14
         Left            =   5085
         TabIndex        =   14
         Top             =   1245
         Width           =   720
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   195
         Index           =   12
         Left            =   105
         TabIndex        =   8
         Top             =   1050
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date From"
         Height          =   195
         Index           =   2
         Left            =   5070
         TabIndex        =   12
         Top             =   945
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transact Date"
         Height          =   195
         Index           =   1
         Left            =   5070
         TabIndex        =   10
         Top             =   660
         Width           =   1020
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company"
         Height          =   195
         Index           =   3
         Left            =   60
         TabIndex        =   4
         Top             =   165
         Width           =   660
      End
   End
   Begin xrControl.xrFrame xrFrame3 
      Height          =   495
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   7770
      _ExtentX        =   13705
      _ExtentY        =   873
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   945
         TabIndex        =   1
         Top             =   75
         Width           =   3990
      End
      Begin VB.TextBox txtSearch 
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
         Index           =   1
         Left            =   6090
         TabIndex        =   3
         Top             =   90
         Width           =   1545
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company"
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
         Left            =   75
         TabIndex        =   0
         Top             =   120
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Refer No."
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
         Left            =   5100
         TabIndex        =   2
         Top             =   105
         Width           =   840
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   8085
      TabIndex        =   25
      Top             =   3030
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
      Picture         =   "frmCP_Price_Protection_Reg.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   8085
      TabIndex        =   26
      Top             =   510
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
      Picture         =   "frmCP_Price_Protection_Reg.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   8085
      TabIndex        =   28
      Top             =   1140
      Width           =   1245
      _ExtentX        =   2196
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
      Picture         =   "frmCP_Price_Protection_Reg.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   8085
      TabIndex        =   29
      Top             =   510
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
      Picture         =   "frmCP_Price_Protection_Reg.frx":166E
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   8085
      TabIndex        =   30
      Top             =   1140
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
      Picture         =   "frmCP_Price_Protection_Reg.frx":1DE8
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   7
      Left            =   8085
      TabIndex        =   33
      Top             =   2415
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Export"
      AccessKey       =   "E"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCP_Price_Protection_Reg.frx":2562
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   8100
      TabIndex        =   36
      Top             =   1770
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
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
      Picture         =   "frmCP_Price_Protection_Reg.frx":2CDC
   End
End
Attribute VB_Name = "frmCP_Price_Protection_Reg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pxeMODULENAME = "frmCP_Price_ProtectionReg"
Private WithEvents oTrans As clsCPPriceProtection
Attribute oTrans.VB_VarHelpID = -1
Private oFormLoadSerial As frmLoadSerial

Private oSkin As clsFormSkin

Dim pbMasterGotFocus As Boolean
Dim pbGridGotFocus As Boolean
Dim pnIndex As Integer
Dim pnCtr As Integer
Dim p_oserial As Recordset

Dim pbSave As Boolean
Dim pnRow As Integer
Dim pbFormLoaded As Boolean
Dim pbLoadedRec As Boolean

Private Sub cmdButton_Click(Index As Integer)
   Dim lnRep As Integer
   
   Select Case Index
   Case 0 'Browse
      If oTrans.SearchTransaction = True Then
         LoadMaster
         LoadDetail
      End If
   Case 1 'Serial
      With oFormLoadSerial
         If oTrans.Master("sTransNox") <> "" Then
            .TransNox = oTrans.Master("sTransNox")
         End If
         .Show
      End With
   Case 2 'Close
      Unload Me
   Case 3 'Save
      If txtField(2).Text <> "" Then
         If oTrans.SaveTransaction() = True Then
            lnRep = MsgBox("Transaction Save Succesfully" & vbCrLf & _
                     "Do you want to Post this Transaction", vbYesNo + vbInformation, "Confirm")
            If lnRep = vbYes Then
               If oTrans.PostTransaction(oTrans.Master("sTransNox")) = True Then
                  If oTrans.OpenTransaction(oTrans.Master("sTransNox")) = True Then
                     MsgBox "Transaction Posted Successfully", vbInformation
                  Else
                     MsgBox "Unable to Post Transaction!" & vbCrLf & _
                        "Please contact GGC SEG/SSG for assitance!", vbCritical, "WARNING"
                  End If
               Else
                  MsgBox "Unable to Post Transaction!" & vbCrLf & _
                  "Please contact GGC SEG/SSG for assitance!", vbCritical, "WARNING"
               End If
            End If
            initButton xeModeReady
         End If
         lnRep = MsgBox("Do you want to Print this Transaction", vbYesNo + vbInformation, "Confirm")
         If lnRep = vbYes Then
            Call PrintTrans
         End If
      End If
   Case 4 'Cancel
        lnRep = MsgBox("Transaction is in Update Mode!!!" & vbCrLf & _
                     "Do you want to Cancel Transaction!!!", vbYesNo + vbQuestion, "Confirm")
      If lnRep = vbYes Then
         LoadDetail
         initButton xeModeReady
      End If
   Case 5 'update
      If txtField(2).Text <> "" Then
         If oTrans.UpdateTransaction() = True Then
            initButton xeStateUnknown
         Else
            MsgBox "Unable to update Transaction!!!", vbInformation
         End If
      Else
         MsgBox "No Record found!!!" & vbCrLf & _
                  "Please verify you entry then try again!!!", vbInformation, "Warning"
      End If
   Case 7
      If txtField(2).Text <> "" Then
         ExportPriceProtection oTrans
         MsgBox "Transaction Exported Successfully!", vbInformation
      Else
         MsgBox "Transaction not Found!", vbInformation
      End If
   End Select
End Sub

Private Sub InitGrid()
   Dim lnCtr As Integer

   With MSFlexGrid1
      .Clear
      .Cols = 4
      .Rows = 3

      .TextMatrix(0, 0) = ""
      .TextMatrix(0, 1) = "MODEL"
      .TextMatrix(0, 2) = "OLD SRP"
      .TextMatrix(0, 3) = "NEW SRP"
      
      .Row = 0
      'column alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = flexAlignLeftCenter
      Next

      .ColWidth(0) = "500"
      .ColWidth(1) = 2400
      .ColWidth(2) = 1000
      .ColWidth(3) = 1000

      .ColAlignment(1) = flexAlignLeftCenter
      .ColAlignment(2) = flexAlignLeftCenter
      .ColAlignment(3) = flexAlignLeftCenter
      
      .Row = 1
      .Col = 1
      .ColSel = .Cols - 1
   End With
   
End Sub

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0
   
   If Not pbFormLoaded Then
      pbFormLoaded = True
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyReturn, vbKeyUp, vbKeyDown
      Select Case KeyCode
      Case vbKeyReturn, vbKeyDown
         If GetFocus = MSFlexGrid1.hwnd Then Exit Sub
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

   Set oFormLoadSerial = New frmLoadSerial
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualRight

   Set oTrans = New clsCPPriceProtection
   Set oTrans.AppDriver = oApp
   oTrans.Branch = oApp.BranchCode
   oTrans.InitTransaction

   Call InitGrid
   Call initButton(xeModeReady)
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub MSFlexGrid1_Click()
   With MSFlexGrid1
      txtDetail(1) = .TextMatrix(.Row, 1)
      txtDetail(2) = .TextMatrix(.Row, 2)
      txtDetail(3) = .TextMatrix(.Row, 3)
      .Col = 1
      .ColSel = .Cols - 1
   End With
   
End Sub

Private Sub MSFlexGrid1_GotFocus()
   pbGridGotFocus = True
End Sub

Private Sub MSFlexGrid1_RowColChange()
   pnRow = MSFlexGrid1.Row
   MSFlexGrid1.Col = 1
   MSFlexGrid1.ColSel = MSFlexGrid1.Cols - 1
End Sub

Private Sub oTrans_DetailRetrieved(ByVal Index As Integer)
   oTrans.Master(Index) = txtField(Index)
End Sub

Private Sub LoadMaster()
   Dim Index As Integer
   
   For pnCtr = 0 To txtField.Count
      Select Case pnCtr
      Case 2
         txtField(pnCtr).Text = oTrans.Master(pnCtr)
         
         txtSearch(0).Text = txtField(pnCtr).Text
         txtSearch(0).Tag = txtSearch(0).Text
      Case 3, 4
         txtField(pnCtr).Text = oTrans.Master(pnCtr)
      Case 1, 5, 6
         txtField(pnCtr).Text = Format(oTrans.Master(pnCtr), "MMMM DD, YYYY")
      Case 7
         txtField(pnCtr).Text = oTrans.Master(pnCtr)
         
         txtSearch(1).Text = txtField(pnCtr).Text
         txtSearch(1).Tag = txtSearch(1).Text
      End Select
   Next

   Label2.Caption = Format(TransStat(oTrans.Master("cTranStat")), ">")
   pbLoadedRec = True
End Sub

Private Sub LoadDetail()
   With MSFlexGrid1
      .Rows = IIf(oTrans.ItemCount = 0, 2, oTrans.ItemCount + 1)
      For pnCtr = 0 To oTrans.ItemCount - 1
         .TextMatrix(pnCtr + 1, 1) = oTrans.Detail(pnCtr, "sModelNme")
         .TextMatrix(pnCtr + 1, 2) = Format(oTrans.Detail(pnCtr, "nAmountxx"), "#,##0.00")
         .TextMatrix(pnCtr + 1, 3) = Format(oTrans.Detail(pnCtr, "nActualxx"), "#,##0.00")
      Next
      
      .Row = 1
      .Col = 1
      .ColSel = .Cols - 1
      txtDetail(1) = .TextMatrix(.Row, 1)
      txtDetail(2) = .TextMatrix(.Row, 2)
      txtDetail(3) = .TextMatrix(.Row, 3)
      
      Label1(15).Caption = Format(oTrans.Master("nTranTotl"), "#,##0.00")
   End With
End Sub

Private Sub txtDetail_Change(Index As Integer)
   If Not pbFormLoaded Then Exit Sub

   With MSFlexGrid1
      Select Case Index
      Case 1
         .TextMatrix(pnRow, Index) = txtDetail(Index)
      Case 2, 3
         .TextMatrix(pnRow, Index) = txtDetail(Index)
      End Select
   End With
End Sub

Private Sub txtDetail_GotFocus(Index As Integer)
   With txtDetail(Index)
      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = oApp.getColor("HT1")
   End With

   pbMasterGotFocus = False
   pbGridGotFocus = False
   pnIndex = Index
End Sub

Private Sub txtDetail_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   With MSFlexGrid1
      Select Case Index
      Case 3
         If KeyCode = vbKeyReturn Then
               ComputeTotal
               Call txtDetail_Validate(Index, False)
            If .Rows - 1 > .Row Then
               txtDetail(2).SetFocus
               .Row = .Row + 1
            End If
         End If
         .Col = 1
         .ColSel = .Cols - 1
         txtDetail(1) = .TextMatrix(.Row, 1)
         txtDetail(2) = .TextMatrix(.Row, 2)
         txtDetail(3) = .TextMatrix(.Row, 3)
      End Select
   End With
End Sub

Private Sub txtDetail_LostFocus(Index As Integer)
   With txtDetail(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtDetail_Validate(Index As Integer, Cancel As Boolean)
   With txtDetail(Index)
      Select Case Index
      Case 2, 3
         If Not IsNumeric(.Text) Then .Text = 0#
         .Text = Format(.Text, "#,##0.00")
            oTrans.Detail(pnRow - 1, Index) = CDbl(.Text)
      End Select
   End With
   Call ComputeTotal
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      Select Case Index
      Case 1, 5, 6
         .Text = Format(.Text, "MM/DD/YYYY")
      End Select

      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = oApp.getColor("HT1")
   End With

   pnIndex = Index
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   With txtField(Index)
      Select Case Index
      Case 1, 5, 6
         If Not IsDate(.Text) Then .Text = oApp.ServerDate
         .Text = Format(.Text, "MMMM DD, YYYY")

         oTrans.Master(Index) = CDate(.Text)
      Case Else
         oTrans.Master(Index) = .Text
      End Select
   End With
End Sub

Private Sub initButton(lnStat As Integer)
   Dim lbShow As Boolean

   lbShow = IIf(lnStat = 0, False, True)
   cmdButton(3).Visible = lbShow
   cmdButton(4).Visible = lbShow

   cmdButton(0).Visible = Not lbShow
   cmdButton(2).Visible = Not lbShow
   cmdButton(5).Visible = Not lbShow
   cmdButton(7).Visible = Not lbShow

   xrFrame1.Enabled = lbShow
   xrFrame2.Enabled = lbShow
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

Private Sub ComputeTotal()
   Dim lnCtr As Integer
   Dim lnDiff As Currency
   Dim lnGrandTotl As Currency
   
   
   
   With MSFlexGrid1
      lnDiff = 0
      lnGrandTotl = 0
      For lnCtr = 0 To oTrans.ItemCount - 1
         lnDiff = oTrans.Detail(lnCtr, "nActualxx") - oTrans.Detail(lnCtr, "nAmountxx")
         Debug.Print lnDiff
         lnGrandTotl = lnGrandTotl + (lnDiff * oTrans.Detail(lnCtr, "nQuantity"))
      Next
'      Call LoadSerial
   End With
   
'   lnGrandTotl = Format(lnDiff, "#,##0.00")
   
   Label1(8).Caption = Format(lnDiff, "#,##0.00")
   Label1(15).Caption = Format(lnGrandTotl, "#,##0.00")
   
   oTrans.Master("nTranTotl") = lnGrandTotl
End Sub

Private Sub txtSearch_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

   With txtField(Index)
      Select Case Index
         Case 0
            If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
               oTrans.SearchTransaction txtSearch(0).Text, False
               LoadMaster
               LoadDetail
            End If
         Case 1
            If KeyCode = vbKeyF3 Then
               oTrans.SearchTransaction txtSearch(0).Text, False
               LoadMaster
               LoadDetail
            End If
      End Select
   End With
End Sub

Function ExportPriceProtection(oPO As clsCPPriceProtection) As Boolean
    Dim xl As New Excel.Application
    Dim xlsheet As Excel.Worksheet
    Dim xlwbook As Excel.Workbook

    Dim lsSQL As String
    Dim lors As Recordset
    Dim loSerial As Recordset
    
    Dim lnOldSRP As Currency
    Dim lnNewSRP As Currency

    Dim XTransNoX As String
    Dim XPromoDte As String
    Dim XSerialNo As String
    Dim XModelNme As String
    Dim XReferNox As String
    
    Dim xOldSrp As String
    Dim xNewSrp As String
    Dim xDiff As String
    Dim xNqty As String
    Dim xNttl As String
        
    Dim XLineNmbr As String
    Dim XInvTyp As String

    Dim XCompanyx As String
    Dim XAddressX As String
    
    Dim lnLineCtr As Integer
    Dim lnItemCtr As Integer
    
    XLineNmbr = ""
    XCompanyx = ""
    XAddressX = ""
    XSerialNo = ""
    xOldSrp = ""
    xNewSrp = ""
    xDiff = ""
    xNqty = ""
    xNttl = ""
    
    'Set the excel sheet format here according to the branch type
'    If InStr(1, oPO.Master("sTransNox")) <> "" Then
   If oPO.Master("sTransNox") <> "" Then
        XTransNoX = "A"
        XModelNme = "B"
        XSerialNo = "C"
        XCompanyx = "B"
        XPromoDte = "B"
        XReferNox = "B"
        
        xOldSrp = "B"
        xNewSrp = "B"
        xDiff = "B"
        xNqty = "B"
        xNttl = "B"
        
        
        XInvTyp = "C4"
        XLineNmbr = "B"
        lnLineCtr = 11
        
        Set xlwbook = xl.Workbooks.Open(oApp.AppPath & "\Reports\" & "Price Protection Format" & ".XLS")
        Set xlsheet = xlwbook.Sheets.Item(1)
         
         lsSQL = "SELECT" & _
                     " c.sSerialID" & _
                     ", c.sSerialNo" & _
                     ", a.sTransNox" & _
                     ", e.sModelNme" & _
                     ", f.nActualxx" & _
                     ", f.nAmountxx" & _
                     ", a.nTranTotl" & _
               " FROM CP_Price_Protection_Master a" & _
               ", CP_Price_Protection_Detail b" & _
               ", CP_Inventory_Serial c" & _
               ", CP_Inventory d" & _
               ", CP_Model e" & _
               ", CP_Price_Protection_Model f" & _
               " WHERE a.sTransNox = b.sTransNox" & _
               " AND b.sSerialID = c.sSerialID" & _
               " AND c.sStockIDx = d.sStockIDx" & _
               " AND d.sModelIDx = e.sModelIdx" & _
               " AND a.sTransNox = f.sTransNox" & _
               " AND d.sModelIDx = f.sModelIdx" & _
               " AND a.sTransNox = f.sTransNox" & _
               " AND a.sTransNox = " & strParm(oPO.Master("sTransNox"))
         Debug.Print lsSQL
         Set loSerial = New Recordset
         loSerial.Open lsSQL, oApp.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
         
         lnNewSRP = 0#
         lnOldSRP = 0#

         loSerial.MoveFirst
         For lnItemCtr = 0 To loSerial.RecordCount - 1
            If XTransNoX <> "" Then xlsheet.Range(XTransNoX & (lnLineCtr + lnItemCtr)).Value = loSerial("sTransNox")
            If XModelNme <> "" Then xlsheet.Range(XModelNme & (lnLineCtr + lnItemCtr)).Value = loSerial("sModelNme")
            If XSerialNo <> "" Then xlsheet.Range(XSerialNo & (lnLineCtr + lnItemCtr)).Value = loSerial("sSerialNo")

            lnNewSRP = lnNewSRP + loSerial("nAmountxx")
            lnOldSRP = lnOldSRP + loSerial("nActualxx")

            loSerial.MoveNext
            
            If XCompanyx <> "" Then xlsheet.Range(XCompanyx & (3)).Value = oPO.Master("sCompnyNm")
            If XPromoDte <> "" Then xlsheet.Range(XLineNmbr & (4)).Value = Format(oPO.Master("dPromoFrm"), "MMMM DD, YYYY") & "-" & Format(oPO.Master("dPromoFrm"), "MMMM DD, YYYY")
            If XReferNox <> "" Then xlsheet.Range(XLineNmbr & (5)).Value = oPO.Master("sReferNox")
         Next

'         If xOldSrp <> "" Then xlsheet.Range(xOldSrp & (7)).Value = lnOldSRP 'oPO.Detail(0, 2)
'         If xNewSrp <> "" Then xlsheet.Range(xNewSrp & (8)).Value = lnNewSRP 'oPO.Detail(0, 3)
         If xDiff <> "" Then xlsheet.Range(xDiff & (7)).Value = lnNewSRP - lnOldSRP 'oPO.Detail(0, 3) - oPO.Detail(0, 2)
         If xNqty <> "" Then xlsheet.Range(xNqty & (8)).Value = loSerial.RecordCount
         
         xlwbook.SaveAs "C:\GGC_Systems\Temp\" & "Price Protection Format" & "(" & Right(oPO.Master("sTransNox"), 10) & ").XLS"
         xl.ActiveWorkbook.Close  'False, oApp.AppPath & "\Temp\" & sFileName & "(" & Right(oPO.Master("sTransNox"), 5) & ").XLS"
         xl.Quit
         
         Set xlwbook = Nothing
         Set xl = Nothing
             
         ExportPriceProtection = True
         Exit Function
    Else
        ExportPriceProtection = False
        MsgBox "Not Found!"
        Exit Function
    End If
    
    Set xlwbook = xl.Workbooks.Open(oApp.AppPath & "\Reports\" & "Price Protection Format" & ".XLS")
    Set xlsheet = xlwbook.Sheets.Item(1)

    Set xlwbook = Nothing
    Set xl = Nothing
End Function

Private Sub LoadSerial()
   Dim lsSQL As String
   
   lsSQL = "SELECT" & _
               " a.sSerialID" & _
               ", a.sSerialNo" & _
               ", b.sTransNox" & _
               ", d.sModelNme" & _
               " FROM CP_Inventory_Serial a" & _
               ", CP_Price_Protection_Detail b" & _
               ", CP_Price_Protection_Model c" & _
               ", CP_Model d" & _
               " WHERE a.sSerialID = b.sSerialId" & _
               " AND b.sTransNox = c.sTransNox" & _
               " AND c.sModelIDx = d.sModelIDx" & _
               " AND b.sTransNox = " & strParm(oTrans.Master("sTransNox"))
         Debug.Print lsSQL
         Set p_oserial = New Recordset
         p_oserial.Open lsSQL, oApp.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
End Sub

Private Function PrintTrans() As Boolean
   Dim lrs As ADODB.Recordset
   Dim lors As ADODB.Recordset
   Dim lsOldProc As String

   lsOldProc = "PrinTrans"
   'On Error GoTo errProc

   PrintTrans = True
   
   Set lrs = New ADODB.Recordset

   lrs.Fields.Append "nField01", adInteger, 5
   lrs.Fields.Append "sField01", adVarChar, 120
   lrs.Fields.Append "sField02", adVarChar, 100
   lrs.Fields.Append "sField03", adVarChar, 100
   lrs.Fields.Append "sField04", adVarChar, 200
   lrs.Fields.Append "sField05", adVarChar, 100
   lrs.Fields.Append "sField06", adVarChar, 100
   lrs.Fields.Append "sField07", adVarChar, 200
   lrs.Fields.Append "sField09", adVarChar, 200
   lrs.Fields.Append "sField10", adVarChar, 200
   lrs.Fields.Append "lField01", adCurrency

   lrs.Open

   With oTrans
      lrs.AddNew
      lrs("sField01").Value = Format(.Master("dTransact"), "MMMM DD, YYYY")
      lrs("sField02").Value = Format(.Master("dPromoFrm"), "YYYY-MM-DD") & "-" & Format(.Master("dPromoTru"), "YYYY-MM-DD")
      lrs("sField03").Value = .Master("sReferNox")
      lrs("sField04").Value = .Master("sRemarksx") '"PRICE PROTECTION"
      lrs("sField05").Value = .Master(2)
      lrs("sField06").Value = "Main Office"
      lrs("sField07").Value = "" '.Master("sRemarksx")
      lrs("sField09").Value = 0#
      lrs("lField01").Value = .Master("nTranTotl")
      
   End With

   'assign important info to the report
   Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\ClaimForm.rpt")
   oReport.DiscardSavedData
   oReport.FieldMappingType = crAutoFieldMapping
   oReport.Database.SetDataSource lrs
   
   PrintTrans = True
   oReport.PrintOutEx False, 1

endPoc:
   Set oReport = Nothing
   Set lrs = Nothing
   Set lors = Nothing
   Exit Function
errProc:
   PrintTrans = False
   ShowError lsOldProc & "( " & " )"
End Function

