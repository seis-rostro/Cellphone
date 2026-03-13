VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmJobOrder 
   BorderStyle     =   0  'None
   Caption         =   "Job Order"
   ClientHeight    =   5010
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11070
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5010
   ScaleWidth      =   11070
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame3 
      Height          =   1470
      Left            =   1575
      Tag             =   "wt0;fb0"
      Top             =   3435
      Width           =   9390
      _ExtentX        =   16563
      _ExtentY        =   2593
      BackColor       =   7716603
      ClipControls    =   0   'False
      BorderStyle     =   1
      Begin xrControl.xrFrame xrFrame1 
         Height          =   1215
         Index           =   1
         Left            =   90
         Tag             =   "wt0;fb0"
         Top             =   105
         Width           =   9180
         _ExtentX        =   16193
         _ExtentY        =   2143
         BackColor       =   12648447
         BorderStyle     =   3
         Begin VB.TextBox txtfield 
            Appearance      =   0  'Flat
            Height          =   975
            Index           =   10
            Left            =   975
            MultiLine       =   -1  'True
            TabIndex        =   27
            Text            =   "frmJO.frx":0000
            Top             =   105
            Width           =   4005
         End
         Begin VB.OptionButton Option1 
            Height          =   195
            Index           =   0
            Left            =   5190
            TabIndex        =   28
            Tag             =   "et0;fb0"
            Top             =   105
            Width           =   210
         End
         Begin VB.OptionButton Option1 
            Height          =   195
            Index           =   1
            Left            =   5190
            TabIndex        =   30
            Tag             =   "et0;fb0"
            Top             =   345
            Width           =   225
         End
         Begin VB.OptionButton Option1 
            Height          =   195
            Index           =   2
            Left            =   5190
            TabIndex        =   32
            Tag             =   "et0;fb0"
            Top             =   585
            Width           =   210
         End
         Begin VB.OptionButton Option1 
            Height          =   195
            Index           =   3
            Left            =   7770
            TabIndex        =   34
            Tag             =   "et0;fb0"
            Top             =   105
            Width           =   195
         End
         Begin VB.OptionButton Option1 
            Height          =   195
            Index           =   4
            Left            =   7770
            TabIndex        =   36
            Tag             =   "et0;fb0"
            Top             =   345
            Width           =   210
         End
         Begin VB.OptionButton Option1 
            Height          =   195
            Index           =   5
            Left            =   7770
            TabIndex        =   38
            Tag             =   "et0;fb0"
            Top             =   585
            Width           =   210
         End
         Begin VB.TextBox txtfield 
            Appearance      =   0  'Flat
            Height          =   240
            Index           =   11
            Left            =   5205
            TabIndex        =   40
            Text            =   "Text1"
            Top             =   840
            Width           =   3870
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Complaint"
            Height          =   195
            Index           =   7
            Left            =   105
            TabIndex        =   25
            Top             =   150
            Width           =   690
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "New Unit / Complete Package"
            Height          =   195
            Index           =   3
            Left            =   5460
            TabIndex        =   29
            Top             =   105
            Width           =   2295
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Battery"
            Height          =   195
            Index           =   9
            Left            =   5460
            TabIndex        =   31
            Top             =   345
            Width           =   495
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sim Card"
            Height          =   195
            Index           =   10
            Left            =   5460
            TabIndex        =   33
            Top             =   585
            Width           =   630
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Back Cover"
            Height          =   255
            Left            =   8040
            TabIndex        =   35
            Top             =   105
            Width           =   900
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Charger"
            Height          =   210
            Left            =   8040
            TabIndex        =   37
            Top             =   345
            Width           =   945
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Download"
            Height          =   180
            Left            =   8040
            TabIndex        =   39
            Top             =   585
            Width           =   780
         End
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   7
      Left            =   90
      TabIndex        =   47
      Top             =   3015
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
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
      Picture         =   "frmJO.frx":0006
      CaptionAlign    =   0
      BackColor       =   14286077
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   405
      Index           =   3
      Left            =   90
      TabIndex        =   45
      Top             =   2595
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
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
      Picture         =   "frmJO.frx":0780
      CaptionAlign    =   0
      BackColor       =   14286077
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   1
      Left            =   90
      TabIndex        =   41
      Top             =   1335
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
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
      Picture         =   "frmJO.frx":0EFA
      CaptionAlign    =   0
      BackColor       =   14286077
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   5
      Left            =   90
      TabIndex        =   46
      Top             =   2595
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
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
      Picture         =   "frmJO.frx":1674
      CaptionAlign    =   0
      BackColor       =   14286077
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   6
      Left            =   90
      TabIndex        =   48
      Top             =   3015
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
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
      Picture         =   "frmJO.frx":1DEE
      CaptionAlign    =   0
      BackColor       =   14286077
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   0
      Left            =   90
      TabIndex        =   42
      Top             =   1335
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
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
      Picture         =   "frmJO.frx":2568
      CaptionAlign    =   0
      BackColor       =   14286077
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   2
      Left            =   90
      TabIndex        =   43
      Top             =   1755
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
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
      Picture         =   "frmJO.frx":2CE2
      CaptionAlign    =   0
      BackColor       =   14286077
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   4
      Left            =   90
      TabIndex        =   44
      Top             =   2175
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
      Caption         =   "Pre&view"
      AccessKey       =   "v"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmJO.frx":345C
      CaptionAlign    =   0
      BackColor       =   14286077
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1290
      Index           =   0
      Left            =   1575
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   9390
      _ExtentX        =   16563
      _ExtentY        =   2275
      BackColor       =   7716603
      BorderStyle     =   1
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   5
         Left            =   7095
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   885
         Width           =   2175
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   2
         Left            =   4050
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   630
         Width           =   1860
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   1350
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   630
         Width           =   1935
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   0
         Left            =   1350
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   135
         Width           =   1920
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   3
         Left            =   1350
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   885
         Width           =   4560
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   4
         Left            =   7095
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   630
         Width           =   2175
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Received"
         Height          =   195
         Index           =   17
         Left            =   5970
         TabIndex        =   10
         Top             =   915
         Width           =   1080
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IMEI No."
         Height          =   195
         Index           =   0
         Left            =   3345
         TabIndex        =   4
         Top             =   660
         Width           =   675
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   240
         Left            =   1395
         Tag             =   "et0;ht2"
         Top             =   180
         Width           =   1935
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction No."
         Height          =   195
         Index           =   13
         Left            =   150
         TabIndex        =   0
         Top             =   165
         Width           =   1140
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Job Order No."
         Height          =   195
         Index           =   12
         Left            =   150
         TabIndex        =   2
         Top             =   660
         Width           =   990
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer "
         Height          =   195
         Index           =   8
         Left            =   150
         TabIndex        =   6
         Top             =   915
         Width           =   705
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contact No."
         Height          =   195
         Index           =   5
         Left            =   5970
         TabIndex        =   8
         Top             =   660
         Width           =   855
      End
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   1530
      Left            =   1575
      Tag             =   "wt0;fb0"
      Top             =   1890
      Width           =   9390
      _ExtentX        =   16563
      _ExtentY        =   2699
      BackColor       =   7716603
      BorderStyle     =   1
      Begin VB.TextBox txtothers 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   0
         Left            =   1350
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   1050
         Width           =   5325
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   6
         Left            =   1350
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   540
         Width           =   2340
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   8
         Left            =   4395
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   795
         Width           =   2295
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   7
         Left            =   1350
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   795
         Width           =   2340
      End
      Begin VB.CheckBox chkField 
         BackColor       =   &H0075BEFB&
         Caption         =   "Back Job "
         Height          =   225
         Index           =   3
         Left            =   6990
         TabIndex        =   24
         Tag             =   "wt0;fb0"
         Top             =   1080
         Width           =   975
      End
      Begin VB.CheckBox chkField 
         BackColor       =   &H0075BEFB&
         Caption         =   "Under Limited Warranty"
         Height          =   255
         Index           =   2
         Left            =   6990
         TabIndex        =   23
         Tag             =   "wt0;fb0"
         Top             =   810
         Width           =   2055
      End
      Begin VB.CheckBox chkField 
         BackColor       =   &H0075BEFB&
         Caption         =   "Void Warranty"
         Height          =   255
         Index           =   1
         Left            =   6990
         TabIndex        =   22
         Tag             =   "wt0;fb0"
         Top             =   555
         Width           =   1440
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         BackColor       =   &H0075BEFB&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   9
         Left            =   8025
         TabIndex        =   26
         TabStop         =   0   'False
         Tag             =   "et0;fb0"
         Text            =   "text1"
         Top             =   1080
         Width           =   1140
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier"
         Height          =   195
         Index           =   23
         Left            =   210
         TabIndex        =   19
         Top             =   1080
         Width           =   570
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Purchase Date"
         Height          =   195
         Index           =   15
         Left            =   210
         TabIndex        =   13
         Top             =   570
         Width           =   1065
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
         Height          =   195
         Index           =   1
         Left            =   3885
         TabIndex        =   17
         Top             =   825
         Width           =   435
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   225
         X2              =   6690
         Y1              =   435
         Y2              =   435
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Other Info"
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
         Index           =   2
         Left            =   210
         TabIndex        =   12
         Top             =   210
         Width           =   870
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   6990
         X2              =   9120
         Y1              =   435
         Y2              =   435
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Warranty Info"
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
         Index           =   16
         Left            =   6990
         TabIndex        =   21
         Top             =   210
         Width           =   1170
      End
      Begin VB.Shape Shape4 
         Height          =   1305
         Left            =   6855
         Top             =   90
         Width           =   2415
      End
      Begin VB.Shape Shape3 
         Height          =   1275
         Left            =   -1545
         Top             =   -3105
         Width           =   6705
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Brand"
         Height          =   195
         Index           =   6
         Left            =   210
         TabIndex        =   15
         Top             =   825
         Width           =   420
      End
      Begin VB.Shape Shape6 
         Height          =   1305
         Left            =   90
         Top             =   90
         Width           =   6705
      End
   End
   Begin MSComCtl2.Animation Progress 
      Height          =   720
      Left            =   300
      TabIndex        =   49
      TabStop         =   0   'False
      Tag             =   "wt0;wb0"
      Top             =   480
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   1270
      _Version        =   393216
      BackColor       =   1842204
      FullWidth       =   61
      FullHeight      =   48
   End
End
Attribute VB_Name = "frmJobOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents oDriver As FormDriver
Attribute oDriver.VB_VarHelpID = -1
Private oSkin As FormSkin
Private bLoaded As Boolean
Private oRS As New ADODB.Recordset

Dim txtfieldGotfocus As Boolean
Dim txtOthersGotfocus As Boolean
Dim pbnewitem As Boolean

Dim psSelected() As String

Dim pnindex As Integer
Dim Index As Integer
Dim pnCtr As Integer
Dim lnCtr As Integer

Private Sub cmdButton_Click(Index As Integer)
Dim lsSearch As String
Dim lsRep As Integer
Dim orig As String
Dim lsSQL As String
Dim lsCondition As String
   
   Select Case Index
      Case 0 'save
             oDriver.RecordSave
     Case 1 'New
            oDriver.RecordNew
      Case 2 'Browse
            oDriver.BrowseRecord
      Case 3 'search
            If txtfieldGotfocus Then
               If pnindex = 2 Then SearchSerial
               If pnindex = 7 Or pnindex = 8 Then oDriver.RecordSearch txtField(pnindex).Text
               If pnindex = 11 Then SearchBackJob False
               If pnindex = 3 Then Search_Client False
            End If
      Case 4 'Print
            If pbnewitem = False Then Print_JobOrder
      Case 5 'Update
            oDriver.RecordUpdate
      Case 6 'cancel
            UnLockFields
            oDriver.RecordCancelUpdate
      Case 7 'Close
            Unload Me
      End Select
End Sub

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0
   
   If bLoaded = False Then
      oDriver.RecordNew
      bLoaded = True
      oDriver.DisableTextbox 0
   End If

End Sub

Private Sub Form_Deactivate()
   Progress.Stop
   Progress.Close
End Sub

Private Sub Form_Load()
Dim lsSQL As String

   CenterChildForm mdiMain, Me

   bLoaded = False
   
   Set oDriver = New FormDriver
   Set oDriver.AppDriver = oApp
   Set oDriver.MainForm = Me
   
   Set oSkin = New FormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransaction

    oDriver.RecQuery = "SELECT" _
                  & " sTransNox, " _
                  & " sJobOrdNo, " _
                  & " sIMEINoxx, " _
                  & " sClientID, " _
                  & " sTelNoxxx, " _
                  & " dTransact, " _
                  & " dPurchase, " _
                  & " sBrandIDx, " _
                  & " sModelIDx, " _
                  & " sBckJobNo, " _
                  & " sComplent, " _
                  & " sCategory, " _
                  & " sSupplier, " _
                  & " cWarranty, " _
                  & " cCategory, " _
                  & " cTranStat, " _
   
   oDriver.RecQuery = oDriver.RecQuery _
                  & " nLaborTot, " _
                  & " nPartsTot, " _
                  & " nMiscChrg, " _
                  & " nTranTotl, " _
                  & " nAmtPaidx, " _
                  & " dPaymentx, " _
                  & " sPaymRecv, " _
                  & " sReferNox, " _
                  & " sModified, " _
                  & " dModified, " _
                  & " vTimeStmp  " _
                  & " FROM CP_JobOrder_Master " _

   oDriver.BrowseQuery = "SELECT" _
                  & " a.sJobOrdNo, " _
                  & " d.sLastName + ', ' + d.sFrstName + ' ' + d.sMiddName as xFullName, " _
                  & " b.sBrandNme+' '+ c.sModelNme as BrandModel, " _
                  & " a.dTransact, " _
                  & " a.nTranTotl  " _
               & " FROM CP_JobOrder_Master a " _
                  & " LEFT JOIN Brand b " _
                     & " ON a.sBrandIDx = b.sBrandIDx " _
                  & " LEFT JOIN Model c " _
                     & " ON a.sModelIDx = c.sModelIDx " _
                  & " LEFT JOIN Client_Master d " _
                     & " ON a.sClientID = d.sClientID " _
               & " WHERE cTranStat = 0 " _
                  & " AND sReferNox = '' "

   oDriver.InitRecForm

   oDriver.BrowseColumn(0) = "sJobOrdNo"
   oDriver.BrowseColumn(1) = "xFullName"
   oDriver.BrowseColumn(2) = "BrandModel"
   oDriver.BrowseColumn(3) = "dTransact"
   oDriver.BrowseColumn(4) = "nTranTotl"
   
   oDriver.BrowseFFormat(3) = "MMMM dd, yyyy"
   oDriver.BrowseFFormat(4) = "#,##0.00"

   oDriver.BrowseFTitle(0) = "J.O. No."
   oDriver.BrowseFTitle(1) = "Customer Name"
   oDriver.BrowseFTitle(2) = "Brand & Model"
   oDriver.BrowseFTitle(3) = "Date"
   oDriver.BrowseFTitle(4) = "Tran. Total"
    
   'Brand
   oDriver.LookupQuery(7) = "SELECT" _
                     & " sBrandIDx, " _
                     & " sBrandNme " _
                  & " FROM Brand " _
                  & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
                  & " ORDER BY sBrandNme"
   
   oDriver.LookupReference(7) = "sBrandIDx»sBrandNme"
   oDriver.LookupColumn(7) = "sBrandNme"
   oDriver.LookupTitle(7) = "Brand Name"

   'Model
   oDriver.LookupQuery(8) = "SELECT" _
                     & " a.sModelIDx, " _
                     & " a.sModelNme, " _
                     & " b.sBrandNme  " _
                  & "FROM Model a LEFT JOIN " _
                     & " Brand b " _
                        & " ON a.sBrandIDx = b.sBrandIDx " _
                  & "WHERE a.cRecdStat = 1 " _
                  & "ORDER BY a.sModelNme "
   
   oDriver.LookupReference(8) = "sModelIDx»sModelNme»sBrandNme"
   oDriver.LookupColumn(8) = "sModelNme»sBrandNme"
   oDriver.LookupTitle(8) = "Model»Brand"
   
   'Customer
   oDriver.LookupQuery(3) = "SELECT" _
                  & " a.sClientID, " _
                  & " a.sLastName + ', ' + a.sFrstName + ' ' + a.sMiddName as xFullName, " _
                  & " a.sAddressx + ', ' + b.sTownName as xAddressx " _
               & " FROM Client_Master a " _
                  & " LEFT JOIN TownCity b " _
                     & " ON a.sTownIDxx = b.sTownIDxx " _
               & " ORDER BY slastname, sfrstname, smiddname "

   oDriver.LookupReference(3) = "a.sClientID»" _
                              & "a.sLastName + ', ' + a.sFrstName + ' ' + a.sMiddName»" _
                              & "a.sAddressx + ', ' + b.sTownName"
   oDriver.LookupColumn(3) = "xFullName»xAddressx"
   oDriver.LookupTitle(3) = "Customer Name»Address"
                    
   oDriver.FieldFormat(0) = "@@-@@@@@@@@"
   oDriver.FieldSize(0) = Len(oDriver.FieldFormat(0))
   
   oDriver.FieldFormat(5) = "MMMM DD, YYYY"
   oDriver.FieldFormat(6) = "MMMM DD, YYYY"
   oDriver.FieldStart = 1
   
End Sub

Private Sub oDriver_EnableOtherControl()
   oDriver.DisableTextbox 0
End Sub

Private Sub oDriver_InitValue()
Dim lsSQL As String
Dim lsCondition As String

   oDriver.FieldReference(0) = True
   If Not oDriver.SetValue(0, getNextCode("CP_JobOrder_Master", "sTransNox", True, oApp.Connection, True, oApp.BranchCode)) Then Exit Sub

   pbnewitem = True
   
   txtField(1).Text = getNextCode("CP_JobOrder_Master", "sJobOrdNo", True, oApp.Connection, True, oApp.BranchCode)
   txtField(5).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
   txtothers(0).Text = ""
   txtothers(0).Tag = ""
   oDriver.FieldValue(15) = 0
   
   'Category
   Option1(0).Value = True
   For lnCtr = 1 To 5
      Option1(lnCtr).Value = False
   Next
   
   'Warranty
   For lnCtr = 1 To 3
      chkField(lnCtr).Value = 0
   Next
      
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oDriver = Nothing
   Set oSkin = Nothing
End Sub

Private Sub SearchSerial()
Dim lsSQL As String
Dim lsCondition As String
Dim lsSearch As String
   
   If oRS.State = adStateOpen Then oRS.Close
         lsSQL = "SELECT" _
                  & " b.sSerialID, " _
                  & " b.sBranchCd, " _
                  & " b.sIMEINoxx, " _
                  & " f.sLastName + ', '+ f.sFrstName+' '+ f.sMiddname as xFullName, " _
                  & " f.sPhoneNox as xContactx, " _
                  & " b.sStockIDx, " _
                  & " a.dModified, " _
                  & " d.sBrandNme, " _
                  & " e.sModelNme, " _
                  & " b.sClientID, " _
                  & " d.sBrandIDx, " _
                  & " e.sModelIDx, " _
                  & " c.sSupplier  " _

         lsSQL = lsSQL _
                  & " FROM CP_Serial_Master b " _
                     & " LEFT JOIN CP_Inventory c " _
                        & " ON b.sStockIDx = c.sStockIDx " _
                     & " LEFT JOIN Brand d " _
                        & " ON c.sBrandIDx = d.sBrandIDx " _
                     & " LEFT JOIN Model e " _
                        & " ON c.sModelIDx = e.sModelIDx " _
                     & " LEFT JOIN Client_Master f " _
                        & " ON b.sClientID = f.sClientID " _
                     & " LEFT JOIN CP_SO_Serial a " _
                        & " ON b.sSerialID = a.sSerialID " _
                  & " WHERE b.sIMEINoxx like '%" & txtField(2).Text & "%' " _
                        & " AND cSoldStat = 1 "
   oRS.Open lsSQL, oApp.Connection, adOpenKeyset, adLockOptimistic, adCmdText
         
   Select Case oRS.RecordCount
      Case 1
         For lnCtr = 2 To 8
            Select Case lnCtr
               Case 2 To 4
                  txtField(lnCtr).Text = IIf(Not IsNull(oRS(lnCtr)), oRS(lnCtr), "")
               Case 6
                  txtField(lnCtr).Text = IIf(Not IsNull(oRS(lnCtr)), Format(oRS(lnCtr), "MMMM dd, yyyy"), "")
               Case 7, 8
                  txtField(lnCtr).Text = IIf(Not IsNull(oRS(lnCtr)), oRS(lnCtr), "")
            End Select
         Next
         oDriver.FieldValue(3) = oRS(9)
         oDriver.FieldValue(7) = oRS(10)
         oDriver.FieldValue(8) = oRS(11)
         oDriver.FieldValue(12) = IIf(Not IsNull(oRS(12)), oRS(12), "")
         txtothers(0).Tag = IIf(Not IsNull(oRS(12)), oRS(12), "")
      Case Is > 1
         lsSearch = KwikSearch(oApp, lsSQL, _
                    "sIMEINoxx»xFullName»sBrandNme»sModelNme»dModified", _
                    "IMEI No.»Customer Name»Brand»Model»Purchase Date", _
                    "@»@»@»@»MMMM dd, yyyy")

         If lsSearch <> "" Then
            psSelected = Split(lsSearch, "»")
            For lnCtr = 2 To 8
               Select Case lnCtr
                  Case 2 To 4
                     txtField(lnCtr).Text = IIf(Not IsNull(psSelected(lnCtr)), psSelected(lnCtr), "")
                  Case 6
                     txtField(lnCtr).Text = IIf(Not IsNull(psSelected(lnCtr)), Format(psSelected(lnCtr), "MMMM dd, yyyy"), "")
                  Case 7, 8
                     txtField(lnCtr).Text = IIf(Not IsNull(psSelected(lnCtr)), psSelected(lnCtr), "")
               End Select
            Next
            oDriver.FieldValue(3) = psSelected(9)
            oDriver.FieldValue(7) = psSelected(10)
            oDriver.FieldValue(8) = psSelected(11)
            oDriver.FieldValue(12) = IIf(Not IsNull(psSelected(12)), psSelected(12), "")
            txtothers(0).Tag = IIf(Not IsNull(psSelected(12)), psSelected(12), "")
         End If
   End Select
   SearchSupplier
   Set oRS = Nothing
End Sub

Private Sub SearchSupplier()
Dim lsSQL As String
   'SearchSupplier
   Set oRS = New ADODB.Recordset
   lsSQL = "SELECT " _
            & " sSupplyID, " _
            & " sSupplyNm  " _
    & " FROM Supplier " _
    & " WHERE sSupplyID = '" & oDriver.FieldValue(12) & "' "
   If oRS.State = adStateOpen Then oRS.Close
   oRS.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   If oRS.RecordCount = 1 Then
      txtothers(0).Text = IIf(IsNull(oRS("sSupplyNm")), "", oRS("sSupplyNm"))
   End If
   Set oRS = Nothing
End Sub
Private Sub SearchBackJob(ByVal SearchValue As Boolean)
Dim lsSearch As String
Dim lsSQL As String
Dim Index As Integer
Dim lnCtr As Integer

   Set oRS = New ADODB.Recordset
         lsSQL = "SELECT" _
               & " a.sTransNOx, " _
               & " a.sJobOrdNo, " _
               & " a.sIMEINoxx, " _
               & " d.sLastName + ', ' + d.sFrstName + ' ' + d.sMiddName xFullName, " _
               & " a.sTelNoxxx, " _
               & " a.dTransact, " _
               & " a.dPurchase, " _
               & " b.sBrandNme, " _
               & " c.sModelNme, " _
               & " a.sBckJobNo, " _
               & " a.sComplent, " _
               & " a.sCategory, " _
               & " a.sClientID, " _
               & " a.sBrandIDx, " _
               & " a.sModelIDx, " _
               & " a.cWarranty, " _
               & " a.cCategory, " _
               & " a.sSupplier, " _
               & " a.nTranTotl  " _

         lsSQL = lsSQL _
            & " FROM CP_JobOrder_Master a " _
               & " LEFT JOIN Brand b " _
                  & " ON a.sBrandIDx = b.sBrandIDx " _
               & " LEFT JOIN Model c " _
                  & " ON a.sModelIDx = c.sModelIDx " _
               & " LEFT JOIN Client_Master d " _
                  & " ON a.sClientID = d.sClientID " _
            & " WHERE a.cTranstat = 1 " _

         
   If SearchValue Then
      lsSQL = lsSQL & " AND a.sJobOrdNo = '" & txtField(9).Text & "' "

   Else
      lsSQL = lsSQL & " AND a.sJobOrdNo like '" & txtField(9).Text & "%' "
   End If
   
   lsSQL = lsSQL & " ORDER BY xFullName "
   oRS.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
            
   Select Case oRS.RecordCount
   Case 0
      MsgBox "No Record Found!!!", vbInformation, "Information"
   Case 1
         For Index = 2 To 17
            Select Case Index
              Case 2, 4, 10 To 11
                  txtField(Index).Text = oRS(Index)
                  oDriver.FieldValue(Index) = oRS(Index)
              Case 5, 6
                  txtField(Index).Text = Format(oRS(Index), "MMMM dd, yyyy")
                  oDriver.FieldValue(Index) = oRS(Index)
              Case 3
                  txtField(Index).Text = oRS(Index)
                  oDriver.FieldValue(Index) = oRS(12)
              Case 7, 8
                  txtField(Index).Text = oRS(Index)
                  oDriver.FieldValue(Index) = oRS(Index + 6)
              Case 9
                  txtField(Index).Text = oRS(1)
                  oDriver.FieldValue(Index) = oRS(1)
              Case 16 'Category
                  lnCtr = oRS(Index)
                  Option1(lnCtr).Value = True
              Case 17 'Supplier
                  oDriver.FieldValue(12) = oRS(Index)
           End Select
         Next
   Case Is > 1
      lsSearch = KwikBrowse(oApp, oRS, _
                        "xFullName»sJobOrdNo»sBrandNme»sModelNme»dTransact»nTranTotl", _
                        "Customer Name»J.O. No.»Brand»Model»Tran. Date»Tran. Total", _
                        "@»@»@»@»MMM dd, yyyy»#,##0.00")

      If lsSearch <> "" Then
         psSelected = Split(lsSearch, "»")
         For Index = 2 To 17
            Select Case Index
              Case 2, 4, 10 To 11
                  txtField(Index).Text = psSelected(Index)
                  oDriver.FieldValue(Index) = psSelected(Index)
              Case 5, 6
                  txtField(Index).Text = Format(psSelected(Index), "MMMM dd, yyyy")
                  oDriver.FieldValue(Index) = psSelected(Index)
              Case 3
                  txtField(Index).Text = psSelected(Index)
                  oDriver.FieldValue(Index) = psSelected(12)
              Case 7, 8
                  txtField(Index).Text = psSelected(Index)
                  oDriver.FieldValue(Index) = psSelected(Index + 6)
              Case 9
                  txtField(Index).Text = psSelected(1)
                  oDriver.FieldValue(Index) = oRS(1)
              Case 16 'Category
                  lnCtr = psSelected(Index)
                  Option1(lnCtr).Value = True
              Case 17 'Supplier
                  oDriver.FieldValue(12) = psSelected(Index)
           End Select
         Next
      End If
      SearchSupplier
      LockFields
   End Select
   Set oRS = Nothing
   
End Sub

Private Sub LockFields()
Dim Index As Integer
   If xrFrame2.Enabled = True Then xrFrame2.Enabled = False
   If xrFrame1(0).Enabled = True Then xrFrame1(0).Enabled = False
   If txtothers(0).Enabled = True Then txtothers(0).Enabled = False
   
End Sub
Private Sub UnLockFields()
   If xrFrame2.Enabled = False Then xrFrame2.Enabled = True
   If xrFrame1(0).Enabled = False Then xrFrame1(0).Enabled = True
   If txtothers(0).Enabled = False Then txtothers(0).Enabled = True
End Sub

Private Sub Search_Client(ByVal SearchValue As Boolean)
Dim lsSearch As String
Dim lrs As ADODB.Recordset
Dim lsSQL As String

   Set lrs = New ADODB.Recordset
   lsSQL = "SELECT" _
               & " a.sClientID, " _
               & " a.sLastName + ', ' + a.sFrstName + ' ' + a.sMiddName xFullName, " _
               & " a.sAddressx + ', ' + b.sTownName as xAddressx, " _
               & " a.sPhoneNox " _
            & " FROM Client_Master a " _
               & " LEFT JOIN TownCity b " _
                  & " ON a.sTownIDxx = b.sTownIDxx " _

   If SearchValue Then
      lsSQL = lsSQL & " WHERE sLastName + ', ' + sFrstName + ' ' + sMiddName = '" & txtField(3).Text & "'"
   Else
      lsSQL = lsSQL & " WHERE sLastName + ', ' + sFrstName + ' ' + sMiddName LIKE '" & txtField(3).Text & "%' "
   End If
   lsSQL = lsSQL & " ORDER BY sLastName + ', ' + sFrstName + ' ' + sMiddName"
   lrs.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   
   
   Select Case lrs.RecordCount
      Case 0
         frmCustomer.Show 1
      Case 1
         txtField(3).Text = lrs("xFullName")
         txtField(4).Text = IIf(IsNull(lrs("sPhoneNox")), "", lrs("sPhoneNox"))
         oDriver.FieldValue(3) = lrs(0)
      Case Is > 1
        lsSearch = KwikBrowse(oApp, lrs, _
                          "sClientID»" _
                        & "xFullName»" _
                     & "xAddressx", _
                          "Client ID»" _
                        & "Name»" _
                        & "Address")

        If lsSearch <> "" Then
            psSelected = Split(lsSearch, "»")
            txtField(3).Text = psSelected(1)
            txtField(4).Text = IIf(IsNull(psSelected(3)), "", psSelected(3))
            oDriver.FieldValue(3) = lrs(0)
        End If
   End Select
   Set lrs = Nothing

End Sub

Private Sub oDriver_LoadOtherData()
Dim lsSQL As String
   
   pbnewitem = False
   Set oRS = New ADODB.Recordset
      lsSQL = "SELECT" _
            & " a.sTransNOx, " _
            & " a.cWarranty, " _
            & " a.cCategory, " _
            & " a.sClientID, " _
            & " b.sLastName + ', ' + b.sFrstName + ' ' + b.sMiddName xFullName " _
         & " FROM CP_JobOrder_Master a " _
            & " LEFT JOIN Client_Master b " _
               & " ON a.sClientID = b.sClientID " _
         & " WHERE cTranstat = 0 " _
            & " AND sTransNox = '" & oDriver.FieldValue(0) & "'"
   If oRS.State = adStateOpen Then oRS.Close
   oRS.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   
   txtField(3).Text = IIf(Not IsNull(oRS("xFullName")), oRS("xFullName"), "")
   oDriver.FieldValue(3) = oRS("sClientID")
   chkfieldValue
   optionValue
   SearchSupplier
   Set oRS = Nothing

End Sub
Private Sub chkField_Click(Index As Integer)
If Index = 3 Then
   If txtField(9).Enabled = True Then txtField(9).SetFocus
End If
End Sub

Private Sub chkfieldValue()
Dim Index As Integer
'0 - WalkIn ; 1-Void Warranty ; 2-Under Limited Warranty ; 3-Back Job

For Index = 1 To 3
   If Index = oRS("cWarranty") Then
      chkField(Index).Value = 1
   Else
      chkField(Index).Value = 0
   End If
Next
End Sub

Private Sub optionValue()
Dim Index As Integer
'0-New Unit; 1-Battery; 2-Sim Card; 3-back Cover; 4-Charger; 5-Others
   Index = oRS("cCategory")
   Option1(Index).Value = True
End Sub

Private Sub oDriver_SaveComplete()
   MsgBox "Transaction Successfully Saved", vbInformation, "Information"
End Sub

Private Sub oDriver_WillSave(Cancel As Boolean)
Dim Index As Integer

   If txtField(1).Text = "" And Option1(5).Value = False Then
      MsgBox "Invalid Job Order No. Detected!!!", vbCritical, "Warning"
      txtField(1).SetFocus
      Cancel = True
   ElseIf txtField(3).Text = "" And Option1(5).Value = False Then
      MsgBox "Invalid Customer Detected!!!", vbCritical, "Warning"
      txtField(3).SetFocus
      Cancel = True
   ElseIf txtField(5).Text = "" Then
      MsgBox "Invalid Date Detected!!!", vbCritical, "Warning"
      txtField(5).SetFocus
      Cancel = True
   ElseIf txtField(7).Text = "" And Option1(5).Value = False Then
      MsgBox "Invalid Brand Detected!!!", vbCritical, "Warning"
      txtField(7).SetFocus
      Cancel = True
   ElseIf txtField(8).Text = "" And Option1(5).Value = False Then
      MsgBox "Invalid Model Detected!!!", vbCritical, "Warning"
      txtField(8).SetFocus
      Cancel = True
   Else
      If pbnewitem Then
         Cancel = Not SaveCP_JO_Detail
            If Cancel Then Exit Sub
      End If
      
      For Index = 1 To 3   'cWarranty
         If chkField(Index).Value = 1 Then
            oDriver.FieldValue(13) = Index
         End If
      Next
      If chkField(1).Value = 0 And chkField(2).Value = 0 And _
         chkField(3).Value = 0 Then oDriver.FieldValue(13) = 0
      
      For Index = 0 To 5   'cCategory
         If Option1(Index).Value = True Then oDriver.FieldValue(14) = Index
      Next
      
      oDriver.FieldValue(1) = Trim(txtField(1).Text)   'JO No
      oDriver.FieldValue(2) = Trim(txtField(2).Text)   'IMEI No.
      oDriver.FieldValue(4) = Trim(txtField(4).Text)   'Cell #
      oDriver.FieldValue(5) = Format(txtField(5).Text, "M/D/YYYY") _
                              & " " & Format(oApp.ServerDate, "hh:nn:ss AM/PM")
      oDriver.FieldValue(6) = Format(txtField(6).Text, "M/d/yyyy")
      oDriver.FieldValue(9) = Trim(txtField(9).Text)   'BackJobNo
      oDriver.FieldValue(10) = Trim(txtField(10).Text) 'sComplent
      oDriver.FieldValue(11) = Trim(txtField(11).Text) 'sCategory
      oDriver.FieldValue(12) = Trim(txtothers(0).Tag) 'sSupplier
      oDriver.FieldValue(15) = 0 'cTranStat
         
   End If
   
End Sub

Private Function SaveCP_JO_Detail() As Boolean
Dim lsSQL As String
Dim lnrow As Long
   
SaveCP_JO_Detail = True
On Error GoTo errProc

         lsSQL = "INSERT INTO CP_JobOrder_Detail " _
                  & "( sTransNox, " _
                  & "  nEntryNox, " _
                  & "  sDescript, " _
                  & "  nPartsAmt, " _
                  & "  nLaborAmt, " _
                  & "  nDiscount, " _
                  & "  nQuantity, " _
                  & "  dModified) " _
                      & "VALUES " _
                      & "('" & oDriver.FieldValue(0) & "', " _
                      & "'1', " _
                      & "'', " _
                      & "'0.00', " _
                      & "'0.00', " _
                      & "'0', " _
                      & "'0', " _
                      & " getdate())"
                                   
         oApp.Connection.Execute lsSQL, lnrow, adCmdText
                   
         If lnrow <= 0 Then
            MsgBox "Unable to Save Job Order_Detail!!!" & vbCrLf & vbCrLf & _
            "Notify ROSALYN LAZO DESCALLAR for Assistance", vbCritical, "Warning"
            SaveCP_JO_Detail = False
            GoTo endProc
         End If
   
endProc:
   Exit Function
errProc:
   SaveCP_JO_Detail = False
   MsgBox Err.Description, vbCritical, "Warning"
End Function

Private Sub txtField_GotFocus(Index As Integer)
   oDriver.ColumnIndex = Index
   txtfieldGotfocus = True
   pnindex = Index
   txtOthersGotfocus = False
   txtField(Index).BackColor = &HE1FEFF
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim lsSQL As String
Dim lsCondition As String
Dim orig As String

   If KeyCode = vbKeyF3 Or KeyCode = 13 Then
      Select Case Index
         Case 2
            SearchSerial
         Case 3
            Search_Client False
         Case 7, 8
            oDriver.RecordSearch txtField(Index).Text
         Case 9
            SearchBackJob False
      End Select
      If oDriver.FieldValue(Index) <> "" Then SetNextFocus
      KeyCode = 0
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyReturn, vbKeyUp, vbKeyDown
      Select Case KeyCode
      Case vbKeyReturn, vbKeyDown
         If pnindex = 12 Then Exit Sub
         SetNextFocus
      Case vbKeyUp
         SetPreviousFocus
      End Select
   End Select
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   If Index = 6 Then
      If Not IsDate(txtField(Index).Text) Then
         If txtField(Index).Text = "" Then Exit Sub
         txtField(Index).Text = Format(oApp.ServerDate, "MMMM dd, yyyy")
      Else
         txtField(Index).Text = Format(txtField(6).Text, "MMMM dd, yyyy")
      End If
   End If
   txtField(Index).BackColor = &HFFFFFF
   If Index = 9 Then txtField(Index).BackColor = &HE7FFFF
End Sub

Private Sub txtfield_Validate(Index As Integer, Cancel As Boolean)
   Cancel = Not oDriver.ValidateField(Index)
   txtField(Index).Text = TitleCase(txtField(Index).Text)
   If Index = 2 Then
      If oDriver.FieldValue(2) <> "" Then SetNextFocus
   End If
End Sub

Private Sub txtothers_GotFocus(Index As Integer)
   oDriver.ColumnIndex = Index
   txtOthersGotfocus = True
   pnindex = Index
   txtOthersGotfocus = False
   txtothers(Index).BackColor = &HE1FEFF
End Sub

Private Sub txtothers_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Or KeyCode = 13 Then
      Select Case Index
         Case 0
            SearchSupplier
            SetNextFocus
            KeyCode = 0
      End Select
   End If
End Sub

Private Sub Print_JobOrder()
Dim lnCtr As Integer
Dim lrs As New ADODB.Recordset
Dim lrsReport As New ADODB.Recordset
Dim lsSQL As String

   Set lrs = New ADODB.Recordset
   Set lrsReport = New ADODB.Recordset

   lrs.Fields.Append "sField01", adVarChar, 150
   lrs.Fields.Append "sField02", adVarChar, 150
   lrs.Fields.Append "sField03", adVarChar, 150
   lrs.Fields.Append "sField04", adVarChar, 150
   lrs.Fields.Append "sField05", adVarChar, 150
   lrs.Fields.Append "sField06", adVarChar, 150
   lrs.Fields.Append "sField07", adVarChar, 150
   lrs.Fields.Append "sField08", adVarChar, 150
   lrs.Fields.Append "sField09", adVarChar, 150
   lrs.Fields.Append "sField10", adVarChar, 150
   lrs.Open

   'Job_Order
    lsSQL = "SELECT" _
               & " a.sJobOrdNo, " _
               & " a.sSupplier, " _
               & " b.sLastName + ', ' + b.sFrstName +' '+ b.sMiddName as xFullName, " _
               & " b.sAddressx + ' ' + c.sTownName as xAddressx, " _
               & " a.sTelNoxxx, " _
               & " d.sBrandNme +' '+ e.sModelNme as xBrandMdl, " _
               & " a.sIMEINoxx, " _
               & " a.sComplent, " _
               & " a.cCategory, " _
               & " a.sCategory, " _
               & " a.dTransact, " _
               & " a.dPurchase, " _
               & " a.cWarranty, " _
               & " a.sBckJobNo, " _
               & " a.cTranStat, " _
               & " f.sBranchNm  " _

   lsSQL = lsSQL & " FROM CP_JobOrder_Master a " _
            & " LEFT JOIN Client_Master b " _
               & " ON a.sClientID = b.sClientID " _
            & " LEFT JOIN TownCity c " _
               & " ON b.sTownIDxx = c.sTownIDxx " _
            & " LEFT JOIN Brand d " _
               & " ON a.sBrandIDx = d.sBrandIDx " _
            & " LEFT JOIN Model e " _
               & " ON a.sModelIDx = e.sModelIDx," _
            & " Branch f " _
         & " WHERE sTransNox = '" & oDriver.FieldValue(0) & "' " _
            & " AND f.sBranchCd = '" & Left(oDriver.FieldValue(0), 2) & "' " _

   If lrsReport.State = adStateOpen Then lrsReport.Close
   lrsReport.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   
   If lrsReport.EOF Then
      MsgBox "Save Transaction!!!" & vbCrLf & _
             "Then Try Again!!!", vbCritical, "Warning"
      Exit Sub
   End If
         Progress.Open App.Path & "\images\FINDCOMP.AVI"
         Progress.Play

         lrs.AddNew
         lrs("sField01").Value = IIf(IsNull(lrsReport("xAddressx")), "", lrsReport("xAddressx"))
         lrs("sField02").Value = lrsReport("sBranchNm")
         Select Case lrsReport("cCategory")
            Case 0
               lrs("sField03").Value = "X"
            Case 1
               lrs("sField04").Value = "X"
            Case 2
               lrs("sField05").Value = "X"
            Case 3
               lrs("sField06").Value = "X"
            Case 4
               lrs("sField07").Value = "X"
         End Select
         
         Select Case lrsReport("cWarranty")
            Case 1
               lrs("sField08").Value = "X"
            Case 2
               lrs("sField09").Value = "X"
            Case 3
               lrs("sField10").Value = "X"
         End Select

         Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\CP_JobOrder.rpt")
         oReport.DiscardSavedData
         oReport.FieldMappingType = crAutoFieldMapping
         oReport.Database.SetDataSource lrs
         
         With oReport
            .Sections("PH").ReportObjects("txtJONo").SetText txtField(1).Text
            .Sections("PH").ReportObjects("txtCustomer").SetText txtField(3).Text
            .Sections("PH").ReportObjects("txtTelephone").SetText txtField(4).Text
            .Sections("PH").ReportObjects("txtBrandModel").SetText txtField(7).Text _
                                             + " " + txtField(8).Text
            .Sections("PH").ReportObjects("txtIMEINo").SetText txtField(2).Text
            .Sections("PH").ReportObjects("txtComplaint").SetText txtField(10).Text
            .Sections("PH").ReportObjects("txtReceived").SetText txtField(5).Text
            .Sections("PH").ReportObjects("txtOthers").SetText txtField(11).Text
            .Sections("PH").ReportObjects("txtBackJob").SetText txtField(9).Text
            .Sections("PH").ReportObjects("txtSupplier").SetText ""
            If txtField(5).Text <> txtField(6).Text Then
               .Sections("PH").ReportObjects("txtPurchased").SetText txtField(6).Text
            Else
               .Sections("PH").ReportObjects("txtPurchased").SetText ""
            End If
         End With
         
      Set lrs = Nothing
      Set lrsReport = Nothing
      
      frmViewer.CRViewer91.ReportSource = oReport
      frmViewer.CRViewer91.ViewReport
      frmViewer.Show

End Sub

Private Sub txtothers_LostFocus(Index As Integer)
   txtothers(Index).BackColor = &HFFFFFF
End Sub

Private Sub txtothers_Validate(Index As Integer, Cancel As Boolean)
   txtothers(Index).Text = TitleCase(txtothers(Index).Text)
End Sub
