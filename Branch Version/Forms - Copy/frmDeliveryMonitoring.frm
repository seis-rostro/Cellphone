VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmDeliveryMonitoring 
   BorderStyle     =   0  'None
   Caption         =   "Guanzon Delivery"
   ClientHeight    =   8340
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8340
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3270
      Left            =   1605
      TabIndex        =   24
      Top             =   4755
      Width           =   10155
      _ExtentX        =   17912
      _ExtentY        =   5768
      _Version        =   393216
      FocusRect       =   0
      SelectionMode   =   1
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   3945
      Left            =   6510
      Tag             =   "wt0;fb0"
      Top             =   675
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   6959
      BackColor       =   12632256
      ClipControls    =   0   'False
      BorderStyle     =   3
      Begin VB.CommandButton cmdDelete 
         Appearance      =   0  'Flat
         Caption         =   "Delete Route"
         Height          =   345
         Left            =   3600
         TabIndex        =   38
         Top             =   3540
         Width           =   1470
      End
      Begin VB.TextBox txtBranch 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   735
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "UEMI Santiago - Multi"
         Top             =   435
         Width           =   4305
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   2715
         Left            =   105
         TabIndex        =   37
         Top             =   780
         Width           =   4950
         _ExtentX        =   8731
         _ExtentY        =   4789
         _Version        =   393216
         FocusRect       =   0
         AllowUserResizing=   2
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ENCODE ROUTES(BRANCH) OF DELIVERY HERE"
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
         Left            =   735
         TabIndex        =   22
         Top             =   165
         Width           =   4335
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Branch"
         Height          =   195
         Index           =   8
         Left            =   165
         TabIndex        =   23
         Top             =   480
         Width           =   510
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   9
      Left            =   90
      TabIndex        =   34
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
      Picture         =   "frmDeliveryMonitoring.frx":0000
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   4155
      Left            =   1590
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   10185
      _ExtentX        =   17965
      _ExtentY        =   7329
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   6
         Left            =   1185
         TabIndex        =   5
         Text            =   "September 25, 2015"
         Top             =   2610
         Width           =   1650
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   750
         Index           =   8
         Left            =   1185
         TabIndex        =   7
         Text            =   "This is a test. Don't take things as it is."
         Top             =   3300
         Width           =   3660
      End
      Begin VB.TextBox txtArrTime 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2880
         TabIndex        =   20
         Text            =   "03:45 PM"
         Top             =   2955
         Width           =   1050
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   7
         Left            =   1185
         TabIndex        =   6
         Text            =   "September 25, 2015"
         Top             =   2955
         Width           =   1650
      End
      Begin VB.TextBox txtDepTime 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2880
         TabIndex        =   18
         Text            =   "04:30 AM"
         Top             =   2610
         Width           =   1050
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   5
         Left            =   1185
         TabIndex        =   3
         Text            =   "Dacasin, Princess Joy"
         Top             =   1920
         Width           =   3660
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   4
         Left            =   1185
         TabIndex        =   2
         Text            =   "Cuison, Michael Torres"
         Top             =   1575
         Width           =   3660
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   3
         Left            =   1185
         TabIndex        =   1
         Text            =   "Adversalo, Rex Soriano"
         Top             =   1230
         Width           =   3660
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   1185
         TabIndex        =   4
         Text            =   "UEMI Tuguegarao - Multi"
         Top             =   2265
         Width           =   3660
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
         Height          =   315
         Index           =   0
         Left            =   1110
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "M001-12-000001"
         Top             =   240
         Width           =   1605
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   1185
         TabIndex        =   0
         Text            =   "AVX 897609"
         Top             =   885
         Width           =   1395
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
         ForeColor       =   &H80000007&
         Height          =   270
         Left            =   3000
         TabIndex        =   39
         Tag             =   "eb0;et0"
         Top             =   315
         Width           =   1680
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   330
         Left            =   3000
         Tag             =   "et0;ht2"
         Top             =   300
         Width           =   1680
      End
      Begin VB.Shape Shape3 
         Height          =   375
         Index           =   0
         Left            =   2970
         Top             =   270
         Width           =   1740
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   195
         Index           =   10
         Left            =   165
         TabIndex        =   21
         Top             =   3360
         Width           =   630
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Est. Arrival"
         Height          =   195
         Index           =   7
         Left            =   165
         TabIndex        =   19
         Top             =   3015
         Width           =   750
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Porter #2"
         Height          =   195
         Index           =   6
         Left            =   165
         TabIndex        =   15
         Top             =   1995
         Width           =   660
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Porter #1"
         Height          =   195
         Index           =   5
         Left            =   165
         TabIndex        =   14
         Top             =   1650
         Width           =   660
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Driver"
         Height          =   195
         Index           =   4
         Left            =   165
         TabIndex        =   13
         Top             =   1305
         Width           =   420
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Destination"
         Height          =   195
         Index           =   3
         Left            =   165
         TabIndex        =   16
         Top             =   2340
         Width           =   795
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Departure"
         Height          =   195
         Index           =   1
         Left            =   165
         TabIndex        =   17
         Top             =   2655
         Width           =   705
      End
      Begin VB.Label lblField 
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
         Index           =   0
         Left            =   165
         TabIndex        =   10
         Top             =   285
         Width           =   915
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   315
         Left            =   1215
         Tag             =   "et0;ht2"
         Top             =   345
         Width           =   1605
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vehicle No"
         Height          =   195
         Index           =   2
         Left            =   165
         TabIndex        =   12
         Top             =   960
         Width           =   780
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   10
      Left            =   90
      TabIndex        =   9
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
      Picture         =   "frmDeliveryMonitoring.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   90
      TabIndex        =   28
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
      Picture         =   "frmDeliveryMonitoring.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   6
      Left            =   90
      TabIndex        =   31
      Top             =   1815
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
      Picture         =   "frmDeliveryMonitoring.frx":166E
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   90
      TabIndex        =   25
      Top             =   555
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
      Picture         =   "frmDeliveryMonitoring.frx":1DE8
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   7
      Left            =   90
      TabIndex        =   32
      Top             =   4335
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
      Picture         =   "frmDeliveryMonitoring.frx":2562
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   90
      TabIndex        =   27
      Top             =   3705
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "Confir&m"
      AccessKey       =   "m"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmDeliveryMonitoring.frx":2CDC
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   90
      TabIndex        =   29
      Top             =   2445
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "Dispa&tch"
      AccessKey       =   "t"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmDeliveryMonitoring.frx":3456
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   90
      TabIndex        =   26
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
      Picture         =   "frmDeliveryMonitoring.frx":3BD0
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   8
      Left            =   90
      TabIndex        =   33
      Top             =   5595
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "Cl&ose"
      AccessKey       =   "o"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmDeliveryMonitoring.frx":434A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   90
      TabIndex        =   30
      Top             =   3075
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "Arri&val"
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
      Picture         =   "frmDeliveryMonitoring.frx":4AC4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   13
      Left            =   90
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   1815
      Width           =   1245
      _ExtentX        =   2196
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
      Picture         =   "frmDeliveryMonitoring.frx":523E
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   11
      Left            =   90
      TabIndex        =   36
      Top             =   4965
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
      Picture         =   "frmDeliveryMonitoring.frx":59B8
   End
End
Attribute VB_Name = "frmDeliveryMonitoring"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pxeMODULENAME = "frmDeliveryMonitoring"
Private WithEvents oTrans As clsDeliveryMonitoring
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin
Private p_sDivision As String

Dim p_nRow As Integer
Dim p_nDet As Integer
Dim p_nCtr As Integer
Dim p_nIndex As Integer
Dim p_bGridFocus As Boolean
Dim loFrm As frmDeliveryMonDetail

Property Let Division(ByVal Value As String)
   p_sDivision = Value
End Property

Private Sub cmdButton_Click(Index As Integer)
   Dim lnRep As Integer

   Select Case Index
   Case 0   'Browse
      If oTrans.SearchTransaction() Then
         loadFields
      Else
         ClearFields
      End If
   Case 1   'Print
      lnRep = MsgBox("Do You want to Print This Transaction", vbYesNo + vbQuestion, "Confirm")
                     If lnRep = vbYes Then
                        PrintTrans
                     End If
   Case 2   'Confirm
      If txtField(0).Text <> "" And oTrans.Master("cTranStat") = xeStateOpen Then
         oTrans.BroadCast (oTrans.Master("sTransNox"))
         Call loadFields
         MsgBox "Transaction Successfully Confirm"
      End If
   Case 3   'Cancel
      lnRep = MsgBox("Transaction is in Update Mode!!!" & vbCrLf & _
                     "Do you want to Cancel Entry of Transaction!!!", vbYesNo + vbQuestion, "Confirm")
         
      If lnRep = vbYes Then
         If oTrans.OpenTransaction("") Then
            Call loadFields
         Else
            Call ClearFields
         End If
         initButton xeModeReady
      Else
         txtField(p_nIndex).SetFocus
      End If
   Case 4   'Dispatch
      If txtField(0).Text <> "" And oTrans.Master("cTranStat") = xeStateOpen Then
         oTrans.Departure (oTrans.Master("sTransNox"))
         Call loadFields
         MsgBox "Transaction Successfully Depart"
      End If
   Case 5   'Arrival
   Case 6   'New
      If oTrans.NewTransaction Then
         oTrans.InitTransaction
   
         InitGridRoute
         InitGridDetail
         loadFields
         
         initButton xeModeAddNew
         txtField(1).SetFocus
      End If
   Case 7   'Update
      If txtField(0).Text <> "" And (oTrans.Master("cTranstat") = xeStateOpen _
                                     Or oTrans.Master("cTranStat") = xeStateClosed) Then
         If oTrans.UpdateTransaction Then
            Call addRoute
            initButton xeModeUpdate
            txtField(1).SetFocus
         End If
      End If
   Case 8   'Close
      Unload Me
      
   Case 9   'Search
      If p_bGridFocus Then
         Call txtBranch_KeyDown(vbKeyF3, 0)
         txtBranch.SetFocus
      Else
         oTrans.SearchMaster p_nIndex, ""
         txtField(p_nIndex).SetFocus
      End If
   Case 10  'Saves
      With MSFlexGrid2
         If .Rows > 1 Then
            p_nCtr = 0
            Do While p_nCtr < .Rows
               If Trim(.TextMatrix(p_nCtr, 1)) = "" Then
                  oTrans.deleteRoute p_nCtr
                  .Rows = .Rows - 1
               End If
               p_nCtr = p_nCtr + 1
            Loop
         End If
         
         If isEntryOk Then
            If oTrans.SaveTransaction Then
               MsgBox "Record Save Successfully!!!", vbInformation, "Confirm"
'                  lnRep = MsgBox("Do You want to Print This Transaction", vbYesNo + vbQuestion, "Confirm")
'                     If lnRep = vbYes Then
'                        PrintTrans
'                     End If
               initButton xeModeReady
            Else
               MsgBox "Unable to Update Record!!!", vbCritical, "Warning"
            End If
         End If
      End With
   Case 11  'Cancel
      If txtField(0).Text <> "" And (oTrans.Master("cTranStat") = xeStateOpen Or _
                                       oTrans.Master("cTranStat") = xeStateClosed) Then
         oTrans.CancelTransaction (oTrans.Master("sTransNox"))
         oTrans.NewTransaction
         oTrans.InitTransaction
   
         InitGridRoute
         InitGridDetail
         loadFields
         
         initButton xeModeUpdate
      Else
         MsgBox "Unable to Update Record!!!", vbCritical, "Warning"
      End If
   Case 12  'Add Route
   Case 13  'Delete Detail
      Call deleteDetail
   End Select
End Sub

Private Sub cmdDelete_Click() 'delete route
   Call deleteRoute
End Sub

Private Sub Form_Activate()
   MSFlexGrid1.Refresh
   MSFlexGrid2.Refresh
   oApp.MenuName = Me.Tag
   Me.ZOrder 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyReturn, vbKeyDown
      If KeyCode = vbKeyReturn And GetFocus = txtBranch.hwnd Then Exit Sub
      SetNextFocus
   Case vbKeyUp
      SetPreviousFocus
   End Select
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String
   Dim loTxt As TextBox
   lsOldProc = "Form_Load"
   ''On Error GoTo errProc
   
   CenterChildForm mdiMain, Me

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualLeft
   
   Set oTrans = New clsDeliveryMonitoring
   Set oTrans.AppDriver = oApp
   oTrans.Division = p_sDivision
   oTrans.InitTransaction
   oTrans.NewTransaction
   
   InitGridRoute
   InitGridDetail
   loadFields
   initButton xeModeUpdate

'   For Each loTxt In txtField
'      loTxt.MaxLength = oTrans.MasFldSize(pnCtr)
'   Next

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
End Sub

Private Sub InitGridRoute()
   
   Dim lnCtr As Integer
   
   With MSFlexGrid2
      .Rows = 2
      .Cols = 2
      .Font = "MS Sans Serif"

      'column title
      .TextMatrix(0, 0) = "No"
      .TextMatrix(0, 1) = "Branch/Route"
      .Row = 0

      'column alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = 1
      Next

      'column width
      .ColWidth(0) = 450
      .ColWidth(1) = 4390
      
      .TextMatrix(1, 0) = "1"
      
      .Row = 1
      .Col = 1
      .ColSel = .Cols - 1
   End With
End Sub

Private Sub InitGridDetail()
   Dim lnCtr As Integer
   With MSFlexGrid1
      .Rows = 2
      .Cols = 5
      .Font = "MS Sans Serif"

      'column title
      .TextMatrix(0, 0) = "No"
      .TextMatrix(0, 1) = "Route"
      .TextMatrix(0, 2) = "Transfer"
      .TextMatrix(0, 3) = "Refer No"
      .TextMatrix(0, 4) = "Recepient"
      .Row = 0

      'column alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = 1
      Next

      'column width
      .ColWidth(0) = 450
      .ColWidth(1) = 3350
      .ColWidth(2) = 1500
      .ColWidth(3) = 1500
      .ColWidth(4) = 3250

      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      .ColAlignment(4) = 1
      
      .TextMatrix(1, 0) = "1"
      
      .Row = 1
      .Col = 1
      .ColSel = .Cols - 1
   End With
End Sub

Private Sub loadFields()
   Dim loTxt As TextBox
   Dim lnRow As Integer
   Dim lnCtr As Integer
   Dim lnCur As Integer
   
   For Each loTxt In txtField
      Select Case loTxt.Index
      Case 0
         loTxt.Text = Format(oTrans.Master(loTxt.Index), "@@@@@@-@@@@@@")
      Case 6, 7
         loTxt.Text = Format(oTrans.Master(loTxt.Index), "MMMM DD, YYYY")
      Case Else
         loTxt.Text = oTrans.Master(loTxt.Index)
      End Select
   Next
   
   txtDepTime = Format(oTrans.Master("dTransact"), "HH:MM AM/PM")
   txtArrTime = Format(oTrans.Master("dEArrival"), "HH:MM AM/PM")
   txtBranch = ""

   With MSFlexGrid1
      If oTrans.ItemCount = 0 Then
         .Rows = 2
         
         .TextMatrix(1, 1) = ""
         .TextMatrix(1, 2) = ""
         .TextMatrix(1, 3) = ""
         .TextMatrix(1, 4) = ""
      Else
         .Rows = oTrans.ItemCount + 1
         lnCur = 0
         For lnCtr = 0 To oTrans.RouteCount - 1
            If Trim(oTrans.Route(lnCtr, "sBranchCd")) <> "" Then
               For lnRow = 0 To oTrans.ItemCount(oTrans.Route(lnCtr, "sBranchCd")) - 1
                  .TextMatrix(lnCur + 1, 0) = lnCur + 1
                  .TextMatrix(lnCur + 1, 1) = oTrans.Route(lnCtr, "sBranchNm")
                  .TextMatrix(lnCur + 1, 2) = oTrans.Detail(lnRow, "sDescript", oTrans.Route(lnCtr, "sBranchCd"))
                  .TextMatrix(lnCur + 1, 3) = oTrans.Detail(lnRow, "sReferNox", oTrans.Route(lnCtr, "sBranchCd"))
                  .TextMatrix(lnCur + 1, 4) = oTrans.Detail(lnRow, "sDestinat", oTrans.Route(lnCtr, "sBranchCd"))
                  lnCur = lnCur + 1
               Next
            End If
         Next
      End If
      
      .Row = 1
      .Col = 1
      .ColSel = .Cols - 1
   End With

   With MSFlexGrid2
      .Rows = oTrans.RouteCount + 1
      
      For lnRow = 0 To oTrans.RouteCount - 1
         .TextMatrix(lnRow + 1, 0) = lnRow + 1
         .TextMatrix(lnRow + 1, 1) = IFNull(oTrans.Route(lnRow, "sBranchNm"), "")
      Next
      
      .Row = 1
      .Col = 1
      .ColSel = .Cols - 1
   End With
   
   Label2.Caption = Format(TransStat(oTrans.Master("cTranStat")), ">")
End Sub

Private Sub ClearFields()
   Dim loTxt As TextBox
   
   For Each loTxt In txtField
      loTxt.Text = ""
   Next
   
   txtDepTime = ""
   txtArrTime = ""
   txtBranch = ""
   
   With MSFlexGrid1
      .Rows = 2
      
      .TextMatrix(1, 0) = 1
      .TextMatrix(1, 1) = ""
      .TextMatrix(1, 2) = ""
      .TextMatrix(1, 3) = ""
      .TextMatrix(1, 4) = ""
      
      .Row = 1
      .Col = 1
      .ColSel = .Cols - 1
   End With
   
   With MSFlexGrid2
      .Rows = 2
      
      .TextMatrix(1, 0) = 1
      .TextMatrix(1, 1) = ""
      
      .Row = 1
      .Col = 1
      .ColSel = .Cols - 1
   End With
End Sub

Private Sub initButton(ByVal fnEdit As xeEditMode)
   If fnEdit = xeModeAddNew Then fnEdit = xeModeUpdate
   
   xrFrame1.Enabled = (fnEdit = xeModeUpdate)
   xrFrame2.Enabled = (fnEdit = xeModeUpdate)
   
   cmdButton(3).Visible = (fnEdit = xeModeUpdate)   'Cancel
   cmdButton(9).Visible = (fnEdit = xeModeUpdate)   'Search
   cmdButton(10).Visible = (fnEdit = xeModeUpdate)  'Save
'   cmdButton(12).Visible = (fnEdit = xeModeUpdate)  'Save
   cmdButton(13).Visible = (fnEdit = xeModeUpdate)  'Save
   
   cmdButton(0).Visible = Not (fnEdit = xeModeUpdate)   'Browse
   cmdButton(1).Visible = Not (fnEdit = xeModeUpdate)   'Print
   cmdButton(2).Visible = Not (fnEdit = xeModeUpdate)   'Confirm
   cmdButton(4).Visible = Not (fnEdit = xeModeUpdate)   'Dispatch
   cmdButton(5).Visible = Not (fnEdit = xeModeUpdate)   'Arrival
   cmdButton(6).Visible = Not (fnEdit = xeModeUpdate)   'New
   cmdButton(7).Visible = Not (fnEdit = xeModeUpdate)   'Update
   cmdButton(8).Visible = Not (fnEdit = xeModeUpdate)   'Close
   cmdButton(11).Visible = Not (fnEdit = xeModeUpdate)   'Close
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

Private Sub MSFlexGrid1_GotFocus()
   p_bGridFocus = False
End Sub

Private Sub MSFlexGrid1_RowColChange()
   With MSFlexGrid1
      p_nDet = .Row
   End With
End Sub

Private Sub MSFlexGrid2_DblClick()
   Call LoadFormDetail
   Call loadFields
End Sub

Private Sub MSFlexGrid2_RowColChange()
   With MSFlexGrid2
      p_nRow = .Row
      txtBranch.Text = .TextMatrix(.Row, 1)
   End With
End Sub

Private Sub oTrans_MasterRetrieved(ByVal Index As Variant, ByVal Value As Variant)
   txtField(Index) = Value
End Sub

Private Sub txtArrTime_GotFocus()
   Call HighlightOn(Me.txtArrTime)
   p_bGridFocus = False
End Sub

Private Sub txtArrTime_LostFocus()
   Call HighlightOff(Me.txtArrTime)
End Sub

Private Sub txtArrTime_Validate(Cancel As Boolean)
   With txtArrTime
      If Not IsDate(.Text) Then .Text = oApp.ServerDate
      .Text = Format(.Text, "HH:MM AM/PM")
      
      oTrans.Master(7) = txtField(7).Text & " " & .Text
   End With
End Sub

Private Sub txtBranch_GotFocus()
   Call HighlightOn(Me.txtBranch)
   p_bGridFocus = True
End Sub

Private Sub txtBranch_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String
   
   lsOldProc = "txtBranch_KeyDown"
   ''On Error GoTo errProc
   
   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
      With txtBranch
         If KeyCode = vbKeyF3 Then
            oTrans.SearchRoute p_nRow - 1, .Text
            Call LoadFormDetail
            Call loadFields
         Else
            If .Text <> "" Then oTrans.SearchRoute p_nRow - 1, .Text
            Call LoadFormDetail
            Call loadFields
         End If
         Call addRoute
         .SetFocus
      End With
      KeyCode = 0
   End If

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " _
                       & "  " & KeyCode _
                       & ", " & Shift _
                       & " )", True
End Sub

Private Sub txtBranch_LostFocus()
   Call HighlightOff(Me.txtBranch)
End Sub

Private Sub txtBranch_Validate(Cancel As Boolean)
   oTrans.Route(p_nRow - 1, "sBranchNm") = txtBranch.Text
   Call addRoute
End Sub

Private Sub txtDepTime_GotFocus()
   Call HighlightOn(Me.txtDepTime)
   p_bGridFocus = False
End Sub

Private Sub txtDepTime_LostFocus()
   Call HighlightOff(Me.txtDepTime)
End Sub

Private Sub txtDepTime_Validate(Cancel As Boolean)
   With txtDepTime
      If Not IsDate(.Text) Then .Text = oApp.ServerDate
      .Text = Format(.Text, "HH:MM AM/PM")
      
       oTrans.Master(6) = txtField(6).Text & " " & .Text
   End With
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   Select Case Index
   Case 6, 7
      txtField(Index) = strShortDate(txtField(Index))
      Call HighlightOn(Me.txtField(Index))
   Case Else
      Call HighlightOn(Me.txtField(Index))
   End Select
   
   p_bGridFocus = False
   p_nIndex = Index
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
   Select Case Index
   Case 6, 7
      Call HighlightOff(Me.txtField(Index))
      If Index = 1 Then Me.txtField(Index) = Format(Me.txtField(Index), "MMM DD, YYYY")
   Case Else
      Call HighlightOff(Me.txtField(Index))
   End Select
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   Dim lsOldProc As String
   
   lsOldProc = "txtField_Validate"
   ''On Error GoTo errProc
   
   With txtField(Index)
      Select Case Index
      Case 6
         oTrans.Master(Index) = .Text & " " & txtDepTime
      Case 7
         oTrans.Master(Index) = .Text & " " & txtArrTime
      Case Else
         oTrans.Master(Index) = .Text
      End Select
   End With

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & ", " & Cancel & " )", True
End Sub

Private Sub addRoute()
   Dim lnRow As Integer
   Dim lbExist As Boolean
   
   With MSFlexGrid2
      .Row = .Rows - 1
      If Trim(oTrans.Route(p_nRow - 1, "sBranchCd")) = "" Then Exit Sub
      For lnRow = 0 To oTrans.RouteCount
         If p_nRow <> lnRow + 1 Then
            If oTrans.Route(p_nRow - 1, "sBranchCd") = oTrans.Route(lnRow, "sBranchCd") Then
               MsgBox "Unable to insert duplicate route!!!" & vbCrLf & _
                        "Please verify your entry then try again...", vbCritical, "WARNING"
               Exit Sub
            End If
         End If
      Next
      .Row = .Rows - 1
      .TextMatrix(p_nRow, 1) = oTrans.Route(p_nRow - 1, "sBranchNm")
      
      If .Row = .Rows - 1 Then
         If Not oTrans.addRoute Then
            MsgBox "Unable to add route!!!" & vbCrLf & _
                     "Please contact GGC SSG/SEG for assistance...", vbCritical, "WARNING"
            Exit Sub
         Else
            If .TextMatrix(.Row, 1) <> "" Then .Rows = .Rows + 1
         End If
      End If
      
      .Row = .Rows - 1
      .Col = 1
      .ColSel = .Cols - 1
      p_nRow = .Row
      
      .TextMatrix(.Rows - 1, 0) = .Rows - 1
      txtBranch.Text = .TextMatrix(.Rows - 1, 1)
   End With
End Sub

Private Sub deleteRoute()
   Dim lnRow As Integer
   Dim lsBranchCd As String
   Dim lnCur As Integer
   Dim lnCtr As Integer
   
   With MSFlexGrid2
      If .Rows > 1 Then
         lsBranchCd = oTrans.Route(p_nRow - 1, "sBranchCd")
         If oTrans.deleteRoute(p_nRow - 1) Then
            If oTrans.RouteCount > 0 Then
               .Rows = oTrans.RouteCount + 1
               
               For lnRow = 0 To oTrans.RouteCount - 1
                  .TextMatrix(lnRow + 1, 1) = IFNull(oTrans.Route(lnRow, "sBranchNm"), "")
               Next
            Else
               .Rows = 2
               
               .TextMatrix(1, 1) = ""
            End If
            
            If oTrans.delBranchDetail(lsBranchCd) Then
               With MSFlexGrid1
                  If oTrans.ItemCount > 0 Then
                     .Rows = oTrans.ItemCount + 1
                     lnCur = 0
                     For lnCtr = 0 To oTrans.RouteCount - 1
                        If Trim(oTrans.Route(lnCtr, "sBranchCd")) <> "" Then
                           For lnRow = 0 To oTrans.ItemCount(oTrans.Route(lnCtr, "sBranchCd")) - 1
                              .TextMatrix(lnCur + 1, 0) = lnCur + 1
                              .TextMatrix(lnCur + 1, 1) = oTrans.Route(lnCtr, "sBranchNm")
                              .TextMatrix(lnCur + 1, 2) = oTrans.Detail(lnRow, "sDescript", oTrans.Route(lnCtr, "sBranchCd"))
                              .TextMatrix(lnCur + 1, 3) = oTrans.Detail(lnRow, "sReferNox", oTrans.Route(lnCtr, "sBranchCd"))
                              .TextMatrix(lnCur + 1, 4) = oTrans.Detail(lnRow, "sDestinat", oTrans.Route(lnCtr, "sBranchCd"))
                              lnCur = lnCur + 1
                           Next
                           lnCur = lnRow
                        End If
                     Next
                  Else
                     .Rows = 2
                     .TextMatrix(lnCur + 1, 1) = ""
                     .TextMatrix(lnCur + 1, 2) = ""
                     .TextMatrix(lnCur + 1, 3) = ""
                     .TextMatrix(lnCur + 1, 4) = ""
                  End If
                           
                  .Row = .Rows - 1
                  .Col = 1
                  .ColSel = .Cols - 1
                  p_nDet = .Row
               End With
            End If
         Else
            MsgBox "Unable to delete route!!!" & vbCrLf & _
                     "Please contact GGC SEG/SSG for assistance...", vbCritical, "WARNING"
         End If
      End If
      
      .Row = .Rows - 1
      .Col = 1
      .ColSel = .Cols - 1
      p_nRow = .Row
      
      txtBranch.Text = .TextMatrix(.Rows - 1, 1)
   End With
End Sub

Private Function isEntryOk() As Boolean
   If txtField(1).Text = "" Then
      MsgBox "Vehicle not found!!!" & vbCrLf & _
             "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
      txtField(1).SetFocus
      GoTo EntryNotOK
   End If
   
   If txtField(2).Text = "" Then
      MsgBox "Destination not found!!!" & vbCrLf & _
             "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
      txtField(2).SetFocus
      GoTo EntryNotOK
   End If
   
   If txtField(3).Text = "" Then
      MsgBox "Driver not found!!!" & vbCrLf & _
             "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
      txtField(3).SetFocus
      GoTo EntryNotOK
   End If
   
   With MSFlexGrid2
      If Trim(.TextMatrix(1, 1)) = "" Then
         MsgBox "Route is required!!!" & vbCrLf & _
                "Please Verify your Entry then Try Again", vbCritical, "Warning"
         .SetFocus
         .Row = 1
         .Col = 1
         GoTo EntryNotOK
      End If
   End With
   
EntryOK:
   isEntryOk = True
   Exit Function
EntryNotOK:
   isEntryOk = False
End Function

Private Sub LoadFormDetail()
   Set loFrm = New frmDeliveryMonDetail
   Set loFrm.Delivery = oTrans
   loFrm.EntryNo = MSFlexGrid2.Row
   loFrm.Show 1
End Sub

Private Sub deleteDetail()
   Dim lsBranchCd As String
   Dim lsReferNox As String
   Dim lsSourceCd As String
   Dim lnCtr As Integer
   Dim lnCur As Integer
   Dim lnRow As Integer
   Dim lnRte As Integer

   With MSFlexGrid1
      If .Rows > 1 Then
         lsBranchCd = oTrans.Detail(p_nDet - 1, "sBranchCd")
         lsReferNox = oTrans.Detail(p_nDet - 1, "sReferNox")
         lsSourceCd = oTrans.Detail(p_nDet - 1, "sSourceCd")
      
         If oTrans.deleteDetail(lsBranchCd, lsReferNox, lsSourceCd) Then
            If oTrans.ItemCount(lsBranchCd) = 0 Then
               With MSFlexGrid2
                  For lnRte = 0 To oTrans.RouteCount - 1
                     If oTrans.Route(lnRte, "sBranchCd") = lsBranchCd Then
                        If oTrans.deleteRoute(lnRte) Then
                           If oTrans.RouteCount > 0 Then
                              .Rows = oTrans.RouteCount + 1
                              For lnCtr = 0 To oTrans.RouteCount - 1
                                 .TextMatrix(lnCtr + 1, 0) = lnCtr + 1
                                 .TextMatrix(lnCtr + 1, 1) = IFNull(oTrans.Route(lnCtr, "sBranchNm"), "")
                              Next
                           Else
                              .Rows = 2
                              .TextMatrix(1, 1) = ""
                           End If
                           
                           .Row = .Rows - 1
                           .Col = 1
                           .ColSel = .Cols - 1
                        End If
                     End If
                  Next
               End With
            End If
            
            If oTrans.ItemCount > 0 Then
               .Rows = oTrans.ItemCount + 1
               lnCur = 0
               For lnCtr = 0 To oTrans.RouteCount - 1
                  If Trim(oTrans.Route(lnCtr, "sBranchCd")) <> "" Then
                     For lnRow = 0 To oTrans.ItemCount(oTrans.Route(lnCtr, "sBranchCd")) - 1
                        .TextMatrix(lnCur + 1, 0) = lnCur + 1
                        .TextMatrix(lnCur + 1, 1) = oTrans.Route(lnCtr, "sBranchNm")
                        .TextMatrix(lnCur + 1, 2) = oTrans.Detail(lnRow, "sDescript", oTrans.Route(lnCtr, "sBranchCd"))
                        .TextMatrix(lnCur + 1, 3) = oTrans.Detail(lnRow, "sReferNox", oTrans.Route(lnCtr, "sBranchCd"))
                        .TextMatrix(lnCur + 1, 4) = oTrans.Detail(lnRow, "sDestinat", oTrans.Route(lnCtr, "sBranchCd"))
                        lnCur = lnCur + 1
                     Next
                     lnCur = lnRow
                  End If
               Next
            Else
               .Rows = 2
               .TextMatrix(lnCur + 1, 1) = ""
               .TextMatrix(lnCur + 1, 2) = ""
               .TextMatrix(lnCur + 1, 3) = ""
               .TextMatrix(lnCur + 1, 4) = ""
            End If
                           
            .Row = .Rows - 1
            .Col = 1
            .ColSel = .Cols - 1
            p_nDet = .Row
         End If
      End If
   End With
End Sub

Public Function PrintTrans(Optional ByVal pbPrint As Boolean = False) As Boolean
   Dim loreport As frmRepViewer
   Dim lsOldProc As String
   Dim lrs As ADODB.Recordset
   Dim lnCtr As Integer
   Dim loSource As Recordset
   Dim lsSQL As String
   
   lsOldProc = "PrinTrans"
   '''On Error GoTo errProc

   PrintTrans = True
   
   Set lrs = New ADODB.Recordset
   
   lrs.Fields.Append "sField01", adVarChar, 120
   lrs.Fields.Append "sField02", adVarChar, 50
   lrs.Fields.Append "sField03", adVarChar, 50
   lrs.Open
   
   lsSQL = "SELECT a.sTransNox" & _
               ", b.sBranchNm" & _
               ", a.sReferNox" & _
               ", a.sSourceCd" & _
               ", c.sSourceNm" & _
            " FROM Delivery_Detail a" & _
                  " LEFT JOIN xxxTransactionSource c" & _
                     " ON a.sSourceCd = c.sSourceID" & _
            ", Branch b" & _
            " WHERE sTransNox = " & strParm(oTrans.Master("sTransNox")) & _
            " AND a.sBranchCd = b.sBranchCd"
            
   Debug.Print lsSQL
   Set loSource = New Recordset
   loSource.Open lsSQL, oApp.Connection, , , adCmdText
   
   Do Until loSource.EOF
      lrs.AddNew
      lrs("sField01").Value = loSource("sBranchNm")
      lrs("sField02").Value = loSource("sReferNox")
      lrs("sField03").Value = IFNull(loSource("sSourceNm"), loSource("sSourceCd"))
      lnCtr = lnCtr + 1
   loSource.MoveNext
   Loop
      
   'assign important info to the report
   
   Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\DeliveryTransfer.rpt")
   oReport.DiscardSavedData
   oReport.FieldMappingType = crAutoFieldMapping
   oReport.Database.SetDataSource lrs
   
   lsSQL = "SELECT a.sBranchNm, a.sAddressx, b.sTownName, c.sProvName" & _
            " FROM Branch a" & _
               " LEFT JOIN TownCity b" & _
                  " ON a.sTownIDxx = b.sTownIDxx" & _
               " LEFT JOIN Province c" & _
                  " ON b.sProvIDxx = c.sProvIDxx" & _
            " WHERE a.sBranchCd = " & strParm(oTrans.Master("sDestinat"))
   
   Set loSource = New Recordset
   loSource.Open lsSQL, oApp.Connection, , , adCmdText
   
   oReport.Sections("RH").ReportObjects("txtHeadTitle").SetText oApp.BranchName
   oReport.Sections("RH").ReportObjects("txtHeadDescription").SetText oApp.Address & ", " & oApp.TownCity & " " & oApp.Province & " " & oApp.ZippCode
   oReport.Sections("RH").ReportObjects("txtRefNo").SetText Right(oTrans.Master("sTransNox"), 8)
   oReport.Sections("RH").ReportObjects("txtDate").SetText Format(txtField(6).Text, "MMMM DD, YYYY")
   oReport.Sections("PH").ReportObjects("txtTo").SetText txtField(2).Text
   oReport.Sections("PH").ReportObjects("txtToAddress").SetText loSource("sAddressx") & " " & strParm(loSource("sTownName")) & ", " & strParm(loSource("sProvName"))
   oReport.Sections("PH").ReportObjects("txtFrom").SetText oApp.BranchName
   oReport.Sections("PH").ReportObjects("txtFromAddress").SetText oApp.Address & ", " & oApp.TownCity & ", " & oApp.Province & " " & oApp.ZippCode
   oReport.Sections("PF").ReportObjects("txtPrepared").SetText oApp.UserName
   oReport.Sections("RF").ReportObjects("txtNote").SetText oTrans.Master("sRemarksx")
   
   oReport.PrintOutEx False, 1
'   Set loreport = New frmRepViewer
'   Set loreport.ReportSource = oReport
'
'   loreport.Show
   
endPoc:
   Call oTrans.CloseTransaction(oTrans.Master(0))
   Set lrs = Nothing
   Set loreport = Nothing
   Set oReport = Nothing
   Set lrs = Nothing
   Exit Function
errProc:
   PrintTrans = False
   ShowError lsOldProc & "( " & " )"
End Function



