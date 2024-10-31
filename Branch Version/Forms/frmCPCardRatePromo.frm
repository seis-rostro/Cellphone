VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmCPCardRatePromo 
   BackColor       =   &H8000000A&
   BorderStyle     =   0  'None
   Caption         =   "CP Credit Card Rates per Model"
   ClientHeight    =   7665
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14520
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   14520
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame2 
      Height          =   915
      Left            =   1575
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   1614
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.CheckBox Check1 
         Caption         =   "Filter"
         Height          =   210
         Index           =   1
         Left            =   3180
         TabIndex        =   5
         Tag             =   "wt0;fb0"
         Top             =   487
         Width           =   690
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Filter"
         Height          =   210
         Index           =   0
         Left            =   3180
         TabIndex        =   2
         Tag             =   "wt0;fb0"
         Top             =   187
         Width           =   690
      End
      Begin VB.TextBox txtFilter 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   735
         TabIndex        =   4
         Top             =   450
         Width           =   2430
      End
      Begin VB.TextBox txtFilter 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   735
         TabIndex        =   1
         Top             =   150
         Width           =   2430
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
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
         Left            =   165
         TabIndex        =   3
         Top             =   510
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank"
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
         Left            =   165
         TabIndex        =   0
         Top             =   210
         Width           =   450
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   90
      TabIndex        =   44
      Top             =   540
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
      Picture         =   "frmCPCardRatePromo.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   600
      Index           =   3
      Left            =   90
      TabIndex        =   43
      Top             =   1770
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
      Picture         =   "frmCPCardRatePromo.frx":077A
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   6090
      Index           =   0
      Left            =   1560
      Tag             =   "wt0;fb0"
      Top             =   1485
      Width           =   4020
      _ExtentX        =   7091
      _ExtentY        =   10742
      BorderStyle     =   1
      Begin VB.CheckBox chk36 
         Caption         =   "with 36 MOS"
         Height          =   210
         Left            =   1485
         TabIndex        =   38
         Tag             =   "wt0;fb0"
         Top             =   5790
         Width           =   1230
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   13
         Left            =   2025
         TabIndex        =   31
         Top             =   4425
         Width           =   1860
      End
      Begin VB.CheckBox chk24 
         Caption         =   "with 24 MOS"
         Height          =   210
         Left            =   1485
         TabIndex        =   37
         Tag             =   "wt0;fb0"
         Top             =   5580
         Width           =   1230
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   11
         Left            =   1350
         TabIndex        =   9
         Top             =   720
         Width           =   2520
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   12
         Left            =   1350
         TabIndex        =   11
         Top             =   1035
         Width           =   2520
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   10
         Left            =   1485
         TabIndex        =   35
         Top             =   5055
         Width           =   2400
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   9
         Left            =   1485
         TabIndex        =   33
         Top             =   4755
         Width           =   2400
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   1350
         TabIndex        =   21
         Top             =   2700
         Width           =   2520
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   1350
         TabIndex        =   19
         Top             =   2370
         Width           =   2520
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   1350
         TabIndex        =   17
         Top             =   2055
         Width           =   2520
      End
      Begin VB.ComboBox cmbShopType 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1350
         TabIndex        =   13
         Text            =   "Multi - Brand"
         Top             =   1365
         Width           =   2535
      End
      Begin VB.CheckBox chkActive 
         Caption         =   "Active"
         Height          =   285
         Left            =   1485
         TabIndex        =   36
         Tag             =   "wt0;fb0"
         Top             =   5325
         Width           =   825
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   1350
         TabIndex        =   15
         Top             =   1740
         Width           =   2520
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   5
         Left            =   2025
         TabIndex        =   23
         Top             =   3105
         Width           =   1860
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   6
         Left            =   2025
         TabIndex        =   25
         Top             =   3435
         Width           =   1860
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   7
         Left            =   2025
         TabIndex        =   27
         Top             =   3765
         Width           =   1860
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   8
         Left            =   2025
         TabIndex        =   29
         Top             =   4095
         Width           =   1860
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
         Left            =   855
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   105
         Width           =   1515
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "36 Mo. Term"
         Height          =   195
         Index           =   0
         Left            =   975
         TabIndex        =   30
         Top             =   4485
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Area Name"
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
         Index           =   11
         Left            =   165
         TabIndex        =   8
         Top             =   795
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Branch Name"
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
         Index           =   10
         Left            =   165
         TabIndex        =   10
         Top             =   1110
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Promo D. Thru"
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
         Left            =   165
         TabIndex        =   34
         Top             =   5070
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Promo D. From"
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
         Left            =   150
         TabIndex        =   32
         Top             =   4785
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model Code"
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
         Left            =   150
         TabIndex        =   20
         Top             =   2775
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model Name"
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
         Left            =   165
         TabIndex        =   18
         Top             =   2445
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Brand Name"
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
         Left            =   180
         TabIndex        =   16
         Top             =   2130
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Name"
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
         Left            =   165
         TabIndex        =   14
         Top             =   1785
         Width           =   990
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3 Mo. Term"
         Height          =   195
         Index           =   3
         Left            =   975
         TabIndex        =   22
         Top             =   3165
         Width           =   810
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "6 Mo. Term"
         Height          =   195
         Index           =   4
         Left            =   975
         TabIndex        =   24
         Top             =   3495
         Width           =   810
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "12 Mo. Term"
         Height          =   195
         Index           =   5
         Left            =   975
         TabIndex        =   26
         Top             =   3825
         Width           =   900
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "24 Mo. Term"
         Height          =   195
         Index           =   6
         Left            =   975
         TabIndex        =   28
         Top             =   4155
         Width           =   900
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3735
         Tag             =   "et0;ht2"
         Top             =   -1095
         Width           =   1500
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shop Type"
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
         TabIndex        =   12
         Top             =   1425
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trans#"
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
         Index           =   1
         Left            =   165
         TabIndex        =   6
         Top             =   135
         Width           =   615
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   90
      TabIndex        =   40
      Top             =   1770
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
      Picture         =   "frmCPCardRatePromo.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   90
      TabIndex        =   41
      Top             =   540
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
      Picture         =   "frmCPCardRatePromo.frx":166E
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   7020
      Left            =   5610
      TabIndex        =   39
      Top             =   540
      Width           =   8790
      _ExtentX        =   15505
      _ExtentY        =   12383
      _Version        =   393216
      Appearance      =   0
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   90
      TabIndex        =   42
      Top             =   1155
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&ADD"
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
      Picture         =   "frmCPCardRatePromo.frx":1DE8
   End
End
Attribute VB_Name = "frmCPCardRatePromo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmCPCardRatePromo"

Private oSkin As clsFormSkin

Dim pnCtr As Integer, pnIndex As Integer
Dim pbGridFocus As Boolean, pbSave As Boolean, pbLoaded As Boolean

Dim psShopType As String
Dim psBrandIDx As String
Dim psBankIDxx As String
Dim psModelIDx As String

Dim psSQLMaster As String
Dim psSQLLookUp(12) As String
Dim poRSMaster As Recordset

Private Sub Check1_Click(Index As Integer)
   Dim lsFilter As String
   
   If Check1(0).Value = Checked Then
      If txtFilter(0) <> "" Then lsFilter = "sBankIDxx = " & strParm(psBankIDxx)
   End If
   
   If Check1(1).Value = xeYes Then
      If lsFilter = "" Then
         If txtFilter(1) <> "" Then lsFilter = "sModelIDx = " & strParm(psModelIDx)
      Else
         If txtFilter(1) <> "" Then lsFilter = lsFilter & " AND sModelIDx = " & strParm(psModelIDx)
      End If
   End If
   
   If lsFilter = "" Then
      poRSMaster.Filter = ""
   Else
      poRSMaster.Filter = lsFilter
   End If
   InitForm
   reLoadDetail
End Sub

Private Sub chk24_Click()
   If chk24.Value = Checked Then
      MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 22) = 1
   Else
      MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 22) = 0
   End If
End Sub

Private Sub chk36_Click()
   If chk36.Value = Checked Then
      MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 23) = 1
   Else
      MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 23) = 0
   End If
End Sub

Private Sub chkActive_Click()
   If chkActive.Value = Checked Then
      MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 20) = 1
   Else
      MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 20) = 0
   End If
End Sub

Private Sub cmbShopType_Click()
   If cmbShopType.ListIndex = 0 Then
      MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 21) = 0
      txtField(2).Enabled = False
      Call txtField_Validate(2, True)
   Else
      txtField(2).Enabled = True
      MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 21) = 1
   End If
        
End Sub

Private Sub cmdButton_Click(Index As Integer)
   Select Case Index
   Case 0 'save
      If txtField(0).Text <> "" Then
         If SaveRecord Then
            MsgBox "Record Save Successfuly..", vbInformation, "Warning"
            LoadDetail
            initButton 0
         Else
            MsgBox "Unable to Save Records.", vbCritical, "Warning"
         End If
      End If
   Case 1 'update
      initButton 1
   Case 3 'cancel
      If MsgBox("This action will discard the updates made." & vbCrLf & _
                  "Do you want to continue?", vbQuestion & vbYesNo, "Confirm") = vbYes Then
         ClearFields
         LoadDetail
         initButton 0
      End If
   Case 4 'close
      Unload Me
   Case 5 'add
      If xrFrame1(0).Enabled Then
         With MSFlexGrid1
            If .TextMatrix(.Rows - 1, 7) <> "" And .TextMatrix(.Rows - 1, 10) <> "" And .TextMatrix(.Rows - 1, 12) <> "" Then
               .Rows = .Rows + 1
               If .Rows > 20 Then .TopRow = .Rows - 20
               .Row = .Rows - 1
            
               MSFlexGrid1_Click
               fillLastRow
               txtField(11).SetFocus
            End If
         End With
      End If
   End Select
End Sub

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0
   pbLoaded = True
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
   '''On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualLeft

   InitForm
   ClearFields
   initButton 0
   
   initSQL
   LoadDetail
   
   
   txtField(0).Text = GetNextCode("CP_Card_Rate_Model_Promo", "sTransNox", True, oApp.Connection, True, oApp.BranchCode)
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   pbLoaded = False
   Set poRSMaster = Nothing
End Sub

Private Sub MSFlexGrid1_Click()
   Dim lnCtr As Integer

   With MSFlexGrid1
      txtField(11) = .TextMatrix(.Row, 1) 'area
      txtField(12) = .TextMatrix(.Row, 2) 'branch
      txtField(1) = .TextMatrix(.Row, 3) 'bank
      txtField(2) = .TextMatrix(.Row, 4) 'brand
      txtField(3) = .TextMatrix(.Row, 5) 'Model
      txtField(4) = .TextMatrix(.Row, 6) 'code
      txtField(0) = .TextMatrix(.Row, 7) 'transnox
      txtField(5) = .TextMatrix(.Row, 13) '3
      txtField(6) = .TextMatrix(.Row, 14) '6
      txtField(7) = .TextMatrix(.Row, 15) '12
      txtField(8) = .TextMatrix(.Row, 16) '24
      txtField(13) = .TextMatrix(.Row, 17) '36
      txtField(9) = .TextMatrix(.Row, 18) 'datefr
      txtField(10) = .TextMatrix(.Row, 19) 'datethru
      chkActive.Value = IFNull(IIf(.TextMatrix(.Row, 20) = "", 0, .TextMatrix(.Row, 20)), 0)
      chk24.Value = IFNull(IIf(.TextMatrix(.Row, 22) = "", 0, .TextMatrix(.Row, 22)), 0)
      chk36.Value = IFNull(IIf(.TextMatrix(.Row, 23) = "", 0, .TextMatrix(.Row, 23)), 0)
      If IIf(.TextMatrix(.Row, 21) = "", 0, .TextMatrix(.Row, 21)) = 0 Then
         cmbShopType.Text = "Multi - Branch"
      Else
         cmbShopType.Text = "Single - Brand"
      End If
      If pbLoaded And xrFrame1(0).Enabled Then
         If .Col < 4 Or .Col > 8 Then .Col = 4
         txtField(.Col).SetFocus
      End If
      
      .Col = 0
      .ColSel = .Cols - 1
     
   End With
      
   pbGridFocus = True
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = oApp.getColor("HT1")
   End With
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String
   lsOldProc = "txtField_KeyDown"
   '''On Error GoTo errProc

   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
      Select Case Index
      Case 1 'search bank
         If searchBank(txtField(Index)) Then SetNextFocus
      Case 2 'search Brand
         If searchBrand(txtField(Index)) Then SetNextFocus
      Case 3 'search model
         If searchModel(txtField(Index)) Then SetNextFocus
      Case 11 'Area
         If searchArea(txtField(Index)) Then SetNextFocus
      Case 12 'Branch
         If searchBranch(txtField(Index)) Then SetNextFocus
      End Select
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

Private Function getFilter(ByVal Index As Integer, ByVal lsDescript As String) As Boolean
   Dim lsOldProc As String
   Dim lsSQL As String
   Dim lasDetail() As String
   
   lsOldProc = pxeMODULENAME & ".searchBrand"
   '''On Error GoTo errProc
   
   Select Case Index
   Case 0
      If txtFilter(0).Tag = lsDescript Then
         getFilter = True
         GoTo endProc
      End If
      
      lsSQL = AddCondition(psSQLLookUp(1), "sBankName LIKE " & strParm(lsDescript & "%"))
      
      lsSQL = KwikSearch(oApp, lsSQL, "sBankIDxx»sBankName", "Bank ID»Bank Name", "@»@")
      If lsSQL = "" Then
         psBankIDxx = ""
         GoTo endProc
      End If
      
      lasDetail = Split(lsSQL, "»")
      txtFilter(0) = lasDetail(1)
      txtFilter(0).Tag = lasDetail(1)
      psBankIDxx = lasDetail(0)
   Case 1
      If txtFilter(1).Tag = lsDescript Then
         getFilter = True
         GoTo endProc
      End If
      
      lsSQL = AddCondition(psSQLLookUp(3), "sModelNme LIKE " & strParm(lsDescript & "%"))
      Debug.Print lsSQL
      lsSQL = KwikSearch(oApp, lsSQL, "sModelIDx»sModelNme»sBrandNme", "Mode ID»Model Name»Brand", "@»@»@")
      If lsSQL = "" Then
         psBankIDxx = ""
         GoTo endProc
      End If
      
      lasDetail = Split(lsSQL, "»")
      txtFilter(1) = lasDetail(2)
      txtFilter(1).Tag = lasDetail(2)
      psModelIDx = lasDetail(0)

   End Select
   
   getFilter = True
endProc:
   Exit Function
errProc:
   ShowError lsOldProc
End Function


Private Function searchBank(ByVal lsDescript As String) As Boolean
   Dim lsOldProc As String
   Dim lsSQL As String
   Dim lasDetail() As String
   
'   lsOldProc = pxeMODULENAME & ".searchBank"
   '''On Error GoTo errProc
   
   If txtField(1).Tag = lsDescript Then
      searchBank = True
      GoTo endProc
   End If
   
   lsSQL = AddCondition(psSQLLookUp(1), "sBankName LIKE " & strParm(lsDescript & "%"))
   
   lsSQL = KwikSearch(oApp, lsSQL, "sBankIDxx»sBankName", "Bank ID»Bank Name", "@»@")
   If lsSQL = "" Then
      txtField(1) = ""
      txtField(1).Tag = ""
      
   MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) = ""
   MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 10) = ""
   Else
      lasDetail = Split(lsSQL, "»")
      txtField(1) = lasDetail(1)
      txtField(1).Tag = lasDetail(1)
      
      MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 10) = lasDetail(0)
       MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) = lasDetail(1)
   End If
   
   searchBank = True
endProc:
   Exit Function
errProc:
   ShowError lsOldProc
End Function

Private Function searchModel(ByVal lsDescript As String) As Boolean
   Dim lsOldProc As String
   Dim lsSQL As String
   Dim lasDetail() As String
   
'   lsOldProc = pxeMODULENAME & ".searchModel"
   '''On Error GoTo errProc
   
   If txtField(3).Tag = txtField(3) Then
      searchModel = True
      GoTo endProc
   End If
   
   lsSQL = AddCondition(psSQLLookUp(3), "a.sModelNme LIKE " & strParm(lsDescript & "%"))
   
   Debug.Print lsSQL
   lsSQL = KwikSearch(oApp, lsSQL, "sModelIDx»sBrandNme»sModelNme»sModelCde", _
                                    "Model ID»Brand»Model Name»Code", "@»@»@»@")
   If lsSQL = "" Then
      txtField(3) = ""
      txtField(4) = ""
      txtField(3).Tag = ""
      
      With MSFlexGrid1
         .TextMatrix(.Row, 5) = ""
         .TextMatrix(.Row, 6) = ""
         .TextMatrix(.Row, 12) = ""
      End With
   Else
      lasDetail = Split(lsSQL, "»")
      txtField(3) = lasDetail(2)
      txtField(4) = lasDetail(3)
      txtField(3).Tag = lasDetail(2)
      
      With MSFlexGrid1
         .TextMatrix(.Row, 5) = txtField(3)
         .TextMatrix(.Row, 6) = txtField(4)
         .TextMatrix(.Row, 12) = lasDetail(0)
      End With
   End If
   searchModel = True
endProc:
   Exit Function
errProc:
   ShowError lsOldProc
End Function

Private Function searchBrand(ByVal lsDescript As String) As Boolean
   Dim lsOldProc As String
   Dim lsSQL As String
   Dim lasDetail() As String
   
'   lsOldProc = pxeMODULENAME & ".searchBrand"
   '''On Error GoTo errProc
   
   If txtField(2).Tag = txtField(2) Then
      searchBrand = True
      GoTo endProc
   End If
   
   psSQLLookUp(2) = "SELECT" & _
                        "  sBrandIDx" & _
                        ", sBrandNme" & _
                     " FROM CP_Brand" & _
                     " WHERE cRecdStat = " & strParm(xeRecStateActive) & _
                     " ORDER BY sBrandNme"
                     
   lsSQL = AddCondition(psSQLLookUp(2), "sBrandNme LIKE " & strParm(lsDescript & "%"))
   
   Debug.Print lsSQL
   lsSQL = KwikSearch(oApp, lsSQL, "sBrandIDx»sBrandNme", _
                                    "Brand ID»Brand")
   If lsSQL = "" Then
      txtField(2) = ""
      txtField(2).Tag = ""
      
   MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = ""
   MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 11) = ""
   MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 21) = 0

   Else
      lasDetail = Split(lsSQL, "»")
      txtField(2) = lasDetail(1)
      txtField(2).Tag = lasDetail(1)
      
      MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = lasDetail(1)
      MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 11) = lasDetail(0)
      MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 21) = 1
   End If
   
   searchBrand = True
endProc:
   Exit Function
errProc:
   ShowError lsOldProc
End Function

Private Sub reLoadDetail()
   Dim lnCtr As Integer
   
   If poRSMaster.RecordCount = 0 Then
      fillLastRow
      Exit Sub
   End If
   
   With MSFlexGrid1
      .Rows = poRSMaster.RecordCount + 1
      For lnCtr = 0 To poRSMaster.RecordCount - 1
         .TextMatrix(lnCtr + 1, 0) = lnCtr + 1
         .TextMatrix(lnCtr + 1, 1) = IFNull(poRSMaster("sAreaDesc"), "")
         .TextMatrix(lnCtr + 1, 2) = IFNull(poRSMaster("sBranchNm"), "")
         .TextMatrix(lnCtr + 1, 3) = IFNull(poRSMaster("sBankName"), "")
         .TextMatrix(lnCtr + 1, 4) = IFNull(poRSMaster("sBrandNme"), "")
         .TextMatrix(lnCtr + 1, 5) = IFNull(poRSMaster("sModelNme"), "")
         .TextMatrix(lnCtr + 1, 6) = IFNull(poRSMaster("sModelCde"), "")
         .TextMatrix(lnCtr + 1, 7) = poRSMaster("sTransNox")
         .TextMatrix(lnCtr + 1, 8) = IFNull(poRSMaster("sAreaCode"), "")
         .TextMatrix(lnCtr + 1, 9) = IFNull(poRSMaster("sBranchCd"), "")
         .TextMatrix(lnCtr + 1, 10) = poRSMaster("sBankIDxx")
         .TextMatrix(lnCtr + 1, 11) = IFNull(poRSMaster("sBrandIdx"), "")
         .TextMatrix(lnCtr + 1, 12) = IFNull(poRSMaster("sModelIDx"))
         .TextMatrix(lnCtr + 1, 13) = Format(IFNull(poRSMaster("n03MoTerm"), 0), "#,##0.00")
         .TextMatrix(lnCtr + 1, 14) = Format(IFNull(poRSMaster("n06MoTerm"), 0), "#,##0.00")
         .TextMatrix(lnCtr + 1, 15) = Format(IFNull(poRSMaster("n12MoTerm"), 0), "#,##0.00")
         .TextMatrix(lnCtr + 1, 16) = Format(IFNull(poRSMaster("n24MoTerm"), 0), "#,##0.00")
         .TextMatrix(lnCtr + 1, 17) = Format(IFNull(poRSMaster("n36MoTerm"), 0), "#,##0.00")
         .TextMatrix(lnCtr + 1, 18) = Format(poRSMaster("dPromoFrm"), "YYYY-MM-DD")
         .TextMatrix(lnCtr + 1, 19) = Format(poRSMaster("dPromoTru"), "YYYY-MM-DD")
         .TextMatrix(lnCtr + 1, 20) = poRSMaster("cRecdStat")
         .TextMatrix(lnCtr + 1, 21) = poRSMaster("cShopType")
         .TextMatrix(lnCtr + 1, 22) = poRSMaster("cWith24Mo")
         .TextMatrix(lnCtr + 1, 23) = IFNull(poRSMaster("cWith36Mo"), 0)
         poRSMaster.MoveNext
         txtField(0).Text = .TextMatrix(lnCtr + 1, 7)
      Next
      
   End With

End Sub


Private Function LoadDetail() As Boolean
   Dim lsOldProc As String
   Dim lnCtr As Integer
   Dim lnRow As Integer
   
'   lsOldProc = pxeMODULENAME & ".LoadDetail"
   
   If TypeName(poRSMaster) = "Nothing" Then
      Set poRSMaster = New Recordset
   End If
   
   If poRSMaster.State = adStateOpen Then poRSMaster.Close
   Debug.Print psSQLMaster
   poRSMaster.Open psSQLMaster, oApp.Connection, adOpenStatic, adLockOptimistic, adCmdText
   Set poRSMaster.ActiveConnection = Nothing
   
   With MSFlexGrid1
      If poRSMaster.EOF Then
         MsgBox "No Record Found.", vbCritical, "Warning"
         .Rows = 2
         
         initButton 1
         fillLastRow
         txtField(0).Text = GetNextCode("CP_Card_Rate_Model_Promo", "sTransNox", True, oApp.Connection, True, oApp.BranchCode)
      Else
         lnRow = poRSMaster.RecordCount
         
         .Rows = lnRow + 1
'         .Rows = poRSMaster.RecordCount - 1
         
         For lnCtr = 0 To poRSMaster.RecordCount - 1
            .TextMatrix(lnCtr + 1, 0) = lnCtr + 1
            .TextMatrix(lnCtr + 1, 1) = IFNull(poRSMaster("sAreaDesc"), "")
            .TextMatrix(lnCtr + 1, 2) = IFNull(poRSMaster("sBranchNm"), "")
            .TextMatrix(lnCtr + 1, 3) = IFNull(poRSMaster("sBankName"), "")
            .TextMatrix(lnCtr + 1, 4) = IFNull(poRSMaster("sBrandNme"), "")
            .TextMatrix(lnCtr + 1, 5) = IFNull(poRSMaster("sModelNme"), "")
            .TextMatrix(lnCtr + 1, 6) = IFNull(poRSMaster("sModelCde"), "")
            .TextMatrix(lnCtr + 1, 7) = poRSMaster("sTransNox")
            .TextMatrix(lnCtr + 1, 8) = IFNull(poRSMaster("sAreaCode"), "")
            .TextMatrix(lnCtr + 1, 9) = IFNull(poRSMaster("sBranchCd"), "")
            .TextMatrix(lnCtr + 1, 10) = poRSMaster("sBankIDxx")
            .TextMatrix(lnCtr + 1, 11) = IFNull(poRSMaster("sBrandIdx"), "")
            .TextMatrix(lnCtr + 1, 12) = IFNull(poRSMaster("sModelIDx"))
            .TextMatrix(lnCtr + 1, 13) = Format(IFNull(poRSMaster("n03MoTerm"), 0), "#,##0.00")
            .TextMatrix(lnCtr + 1, 14) = Format(IFNull(poRSMaster("n06MoTerm"), 0), "#,##0.00")
            .TextMatrix(lnCtr + 1, 15) = Format(IFNull(poRSMaster("n12MoTerm"), 0), "#,##0.00")
            .TextMatrix(lnCtr + 1, 16) = Format(IFNull(poRSMaster("n24MoTerm"), 0), "#,##0.00")
            .TextMatrix(lnCtr + 1, 17) = Format(IFNull(poRSMaster("n36MoTerm"), 0), "#,##0.00")
            .TextMatrix(lnCtr + 1, 18) = Format(poRSMaster("dPromoFrm"), "YYYY-MM-DD")
            .TextMatrix(lnCtr + 1, 19) = Format(poRSMaster("dPromoTru"), "YYYY-MM-DD")
            .TextMatrix(lnCtr + 1, 20) = poRSMaster("cRecdStat")
            .TextMatrix(lnCtr + 1, 21) = poRSMaster("cShopType")
            .TextMatrix(lnCtr + 1, 22) = IFNull(poRSMaster("cWith24Mo"), 0)
            .TextMatrix(lnCtr + 1, 23) = IFNull(poRSMaster("cWith36Mo"), 0)
            Debug.Print poRSMaster("sModelIDx")
         poRSMaster.MoveNext
      Next
         If poRSMaster.EOF Then
            txtField(0).Text = GetNextCode("CP_Card_Rate_Model_Promo", "sTransNox", True, oApp.Connection, True, oApp.BranchCode)
         Else
            txtField(0).Text = IFNull(poRSMaster("sTransNox"))
         End If
      End If
   End With
   
endProc:
   Exit Function
errProc:
   ShowError (lsOldProc)
End Function

Private Sub initButton(lnStat As Integer)
   Dim lbShow As Boolean

   lbShow = IIf(lnStat = 0, False, True)

   cmdButton(1).Visible = Not lbShow
   cmdButton(4).Visible = Not lbShow
   cmdButton(3).Visible = lbShow
   cmdButton(0).Visible = lbShow

   xrFrame1(0).Enabled = lbShow
   xrFrame2.Enabled = Not lbShow
   
   If pbLoaded Then
      If lbShow Then
         txtField(1).SetFocus
      Else
         txtFilter(0).SetFocus
      End If
   End If
End Sub

Private Sub InitForm()
   Dim lnCtr As Integer
   
   With MSFlexGrid1
      .Clear
   
      .Cols = 24
      .Rows = 2
      .Font = "MS Sans Serif"
      
      'Column Title
      .TextMatrix(0, 1) = "Area"
      .TextMatrix(0, 2) = "Branch"
      .TextMatrix(0, 3) = "Bank"
      .TextMatrix(0, 4) = "Brand"
      .TextMatrix(0, 5) = "Model Name"
      .TextMatrix(0, 6) = "Code"
      .TextMatrix(0, 7) = "" 'Trans #
      .TextMatrix(0, 8) = "" 'AreaDesc
      .TextMatrix(0, 9) = "" 'Branchcd
      .TextMatrix(0, 10) = "" 'Bank id
      .TextMatrix(0, 11) = "" 'BrandID
      .TextMatrix(0, 12) = "" 'ModelID
      .TextMatrix(0, 13) = "" '3 MOS
      .TextMatrix(0, 14) = "" '6 MOS
      .TextMatrix(0, 15) = "" '12 MOS
      .TextMatrix(0, 16) = "" '24 MOS
      .TextMatrix(0, 17) = "" '36 MOS
      .TextMatrix(0, 18) = "" 'dPromoFR
      .TextMatrix(0, 19) = "" 'dPromoTO
      .TextMatrix(0, 20) = "" 'recStat
      .TextMatrix(0, 21) = 0 'cShoptype
      .TextMatrix(0, 22) = "" 'w/ 24 MOS
      .TextMatrix(0, 23) = "" 'w/ 36 MOS
      
      .Row = 0

      'Column Alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = 3
      Next
      .Row = 1

      'Column Width
      .ColWidth(0) = 300
      .ColWidth(1) = 1000
      .ColWidth(2) = 2200
      .ColWidth(3) = 1300
      .ColWidth(4) = 1000
      .ColWidth(5) = 1900
      .ColWidth(6) = 1000
      
      For lnCtr = 7 To 20
         .ColWidth(lnCtr) = 0
      Next
      
'      MSFlexGrid1_Click
   End With
   
   txtField(9).Text = oApp.ServerDate
   txtField(10).Text = oApp.ServerDate
   
   cmbShopType.AddItem "Multi - Brand"
   cmbShopType.AddItem "Single - Brand"
   cmbShopType.ListIndex = 0

End Sub

Private Sub ClearFields()
   Dim loTxt As TextBox
   
   For Each loTxt In txtField
      loTxt = ""
   Next
   
   psShopType = ""
   psBankIDxx = ""
   psBrandIDx = ""
   psModelIDx = ""
   
   txtFilter(0) = ""
   txtFilter(1) = ""
   
   Check1(0).Value = xeNo
   Check1(1).Value = xeNo
   
   InitForm
End Sub

Private Function isEntryOk() As Boolean

EntryOK:
   isEntryOk = True
   Exit Function
EntryNotOK:
   isEntryOk = False
End Function

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
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

Private Sub initSQL()

   psSQLMaster = "SELECT" & _
                     " a.sTransNox" & _
                     ", a.sBankIDxx" & _
                     ", a.sBrandIDx" & _
                     ", a.sModelIDx" & _
                     ", a.n03MoTerm" & _
                     ", a.n06MoTerm" & _
                     ", a.n12MoTerm" & _
                     ", a.n24MoTerm" & _
                     ", a.n36MoTerm" & _
                     ", a.dPromoFrm" & _
                     ", a.dPromoTru" & _
                     ", a.cRecdStat" & _
                     ", a.sAreaCode" & _
                     ", g.sBranchCd"

   psSQLMaster = psSQLMaster & _
                     ", a.cWith24Mo" & _
                     ", a.cWith36Mo" & _
                     ", a.cShopType" & _
                     ", b.sBankName" & _
                     ", c.sModelCde" & _
                     ", c.sModelNme" & _
                     ", d.sBrandNme" & _
                     ", f.sAreaDesc" & _
                     ", g.sBranchNm" & _
                     ", a.sApproved" & _
                     ", a.sModified" & _
                     ", a.dModified"

   psSQLMaster = psSQLMaster & _
                  " FROM CP_Card_Rate_Model_Promo a" & _
                     " LEFT JOIN Banks b" & _
                        " ON a.sBankIDxx = b.sBankIDxx" & _
                     " LEFT JOIN CP_Model c" & _
                        " ON a.sModelIDx = c.sModelIDx" & _
                     " LEFT JOIN CP_Brand d" & _
                        " ON a.sBrandIDx = d.sBrandIDx" & _
                     " LEFT JOIN Branch_Others e" & _
                        " ON a.sBranchCd = e.sBranchCd" & _
                     " LEFT JOIN Branch_Area f" & _
                        " ON a.sAreaCode = f.sAreaCode" & _
                     " LEFT JOIN Branch g" & _
                        " ON a.sBranchCd = g.sBranchCd" & _
                  " ORDER BY a.sTransNox,b.sBankName, d.sBrandNme, c.sModelNme"
                                       
   psSQLLookUp(1) = "SELECT" & _
                        "  sBankIDxx" & _
                        ", sBankName" & _
                     " FROM Banks" & _
                     " WHERE cRecdStat = " & strParm(xeRecStateActive) & _
                     " ORDER BY sBankName"
   
   psSQLLookUp(2) = "SELECT" & _
                        "  sBrandIDx" & _
                        ", sBrandNme" & _
                     " FROM CP_Brand" & _
                     " WHERE cRecdStat = " & strParm(xeRecStateActive) & _
                     " ORDER BY sBrandNme"
                     
   psSQLLookUp(3) = "SELECT" & _
                        "  a.sModelIDx" & _
                        ", b.sBrandNme" & _
                        ", a.sModelNme" & _
                        ", a.sModelCde" & _
                     " FROM CP_Model a" & _
                        " LEFT JOIN CP_Brand b" & _
                           " ON a.sBrandIDx = b.sBrandIDx" & _
                     " WHERE a.cRecdStat = " & strParm(xeRecStateActive) & _
                     " ORDER BY b.sBrandNme, a.sModelNme"
                     
   psSQLLookUp(11) = "SELECT" & _
                        " sAreaCode" & _
                        ", sAreaDesc" & _
                     " FROM Branch_Area " & _
                     " ORDER BY sAreaDesc"
   
   psSQLLookUp(12) = "SELECT" & _
                        " a.sBranchCd" & _
                        ", a.sBranchNm" & _
                        ", c.sAreaDesc" & _
                        ", c.sAreaCode" & _
                     " FROM Branch a" & _
                     ", Branch_Others b" & _
                     ", Branch_Area c" & _
                     " WHERE a.sBranchCd = b.sBranchCd" & _
                     " AND b.sAreaCode = c.sAreaCode" & _
                     " ORDER BY sBranchNm"
   
End Sub

Private Function SaveRecord() As Boolean
   Dim lsProcName As String
   Dim lnCtr As Integer
   Dim lnCoumn As Integer
   Dim lsSQL As String
   Dim lsSQLUpdate As String
   Dim lnColmCtr As Integer
   Dim lbRecord As Boolean
   
   lsProcName = pxeMODULENAME & ".SaveRecord"
   ''On Error GoTo errProc
   
   If TypeName(poRSMaster) = "Nothing" Then GoTo endProc
   
   With MSFlexGrid1
      oApp.BeginTrans
      
      If poRSMaster.RecordCount > 0 Then poRSMaster.MoveFirst
      lsSQLUpdate = ""
      For lnCtr = 1 To .Rows - 1
        
         poRSMaster.Filter = "sTransNox = " & strParm(.TextMatrix(lnCtr, 7))
         ' strParm(.TextMatrix(lnCtr, 7))
         If poRSMaster.EOF = True And .TextMatrix(lnCtr, 12) <> "" Then
            lsSQL = "INSERT INTO CP_Card_Rate_Model_Promo SET" & _
                     " sTransNox = " & strParm(GetNextCode("CP_Card_Rate_Model_Promo", "sTransNox", True, oApp.Connection, True, oApp.BranchCode)) & _
                     ", sAreaCode = " & strParm(.TextMatrix(lnCtr, 8)) & _
                     ", sBranchCd = " & strParm(.TextMatrix(lnCtr, 9)) & _
                     ", cShopType = " & strParm(.TextMatrix(lnCtr, 21)) & _
                     ", sBankIDxx = " & strParm(.TextMatrix(lnCtr, 10)) & _
                     ", sBrandIDx = " & strParm(.TextMatrix(lnCtr, 11)) & _
                     ", sModelIDx = " & strParm(.TextMatrix(lnCtr, 12)) & _
                     ", n03MoTerm = " & CDbl(.TextMatrix(lnCtr, 13)) & _
                     ", n06MoTerm = " & CDbl(.TextMatrix(lnCtr, 14)) & _
                     ", n12MoTerm = " & CDbl(.TextMatrix(lnCtr, 15)) & _
                     ", n24MoTerm = " & CDbl(.TextMatrix(lnCtr, 16)) & _
                     ", n36MoTerm = " & CDbl(.TextMatrix(lnCtr, 17)) & _
                     ", dPromoFrm = " & dateParm(CDate(.TextMatrix(lnCtr, 18))) & _
                     ", dPromoTru = " & dateParm(CDate(.TextMatrix(lnCtr, 19))) & _
                     ", sApproved = " & strParm(oApp.UserID) & _
                     ", cRecdStat = " & strParm(.TextMatrix(lnCtr, 20)) & _
                     ", cWith24Mo = " & strParm(.TextMatrix(lnCtr, 22)) & _
                     ", cWith36Mo = " & strParm(.TextMatrix(lnCtr, 23)) & _
                     ", sModified = " & strParm(oApp.UserID) & _
                     ", dModified = " & dateParm(oApp.ServerDate)
               If oApp.Execute(lsSQL, "CP_Card_Rate_Model_Promo") = 0 Then GoTo endWithRoll
          
          ElseIf .TextMatrix(lnCtr, 12) <> "" And .TextMatrix(lnCtr, 10) <> "" Then
            lsSQLUpdate = ""
            If .TextMatrix(lnCtr, 8) <> poRSMaster("sAreaCode") Then
               lsSQLUpdate = "sAreaCode = " & strParm(.TextMatrix(lnCtr, 8)) & ","
            End If
            If .TextMatrix(lnCtr, 9) <> poRSMaster("sBranchCd") Then
               lsSQLUpdate = lsSQLUpdate + "sBranchCd = " & strParm("sBranchCd = " & .TextMatrix(lnCtr, 9)) & ","
            End If
            If .TextMatrix(lnCtr, 21) <> poRSMaster("cShopType") Then
               lsSQLUpdate = lsSQLUpdate + "cShopType = " & strParm(IIf(.TextMatrix(lnCtr, 20) = "Multi - Brand", 0, 1)) & ","
            End If
            If .TextMatrix(lnCtr, 10) <> poRSMaster("sBankIDxx") Then
               lsSQLUpdate = lsSQLUpdate + "sBankIDxx = " & strParm(.TextMatrix(lnCtr, 10)) & ","
            End If
            If .TextMatrix(lnCtr, 11) <> poRSMaster("sBrandIDx") Then
               lsSQLUpdate = lsSQLUpdate + "sBrandIDx = " & strParm(.TextMatrix(lnCtr, 11)) & ","
            End If
            If .TextMatrix(lnCtr, 12) <> poRSMaster("sModelIDx") Then
               lsSQLUpdate = lsSQLUpdate + "sModelIDx = " & strParm(.TextMatrix(lnCtr, 12)) & ","
            End If
            If .TextMatrix(lnCtr, 13) <> Format(poRSMaster("n03MoTerm"), "#,##0.00") Then
               lsSQLUpdate = lsSQLUpdate + "n03MoTerm = " & .TextMatrix(lnCtr, 13) & ","
            End If
            If .TextMatrix(lnCtr, 14) <> Format(poRSMaster("n06MoTerm"), "#,##0.00") Then
               lsSQLUpdate = lsSQLUpdate + "n06MoTerm = " & .TextMatrix(lnCtr, 14) & ","
            End If
            If .TextMatrix(lnCtr, 15) <> Format(poRSMaster("n12MoTerm"), "#,##0.00") Then
               lsSQLUpdate = lsSQLUpdate + "n12MoTerm = " & .TextMatrix(lnCtr, 15) & ","
            End If
            If .TextMatrix(lnCtr, 16) <> Format(poRSMaster("n24MoTerm"), "#,##0.00") Then
               lsSQLUpdate = lsSQLUpdate + "n24MoTerm = " & .TextMatrix(lnCtr, 16) & ","
            End If
            If .TextMatrix(lnCtr, 17) <> Format(poRSMaster("n36MoTerm"), "#,##0.00") Then
               lsSQLUpdate = lsSQLUpdate + "n36MoTerm = " & .TextMatrix(lnCtr, 17) & ","
            End If
            If .TextMatrix(lnCtr, 18) <> Format(poRSMaster("dPromoFrm"), "YYYY-MM-DD") Then
               lsSQLUpdate = lsSQLUpdate + "dPromoFrm = " & dateParm(Format(.TextMatrix(lnCtr, 18), "YYYY-MM-DD")) & ","
            End If
            If .TextMatrix(lnCtr, 19) <> Format(poRSMaster("dPromoTru"), "YYYY-MM-DD") Then
               lsSQLUpdate = lsSQLUpdate + "dPromoTru = " & dateParm(Format(.TextMatrix(lnCtr, 19), "YYYY-MM-DD")) & ","
            End If
            If .TextMatrix(lnCtr, 20) <> poRSMaster("cRecdStat") Then
               lsSQLUpdate = lsSQLUpdate + "cRecdStat = " & strParm(.TextMatrix(lnCtr, 20)) & ","
            End If
            
            If .TextMatrix(lnCtr, 22) <> poRSMaster("cWith24Mo") Then
               lsSQLUpdate = lsSQLUpdate + "cWith24Mo = " & strParm(.TextMatrix(lnCtr, 22)) & ","
            End If
            If .TextMatrix(lnCtr, 23) <> IFNull(poRSMaster("cWith36Mo"), "") Then
               lsSQLUpdate = lsSQLUpdate + "cWith36Mo = " & strParm(.TextMatrix(lnCtr, 23)) & ","
            End If
            
            If lsSQLUpdate <> "" Then
               lsSQLUpdate = Left(lsSQLUpdate, Len(lsSQLUpdate) - 1)
   
               lsSQL = "UPDATE CP_Card_Rate_Model_Promo " & _
                     " SET " & _
                     lsSQLUpdate & _
                     " WHERE sTransNox = " & strParm(.TextMatrix(lnCtr, 7))
               
               Debug.Print lsSQL
               If oApp.Execute(lsSQL, "CP_Card_Rate_Model_Promo") = 0 Then GoTo endWithRoll

            End If
               
          End If
      Next
      oApp.CommitTrans
      
   End With
   
   SaveRecord = True
   Set poRSMaster = Nothing
endProc:
   Exit Function
endWithRoll:
   oApp.RollbackTrans
   GoTo endProc
errProc:
   ShowError lsProcName
End Function

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
   Case 1, 2
      If cmbShopType.ListIndex = 0 Then 'if multi brand then set brand = ""
         txtField(2).Text = ""
      End If
      
      If Index = 1 Then
         MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) = txtField(Index)
      ElseIf Index = 2 Then
         MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = txtField(Index)
      End If
      
   Case 5, 6, 7, 8, 13
      If Not IsNumeric(txtField(Index)) Then txtField(Index) = 0
         
      txtField(Index) = IIf(txtField(Index) = 0, "0.00", Format(txtField(Index), "#,##0.00"))
      
      With MSFlexGrid1
            .TextMatrix(.Row, 13) = txtField(5)
            .TextMatrix(.Row, 14) = txtField(6)
            .TextMatrix(.Row, 15) = txtField(7)
            .TextMatrix(.Row, 16) = txtField(8)
            .TextMatrix(.Row, 17) = txtField(13)
      End With
   Case 9
      If Not IsDate(txtField(Index)) Then txtField(Index) = oApp.ServerDate
      txtField(Index) = Format(txtField(Index).Text, "MM DD, YYYY")
      MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 18) = txtField(Index)
   Case 10
      If Not IsDate(txtField(Index)) Then txtField(Index) = oApp.ServerDate
      txtField(Index) = Format(txtField(Index).Text, "MM DD, YYYY")
      MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 19) = txtField(Index)
   Case 11
      If txtField(Index) = "" Then
         txtField(12) = ""
         txtField(11) = ""
         txtField(11).Tag = ""
         MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1) = ""
         MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 8) = ""
      End If
   End Select
End Sub

Private Sub txtFilter_GotFocus(Index As Integer)
   With txtFilter(Index)
      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = oApp.getColor("HT1")
   End With
End Sub

Private Sub txtFilter_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyReturn, vbKeyF3
      If getFilter(Index, txtFilter(Index)) Then SetNextFocus
   End Select
   KeyCode = 0
End Sub

Private Sub txtFilter_LostFocus(Index As Integer)
   With txtFilter(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub fillLastRow()
   With MSFlexGrid1
      .TextMatrix(.Rows - 1, 0) = .Rows - 1
      .TextMatrix(.Rows - 1, 1) = ""
      .TextMatrix(.Rows - 1, 2) = ""
      .TextMatrix(.Rows - 1, 3) = ""
      .TextMatrix(.Rows - 1, 4) = ""
      .TextMatrix(.Rows - 1, 5) = ""
      .TextMatrix(.Rows - 1, 6) = ""
      .TextMatrix(.Rows - 1, 7) = ""
      .TextMatrix(.Rows - 1, 8) = ""
      .TextMatrix(.Rows - 1, 9) = ""
      .TextMatrix(.Rows - 1, 10) = ""
      .TextMatrix(.Rows - 1, 11) = ""
      .TextMatrix(.Rows - 1, 12) = ""
      .TextMatrix(.Rows - 1, 13) = ""
      .TextMatrix(.Rows - 1, 14) = ""
      .TextMatrix(.Rows - 1, 15) = ""
      .TextMatrix(.Rows - 1, 16) = ""
      .TextMatrix(.Rows - 1, 17) = ""
      .TextMatrix(.Rows - 1, 18) = ""
      .TextMatrix(.Rows - 1, 19) = ""
      .TextMatrix(.Rows - 1, 20) = 1
      .TextMatrix(.Rows - 1, 21) = 0
      .TextMatrix(.Rows - 1, 22) = 0 'w/ 24 months
      .TextMatrix(.Rows - 1, 23) = 0 'w/ 36 months
      
      txtField(0).Text = GetNextCode("CP_Card_Rate_Model_Promo", "sTransNox", True, oApp.Connection, True, oApp.BranchCode)
      .TextMatrix(.Rows - 1, 7) = txtField(0).Text
      
   End With
   
   
End Sub

Private Function searchBranch(ByVal lsDescript As String) As Boolean
   Dim lsOldProc As String
   Dim lsSQL As String
   Dim lasDetail() As String
   
'   lsOldProc = pxeMODULENAME & ".SearchBranch"
   '''On Error GoTo errProc
   
   If txtField(12).Tag = txtField(12) Then
      searchBranch = True
      GoTo endProc
   End If
   
   lsSQL = AddCondition(psSQLLookUp(12), " sBranchNm LIKE " & strParm(lsDescript & "%"))
   
   Debug.Print lsSQL
   lsSQL = KwikSearch(oApp, lsSQL, "sBranchCd»sBranchNm", _
                                    "Code»Branch Name", "@»@")
   If lsSQL = "" Then
      txtField(12) = ""
      txtField(11) = ""
      txtField(12).Tag = ""
      MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = ""
      MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 8) = ""
      MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 9) = ""
   Else
      lasDetail = Split(lsSQL, "»")
      txtField(11) = lasDetail(2)
      txtField(12) = lasDetail(1)
       MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = lasDetail(1)
       MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 8) = lasDetail(3)
       MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 9) = lasDetail(0)
       MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1) = lasDetail(2)
   End If
   
   searchBranch = True
endProc:
   Exit Function
errProc:
   ShowError lsOldProc
End Function

Private Function searchArea(ByVal lsDescript As String) As Boolean
   Dim lsOldProc As String
   Dim lsSQL As String
   Dim lasDetail() As String
   
'   lsOldProc = pxeMODULENAME & ".searchArea"
   '''On Error GoTo errProc
   
   If txtField(11).Tag = txtField(11) Then
      searchArea = True
      GoTo endProc
   End If
   
   lsSQL = AddCondition(psSQLLookUp(11), " sAreaDesc LIKE " & strParm(lsDescript & "%"))
   
   Debug.Print lsSQL
   lsSQL = KwikSearch(oApp, lsSQL, "sAreaCode»sAreaDesc", _
                                    "Code»Area Name", "@»@")
   If lsSQL = "" Then
      txtField(12) = ""
      txtField(11) = ""
      txtField(11).Tag = ""
      MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1) = ""
      MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 8) = ""
   Else
      If lsSQL <> "" Then
         lasDetail = Split(lsSQL, "»")
         txtField(11) = lasDetail(1)
         MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1) = lasDetail(1)
         MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 8) = lasDetail(0)
      End If
   End If
   
   searchArea = True
endProc:
   Exit Function
errProc:
   ShowError lsOldProc
End Function
