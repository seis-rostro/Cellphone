VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmClientInfo 
   BorderStyle     =   0  'None
   Caption         =   "Client Maintenance"
   ClientHeight    =   6840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10380
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6840
   ScaleWidth      =   10380
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   4680
      Index           =   0
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   1140
      Width           =   10200
      _ExtentX        =   17992
      _ExtentY        =   8255
      BorderStyle     =   1
      Begin VB.TextBox txtOther 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   615
         Index           =   1
         Left            =   6285
         MultiLine       =   -1  'True
         TabIndex        =   42
         Top             =   3270
         Width           =   3660
      End
      Begin VB.TextBox txtOther 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   6285
         TabIndex        =   40
         Top             =   2970
         Width           =   3660
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   13
         Left            =   6285
         TabIndex        =   36
         Top             =   2370
         Width           =   3735
      End
      Begin VB.OptionButton optGender 
         Caption         =   "Female"
         Height          =   315
         Index           =   1
         Left            =   2040
         TabIndex        =   14
         Tag             =   "wt0;fb0"
         Top             =   1725
         Width           =   840
      End
      Begin VB.OptionButton optGender 
         Caption         =   "Male"
         Height          =   315
         Index           =   0
         Left            =   1185
         TabIndex        =   13
         Tag             =   "wt0;fb0"
         Top             =   1725
         Width           =   660
      End
      Begin VB.ComboBox cmbField 
         Height          =   315
         ItemData        =   "frmClientInfo.frx":0000
         Left            =   1185
         List            =   "frmClientInfo.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   2040
         Width           =   2070
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   14
         Left            =   6285
         TabIndex        =   38
         Top             =   2670
         Width           =   2070
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   945
         Index           =   12
         Left            =   6285
         MultiLine       =   -1  'True
         TabIndex        =   34
         Top             =   1410
         Width           =   3735
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   11
         Left            =   6285
         TabIndex        =   32
         Top             =   1110
         Width           =   2070
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   10
         Left            =   6285
         TabIndex        =   30
         Top             =   810
         Width           =   2070
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   9
         Left            =   1185
         TabIndex        =   28
         Top             =   4200
         Width           =   2070
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   8
         Left            =   1185
         TabIndex        =   26
         Top             =   3900
         Width           =   3735
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   6
         Left            =   1185
         TabIndex        =   22
         Top             =   2970
         Width           =   3735
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   1185
         TabIndex        =   20
         Top             =   2670
         Width           =   2070
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   1185
         TabIndex        =   18
         Top             =   2370
         Width           =   2070
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   615
         Index           =   7
         Left            =   1185
         MultiLine       =   -1  'True
         TabIndex        =   24
         Top             =   3270
         Width           =   3735
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   1185
         TabIndex        =   11
         Top             =   1410
         Width           =   3735
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   1185
         TabIndex        =   9
         Top             =   1110
         Width           =   3735
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   1185
         TabIndex        =   7
         Top             =   810
         Width           =   3735
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   1185
         TabIndex        =   5
         Top             =   210
         Width           =   1920
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   225
         Index           =   9
         Left            =   5085
         TabIndex        =   41
         Top             =   3330
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Spouse Name"
         Height          =   195
         Index           =   19
         Left            =   5085
         TabIndex        =   39
         Tag             =   "et0;fb0"
         Top             =   3030
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client #"
         Height          =   165
         Index           =   16
         Left            =   5085
         TabIndex        =   37
         Top             =   2730
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company Name"
         Height          =   165
         Index           =   15
         Left            =   5085
         TabIndex        =   35
         Top             =   2430
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Additional Info"
         Height          =   195
         Index           =   14
         Left            =   5085
         TabIndex        =   33
         Top             =   1470
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TIN"
         Height          =   195
         Index           =   13
         Left            =   5085
         TabIndex        =   31
         Top             =   1185
         Width           =   270
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email Add"
         Height          =   195
         Index           =   12
         Left            =   5085
         TabIndex        =   29
         Top             =   870
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone #"
         Height          =   195
         Index           =   11
         Left            =   150
         TabIndex        =   27
         Top             =   4260
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Town"
         Height          =   195
         Index           =   8
         Left            =   150
         TabIndex        =   25
         Top             =   3960
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gender"
         Height          =   195
         Index           =   22
         Left            =   150
         TabIndex        =   12
         Top             =   1785
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Birth Place"
         Height          =   195
         Index           =   7
         Left            =   150
         TabIndex        =   21
         Top             =   3030
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Birth Date"
         Height          =   195
         Index           =   6
         Left            =   150
         TabIndex        =   19
         Top             =   2730
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Citizenship"
         Height          =   195
         Index           =   5
         Left            =   150
         TabIndex        =   17
         Top             =   2430
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Civil Status"
         Height          =   195
         Index           =   4
         Left            =   150
         TabIndex        =   15
         Top             =   2100
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   225
         Index           =   3
         Left            =   150
         TabIndex        =   23
         Top             =   3330
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "First Name"
         Height          =   195
         Index           =   10
         Left            =   150
         TabIndex        =   8
         Top             =   1170
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name"
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   6
         Top             =   870
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Middle Name"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   10
         Top             =   1470
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client ID"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   4
         Top             =   270
         Width           =   600
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   285
         Left            =   1290
         Tag             =   "et0;ht2"
         Top             =   315
         Width           =   1920
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   0
      Left            =   9540
      TabIndex        =   50
      Top             =   6030
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmClientInfo.frx":0039
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   1
      Left            =   8760
      TabIndex        =   49
      Top             =   6030
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmClientInfo.frx":07B3
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   2
      Left            =   6420
      TabIndex        =   44
      Top             =   6030
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmClientInfo.frx":0F2D
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   4
      Left            =   5640
      TabIndex        =   43
      Top             =   6030
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmClientInfo.frx":16A7
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   5
      Left            =   9540
      TabIndex        =   51
      Top             =   6030
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmClientInfo.frx":1E21
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   705
      Index           =   7
      Left            =   7200
      TabIndex        =   45
      Top             =   6030
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmClientInfo.frx":259B
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   3
      Left            =   6420
      TabIndex        =   46
      Top             =   6030
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmClientInfo.frx":2D15
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   6
      Left            =   7200
      TabIndex        =   47
      Top             =   6030
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmClientInfo.frx":348F
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   8
      Left            =   7980
      TabIndex        =   48
      Top             =   6030
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
      Caption         =   "&Ledger"
      AccessKey       =   "L"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmClientInfo.frx":3C09
      PicturePos      =   1
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   555
      Index           =   1
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   10200
      _ExtentX        =   17992
      _ExtentY        =   979
      Begin VB.TextBox txtOther 
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
         Index           =   3
         Left            =   4890
         TabIndex        =   3
         Top             =   105
         Width           =   5160
      End
      Begin VB.TextBox txtOther 
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
         Index           =   2
         Left            =   1185
         TabIndex        =   1
         Top             =   105
         Width           =   1920
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Clien&t ID"
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
         Index           =   21
         Left            =   165
         TabIndex        =   0
         Top             =   150
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Full Name"
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
         Index           =   20
         Left            =   3960
         TabIndex        =   2
         Top             =   150
         Width           =   870
      End
   End
End
Attribute VB_Name = "frmClientInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmClientInfo"

Private WithEvents oDriver As clsFormDriver
Attribute oDriver.VB_VarHelpID = -1
Private oLedger As frmClientLedger
Private oSkin As clsFormSkin
Private bLoaded As Boolean

Dim pnIndex As Integer

Private Sub cmbField_Click()
   If Not cmbField.ListIndex = 1 Then txtOther(0).Text = ""
End Sub

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   
   lsOldProc = "cmdButton_Click"
   On Error GoTo errProc
   
   Select Case Index
   Case 0
      oDriver.RecordCancelUpdate
   Case 1
      oDriver.BrowseRecord
   Case 2
      oDriver.RecordSave
   Case 3
      oDriver.RecordUpdate
   Case 5
      Unload Me
   Case 6
      oDriver.RecordDelete
   Case 7
      oDriver.RecordSearch
   Case 8
      If txtField(0).Text <> "" Then
         oLedger.ClientID = oDriver.FieldValue(0)
         Load oLedger
         If oLedger.browseLedger Then
            oLedger.Show 1
         Else
            MsgBox "No Ledger found!!!", vbCritical, "Warning"
         End If
      End If
   End Select

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Sub Form_Activate()
   Dim lsOldProc As String
   
   lsOldProc = "Form_Activate"
   On Error GoTo errProc
   
   oApp.MenuName = Me.Tag
   Me.ZOrder 0

   If bLoaded Then
      oDriver.RecordCancelUpdate
      oDriver.DisableTextbox 0
      bLoaded = False
   End If
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyUp, vbKeyDown
      Select Case KeyCode
      Case vbKeyDown
         SetNextFocus
      Case vbKeyUp
         SetPreviousFocus
      End Select
   Case vbKeyF12
      'oDriver.ViewUserModify
   End Select
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String
   
   lsOldProc = "Form_Load"
   On Error GoTo errProc

   CenterChildForm mdiMain, Me
   bLoaded = True
   
   Set oLedger = New frmClientLedger
   
   Set oDriver = New clsFormDriver
   Set oDriver.AppDriver = oApp
   Set oDriver.MainForm = Me
   
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin
   
   oDriver.RecQuery = "SELECT" _
                        & "  sClientID" _
                        & ", sLastName" _
                        & ", sFrstName" _
                        & ", sMiddName" _
                        & ", sCitizenx" _
                        & ", dBirthDte" _
                        & ", sBirthPlc" _
                        & ", sAddressx" _
                        & ", sTownIDxx" _
                        & ", sPhoneNox" _
                        & ", sEmailAdd" _
                        & ", sTaxIDNox" _
                        & ", sAddlInfo" _
                        & ", sCompnyNm" _
                        & ", sClientNo" _
                        & ", cGenderCd" _
                        & ", cCvilStat" _
                        & ", sSpouseID" _
                        & ", sModified" _
                        & ", dModified" _
                     & " FROM Client_Master"

   oDriver.BrowseQuery = "SELECT " _
                           & "  a.sClientID" _
                           & ", CONCAT(a.sLastName, ', ', sFrstName, ' ', sMiddName) xFullName" _
                           & ", CONCAT(a.sAddressx, ', ', b.sTownName, ', ', c.sProvName) xAddressx" _
                        & " FROM Client_Master a" _
                           & " Left Join TownCity b" _
                              & " On a.sTownIDxx = b.sTownIDxx" _
                           & " Left Join Province c" _
                              & " On b.sProvIDxx = c.sProvIDxx" _
                        & " WHERE a.cRecdStat =" & strParm(xeRecStateActive) _
                        & " ORDER BY CONCAT(a.sLastName, ', ', sFrstName, ' ', sMiddName)"
   oDriver.InitRecForm
   
   oDriver.BrowseFTitle(0) = "Client ID"
   oDriver.BrowseFTitle(1) = "Full Name"
   oDriver.BrowseFTitle(2) = "Address"

   oDriver.LookupQuery(4) = "SELECT" _
                              & "  sCntryCde" _
                              & ", sNational" _
                              & ", sCntryNme" _
                           & " FROM Country" _
                           & " WHERE cRecdStat = " & strParm(xeRecStateActive) _

   oDriver.LookupReference(4) = "sCntryCde»sNational»sCntryNme"
   oDriver.LookupColumn(4) = "sCntryCde»sNational»sCntryNme"
   oDriver.LookupTitle(4) = "Code»Nationality»Country"

   oDriver.LookupQuery(8) = "SELECT" _
                              & "  a.sTownIDxx" _
                              & ", a.sTownName" _
                              & ", b.sProvName" _
                              & ", a.sZippCode" _
                           & " FROM TownCity a" _
                              & " LEFT JOIN Province b" _
                                 & " ON a.sProvIDxx = b.sProvIDxx" _
                           & " WHERE a.cRecdStat = " & strParm(xeRecStateActive) _
                           & " ORDER BY a.sTownName,b.sProvName"

   oDriver.LookupReference(8) = "a.sTownIDxx»a.sTownName»b.sProvName»a.sZippCode"
   oDriver.LookupColumn(8) = "sTownIDxx»sTownName»sProvName»sZippCode"
   oDriver.LookupTitle(8) = "TownID»Town»Province»ZippCode"
   
   oDriver.FieldFormat(0) = "@@@@-@@@@@@"
   oDriver.FieldFormat(5) = "MMMM DD, YYYY"
   oDriver.FieldStart = 1

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
   Set oDriver = Nothing
   Set oLedger = Nothing
End Sub

Private Sub oDriver_DisableOtherControl()
   xrFrame1(1).Enabled = True
   
   optGender(0).Enabled = False
   optGender(1).Enabled = False
   cmbField.Enabled = False
   
   If Not oApp.isMainOffice Then
      oDriver.hideButton 3
      oDriver.hideButton 4
      oDriver.hideButton 6
   End If
End Sub

Private Sub oDriver_EnableOtherControl()
   oDriver.DisableTextbox 0

   optGender(0).Enabled = True
   optGender(1).Enabled = True
   cmbField.Enabled = True
   
   xrFrame1(1).Enabled = False
End Sub

Private Sub oDriver_InitValue()
   Dim lnCtr As Integer
   Dim lsOldProc As String
   
   lsOldProc = "oDriver_InitValue"
   On Error GoTo errProc
   
   oDriver.FieldValue(0) = GetNextCode("Client_Master", "sClientID", True, oApp.Connection, True, oApp.BranchCode)
   txtField(0).Text = oDriver.FieldValue(0)
   
   For lnCtr = 0 To txtOther.Count - 1
      txtOther(lnCtr).Text = ""
      txtOther(lnCtr).Tag = ""
   Next
   
   optGender(0).Value = False
   optGender(1).Value = False
   
   cmbField.ListIndex = -1
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )"
End Sub

Private Sub oDriver_LoadOtherData()
   Dim lrs As ADODB.Recordset
   Dim lnCtr As Integer
   Dim lsOldProc As String
   
   lsOldProc = "oDriver_LoadOtherData"
   On Error GoTo errProc
   
   If Not IsNull(oDriver.FieldValue(15)) Then
      If Trim(oDriver.FieldValue(15)) <> "" Then
         optGender(oDriver.FieldValue(15)).Value = True
      Else
         optGender(0).Value = False
         optGender(1).Value = False
      End If
   Else
      optGender(0).Value = False
      optGender(1).Value = False
   End If
      
   If Not IsNull(oDriver.FieldValue(16)) Then
      If Trim(oDriver.FieldValue(16)) <> "" Then
         cmbField.ListIndex = oDriver.FieldValue(16)
      Else
         cmbField.ListIndex = -1
      End If
   Else
      cmbField.ListIndex = -1
   End If
   
   txtOther(0).Text = ""
   txtOther(2).Text = txtField(0).Text
   txtOther(2).Tag = txtOther(2).Text
   txtOther(3).Text = txtField(1).Text & ", " & txtField(2).Text & " " & txtField(3).Text
   txtOther(3).Tag = txtOther(3).Text
   
   Set lrs = New ADODB.Recordset
   lrs.Open "SELECT" _
               & "  CONCAT(a.sLastName, ', ', a.sFrstName, ' ', a.sMiddName) xFullName" _
               & ", CONCAT(a.sAddressx, ', ', b.sTownName, ' ', b.sZippCode, ', ', c.sProvName) xAddressx" _
            & " FROM Client_Master a" _
               & " LEFT JOIN TownCity b" _
                  & " ON a.sTownIDxx = b.sTownIDxx" _
               & " LEFT JOIN Province c" _
                  & " On b.sProvIDxx = c.sProvIDxx" _
            & " WHERE sClientID = " & strParm(IIf(IsNull(oDriver.FieldValue(17)), "", oDriver.FieldValue(17))) _
            , oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   
   If lrs.EOF Then GoTo endProc
   txtOther(0).Text = lrs("xFullName")
   txtOther(0).Tag = txtOther(0).Text
   txtOther(1).Text = IIf(IsNull(lrs("xAddressx")), "", lrs("xAddressx"))
   txtOther(1).Tag = txtOther(1).Text
   
endProc:
   Set lrs = Nothing
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )"
End Sub

Private Sub oDriver_WillSave(Cancel As Boolean)
   Dim lsOldProc As String
   
   lsOldProc = "oDriver_WillSave"
   On Error GoTo errProc

   
   If optGender(0).Value Or optGender(1).Value Then
      oDriver.FieldValue(15) = IIf(optGender(0).Value, 0, 1)
   End If
   If cmbField.ListIndex > -1 Then oDriver.FieldValue(16) = cmbField.ListIndex
   

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Cancel & " )"
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   Dim lsOldProc As String
   
   lsOldProc = "txtField_Validate"
   On Error GoTo errProc
   
   With txtField(Index)
      .Text = TitleCase(.Text)
   End With
   
   Cancel = Not oDriver.ValidateField(Index)
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " _
                       & "  " & Index _
                       & ", " & Cancel _
                       & " )", True
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String
   
   lsOldProc = "txtField_KeyDown"
   On Error GoTo errProc
   
   If KeyCode = vbKeyReturn Or KeyCode = vbKeyF3 Then
      With txtField(Index)
         If KeyCode = vbKeyF3 Then
            oDriver.RecordSearch .Text
            If .Text <> "" Then SetNextFocus
         Else
            If .Text <> "" Then oDriver.RecordSearch .Text
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

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("HT1")
   End With
   oDriver.ColumnIndex = Index
   pnIndex = Index
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtOther_GotFocus(Index As Integer)
   With txtOther(Index)
      .BackColor = oApp.getColor("HT1")
   End With
End Sub

Private Sub txtOther_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String
   
   lsOldProc = "txtOther_KeyDown"
   On Error GoTo errProc
   
   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
      With txtOther(Index)
         Select Case Index
         Case 2, 3
            If Trim(.Text) = "" Then
               oDriver_InitValue
               Exit Sub
            End If
                
            If .Tag <> .Text Then SearchCustomer Index, .Text, IIf(Index = 2, True, False)
         End Select
         .Tag = .Text
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

Private Sub txtOther_Validate(Index As Integer, Cancel As Boolean)
   Dim lsOldProc As String
   
   lsOldProc = "txtOther_Validate"
   On Error GoTo errProc
   
   With txtOther(Index)
      Select Case Index
      Case 2, 3
         If Trim(.Text) = "" Then
            oDriver_InitValue
            txtField(0).Text = ""
            Exit Sub
         End If
                
         If .Tag <> .Text Then SearchCustomer Index, .Text, IIf(Index = 2, True, False)
      End Select
      .Tag = .Text
   End With
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " _
                       & "  " & Index _
                       & ", " & Cancel _
                       & " )", True
End Sub

Private Sub txtOther_LostFocus(Index As Integer)
   With txtOther(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub SearchCustomer(ByVal lnIndex _
                           , ByVal lsValue As String _
                           , ByVal lbByCode As Boolean)
   Dim lsOldProc As String
   Dim lsSQL As String
   Dim lrs As ADODB.Recordset
   Dim lsBrowse As String
   Dim lsSelected() As String
   
   lsOldProc = "SearchCustomer"
   On Error GoTo errProc
   
   If lbByCode Then
      oDriver.LookupValue(0) = lsValue
      oDriver.LoadRecord
   Else
      lsSQL = AddCondition(oDriver.BrowseQuery _
                           , "CONCAT(a.sLastName, ', ', a.sFrstName) LIKE " & strParm(lsValue & "%"))
      Set lrs = New ADODB.Recordset
      lrs.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
      If lrs.EOF Then
         oDriver_InitValue
         txtField(0).Text = ""
         GoTo endProc
      End If
      
      With txtOther(lnIndex)
         If lrs.RecordCount = 1 Then
            oDriver.LookupValue(0) = lrs("sClientID")
            oDriver.LoadRecord
         Else
            lsBrowse = KwikBrowse(oApp, lrs, _
                                 "sClientID»xFullName»xAddressx", _
                                 oDriver.BrowseFTitle(0) & "»" & _
                                 oDriver.BrowseFTitle(1) & "»" & _
                                 oDriver.BrowseFTitle(2))
            If lsBrowse <> "" Then
               lsSelected = Split(lsBrowse, "»")
               oDriver.LookupValue(0) = lsSelected(0)
               oDriver.LoadRecord
            Else
               If .Tag <> "" Then .Text = .Tag
            End If
            
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
         End If
         .Tag = .Text
      End With
   End If
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " _
                       & "  " & lsSQL _
                       & " )", True
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
