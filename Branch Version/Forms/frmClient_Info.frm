VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmClient_Info 
   BorderStyle     =   0  'None
   Caption         =   "Client Maintenance"
   ClientHeight    =   6495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10380
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6495
   ScaleWidth      =   10380
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   4395
      Index           =   0
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   1095
      Width           =   10200
      _ExtentX        =   17992
      _ExtentY        =   7752
      BorderStyle     =   1
      Begin VB.TextBox txtOther 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   5070
         TabIndex        =   48
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   3090
         Width           =   1920
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   13
         Left            =   6285
         TabIndex        =   36
         Text            =   "Text1"
         Top             =   2280
         Width           =   3735
      End
      Begin VB.OptionButton optGender 
         Caption         =   "Female"
         Height          =   315
         Index           =   1
         Left            =   2040
         TabIndex        =   14
         Tag             =   "wt0;fb0"
         Top             =   1620
         Width           =   840
      End
      Begin VB.OptionButton optGender 
         Caption         =   "Male"
         Height          =   315
         Index           =   0
         Left            =   1185
         TabIndex        =   13
         Tag             =   "wt0;fb0"
         Top             =   1605
         Width           =   660
      End
      Begin VB.ComboBox cmbField 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "frmClient_Info.frx":0000
         Left            =   1185
         List            =   "frmClient_Info.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1950
         Width           =   2070
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   14
         Left            =   6285
         TabIndex        =   38
         Text            =   "Text1"
         Top             =   2550
         Width           =   2070
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   945
         Index           =   12
         Left            =   6285
         MultiLine       =   -1  'True
         TabIndex        =   34
         Text            =   "frmClient_Info.frx":0039
         Top             =   1305
         Width           =   3735
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   11
         Left            =   6285
         TabIndex        =   32
         Text            =   "Text1"
         Top             =   1035
         Width           =   2070
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   10
         Left            =   6285
         TabIndex        =   30
         Text            =   "Text1"
         Top             =   765
         Width           =   2070
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   9
         Left            =   1185
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   3975
         Width           =   3735
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   8
         Left            =   1185
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   3705
         Width           =   3735
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   6
         Left            =   1185
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   2820
         Width           =   3735
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   5
         Left            =   1185
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   2550
         Width           =   2070
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   4
         Left            =   1185
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   2280
         Width           =   2070
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   585
         Index           =   7
         Left            =   1185
         MultiLine       =   -1  'True
         TabIndex        =   24
         Text            =   "frmClient_Info.frx":003F
         Top             =   3090
         Width           =   3735
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   3
         Left            =   1185
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   1305
         Width           =   3735
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   2
         Left            =   1185
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   1035
         Width           =   3735
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   1185
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   765
         Width           =   3735
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   0
         Left            =   1185
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   135
         Width           =   1920
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email Add"
         Height          =   195
         Index           =   12
         Left            =   5085
         TabIndex        =   29
         Top             =   765
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TIN"
         Height          =   195
         Index           =   13
         Left            =   5085
         TabIndex        =   31
         Top             =   1050
         Width           =   270
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Additional Info"
         Height          =   195
         Index           =   14
         Left            =   5085
         TabIndex        =   33
         Top             =   1290
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company Name"
         Height          =   195
         Index           =   15
         Left            =   5085
         TabIndex        =   35
         Top             =   2295
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client #"
         Height          =   195
         Index           =   16
         Left            =   5085
         TabIndex        =   37
         Top             =   2565
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone #"
         Height          =   195
         Index           =   11
         Left            =   165
         TabIndex        =   27
         Top             =   3990
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Town"
         Height          =   195
         Index           =   8
         Left            =   165
         TabIndex        =   25
         Top             =   3720
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gender"
         Height          =   195
         Index           =   22
         Left            =   165
         TabIndex        =   12
         Top             =   1650
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Birth Place"
         Height          =   195
         Index           =   7
         Left            =   165
         TabIndex        =   21
         Top             =   2835
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Birth Date"
         Height          =   195
         Index           =   6
         Left            =   165
         TabIndex        =   19
         Top             =   2565
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Citizenship"
         Height          =   195
         Index           =   5
         Left            =   165
         TabIndex        =   17
         Top             =   2295
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Civil Status"
         Height          =   195
         Index           =   4
         Left            =   165
         TabIndex        =   15
         Top             =   1995
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   225
         Index           =   3
         Left            =   165
         TabIndex        =   23
         Top             =   3105
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "First Name"
         Height          =   195
         Index           =   10
         Left            =   165
         TabIndex        =   8
         Top             =   1050
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name"
         Height          =   195
         Index           =   2
         Left            =   165
         TabIndex        =   6
         Top             =   765
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Middle Name"
         Height          =   195
         Index           =   1
         Left            =   165
         TabIndex        =   10
         Top             =   1290
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client ID"
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   4
         Top             =   150
         Width           =   600
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   240
         Left            =   1235
         Tag             =   "et0;ht2"
         Top             =   165
         Width           =   1920
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   0
      Left            =   9540
      TabIndex        =   46
      Top             =   5700
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
      Picture         =   "frmClient_Info.frx":0045
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   1
      Left            =   8775
      TabIndex        =   45
      Top             =   5700
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
      Picture         =   "frmClient_Info.frx":07BF
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   2
      Left            =   7245
      TabIndex        =   40
      Top             =   5700
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
      Picture         =   "frmClient_Info.frx":0F39
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   4
      Left            =   6480
      TabIndex        =   39
      Top             =   5700
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
      Picture         =   "frmClient_Info.frx":16B3
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   7
      Left            =   8010
      TabIndex        =   42
      Top             =   5700
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
      Picture         =   "frmClient_Info.frx":1E2D
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   8
      Left            =   5715
      TabIndex        =   44
      Top             =   5700
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
      Picture         =   "frmClient_Info.frx":25A7
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   3
      Left            =   7245
      TabIndex        =   41
      Top             =   5700
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
      Picture         =   "frmClient_Info.frx":2D21
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   6
      Left            =   8010
      TabIndex        =   43
      Top             =   5700
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
      Picture         =   "frmClient_Info.frx":349B
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   5
      Left            =   9540
      TabIndex        =   47
      Top             =   5700
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
      Picture         =   "frmClient_Info.frx":3C15
      PicturePos      =   1
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   495
      Index           =   1
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   10200
      _ExtentX        =   17992
      _ExtentY        =   873
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtOther 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   4
         Left            =   4905
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   105
         Width           =   5160
      End
      Begin VB.TextBox txtOther 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   3
         Left            =   1185
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   105
         Width           =   1920
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Client ID"
         Height          =   285
         Index           =   21
         Left            =   165
         TabIndex        =   0
         Top             =   105
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Full Name"
         Height          =   285
         Index           =   20
         Left            =   4080
         TabIndex        =   2
         Top             =   105
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmClient_Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents oDriver As FormDriver
Attribute oDriver.VB_VarHelpID = -1
Private oLedger As frmClient_Ledger
Private oSkin As FormSkin
Private bLoaded As Boolean

Dim pnindex As Integer

Private Sub cmdButton_Click(Index As Integer)
   Select Case Index
   Case 0
      oDriver.RecordCancelUpdate
   Case 1
      BrowseRecord True
   Case 2
      oDriver.RecordSave
   Case 3
      oDriver.RecordUpdate
      txtField(1).SetFocus
   Case 4
      oDriver.RecordNew
   Case 5
      Unload Me
   Case 6
      MsgBox "Delete Not Permitted!!!" & vbCrLf & vbCrLf & _
      "Please Notify ROSALYN LAZO DESCALLAR" & vbCrLf & _
      "for Assistance!!!", vbCritical, "Warning"
'      oDriver.RecordDelete
   Case 7
      oDriver.RecordSearch
   Case 8
      If txtField(0).Text <> "" Then
         oLedger.ClientID = oDriver.FieldValue(0)
'         Load oLedger
         If oLedger.BrowseLedger Then
            oLedger.Show 1
         Else
            MsgBox "No Ledger found!!!", vbCritical, "Warning"
         End If
      End If
   End Select
End Sub

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0

   If bLoaded Then
      oDriver.RecordNew
      oDriver.DisableTextbox 0
      bLoaded = False
      oDriver.HideButton 3
      oDriver.HideButton 6
      oDriver.HideButton 8
   End If
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
   End Select
End Sub

Private Sub Form_Load()
   CenterChildForm mdiMain, Me
   bLoaded = True
   
   Set oLedger = New frmClient_Ledger
   
   Set oDriver = New FormDriver
   Set oDriver.AppDriver = oApp
   Set oDriver.MainForm = Me
   
   Set oSkin = New FormSkin
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
                        & ", cRecdStat" _
                        & ", sModified" _
                        & ", dModified" _
                        & ", vTimeStmp" _
                     & " FROM Client_Master"
   
   oDriver.BrowseQuery = "SELECT " _
                           & " sClientID, " _
                           & " sLastName +  ', ' + sFrstName + ' ' + sMiddName as FullName " _
                        & " FROM Client_Master" _
                        & " ORDER BY sClientID"
   
   oDriver.InitRecForm
   
   oDriver.BrowseFTitle(0) = "Client ID"
   oDriver.BrowseFTitle(1) = "Full Name"

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
   oDriver.LookupColumn(8) = "sTownName»sProvName»sZippCode"
   oDriver.LookupTitle(8) = "Town»Province»ZippCode"
   
   oDriver.FieldFormat(5) = "MMMM DD, YYYY"
   oDriver.FieldStart = 1
   
   ClearFields
End Sub

Private Sub ClearFields()
   Dim lnCtr As Integer
   
   For lnCtr = 0 To txtField.Count - 1
      txtField(lnCtr).Text = ""
   Next
   
   txtOther(3).Text = ""
   txtOther(4).Text = ""
   txtOther(3).Tag = ""
   txtOther(4).Tag = ""
   
   optGender(0).Value = False
   optGender(1).Value = False
   
   cmbField.ListIndex = -1
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
   Set oDriver = Nothing
   Set oLedger = Nothing
End Sub

Private Sub oDriver_DisableOtherControl()
   txtOther(3).Enabled = True
   txtOther(4).Enabled = True
   
   optGender(0).Enabled = False
   optGender(1).Enabled = False
   txtOther(1).Visible = False
   cmbField.Enabled = False
   
End Sub

Private Sub oDriver_EnableOtherControl()
   oDriver.DisableTextbox 0
   txtOther(3).Enabled = False
   txtOther(4).Enabled = False

   optGender(0).Enabled = True
   optGender(1).Enabled = True
   txtOther(1).Visible = False
   cmbField.Enabled = True
End Sub

Private Sub oDriver_InitValue()
   oDriver.FieldValue(0) = getNextCode("Client_Master", "sClientID", True, oApp.Connection, True, oApp.BranchCode)
   txtField(0) = oDriver.FieldValue(0)
End Sub

Private Sub oDriver_LoadOtherData()
   Dim lrs As ADODB.Recordset
   Dim lnCtr As Integer
   Dim lsSQL As String
   
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
      
   txtOther(3).Text = oDriver.FieldValue(0)
   txtOther(3).Tag = oDriver.FieldValue(0)
   txtOther(4).Text = oDriver.FieldValue(1) & ", " & oDriver.FieldValue(2) & " " & oDriver.FieldValue(3)
End Sub

Private Sub oDriver_WillSave(Cancel As Boolean)
   If optGender(0).Value Or optGender(1).Value Then
      oDriver.FieldValue(15) = IIf(optGender(0).Value, 0, 1)
   End If
   oDriver.FieldValue(18) = 1
   If cmbField.ListIndex > -1 Then oDriver.FieldValue(16) = cmbField.ListIndex
End Sub

Private Function BrowseRecord(ByVal Search As Boolean, Optional lsValue As Variant, Optional lbCode As Variant)
Dim lrs As ADODB.Recordset
Dim lsSQL As String
Dim lsSearch As String
Dim lsSelected() As String

On Error GoTo errProc
   
   lsSQL = "Select" _
               & "  a.sClientID" _
               & ", a.sLastName + ', ' + a.sFrstName + ' ' + a.sMiddName as xFullName" _
               & ", a.sAddressx + ', ' + b.sTownName + ', ' + c.sProvName + ' ' + b.sZippCode xAddressx" _
            & " From Client_Master a" _
               & " Left Join TownCity b" _
                  & " Left Join Province c" _
                  & " On b.sProvIDxx = c.sProvIDxx" _
                     & " On a.sTownIDxx = b.sTownIDxx"
   
   If Not IsMissing(lsValue) Then
      Select Case lbCode
      Case 0
         lsSQL = lsSQL & " Where a.sClientID = " & strParm(lsValue)
      Case 1
         lsSQL = lsSQL & " Where a.sLastName + ', ' + a.sFrstName + ' ' + a.sMiddName = " & strParm(lsValue)
      Case 2
         lsSQL = lsSQL & " Where a.sLastName + ', ' + a.sFrstName + ' ' + a.sMiddName Like " & strParm(lsValue & "%")
      End Select
   End If
   Set lrs = New ADODB.Recordset
   lrs.Open lsSQL, oApp.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
   
   If lrs.EOF Then
      ClearFields
   ElseIf lrs.RecordCount = 1 Then
      oDriver.LookupValue(0) = lrs("sClientID")
      oDriver.LoadRecord
   Else
      lsSearch = KwikSearch(oApp _
                           , lsSQL _
                           , "sClientID»xFullName»xAddressx" _
                           , "ClientID»FullName»Address" _
                           , "@»@»@" _
                           , Search _
                           , "a.sClientID»a.sLastName + ', ' + a.sFrstName + ' ' + a.sMiddName»a.sAddressx + ', ' + b.sTownName + ', ' + c.sProvName + ' ' + b.sZippCode")
      
      If lsSearch <> "" Then
         lsSelected = Split(lsSearch, "»")
         oDriver.LookupValue(0) = lsSelected(0)
         oDriver.LoadRecord
      End If
   End If
   
   txtOther(pnindex).SelStart = 3
   txtOther(pnindex).SelLength = Len(txtOther(pnindex).Text)
   
   txtField(3).Tag = txtField(3).Text
   txtField(4).Tag = txtField(4).Text
   
endProc:
   lrs.Close
   Exit Function
errProc:
   MsgBox Err.Description, vbCritical, "Warning"
End Function

Private Sub txtField_LostFocus(Index As Integer)
   txtField(Index).BackColor = &H80000005
End Sub

Private Sub txtfield_Validate(Index As Integer, Cancel As Boolean)
   txtField(Index).Text = TitleCase(txtField(Index).Text)
   Cancel = Not oDriver.ValidateField(Index)
   If Index = 7 Or Index = 12 Then oDriver.FieldValue(Index) = Replace(oDriver.FieldValue(Index), vbCrLf, " ")
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Or KeyCode = 13 Then
      If Index = 8 Then
         oDriver.RecordSearch txtField(Index).Text
         SetNextFocus
      End If
   KeyCode = 0
  End If
End Sub

Private Sub txtOther_GotFocus(Index As Integer)
   If Index = 3 Or Index = 4 Then
      pnindex = Index
      If txtOther(Index).Text <> "" Then
         txtOther(Index).SelStart = 0
         txtOther(Index).SelLength = Len(txtOther(Index).Text)
      End If
      txtOther(Index).BackColor = &HE1FEFF
   End If
End Sub

Private Sub txtOther_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
      If Index = 4 And Trim(txtOther(Index).Text) <> "" Then BrowseRecord IIf(Trim(txtOther(Index).Text) = "", True, False), txtOther(Index).Text, 4
      SetNextFocus
      KeyCode = 0
   End If
End Sub

Private Sub txtOther_LostFocus(Index As Integer)
If Index = 3 Or Index = 4 Then
   txtOther(Index).Text = TitleCase(txtOther(Index).Text)
   txtOther(Index).BackColor = &H80000005
End If
End Sub

Private Sub txtOther_Validate(Index As Integer, Cancel As Boolean)
   If Index = 3 Or Index = 4 Then
      If txtOther(Index).Text <> txtOther(Index).Tag Then
         BrowseRecord False, txtOther(Index).Text
      End If
   End If
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   oDriver.ColumnIndex = Index
   pnindex = Index
   txtField(Index).BackColor = &HE1FEFF
End Sub
