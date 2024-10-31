VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmSupplier 
   BorderStyle     =   0  'None
   Caption         =   "Supplier Maintenance"
   ClientHeight    =   3915
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6615
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   3915
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   2340
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   4128
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   6
         Left            =   1425
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   1875
         Width           =   2730
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   2
         Left            =   1425
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   855
         Width           =   4680
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   3
         Left            =   1425
         MaxLength       =   50
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   1110
         Width           =   4680
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   4
         Left            =   1425
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   1365
         Width           =   4680
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   1425
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   600
         Width           =   4680
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
         Height          =   240
         Index           =   0
         Left            =   1425
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   195
         Width           =   1425
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   5
         Left            =   1425
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   1620
         Width           =   2730
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fax No."
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   12
         Top             =   1905
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Person"
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   4
         Top             =   870
         Width           =   1125
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   240
         Left            =   1470
         Tag             =   "et0;ht2"
         Top             =   240
         Width           =   1425
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier ID"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   0
         Top             =   225
         Width           =   810
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Town"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   8
         Top             =   1410
         Width           =   810
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Name"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   615
         Width           =   1125
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   1140
         Width           =   810
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Contact No."
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   10
         Top             =   1650
         Width           =   1215
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   0
      Left            =   5760
      TabIndex        =   20
      Top             =   3120
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
      Picture         =   "frmSupplier.frx":0000
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   1
      Left            =   4995
      TabIndex        =   19
      Top             =   3120
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
      Picture         =   "frmSupplier.frx":077A
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   2
      Left            =   3465
      TabIndex        =   15
      Top             =   3120
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
      Picture         =   "frmSupplier.frx":0EF4
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   4
      Left            =   2700
      TabIndex        =   14
      Top             =   3120
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
      Picture         =   "frmSupplier.frx":166E
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   5
      Left            =   5760
      TabIndex        =   21
      Top             =   3120
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
      Picture         =   "frmSupplier.frx":1DE8
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   705
      Index           =   7
      Left            =   4230
      TabIndex        =   17
      Top             =   3120
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
      Picture         =   "frmSupplier.frx":2562
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   3
      Left            =   3465
      TabIndex        =   16
      Top             =   3120
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
      Picture         =   "frmSupplier.frx":2CDC
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   6
      Left            =   4230
      TabIndex        =   18
      Top             =   3120
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
      Picture         =   "frmSupplier.frx":3456
      PicturePos      =   1
   End
End
Attribute VB_Name = "frmSupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents oDriver As FormDriver
Attribute oDriver.VB_VarHelpID = -1
Private oSkin As FormSkin
Private bLoaded As Boolean
Dim pnindex As Integer

Private Sub cmdButton_Click(Index As Integer)
    Dim lnctr As Integer
   Select Case Index
   Case 0
      oDriver.RecordCancelUpdate
   Case 1
      oDriver.BrowseRecord
   Case 2
      oDriver.RecordSave
   Case 3
      oDriver.RecordUpdate
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
      If pnindex = 4 Then oDriver.RecordSearch txtfield(4).Text
   End Select
End Sub

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0

   If bLoaded = False Then
      oDriver.RecordNew
      oDriver.DisableTextbox 0
      bLoaded = True
   End If
End Sub

Private Sub Form_Load()
   CenterChildForm mdiMain, Me
   
   bLoaded = False
   
   Set oDriver = New FormDriver
   Set oDriver.AppDriver = oApp
   Set oDriver.MainForm = Me
   
   Set oSkin = New FormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin
   
   oDriver.RecQuery = "SELECT " _
                        & " sSupplyID, " _
                        & " sSupplyNm, " _
                        & " sCPersonx, " _
                        & " sAddressx, " _
                        & " sTownIDxx, " _
                        & " sTelNoxxx, " _
                        & " sFaxNoxxx, " _
                        & " cRecdStat, " _
                        & " sModified, " _
                        & " dModified, " _
                        & " vTimeStmp  " _
                        & " FROM Supplier  " _

   
   oDriver.BrowseQuery = "SELECT" _
                        & " sSupplyID, " _
                        & " sSupplyNm  " _
                        & " FROM Supplier  " _
                        & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
                        & " ORDER BY sSupplyID"
   
   oDriver.InitRecForm
   
   oDriver.BrowseFTitle(0) = "Supplier ID"
   oDriver.BrowseFTitle(1) = "Supplier Name "
   oDriver.BrowseFFormat(0) = "@@-@@@"
   
   oDriver.LookupQuery(4) = "SELECT" _
                        & "  a.sTownIDxx" _
                        & ", a.sTownName" _
                        & ", b.sProvName" _
                        & ", a.sZippCode" _
                     & " FROM TownCity a" _
                        & " LEFT JOIN Province b" _
                           & " ON a.sProvIDxx = b.sProvIDxx" _
                     & " WHERE a.cRecdStat = " & strParm(xeRecStateActive) _
                     & " ORDER BY a.sTownName,b.sProvName"

   oDriver.LookupReference(4) = "a.sTownIDxx»a.sTownName»b.sProvName»a.sZippCode"
   oDriver.LookupColumn(4) = "sTownName»sProvName»sZippCode"
   oDriver.LookupTitle(4) = "Town»Province»ZippCode"
   
   
   oDriver.FieldFormat(0) = "@@-@@@"
   oDriver.FieldSize(0) = Len(oDriver.FieldFormat(0))
   oDriver.FieldStart = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oDriver = Nothing
   Set oSkin = Nothing
End Sub

Private Sub oDriver_DisableOtherControl()
   oDriver.DisableTextbox 0
End Sub

Private Sub oDriver_EnableOtherControl()
   oDriver.DisableTextbox 0
End Sub

Private Sub oDriver_InitValue()
   If oDriver.SetValue(0, getNextCode("Supplier", "sSupplyID", False, oApp.Connection, True, oApp.BranchCode)) = False Then Exit Sub
   oDriver.FieldReference(0) = True
   oDriver.FieldValue(7) = xeRecStateActive
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   oDriver.ColumnIndex = Index
   txtfield(Index).BackColor = &HE1FEFF
   pnindex = Index
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = vbKeyF3 Then
   If Index = 4 Then
      oDriver.RecordSearch txtfield(Index).Text
      If txtfield(Index).Text <> "" Then SetNextFocus
   End If
End If
KeyCode = 0
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   txtfield(Index).BackColor = &H80000005
End Sub

Private Sub txtfield_Validate(Index As Integer, Cancel As Boolean)
   txtfield(Index).Text = TitleCase(txtfield(Index).Text)
   Cancel = Not oDriver.ValidateField(Index)
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
   End Select
End Sub




