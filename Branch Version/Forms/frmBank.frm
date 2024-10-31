VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmBank 
   BorderStyle     =   0  'None
   Caption         =   "Bank Maintenance"
   ClientHeight    =   3360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8310
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3360
   ScaleWidth      =   8310
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   2
      Left            =   5205
      TabIndex        =   15
      Top             =   2580
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
      Picture         =   "frmBank.frx":0000
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   0
      Left            =   7500
      TabIndex        =   20
      Top             =   2580
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
      Picture         =   "frmBank.frx":077A
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   1
      Left            =   6735
      TabIndex        =   19
      Top             =   2580
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
      Picture         =   "frmBank.frx":0EF4
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   3
      Left            =   5205
      TabIndex        =   16
      Top             =   2580
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
      Picture         =   "frmBank.frx":166E
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   4
      Left            =   4440
      TabIndex        =   14
      Top             =   2580
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
      Picture         =   "frmBank.frx":1DE8
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   5
      Left            =   7500
      TabIndex        =   21
      Top             =   2580
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
      Picture         =   "frmBank.frx":2562
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   6
      Left            =   5970
      TabIndex        =   17
      Top             =   2580
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
      Picture         =   "frmBank.frx":2CDC
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   705
      Index           =   7
      Left            =   5970
      TabIndex        =   18
      Top             =   2580
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
      Picture         =   "frmBank.frx":3456
      PicturePos      =   1
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1830
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   3228
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   6
         Left            =   5235
         TabIndex        =   13
         Top             =   1140
         Width           =   2760
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   5
         Left            =   5235
         TabIndex        =   11
         Top             =   885
         Width           =   2760
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   4
         Left            =   5235
         TabIndex        =   9
         Top             =   630
         Width           =   2760
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   3
         Left            =   1230
         TabIndex        =   7
         Top             =   1410
         Width           =   2760
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
         Left            =   1230
         TabIndex        =   1
         Top             =   165
         Width           =   900
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   1230
         TabIndex        =   3
         Top             =   630
         Width           =   2760
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   510
         Index           =   2
         Left            =   1230
         TabIndex        =   5
         Top             =   885
         Width           =   2760
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fax No."
         Height          =   195
         Index           =   6
         Left            =   4065
         TabIndex        =   12
         Top             =   1155
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Telephon No."
         Height          =   195
         Index           =   5
         Left            =   4065
         TabIndex        =   10
         Top             =   900
         Width           =   1185
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Person"
         Height          =   195
         Index           =   4
         Left            =   4065
         TabIndex        =   8
         Top             =   645
         Width           =   1290
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Town"
         Height          =   195
         Index           =   3
         Left            =   165
         TabIndex        =   6
         Top             =   1425
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bank ID"
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
         TabIndex        =   0
         Top             =   195
         Width           =   810
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Name"
         Height          =   195
         Index           =   1
         Left            =   165
         TabIndex        =   2
         Top             =   645
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   4
         Top             =   900
         Width           =   810
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   240
         Left            =   1275
         Tag             =   "et0;ht2"
         Top             =   210
         Width           =   900
      End
   End
End
Attribute VB_Name = "frmBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  Programmed By  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  Rosalyn Lazo Descallar  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  Date Started  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  November 10, 2006  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'

Option Explicit

Private WithEvents oDriver As FormDriver
Attribute oDriver.VB_VarHelpID = -1
Private oSkin As FormSkin
Private bLoaded As Boolean

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
      oDriver.RecordSearch txtfield(3).Text
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
   
   oDriver.RecQuery = "SELECT sBankIDxx, " _
                        & " sBankName , " _
                        & " sAddressx ," _
                        & " sTownIdxx ," _
                        & " sContactP ," _
                        & " sTelNoxxx ," _
                        & " sFaxNoxxx ," _
                        & " cRecdStat ," _
                        & " sModified ," _
                        & " dModified ," _
                        & " vTimestmp  " _
                     & " FROM Banks "
   
   oDriver.BrowseQuery = "SELECT" _
                        & "  a.sBankIDxx, " _
                        & "  a.sBankName, " _
                        & " a.sAddressx + ' ' + b.sTownName xAddressx " _
                     & " FROM Banks a  " _
                        & " LEFT JOIN TownCity b " _
                           & " ON a.sTownIDxx = b.sTownIDxx " _
                     & " WHERE a.cRecdStat = " & strParm(xeRecStateActive) _
                     & " ORDER BY sBankIDxx"
                        
   oDriver.InitRecForm
   
   oDriver.BrowseColumn(0) = "sBankIDxx"
   oDriver.BrowseColumn(1) = "sBankName"
   oDriver.BrowseColumn(2) = "xAddressx"
   
   oDriver.BrowseFTitle(0) = "Bank ID"
   oDriver.BrowseFTitle(1) = "Bank Name"
   oDriver.BrowseFTitle(2) = "Address"
   oDriver.BrowseFFormat(0) = "@@-@@@"
   
   'Town
   oDriver.LookupQuery(3) = "SELECT" _
                           & "  a.sTownIDxx" _
                           & ", a.sTownName" _
                           & ", b.sProvName" _
                           & ", a.sZippCode" _
                        & " FROM TownCity a" _
                           & " LEFT JOIN Province b" _
                              & " ON a.sProvIDxx = b.sProvIDxx" _
                        & " WHERE a.cRecdStat = " & strParm(xeRecStateActive) _
                        & " ORDER BY a.sTownName,b.sProvName"

   oDriver.LookupReference(3) = "a.sTownIDxx»a.sTownName»b.sProvName»a.sZippCode"
   oDriver.LookupColumn(3) = "sTownName»sProvName»sZippCode"
   oDriver.LookupTitle(3) = "Town»Province»ZippCode"

   
   oDriver.FieldFormat(0) = "@@-@@@@@"
   oDriver.FieldSize(0) = Len(oDriver.FieldFormat(0))
   oDriver.FieldStart = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oDriver = Nothing
   Set oSkin = Nothing
End Sub

Private Sub oDriver_EnableOtherControl()
   oDriver.DisableTextbox 0
End Sub

Private Sub oDriver_InitValue()
   If oDriver.SetValue(0, getNextCode("Banks", "sBankIDxx", False, oApp.Connection, True, oApp.BranchCode)) = False Then Exit Sub
   oDriver.FieldReference(0) = True
   oDriver.FieldValue(7) = xeRecStateActive
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   oDriver.ColumnIndex = Index
   txtfield(Index).BackColor = &HE1FEFF
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = vbKeyF3 Then
   If Index = 3 Then
      oDriver.RecordSearch txtfield(Index).Text
   End If
   If txtfield(Index).Text <> "" Then SetNextFocus
End If
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   txtfield(Index).BackColor = &H80000005
End Sub

Private Sub txtfield_Validate(Index As Integer, Cancel As Boolean)
   If Index = 1 Then
      With txtfield(Index)
         If Trim(.Text) <> "" Then .Text = UCase(Left(.Text, 1)) & Right(.Text, Len(.Text) - 1)
      End With
   End If
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



