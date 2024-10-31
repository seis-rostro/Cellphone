VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmBanks 
   BorderStyle     =   0  'None
   Caption         =   "Banks Maintenance"
   ClientHeight    =   4380
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5400
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4380
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   2805
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   4948
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   6
         Left            =   1425
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   2280
         Width           =   3570
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   5
         Left            =   1425
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   1950
         Width           =   3570
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   4
         Left            =   1425
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   1620
         Width           =   3570
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   1425
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   630
         Width           =   3570
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
         Left            =   1425
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   195
         Width           =   900
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   3
         Left            =   1425
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   1290
         Width           =   3570
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   1425
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   960
         Width           =   3570
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fax No"
         Height          =   195
         Index           =   6
         Left            =   210
         TabIndex        =   21
         Top             =   2340
         Width           =   810
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tel. No."
         Height          =   195
         Index           =   5
         Left            =   210
         TabIndex        =   19
         Top             =   2010
         Width           =   810
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Town/City"
         Height          =   195
         Index           =   4
         Left            =   210
         TabIndex        =   17
         Top             =   1680
         Width           =   810
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   315
         Left            =   1500
         Tag             =   "et0;ht2"
         Top             =   285
         Width           =   900
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Person"
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   4
         Top             =   1035
         Width           =   1110
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   195
         Index           =   3
         Left            =   210
         TabIndex        =   6
         Top             =   1350
         Width           =   810
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Name"
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   2
         Top             =   705
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
         Left            =   210
         TabIndex        =   0
         Top             =   240
         Width           =   810
      End
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   705
      Index           =   0
      Left            =   4545
      TabIndex        =   15
      Top             =   3600
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
      Picture         =   "frmBanks.frx":0000
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   705
      Index           =   1
      Left            =   3765
      TabIndex        =   13
      Top             =   3600
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
      Picture         =   "frmBanks.frx":077A
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   705
      Index           =   2
      Left            =   2205
      TabIndex        =   10
      Top             =   3600
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
      Picture         =   "frmBanks.frx":0EF4
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   705
      Index           =   3
      Left            =   2205
      TabIndex        =   9
      Top             =   3600
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
      Picture         =   "frmBanks.frx":166E
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   705
      Index           =   4
      Left            =   1425
      TabIndex        =   8
      Top             =   3600
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
      Picture         =   "frmBanks.frx":1DE8
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   705
      Index           =   5
      Left            =   4545
      TabIndex        =   14
      Top             =   3600
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
      Picture         =   "frmBanks.frx":2562
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   705
      Index           =   6
      Left            =   2985
      TabIndex        =   12
      Top             =   3600
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
      Picture         =   "frmBanks.frx":2CDC
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   705
      Index           =   7
      Left            =   2985
      TabIndex        =   11
      Top             =   3600
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
      Picture         =   "frmBanks.frx":3456
      PicturePos      =   1
   End
End
Attribute VB_Name = "frmBanks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmBanks"

Private WithEvents oDriver As clsFormDriver
Attribute oDriver.VB_VarHelpID = -1
Private oSkin As clsFormSkin
Private bLoaded As Boolean

Dim pnIndex As Integer

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
   Case 4
      oDriver.RecordNew
   Case 5
      Unload Me
   Case 6
      oDriver.RecordDelete
   Case 7
      oDriver.RecordSearch
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
   
   If bLoaded = False Then
      oDriver.RecordNew
      oDriver.DisableTextbox 0
      bLoaded = True
   End If

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String
   
   lsOldProc = "Form_Load"
   On Error GoTo errProc
   
   CenterChildForm mdiMain, Me
   
   bLoaded = False
   
   Set oDriver = New clsFormDriver
   Set oDriver.AppDriver = oApp
   Set oDriver.MainForm = Me
   
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin
   
   oDriver.RecQuery = "SELECT * FROM Banks"
   oDriver.BrowseQuery = "SELECT" _
                           & "  a.sBankIDxx" _
                           & ", a.sBankName" _
                           & ", CONCAT(a.sAddressx, ', ', b.sTownName, ', ', c.sProvName, ' ', b.sZippCode) xAddressx" _
                        & " FROM Banks a" _
                           & " LEFT Join TownCity b" _
                              & " ON a.sTownIDxx = b.sTownIDxx" _
                           & " LEFT Join Province c" _
                              & " ON b.sProvIDxx = c.sProvIDxx" _
                        & " WHERE a.cRecdStat = " & strParm(xeRecStateActive) _
                        & " ORDER BY a.sBankName"
   
   oDriver.InitRecForm
   
   oDriver.BrowseFTitle(0) = "Code"
   oDriver.BrowseFTitle(1) = "Bank"
   oDriver.BrowseFTitle(2) = "Address"
   
   oDriver.LookupQuery(4) = "SELECT" _
                              & "  a.sTownIDxx" _
                              & ", CONCAT(a.sTownName, ', ', b.sProvName, ' ', a.sZippCode) xTownName" _
                           & " FROM TownCity a" _
                              & " LEFT Join Province b" _
                                 & " On a.sProvIDxx = b.sProvIDxx" _
                           & " WHERE a.cRecdStat = " & strParm(xeRecStateActive) _
                           & " ORDER BY a.sTownName"
   
   oDriver.LookupReference(4) = "sTownIDxx�xTownName"
   oDriver.LookupColumn(4) = "xTownName"
   oDriver.LookupTitle(4) = "TownName"
   
   oDriver.FieldFormat(0) = "@@-@@@@@"
   oDriver.FieldSize(0) = Len(oDriver.FieldFormat(0))
   oDriver.FieldStart = 1

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oDriver = Nothing
   Set oSkin = Nothing
End Sub

Private Sub oDriver_EnableOtherControl()
   oDriver.DisableTextbox 0
End Sub

Private Sub oDriver_InitValue()
   Dim lsOldProc As String
   
   lsOldProc = "oDriver_InitValue"
   On Error GoTo errProc

   If oDriver.SetValue(0, GetNextCode("Banks", "sBankIDxx", False, oApp.Connection, True, oApp.BranchCode)) = False Then Exit Sub
   oDriver.FieldReference(0) = True
   oDriver.FieldValue(7) = 1
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )"
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String
   
   lsOldProc = "txtField_KeyDown"
   On Error GoTo errProc
   
   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
      With txtField(pnIndex)
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
   oDriver.ColumnIndex = Index
   With txtField(Index)
      .BackColor = oApp.getColor("HT1")
   End With
   pnIndex = Index
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
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

Private Sub oDriver_WillSave(Cancel As Boolean)
   If oDriver.FieldValue(1) = "" Then
      MsgBox "Invalid Bank Name Detected!!!" & vbCrLf & _
             "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
      txtField(1).SetFocus
      Cancel = True
   ElseIf oDriver.FieldValue(4) = "" Then
      MsgBox "Invalid Town/City Detected!!!" & vbCrLf & _
             "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
      txtField(4).SetFocus
      Cancel = True
   End If
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   Dim lsOldProc As String
   
   lsOldProc = "txtField_Validate"
   On Error GoTo errProc
   
   txtField(Index).Text = TitleCase(txtField(Index).Text)
   Cancel = Not oDriver.ValidateField(Index)
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " _
                       & "  " & Index _
                       & ", " & Cancel _
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
