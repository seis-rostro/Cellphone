VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmCustomerAccessories 
   BorderStyle     =   0  'None
   Caption         =   "Accessories  Maintenance"
   ClientHeight    =   2955
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4740
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2955
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1365
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   2408
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   1560
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   720
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
         Height          =   315
         Index           =   0
         Left            =   1560
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   285
         Width           =   1425
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   315
         Left            =   1650
         Tag             =   "et0;ht2"
         Top             =   375
         Width           =   1425
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Accessories ID"
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
         Left            =   195
         TabIndex        =   0
         Top             =   315
         Width           =   1365
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   1
         Left            =   195
         TabIndex        =   2
         Top             =   780
         Width           =   1065
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   0
      Left            =   3900
      TabIndex        =   9
      Top             =   2175
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
      Picture         =   "frmCustomerAccessories.frx":0000
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   1
      Left            =   3120
      TabIndex        =   8
      Top             =   2175
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
      Picture         =   "frmCustomerAccessories.frx":077A
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   2
      Left            =   2340
      TabIndex        =   6
      Top             =   2175
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
      Picture         =   "frmCustomerAccessories.frx":0EF4
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   3
      Left            =   1560
      TabIndex        =   5
      Top             =   2175
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
      Picture         =   "frmCustomerAccessories.frx":166E
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   4
      Left            =   780
      TabIndex        =   4
      Top             =   2175
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
      Picture         =   "frmCustomerAccessories.frx":1DE8
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   5
      Left            =   3900
      TabIndex        =   10
      Top             =   2175
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
      Picture         =   "frmCustomerAccessories.frx":2562
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   6
      Left            =   2340
      TabIndex        =   7
      Top             =   2175
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
      Picture         =   "frmCustomerAccessories.frx":2CDC
      PicturePos      =   1
   End
End
Attribute VB_Name = "frmCustomerAccessories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmCustomerAccessories"

Private WithEvents oDriver As clsFormDriver
Attribute oDriver.VB_VarHelpID = -1
Private oSkin As clsFormSkin
Private bLoaded As Boolean

Dim pnIndex As Integer

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   
   lsOldProc = "cmdButton_Click"
   'On Error GoTo errProc
   
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
   'On Error GoTo errProc
   
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
   'On Error GoTo errProc
   
   CenterChildForm mdiMain, Me
   
   bLoaded = False
   
   Set oDriver = New clsFormDriver
   Set oDriver.AppDriver = oApp
   Set oDriver.MainForm = Me
   
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin
   
   oDriver.RecQuery = "SELECT" _
                           & "  sAccessID" _
                           & ", sDescript" _
                           & ", cRecdStat" _
                           & ", sModified" _
                           & ", dModified" _
                        & " FROM CP_Customer_Accessories"
   oDriver.BrowseQuery = "SELECT" _
                           & "  sAccessID" _
                           & ", sDescript" _
                        & " FROM CP_Customer_Accessories" _
                        & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
                        & " ORDER BY sAccessID"
   
   oDriver.InitRecForm
   
   oDriver.BrowseFTitle(0) = "Code"
   oDriver.BrowseFTitle(1) = "Accessories"
   oDriver.BrowseFFormat(0) = "@@-@@@@@"
   
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
   'On Error GoTo errProc
   
   If oDriver.SetValue(0, GetNextCode("CP_Customer_Accessories", "sAccessID", False, oApp.Connection, True, oApp.BranchCode)) = False Then Exit Sub
   oDriver.FieldReference(0) = True
   oDriver.FieldValue(2) = 1

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )"
End Sub

Private Sub oDriver_WillSave(Cancel As Boolean)
   If oDriver.FieldValue(1) = "" Then
      MsgBox "Invalid Description detected!!!", vbCritical, "Warning"
      txtField(1).SetFocus
      Cancel = True
   End If
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

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   Dim lsOldProc As String
   
   lsOldProc = "txtField_Validate"
   'On Error GoTo errProc
   
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