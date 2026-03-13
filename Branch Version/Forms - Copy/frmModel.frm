VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmModel 
   BorderStyle     =   0  'None
   Caption         =   "Model Maintenance"
   ClientHeight    =   2850
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4365
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   2850
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   0
      Left            =   3525
      TabIndex        =   12
      Top             =   2070
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
      Picture         =   "frmModel.frx":0000
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   1
      Left            =   2760
      TabIndex        =   11
      Top             =   2070
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
      Picture         =   "frmModel.frx":077A
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   2
      Left            =   1230
      TabIndex        =   7
      Top             =   2070
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
      Picture         =   "frmModel.frx":0EF4
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   3
      Left            =   1230
      TabIndex        =   8
      Top             =   2070
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
      Picture         =   "frmModel.frx":166E
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   4
      Left            =   465
      TabIndex        =   6
      Top             =   2070
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
      Picture         =   "frmModel.frx":1DE8
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   5
      Left            =   3525
      TabIndex        =   13
      Top             =   2070
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
      Picture         =   "frmModel.frx":2562
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   6
      Left            =   1995
      TabIndex        =   9
      Top             =   2070
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
      Picture         =   "frmModel.frx":2CDC
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   705
      Index           =   7
      Left            =   1995
      TabIndex        =   10
      Top             =   2070
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
      Picture         =   "frmModel.frx":3456
      PicturePos      =   1
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1320
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   4185
      _ExtentX        =   7382
      _ExtentY        =   2328
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   2
         Left            =   1230
         TabIndex        =   3
         Top             =   690
         Width           =   2760
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   1230
         TabIndex        =   5
         Top             =   945
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
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   240
         Left            =   1275
         Tag             =   "et0;ht2"
         Top             =   210
         Width           =   900
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Brand Name"
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   2
         Top             =   720
         Width           =   810
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Model Name"
         Height          =   195
         Index           =   1
         Left            =   165
         TabIndex        =   4
         Top             =   960
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Model ID"
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
   End
End
Attribute VB_Name = "frmModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents oDriver As FormDriver
Attribute oDriver.VB_VarHelpID = -1
Private oSkin As FormSkin
Private bLoaded As Boolean

Dim pbLoading As Boolean
Dim pnindex As Integer

Private Sub cmdButton_Click(Index As Integer)
   Select Case Index
   Case 0
      pbLoading = True
      oDriver.RecordCancelUpdate
   Case 1
      oDriver.BrowseRecord
   Case 2
      oDriver.RecordSave
   Case 3
      pbLoading = False
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
      oDriver.RecordSearch
      txtfield(pnindex).SetFocus
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
                        & "sModelIDx, " _
                        & "sModelNme, " _
                        & "sBrandIDx, " _
                        & "cRecdStat, " _
                        & "sModified, " _
                        & "dModified, " _
                        & "vTimeStmp  " _
                    & "FROM Model "
                    
   oDriver.BrowseQuery = "SELECT" _
                       & " a.sModelIDx, " _
                       & " a.sModelNme, " _
                       & " b.sBrandNme  " _
                    & " FROM Model a, " _
                       & " Brand b " _
                    & " WHERE a.cRecdStat = " & strParm(xeRecStateActive) _
                        & " AND a.sBrandIDx = b.sBrandIDx" _
                    & " ORDER BY a.sModelIDx "
   
   oDriver.InitRecForm

   oDriver.BrowseFTitle(0) = "Model ID"
   oDriver.BrowseFTitle(1) = "Model"
   oDriver.BrowseFTitle(2) = "Brand"
   
   oDriver.LookupQuery(2) = "SELECT" _
                              & "  sBrandIDx" _
                              & ", sBrandNme" _
                           & " FROM Brand " _
                           & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
                           & " ORDER BY sBrandNme"
   
    oDriver.LookupReference(2) = "sBrandIDx»sBrandNme"
    oDriver.LookupColumn(2) = "sBrandNme"
    oDriver.LookupTitle(2) = "Brand"
   
    oDriver.FieldFormat(0) = "@@-@@@@@"
    oDriver.FieldSize(0) = Len(oDriver.FieldFormat(0))
    oDriver.FieldStart = 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oDriver = Nothing
   Set oSkin = Nothing
End Sub

Private Sub oDriver_EnableOtherControl()
   oDriver.DisableTextbox 0
End Sub

Private Sub oDriver_InitValue()
   If oDriver.SetValue(0, getNextCode("Model", "sModelIDx", False, oApp.Connection, True, oApp.BranchCode)) = False Then Exit Sub
   oDriver.FieldReference(0) = True
   oDriver.FieldValue(3) = xeRecStateActive
   pbLoading = False
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   If pbLoading Then Exit Sub
   oDriver.ColumnIndex = Index
   pnindex = Index
   txtfield(Index).BackColor = &HE1FEFF
   txtfield(Index).Text = TitleCase(txtfield(Index).Text)
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Or KeyCode = 13 Then
      oDriver.RecordSearch txtfield(2).Text
      If txtfield(Index).Text <> "" Then SetNextFocus
      KeyCode = 0
   End If
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
      MsgBox "Invalid Model Name detected!!!", vbCritical, "Warning"
      txtfield(1).SetFocus
      Cancel = True
   ElseIf oDriver.FieldValue(2) = "" Then
      MsgBox "Invalid Brand detected!!!", vbCritical, "Warning"
      txtfield(2).SetFocus
      Cancel = True
   End If
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   txtfield(Index).BackColor = &H80000005
End Sub

Private Sub txtfield_Validate(Index As Integer, Cancel As Boolean)
   With txtfield(Index)
      If Trim(.Text) <> "" Then .Text = UCase(Left(.Text, 1)) & Right(.Text, Len(.Text) - 1)
   End With
   
   Cancel = Not oDriver.ValidateField(Index)
End Sub


