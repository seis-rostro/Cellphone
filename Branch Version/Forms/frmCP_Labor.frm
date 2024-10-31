VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmCP_Labor 
   BorderStyle     =   0  'None
   Caption         =   "Labor Maintenance"
   ClientHeight    =   5220
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7125
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5220
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame2 
      Height          =   1950
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   2085
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   3440
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.OptionButton optField 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Specialized 2"
         Height          =   300
         Index           =   2
         Left            =   4515
         TabIndex        =   19
         Tag             =   "wt0;fb0"
         Top             =   1125
         Width           =   1530
      End
      Begin VB.OptionButton optField 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Specialized 1"
         Height          =   300
         Index           =   1
         Left            =   4515
         TabIndex        =   18
         Tag             =   "wt0;fb0"
         Top             =   870
         Width           =   1530
      End
      Begin VB.OptionButton optField 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Regular"
         Height          =   300
         Index           =   0
         Left            =   4515
         TabIndex        =   17
         Tag             =   "wt0;fb0"
         Top             =   615
         Width           =   1530
      End
      Begin VB.CheckBox chkField 
         BackColor       =   &H00C0C0C0&
         Caption         =   "In House"
         Height          =   195
         Left            =   1305
         TabIndex        =   15
         Tag             =   "wt0;fb0"
         Top             =   1530
         Width           =   2145
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   6
         Left            =   1305
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   1155
         Width           =   2325
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   5
         Left            =   1305
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   825
         Width           =   2325
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   4
         Left            =   1305
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   495
         Width           =   2325
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Labor Type"
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
         Left            =   4365
         TabIndex        =   16
         Tag             =   "wt0;fb0"
         Top             =   405
         Width           =   975
      End
      Begin VB.Shape Shape2 
         Height          =   975
         Left            =   4245
         Top             =   495
         Width           =   2235
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Level 03"
         Height          =   195
         Index           =   5
         Left            =   615
         TabIndex        =   13
         Top             =   1215
         Width           =   615
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Level 02"
         Height          =   195
         Index           =   4
         Left            =   615
         TabIndex        =   11
         Top             =   870
         Width           =   615
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Labor Prices"
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
         Left            =   150
         TabIndex        =   8
         Top             =   135
         Width           =   1080
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Level 01"
         Height          =   195
         Index           =   2
         Left            =   615
         TabIndex        =   9
         Top             =   540
         Width           =   615
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1515
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   2672
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   3
         Left            =   1155
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   1005
         Width           =   5580
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   1155
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   675
         Width           =   900
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   2070
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   675
         Width           =   4665
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
         Left            =   1155
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   150
         Width           =   1635
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Brand Name"
         Height          =   195
         Index           =   8
         Left            =   150
         TabIndex        =   6
         Top             =   1035
         Width           =   885
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   7
         Left            =   5925
         TabIndex        =   4
         Top             =   420
         Width           =   795
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   315
         Left            =   1215
         Tag             =   "et0;ht2"
         Top             =   225
         Width           =   1635
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Labor Code"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   2
         Top             =   705
         Width           =   825
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Labor ID"
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
         Left            =   150
         TabIndex        =   0
         Top             =   210
         Width           =   750
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   0
      Left            =   6270
      TabIndex        =   26
      Top             =   4425
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
      Picture         =   "frmCP_Labor.frx":0000
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   1
      Left            =   5490
      TabIndex        =   24
      Top             =   4425
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
      Picture         =   "frmCP_Labor.frx":077A
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   2
      Left            =   4710
      TabIndex        =   23
      Top             =   4425
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
      Picture         =   "frmCP_Labor.frx":0EF4
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   4
      Left            =   3930
      TabIndex        =   20
      Top             =   4425
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
      Picture         =   "frmCP_Labor.frx":166E
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   5
      Left            =   6270
      TabIndex        =   25
      Top             =   4425
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
      Picture         =   "frmCP_Labor.frx":1DE8
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   3
      Left            =   4710
      TabIndex        =   21
      Top             =   4425
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
      Picture         =   "frmCP_Labor.frx":2562
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   6
      Left            =   4710
      TabIndex        =   22
      Top             =   4425
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
      Picture         =   "frmCP_Labor.frx":2CDC
      PicturePos      =   1
   End
End
Attribute VB_Name = "frmCP_Labor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmCP_Labor"

Private WithEvents oDriver As clsFormDriver
Attribute oDriver.VB_VarHelpID = -1
Private oSkin As clsFormSkin
Private bLoaded As Boolean

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   
   lsOldProc = "cmdButton_Click"
   ''On Error GoTo errProc
   
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
   ''On Error GoTo errProc

   oApp.MenuName = Me.Tag
   Me.ZOrder 0

   If Not bLoaded Then
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
   ''On Error GoTo errProc

   CenterChildForm mdiMain, Me
   
   bLoaded = False
   
   Set oDriver = New clsFormDriver
   Set oDriver.AppDriver = oApp
   Set oDriver.MainForm = Me
   
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin
   
   oDriver.RecQuery = "SELECT * FROM CP_Labor"
   oDriver.BrowseQuery = "SELECT" _
                           & "  a.sLaborIDx" _
                           & ", a.sLaborNme" _
                           & ", a.sLaborCde" _
                           & ", b.sBrandNme" _
                           & ", a.nPriceLv1" _
                           & ", a.nPriceLv2" _
                           & ", a.nPriceLv3" _
                        & " FROM CP_Labor a" _
                           & ", CP_Brand b" _
                        & " WHERE a.cRecdStat = " & strParm(xeRecStateActive) _
                           & " AND a.sBrandIDx = b.sBrandIDx" _
                        & " ORDER BY a.sLaborIDx"
   
   oDriver.InitRecForm
   
   oDriver.BrowseFTitle(0) = "Labor ID"
   oDriver.BrowseFTitle(1) = "Labor Name"
   oDriver.BrowseFTitle(2) = "Code"
   oDriver.BrowseFTitle(3) = "Brand"
   oDriver.BrowseFTitle(4) = "Level 01"
   oDriver.BrowseFTitle(5) = "Level 02"
   oDriver.BrowseFTitle(6) = "Level 03"
   
   oDriver.BrowseFFormat(4) = "#,##0.00"
   oDriver.BrowseFFormat(5) = "#,##0.00"
   oDriver.BrowseFFormat(6) = "#,##0.00"
   
   oDriver.LookupQuery(3) = "SELECT" _
                              & "  sBrandIDx" _
                              & ", sBrandNme" _
                           & " FROM CP_Brand" _
                           & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
                           & " ORDER BY sBrandNme"
   oDriver.LookupReference(3) = "sBrandIDx»sBrandNme"
   oDriver.LookupColumn(3) = "sBrandNme"
   oDriver.LookupTitle(3) = "Brand Name"
   
   oDriver.FieldStart = 2
   
   oDriver.FieldFormat(4) = "#,##0.00"
   oDriver.FieldFormat(5) = "#,##0.00"
   oDriver.FieldFormat(6) = "#,##0.00"
   oDriver.FieldFormat(7) = "#,##0.00"

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oDriver = Nothing
   Set oSkin = Nothing
End Sub

Private Sub oDriver_DisableOtherControl()
   OptField(0).Enabled = False
   OptField(1).Enabled = False
   OptField(2).Enabled = False
   oDriver.hideButton 3
   chkField.Enabled = False
End Sub

Private Sub oDriver_EnableOtherControl()
   oDriver.DisableTextbox 0
   OptField(0).Enabled = True
   OptField(1).Enabled = True
   OptField(2).Enabled = True
   chkField.Enabled = True
End Sub

Private Sub oDriver_InitValue()
   Dim lsOldProc As String
   
   lsOldProc = "oDriver_InitValue"
   ''On Error GoTo errProc

   If Not oDriver.SetValue(0, GetNextCode("CP_Labor", "sLaborIDx", False, oApp.Connection, True, oApp.BranchCode)) Then Exit Sub
   oDriver.FieldReference(0) = True
   oDriver.FieldValue(9) = xeRecStateActive
   chkField.Value = 0
   OptField(0).Value = True

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )"
End Sub

Private Sub oDriver_LoadOtherData()
   chkField.Value = oDriver.FieldValue(7)
   OptField(oDriver.FieldValue(8)).Value = True
End Sub

Private Sub oDriver_WillSave(Cancel As Boolean)
   Dim lnCtr As Integer
   
   oDriver.FieldValue(7) = chkField.Value
   If OptField(0).Value Then oDriver.FieldValue(8) = "0"
   If OptField(1).Value Then oDriver.FieldValue(8) = "1"
   If OptField(2).Value Then oDriver.FieldValue(8) = "2"
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   oDriver.ColumnIndex = Index
   With txtField(Index)
      .BackColor = oApp.getColor("HT1")
   End With
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   Dim lsOldProc As String
   
   lsOldProc = "txtField_Validate"
   ''On Error GoTo errProc
   
   If Index = 1 Or Index = 2 Then
      txtField(Index).Text = UCase(txtField(Index).Text)
   Else
      txtField(Index).Text = TitleCase(txtField(Index).Text)
   End If
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
