VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmAssets 
   BorderStyle     =   0  'None
   Caption         =   "Assets"
   ClientHeight    =   5250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4575
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5250
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   3705
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   6535
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtOther 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1035
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   1290
         Width           =   3120
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   1035
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   630
         Width           =   3120
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
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   90
         Width           =   1935
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   1035
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   960
         Width           =   3120
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Classify"
         Height          =   1725
         Left            =   75
         TabIndex        =   8
         Tag             =   "wt0;fb0"
         Top             =   1755
         Width           =   1770
         Begin VB.OptionButton optClassify 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Land"
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   24
            Tag             =   "wt0;fb0"
            Top             =   1170
            Width           =   1170
         End
         Begin VB.OptionButton optClassify 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Building"
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   23
            Tag             =   "wt0;fb0"
            Top             =   1395
            Width           =   1170
         End
         Begin VB.OptionButton optClassify 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Vehicle"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   22
            Tag             =   "wt0;fb0"
            Top             =   930
            Width           =   1170
         End
         Begin VB.OptionButton optClassify 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Furnitures"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   9
            Tag             =   "wt0;fb0"
            Top             =   255
            Width           =   1170
         End
         Begin VB.OptionButton optClassify 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Fixtures"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   10
            Tag             =   "wt0;fb0"
            Top             =   480
            Width           =   1170
         End
         Begin VB.OptionButton optClassify 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Equipment"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   11
            Tag             =   "wt0;fb0"
            Top             =   705
            Width           =   1170
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Depreciation Method"
         Height          =   1035
         Left            =   1890
         TabIndex        =   12
         Tag             =   "wt0;fb0"
         Top             =   1740
         Width           =   2385
         Begin VB.OptionButton optDM 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Straight - Line"
            Height          =   255
            Index           =   0
            Left            =   270
            TabIndex        =   13
            Tag             =   "wt0;fb0"
            Top             =   285
            Value           =   -1  'True
            Width           =   1830
         End
         Begin VB.OptionButton optDM 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Reducing Balance"
            Height          =   255
            Index           =   1
            Left            =   270
            TabIndex        =   14
            Tag             =   "wt0;fb0"
            Top             =   495
            Width           =   1830
         End
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Acct Code"
         Height          =   285
         Index           =   3
         Left            =   120
         TabIndex        =   6
         Top             =   1350
         Width           =   885
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   285
         Index           =   2
         Left            =   105
         TabIndex        =   2
         Top             =   660
         Width           =   885
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   1
         Left            =   1080
         Tag             =   "et0;ht2"
         Top             =   180
         Width           =   1950
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Asset ID"
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
         Left            =   90
         TabIndex        =   0
         Top             =   150
         Width           =   885
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Acct Desc"
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   1020
         Width           =   885
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   0
      Left            =   3765
      TabIndex        =   15
      Top             =   4455
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
      Picture         =   "frmAssets.frx":0000
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   1
      Left            =   3000
      TabIndex        =   16
      Top             =   4455
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
      Picture         =   "frmAssets.frx":077A
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   2
      Left            =   2235
      TabIndex        =   17
      Top             =   4455
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
      Picture         =   "frmAssets.frx":0EF4
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   3
      Left            =   1470
      TabIndex        =   18
      Top             =   4455
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
      Picture         =   "frmAssets.frx":166E
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   4
      Left            =   705
      TabIndex        =   19
      Top             =   4455
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
      Picture         =   "frmAssets.frx":1DE8
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   5
      Left            =   3765
      TabIndex        =   20
      Top             =   4455
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
      Picture         =   "frmAssets.frx":2562
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   6
      Left            =   2235
      TabIndex        =   21
      Top             =   4455
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
      Picture         =   "frmAssets.frx":2CDC
      PicturePos      =   1
   End
End
Attribute VB_Name = "frmAssets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmAssets"

Private WithEvents oDriver As clsFormDriver
Attribute oDriver.VB_VarHelpID = -1
Private oSkin As clsFormSkin
Private bLoaded As Boolean

Dim pnIndex As Integer

Private Sub Form_Load()
   Dim lsOldProc As String
      
   lsOldProc = "Form_Load"
   ''On Error GoTo errProc
   
   CenterChildForm mdiMain, Me
   
   bLoaded = False
   
   Set oDriver = New clsFormDriver
   Set oDriver.AppDriver = oapp
   Set oDriver.MainForm = Me
   
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oapp
   Set oSkin.Form = Me
   oSkin.ApplySkin
   
   oDriver.RecQuery = "SELECT * FROM Assets"
   oDriver.BrowseQuery = "SELECT" _
                              & "  sAssetIDx" _
                              & ", sDescript" _
                              & ", sAcctCode" _
                              & ", cClassify" _
                              & ", cDepMethd" _
                              & ", cRecdStat" _
                              & ", sModified" _
                              & ", dModified" _
                           & " FROM Assets" _
                           & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
                           & " ORDER BY sAssetIDx"
                           
   oDriver.InitRecForm
      
   oDriver.BrowseFTitle(0) = "Code"
   oDriver.BrowseFTitle(1) = "Description"
   oDriver.BrowseFFormat(0) = "@@@@"
   
   oDriver.LookupQuery(2) = "SELECT" _
                           & "  sAcctCode" _
                           & ", sDescript" _
                        & " FROM Account_Chart" _
                        & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
                        & " ORDER BY sAcctCode"
   
   oDriver.LookupReference(2) = "sAcctCode�sDescript"
   oDriver.LookupColumn(2) = "sAcctCode�sDescript"
   oDriver.LookupTitle(2) = "Code�Description"
   
   oDriver.FieldFormat(0) = "@@@@@@"
   oDriver.FieldSize(0) = Len(oDriver.FieldFormat(0))
   oDriver.FieldStart = 1

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   
   lsOldProc = "cmdButton_Click"
   ''On Error GoTo errProc
   
   txtField_LostFocus pnIndex
   Select Case Index
   Case 0
      oDriver.RecordCancelUpdate
   Case 1
      oDriver.BrowseRecord
   Case 2
      If txtField(1).Text <> "" Or txtField(2).Text <> "" Then
         oDriver.RecordSave
         MsgBox "Record save successfully!!!", vbInformation
      End If
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
      oDriver_LoadOtherData
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
   
   oapp.MenuName = Me.Tag
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

Private Sub Form_Unload(Cancel As Integer)
   Set oDriver = Nothing
   Set oSkin = Nothing
End Sub

Private Sub oDriver_DisableOtherControl()
   optClassify(0).Enabled = False
   optClassify(1).Enabled = False
   optClassify(2).Enabled = False
   optClassify(3).Enabled = False
   optClassify(4).Enabled = False
   optClassify(5).Enabled = False
   
   
   optDM(0).Enabled = False
   optDM(1).Enabled = False
End Sub

Private Sub oDriver_EnableOtherControl()
   oDriver.DisableTextbox 0
   
   optClassify(0).Enabled = True
   optClassify(1).Enabled = True
   optClassify(2).Enabled = True
   optClassify(3).Enabled = True
   optClassify(4).Enabled = True
   optClassify(5).Enabled = True
   
   optDM(0).Enabled = True
   optDM(1).Enabled = True
   
End Sub

Private Sub oDriver_InitValue()
   Dim lsOldProc As String
   
   lsOldProc = "oDriver_InitValue"
   ''On Error GoTo errProc

   If oDriver.SetValue(0, GetNextCode("Assets", "sAssetIDx", False, oapp.Connection, False, "")) = False Then Exit Sub
   oDriver.FieldReference(1) = True
   oDriver.FieldValue(1) = ""
   oDriver.FieldValue(2) = ""
   oDriver.FieldValue(3) = 0
   oDriver.FieldValue(4) = 0
   oDriver.FieldValue(5) = xeRecStateActive
      
   optClassify(oDriver.FieldValue(3)).Value = True
   optDM(oDriver.FieldValue(4)).Value = True
   txtOther = ""

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )"
End Sub

Private Sub oDriver_LoadOtherData()
   Dim lors As Recordset

   Set lors = New Recordset
   lors.Open "SELECT" & _
                  "  sAcctCode" & _
                  ", sDescript" & _
               " FROM Account_Chart" & _
               " WHERE sAcctCode = " & strParm(oDriver.FieldValue(2)) _
   , oapp.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText

   txtOther = ""
   If Not lors.EOF Then txtOther = lors("sAcctCode")
   Set lors = Nothing
   
   optClassify(oDriver.FieldValue(3)).Value = True
   optDM(oDriver.FieldValue(4)).Value = True
End Sub

Private Sub optClassify_Click(Index As Integer)
   oDriver.FieldValue(3) = Index
End Sub

Private Sub optDM_Click(Index As Integer)
   oDriver.FieldValue(4) = Index
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   oDriver.ColumnIndex = Index
   With txtField(Index)
      .BackColor = oapp.getColor("HT1")
   End With
   pnIndex = Index
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String
   
   lsOldProc = "txtField_KeyDown"
   ''On Error GoTo errProc
   
   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
      With txtField(pnIndex)
         Select Case Index
         Case 2
            If KeyCode = vbKeyF3 Then
               oDriver.RecordSearch .Text
               If .Text <> "" Then SetNextFocus
            Else
               If .Text <> "" Then oDriver.RecordSearch .Text
            End If
            Call oDriver_LoadOtherData
         End Select
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
   With txtField(Index)
      .BackColor = oapp.getColor("EB")
   End With
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   Dim lsOldProc As String
   
   lsOldProc = "txtField_Validate"
   ''On Error GoTo errProc
   
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
   With oapp
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