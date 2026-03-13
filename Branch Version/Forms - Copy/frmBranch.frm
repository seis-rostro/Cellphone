VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmBranch 
   BorderStyle     =   0  'None
   Caption         =   "Branch Maintenance"
   ClientHeight    =   4455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6510
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4455
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   2835
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   5001
      BorderStyle     =   1
      Begin VB.CheckBox chkWarehouse 
         Caption         =   "Warehousse"
         Height          =   195
         Left            =   2640
         TabIndex        =   13
         Tag             =   "et0;fb0"
         Top             =   2445
         Width           =   1470
      End
      Begin VB.CheckBox chkMainOffice 
         Caption         =   "Main Office"
         Height          =   195
         Left            =   1380
         TabIndex        =   12
         Tag             =   "et0;fb0"
         Top             =   2445
         Width           =   1185
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   5
         Left            =   1380
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   2055
         Width           =   2730
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   1380
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   1065
         Width           =   4680
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   3
         Left            =   1380
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   1395
         Width           =   4680
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   4
         Left            =   1380
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   1725
         Width           =   4680
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   1380
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   735
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
         Height          =   315
         Index           =   0
         Left            =   1380
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   300
         Width           =   630
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Company"
         Height          =   195
         Index           =   5
         Left            =   195
         TabIndex        =   4
         Top             =   1125
         Width           =   1125
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   315
         Left            =   1500
         Tag             =   "et0;ht2"
         Top             =   390
         Width           =   630
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Branch ID"
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
         Top             =   360
         Width           =   810
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Town"
         Height          =   195
         Index           =   4
         Left            =   195
         TabIndex        =   8
         Top             =   1800
         Width           =   810
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Manager"
         Height          =   195
         Index           =   3
         Left            =   195
         TabIndex        =   10
         Top             =   2115
         Width           =   810
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Branch Name"
         Height          =   195
         Index           =   1
         Left            =   195
         TabIndex        =   2
         Top             =   795
         Width           =   1125
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   195
         Index           =   2
         Left            =   195
         TabIndex        =   6
         Top             =   1455
         Width           =   810
      End
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   705
      Index           =   0
      Left            =   5655
      TabIndex        =   20
      Top             =   3660
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
      Picture         =   "frmBranch.frx":0000
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   705
      Index           =   1
      Left            =   4875
      TabIndex        =   19
      Top             =   3660
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
      Picture         =   "frmBranch.frx":077A
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   705
      Index           =   2
      Left            =   3315
      TabIndex        =   17
      Top             =   3660
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
      Picture         =   "frmBranch.frx":0EF4
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   705
      Index           =   4
      Left            =   2535
      TabIndex        =   14
      Top             =   3660
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
      Picture         =   "frmBranch.frx":166E
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   705
      Index           =   5
      Left            =   5655
      TabIndex        =   21
      Top             =   3660
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
      Picture         =   "frmBranch.frx":1DE8
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   705
      Index           =   7
      Left            =   4095
      TabIndex        =   18
      Top             =   3660
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
      Picture         =   "frmBranch.frx":2562
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   705
      Index           =   3
      Left            =   3315
      TabIndex        =   15
      Top             =   3660
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
      Picture         =   "frmBranch.frx":2CDC
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   705
      Index           =   6
      Left            =   4095
      TabIndex        =   16
      Top             =   3660
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
      Picture         =   "frmBranch.frx":3456
      PicturePos      =   1
   End
End
Attribute VB_Name = "frmBranch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmBranch"

Private WithEvents oDriver As clsFormDriver
Attribute oDriver.VB_VarHelpID = -1
Private oSkin As clsFormSkin
Private bLoaded As Boolean

Dim pnCtr As Integer
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
   Dim lsSQL As String
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
   
   With oDriver
      .RecQuery = "SELECT" _
                     & "  sBranchCd" _
                     & ", sBranchNm" _
                     & ", sCompnyID" _
                     & ", sAddressx" _
                     & ", sTownIDxx" _
                     & ", sManagerx" _
                     & ", cMainOffc" _
                     & ", cWareHous" _
                     & ", cRecdStat" _
                     & ", sModified" _
                     & ", dModified" _
                  & " FROM Branch"
           
      .BrowseQuery = "SELECT" _
                        & "  a.sBranchCd" _
                        & ", a.sBranchNm" _
                        & ", b.sCompnyNm" _
                     & " FROM Branch a" _
                        & " Left Join Company b" _
                           & " On a.sCompnyID = b.sCompnyID" _
                     & " WHERE a.cRecdStat = " & strParm(xeRecStateActive) _
                     & " ORDER BY a.sBranchCd, a.sBranchNm"
      .InitRecForm
   
      .BrowseFTitle(0) = "Code"
      .BrowseFTitle(1) = "Branch"
      .BrowseFTitle(2) = "Company"
      .BrowseFFormat(0) = "@"

      .LookupQuery(2) = "SELECT" _
                           & "  sCompnyID" _
                           & ", sCompnyNm" _
                        & " FROM Company" _
                        & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
                        & " ORDER BY sCompnyNm"
   
      .LookupReference(2) = "sCompnyID»sCompnyNm"
      .LookupColumn(2) = "sCompnyNm"
      .LookupTitle(2) = "Company"
      
      .LookupQuery(4) = "SELECT" _
                           & "  sTownIDxx" _
                           & ", sTownName" _
                        & " FROM TownCity" _
                        & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
                        & " ORDER BY sTownName"
   
      .LookupReference(4) = "sTownIDxx»sTownName"
      .LookupColumn(4) = "sTownName"
      .LookupTitle(4) = "TownCity"
   
      .FieldFormat(0) = "@@"
      .FieldSize(0) = Len(oDriver.FieldFormat(0))
      .FieldStart = 1
   End With
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
   Set oDriver = Nothing
End Sub

Private Sub oDriver_DisableOtherControl()
   chkMainOffice.Enabled = False
End Sub

Private Sub oDriver_EnableOtherControl()
   oDriver.DisableTextbox 0
   
   chkMainOffice.Enabled = True
End Sub

Private Sub oDriver_InitValue()
   Dim lsOldProc As String
   
   lsOldProc = "oDriver_InitValue"
   'On Error GoTo errProc
   
   If oDriver.SetValue(0, getNextBranch) = False Then Exit Sub
   oDriver.FieldReference(0) = True
   oDriver.FieldValue(1) = ""
   oDriver.FieldValue(2) = ""
   oDriver.FieldValue(3) = ""
   oDriver.FieldValue(4) = ""
   oDriver.FieldValue(5) = ""
   oDriver.FieldValue(6) = 0
   oDriver.FieldValue(7) = 0
   oDriver.FieldValue(8) = xeRecStateActive
   
   chkMainOffice.Value = oDriver.FieldValue(6)
   chkWarehouse.Value = oDriver.FieldValue(7)
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )"
End Sub

Private Sub oDriver_LoadOtherData()
   chkMainOffice.Value = oDriver.FieldValue(6)
   chkWarehouse.Value = oDriver.FieldValue(7)
End Sub

Private Sub oDriver_WillSave(Cancel As Boolean)
   If oDriver.FieldValue(1) = "" Then
      MsgBox "Invalid Description detected!!!", vbCritical, "Warning"
      txtField(1).SetFocus
      Cancel = True
   End If
   
   oDriver.FieldValue(6) = chkMainOffice.Value
   oDriver.FieldValue(7) = chkWarehouse.Value
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("HT1")
   End With
   
   oDriver.ColumnIndex = Index
   pnIndex = Index
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String
   
   lsOldProc = "txtField_KeyDown"
   'On Error GoTo errProc
   
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

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   Dim lsOldProc As String
   
   lsOldProc = "txtField_LostFocus"
   'On Error GoTo errProc
      
   txtField(Index).Text = TitleCase(txtField(Index).Text)
   Cancel = Not oDriver.ValidateField(Index)
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
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

Private Function getNextBranch() As String
   Dim loRS As Recordset
   Dim lsOldProc As String
   Dim lsSQL As String
   
   lsOldProc = "getNextBranch"
   lsSQL = "SELECT sBranchCd" & _
            " FROM Branch" & _
            " WHERE sBranchCd REGEXP " & strParm("^[0-9]") & _
            " ORDER BY sBranchCd DESC" & _
            " LIMIT 1"
   
   Set loRS = New Recordset
   loRS.Open lsSQL, oApp.Connection, , , adCmdText
   If loRS.EOF Then
      lsSQL = "01"
   Else
      lsSQL = Format(loRS("sBranchCd") + 1, "00")
   End If
   
   getNextBranch = lsSQL
   
endProc:
   Set loRS = Nothing
   Exit Function
errProc:
   ShowError lsOldProc & " ( ) "
   GoTo endProc
End Function
