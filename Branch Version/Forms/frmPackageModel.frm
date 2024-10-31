VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmPackageModel 
   BorderStyle     =   0  'None
   Caption         =   "Package Per Model Maintenance"
   ClientHeight    =   3765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4305
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3765
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   2190
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   4110
      _ExtentX        =   7250
      _ExtentY        =   3863
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   1140
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   1545
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
         Left            =   1140
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   315
         Width           =   1425
      End
      Begin VB.TextBox txtOther 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   1140
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   1215
         Width           =   2760
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   1140
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   885
         Width           =   2760
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   1
         Left            =   75
         TabIndex        =   4
         Top             =   1290
         Width           =   885
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   315
         Left            =   1215
         Tag             =   "et0;ht2"
         Top             =   405
         Width           =   1425
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "PackModID"
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
         Left            =   105
         TabIndex        =   0
         Top             =   375
         Width           =   1020
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
         Height          =   195
         Index           =   4
         Left            =   75
         TabIndex        =   6
         Top             =   1620
         Width           =   810
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Barcode"
         Height          =   195
         Index           =   2
         Left            =   75
         TabIndex        =   2
         Top             =   945
         Width           =   885
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   0
      Left            =   3465
      TabIndex        =   14
      Top             =   2970
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
      Picture         =   "frmPackageModel.frx":0000
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   1
      Left            =   2685
      TabIndex        =   13
      Top             =   2970
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
      Picture         =   "frmPackageModel.frx":077A
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   2
      Left            =   1125
      TabIndex        =   11
      Top             =   2970
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
      Picture         =   "frmPackageModel.frx":0EF4
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   4
      Left            =   345
      TabIndex        =   8
      Top             =   2970
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
      Picture         =   "frmPackageModel.frx":166E
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   5
      Left            =   3465
      TabIndex        =   15
      Top             =   2970
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
      Picture         =   "frmPackageModel.frx":1DE8
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   7
      Left            =   1905
      TabIndex        =   12
      Top             =   2970
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
      Picture         =   "frmPackageModel.frx":2562
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   3
      Left            =   1125
      TabIndex        =   9
      Top             =   2970
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
      Picture         =   "frmPackageModel.frx":2CDC
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   6
      Left            =   1905
      TabIndex        =   10
      Top             =   2970
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
      Picture         =   "frmPackageModel.frx":3456
      PicturePos      =   1
   End
End
Attribute VB_Name = "frmPackageModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmPackageModel"

Private WithEvents oDriver As clsFormDriver
Attribute oDriver.VB_VarHelpID = -1
Private oSkin As clsFormSkin
Private bLoaded As Boolean

Dim pnIndex As Integer

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   
   lsOldProc = "cmdButton_Click"
   'On Error GoTo errProc
   
   txtField_LostFocus pnIndex
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
   Dim lsConcatDesc As String
   
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
                        & "  sPckModID" _
                        & ", sStockIDx" _
                        & ", sModelIDx" _
                        & ", nQuantity" _
                        & ", cRecdStat" _
                        & ", sModified" _
                        & ", dModified" _
                     & " FROM Package_Model"
                     
   oDriver.BrowseQuery = "SELECT" _
                           & "  a.sPckModID" _
                           & ", b.sBarrCode" _
                           & ", b.sDescript" _
                           & ", c.sModelNme" _
                        & " FROM Package_Model a" _
                           & ", CP_Inventory b" _
                           & ", CP_Model c" _
                        & " WHERE a.sStockIDx = b.sStockIDx" _
                           & " AND a.sModelIDx = c.sModelIDx" _
                           & " AND a.cRecdStat = " & strParm(xeRecStateActive) _
                        & " ORDER BY b.sDescript" _
                           & ", c.sModelNme"
   
   oDriver.InitRecForm
   
   oDriver.BrowseFReference(0) = True
   
   oDriver.BrowseFTitle(0) = "Code"
   oDriver.BrowseFTitle(1) = "Barcode"
   oDriver.BrowseFTitle(2) = "Description"
   oDriver.BrowseFTitle(3) = "Model"

   lsConcatDesc = "CONCAT(a.sDescript, ' '" _
                        & ", IF(b.sBrandNme IS NULL, '', b.sBrandNme), ' '" _
                        & ", IF(c.sModelNme IS NULL, '', c.sModelNme), ' '" _
                        & ", IF(d.sColorNme IS NULL, '', d.sColorNme), ' '" _
                        & ", IF(f.sSizeName IS NULL, '', f.sSizeName))"
   
   oDriver.LookupQuery(1) = "SELECT" _
                              & "  a.sStockIDx" _
                              & ", a.sBarrcode" _
                              & ", " & lsConcatDesc & " xDescript" _
                              & ", b.sBrandNme" _
                              & ", c.sModelNme" _
                              & ", f.sSizeName" _
                              & ", d.sColorNme" _
                              & ", e.sCategrNm" _
                           & " FROM CP_Inventory a" _
                              & " LEFT JOIN CP_Brand b" _
                                 & " ON a.sBrandIDx = b.sBrandIDx" _
                              & " LEFT JOIN CP_Model c" _
                                 & " ON a.sModelIDx = c.sModelIDx" _
                              & " LEFT JOIN Color d" _
                                 & " ON a.sColorIDx = d.sColorIDx" _
                              & " LEFT JOIN Size f" _
                                 & " ON a.sSizeIDxx = f.sSizeIDxx" _
                              & " LEFT JOIN Category e" _
                                 & " ON a.sCategID1 = e.sCategrID" _
                           & " WHERE a.cRecdStat = " & strParm(xeRecStateActive)

   oDriver.LookupReference(1) = "a.sStockIDx»a.sBarrCode»" & lsConcatDesc & "»b.sBrandNme»" _
                                 & "c.sModelNme»f.sSizeName»d.sColorNme"
   oDriver.LookupColumn(1) = "sBarrCode»xDescript»sBrandNme»sModelNme»sColorNme»sSizeName"
   oDriver.LookupTitle(1) = "BarCode»Descript»Brand»Model»Color»Size"

   oDriver.LookupQuery(2) = "SELECT" _
                              & "  a.sModelIDx" _
                              & ", a.sModelNme" _
                              & ", b.sBrandNme" _
                           & " FROM CP_Model a" _
                              & " LEFT JOIN CP_Brand b" _
                                 & " ON a.sBrandIDx = b.sBrandIDx" _
                           & " WHERE a.cRecdStat = " & strParm(xeRecStateActive) _
                           & " ORDER BY a.sModelNme"

   oDriver.LookupReference(2) = "a.sModelIDx»a.sModelNme"
   oDriver.LookupColumn(2) = "sModelNme»sBrandNme"
   oDriver.LookupTitle(2) = "Model»Brand"

   oDriver.FieldFormat(0) = "@@@@-@@@"
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

Private Sub oDriver_InitValue()
   Dim lsOldProc As String
   
   lsOldProc = "oDriver_InitValue"
   'On Error GoTo errProc

   If oDriver.SetValue(0, GetNextCode("Package_Model", "sPckModID", True, oApp.Connection, True, oApp.BranchCode)) = False Then Exit Sub
   oDriver.FieldReference(0) = True
   oDriver.FieldValue(3) = 0
   oDriver.FieldValue(4) = xeRecStateActive
   
   txtOther(0).Text = ""
   
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )"
End Sub

Private Sub oDriver_EnableOtherControl()
   oDriver.DisableTextbox 0
End Sub

Private Sub oDriver_LoadOtherData()
   txtOther(0).Text = getDescript(oDriver.FieldValue(1))
End Sub

Private Sub oDriver_WillSave(Cancel As Boolean)
   If oDriver.FieldValue(1) = "" Then
      MsgBox "Invalid BarrCode detected!!!", vbCritical, "Warning"
      txtField(1).SetFocus
      Cancel = True
   ElseIf oDriver.FieldValue(2) = "" Then
      MsgBox "Invalid CP Model detected!!!", vbCritical, "Warning"
      txtField(2).SetFocus
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
   
   If Index = 1 Then txtOther(0).Text = getDescript(oDriver.FieldValue(1))
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

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   Dim lsOldProc As String
   
   lsOldProc = "txtField_Validate"
   'On Error GoTo errProc
   
   Cancel = Not oDriver.ValidateField(Index)
   If Index = 1 Then txtOther(0).Text = getDescript(oDriver.FieldValue(1))
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " _
                       & "  " & Index _
                       & ", " & Cancel _
                       & " )", True
End Sub

Private Function getDescript(lsStockIDx As String) As String
   Dim lrs As Recordset
   
   Set lrs = New Recordset
   lrs.Open "SELECT * FROM CP_Inventory" & _
               " WHERE sStockIDx = " & strParm(lsStockIDx) _
   , oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   
   If lrs.EOF Then
      getDescript = ""
      Exit Function
   End If
   
   getDescript = lrs("sDescript")
End Function

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
