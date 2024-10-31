VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmCP_Model 
   BorderStyle     =   0  'None
   Caption         =   "CP Model"
   ClientHeight    =   3435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5535
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3435
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1860
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   5355
      _ExtentX        =   9446
      _ExtentY        =   3281
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   3
         Left            =   1185
         TabIndex        =   7
         Top             =   1395
         Width           =   2055
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   1185
         TabIndex        =   5
         Top             =   1065
         Width           =   3945
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   1185
         TabIndex        =   3
         Top             =   735
         Width           =   3945
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
         Left            =   1185
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model Code"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   6
         Top             =   1455
         Width           =   855
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Brand Name"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   4
         Top             =   1125
         Width           =   885
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   315
         Left            =   1275
         Tag             =   "et0;ht2"
         Top             =   330
         Width           =   2055
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   2
         Top             =   795
         Width           =   795
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
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
         Index           =   0
         Left            =   135
         TabIndex        =   0
         Top             =   300
         Width           =   780
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   0
      Left            =   4695
      TabIndex        =   14
      Top             =   2655
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
      Picture         =   "frmCP_Model.frx":0000
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   1
      Left            =   3915
      TabIndex        =   13
      Top             =   2655
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
      Picture         =   "frmCP_Model.frx":077A
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   2
      Left            =   2355
      TabIndex        =   8
      Top             =   2655
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
      Picture         =   "frmCP_Model.frx":0EF4
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   4
      Left            =   1575
      TabIndex        =   10
      Top             =   2655
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
      Picture         =   "frmCP_Model.frx":166E
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   5
      Left            =   4695
      TabIndex        =   15
      Top             =   2655
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
      Picture         =   "frmCP_Model.frx":1DE8
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   705
      Index           =   7
      Left            =   3135
      TabIndex        =   9
      Top             =   2655
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
      Picture         =   "frmCP_Model.frx":2562
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   3
      Left            =   2355
      TabIndex        =   11
      Top             =   2655
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
      Picture         =   "frmCP_Model.frx":2CDC
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   6
      Left            =   3135
      TabIndex        =   12
      Top             =   2655
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
      Picture         =   "frmCP_Model.frx":3456
      PicturePos      =   1
   End
End
Attribute VB_Name = "frmCP_Model"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmCPModel"

Private WithEvents oDriver As clsFormDriver
Attribute oDriver.VB_VarHelpID = -1
Private oSkin As clsFormSkin
Private bLoaded As Boolean

Dim pnCtr As Integer
Dim pnIndex As Integer

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
   
   With oDriver
      .RecQuery = "SELECT" _
                     & "  sModelIDx" _
                     & ", sModelNme" _
                     & ", sBrandIDx" _
                     & ", sModelCde" _
                     & ", cRecdStat" _
                     & ", sModified" _
                     & ", dModified" _
                  & " FROM CP_Model"
           
      .BrowseQuery = "SELECT" _
                        & "  a.sModelIDx" _
                        & ", a.sModelNme" _
                        & ", a.sModelCde" _
                        & ", b.sBrandNme" _
                     & " FROM CP_Model a" _
                        & " Left Join CP_Brand b" _
                           & " On a.sBrandIDx = b.sBrandIDx" _
                     & " WHERE a.cRecdStat = " & strParm(xeRecStateActive) _
                     & " ORDER BY a.sModelIDx"
      .InitRecForm
   
      .BrowseFTitle(0) = "ID"
      .BrowseFTitle(1) = "Model Desc"
      .BrowseFTitle(2) = "Model Code"
      .BrowseFTitle(3) = "Brand"
      .BrowseFFormat(0) = "@@-@@@"

      .LookupQuery(2) = "SELECT" _
                           & "  sBrandIDx" _
                           & ", sBrandNme" _
                        & " FROM CP_Brand " _
                        & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
                        & " ORDER BY sBrandNme"
   
      .LookupReference(2) = "sBrandIDx»sBrandNme"
      .LookupColumn(2) = "sBrandNme"
      .LookupTitle(2) = "Brand"
   
      .FieldFormat(0) = "@@-@@@"
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

Private Sub oDriver_EnableOtherControl()
   oDriver.DisableTextbox 0
End Sub

Private Sub oDriver_InitValue()
   Dim lsOldProc As String
   
   lsOldProc = "oDriver_InitValue"
   ''On Error GoTo errProc
   
   If oDriver.SetValue(0, GetNextCode("CP_Model", "sModelIDx", True, oApp.Connection, True, oApp.BranchCode)) = False Then Exit Sub
   oDriver.FieldReference(0) = True
   oDriver.FieldValue(1) = ""
   oDriver.FieldValue(2) = ""
   oDriver.FieldValue(3) = ""
   oDriver.FieldValue(4) = xeRecStateActive
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )"
End Sub

Private Sub oDriver_WillSave(Cancel As Boolean)
'   If oDriver.FieldValue(1) = "" Then
'      MsgBox "Invalid Model Description detected!!!", vbCritical, "Warning"
'      txtField(1).SetFocus
'      Cancel = True
'      Exit Sub
'   End If
   
   If oDriver.FieldValue(3) = "" Then
      MsgBox "Invalid Model Code detected!!!", vbCritical, "Warning"
      txtField(3).SetFocus
      Cancel = True
   End If
   
   If oDriver.EditMode = xeModeAddNew Then Cancel = Not insertModelPrice
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
   ''On Error GoTo errProc
   
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
   ''On Error GoTo errProc
      
   txtField(Index).Text = TitleCase(txtField(Index).Text)
   Cancel = Not oDriver.ValidateField(Index)
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Function insertModelPrice() As Boolean
   Dim lsSQL As String
   
   lsSQL = "INSERT INTO CP_Model_Price" & _
            " SET sModelIDx = " & strParm(oDriver.FieldValue(0)) & _
               ", nSelPrice = 0.00" & _
               ", nSelPrce2 = 0.00" & _
               ", nSelPrce3 = 0.00" & _
               ", nSelPrce4 = 0.00" & _
               ", nLastPrce = 0.00" & _
               ", sApproved = " & strParm(Encrypt(oApp.UserID)) & _
               ", dPricexxx = " & dateParm(oApp.ServerDate) & _
               ", cRecdStat = " & strParm(xeYes) & _
               ", sPrepared = " & strParm(Encrypt(oApp.UserID)) & _
               ", dPrepared = " & dateParm(oApp.ServerDate) & _
               ", sModified = " & strParm(Encrypt(oApp.UserID)) & _
               ", dModified = " & dateParm(oApp.ServerDate)
   
   If oApp.Execute(lsSQL, "CP_Model_Price") <= 0 Then
      MsgBox "Unable to Insert Model Price. Please Inform MIS Dept.", vbCritical, "Warning"
      Exit Function
   End If

   insertModelPrice = True
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
