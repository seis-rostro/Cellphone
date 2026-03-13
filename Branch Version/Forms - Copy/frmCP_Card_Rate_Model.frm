VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmCP_Card_Rate_Modelx 
   BorderStyle     =   0  'None
   Caption         =   "CP Model"
   ClientHeight    =   4740
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5535
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4740
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   3150
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   5355
      _ExtentX        =   9446
      _ExtentY        =   5556
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   9
         Left            =   1185
         TabIndex        =   15
         Top             =   2490
         Width           =   2055
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   7
         Left            =   1185
         TabIndex        =   13
         Top             =   2160
         Width           =   2055
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
         Index           =   1
         Left            =   1185
         TabIndex        =   1
         Top             =   180
         Width           =   3930
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   6
         Left            =   1185
         TabIndex        =   11
         Top             =   1830
         Width           =   2055
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   5
         Left            =   1185
         TabIndex        =   9
         Top             =   1500
         Width           =   2055
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   4
         Left            =   1185
         TabIndex        =   7
         Top             =   1170
         Width           =   2055
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   3
         Left            =   1185
         TabIndex        =   5
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   1185
         TabIndex        =   3
         Top             =   510
         Width           =   3930
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
         Left            =   3840
         TabIndex        =   17
         Top             =   1845
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Priced Date"
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
         Index           =   8
         Left            =   105
         TabIndex        =   14
         Top             =   2550
         Width           =   1020
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Min. 6 Mos."
         Height          =   195
         Index           =   7
         Left            =   300
         TabIndex        =   4
         Top             =   900
         Width           =   825
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
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
         Left            =   600
         TabIndex        =   2
         Top             =   570
         Width           =   525
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Name"
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
         Index           =   1
         Left            =   135
         TabIndex        =   0
         Top             =   240
         Width           =   990
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "24 Mo. Term"
         Height          =   195
         Index           =   6
         Left            =   225
         TabIndex        =   12
         Top             =   2220
         Width           =   900
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "12 Mo. Term"
         Height          =   195
         Index           =   5
         Left            =   225
         TabIndex        =   10
         Top             =   1890
         Width           =   900
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "6 Mo. Term"
         Height          =   195
         Index           =   4
         Left            =   315
         TabIndex        =   8
         Top             =   1560
         Width           =   810
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3 Mo. Term"
         Height          =   195
         Index           =   3
         Left            =   315
         TabIndex        =   6
         Top             =   1230
         Width           =   810
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   315
         Left            =   3930
         Tag             =   "et0;ht2"
         Top             =   1935
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "cRecdStat"
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
         Left            =   4290
         TabIndex        =   16
         Top             =   1575
         Visible         =   0   'False
         Width           =   915
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   0
      Left            =   4695
      TabIndex        =   24
      Top             =   3945
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
      Picture         =   "frmCP_Card_Rate_Model.frx":0000
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   1
      Left            =   3915
      TabIndex        =   23
      Top             =   3945
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
      Picture         =   "frmCP_Card_Rate_Model.frx":077A
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   2
      Left            =   2355
      TabIndex        =   18
      Top             =   3945
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
      Picture         =   "frmCP_Card_Rate_Model.frx":0EF4
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   4
      Left            =   1575
      TabIndex        =   20
      Top             =   3945
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
      Picture         =   "frmCP_Card_Rate_Model.frx":166E
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   5
      Left            =   4695
      TabIndex        =   25
      Top             =   3945
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
      Picture         =   "frmCP_Card_Rate_Model.frx":1DE8
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   705
      Index           =   7
      Left            =   3135
      TabIndex        =   19
      Top             =   3945
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
      Picture         =   "frmCP_Card_Rate_Model.frx":2562
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   3
      Left            =   2355
      TabIndex        =   21
      Top             =   3945
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
      Picture         =   "frmCP_Card_Rate_Model.frx":2CDC
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   6
      Left            =   3135
      TabIndex        =   22
      Top             =   3945
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
      Picture         =   "frmCP_Card_Rate_Model.frx":3456
      PicturePos      =   1
   End
End
Attribute VB_Name = "frmCP_Card_Rate_Modelx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmCP_Card_Rate_Model"

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
      .RecQuery = "SELECT" & _
                     "  cRecdStat" & _
                     ", sBankIDxx" & _
                     ", sModelIDx" & _
                     ", nMin6Monx" & _
                     ", n03MoTerm" & _
                     ", n06MoTerm" & _
                     ", n12MoTerm" & _
                     ", n24MoTerm" & _
                     ", sApproved" & _
                     ", dPricexxx" & _
                     ", sModified" & _
                     ", dModified" & _
                  " FROM CP_Card_Rate_Model"
           
      .BrowseQuery = "SELECT" & _
                        "  a.sBankIDxx" & _
                        ", a.sModelIDx" & _
                        ", c.sModelCde" & _
                        ", b.sBankName" & _
                        ", c.sModelNme" & _
                     " FROM CP_Card_Rate_Model a" & _
                        " LEFT JOIN Banks b" & _
                           " ON a.sBankIDxx = b.sBankIDxx" & _
                        " LEFT JOIN CP_Model c" & _
                           " ON a.sModelIDx = c.sModelIDx" & _
                     " WHERE a.cRecdStat = " & strParm(xeRecStateActive) & _
                     " ORDER BY b.sBankName"
                     
      .InitRecForm
   
      .BrowseFTitle(0) = "Bank ID"
      .BrowseFTitle(1) = "Model ID"
      .BrowseFTitle(2) = "Model Code"
      .BrowseFTitle(3) = "Bank Name"
      .BrowseFTitle(4) = "Model Name"
      
      .BrowseFReference(0) = True
      .BrowseFReference(1) = True

      .LookupQuery(1) = "SELECT" & _
                           "  sBankIDxx" & _
                           ", sBankName" & _
                        " FROM Banks" & _
                        " WHERE cRecdStat = " & strParm(xeRecStateActive) & _
                        " ORDER BY sBankName"

      .LookupReference(1) = "sBankIDxx製BankName"
      .LookupColumn(1) = "sBankIDxx製BankName"
      .LookupTitle(1) = "Bank ID翡ank Name"
      
      .LookupQuery(2) = "SELECT" & _
                           "  a.sModelIDx" & _
                           ", a.sModelNme" & _
                           ", a.sModelCde" & _
                           ", b.sBrandNme" & _
                        " FROM CP_Model a" & _
                           " LEFT JOIN CP_Brand b" & _
                              " ON a.sBrandIDx = b.sBrandIDx" & _
                        " WHERE a.cRecdStat = " & strParm(xeRecStateActive) & _
                        " ORDER BY a.sModelNme, b.sBrandNme"

      .LookupReference(2) = "sModelIDx製ModelNme製ModelCde製BrandNme"
      .LookupColumn(2) = "sModelIDx製ModelNme製ModelCde製BrandNme"
      .LookupTitle(2) = "Model ID膂odel Name膂odel Code翡rand"
   
      .FieldFormat(0) = "@"
      .FieldFormat(1) = "@"
      .FieldFormat(2) = "@"
      .FieldFormat(3) = "#,##0.00"
      .FieldFormat(4) = "#,##0.00"
      .FieldFormat(5) = "#,##0.00"
      .FieldFormat(6) = "#,##0.00"
      .FieldFormat(7) = "#,##0.00"
      .FieldFormat(8) = "@"
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
   'On Error GoTo errProc
   
   'If oDriver.SetValue(0, GetNextCode("CP_Model", "sModelIDx", True, oApp.Connection, True, oApp.BranchCode)) = False Then Exit Sub
   oDriver.FieldValue(0) = xeRecStateActive
   'oDriver.FieldReference(0) = True
   
   oDriver.FieldValue(1) = ""
   oDriver.FieldValue(2) = 0#
   oDriver.FieldValue(3) = 0#
   oDriver.FieldValue(4) = 0#
   oDriver.FieldValue(5) = 0#
   oDriver.FieldValue(6) = 0#
   oDriver.FieldValue(7) = 0#
   oDriver.FieldValue(8) = Encrypt(oApp.UserID)
   oDriver.FieldValue(7) = oApp.SysDate
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )"
End Sub

Private Sub oDriver_LoadOtherData()
   Dim lsSQL As String
   Dim loRS As Recordset
   
   Set loRS = New Recordset
   lsSQL = "SELECT sBankIDxx, sBankName FROM Banks WHERE sBankIDxx = " & strParm(oDriver.FieldValue(1))
   
   loRS.Open lsSQL, oApp.Connection, , , adCmdText
   Set loRS.ActiveConnection = Nothing
   
   If Not loRS.EOF Then txtField(1) = loRS("sBankName")
End Sub

Private Sub oDriver_WillSave(Cancel As Boolean)
   Dim Index As Integer
   
   For Index = 0 To 6
      Select Case Index
      Case 1
         If oDriver.FieldValue(Index) = "" Then
            MsgBox "Invalid Bank Detected!!!", vbCritical, "Warning"
            txtField(Index).SetFocus
            Cancel = True
            Exit Sub
         End If
      Case 2, 3, 4, 5, 6
         If oDriver.FieldValue(Index) = "" Then
            MsgBox "Invalid Discount Rate detected!!!", vbCritical, "Warning"
            txtField(Index).SetFocus
            Cancel = True
            Exit Sub
         End If
      End Select
   Next
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
      
      If Index = 9 Then
         .Text = Format(CDate(.Text), "Mmm dd, yyyy")
      End If
   End With
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   Dim lsOldProc As String
   
   lsOldProc = "txtField_LostFocus"
   'On Error GoTo errProc
      
   If Not IsDate(txtField(9)) Then txtField(9) = oApp.SysDate
     
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
