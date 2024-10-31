VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmCPCardRate 
   BorderStyle     =   0  'None
   Caption         =   "CP Credit Card Rates"
   ClientHeight    =   5715
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17190
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   17190
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   90
      TabIndex        =   25
      Top             =   540
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
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
      Picture         =   "frmCPCardRate.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   600
      Index           =   3
      Left            =   90
      TabIndex        =   24
      Top             =   1770
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
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
      Picture         =   "frmCPCardRate.frx":077A
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   5070
      Index           =   0
      Left            =   1575
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   5010
      _ExtentX        =   8837
      _ExtentY        =   8943
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   10
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   990
         Width           =   3255
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   4
         Left            =   2700
         TabIndex        =   9
         Top             =   1980
         Width           =   2055
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   9
         Left            =   2700
         TabIndex        =   19
         Top             =   3630
         Width           =   2055
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   3
         Left            =   2700
         TabIndex        =   7
         Top             =   1650
         Width           =   2055
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   2700
         TabIndex        =   5
         Top             =   1320
         Width           =   2055
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   5
         Left            =   2700
         TabIndex        =   11
         Top             =   2310
         Width           =   2055
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   6
         Left            =   2700
         TabIndex        =   13
         Top             =   2640
         Width           =   2055
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   7
         Left            =   2700
         TabIndex        =   15
         Top             =   2970
         Width           =   2055
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   8
         Left            =   2700
         TabIndex        =   17
         Top             =   3300
         Width           =   2055
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   675
         Width           =   3255
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
         Height          =   285
         Index           =   0
         Left            =   1500
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   195
         Width           =   1680
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shop Type"
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
         TabIndex        =   27
         Top             =   1050
         Width           =   930
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Min. 24 Mos."
         Height          =   195
         Index           =   7
         Left            =   1650
         TabIndex        =   8
         Top             =   2040
         Width           =   915
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "36 Mo. Term"
         Height          =   195
         Index           =   1
         Left            =   1650
         TabIndex        =   18
         Top             =   3690
         Width           =   900
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Min. 12 Mos."
         Height          =   195
         Index           =   0
         Left            =   1650
         TabIndex        =   6
         Top             =   1710
         Width           =   915
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Min. 6 Mos."
         Height          =   195
         Index           =   2
         Left            =   1650
         TabIndex        =   4
         Top             =   1380
         Width           =   825
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3 Mo. Term"
         Height          =   195
         Index           =   3
         Left            =   1650
         TabIndex        =   10
         Top             =   2340
         Width           =   810
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "6 Mo. Term"
         Height          =   195
         Index           =   4
         Left            =   1650
         TabIndex        =   12
         Top             =   2700
         Width           =   810
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "12 Mo. Term"
         Height          =   195
         Index           =   5
         Left            =   1650
         TabIndex        =   14
         Top             =   3030
         Width           =   900
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "24 Mo. Term"
         Height          =   195
         Index           =   6
         Left            =   1650
         TabIndex        =   16
         Top             =   3360
         Width           =   900
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1590
         Tag             =   "et0;ht2"
         Top             =   285
         Width           =   1680
      End
      Begin VB.Label Label1 
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
         Index           =   0
         Left            =   165
         TabIndex        =   2
         Top             =   735
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Index           =   1
         Left            =   165
         TabIndex        =   0
         Top             =   240
         Width           =   705
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   90
      TabIndex        =   20
      Top             =   1770
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
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
      Picture         =   "frmCPCardRate.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   90
      TabIndex        =   21
      Top             =   540
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
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
      Picture         =   "frmCPCardRate.frx":166E
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5070
      Left            =   6585
      TabIndex        =   22
      Top             =   540
      Width           =   10500
      _ExtentX        =   18521
      _ExtentY        =   8943
      _Version        =   393216
      Appearance      =   0
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   90
      TabIndex        =   23
      Top             =   1155
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&ADD"
      AccessKey       =   "A"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCPCardRate.frx":1DE8
   End
End
Attribute VB_Name = "frmCPCardRate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmCPCardRate"

''''''''''''''''''''''''''''''''
'iMac
''''''''''''''''''''''''''''''''
'Harcode Rate for Major Cards
''''''''''''''''''''''''''''''''
Private Const pxeMajorCrd = "MCRXXX"
Private Const pxeMin06Mon = 0#
Private Const pxeMin12Mon = 0#
Private Const pxeMin24Mon = 0#
Private Const pxe03Months = 0#
Private Const pxe06Months = 7#
Private Const pxe12Months = 15#
Private Const pxe24Months = 0#
Private Const pxe36Months = 0#

Private oSkin As clsFormSkin

Dim pnCtr As Integer, pnIndex As Integer
Dim pbGridFocus As Boolean, pbSave As Boolean, pbLoaded As Boolean

Dim psSQLMaster As String
Dim psSQLBrowse As String
Dim psSQLLookUp As String
Dim poRSMaster As Recordset

'she 2020-10-06
Dim lsShopIdx As String
Dim lsShopTp As String


Private Sub cmdButton_Click(Index As Integer)
   Select Case Index
   Case 0 'save
      If SaveRecord Then
         MsgBox "Record Save Successfuly..", vbInformation, "Warning"
         LoadDetail
         initButton 0
      Else
         MsgBox "Unable to Save Records.", vbCritical, "Warning"
      End If
   Case 1 'update
      initButton 1
   Case 3 'cancel
      If MsgBox("This action will discard the updates made." & vbCrLf & _
                  "Do you want to continue?", vbQuestion & vbYesNo, "Confirm") = vbYes Then
         LoadDetail
         initButton 0
      End If
   Case 4 'close
      Unload Me
   Case 5 'add
      If xrFrame1(0).Enabled Then
         With MSFlexGrid1
            If .TextMatrix(.Rows - 1, 5) <> "" Or .TextMatrix(.Rows - 1, 6) <> "" Or .TextMatrix(.Rows - 1, 7) <> "" Then
               .Rows = .Rows + 1
               .Row = .Rows - 1
               
               fillLastRow
               
               MSFlexGrid1_Click
            End If
         End With
      End If
   End Select
End Sub

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0

   pbLoaded = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyReturn, vbKeyUp, vbKeyDown
      Select Case KeyCode
      Case vbKeyReturn, vbKeyDown
         If GetFocus = MSFlexGrid1.hwnd Then Exit Sub
         SetNextFocus
      Case vbKeyUp
         SetPreviousFocus
      End Select
   End Select
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String

   lsOldProc = "Form_Load"
   '''On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualLeft

   InitForm
   ClearFields
   initButton 0
   
   initSQL
   LoadDetail

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   pbLoaded = False
   Set poRSMaster = Nothing
End Sub

Private Sub MSFlexGrid1_Click()
   Dim lnCtr As Integer
   
   With MSFlexGrid1
      txtField(0) = .TextMatrix(.Row, 11)
      txtField(1) = .TextMatrix(.Row, 1)
      txtField(2) = .TextMatrix(.Row, 2)
      txtField(3) = .TextMatrix(.Row, 3)
      txtField(4) = .TextMatrix(.Row, 4)
      txtField(5) = .TextMatrix(.Row, 5)
      txtField(6) = .TextMatrix(.Row, 6)
      txtField(7) = .TextMatrix(.Row, 7)
      txtField(8) = .TextMatrix(.Row, 8)
      txtField(9) = .TextMatrix(.Row, 9)
      txtField(10) = .TextMatrix(.Row, 12)
      
      If pbLoaded And xrFrame1(0).Enabled Then txtField(.Col).SetFocus
      
      .Col = 0
      .ColSel = .Cols - 1
   End With
   
   pbGridFocus = True
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   Select Case Index
   Case 0
      If xrFrame1(0).Enabled Then txtField(1).SetFocus
   Case 1
      txtField(Index).Locked = Not txtField(0) = ""
      txtField(Index).Locked = Not (MSFlexGrid1.Row > poRSMaster.RecordCount)
   End Select

   With txtField(Index)
      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = oApp.getColor("HT1")
   End With
   
   pbGridFocus = False
   pnIndex = Index
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String
   lsOldProc = "txtField_KeyDown"
   '''On Error GoTo errProc

   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
      If Index = 1 Then
         If searchBank(txtField(Index)) Then
            txtField(2).SetFocus
         End If
      ElseIf Index = 10 Then
         If SearchShopType(txtField(Index)) Then
            txtField(10).SetFocus
         End If
      End If
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

Private Function searchBank(ByVal lsDescript As String) As Boolean
   Dim lsOldProc As String
   Dim lsSQL As String
   Dim lasDetail() As String
   
   lsOldProc = pxeMODULENAME & ".searchBank"
   '''On Error GoTo errProc
   
   If txtField(1).Tag = lsDescript Then
      searchBank = True
      GoTo endProc
   End If
   
   lsSQL = AddCondition(psSQLLookUp, "sBankName LIKE " & strParm(lsDescript & "%"))
   
   lsSQL = KwikSearch(oApp, lsSQL, "sBankIDxx»sBankName", "Bank ID»Bank Name", "@»@")
   If lsSQL = "" Then GoTo endProc
   
   lasDetail = Split(lsSQL, "»")
   txtField(0) = lasDetail(0)
   txtField(1) = lasDetail(1)
   txtField(1).Tag = txtField(1)
   
   With MSFlexGrid1
      .TextMatrix(.Row, 1) = txtField(1)
      .TextMatrix(.Row, 2) = "0.00"
      .TextMatrix(.Row, 3) = "0.00"
      .TextMatrix(.Row, 4) = "0.00"
      .TextMatrix(.Row, 5) = "0.00"
      .TextMatrix(.Row, 6) = "0.00"
      .TextMatrix(.Row, 7) = "0.00"
      .TextMatrix(.Row, 8) = "0.00"
      .TextMatrix(.Row, 9) = "0.00"
      .TextMatrix(.Row, 10) = ""
      .TextMatrix(.Row, 11) = lasDetail(0)
   End With
endProc:
   Exit Function
errProc:
   ShowError lsOldProc
End Function

Private Function LoadDetail() As Boolean
   Dim lsOldProc As String
   Dim lnCtr As Integer
   Dim lnRow As Integer
   
   lsOldProc = pxeMODULENAME & ".LoadDetail"
   
   If TypeName(poRSMaster) = "Nothing" Then
      Set poRSMaster = New Recordset
   End If
   
   Debug.Print psSQLMaster
   If poRSMaster.State = adStateOpen Then poRSMaster.Close
   poRSMaster.Open psSQLMaster, oApp.Connection, adOpenStatic, adLockOptimistic, adCmdText
   Set poRSMaster.ActiveConnection = Nothing
   
   Debug.Print psSQLMaster
   With MSFlexGrid1
      If poRSMaster.EOF Then
         MsgBox "No Record Found.", vbCritical, "Warning"
         .Rows = 2
         
         fillLastRow
      Else
         lnRow = poRSMaster.RecordCount
         
         .Rows = lnRow + 1
         
         If .Rows > 12 Then
            .ColWidth(1) = 2150
         Else
            .ColWidth(1) = 2400
         End If
      
         For lnCtr = 0 To lnRow - 1
            .TextMatrix(lnCtr + 1, 0) = lnCtr + 1
            .TextMatrix(lnCtr + 1, 1) = IFNull(poRSMaster("sBankName"), "")
            .TextMatrix(lnCtr + 1, 2) = IIf(IFNull(poRSMaster("nMin6Monx"), 0) = 0, "0.00", Format(poRSMaster("nMin6Monx"), "#,###.#0"))
            .TextMatrix(lnCtr + 1, 3) = IIf(IFNull(poRSMaster("nMin12Mon"), 0) = 0, "0.00", Format(poRSMaster("nMin12Mon"), "#,###.#0"))
            .TextMatrix(lnCtr + 1, 4) = IIf(IFNull(poRSMaster("nMin24Mon"), 0) = 0, "0.00", Format(poRSMaster("nMin24Mon"), "#,###.#0"))
            .TextMatrix(lnCtr + 1, 5) = IIf(IFNull(poRSMaster("n03MoTerm"), 0) = 0, "0.00", Format(poRSMaster("n03MoTerm"), "#,###.#0"))
            .TextMatrix(lnCtr + 1, 6) = IIf(IFNull(poRSMaster("n06MoTerm"), 0) = 0, "0.00", Format(poRSMaster("n06MoTerm"), "#,###.#0"))
            .TextMatrix(lnCtr + 1, 7) = IIf(IFNull(poRSMaster("n12MoTerm"), 0) = 0, "0.00", Format(poRSMaster("n12MoTerm"), "#,###.#0"))
            .TextMatrix(lnCtr + 1, 8) = IIf(IFNull(poRSMaster("n24MoTerm"), 0) = 0, "0.00", Format(poRSMaster("n24MoTerm"), "#,###.#0"))
            .TextMatrix(lnCtr + 1, 9) = IIf(IFNull(poRSMaster("n36MoTerm"), 0) = 0, "0.00", Format(poRSMaster("n36MoTerm"), "#,###.#0"))
            .TextMatrix(lnCtr + 1, 10) = poRSMaster("sBrandIDx")
            .TextMatrix(lnCtr + 1, 11) = poRSMaster("sBankIDxx")
            .TextMatrix(lnCtr + 1, 12) = IFNull(poRSMaster("sBrandNme"), "")
            poRSMaster.MoveNext
         Next
      End If
   End With
   
   MSFlexGrid1_Click
endProc:
   Exit Function
errProc:
   ShowError (lsOldProc)
End Function

Private Sub initButton(lnStat As Integer)
   Dim lbShow As Boolean

   lbShow = IIf(lnStat = 0, False, True)

   cmdButton(1).Visible = Not lbShow
   cmdButton(4).Visible = Not lbShow
   cmdButton(3).Visible = lbShow
   cmdButton(0).Visible = lbShow

   xrFrame1(0).Enabled = lbShow
   
   If pbLoaded Then
      If lbShow Then
         txtField(2).SetFocus
      Else
         cmdButton(1).SetFocus
      End If
   End If
End Sub

Private Sub InitForm()
   Dim lnCtr As Integer
   
   With MSFlexGrid1
      .Clear
   
      .Cols = 13
      .Rows = 2
      .Font = "MS Sans Serif"

      'Column Title
      .TextMatrix(0, 1) = "Bank Name"
      .TextMatrix(0, 2) = "6mo Min."
      .TextMatrix(0, 3) = "12mo Min."
      .TextMatrix(0, 4) = "24mo Min."
      .TextMatrix(0, 5) = "3mo Rt"
      .TextMatrix(0, 6) = "6mo Rt"
      .TextMatrix(0, 7) = "12mo Rt"
      .TextMatrix(0, 8) = "24mo Rt"
      .TextMatrix(0, 9) = "36mo Rt"
      .TextMatrix(0, 10) = ""
      .TextMatrix(0, 11) = ""
      .TextMatrix(0, 12) = ""

      .Row = 0

      'Column Alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = 3
      Next
      .Row = 1

      'Column Width
      .ColWidth(0) = 380
      .ColWidth(1) = 2400
      .ColWidth(2) = 1100
      .ColWidth(3) = 1100
      .ColWidth(4) = 1100
      .ColWidth(5) = 1100
      .ColWidth(6) = 1100
      .ColWidth(7) = 1100
      .ColWidth(8) = 1100
      .ColWidth(9) = 1100
      .ColWidth(10) = 0
      .ColWidth(11) = 0
      .ColWidth(12) = 0

      .ColAlignment(1) = 1
      
      MSFlexGrid1_Click
   End With
End Sub

Private Sub ClearFields()
   Dim loTxt As TextBox
   
   For Each loTxt In txtField
      loTxt = ""
   Next
   
   InitForm
End Sub

Private Function isEntryOk() As Boolean

EntryOK:
   isEntryOk = True
   Exit Function
EntryNotOK:
   isEntryOk = False
End Function

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
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

Private Sub initSQL()
   psSQLMaster = "SELECT" & _
                     "  a.sBankIDxx" & _
                     ", a.sBrandIdx" & _
                     ", b.sBankName" & _
                     ", a.nMin6Monx" & _
                     ", a.nMin12Mon" & _
                     ", a.nMin24Mon" & _
                     ", a.n03MoTerm" & _
                     ", a.n06MoTerm" & _
                     ", a.n12MoTerm" & _
                     ", a.n24MoTerm" & _
                     ", a.n36MoTerm" & _
                     ", c.sBrandNme" & _
                     ", a.sApproved" & _
                     ", a.cRecdStat" & _
                     ", a.sModified" & _
                     ", a.dModified" & _
                  " FROM CP_Card_Rate a" & _
                     " LEFT JOIN Banks b ON a.sBankIDxx = b.sBankIDxx" & _
                     " LEFT JOIN CP_Brand c ON a.sBrandIDx = c.sBrandIdx" & _
                  " WHERE a.cRecdStat = " & strParm(xeRecStateActive) & _
                  " ORDER BY b.sBankName"
   
   psSQLBrowse = "SELECT" & _
                        "  a.sBankIDxx" & _
                        ", b.sBankName" & _
                     " FROM CP_Card_Rate a" & _
                        " LEFT JOIN Banks b" & _
                           " ON a.sBankIDxx = b.sBankIDxx" & _
                     " WHERE a.cRecdStat = " & strParm(xeRecStateActive) & _
                     " ORDER BY b.sBankName"
                     
   psSQLLookUp = "SELECT" & _
                        "  sBankIDxx" & _
                        ", sBankName" & _
                     " FROM Banks" & _
                     " WHERE cRecdStat = " & strParm(xeRecStateActive) & _
                     " ORDER BY sBankName"
    'she 2020-10-06
'   " AND sBankIDxx" & _  NOT IN (SELECT" & _
                                                   " sBankIDxx" & _
                                                   " FROM CP_Card_Rate" & _
                                                   " WHERE cRecdStat = '1')"
End Sub

Private Function SaveRecord() As Boolean
   Dim lsProcName As String
   Dim lnCtr As Integer
   Dim lsSQL As String

   lsProcName = pxeMODULENAME & ".SaveRecord"
   '''On Error GoTo errProc
   
   If TypeName(poRSMaster) = "Nothing" Then GoTo endProc
   
   With MSFlexGrid1
      oApp.BeginTrans
      poRSMaster.Filter = ""
      If poRSMaster.RecordCount > 0 Then poRSMaster.MoveFirst
      
      For lnCtr = 1 To .Rows - 1
         If .TextMatrix(.Rows - 1, 2) <> "" Or _
            .TextMatrix(.Rows - 1, 3) <> "" Or _
            .TextMatrix(.Rows - 1, 4) <> "" Or _
            .TextMatrix(.Rows - 1, 5) <> "" Or _
            .TextMatrix(.Rows - 1, 6) <> "" Or _
            .TextMatrix(.Rows - 1, 7) <> "" Or _
            .TextMatrix(.Rows - 1, 8) <> "" Or _
            .TextMatrix(.Rows - 1, 9) <> "" Then
            
            poRSMaster.Filter = "sBankIDxx = " & strParm(.TextMatrix(lnCtr, 11)) & " AND sBrandIDx = " & strParm(.TextMatrix(lnCtr, 10))
            
            If poRSMaster.EOF = False Then
               If poRSMaster("sBankIDxx") = pxeMajorCrd Then
                  'any condition met means the record is modified, so save this entry
                  If poRSMaster("nMin6Monx").OriginalValue <> CDbl(pxeMin06Mon) Or _
                     poRSMaster("nMin12Mon").OriginalValue <> CDbl(pxeMin12Mon) Or _
                     poRSMaster("nMin24Mon").OriginalValue <> CDbl(pxeMin24Mon) Or _
                     poRSMaster("n03MoTerm").OriginalValue <> CDbl(pxe03Months) Or _
                     poRSMaster("n06MoTerm").OriginalValue <> CDbl(pxe06Months) Or _
                     poRSMaster("n12MoTerm").OriginalValue <> CDbl(pxe12Months) Or _
                     poRSMaster("n24MoTerm").OriginalValue <> CDbl(pxe24Months) Or _
                     poRSMaster("n36MoTerm").OriginalValue <> CDbl(pxe36Months) Then
                       
                        lsSQL = "UPDATE CP_Card_Rate SET" & _
                                 "  nMin6Monx = " & CDbl(pxeMin06Mon) & _
                                 ", nMin12Mon = " & CDbl(pxeMin12Mon) & _
                                 ", nMin24Mon = " & CDbl(pxeMin24Mon) & _
                                 ", n03MoTerm = " & CDbl(pxe03Months) & _
                                 ", n06MoTerm = " & CDbl(pxe06Months) & _
                                 ", n12MoTerm = " & CDbl(pxe12Months) & _
                                 ", n24MoTerm = " & CDbl(pxe24Months) & _
                                 ", n36MoTerm = " & CDbl(pxe36Months) & _
                                 ", sApproved = " & strParm(Encrypt(oApp.UserID)) & _
                                 ", sModified = " & strParm(Encrypt(oApp.UserID)) & _
                                 ", dModified = " & dateParm(oApp.ServerDate) & _
                                 " WHERE sBankIDxx = " & strParm(pxeMajorCrd)
                     
                        If oApp.Execute(lsSQL, "CP_Card_Rate") = 0 Then GoTo endWithRoll
                    End If
               Else
                  'any condition met means the record is modified, so save this entry
                  If poRSMaster("nMin6Monx").OriginalValue <> CDbl(.TextMatrix(lnCtr, 2)) Or _
                     poRSMaster("nMin12Mon").OriginalValue <> CDbl(.TextMatrix(lnCtr, 3)) Or _
                     IFNull(poRSMaster("nMin24Mon").OriginalValue, 0) <> CDbl(.TextMatrix(lnCtr, 4)) Or _
                     poRSMaster("n03MoTerm").OriginalValue <> CDbl(.TextMatrix(lnCtr, 5)) Or _
                     poRSMaster("n06MoTerm").OriginalValue <> CDbl(.TextMatrix(lnCtr, 6)) Or _
                     poRSMaster("n12MoTerm").OriginalValue <> CDbl(.TextMatrix(lnCtr, 7)) Or _
                     poRSMaster("n24MoTerm").OriginalValue <> CDbl(.TextMatrix(lnCtr, 8)) Or _
                     poRSMaster("n36MoTerm").OriginalValue <> CDbl(.TextMatrix(lnCtr, 9)) Then
                     
                
                     lsSQL = "UPDATE CP_Card_Rate SET" & _
                                "  nMin6Monx = " & CDbl(.TextMatrix(lnCtr, 2)) & _
                                ", nMin12Mon = " & CDbl(.TextMatrix(lnCtr, 3)) & _
                                ", nMin24Mon = " & CDbl(.TextMatrix(lnCtr, 4)) & _
                                ", n03MoTerm = " & CDbl(.TextMatrix(lnCtr, 5)) & _
                                ", n06MoTerm = " & CDbl(.TextMatrix(lnCtr, 6)) & _
                                ", n12MoTerm = " & CDbl(.TextMatrix(lnCtr, 7)) & _
                                ", n24MoTerm = " & CDbl(.TextMatrix(lnCtr, 8)) & _
                                ", n36MoTerm = " & CDbl(.TextMatrix(lnCtr, 9)) & _
                                ", sBrandIDx = " & strParm(.TextMatrix(lnCtr, 10)) & _
                                ", sApproved = " & strParm(oApp.UserID) & _
                                ", sModified = " & strParm(oApp.UserID) & _
                                ", dModified = " & dateParm(oApp.ServerDate) & _
                              " WHERE sBankIDxx = " & strParm(.TextMatrix(lnCtr, 11)) & _
                              " AND sBrandIDx = " & strParm(.TextMatrix(lnCtr, 10))
                     
                      If oApp.Execute(lsSQL, "CP_Card_Rate") = 0 Then GoTo endWithRoll
                    End If
                End If
            Else
               'we have new entries, we create insert statements
               lsSQL = "INSERT INTO CP_Card_Rate SET" & _
                           "  sBankIDxx = " & strParm(.TextMatrix(lnCtr, 11)) & _
                           ", sBrandIdx = " & strParm(.TextMatrix(lnCtr, 10)) & _
                           ", nMin6Monx = " & CDbl(.TextMatrix(lnCtr, 2)) & _
                           ", nMin12Mon = " & CDbl(.TextMatrix(lnCtr, 3)) & _
                           ", nMin24Mon = " & CDbl(.TextMatrix(lnCtr, 4)) & _
                           ", n03MoTerm = " & CDbl(.TextMatrix(lnCtr, 5)) & _
                           ", n06MoTerm = " & CDbl(.TextMatrix(lnCtr, 6)) & _
                           ", n12MoTerm = " & CDbl(.TextMatrix(lnCtr, 7)) & _
                           ", n24MoTerm = " & CDbl(.TextMatrix(lnCtr, 8)) & _
                           ", n36MoTerm = " & CDbl(.TextMatrix(lnCtr, 9)) & _
                           ", sApproved = " & strParm(oApp.UserID) & _
                           ", cRecdStat = " & strParm(xeRecStateActive) & _
                           ", sModified = " & strParm(oApp.UserID) & _
                           ", dModified = " & dateParm(oApp.ServerDate)
               Debug.Print lsSQL
               If oApp.Execute(lsSQL, "CP_Card_Rate") = 0 Then GoTo endWithRoll
            End If
         End If
         poRSMaster.Filter = ""
      Next
      
      oApp.CommitTrans
   End With
   
   SaveRecord = True
   Set poRSMaster = Nothing
endProc:
   Exit Function
endWithRoll:
   oApp.RollbackTrans
   GoTo endProc
errProc:
   ShowError lsProcName
End Function

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   With MSFlexGrid1
      If .TextMatrix(.Row, 11) = pxeMajorCrd Then
         MSFlexGrid1_Click
         Exit Sub
      End If
      
      Select Case Index
      Case 2, 3, 4, 5, 6, 7, 8, 9
         If Not IsNumeric(txtField(Index)) Then txtField(Index) = 0
            
         txtField(Index) = IIf(txtField(Index) = 0, "0.00", Format(txtField(Index), "#,##0.00"))
         .TextMatrix(.Row, Index) = txtField(Index)
      End Select
   End With
End Sub

Private Sub fillLastRow()
   With MSFlexGrid1
      .TextMatrix(.Rows - 1, 0) = .Rows - 1
      .TextMatrix(.Rows - 1, 1) = ""
      .TextMatrix(.Rows - 1, 2) = "0.00"
      .TextMatrix(.Rows - 1, 3) = "0.00"
      .TextMatrix(.Rows - 1, 4) = "0.00"
      .TextMatrix(.Rows - 1, 5) = "0.00"
      .TextMatrix(.Rows - 1, 6) = "0.00"
      .TextMatrix(.Rows - 1, 7) = "0.00"
      .TextMatrix(.Rows - 1, 8) = "0.00"
      .TextMatrix(.Rows - 1, 9) = "0.00"
      .TextMatrix(.Rows - 1, 10) = ""
      .TextMatrix(.Rows - 1, 11) = ""
   End With
End Sub

'she 2020-10-06
'special treatment for city bank
Private Function SearchShopType(ByVal lsDescript As String) As Boolean
   Dim lsOldProc As String
   Dim lsSQL As String
   Dim lasDetail() As String
   Dim loShopType As Recordset
   
   lsOldProc = pxeMODULENAME & ".SearchShopType"
   '''On Error GoTo errProc

   
   lsSQL = "SELECT sBrandIDx, sBrandNme " & _
            " FROM CP_Brand" & _
            " WHERE sBrandNme LIKE " & strParm(lsDescript + "%")
   Debug.Print lsSQL
   Set loShopType = New Recordset
   loShopType.Open lsSQL, oApp.Connection, , , adCmdText
   
   If Not loShopType.EOF Then
        lsSQL = KwikSearch(oApp, lsSQL, "sBrandIDx»sBrandNme", "ID»Shop", "@»@")
        If lsSQL = "" Then GoTo endProc
        
        lasDetail = Split(lsSQL, "»")
        txtField(10) = lasDetail(1) & " CONCEPT"
        txtField(10).Tag = txtField(10)
        
        lsShopIdx = lasDetail(0)
        lsShopTp = lasDetail(1) & " CONCEPT"
   Else
        lsShopIdx = ""
        lsShopTp = ""
   End If
  
  With MSFlexGrid1
      .TextMatrix(.Row, 10) = lasDetail(0)
   End With
   
endProc:
   Exit Function
errProc:
   ShowError lsOldProc
End Function
