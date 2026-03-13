VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmCPCardRateModel 
   BorderStyle     =   0  'None
   Caption         =   "CP Credit Card Rates per Model"
   ClientHeight    =   5715
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame2 
      Height          =   915
      Left            =   1575
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   1614
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.CheckBox Check1 
         Caption         =   "Filter"
         Height          =   210
         Index           =   1
         Left            =   3180
         TabIndex        =   5
         Tag             =   "wt0;fb0"
         Top             =   487
         Width           =   690
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Filter"
         Height          =   210
         Index           =   0
         Left            =   3180
         TabIndex        =   2
         Tag             =   "wt0;fb0"
         Top             =   187
         Width           =   690
      End
      Begin VB.TextBox txtFilter 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   735
         TabIndex        =   4
         Top             =   450
         Width           =   2430
      End
      Begin VB.TextBox txtFilter 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   735
         TabIndex        =   1
         Top             =   150
         Width           =   2430
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Brand"
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
         Index           =   5
         Left            =   165
         TabIndex        =   3
         Top             =   510
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank"
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
         Index           =   4
         Left            =   165
         TabIndex        =   0
         Top             =   210
         Width           =   450
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   90
      TabIndex        =   31
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
      Picture         =   "frmCPCardRateModel.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   600
      Index           =   3
      Left            =   90
      TabIndex        =   30
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
      Picture         =   "frmCPCardRateModel.frx":077A
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   4140
      Index           =   0
      Left            =   1575
      Tag             =   "wt0;fb0"
      Top             =   1470
      Width           =   4020
      _ExtentX        =   7091
      _ExtentY        =   7303
      BorderStyle     =   1
      Begin VB.CheckBox chkActive 
         Caption         =   "Active"
         Height          =   285
         Left            =   3105
         TabIndex        =   32
         Top             =   3645
         Width           =   825
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   5
         Left            =   2025
         TabIndex        =   17
         Top             =   1980
         Width           =   1860
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   1350
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1275
         Width           =   2535
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   1350
         TabIndex        =   11
         Top             =   975
         Width           =   2535
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   4
         Left            =   2025
         TabIndex        =   15
         Top             =   1650
         Width           =   1860
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   6
         Left            =   2025
         TabIndex        =   19
         Top             =   2310
         Width           =   1860
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   7
         Left            =   2025
         TabIndex        =   21
         Top             =   2640
         Width           =   1860
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   8
         Left            =   2025
         TabIndex        =   23
         Top             =   2970
         Width           =   1860
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   9
         Left            =   2025
         TabIndex        =   25
         Top             =   3300
         Width           =   1860
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   1350
         TabIndex        =   9
         Top             =   675
         Width           =   2535
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
         Index           =   10
         Left            =   1500
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   165
         Width           =   1515
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Min. 12 Mos."
         Height          =   195
         Index           =   0
         Left            =   975
         TabIndex        =   16
         Top             =   2040
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model Code"
         Height          =   195
         Index           =   3
         Left            =   165
         TabIndex        =   12
         Top             =   1335
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model Name"
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
         TabIndex        =   10
         Top             =   1035
         Width           =   1065
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Min. 6 Mos."
         Height          =   195
         Index           =   2
         Left            =   975
         TabIndex        =   14
         Top             =   1710
         Width           =   825
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3 Mo. Term"
         Height          =   195
         Index           =   3
         Left            =   975
         TabIndex        =   18
         Top             =   2370
         Width           =   810
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "6 Mo. Term"
         Height          =   195
         Index           =   4
         Left            =   975
         TabIndex        =   20
         Top             =   2700
         Width           =   810
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "12 Mo. Term"
         Height          =   195
         Index           =   5
         Left            =   975
         TabIndex        =   22
         Top             =   3030
         Width           =   900
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "24 Mo. Term"
         Height          =   195
         Index           =   6
         Left            =   975
         TabIndex        =   24
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
         Width           =   1500
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
         TabIndex        =   8
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
         TabIndex        =   6
         Top             =   240
         Width           =   705
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   90
      TabIndex        =   26
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
      Picture         =   "frmCPCardRateModel.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   90
      TabIndex        =   27
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
      Picture         =   "frmCPCardRateModel.frx":166E
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5070
      Left            =   5610
      TabIndex        =   28
      Top             =   540
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   8943
      _Version        =   393216
      Appearance      =   0
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   90
      TabIndex        =   29
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
      Picture         =   "frmCPCardRateModel.frx":1DE8
   End
End
Attribute VB_Name = "frmCPCardRateModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmCPCardRateModel"

Private oSkin As clsFormSkin

Dim pnCtr As Integer, pnIndex As Integer
Dim pbGridFocus As Boolean, pbSave As Boolean, pbLoaded As Boolean

Dim psBrandIDx As String
Dim psBankIDxx As String
Dim psModelIDx As String

Dim psSQLMaster As String
Dim psSQLLookUp(2) As String
Dim poRSMaster As Recordset

Private Sub Check1_Click(Index As Integer)
   Dim lsFilter As String
   
   If Check1(0).Value = xeYes Then
      If txtFilter(0) <> "" Then lsFilter = "sBankIDxx = " & strParm(psBankIDxx)
   Else
      lsFilter = ""
   End If
   
   If Check1(1).Value = xeYes Then
      If lsFilter = "" Then
         If txtFilter(1) <> "" Then lsFilter = "sModelIDx = " & strParm(psModelIDx)
      Else
         If txtFilter(1) <> "" Then lsFilter = lsFilter & " AND sModelIDx = " & strParm(psModelIDx)
      End If
   End If
   
   If lsFilter = "" Then
      poRSMaster.Filter = ""
   Else
      poRSMaster.Filter = lsFilter
   End If
   InitForm
   reLoadDetail
End Sub

Private Sub cmdButton_Click(Index As Integer)
   Select Case Index
   Case 0 'save
      If SaveRecord Then
         MsgBox "Record Save Successfuly..", vbInformation, "Warning"
         Clearfields
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
         Clearfields
         LoadDetail
         initButton 0
      End If
   Case 4 'close
      Unload Me
   Case 5 'add
      If xrFrame1(0).Enabled Then
         With MSFlexGrid1
            If .TextMatrix(.Rows - 1, 10) <> "" Then
               .Rows = .Rows + 1
               If .Rows > 20 Then .TopRow = .Rows - 20
               .Row = .Rows - 1
               
               fillLastRow
               
               MSFlexGrid1_Click
               txtField(1).SetFocus
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
   ''On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualLeft

   InitForm
   Clearfields
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
      txtField(10) = .TextMatrix(.Row, 10)
      txtField(1) = .TextMatrix(.Row, 1)
      txtField(2) = .TextMatrix(.Row, 2)
      txtField(3) = .TextMatrix(.Row, 3)
      txtField(4) = .TextMatrix(.Row, 4)
      txtField(5) = .TextMatrix(.Row, 5)
      txtField(6) = .TextMatrix(.Row, 6)
      txtField(7) = .TextMatrix(.Row, 7)
      txtField(8) = .TextMatrix(.Row, 8)
      txtField(9) = .TextMatrix(.Row, 9)
      
      If pbLoaded And xrFrame1(0).Enabled Then
         If .Col < 4 Or .Col > 8 Then .Col = 4
         txtField(.Col).SetFocus
      End If
      
      .Col = 0
      .ColSel = .Cols - 1
   End With
   
   pbGridFocus = True
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With MSFlexGrid1
      Select Case Index
      Case 0
         If xrFrame1(0).Enabled Then txtField(1).SetFocus
      Case 1
         txtField(Index).Locked = Not txtField(9) = ""
         txtField(Index).Locked = Not (MSFlexGrid1.Row > poRSMaster.RecordCount)
      Case 2
         txtField(Index).Locked = Not .TextMatrix(.Row, 10) = ""
         txtField(Index).Locked = Not (MSFlexGrid1.Row > poRSMaster.RecordCount)
      End Select
   End With

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
   ''On Error GoTo errProc

   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
      Select Case Index
      Case 1 'search bank
         If searchBank(txtField(Index)) Then SetNextFocus
      Case 2 'search model
         If searchModel(txtField(Index)) Then SetNextFocus
      End Select
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

Private Function getFilter(ByVal Index As Integer, ByVal lsDescript As String) As Boolean
   Dim lsOldProc As String
   Dim lsSQL As String
   Dim lasDetail() As String
   
   lsOldProc = pxeMODULENAME & ".searchBrand"
   ''On Error GoTo errProc
   
   Select Case Index
   Case 0
      If txtFilter(0).Tag = lsDescript Then
         getFilter = True
         GoTo endProc
      End If
      
      lsSQL = AddCondition(psSQLLookUp(0), "sBankName LIKE " & strParm(lsDescript & "%"))
      
      lsSQL = KwikSearch(oApp, lsSQL, "sBankIDxx»sBankName", "Bank ID»Bank Name", "@»@")
      If lsSQL = "" Then
         psBankIDxx = ""
         GoTo endProc
      End If
      
      lasDetail = Split(lsSQL, "»")
      txtFilter(0) = lasDetail(1)
      txtFilter(0).Tag = lasDetail(1)
      psBankIDxx = lasDetail(0)
   Case 1
      If txtFilter(1).Tag = lsDescript Then
         getFilter = True
         GoTo endProc
      End If
      
      lsSQL = AddCondition(psSQLLookUp(1), "sModelNme LIKE " & strParm(lsDescript & "%"))
      Debug.Print lsSQL
      lsSQL = KwikSearch(oApp, lsSQL, "sModelIDx»sModelNme", "Mode ID»Model Name", "@»@")
      If lsSQL = "" Then
         psBankIDxx = ""
         GoTo endProc
      End If
      
      lasDetail = Split(lsSQL, "»")
      txtFilter(1) = lasDetail(2)
      txtFilter(1).Tag = lasDetail(2)
      psModelIDx = lasDetail(0)

'she 2016-04-21 lipat ko ung query sa modelname
      
'      lsSQL = AddCondition(psSQLLookUp(2), "sBrandNme LIKE " & strParm(lsDescript & "%"))
'
'      lsSQL = KwikSearch(oApp, lsSQL, "sBrandIDx»sBrandNme", "Brand ID»Brand Name", "@»@")
'      If lsSQL = "" Then
'         psBankIDxx = ""
'         GoTo endProc
'      End If
'
'      lasDetail = Split(lsSQL, "»")
'      txtFilter(1) = lasDetail(1)
'      txtFilter(1).Tag = lasDetail(1)
'      psBrandIDx = lasDetail(0)
   End Select
   
   getFilter = True
endProc:
   Exit Function
errProc:
   ShowError lsOldProc
End Function


Private Function searchBank(ByVal lsDescript As String) As Boolean
   Dim lsOldProc As String
   Dim lsSQL As String
   Dim lasDetail() As String
   
   lsOldProc = pxeMODULENAME & ".searchBank"
   ''On Error GoTo errProc
   
   If txtField(1).Tag = lsDescript Then
      searchBank = True
      GoTo endProc
   End If
   
   lsSQL = AddCondition(psSQLLookUp(0), "sBankName LIKE " & strParm(lsDescript & "%"))
   
   lsSQL = KwikSearch(oApp, lsSQL, "sBankIDxx»sBankName", "Bank ID»Bank Name", "@»@")
   If lsSQL = "" Then GoTo endProc
   
   lasDetail = Split(lsSQL, "»")
   txtField(10) = lasDetail(0)
   txtField(1) = lasDetail(1)
   txtField(1).Tag = txtField(1)
   
   With MSFlexGrid1
      .TextMatrix(.Row, 10) = txtField(10)
      .TextMatrix(.Row, 1) = txtField(1)
   End With
   
   searchBank = True
endProc:
   Exit Function
errProc:
   ShowError lsOldProc
End Function

Private Function searchModel(ByVal lsDescript As String) As Boolean
   Dim lsOldProc As String
   Dim lsSQL As String
   Dim lasDetail() As String
   
   lsOldProc = pxeMODULENAME & ".searchModel"
   ''On Error GoTo errProc
   
   If txtField(2).Tag = txtField(2) Then
      searchModel = True
      GoTo endProc
   End If
   
   lsSQL = AddCondition(psSQLLookUp(1), "a.sModelNme LIKE " & strParm(lsDescript & "%"))
   
   Debug.Print lsSQL
   lsSQL = KwikSearch(oApp, lsSQL, "sModelIDx»sBrandNme»sModelNme»sModelCde", _
                                    "Model ID»Brand»Model Name»Code", "@»@»@»@")
   If lsSQL = "" Then GoTo endProc
   
   lasDetail = Split(lsSQL, "»")
   txtField(2) = lasDetail(2)
   txtField(3) = lasDetail(3)
   txtField(2).Tag = lasDetail(2)
   
   With MSFlexGrid1
      .TextMatrix(.Row, 11) = lasDetail(0)
      .TextMatrix(.Row, 2) = txtField(2)
      .TextMatrix(.Row, 3) = txtField(3)
   End With
   
   searchModel = True
endProc:
   Exit Function
errProc:
   ShowError lsOldProc
End Function

Private Sub reLoadDetail()
   Dim lnCtr As Integer
   
   If poRSMaster.RecordCount = 0 Then
      fillLastRow
      Exit Sub
   End If
   
   With MSFlexGrid1
      .Rows = poRSMaster.RecordCount + 1
      For lnCtr = 0 To poRSMaster.RecordCount - 1
         .TextMatrix(lnCtr + 1, 0) = lnCtr + 1
         .TextMatrix(lnCtr + 1, 1) = poRSMaster("sBankName")
         .TextMatrix(lnCtr + 1, 2) = poRSMaster("sModelNme")
         .TextMatrix(lnCtr + 1, 3) = IFNull(poRSMaster("sModelCde"))
         .TextMatrix(lnCtr + 1, 4) = IIf(IFNull(poRSMaster("nMin6Monx"), 0) = 0, "0.00", Format(poRSMaster("nMin6Monx"), "#,###.#0"))
         .TextMatrix(lnCtr + 1, 5) = IIf(IFNull(poRSMaster("nMin12Mon"), 0) = 0, "0.00", Format(poRSMaster("nMin12Mon"), "#,###.#0"))
         .TextMatrix(lnCtr + 1, 6) = IIf(IFNull(poRSMaster("n03MoTerm"), 0) = 0, "0.00", Format(poRSMaster("n03MoTerm"), "#,###.#0"))
         .TextMatrix(lnCtr + 1, 7) = IIf(IFNull(poRSMaster("n06MoTerm"), 0) = 0, "0.00", Format(poRSMaster("n06MoTerm"), "#,###.#0"))
         .TextMatrix(lnCtr + 1, 8) = IIf(IFNull(poRSMaster("n12MoTerm"), 0) = 0, "0.00", Format(poRSMaster("n12MoTerm"), "#,###.#0"))
         .TextMatrix(lnCtr + 1, 9) = IIf(IFNull(poRSMaster("n24MoTerm"), 0) = 0, "0.00", Format(poRSMaster("n24MoTerm"), "#,###.#0"))
         .TextMatrix(lnCtr + 1, 10) = poRSMaster("sBankIDxx")
         .TextMatrix(lnCtr + 1, 11) = poRSMaster("sModelIDx")
         .TextMatrix(lnCtr + 1, 12) = poRSMaster("cRecdStat")
         poRSMaster.MoveNext
      Next
   End With
   
   MSFlexGrid1_Click
End Sub


Private Function LoadDetail() As Boolean
   Dim lsOldProc As String
   Dim lnCtr As Integer
   Dim lnRow As Integer
   
   lsOldProc = pxeMODULENAME & ".LoadDetail"
   
   If TypeName(poRSMaster) = "Nothing" Then
      Set poRSMaster = New Recordset
   End If
   
   If poRSMaster.State = adStateOpen Then poRSMaster.Close
   Debug.Print psSQLMaster
   poRSMaster.Open psSQLMaster, oApp.Connection, adOpenStatic, adLockOptimistic, adCmdText
   Set poRSMaster.ActiveConnection = Nothing
   
   With MSFlexGrid1
      If poRSMaster.EOF Then
         MsgBox "No Record Found.", vbCritical, "Warning"
         .Rows = 2
         
         initButton 1
         fillLastRow
      Else
         lnRow = poRSMaster.RecordCount
         
         .Rows = lnRow + 1
         
         If .Rows > 12 Then
            .ColWidth(1) = 2000
         Else
            .ColWidth(1) = 2000
         End If
      
         For lnCtr = 0 To lnRow - 1
            .TextMatrix(lnCtr + 1, 0) = lnCtr + 1
            .TextMatrix(lnCtr + 1, 1) = poRSMaster("sBankName")
            .TextMatrix(lnCtr + 1, 2) = poRSMaster("sModelNme")
            .TextMatrix(lnCtr + 1, 3) = IFNull(poRSMaster("sModelCde"))
            .TextMatrix(lnCtr + 1, 4) = IIf(IFNull(poRSMaster("nMin6Monx"), 0) = 0, "0.00", Format(poRSMaster("nMin6Monx"), "#,###.#0"))
            .TextMatrix(lnCtr + 1, 5) = IIf(IFNull(poRSMaster("nMin12Mon"), 0) = 0, "0.00", Format(poRSMaster("nMin12Mon"), "#,###.#0"))
            .TextMatrix(lnCtr + 1, 6) = IIf(IFNull(poRSMaster("n03MoTerm"), 0) = 0, "0.00", Format(poRSMaster("n03MoTerm"), "#,###.#0"))
            .TextMatrix(lnCtr + 1, 7) = IIf(IFNull(poRSMaster("n06MoTerm"), 0) = 0, "0.00", Format(poRSMaster("n06MoTerm"), "#,###.#0"))
            .TextMatrix(lnCtr + 1, 8) = IIf(IFNull(poRSMaster("n12MoTerm"), 0) = 0, "0.00", Format(poRSMaster("n12MoTerm"), "#,###.#0"))
            .TextMatrix(lnCtr + 1, 9) = IIf(IFNull(poRSMaster("n24MoTerm"), 0) = 0, "0.00", Format(poRSMaster("n24MoTerm"), "#,###.#0"))
            .TextMatrix(lnCtr + 1, 10) = poRSMaster("sBankIDxx")
            .TextMatrix(lnCtr + 1, 11) = poRSMaster("sModelIDx")
            .TextMatrix(lnCtr + 1, 12) = poRSMaster("cRecdStat")
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
   xrFrame2.Enabled = Not lbShow
   
   If pbLoaded Then
      If lbShow Then
         txtField(4).SetFocus
      Else
         txtFilter(0).SetFocus
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
      .TextMatrix(0, 2) = "Model Name"
      .TextMatrix(0, 3) = "Model Code"
      .TextMatrix(0, 4) = "6moMin."
      .TextMatrix(0, 5) = "12moMin."
      .TextMatrix(0, 6) = "3moRt"
      .TextMatrix(0, 7) = "6moRt"
      .TextMatrix(0, 8) = "12moRt"
      .TextMatrix(0, 9) = "24moRt"
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
      .ColWidth(1) = 1600
      .ColWidth(2) = 1600
      .ColWidth(3) = 1500
      .ColWidth(4) = 800
      .ColWidth(6) = 800
      .ColWidth(7) = 700
      .ColWidth(8) = 700
      .ColWidth(9) = 700
      .ColWidth(10) = 0
      .ColWidth(11) = 0
      .ColWidth(12) = 0

      .ColAlignment(1) = 1
      
      MSFlexGrid1_Click
   End With
End Sub

Private Sub Clearfields()
   Dim loTxt As TextBox
   
   For Each loTxt In txtField
      loTxt = ""
   Next
   
   psBankIDxx = ""
   psBrandIDx = ""
   psModelIDx = ""
   
   txtFilter(0) = ""
   txtFilter(1) = ""
   
   Check1(0).Value = xeNo
   Check1(1).Value = xeNo
   
   InitForm
End Sub

Private Function isEntryOK() As Boolean

EntryOK:
   isEntryOK = True
   Exit Function
EntryNotOK:
   isEntryOK = False
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
                     ", a.sModelIDx" & _
                     ", a.nMin6Monx" & _
                     ", a.nMin12Mon" & _
                     ", a.n03MoTerm" & _
                     ", a.n06MoTerm" & _
                     ", a.n12MoTerm" & _
                     ", a.n24MoTerm" & _
                     ", a.sApproved" & _
                     ", a.dPricexxx" & _
                     ", a.cRecdStat" & _
                     ", b.sBankName" & _
                     ", c.sModelCde" & _
                     ", c.sModelNme" & _
                     ", b.sBankIDxx" & _
                     ", d.sBrandIDx" & _
                  " FROM CP_Card_Rate_Model a" & _
                     " LEFT JOIN Banks b" & _
                        " ON a.sBankIDxx = b.sBankIDxx" & _
                     " LEFT JOIN CP_Model c" & _
                        " ON a.sModelIDx = c.sModelIDx" & _
                     " LEFT JOIN CP_Brand d" & _
                        " ON c.sBrandIDx = d.sBrandIDx" & _
                  " ORDER BY b.sBankName, d.sBrandNme, c.sModelNme"
                  
                     
   psSQLLookUp(0) = "SELECT" & _
                        "  sBankIDxx" & _
                        ", sBankName" & _
                     " FROM Banks" & _
                     " WHERE cRecdStat = " & strParm(xeRecStateActive) & _
                     " ORDER BY sBankName"
                     
   psSQLLookUp(1) = "SELECT" & _
                        "  a.sModelIDx" & _
                        ", b.sBrandNme" & _
                        ", a.sModelNme" & _
                        ", a.sModelCde" & _
                     " FROM CP_Model a" & _
                        " LEFT JOIN CP_Brand b" & _
                           " ON a.sBrandIDx = b.sBrandIDx" & _
                     " WHERE a.cRecdStat = " & strParm(xeRecStateActive) & _
                     " ORDER BY b.sBrandNme, a.sModelNme"
   
   psSQLLookUp(2) = "SELECT" & _
                        "  sBrandIDx" & _
                        ", sBrandNme" & _
                     " FROM CP_Brand" & _
                     " WHERE cRecdStat = " & strParm(xeRecStateActive) & _
                     " ORDER BY sBrandNme"
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
         If .TextMatrix(lnCtr, 10) <> "" Then
            poRSMaster.Filter = "sBankIDxx = " & strParm(.TextMatrix(lnCtr, 10)) & _
                              " AND sModelIDx = " & strParm(.TextMatrix(lnCtr, 11))
            
            If poRSMaster.EOF = False Then
               'any condition met means the record is modified, so save this entry
               If poRSMaster("nMin6Monx").OriginalValue <> CDbl(.TextMatrix(lnCtr, 4)) Or _
                  poRSMaster("nMin12Mon").OriginalValue <> CDbl(.TextMatrix(lnCtr, 5)) Or _
                  poRSMaster("n03MoTerm").OriginalValue <> CDbl(.TextMatrix(lnCtr, 6)) Or _
                  poRSMaster("n06MoTerm").OriginalValue <> CDbl(.TextMatrix(lnCtr, 7)) Or _
                  poRSMaster("n12MoTerm").OriginalValue <> CDbl(.TextMatrix(lnCtr, 8)) Or _
                  poRSMaster("n24MoTerm").OriginalValue <> CDbl(.TextMatrix(lnCtr, 9)) Then
                  
                  lsSQL = "UPDATE CP_Card_Rate_Model SET" & _
                           "  nMin6Monx = " & CDbl(.TextMatrix(lnCtr, 4)) & _
                           ", nMin12Mon = " & CDbl(.TextMatrix(lnCtr, 5)) & _
                           ", n03MoTerm = " & CDbl(.TextMatrix(lnCtr, 6)) & _
                           ", n06MoTerm = " & CDbl(.TextMatrix(lnCtr, 7)) & _
                           ", n12MoTerm = " & CDbl(.TextMatrix(lnCtr, 8)) & _
                           ", n24MoTerm = " & CDbl(.TextMatrix(lnCtr, 9)) & _
                           ", cRecdStat = " & strParm(xeYes) & _
                           ", dPricexxx = " & dateParm(oApp.ServerDate) & _
                           ", sApproved = " & strParm(Encrypt(oApp.UserID)) & _
                           ", sModified = " & strParm(Encrypt(oApp.UserID)) & _
                           ", dModified = " & dateParm(oApp.ServerDate) & _
                           " WHERE sBankIDxx = " & strParm(.TextMatrix(lnCtr, 10)) & _
                              " AND sModelIDx = " & strParm(.TextMatrix(lnCtr, 11))
                              
                  If oApp.Execute(lsSQL, "CP_Card_Rate_Model") = 0 Then GoTo endwithRoll
               End If
            Else
               'we have new entries, we create insert statements
               lsSQL = "INSERT INTO CP_Card_Rate_Model SET" & _
                           "  sBankIDxx = " & strParm(.TextMatrix(lnCtr, 10)) & _
                           ", sModelIDx = " & strParm(.TextMatrix(lnCtr, 11)) & _
                           ", nMin6Monx = " & CDbl(.TextMatrix(lnCtr, 4)) & _
                           ", nMin12Mon = " & CDbl(.TextMatrix(lnCtr, 5)) & _
                           ", n03MoTerm = " & CDbl(.TextMatrix(lnCtr, 6)) & _
                           ", n06MoTerm = " & CDbl(.TextMatrix(lnCtr, 7)) & _
                           ", n12MoTerm = " & CDbl(.TextMatrix(lnCtr, 8)) & _
                           ", n24MoTerm = " & CDbl(.TextMatrix(lnCtr, 9)) & _
                           ", dPricexxx = " & dateParm(oApp.ServerDate) & _
                           ", sApproved = " & strParm(Encrypt(oApp.UserID)) & _
                           ", cRecdStat = " & strParm(xeYes) & _
                           ", sModified = " & strParm(Encrypt(oApp.UserID)) & _
                           ", dModified = " & dateParm(oApp.ServerDate)
               Debug.Print lsSQL
               If oApp.Execute(lsSQL, "CP_Card_Rate_Model") = 0 Then GoTo endwithRoll
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
endwithRoll:
   oApp.RollbackTrans
   GoTo endProc
errProc:
   ShowError lsProcName
End Function

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
   Case 4, 5, 6, 7, 8, 9
      If Not IsNumeric(txtField(Index)) Then txtField(Index) = 0
         
      txtField(Index) = IIf(txtField(Index) = 0, "0.00", Format(txtField(Index), "#,###.#0"))
      
      With MSFlexGrid1
         .TextMatrix(.Row, Index) = txtField(Index)
      End With
   End Select
End Sub

Private Sub txtFilter_GotFocus(Index As Integer)
   With txtFilter(Index)
      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = oApp.getColor("HT1")
   End With
End Sub

Private Sub txtFilter_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyReturn, vbKeyF3
      If getFilter(Index, txtFilter(Index)) Then SetNextFocus
   End Select
   KeyCode = 0
End Sub

Private Sub txtFilter_LostFocus(Index As Integer)
   With txtFilter(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub fillLastRow()
   With MSFlexGrid1
      .TextMatrix(.Rows - 1, 0) = .Rows - 1
      .TextMatrix(.Rows - 1, 1) = ""
      .TextMatrix(.Rows - 1, 2) = ""
      .TextMatrix(.Rows - 1, 3) = ""
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
