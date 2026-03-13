VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0"; "xrGridControl.ocx"
Begin VB.Form frmUnpostedTransaction 
   BorderStyle     =   0  'None
   Caption         =   "UNENCODED TRANSACTION"
   ClientHeight    =   8115
   ClientLeft      =   0
   ClientTop       =   15
   ClientWidth     =   12585
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8115
   ScaleWidth      =   12585
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   3780
      Index           =   0
      Left            =   1680
      Tag             =   "wt0;fb0"
      Top             =   585
      Width           =   10680
      _ExtentX        =   18838
      _ExtentY        =   6668
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   4
         Left            =   7365
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   1380
         Width           =   2760
      End
      Begin VB.OptionButton optCashPaym 
         Caption         =   "Check"
         Height          =   195
         Index           =   1
         Left            =   3000
         TabIndex        =   17
         Tag             =   "wt0;fb0"
         Top             =   1785
         Width           =   1005
      End
      Begin VB.OptionButton optCashPaym 
         Caption         =   "Cash"
         Height          =   195
         Index           =   0
         Left            =   1725
         TabIndex        =   16
         Tag             =   "wt0;fb0"
         Top             =   1785
         Value           =   -1  'True
         Width           =   1005
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1140
         Index           =   3
         Left            =   1680
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   2550
         Width           =   8820
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   2
         Left            =   1680
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   2100
         Width           =   2760
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   1
         Left            =   1680
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1245
         Width           =   2760
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   0
         Left            =   1680
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   315
         Width           =   2760
      End
      Begin VB.ComboBox cmbTranType 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmUnpostedTransaction.frx":0000
         Left            =   1680
         List            =   "frmUnpostedTransaction.frx":0002
         TabIndex        =   14
         Top             =   840
         Width           =   2775
      End
      Begin VB.Shape Shape4 
         Height          =   330
         Index           =   0
         Left            =   7395
         Top             =   870
         Width           =   2700
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "UNKNOWN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   7365
         TabIndex        =   21
         Tag             =   "eb0;et0"
         Top             =   900
         Width           =   2760
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H8000000F&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   300
         Index           =   0
         Left            =   7425
         Tag             =   "et0;et0"
         Top             =   885
         Width           =   2655
      End
      Begin VB.Shape Shape3 
         Height          =   390
         Index           =   2
         Left            =   7365
         Top             =   840
         Width           =   2760
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   6420
         TabIndex        =   20
         Top             =   1380
         Width           =   810
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "REMARKS:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   480
         TabIndex        =   15
         Top             =   2610
         Width           =   1035
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AMOUNT:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   480
         TabIndex        =   13
         Top             =   2175
         Width           =   930
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "REFER NO:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   480
         TabIndex        =   12
         Top             =   1350
         Width           =   1080
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TRAN. TYPE:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   6
         Top             =   930
         Width           =   1005
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DATE:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   480
         TabIndex        =   0
         Top             =   405
         Width           =   690
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   105
      TabIndex        =   7
      Top             =   1350
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
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
      Picture         =   "frmUnpostedTransaction.frx":0004
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   105
      TabIndex        =   8
      Top             =   720
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
      Picture         =   "frmUnpostedTransaction.frx":077E
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   105
      TabIndex        =   9
      Top             =   1965
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
      Picture         =   "frmUnpostedTransaction.frx":0EF8
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   105
      TabIndex        =   10
      Top             =   720
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
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
      Picture         =   "frmUnpostedTransaction.frx":1672
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   105
      TabIndex        =   11
      Top             =   1965
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
      Picture         =   "frmUnpostedTransaction.frx":1DEC
   End
   Begin xrGridEditor.GridEditor GridEditor1 
      Height          =   3630
      Left            =   1695
      TabIndex        =   5
      Tag             =   "et0;eb0;et0;bc2"
      Top             =   4395
      Width           =   10650
      _ExtentX        =   18785
      _ExtentY        =   6403
      AllowBigSelection=   -1  'True
      AutoAdd         =   -1  'True
      AutoNumber      =   -1  'True
      BACKCOLOR       =   -2147483643
      BACKCOLORBKG    =   8421504
      BACKCOLORFIXED  =   -2147483633
      BACKCOLORSEL    =   -2147483635
      BORDERSTYLE     =   1
      COLS            =   2
      FILLSTYLE       =   0
      FIXEDCOLS       =   1
      FIXEDROWS       =   1
      FOCUSRECT       =   1
      EDITORBACKCOLOR =   -2147483643
      EDITORFORECOLOR =   -2147483640
      FORECOLOR       =   -2147483640
      FORECOLORFIXED  =   -2147483630
      FORECOLORSEL    =   -2147483634
      FORMATSTRING    =   ""
      Object.HEIGHT          =   3630
      GRIDCOLOR       =   12632256
      GRIDCOLORFIXED  =   0
      BeginProperty GRIDFONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GRIDLINES       =   1
      GRIDLINESFIXED  =   2
      GRIDLINEWIDTH   =   1
      MOUSEICON       =   "frmUnpostedTransaction.frx":2566
      MOUSEPOINTER    =   0
      REDRAW          =   -1  'True
      RIGHTTOLEFT     =   0   'False
      ROWS            =   2
      SCROLLBARS      =   3
      SCROLLTRACK     =   0   'False
      SELECTIONMODE   =   0
      Object.TOOLTIPTEXT     =   ""
      WORDWRAP        =   0   'False
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   6
      Left            =   105
      TabIndex        =   18
      Top             =   1350
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
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
      Picture         =   "frmUnpostedTransaction.frx":2582
   End
End
Attribute VB_Name = "frmUnpostedTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Const pxeMODULENAME = "frmUnpostedTransaction"
Private oSkin As clsFormSkin

Dim pbGridFocus As Boolean
Dim pnCtr As Integer, pnIndex As Integer
Dim pnDTREntryNo As Integer
Dim lsTranType As String
Dim lbAllValueOkay As String
Dim pbLoad As Boolean

Private Sub cmbTranType_Click()
   GridEditor1.TextMatrix(pnCtr, 1) = cmbTranType.Text
End Sub

Private Sub cmbTranType_KeyPress(keyascii As Integer)
   keyascii = 0
        Exit Sub
End Sub

Private Sub cmdButton_Click(Index As Integer)
   Dim lnCtr As Integer
   Dim lnRep As Integer
   Dim lsSQL As String

   With GridEditor1
      Select Case Index
      Case 0 'SAVE
         For lnCtr = 1 To .Rows - 1
            If .Rows > 2 And .TextMatrix(lnCtr, 1) = "" And .TextMatrix(lnCtr, 2) = "" And .TextMatrix(lnCtr, 4) = "" Then
               .Rows = .Rows - 1
            End If
         Next
         
         If SaveDTRTrans = True Then
            MsgBox "Transaction save Successfully!!!", vbInformation, "Confirm"
         Else
            MsgBox "Unable to save Transaction!!!", vbCritical, "NOTICE"
            GoTo endProc
         End If
         
         initButton (False)
         isFieldEnable (False)
         InitFields
         InitGrid
      Case 1 'Delete row
         With GridEditor1
            .deleteRow
         End With
      Case 2 'cancel
         initButton (False)
         isFieldEnable (False)
         txtField(0).Enabled = False
         InitFields
         InitGrid
         pbLoad = False
      Case 3 'new
         initButton (True)
         isFieldEnable (False)
         ClearFields
         InitGrid
         InitFields
         pbLoad = False
      Case 4
         Unload Me
      Case 5 'Confirm
         If pbLoad = True And Label2.Caption = "OPEN" Then
            lnRep = MsgBox("Are you sure you want to Confirm this Transaction???", vbQuestion + vbYesNo, "CONFIRM")
            If lnRep = vbYes Then
               lsSQL = "UPDATE DTR_Summary SET cPostedxx = '1'" & _
                        " WHERE sBranchCd = " & strParm(oApp.BranchCode) & _
                        " AND sTranDate = " & strParm(txtField(0).Text)
               If oApp.Execute(lsSQL, "DTR_Summary") <= 0 Then
                  Debug.Print lsSQL
                  MsgBox "Unable to Save Transaction!"
                  GoTo endProc
               End If
               MsgBox "Transaction Confirmed!!"
            End If
         Else
            MsgBox "Transaction was already " & Label2.Caption
                  GoTo endProc
         End If
      Case 6 'Browse
         Call SearchTransaction
      End Select
   End With
endProc:
   Exit Sub
errProc:
'   ShowError lsOldProc, True
End Sub

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0

   With GridEditor1
      .Refresh
   End With
End Sub

Private Sub Form_Load()
   CenterChildForm mdiMain, Me

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualLeft
   
   InitFields
   InitGrid
   initButton (True)
   Call isFieldEnable(False)
   
   cmbTranType.ListIndex = 0
   
   pbLoad = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyReturn, vbKeyUp, vbKeyDown
      Select Case KeyCode
      Case vbKeyReturn, vbKeyDown
         If GetFocus = GridEditor1.hwnd Then Exit Sub
         SetNextFocus
      Case vbKeyUp
         SetPreviousFocus
      End Select
   Case vbKeyF12
   End Select
End Sub

Private Sub GridEditor1_RowColChange()
   With GridEditor1
         cmbTranType.Text = .TextMatrix(.Row, 1)
         txtField(1).Text = .TextMatrix(.Row, 2)
         txtField(2).Text = .TextMatrix(.Row, 3)
         txtField(3).Text = .TextMatrix(.Row, 4)
         If .TextMatrix(.Row, 5) = 1 Then
            optCashPaym(0).Value = Checked
            optCashPaym(1).Value = Unchecked
         Else
            optCashPaym(0).Value = Unchecked
            optCashPaym(1).Value = Checked
         End If
   End With
      
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = oApp.getColor("HT1")
   End With

   pbGridFocus = False
   pnIndex = Index
   Call setGridValue
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

Private Sub InitFields()

   txtField(0).Text = Format(oApp.ServerDate, "MMM DD, YYYY")
   txtField(1).Text = ""
   txtField(2).Text = "0.00"
   txtField(3).Text = ""
   txtField(4).Text = "0.00"
   
   cmbTranType.AddItem "CP Sales"
   cmbTranType.AddItem "Misc"
   cmbTranType.AddItem "Monthly Payment"
   cmbTranType.AddItem "CP Load"
   
End Sub

Function VerifyDTRDate(lsDate As String) As Boolean
   Dim lsSQL As String
   Dim loData As Recordset
   
   lsSQL = "SELECT b.dUnenCode, a.* " & _
            " FROM DTR_Summary a" & _
            ", Branch_Others b" & _
            " WHERE a.sBranchCd = " & strParm(oApp.BranchCode) & _
            " AND a.sBranchCd = b.sBranchCd" & _
            " AND a.sTranDate = " & strParm(Format(lsDate, "YYYYMMDD"))
   
   Set loData = New Recordset
   loData.Open lsSQL, oApp.Connection, , , adCmdText
   Debug.Print lsSQL
   If Not loData.EOF Then
      If loData("cPostedxx") >= 1 And DateDiff("d", loData("dUnenCode"), Format(lsDate, "YYYY-MM-DD")) >= 0 Then
         VerifyDTRDate = False
         Call isFieldEnable(False)
         MsgBox "Unable to Update DTR Summary!!!" & vbCrLf & _
                  "DTR Summary already " & Label2.Caption
    ElseIf DateDiff("d", loData("dUnenCode"), CDate(lsDate)) <= 0 Then
         VerifyDTRDate = False
         Call isFieldEnable(False)
      ElseIf loData("cPostedxx") = xeStateOpen Then
        VerifyDTRDate = True
        Call isFieldEnable(True)
      Else
         VerifyDTRDate = False
         Call isFieldEnable(False)
          MsgBox "Unable to encode transactions!!!" & vbCrLf & _
                  "Pls check your entry then try again!!" & "WARNING"
      End If
   Else
      VerifyDTRDate = False
      Call isFieldEnable(False)
      MsgBox "No Record found! Unable to encode transactions!!!" & vbCrLf & _
                  "Pls check your entry then try again!!" & "WARNING"
   End If
   
   'checking ulit sa detail
   pnDTREntryNo = 1
   
   lsSQL = "SELECT * FROM DTR_Summary_Detail" & _
            " WHERE sBranchCd = " & strParm(oApp.BranchCode) & _
            " AND sTranDate = " & strParm(Format(lsDate, "YYYYMMDD")) & _
            " ORDER BY nEntryNox DESC LIMIT 1"
   Debug.Print lsSQL
   Set loData = New Recordset
   loData.Open lsSQL, oApp.Connection, , , adCmdText
   
   If Not loData.EOF Then
      pnDTREntryNo = loData("nEntryNox") + 1
   Else
      pnDTREntryNo = 1
   End If
      
endProc:
   Exit Function
End Function

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   With GridEditor1
      Select Case Index
      Case 0
         If KeyCode = vbKeyReturn Or KeyCode = vbKeyF3 Then
            If VerifyDTRDate(txtField(0).Text) = True Then
               isFieldEnable (True)
            End If
         End If
      Case 3
         If KeyCode = vbKeyReturn Then
            If isAllDataValid = True And VerifyDTRDate(txtField(0).Text) = True Then
               If .Row <> .Rows And .Rows >= 2 Then
                  Call setGridValue
                  GridEditor1.Rows = GridEditor1.Rows + 1
                  pnCtr = GridEditor1.Rows - 1
                  Call ClearFields
               End If
            End If
         End If
      End Select
   End With
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = &H80000005
   End With
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
   Case 0
      If Not IsDate(txtField(Index).Text) Then
         txtField(Index).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
      Else
         txtField(Index).Text = Format(txtField(Index).Text, "MMMM DD, YYYY")
      End If
   Case 1
      If Not IsNumeric(txtField(Index).Text) Then
         txtField(Index).Text = ""
      Else
         txtField(Index).Text = txtField(Index).Text
      End If
   Case 2
      If Not IsNumeric(txtField(Index).Text) Then
         txtField(Index).Text = "0.00"
      Else
         txtField(Index).Text = Format(txtField(Index).Text, "#,##0.00")
      End If
   End Select
End Sub

Private Sub InitGrid()
   Dim lnCtr As Integer

   With GridEditor1
      .Rows = 2
      .Cols = 6
      .Font = "MS Sans Serif"

      'Column Title
      .TextMatrix(0, 1) = "Transaction Type"
      .TextMatrix(0, 2) = "Refer#"
      .TextMatrix(0, 3) = "Amount"
      .TextMatrix(0, 4) = "Remarks"
      .TextMatrix(0, 5) = "cPaym"
      .Row = 0

      'Column Alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = 3
      Next

      .ColAlignment(1) = 1

      'Column Width
      .ColWidth(0) = 330
      .ColWidth(1) = 2500
      .ColWidth(2) = 2000
      .ColWidth(3) = 2000
      .ColWidth(4) = 3000
      .ColWidth(5) = 700
      .ColWidth(6) = 700

      .ColEnabled(3) = False
      .ColEnabled(4) = False
      .ColEnabled(5) = False
      

      .Row = 1
      .Col = 1
      
      pnCtr = 1
   End With
End Sub

Private Sub setGridValue()
    With GridEditor1
      .TextMatrix(pnCtr, 1) = cmbTranType.Text
      .TextMatrix(pnCtr, 2) = txtField(1).Text
      .TextMatrix(pnCtr, 3) = Format(txtField(2).Text, "#,##0.00")
      .TextMatrix(pnCtr, 4) = txtField(3).Text
      .TextMatrix(pnCtr, 5) = IIf(optCashPaym(0).Value = Checked, 1, 0)
   End With
   Call ComputeTotal
End Sub

Function isAllDataValid() As Boolean

   If txtField(0).Text <> "" Then
      isAllDataValid = True
   Else
      MsgBox "Invalid Date Detected!!"
      isAllDataValid = False
      GoTo endProc
   End If
   
   If txtField(1).Text <> "" Then
      isAllDataValid = True
   Else
      MsgBox "Invalid Reference No Detected!!"
      isAllDataValid = False
      GoTo endProc
   End If
   
   If txtField(2).Text <> "" Then
      isAllDataValid = True
   Else
      MsgBox "Invalid Amount Detected!!"
      isAllDataValid = False
      GoTo endProc
   End If
   
   If txtField(3).Text <> "" Then
      isAllDataValid = True
   Else
      MsgBox "Invalid Remarks Detected!!"
      isAllDataValid = False
      GoTo endProc
   End If
   
endProc:
   Exit Function
End Function

Private Sub ClearFields()
   txtField(1).Text = ""
   txtField(2).Text = "0.00"
   txtField(3).Text = ""
End Sub

Private Sub ComputeTotal()
   Dim lnCtr As Integer
   Dim lnTranTotal As Currency

   lnTranTotal = 0#
   With GridEditor1
      For lnCtr = 1 To GridEditor1.Rows - 1
         lnTranTotal = lnTranTotal + IIf(.TextMatrix(lnCtr, 3) = "", 0, .TextMatrix(lnCtr, 3))
      Next
   End With
   txtField(4).Text = Format(lnTranTotal, "#,##0.00")
End Sub

Function SaveDTRTrans() As Boolean
   Dim lsSQL As String
   Dim lnCtr As Integer

   With GridEditor1
   
   oApp.BeginTrans
      For lnCtr = 1 To .Rows - 1
         Call CheckAllDetails(lnCtr, .TextMatrix(lnCtr, 1))
         If lbAllValueOkay = True Then
            lsSQL = "INSERT INTO DTR_Summary_Detail" & _
                     " SET sBranchCd = " & strParm(oApp.BranchCode) & _
                     ", sTranDate = " & strParm(Format(txtField(0).Text, "YYYYMMDD")) & _
                     ", nEntryNox = " & strParm(pnDTREntryNo) & _
                     ", sTranType = " & strParm(lsTranType) & _
                     ", sReferNox = " & strParm(.TextMatrix(lnCtr, 2)) & _
                     ", nTranAmtx = " & CDbl(.TextMatrix(lnCtr, 3)) & _
                     ", sRemarksx = " & strParm(.TextMatrix(lnCtr, 4)) & _
                     ", cCashPaym = " & strParm(.TextMatrix(lnCtr, 5)) & _
                     ", cHasEntry = '0' " & _
                     ", sModified = " & strParm(oApp.UserID) & _
                     ", dModified = " & dateParm(oApp.ServerDate)
                     Debug.Print lsSQL
            If oApp.Execute(lsSQL, "DTR_Summary_Detail") <= 0 Then
               MsgBox "Unable to Save DTR Summary Detail!!!"
               GoTo endProc
            End If
            
            pnDTREntryNo = pnDTREntryNo + 1
            
            SaveDTRTrans = True
         Else
            MsgBox "Unable to Save DTR Summary Detail!!!"
            SaveDTRTrans = False
            oApp.RollbackTrans
            GoTo endProc
         End If
      Next

   oApp.CommitTrans
   
   End With
   
endProc:
   Exit Function
End Function

Private Sub isFieldEnable(lbEnable As Boolean)
         
   txtField(0).Enabled = IIf(lbEnable = True, False, True)
   txtField(1).Enabled = lbEnable
   txtField(2).Enabled = lbEnable
   txtField(3).Enabled = lbEnable
   txtField(4).Enabled = lbEnable
   
   cmbTranType.Enabled = lbEnable
   
End Sub

Function CheckAllDetails(lnCtrRow As Integer, lsTrans As String)
   
   With GridEditor1
      Select Case lsTrans
      Case "CP Sales"
         lsTranType = "CPSl"
      Case "Misc"
         lsTranType = "MCSc"
      Case "Monthly Payment"
         lsTranType = "MPPy"
      Case "CP Load"
         lsTranType = "CPLd"
      End Select
   
      If .TextMatrix(lnCtrRow, 1) = "" Then
         MsgBox "Invalid Transaction Type Detected!!"
         lbAllValueOkay = False
         GoTo endProc
      Else
         lbAllValueOkay = True
      End If
      
      If .TextMatrix(lnCtrRow, 2) = "" Then
         MsgBox "Invalid Reference No Detected!!"
         lbAllValueOkay = False
         GoTo endProc
      Else
         lbAllValueOkay = True
      End If
      
      If .TextMatrix(lnCtrRow, 3) = "" Or .TextMatrix(lnCtrRow, 3) <= 0# Then
         MsgBox "Invalid Amount Detected!!"
         lbAllValueOkay = False
         GoTo endProc
      Else
         lbAllValueOkay = True
      End If
      
      If .TextMatrix(lnCtrRow, 4) = "" Then
         MsgBox "Invalid Remarks/Note Detected!!"
         lbAllValueOkay = False
         GoTo endProc
      Else
         lbAllValueOkay = True
      End If
      
   End With
endProc:
   Exit Function
End Function

Private Sub initButton(lbEnable As Boolean)
   cmdButton(0).Visible = lbEnable
   cmdButton(1).Visible = lbEnable
   cmdButton(2).Visible = lbEnable
   cmdButton(3).Visible = IIf(lbEnable = True, False, True)
   cmdButton(4).Visible = IIf(lbEnable = True, False, True)
   cmdButton(6).Visible = IIf(lbEnable = True, False, True)
End Sub

Private Function SearchTransaction() As Boolean
   Dim lsSQL As String
   Dim loData As Recordset
   Dim loDataDetail As Recordset
   Dim lasMaster() As String
   Dim lsMaster As String
   
   Dim lnCtr As Integer
   
   lsSQL = "SELECT" & _
               " b.sBranchNm" & _
               ", a.*" & _
            " FROM DTR_Summary a" & _
            ", Branch b" & _
            " WHERE a.sBranchCd = b.sBranchCd" & _
            " AND a.sBranchCd = " & strParm(oApp.BranchCode)
   
   Set loData = New Recordset
   loData.Open lsSQL, oApp.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
   If loData.EOF Then GoTo endProc
   
   If Not loData.EOF Then
      lsMaster = KwikBrowse(oApp, loData _
                           , "sBranchNm»sTranDate" _
                           , "Branch»Date")
      
      If lsMaster = Empty Then GoTo endProc
      
      lasMaster = Split(lsMaster, "»")
      
      txtField(0).Text = loData("sTranDate")
      Label2 = TransStat(loData("cPostedxx"))
      
      lsSQL = "SELECT" & _
               " c.sBranchNm" & _
               ", a.*" & _
            " FROM DTR_Summary_Detail a" & _
            ", DTR_Summary b" & _
            ", Branch c" & _
            " WHERE a.sBranchCd = b.sBranchCd" & _
            " AND a.sTranDate = b.sTranDate" & _
            " AND b.sBranchCd = c.sBranchCd" & _
            " AND a.sBranchCd = " & strParm(oApp.BranchCode) & _
            " AND a.sTranDate = " & strParm(loData("sTranDate"))
   
      Set loDataDetail = New Recordset
      loDataDetail.Open lsSQL, oApp.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
      
      If loDataDetail.EOF Then GoTo endProc
      End If
   
   With GridEditor1
      .Rows = loDataDetail.RecordCount + 1
      For lnCtr = 1 To loDataDetail.RecordCount
         Select Case loDataDetail("sTranType")
         Case "CPSl"
            .TextMatrix(lnCtr, 1) = "CP Sales"
         Case "MCSc"
            .TextMatrix(lnCtr, 1) = "Misc"
         Case "MPPy"
            .TextMatrix(lnCtr, 1) = "Monthly Payment"
         Case "CPLd"
            .TextMatrix(lnCtr, 1) = "CP Load"
         End Select
         .TextMatrix(lnCtr, 2) = loDataDetail("sReferNox")
         .TextMatrix(lnCtr, 3) = Format(loDataDetail("nTranAmtx"), "#,##0.00")
         .TextMatrix(lnCtr, 4) = loDataDetail("sRemarksx")
         .TextMatrix(lnCtr, 5) = loDataDetail("cCashPaym")
         loDataDetail.MoveNext
      Next
      pbLoad = True
   End With
  
endProc:
End Function

