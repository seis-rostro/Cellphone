VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmCP_MCS_Override 
   BorderStyle     =   0  'None
   Caption         =   "MCS Override"
   ClientHeight    =   9585
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9615
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   143.274
   ScaleMode       =   0  'User
   ScaleWidth      =   82.603
   ShowInTaskbar   =   0   'False
   Tag             =   "wt0;fb0"
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   90
      TabIndex        =   16
      Top             =   1155
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
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
      Picture         =   "frmCP_MCS_Override.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   90
      TabIndex        =   17
      Top             =   2400
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Filter"
      AccessKey       =   "F"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCP_MCS_Override.frx":077A
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   5985
      Left            =   1560
      Tag             =   "wt0;fb0"
      Top             =   3495
      Width           =   7950
      _ExtentX        =   14023
      _ExtentY        =   10557
      BackColor       =   14737632
      ClipControls    =   0   'False
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   5775
         Left            =   75
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   90
         Width           =   7785
         _ExtentX        =   13732
         _ExtentY        =   10186
         _Version        =   393216
         FocusRect       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   600
      Index           =   3
      Left            =   90
      TabIndex        =   18
      Top             =   3030
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
      Picture         =   "frmCP_MCS_Override.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   90
      TabIndex        =   15
      Top             =   525
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
      Picture         =   "frmCP_MCS_Override.frx":166E
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   90
      TabIndex        =   20
      Top             =   1155
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
      Picture         =   "frmCP_MCS_Override.frx":1DE8
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   90
      TabIndex        =   19
      Top             =   525
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
      Picture         =   "frmCP_MCS_Override.frx":2562
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   2955
      Left            =   1560
      Tag             =   "wt0;fb0"
      Top             =   525
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   5212
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   6
         Left            =   6045
         TabIndex        =   13
         Top             =   1230
         Width           =   1755
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   600
         Index           =   3
         Left            =   1365
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1215
         Width           =   3675
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Left            =   1365
         TabIndex        =   1
         Top             =   135
         Width           =   1815
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   1365
         TabIndex        =   5
         Top             =   915
         Width           =   3675
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   915
         Index           =   4
         Left            =   1365
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   1830
         Width           =   3675
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   1365
         MaxLength       =   50
         TabIndex        =   3
         Top             =   615
         Width           =   1815
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   6045
         TabIndex        =   11
         Top             =   930
         Width           =   1755
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   195
         Index           =   11
         Left            =   645
         TabIndex        =   6
         Top             =   1260
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Thru"
         Height          =   195
         Index           =   14
         Left            =   5205
         TabIndex        =   12
         Top             =   1275
         Width           =   720
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   285
         Left            =   1455
         Tag             =   "et0;ht2"
         Top             =   255
         Width           =   1830
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trans. No."
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
         Index           =   9
         Left            =   150
         TabIndex        =   0
         Top             =   180
         Width           =   915
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   195
         Index           =   12
         Left            =   435
         TabIndex        =   8
         Top             =   1815
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date From"
         Height          =   195
         Index           =   2
         Left            =   5205
         TabIndex        =   10
         Top             =   975
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transact. Date"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   2
         Top             =   660
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company"
         Height          =   195
         Index           =   3
         Left            =   555
         TabIndex        =   4
         Top             =   960
         Width           =   660
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   6
      Left            =   90
      TabIndex        =   21
      Top             =   1785
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Del. Row"
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
      Picture         =   "frmCP_MCS_Override.frx":2CDC
   End
End
Attribute VB_Name = "frmCP_MCS_Override"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
'Private Const pxeMODULENAME = "frmCP_MCS_Override"
'Private WithEvents oTrans As clsCPMCSOverride
'Private oSkin As clsFormSkin
'Private oForm As frmMCSSerialCriteria
'
'Dim pbGridFocus As Boolean
'Dim pnIndex As Integer
'Dim pnCtr As Integer
'
'Dim pbSave As Boolean
'Dim pnRow As Integer
'Dim pbLoaded As Boolean
'Dim pbHasDet As Boolean
'
'Private Sub cmdButton_Click(Index As Integer)
'   Dim lnRep As Long
'
'   Select Case Index
'   Case 0 'save
'      If oTrans.SaveTransaction Then
'         MsgBox "Transaction saved successfuly.", vbInformation, pxeMODULENAME
'         lnRep = MsgBox("Do you want to print this transaction?", vbQuestion & vbYesNo, pxeMODULENAME)
'
'         If lnRep = vbYes Then
'            If oTrans.CloseTransaction(oTrans.Master("sTransNox")) Then MsgBox "Printing..."
'         End If
'
'         InitEntry
'      Else
'         MsgBox "Unable to save transaction.", vbCritical, pxeMODULENAME
'      End If
'   Case 1 'search
'      Select Case pnIndex
'      Case 2
'         Call txtField_KeyDown(pnIndex, vbKeyF3, 0)
'      End Select
'   Case 2 'filter serial
'      oForm.Show 1
'      If Not oForm.Cancelled Then
'         If oTrans.LoadSerial Then
'            Call loadSerialTemp
'         Else
'            Call clearGrid1
'         End If
'      End If
'   Case 3 'cancel
'      lnRep = MsgBox("Transaction is in Update Mode!!!" & vbCrLf & _
'                     "Do you want to Cancel Transaction!!!", vbYesNo + vbQuestion, "Confirm")
'      If lnRep = vbYes Then
'         InitEntry
'         initButton xeModeReady
'      End If
'   Case 4 'new
'      If oTrans.NewTransaction Then
'         Call initButton(xeModeAddNew)
'         Call InitEntry
'         txtField(1).SetFocus
'      End If
'   Case 5 'close
'      Unload Me
'   Case 6 'delete
'      With MSFlexGrid1
'         If .Rows > 2 Then
'            If oTrans.deleteDetail(.Row - 1) Then loadSerialTemp
'         End If
'      End With
'   End Select
'End Sub
'
'Private Sub Form_Activate()
'   oApp.MenuName = Me.Tag
'   Me.ZOrder 0
'
'   If Not pbLoaded Then
'      pbLoaded = True
'   End If
'End Sub
'
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'   Select Case KeyCode
'   Case vbKeyReturn, vbKeyUp, vbKeyDown
'      Select Case KeyCode
'      Case vbKeyReturn, vbKeyDown
'         Select Case GetFocus
'         Case MSFlexGrid1.hwnd
'            Exit Sub
'         End Select
'         SetNextFocus
'      Case vbKeyUp
'         SetPreviousFocus
'      Case vbKeySpace
'         If GetFocus = MSFlexGrid1.hwnd Then Call HighLightRowG1
'      End Select
'   End Select
'End Sub
'
'Private Sub Form_Load()
'   Dim lsOldProc As String
'
'   lsOldProc = "Form_Load"
'   ''On Error GoTo errProc
'
'   CenterChildForm mdiMain, Me
'
'   Set oSkin = New clsFormSkin
'   Set oSkin.AppDriver = oApp
'   Set oSkin.Form = Me
'   oSkin.ApplySkin xeFormTransEqualLeft
'
'   Set oTrans = New clsCPMCSOverride
'   Set oTrans.AppDriver = oApp
'   oTrans.Branch = oApp.BranchCode
'   oTrans.InitTransaction
'   oTrans.NewTransaction
'
'   Set oForm = New frmMCSSerialCriteria
'   Set oForm.Trans = oTrans
'
'   Call InitGrid
'   Call InitEntry
'   Call initButton(xeModeAddNew)
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " & " )", True
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'   Set oSkin = Nothing
'   Set oTrans = Nothing
'   Set oForm = Nothing
'
'   pbLoaded = False
'End Sub
'
'Private Sub initButton(lnStat As Integer)
'   Dim lbShow As Boolean
'
'   lbShow = IIf(lnStat = 0, False, True)
'   cmdButton(4).Visible = Not lbShow
'   cmdButton(5).Visible = Not lbShow
'
'   cmdButton(0).Visible = lbShow
'   cmdButton(1).Visible = lbShow
'   cmdButton(2).Visible = lbShow
'   cmdButton(3).Visible = lbShow
'
'   xrFrame1.Enabled = lbShow
'End Sub
'
'Private Sub InitEntry()
'   With oTrans
'      .NewTransaction
'
'      txtField(0) = Format(.Master("sTransNox"), "@@@@@@-@@@@@@")
'      txtField(1) = Format(.Master("dTransact"), "MMMM DD, YYYY")
'      txtField(2) = .Master("sCompnyNm")
'      txtField(3) = ""
'      txtField(4) = .Master("sRemarksx")
'      txtField(5) = Format(.Master("dPromoFrm"), "MMMM DD, YYYY")
'      txtField(6) = Format(.Master("dPromoTru"), "MMMM DD, YYYY")
'
'      pnRow = 0
'   End With
'
'   clearGrid1
'End Sub
'
'Private Sub InitGrid()
'   Dim lnCtr As Integer
'
'   With MSFlexGrid1
'      .Clear
'
'      .Cols = 4
'      .Rows = 2
'      .Font = "MS Sans Serif"
'      .RowHeight(0) = 350
'
'      .TextMatrix(0, 1) = "IMEI NO"
'      .TextMatrix(0, 2) = "BRAND"
'      .TextMatrix(0, 3) = "MODEL"
'      .Row = 0
'
'      'Column Alignment
'      For lnCtr = 0 To .Cols - 1
'         .Col = lnCtr
'         .CellFontBold = True
'         .CellAlignment = 4
'      Next
'
'      .ColWidth(0) = 450
'      .ColWidth(1) = 2500
'      .ColWidth(2) = 2400
'      .ColWidth(3) = 2400
'
'      .ColAlignment(1) = 1
'      .ColAlignment(2) = 1
'      .ColAlignment(3) = 1
'   End With
'End Sub
'
'Private Sub MSFlexGrid1_Click()
'   Call HighLightRowG1
'End Sub
'
'Private Sub oTrans_MasterRetrieved(ByVal Index As Integer)
'   txtField(Index) = IFNull(oTrans.Master(Index), "")
'End Sub
'
'Private Sub txtField_GotFocus(Index As Integer)
'   With txtField(Index)
'      Select Case Index
'      Case 1, 5, 6
'         .Text = Format(.Text, "MM/DD/YYYY")
'      End Select
'
'      .SelStart = 0
'      .SelLength = Len(.Text)
'      .BackColor = oApp.getColor("HT1")
'   End With
'
'   pnIndex = Index
'End Sub
'
'Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'   Dim lsOldProc As String
'
'   lsOldProc = "txtField_KeyDown"
'   ''On Error GoTo errProc
'
'   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
'      With txtField(Index)
'         If KeyCode = vbKeyF3 Then
'            oTrans.SearchMaster Index, .Text
'            If .Text <> "" Then SetNextFocus
'         Else
'            If .Text <> "" Then oTrans.SearchMaster Index, .Text
'         End If
'      End With
'      KeyCode = 0
'   End If
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " _
'                       & "  " & Index _
'                       & ", " & KeyCode _
'                       & ", " & Shift _
'                       & " )", True
'End Sub
'
'Private Sub txtField_LostFocus(Index As Integer)
'   With txtField(Index)
'      .BackColor = oApp.getColor("EB")
'   End With
'End Sub
'
'Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
'   With txtField(Index)
'      Select Case Index
'      Case 1, 5, 6
'         If Not IsDate(.Text) Then .Text = oApp.ServerDate
'         .Text = Format(.Text, "MMMM DD, YYYY")
'
'         oTrans.Master(Index) = CDate(.Text)
'      Case Else
'         oTrans.Master(Index) = .Text
'      End Select
'   End With
'End Sub
'
'Private Sub clearGrid1()
'   With MSFlexGrid1
'      .Rows = 2
'
'      .TextMatrix(1, 0) = "1"
'      .TextMatrix(1, 1) = ""
'      .TextMatrix(1, 2) = ""
'      .TextMatrix(1, 3) = ""
'
'      .Row = 1
'      .Col = 1
'      .ColSel = .Cols - 1
'   End With
'End Sub
'
'Private Sub loadSerialTemp()
'   Dim lnCtr As Integer
'
'   With MSFlexGrid1
'      .Rows = oTrans.ItemCount + 1
'      .Redraw = False
'      For pnCtr = 0 To oTrans.ItemCount - 1
'         .Row = pnCtr + 1
'         .TextMatrix(pnCtr + 1, 0) = pnCtr + 1
'         .TextMatrix(pnCtr + 1, 1) = oTrans.Detail(pnCtr, "sSerialNo")
'         .TextMatrix(pnCtr + 1, 2) = oTrans.Detail(pnCtr, "sBrandNme")
'         .TextMatrix(pnCtr + 1, 3) = oTrans.Detail(pnCtr, "sModelNme")
'
'         For lnCtr = 1 To .Cols - 1
'            .Col = lnCtr
'            If .CellFontBold Then
'               .CellFontBold = False
'               .CellBackColor = SystemColorConstants.vbWindowBackground
'            End If
'         Next
'      Next
'
'      If .Rows > 23 Then
'         .ColWidth(2) = 2300
'         .ColWidth(3) = 2300
'      Else
'         .ColWidth(2) = 2400
'         .ColWidth(3) = 2400
'      End If
'
'      .Row = 1
'      .Col = 1
'      .ColSel = .Cols - 1
'      .Redraw = True
'   End With
'End Sub
'
'Private Sub HighLightRowG1()
'   Dim lnCtr As Integer
'
'   If Not oTrans.SerialCount = 0 Then Exit Sub
'   With MSFlexGrid1
'      For lnCtr = 1 To .Cols - 1
'         .Col = lnCtr
'         If .CellFontBold Then
'            .CellFontBold = False
'            .CellBackColor = SystemColorConstants.vbWindowBackground
'            oTrans.Serial(.Row - 1, "cSelectxx") = xeNo
'         Else
'            .CellFontBold = True
'            .CellBackColor = &HC0C0C0
'            oTrans.Serial(.Row - 1, "cSelectxx") = xeYes
'         End If
'      Next
'
'      .Col = 1
'      .ColSel = .Cols - 1
'   End With
'End Sub
'
'Private Sub ShowError(ByVal lsProcName As String, Optional bEnd As Boolean = False)
'   With oApp
'      .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
'      If bEnd Then
'         .xShowError
'         End
'      Else
'         With Err
'            .Raise .Number, .Source, .Description
'         End With
'      End If
'   End With
'End Sub
