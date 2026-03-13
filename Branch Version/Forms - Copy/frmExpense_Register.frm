VERSION 5.00
Object = "{0A7B56A6-35D0-4533-91DA-1715D3A0DD3E}#1.1#0"; "xrGridControl.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmExpense_Register 
   BorderStyle     =   0  'None
   Caption         =   "Expense Register"
   ClientHeight    =   5265
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5265
   ScaleWidth      =   7740
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   2
      Left            =   6390
      TabIndex        =   10
      Top             =   1620
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
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
      Picture         =   "frmExpense_Register.frx":0000
      CaptionAlign    =   0
      BackColor       =   14286077
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   510
      Index           =   1
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   6075
      _ExtentX        =   10716
      _ExtentY        =   900
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   0
         Left            =   4065
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   105
         Width           =   1845
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   2
         Left            =   675
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   105
         Width           =   1980
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction No"
         Height          =   285
         Index           =   9
         Left            =   2880
         TabIndex        =   2
         Top             =   120
         Width           =   1185
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   585
      End
   End
   Begin xrControl.xrFrame xrFrame3 
      Height          =   3585
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   1080
      Width           =   6075
      _ExtentX        =   10716
      _ExtentY        =   6324
      BackColor       =   12632256
      ClipControls    =   0   'False
      BorderStyle     =   1
      Begin xrGridEditor.GridEditor GridEditor1 
         Height          =   3435
         Left            =   60
         TabIndex        =   4
         Tag             =   "et0;eb0;et0;bc2"
         Top             =   60
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   6059
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
         Object.HEIGHT          =   3435
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
         MOUSEICON       =   "frmExpense_Register.frx":077A
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
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   480
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   4695
      Width           =   6075
      _ExtentX        =   10716
      _ExtentY        =   847
      BackColor       =   12632256
      ClipControls    =   0   'False
      BorderStyle     =   1
      Begin VB.TextBox txtfield 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   4
         Left            =   3660
         TabIndex        =   6
         TabStop         =   0   'False
         Tag             =   "ht0;ft0"
         Text            =   "Total Amount"
         Top             =   90
         Width           =   2310
      End
      Begin VB.Label lblField 
         BackStyle       =   0  'Transparent
         Caption         =   "Total "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   100
         Left            =   3030
         TabIndex        =   5
         Tag             =   "ht0"
         Top             =   105
         Width           =   585
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   1
      Left            =   6390
      TabIndex        =   8
      Top             =   1200
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
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
      Picture         =   "frmExpense_Register.frx":0796
      CaptionAlign    =   0
      BackColor       =   14286077
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   3
      Left            =   6390
      TabIndex        =   12
      Top             =   2040
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
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
      Picture         =   "frmExpense_Register.frx":0F10
      CaptionAlign    =   0
      BackColor       =   14286077
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   0
      Left            =   6390
      TabIndex        =   7
      Top             =   780
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
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
      Picture         =   "frmExpense_Register.frx":168A
      CaptionAlign    =   0
      BackColor       =   14286077
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   4
      Left            =   6390
      TabIndex        =   9
      Top             =   1200
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
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
      Picture         =   "frmExpense_Register.frx":1E04
      CaptionAlign    =   0
      BackColor       =   14286077
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   5
      Left            =   6390
      TabIndex        =   13
      Top             =   2040
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
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
      Picture         =   "frmExpense_Register.frx":257E
      CaptionAlign    =   0
      BackColor       =   14286077
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   6
      Left            =   6390
      TabIndex        =   11
      Top             =   1620
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
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
      Picture         =   "frmExpense_Register.frx":2CF8
      CaptionAlign    =   0
      BackColor       =   14286077
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
End
Attribute VB_Name = "frmExpense_Register"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents oDriver As FormDriver
Attribute oDriver.VB_VarHelpID = -1
Private oSkin As FormSkin
Private bLoaded As Boolean
Private oRS As New ADODB.Recordset

Dim txtfieldGotfocus As Boolean
Dim txtOthersGotfocus As Boolean
Dim pbnewitem As Boolean

Dim psSelected() As String

Dim pnindex As Integer
Dim Index As Integer
Dim pnCtr As Integer
Dim lnCtr As Integer

Private Sub cmdButton_Click(Index As Integer)
Dim Cancel As Boolean
Dim lnCtr As Integer
Dim Total As Double
   
   Select Case Index
      Case 0   'Save
        Cancel = Not UpdateCP_Expense_Detail
      Case 1 'Delete Row
         With GridEditor1
            If .Rows <> 2 Then
               .DeleteRow
            End If
            For lnCtr = 1 To .Rows - 1
               Total = CDbl(.TextMatrix(lnCtr, 2)) + Total
            Next
            txtfield(4).Text = Format(Total, "#,##0.00")
         End With
      Case 2  'Browse
         If pnindex = 2 Then Search_Expense
         If txtfield(0).Text <> "" Then InitButton xeModeReady
      Case 3 'Cancel
         InitButton xeModeReady
         ShowGrid
      Case 4   'New
         InitTxtField
         EmptyGrid
         InitButton xeModeAddNew
      Case 5
         Unload Me
      Case 6   'Update
         If txtfield(0).Text <> "" Then
            GridEditor1.SetFocus
            InitButton xeModeAddNew
         End If
   End Select

End Sub
Private Sub Form_Activate()
   oDriver.DisableTextbox 0
   oDriver.DisableTextbox 4
   txtfield(2).SetFocus
End Sub

Private Sub Form_Load()

CenterChildForm mdiMain, Me
bLoaded = False

Set oDriver = New FormDriver
Set oDriver.AppDriver = oApp
Set oDriver.MainForm = Me

InitTxtField
InitGrid
EmptyGrid

InitButton xeModeAddNew

Set oSkin = New FormSkin
Set oSkin.AppDriver = oApp
Set oSkin.Form = Me
oSkin.ApplySkin xeFormTransDetail
   
         
End Sub

Private Sub InitButton(lnStat As Integer)
   Dim lbShow As Boolean

   lbShow = IIf(lnStat = xeModeReady, False, True)
   cmdButton(0).Visible = lbShow
   cmdButton(1).Visible = lbShow
   cmdButton(2).Visible = lbShow
   cmdButton(3).Visible = lbShow
   
   With GridEditor1
      .ColEnabled(1) = lbShow
      .ColEnabled(2) = lbShow
   End With
   
   cmdButton(4).Visible = Not lbShow
   cmdButton(5).Visible = Not lbShow
   cmdButton(6).Visible = Not lbShow
   
End Sub

Private Sub InitGrid()

   With GridEditor1
      .Rows = 2
      .Cols = 3
      .Font = "MS Sans Serif"
      
      'column title
      .TextMatrix(0, 1) = "Particulars"
      .TextMatrix(0, 2) = "Amount"
      
      .Row = 0
      
      'column width
      .ColWidth(0) = 500
      .ColWidth(1) = 3850
      .ColWidth(2) = 1500
              
      .ColFormat(2) = "#,##0.00"
      
      .ColAlignment(1) = 1
      .ColAlignment(2) = 6
      
      .ColDefault(2) = "0.00"
      
      .Row = 1
      
   End With
   
End Sub
Private Sub InitTxtField()
Dim Index As Integer

For Index = 0 To 8
   Select Case Index
      Case 0
         txtfield(Index) = ""
      Case 2
         txtfield(Index) = Format(oApp.ServerDate, "MMMM dd, yyyy")
      Case 4
         txtfield(Index) = "0.00"
   End Select
Next
End Sub
Private Sub Search_Expense()
Dim lsSearch As String
Dim lsSQL As String
Dim lnCtr As Integer
Dim lrs As ADODB.Recordset
Dim Index As Integer

   Set lrs = New ADODB.Recordset
   lsSQL = "SELECT" _
               & " a.sTransNox, " _
               & " a.sModified, " _
               & " a.dTranDate, " _
               & " a.sBranchCd,  " _
               & " a.nTotalExp " _
            & " FROM CP_Expense_Master a " _
            & " WHERE dTranDate between '" & Format((txtfield(2).Text), "MMMM dd, yyyy") & "' " _
               & " AND '" & Format(txtfield(2).Text, "MMMM dd, yyyy") & " 23:59:59" & "'" _
               & " AND a.sBranchCd = '" & oApp.BranchCode & "'" _
            & " ORDER BY dtrandate "
   lrs.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   
   If lrs.RecordCount = 1 Then
      For Index = 0 To 4
         Select Case Index
            Case 0, 2
               txtfield(Index).Text = lrs(Index)
               If Index = 2 Then txtfield(Index).Text = Format(lrs(Index), "MMMM dd, yyyy")
            Case 4
               txtfield(Index).Text = Format(lrs(Index), "#,##0.00")
         End Select
      Next
      ShowGrid
   ElseIf lrs.RecordCount > 1 Then
      lsSearch = KwikBrowse(oApp, lrs, _
                        "dTranDate»nTotalExp", _
                        "Date»Total Expense", _
                        "MMMM dd, yyyy»#,##0.00")
      
      If lsSearch <> "" Then
          psSelected = Split(lsSearch, "»")
          ShowMoreRec
          ShowGrid
      End If
      
   Else
      MsgBox "No Record Found!!!", vbInformation, Me.Caption
   End If
   Set lrs = Nothing

End Sub

Private Sub ShowMoreRec()
Dim Index As Integer
   For Index = 0 To 4
      Select Case Index
         Case 0, 2
            txtfield(Index).Text = psSelected(Index)
            If Index = 2 Then txtfield(Index).Text = Format(psSelected(Index), "MMMM dd, yyyy")
         Case 4
            txtfield(Index).Text = Format(psSelected(Index), "#,##0.00")
      End Select
   Next
End Sub

Private Function UpdateCP_Expense_Detail() As Boolean
Dim lsSQL As String
Dim lnrow As Long
Dim lnCtr As Integer

UpdateCP_Expense_Detail = True
oApp.Connection.BeginTrans
On Error GoTo errProc

      lsSQL = "DELETE CP_Expense_Detail " _
               & "WHERE sTransNox = '" & txtfield(0).Text & "' "
      oApp.Connection.Execute lsSQL, lnrow, adCmdText
      If lnrow <> 0 Then oApp.RegisDelete lsSQL
   
      With GridEditor1
         For lnCtr = 1 To .Rows - 1
            lsSQL = "INSERT INTO CP_Expense_Detail " _
                     & "( sTransNox, " _
                     & "  nEntryNox, " _
                     & "  sDescript, " _
                     & "  nAmountxx, " _
                     & "  sModified, " _
                     & "  dModified) " _
               & "VALUES " _
                     & "('" & txtfield(0).Text & "', " _
                     & "'" & .TextMatrix(lnCtr, 0) & "', " _
                     & "'" & .TextMatrix(lnCtr, 1) & "', " _
                     & "'" & CDbl(.TextMatrix(lnCtr, 2)) & "', " _
                     & "'" & Encrypt(oApp.UserID) & "', " _
                     & " getdate())"
            oApp.Connection.Execute lsSQL, lnrow, adCmdText
         Next
         
         lsSQL = "UPDATE CP_Expense_Master SET" _
                     & " dTrandate = '" & txtfield(2).Text & "', " _
                     & " nTotalExp = '" & CDbl(txtfield(4).Text) & "', " _
                     & " sModified = '" & Encrypt(oApp.UserID) & "', " _
                     & " dModified = getdate() " _
               & " WHERE sTransNox = '" & txtfield(0).Text & "' "
         oApp.Connection.Execute lsSQL, lnrow, adCmdText
      
         If lnrow <= 0 Then
            MsgBox "Unable to Update Expense_Detail!!!" & vbCrLf & vbCrLf & _
            "Notify ROSALYN LAZO DESCALLAR for Assistance", vbCritical, "Warning"
            UpdateCP_Expense_Detail = False
            GoTo endProc
         Else
            MsgBox "Record Successfully Updated!!!", vbInformation, Me.Caption
            InitButton xeModeReady
         End If
         
      End With

endProc:
   oApp.Connection.CommitTrans
   Exit Function
errProc:
   oApp.Connection.RollbackTrans
   UpdateCP_Expense_Detail = False
   MsgBox Err.Description, vbCritical, "Warning"
End Function
Private Sub ShowGrid()
Dim lsSQL As String
Dim showdetail As New ADODB.Recordset
Dim lnCtr As Integer

   lsSQL = "SELECT " _
               & " a.sTransNox, " _
               & " a.nEntryNox, " _
               & " a.sDescript, " _
               & " a.nAmountxx, " _
               & " a.sModified, " _
               & " a.dModified  " _
         & " FROM CP_Expense_Detail a " _
            & " LEFT JOIN CP_Expense_Master b " _
               & " ON a.sTransNox = b.sTransNox  " _
         & " WHERE a.sTransNox = '" & txtfield(0).Text & "'" _
            & " AND b.sBranchCd = '" & oApp.BranchCode & "'" _
         & " ORDER BY nEntryNox "
   Set showdetail = New ADODB.Recordset
   showdetail.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   
   With GridEditor1
      If showdetail.RecordCount <> 0 Then
         .Rows = showdetail.RecordCount + 1
         For lnCtr = 0 To showdetail.RecordCount - 1
            .TextMatrix(lnCtr + 1, 0) = showdetail("nEntryNox")
            .TextMatrix(lnCtr + 1, 1) = showdetail("sDescript")
            .TextMatrix(lnCtr + 1, 2) = Format(showdetail("nAmountxx"), "#,##0.00")
            showdetail.MoveNext
            .ColEnabled(lnCtr) = False
         Next
      Else
         EmptyGrid
      End If
   End With
   
   Set showdetail = Nothing

End Sub
Private Sub EmptyGrid()
   With GridEditor1
      .Rows = 2
      .TextMatrix(1, 1) = ""
      .TextMatrix(1, 2) = "0.00"
   End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyReturn, vbKeyUp, vbKeyDown
      Select Case KeyCode
      Case vbKeyReturn, vbKeyDown
         If GetFocus = GridEditor1.hWnd Then Exit Sub
         SetNextFocus
      Case vbKeyUp
         SetPreviousFocus
      End Select
   End Select
End Sub

Private Sub GridEditor1_AddingRow(Cancel As Boolean)
   With GridEditor1
      If .TextMatrix(.Row, 1) = "" Then
         Cancel = True
      ElseIf Not IsNumeric(.TextMatrix(.Row, 2)) Or .TextMatrix(.Row, 2) = 0# Then
         Cancel = True
      End If
   End With
End Sub

Private Sub GridEditor1_EditorValidate(Cancel As Boolean)
Dim Total As Double
Dim lnCtr As Integer
   Total = 0#
   With GridEditor1
      If .Col = 2 Then
         If Not IsNumeric(.TextMatrix(.Row, 2)) Then
            MsgBox "Invalid Amount!!!", vbCritical, "Warning"
            .TextMatrix(.Row, 2) = 0#
         Else
            For lnCtr = 1 To .Rows - 1
               Total = CDbl(.TextMatrix(lnCtr, 2)) + Total
            Next
            txtfield(4).Text = Format(Total, "#,##0.00")
         End If
      End If
   End With
End Sub

Private Sub oDriver_DisableOtherControl()
   oDriver.DisableTextbox 0
   oDriver.DisableTextbox 4
End Sub

Private Sub oDriver_EnableOtherControl()
   oDriver.DisableTextbox 0
   oDriver.DisableTextbox 4
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   txtfieldGotfocus = True
   pnindex = Index
   oDriver.ColumnIndex = Index
   txtfield(Index).BackColor = &HE1FEFF
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Or KeyCode = 13 Then
      Select Case Index
      Case 0, 2
         Search_Expense
         If txtfield(Index).Text <> "" Then SetNextFocus
      End Select
      KeyCode = 0
   End If
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   If Index = 2 Then
      If Not IsDate(txtfield(Index).Text) Then
         txtfield(Index).Text = Format(oApp.ServerDate, "MMMM dd, yyyy")
      Else
         txtfield(Index).Text = Format(txtfield(Index).Text, "MMMM dd, yyyy")
      End If
   End If
   txtfield(Index).BackColor = &H80000005
End Sub

Private Sub txtfield_Validate(Index As Integer, Cancel As Boolean)
   If Not IsDate(txtfield(2).Text) Then
      txtfield(2).Text = Format(oApp.ServerDate, "MMMM dd, yyyy")
   Else
      txtfield(2).Text = Format(txtfield(2).Text, "MMMM dd, yyyy")
   End If
End Sub

