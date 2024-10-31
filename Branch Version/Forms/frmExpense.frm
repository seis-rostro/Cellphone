VERSION 5.00
Object = "{0A7B56A6-35D0-4533-91DA-1715D3A0DD3E}#1.1#0"; "xrGridControl.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmExpense 
   BorderStyle     =   0  'None
   Caption         =   "Branch Expenses"
   ClientHeight    =   5295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7755
   DrawStyle       =   1  'Dash
   Icon            =   "frmExpense.frx":0000
   KeyPreview      =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5295
   ScaleWidth      =   7755
   ShowInTaskbar   =   0   'False
   Begin xrGridEditor.GridEditor GridEditor1 
      Height          =   3435
      Left            =   1650
      TabIndex        =   4
      Tag             =   "et0;eb0;et0;bc2"
      Top             =   1170
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
      MOUSEICON       =   "frmExpense.frx":000C
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
   Begin xrControl.xrFrame xrFrame1 
      Height          =   510
      Index           =   1
      Left            =   1575
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
         Index           =   1
         Left            =   3960
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   105
         Width           =   1980
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   0
         Left            =   1350
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   105
         Width           =   1860
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   285
         Index           =   0
         Left            =   3450
         TabIndex        =   2
         Top             =   120
         Width           =   585
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction No"
         Height          =   285
         Index           =   9
         Left            =   105
         TabIndex        =   0
         Top             =   120
         Width           =   1185
      End
   End
   Begin xrControl.xrFrame xrFrame3 
      Height          =   3585
      Left            =   1575
      Tag             =   "wt0;fb0"
      Top             =   1095
      Width           =   6075
      _ExtentX        =   10716
      _ExtentY        =   6324
      BackColor       =   12632256
      ClipControls    =   0   'False
      BorderStyle     =   1
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   480
      Left            =   1575
      Tag             =   "wt0;fb0"
      Top             =   4710
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
         Index           =   2
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
      Left            =   90
      TabIndex        =   8
      Top             =   2445
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
      Picture         =   "frmExpense.frx":0028
      CaptionAlign    =   0
      BackColor       =   14286077
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   0
      Left            =   90
      TabIndex        =   7
      Top             =   2025
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
      Picture         =   "frmExpense.frx":07A2
      CaptionAlign    =   0
      BackColor       =   14286077
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   3
      Left            =   90
      TabIndex        =   11
      Top             =   3285
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
      Picture         =   "frmExpense.frx":0F1C
      CaptionAlign    =   0
      BackColor       =   14286077
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   2
      Left            =   90
      TabIndex        =   10
      Top             =   2865
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
      Picture         =   "frmExpense.frx":1696
      CaptionAlign    =   0
      BackColor       =   14286077
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   4
      Left            =   90
      TabIndex        =   9
      Top             =   2445
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
      Picture         =   "frmExpense.frx":1E10
      CaptionAlign    =   0
      BackColor       =   14286077
      BackColorDown   =   8775418
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   5
      Left            =   90
      TabIndex        =   12
      Top             =   3285
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
      Picture         =   "frmExpense.frx":258A
      CaptionAlign    =   0
   End
End
Attribute VB_Name = "frmExpense"
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
Dim pbnewitem As Boolean
Dim pnindex As Integer

Private Sub cmdButton_Click(Index As Integer)
Dim lnCtr As Integer
Dim Total As Double

   Select Case Index
      Case 0 'Save
         oDriver.RecordSave
      Case 1 'Delete
         With GridEditor1
            If .Rows <> 2 Then
               .DeleteRow
            End If
            For lnCtr = 1 To .Rows - 1
               Total = CDbl(.TextMatrix(lnCtr, 2)) + Total
            Next
            txtField(2).Text = Format(Total, "#,##0.00")
         End With
      Case 2 'Browse
         oDriver.BrowseRecord
         InitButton xeModeReady
         ShowGrid
      Case 3 'close
         Unload Me
      Case 4 'New
         oDriver.RecordNew
         InitButton xeModeAddNew
         EmptyGrid
      Case 5 'Cancel
         oDriver.RecordCancelUpdate
         InitButton xeModeReady
         ShowGrid
      End Select
End Sub

Private Sub InitButton(lnStat As Integer)
   Dim lbShow As Boolean

   lbShow = IIf(lnStat = xeModeReady, False, True)
   cmdButton(0).Visible = lbShow
   cmdButton(1).Visible = lbShow
   cmdButton(5).Visible = lbShow

   With GridEditor1
      .ColEnabled(1) = lbShow
      .ColEnabled(2) = lbShow
   End With
   
   cmdButton(3).Visible = Not lbShow
   cmdButton(4).Visible = Not lbShow
   
End Sub

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0
   If bLoaded = False Then
      oDriver.RecordNew
      oDriver.ShowButton 1
      oDriver.DisableTextbox 0
      oDriver.DisableTextbox 2
      bLoaded = True
   End If
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
            txtField(2).Text = Format(Total, "#,##0.00")
         End If
      End If
   End With
End Sub

Private Sub oDriver_DisableOtherControl()
   oDriver.DisableTextbox 0
   oDriver.DisableTextbox 2
End Sub

Private Sub oDriver_EnableOtherControl()
   oDriver.DisableTextbox 0
   oDriver.DisableTextbox 2
End Sub

Private Sub oDriver_InitValue()
Dim lsSQL As String
Dim lsCondition As String

   oDriver.FieldReference(0) = True
   If Not oDriver.SetValue(0, getNextCode("CP_Expense_Master", "sTransNox", True, oApp.Connection, True, oApp.BranchCode)) Then Exit Sub
   pbnewitem = True
   
   txtField(2).Text = "0.00"
   oDriver.FieldValue(3) = oApp.BranchCode
   
End Sub

Private Sub Form_Load()
Dim lsSQL As String
Dim lnCtr As Integer


   CenterChildForm mdiMain, Me

   bLoaded = False
   
   Set oDriver = New FormDriver
   Set oDriver.AppDriver = oApp
   Set oDriver.MainForm = Me
       
   InitGrid
   InitButton xeModeAddNew
    
   Set oSkin = New FormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransaction
   
   oDriver.RecQuery = "SELECT" _
                           & " sTransNox, " _
                           & " dTrandate, " _
                           & " nTotalExp, " _
                           & " sBranchcd, " _
                           & " sModified, " _
                           & " dModified, " _
                           & " vTimeStmp  " _
                    & " FROM CP_Expense_Master " _

   oDriver.BrowseQuery = "SELECT" _
               & " sTransNox, " _
               & " dTranDate, " _
               & " nTotalExp  " _
            & " FROM CP_Expense_Master " _
            & " WHERE sBranchCd = '" & oApp.BranchCode & "' " _
            & " ORDER By dTrandate Desc "
   
   oDriver.InitRecForm

   oDriver.BrowseColumn(0) = "sTransNox"
   oDriver.BrowseColumn(1) = "dTranDate"
   oDriver.BrowseColumn(2) = "nTotalExp"
   
   oDriver.BrowseFTitle(0) = "Tran. No."
   oDriver.BrowseFTitle(1) = "Date"
   oDriver.BrowseFTitle(2) = "Total Expense"
   
   oDriver.BrowseFFormat(1) = "MMMM dd, yyyy"
   oDriver.BrowseFFormat(2) = "#,##0.00"
    
    
    oDriver.FieldFormat(0) = "@@-@@@@@@@@"
    oDriver.FieldSize(0) = Len(oDriver.FieldFormat(0))
    
    oDriver.FieldFormat(1) = "MMMM DD, YYYY"
    oDriver.FieldStart = 1
    oDriver.FieldFormat(2) = "#,##0.00"

End Sub

Private Sub InitGrid()
    
    With GridEditor1
        .Rows = 2
        .Cols = 3
        .Font = "MS Sans Serif"
        
        'column title
        .TextMatrix(0, 1) = "Particulars"
        .TextMatrix(0, 2) = "Amount"
                
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

Private Sub EmptyGrid()

With GridEditor1
   .Rows = 2
   .TextMatrix(1, 1) = ""
   .TextMatrix(1, 2) = "0.00"
End With

End Sub
Private Sub ShowGrid()
Dim lsSQL As String
Dim showdetail As New ADODB.Recordset
Dim lnCtr As Integer

   lsSQL = "SELECT " _
               & " sTransNox, " _
               & " nEntryNox, " _
               & " sDescript, " _
               & " nAmountxx, " _
               & " sModified, " _
               & " dModified  " _
         & " FROM CP_Expense_Detail " _
         & " WHERE sTransNox = '" & oDriver.FieldValue(0) & "'" _
         & " ORDER BY nEntryNox "
   
   Set showdetail = New ADODB.Recordset
   showdetail.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   
   With GridEditor1
      .Rows = showdetail.RecordCount + 1
      For lnCtr = 0 To showdetail.RecordCount - 1
         .TextMatrix(lnCtr + 1, 0) = showdetail("nEntryNox")
         .TextMatrix(lnCtr + 1, 1) = showdetail("sDescript")
         .TextMatrix(lnCtr + 1, 2) = Format(showdetail("nAmountxx"), "#,##0.00")
         showdetail.MoveNext
      Next
   End With
   
   Set showdetail = Nothing

End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set oDriver = Nothing
   Set oSkin = Nothing
End Sub

Private Sub GridEditor1_GotFocus()
   GridEditor1.Col = 1
End Sub

Private Sub oDriver_SaveComplete()
   InitButton xeModeReady
   EmptyGrid
End Sub

Private Sub oDriver_WillSave(Cancel As Boolean)
   If oDriver.FieldValue(1) = "" Then
      MsgBox "Invalid Date detected!!!", vbCritical, "Warning"
      txtField(1).SetFocus
      Cancel = True
   Else
      If pbnewitem Then
         Cancel = Not SaveCP_Expense_Detail
            If Cancel Then Exit Sub
      End If
      oDriver.FieldValue(2) = CDbl(txtField(2).Text)
   End If
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   oDriver.ColumnIndex = Index
   txtfieldGotfocus = True
   pnindex = Index
   txtField(Index).BackColor = &HE1FEFF
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   txtField(Index).BackColor = &H80000005
End Sub

Private Sub txtfield_Validate(Index As Integer, Cancel As Boolean)
   Cancel = Not oDriver.ValidateField(Index)
   If Index = 1 Then
      If Not IsDate(txtField(Index).Text) Then
         txtField(Index).Text = Format(oApp.ServerDate, "MMMM dd, yyyy")
      Else
         txtField(Index).Text = Format(txtField(Index).Text, "MMMM dd, yyyy")
      End If
   End If
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

Private Function SaveCP_Expense_Detail() As Boolean
Dim lsSQL As String
Dim lnCtr As Integer
Dim lnrow As Long

SaveCP_Expense_Detail = True
On Error GoTo errProc

   With GridEditor1
         For lnCtr = 1 To .Rows - 1
            If .TextMatrix(lnCtr, 1) <> "" Then
               lsSQL = "INSERT INTO CP_Expense_Detail " _
                     & "( sTransNox, " _
                     & "  nEntryNox, " _
                     & "  sDescript, " _
                     & "  nAmountxx, " _
                     & "  sModified, " _
                     & "  dModified) " _
               & "VALUES " _
                     & "('" & oDriver.FieldValue(0) & "', " _
                     & "'" & .TextMatrix(lnCtr, 0) & "', " _
                     & "'" & .TextMatrix(lnCtr, 1) & "', " _
                     & "'" & CDbl(.TextMatrix(lnCtr, 2)) & "', " _
                     & "'" & Encrypt(oApp.UserID) & "', " _
                     & " getdate())"
               oApp.Connection.Execute lsSQL, lnrow, adCmdText
            End If
         Next

         If lnrow <= 0 Then
            MsgBox "Unable to Save Expense_Detail!!!" & vbCrLf & vbCrLf & _
            "Notify ROSALYN LAZO DESCALLAR for Assistance", vbCritical, "Warning"
            SaveCP_Expense_Detail = False
            GoTo endProc
         Else
            MsgBox "Record Successfully Saved!!!", vbInformation, Me.Caption
         End If

   End With

endProc:
   Exit Function
errProc:
   SaveCP_Expense_Detail = False
   MsgBox Err.Description, vbCritical, "Warning"
End Function


