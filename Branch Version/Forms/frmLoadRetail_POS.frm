VERSION 5.00
Object = "{0A7B56A6-35D0-4533-91DA-1715D3A0DD3E}#1.1#0"; "xrGridControl.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmLoadRetail_POS 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Load Retail POS"
   ClientHeight    =   6885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   7620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin xrControl.xrFrame xrFrame1 
      Height          =   660
      Index           =   1
      Left            =   75
      Tag             =   "wt0;fb0"
      Top             =   5490
      Width           =   7470
      _ExtentX        =   13176
      _ExtentY        =   1164
      BackColor       =   7716603
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   435
         Index           =   0
         Left            =   5145
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   90
         Width           =   2220
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   435
         Index           =   2
         Left            =   1965
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   90
         Width           =   2220
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Deduct"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   4320
         TabIndex        =   7
         Top             =   180
         Width           =   780
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   510
         TabIndex        =   1
         Top             =   180
         Width           =   1470
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   5400
      Index           =   0
      Left            =   75
      Tag             =   "wt0;fb0"
      Top             =   60
      Width           =   7470
      _ExtentX        =   13176
      _ExtentY        =   9525
      BackColor       =   7716603
      Begin xrGridEditor.GridEditor GridEditor1 
         Height          =   5235
         Left            =   60
         TabIndex        =   0
         Top             =   60
         Width           =   7320
         _ExtentX        =   12912
         _ExtentY        =   9234
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
         Object.HEIGHT          =   5235
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
         MOUSEICON       =   "frmLoadRetail_POS.frx":0000
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
   Begin xrControl.xrFrame xrFrame1 
      Height          =   660
      Index           =   2
      Left            =   75
      Tag             =   "wt0;fb0"
      Top             =   6150
      Width           =   7470
      _ExtentX        =   13176
      _ExtentY        =   1164
      BackColor       =   7716603
      Begin xrControl.xrButton cmdButton 
         Height          =   405
         Index           =   2
         Left            =   6105
         TabIndex        =   5
         Top             =   105
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
         Picture         =   "frmLoadRetail_POS.frx":001C
         CaptionAlign    =   0
         BackColor       =   14286077
         BackColorDown   =   8775418
         BorderColorFocus=   8775418
      End
      Begin xrControl.xrButton cmdButton 
         Height          =   405
         Index           =   0
         Left            =   3585
         TabIndex        =   3
         Top             =   105
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   714
         Caption         =   "&OK"
         AccessKey       =   "O"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmLoadRetail_POS.frx":0796
         CaptionAlign    =   0
         BackColor       =   14286077
         BackColorDown   =   8775418
         BorderColorFocus=   8775418
      End
      Begin xrControl.xrButton cmdButton 
         Height          =   405
         Index           =   1
         Left            =   4845
         TabIndex        =   4
         Top             =   105
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   714
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
         Picture         =   "frmLoadRetail_POS.frx":0F10
         CaptionAlign    =   0
         BackColor       =   14286077
         BackColorDown   =   8775418
         BorderColorFocus=   8775418
      End
   End
End
Attribute VB_Name = "frmLoadRetail_POS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents oDriver As FormDriver
Attribute oDriver.VB_VarHelpID = -1
Private bLoaded As Boolean
Private oRS As New ADODB.Recordset

Dim txtfieldGotfocus As Boolean
Dim pbnewitem As Boolean
Dim psSelected() As String
Dim lrsTarget As New ADODB.Recordset

Dim psForm As String
Dim psStock As String
Dim psBarrcode As String
Dim psBrand As String

Property Let oForm(oForm As String)
   psForm = oForm
End Property

Property Let oStock(oStock As String)
   psStock = oStock
End Property

Property Let oBarrcode(oBarrcode As String)
   psBarrcode = oBarrcode
End Property

Property Let oBrand(oBrand As String)
   psBrand = oBrand
End Property

Private Sub cmdButton_Click(Index As Integer)
Dim lsCancel As String
Dim lsSQL As String
Dim lnrow As Long

With GridEditor1
   Select Case Index
      Case 0   'OK
         If .TextMatrix(.Rows - 1, 1) = "" Then
            .Rows = .Rows - 1
         End If
         ShowGrid
         frmPOS_Register.Update_Load = True
         Me.Hide
      Case 1   'delete
         If .Rows <> 2 Then
            .DeleteRow
         End If
         Grand_Total
      Case 2   'cancel
         If .Rows > 1 Then
            lsCancel = MsgBox("Are you sure you want to Cancel this Transaction?" & vbCrLf & _
            "This Entry will be Erased!!!", vbQuestion + vbYesNo, "Confirm")
            If lsCancel <> vbYes Then Exit Sub
            frmPOS_Register.Update_Load = False
            EmptyGrid
            Unload Me
         End If
   End Select
End With
End Sub

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0
   
   If bLoaded = False Then
      bLoaded = True
   End If
   GridEditor1.Refresh
   GridEditor1.SetFocus

End Sub

Private Sub Form_Load()
Dim lsSQL As String

   CenterChildForm mdiMain, Me
   bLoaded = False
    
   Set oDriver = New FormDriver
   Set oDriver.AppDriver = oApp
   Set oDriver.MainForm = Me
   
   InitGrid
   GridEditor1.Refresh

End Sub

Private Sub EmptyGrid()
Dim pnCtr As Integer
   With frmCP_POS.MSFlexGrid1
      .Rows = 2
      For pnCtr = 1 To .Cols - 1
         .TextMatrix(1, pnCtr) = ""
      Next
   End With
End Sub

Private Sub GridEditor1_AddingRow(Cancel As Boolean)
   With GridEditor1
      If .TextMatrix(.Row, 1) = "" Then
         Cancel = True
      ElseIf .TextMatrix(.Row, 2) = "" Then
         Cancel = True
      ElseIf .TextMatrix(.Row, 3) = 0 Then
         Cancel = True
      ElseIf .TextMatrix(.Row, 4) = "" Then
         Cancel = True
      ElseIf .TextMatrix(.Row, 5) <= 0 Then
         Cancel = True
      End If
   End With
End Sub

Private Sub Grand_Total()
Dim Total As Double
Dim Deduc As Double
Dim lnCtr As Integer
   Total = 0#
   Deduc = 0#
   With GridEditor1
      If Not IsNumeric(.TextMatrix(.Row, 3)) Then
         MsgBox "Invalid Amount!!!", vbCritical, "Warning"
         .TextMatrix(.Row, 3) = 0#
         .TextMatrix(.Row, 4) = 0#
         .TextMatrix(.Row, 5) = 0#
      Else
         For lnCtr = 1 To .Rows - 1
            Total = CDbl(.TextMatrix(lnCtr, 3)) + Total
            Deduc = CDbl(.TextMatrix(lnCtr, 4)) + Deduc
         Next
         txtField(2).Text = Format(Total, "#,##0.00")
         txtField(0).Text = Format(Deduc, "#,##0.00")
      End If
   End With
End Sub

Private Sub Show_Deduction()
Dim Deduc As Double
Dim lnCtr As Integer
   
   Deduc = 0#
   With GridEditor1
      If Not IsNumeric(.TextMatrix(.Row, 3)) Then
         .TextMatrix(.Row, 3) = 0#
         .TextMatrix(.Row, 4) = 0#
         .TextMatrix(.Row, 5) = 0#
      Else
         For lnCtr = 1 To .Rows - 1
            If lnCtr = 1 Then
               .TextMatrix(lnCtr, 5) = Format(txtField(0).Tag - .TextMatrix(lnCtr, 4), "#,##0.00")
            Else
               .TextMatrix(lnCtr, 5) = Format(.TextMatrix(lnCtr - 1, 5) - .TextMatrix(lnCtr, 4), "#,##0.00")
            End If
         Next
      End If
   End With
End Sub

Private Sub GridEditor1_KeyDown(KeyCode As Integer, Shift As Integer)
   With GridEditor1
      If KeyCode = vbKeyF3 Then
         If .Col = 3 Then
            Search_Matrix True
         End If
      End If
   End With
End Sub

Private Sub Search_Matrix(ByVal SearchValue As Boolean)
Dim lsSearch As String
Dim lrs As ADODB.Recordset
Dim lsSQL As String
   
   Set lrs = New ADODB.Recordset
   
With GridEditor1
   
   lsSQL = "SELECT" _
            & " a.sStockIDx, " _
            & " b.sBarrcode, " _
            & " a.sMatrixNm, " _
            & " a.nAmountxx, " _
            & " a.nSelPrice  " _
         & " FROM ELoad_Matrix a " _
            & " LEFT JOIN CP_Inventory b " _
                & " ON a.sStockIDx = b.sStockIDx " _
         & " WHERE cRecdStat = 1 " _
         & " AND a.sStockIDx = '" & psStock & "' "
   
   If .TextMatrix(.Row, 3) <> 0 Then
      If SearchValue Then
         lsSQL = lsSQL & " AND a.nSelPrice = '" & CDbl(.TextMatrix(.Row, 3)) & "'"
      End If
   End If
   
   lsSQL = lsSQL & " ORDER BY a.nSelPrice"
            
   lrs.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   
   If lrs.RecordCount = 1 Then
      .TextMatrix(.Row, 3) = Format(lrs("nSelPrice"), "#,##0.00")
      .TextMatrix(.Row, 4) = Format(lrs("nAmountxx"), "#,##0.00")
      Show_Deduction
   ElseIf lrs.RecordCount > 1 Then
        lsSearch = KwikBrowse(oApp, lrs, _
                          "sBarrcode�" _
                        & "sMatrixNm�" _
                        & "nSelPrice", _
                          "Bar Code�" _
                        & "Description�" _
                        & "Load Amount")
                        
        If lsSearch <> "" Then
            psSelected = Split(lsSearch, "�")
            .TextMatrix(.Row, 3) = Format(psSelected(4), "#,##0.00")
            .TextMatrix(.Row, 4) = Format(psSelected(3), "#,##0.00")
        End If
        Show_Deduction
   ElseIf lrs.RecordCount = 0 Then
      .TextMatrix(.Row, 3) = 0
      .TextMatrix(.Row, 4) = ""
      .TextMatrix(.Row, 5) = 0
      MsgBox "Invalid Amount", vbCritical, "Warning!!!"
   End If
   End With
   
   Grand_Total
   Set lrs = Nothing
   
End Sub

Private Sub InitGrid()
    With GridEditor1
      .Rows = 2
      .Cols = 6
      .Font = "MS Sans Serif"
              
      'column title
      .TextMatrix(0, 1) = "Cellphone No."
      .TextMatrix(0, 2) = "Reference No."
      .TextMatrix(0, 3) = "Amount"
      .TextMatrix(0, 4) = "Load Matrix"
      .TextMatrix(0, 5) = "Balance"
      .Row = 0
      
      'column width
      .ColWidth(0) = 500
      .ColWidth(1) = 1500
      .ColWidth(2) = 1500
      .ColWidth(3) = 1100
      .ColWidth(4) = 1100
      .ColWidth(5) = 1500
              
      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 6
      .ColAlignment(4) = 6
      .ColAlignment(5) = 6
            
      .ColEnabled(4) = False
      .ColEnabled(5) = False
            
      .ColNumberOnly(3) = True
      .ColDefault(3) = 0
      
      .Row = 1
    End With
End Sub

Private Sub ShowGrid()
Dim lnCtr

   If psForm = "" Then
      With GridEditor1
         frmCP_POS.MSFlexGrid1.Rows = .Rows + 1
         For lnCtr = 1 To .Rows - 1
            frmCP_POS.MSFlexGrid1.TextMatrix(lnCtr, 0) = .TextMatrix(lnCtr, 0)
            frmCP_POS.MSFlexGrid1.TextMatrix(lnCtr, 1) = frmCP_POS.txtothers(0).Text
            frmCP_POS.MSFlexGrid1.TextMatrix(lnCtr, 2) = frmCP_POS.txtothers(1).Text
            frmCP_POS.MSFlexGrid1.TextMatrix(lnCtr, 3) = Format(.TextMatrix(lnCtr, 3), "#,##0.00")
            frmCP_POS.MSFlexGrid1.TextMatrix(lnCtr, 4) = 1
            frmCP_POS.MSFlexGrid1.TextMatrix(lnCtr, 5) = 0
            frmCP_POS.MSFlexGrid1.TextMatrix(lnCtr, 6) = 0
            frmCP_POS.MSFlexGrid1.TextMatrix(lnCtr, 7) = .TextMatrix(lnCtr, 2)
            frmCP_POS.MSFlexGrid1.TextMatrix(lnCtr, 8) = frmCP_POS.txtothers(0).Tag
            frmCP_POS.MSFlexGrid1.TextMatrix(lnCtr, 9) = Format(.TextMatrix(lnCtr, 3), "#,##0.00")
            frmCP_POS.MSFlexGrid1.TextMatrix(lnCtr, 10) = ""
            frmCP_POS.MSFlexGrid1.TextMatrix(lnCtr, 11) = .TextMatrix(lnCtr, 4)
            frmCP_POS.MSFlexGrid1.TextMatrix(lnCtr, 12) = frmCP_POS.txtothers(1).Tag
            frmCP_POS.MSFlexGrid1.TextMatrix(lnCtr, 13) = .TextMatrix(lnCtr, 1)
         Next
      End With
   
      With frmCP_POS
         .txtField(2).Text = Format(txtField(2).Text, "#,##0.00")
         .txtothers(7).Text = Format((txtField(2).Text) / (1.12), "#,##0.00")
         .txtothers(9).Text = Format((txtField(2).Text - .txtothers(7).Text), "#,##0.00")
      End With
   ElseIf psForm = "Register" Then
      With GridEditor1
         frmPOS_Register.MSFlexGrid1.Rows = .Rows
         For lnCtr = 1 To .Rows - 1
            frmPOS_Register.MSFlexGrid1.TextMatrix(lnCtr, 0) = .TextMatrix(lnCtr, 0)
            frmPOS_Register.MSFlexGrid1.TextMatrix(lnCtr, 1) = psBarrcode
            frmPOS_Register.MSFlexGrid1.TextMatrix(lnCtr, 2) = psBrand
            frmPOS_Register.MSFlexGrid1.TextMatrix(lnCtr, 3) = Format(.TextMatrix(lnCtr, 3), "#,##0.00")
            frmPOS_Register.MSFlexGrid1.TextMatrix(lnCtr, 4) = 1
            frmPOS_Register.MSFlexGrid1.TextMatrix(lnCtr, 5) = 0
            frmPOS_Register.MSFlexGrid1.TextMatrix(lnCtr, 6) = 0
            frmPOS_Register.MSFlexGrid1.TextMatrix(lnCtr, 7) = psStock
            frmPOS_Register.MSFlexGrid1.TextMatrix(lnCtr, 8) = .TextMatrix(lnCtr, 1)
            frmPOS_Register.MSFlexGrid1.TextMatrix(lnCtr, 9) = .TextMatrix(lnCtr, 2)
            frmPOS_Register.MSFlexGrid1.TextMatrix(lnCtr, 10) = Format(.TextMatrix(lnCtr, 3), "#,##0.00")
            frmPOS_Register.MSFlexGrid1.TextMatrix(lnCtr, 11) = 3
            frmPOS_Register.MSFlexGrid1.TextMatrix(lnCtr, 12) = .TextMatrix(lnCtr, 4)
         Next
      End With
   
      With frmPOS_Register
         .txtField(5).Text = Format(txtField(2).Text, "#,##0.00")
         .txtField(4).Text = Format(txtField(2).Text, "#,##0.00")
      End With
   End If
End Sub

