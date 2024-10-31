VERSION 5.00
Object = "{0A7B56A6-35D0-4533-91DA-1715D3A0DD3E}#1.1#0"; "xrGridControl.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmLoadWallet_POS 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6645
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6885
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin xrControl.xrFrame xrFrame1 
      Height          =   660
      Index           =   1
      Left            =   75
      Tag             =   "wt0;fb0"
      Top             =   5490
      Width           =   6495
      _ExtentX        =   11456
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
         Index           =   2
         Left            =   4155
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   90
         Width           =   2220
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
         Left            =   2625
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
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   9525
      BackColor       =   7716603
      Begin xrGridEditor.GridEditor GridEditor1 
         Height          =   5175
         Left            =   90
         TabIndex        =   0
         Top             =   90
         Width           =   6285
         _ExtentX        =   11086
         _ExtentY        =   9128
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
         Object.HEIGHT          =   5175
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
         MOUSEICON       =   "frmLoadWallet_POS.frx":0000
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
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   1164
      BackColor       =   7716603
      Begin xrControl.xrButton cmdButton 
         Height          =   405
         Index           =   2
         Left            =   5130
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
         Picture         =   "frmLoadWallet_POS.frx":001C
         CaptionAlign    =   0
         BackColor       =   14286077
         BackColorDown   =   8775418
         BorderColorFocus=   8775418
      End
      Begin xrControl.xrButton cmdButton 
         Height          =   405
         Index           =   0
         Left            =   2610
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
         Picture         =   "frmLoadWallet_POS.frx":0796
         CaptionAlign    =   0
         BackColor       =   14286077
         BackColorDown   =   8775418
         BorderColorFocus=   8775418
      End
      Begin xrControl.xrButton cmdButton 
         Height          =   405
         Index           =   1
         Left            =   3870
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
         Picture         =   "frmLoadWallet_POS.frx":0F10
         CaptionAlign    =   0
         BackColor       =   14286077
         BackColorDown   =   8775418
         BorderColorFocus=   8775418
      End
   End
End
Attribute VB_Name = "frmLoadWallet_POS"
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
Dim lsSearch As String
Dim lsCancel As Integer
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
            Unload Me
            EmptyGrid
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

End Sub

Private Sub Form_Load()
Dim lsSQL As String

   CenterChildForm mdiMain, Me
   bLoaded = False
    
   Set oDriver = New FormDriver
   Set oDriver.AppDriver = oApp
   Set oDriver.MainForm = Me
   
   InitGrid
   
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
      End If
   End With
End Sub

Private Sub Grand_Total()
Dim Total As Double
Dim lnCtr As Integer
   Total = 0#
   With GridEditor1
      If Not IsNumeric(.TextMatrix(.Row, 5)) Then
         MsgBox "Invalid Amount!!!", vbCritical, "Warning"
         .TextMatrix(.Row, 5) = 0#
      Else
         For lnCtr = 1 To .Rows - 1
            Total = CDbl(.TextMatrix(lnCtr, 5)) + Total
         Next
         txtfield(2).Text = Format(Total, "#,##0.00")
      End If
   End With
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
      .TextMatrix(0, 4) = "%"
      .TextMatrix(0, 5) = "Sub Total"
      .Row = 0
      
      'column width
      .ColWidth(0) = 500
      .ColWidth(1) = 1500
      .ColWidth(2) = 1500
      .ColWidth(3) = 900
      .ColWidth(4) = 500
      .ColWidth(5) = 1200
              
      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 6
      .ColAlignment(4) = 6
      .ColAlignment(5) = 6
                        
      .ColNumberOnly(3) = True
      .ColNumberOnly(4) = True
      .ColNumberOnly(5) = True
      .ColDefault(3) = 0
      .ColDefault(4) = 0
      .ColDefault(5) = 0
      .ColEnabled(5) = False
      
      .ColFormat(3) = "#,#00.00"
      .ColFormat(5) = "#,#00.00"
      
      .Row = 1
    End With
End Sub


Private Sub ShowGrid()
Dim lnCtr

   With GridEditor1
      Select Case psForm
      Case ""
         frmCP_POS.MSFlexGrid1.Rows = .Rows + 1
         For lnCtr = 1 To .Rows - 1
            frmCP_POS.MSFlexGrid1.TextMatrix(lnCtr, 0) = .TextMatrix(lnCtr, 0)
            frmCP_POS.MSFlexGrid1.TextMatrix(lnCtr, 1) = frmCP_POS.txtothers(0).Text
            frmCP_POS.MSFlexGrid1.TextMatrix(lnCtr, 2) = frmCP_POS.txtothers(1).Text
            frmCP_POS.MSFlexGrid1.TextMatrix(lnCtr, 3) = Format(.TextMatrix(lnCtr, 3), "#,##0.00")
            frmCP_POS.MSFlexGrid1.TextMatrix(lnCtr, 4) = 1
            frmCP_POS.MSFlexGrid1.TextMatrix(lnCtr, 5) = .TextMatrix(lnCtr, 4)
            frmCP_POS.MSFlexGrid1.TextMatrix(lnCtr, 6) = 0
            frmCP_POS.MSFlexGrid1.TextMatrix(lnCtr, 7) = .TextMatrix(lnCtr, 2)
            frmCP_POS.MSFlexGrid1.TextMatrix(lnCtr, 8) = frmCP_POS.txtothers(0).Tag
            frmCP_POS.MSFlexGrid1.TextMatrix(lnCtr, 9) = Format(.TextMatrix(lnCtr, 5), "#,##0.00")
            frmCP_POS.MSFlexGrid1.TextMatrix(lnCtr, 10) = ""
            frmCP_POS.MSFlexGrid1.TextMatrix(lnCtr, 11) = .TextMatrix(lnCtr, 3)
            frmCP_POS.MSFlexGrid1.TextMatrix(lnCtr, 12) = frmCP_POS.txtothers(1).Tag
            frmCP_POS.MSFlexGrid1.TextMatrix(lnCtr, 13) = .TextMatrix(lnCtr, 1)
         Next
         With frmCP_POS
            .txtfield(2).Text = Format(txtfield(2).Text, "#,##0.00")
            .txtothers(7).Text = Format((txtfield(2).Text) / (1.12), "#,##0.00")
            .txtothers(9).Text = Format((txtfield(2).Text - .txtothers(7).Text), "#,##0.00")
         End With
      Case "Register"
         frmPOS_Register.MSFlexGrid1.Rows = .Rows
         For lnCtr = 1 To .Rows - 1
            frmPOS_Register.MSFlexGrid1.TextMatrix(lnCtr, 0) = .TextMatrix(lnCtr, 0)
            frmPOS_Register.MSFlexGrid1.TextMatrix(lnCtr, 1) = psBarrcode
            frmPOS_Register.MSFlexGrid1.TextMatrix(lnCtr, 2) = frmPOS_Register.MSFlexGrid1.TextMatrix(1, 2)
            frmPOS_Register.MSFlexGrid1.TextMatrix(lnCtr, 3) = Format(.TextMatrix(lnCtr, 3), "#,##0.00")
            frmPOS_Register.MSFlexGrid1.TextMatrix(lnCtr, 4) = 1
            frmPOS_Register.MSFlexGrid1.TextMatrix(lnCtr, 5) = .TextMatrix(lnCtr, 4)
            frmPOS_Register.MSFlexGrid1.TextMatrix(lnCtr, 6) = 0
            frmPOS_Register.MSFlexGrid1.TextMatrix(lnCtr, 7) = psStock
            frmPOS_Register.MSFlexGrid1.TextMatrix(lnCtr, 8) = .TextMatrix(lnCtr, 1)
            frmPOS_Register.MSFlexGrid1.TextMatrix(lnCtr, 9) = .TextMatrix(lnCtr, 2)
            frmPOS_Register.MSFlexGrid1.TextMatrix(lnCtr, 10) = Format(.TextMatrix(lnCtr, 5), "#,##0.00")
            frmPOS_Register.MSFlexGrid1.TextMatrix(lnCtr, 11) = 2
            frmPOS_Register.MSFlexGrid1.TextMatrix(lnCtr, 12) = .TextMatrix(lnCtr, 3)
         Next
         With frmPOS_Register
            .txtfield(4).Text = Format(txtfield(2).Text, "#,##0.00")
            .txtfield(5).Text = Format(txtfield(2).Text, "#,##0.00")
         End With
      End Select
   End With


End Sub

Private Sub GridEditor1_EditorValidate(Cancel As Boolean)
With GridEditor1
   Select Case .Col
      Case 3
         If Not IsNumeric(.TextMatrix(.Row, .Col)) Then
            .TextMatrix(.Row, .Col) = 0#
         Else
            If .TextMatrix(.Row, 4) <> 0 Then
               .TextMatrix(.Row, 5) = Format(.TextMatrix(.Row, 3) - (.TextMatrix(.Row, 3) * _
                                    (.TextMatrix(.Row, 4) / 100)), "#,#00.00")
            End If
         End If
         Grand_Total
      Case 4
         If Not IsNumeric(.TextMatrix(.Row, .Col)) Then
            .TextMatrix(.Row, .Col) = 0#
         Else
            .TextMatrix(.Row, 5) = Format(.TextMatrix(.Row, 3) - (.TextMatrix(.Row, 3) * _
                                    (.TextMatrix(.Row, 4) / 100)), "#,#00.00")
         End If
         Grand_Total
   End Select
End With
End Sub
