VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmInhouseFinCalculator 
   BorderStyle     =   0  'None
   Caption         =   "Inhouse Financing Calculator"
   ClientHeight    =   5310
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10320
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5310
   ScaleWidth      =   10320
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   8970
      TabIndex        =   28
      Top             =   1170
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "Cl&ose"
      AccessKey       =   "o"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmInhouseFinCalculator.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   8970
      TabIndex        =   27
      Top             =   540
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Compute"
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
      Picture         =   "frmInhouseFinCalculator.frx":077A
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   4635
      Index           =   0
      Left            =   120
      Tag             =   "wt0;fb0"
      Top             =   600
      Width           =   8580
      _ExtentX        =   15134
      _ExtentY        =   8176
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
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
         Index           =   14
         Left            =   1500
         TabIndex        =   3
         Text            =   "0.00"
         Top             =   690
         Width           =   1605
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
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
         Index           =   13
         Left            =   6810
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   2120
         Width           =   1140
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
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
         Index           =   12
         Left            =   4350
         TabIndex        =   31
         Top             =   2120
         Width           =   1140
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
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
         Index           =   11
         Left            =   4350
         TabIndex        =   13
         Top             =   1815
         Width           =   1140
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
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
         Left            =   4350
         TabIndex        =   11
         Top             =   1515
         Width           =   1140
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
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
         Index           =   9
         Left            =   4350
         TabIndex        =   9
         Top             =   1215
         Width           =   1140
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
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
         Index           =   8
         Left            =   6810
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   1815
         Width           =   1140
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
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
         Index           =   7
         Left            =   6810
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   1515
         Width           =   1140
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
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
         Index           =   6
         Left            =   6810
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   1215
         Width           =   1140
      End
      Begin VB.CheckBox Check1 
         Caption         =   "w/Insurance"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1500
         TabIndex        =   16
         Tag             =   "wt0;fb0"
         Top             =   1800
         Width           =   1590
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
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
         Index           =   5
         Left            =   1500
         TabIndex        =   5
         Text            =   "0.00"
         Top             =   990
         Width           =   1605
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
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
         Index           =   4
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   2100
         Width           =   1605
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
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
         Index           =   3
         Left            =   6810
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   915
         Width           =   1140
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
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
         Index           =   2
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1290
         Width           =   1605
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
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
         Index           =   1
         Left            =   4350
         TabIndex        =   7
         Top             =   915
         Width           =   1140
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
         Top             =   165
         Width           =   2490
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   1545
         Left            =   90
         TabIndex        =   30
         Tag             =   "et0;eb0;et0;bc2"
         Top             =   2880
         Width           =   8325
         _ExtentX        =   14684
         _ExtentY        =   2725
         _Version        =   393216
         Rows            =   5
         Enabled         =   0   'False
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Promo Discount"
         Height          =   195
         Index           =   15
         Left            =   225
         TabIndex        =   2
         Top             =   735
         Width           =   1125
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SC 24 Months"
         Height          =   195
         Index           =   14
         Left            =   5655
         TabIndex        =   34
         Top             =   2115
         Width           =   1005
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DP"
         Height          =   195
         Index           =   13
         Left            =   4080
         TabIndex        =   33
         Top             =   2115
         Width           =   225
      End
      Begin VB.Shape Shape3 
         Height          =   1935
         Left            =   3720
         Top             =   690
         Width           =   4620
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DP"
         Height          =   195
         Index           =   12
         Left            =   4080
         TabIndex        =   12
         Top             =   1845
         Width           =   225
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DP"
         Height          =   195
         Index           =   11
         Left            =   4080
         TabIndex        =   10
         Top             =   1545
         Width           =   225
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DP"
         Height          =   195
         Index           =   10
         Left            =   4080
         TabIndex        =   8
         Top             =   1245
         Width           =   225
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SC 12 Months"
         Height          =   195
         Index           =   9
         Left            =   5655
         TabIndex        =   25
         Top             =   1845
         Width           =   1005
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SC 9 Months"
         Height          =   195
         Index           =   8
         Left            =   5745
         TabIndex        =   23
         Top             =   1545
         Width           =   915
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SC 6 Months"
         Height          =   195
         Index           =   7
         Left            =   5745
         TabIndex        =   21
         Top             =   1245
         Width           =   915
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Other Items"
         Height          =   195
         Index           =   6
         Left            =   540
         TabIndex        =   4
         Top             =   1035
         Width           =   810
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Insurance Fee"
         Height          =   195
         Index           =   4
         Left            =   345
         TabIndex        =   17
         Top             =   2130
         Width           =   1020
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SC 3 Months"
         Height          =   195
         Index           =   3
         Left            =   5745
         TabIndex        =   19
         Top             =   945
         Width           =   915
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rebates"
         Height          =   195
         Index           =   2
         Left            =   750
         TabIndex        =   14
         Top             =   1335
         Width           =   600
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   8400
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DP"
         Height          =   195
         Index           =   1
         Left            =   4080
         TabIndex        =   6
         Top             =   945
         Width           =   225
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
         Height          =   195
         Index           =   5
         Left            =   1020
         TabIndex        =   29
         Top             =   5730
         Width           =   795
      End
      Begin VB.Shape Shape2 
         Height          =   465
         Index           =   0
         Left            =   120
         Top             =   5640
         Width           =   7350
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FFFFFF&
         Height          =   465
         Index           =   1
         Left            =   120
         Top             =   5655
         Width           =   7350
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1575
         Tag             =   "et0;ht2"
         Top             =   255
         Width           =   2490
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
         Index           =   0
         Left            =   285
         TabIndex        =   0
         Top             =   210
         Width           =   525
      End
   End
End
Attribute VB_Name = "frmInhouseFinCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmMCInsCalculator"

Private oPriceList As clsCPPriceList
Private oSkin As clsFormSkin

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String

   lsOldProc = "cmdButton_Click"
   '''On Error GoTo errProc
   Select Case Index
   Case 0
      ' process monthly payment
      Call prcMonthlyAmort
   Case 1
      Unload Me
   End Select

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
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

Private Sub Form_Load()
   Dim lsOldProc As String

   lsOldProc = "Form_Load"
   '''On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualRight

   Set oPriceList = New clsCPPriceList
   Set oPriceList.AppDriver = oApp
   oPriceList.DateTransact = oApp.ServerDate
   oPriceList.InitTransaction

   InitForm

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub InitForm()
   Dim lnCtr As Integer

   With MSFlexGrid1
      .Cols = 7
      .Rows = 5
      .Font = "MS Sans Serif"

      'Column Title
      .TextMatrix(0, 2) = "3 mos"
      .TextMatrix(0, 3) = "6 mos"
      .TextMatrix(0, 4) = "9 mos"
      .TextMatrix(0, 5) = "12 mos"
      .TextMatrix(0, 6) = "24 mos"

      .Row = 0
      'Column Alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = 3
      Next
      .Row = 1

      .TextMatrix(1, 1) = "Gross Monthly"
      .TextMatrix(2, 1) = "Net Monthly"
      .TextMatrix(3, 1) = "Gross Total"
      .TextMatrix(4, 1) = "Net Total"

      .Col = 1
      'Column Alignment
      For lnCtr = 0 To .Rows - 1
         .Row = lnCtr
         .CellFontBold = True
      Next

      'Column Width
      .ColWidth(0) = 325
      .ColWidth(1) = 1950
      .ColWidth(2) = 1180
      .ColWidth(3) = 1180
      .ColWidth(4) = 1180
      .ColWidth(5) = 1180
      .ColWidth(6) = 1180

      .ColAlignment(2) = flexAlignRightCenter
      .ColAlignment(3) = flexAlignRightCenter
      .ColAlignment(4) = flexAlignRightCenter
      .ColAlignment(5) = flexAlignRightCenter
      .ColAlignment(6) = flexAlignRightCenter
   End With
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      If Index = 1 Then .Text = Format(.Text, "#0")

      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = oApp.getColor("HT1")
   End With
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String

   lsOldProc = "txtField_KeyDown"
   '''On Error GoTo errProc

   If (KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn) And Index = 0 Then
      With txtField(Index)
         oPriceList.CPModel = .Text
         If oPriceList.CPModel <> "" Then
            Call LoadMaster
            SetNextFocus
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

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   With txtField(Index)
      Select Case Index
      Case 0
         oPriceList.CPModel = .Text
         If oPriceList.CPModel <> "" Then
            Call LoadMaster
         End If
      Case 1, 9, 10, 11, 12
         If IsNumeric(.Text) = False Then .Text = 0
         Select Case Index
         Case 1
            oPriceList.DownPayment(0) = CDbl(.Text)
            .Text = Format(oPriceList.DownPayment(0), "#,##0.00")
         Case 9
            oPriceList.DownPayment(1) = CDbl(.Text)
            .Text = Format(oPriceList.DownPayment(1), "#,##0.00")
         Case 10
            oPriceList.DownPayment(2) = CDbl(.Text)
            .Text = Format(oPriceList.DownPayment(2), "#,##0.00")
         Case 11
            oPriceList.DownPayment(3) = CDbl(.Text)
            .Text = Format(oPriceList.DownPayment(3), "#,##0.00")
         Case 12
            oPriceList.DownPayment(4) = CDbl(.Text)
            .Text = Format(oPriceList.DownPayment(4), "#,##0.00")
         End Select
         
         Call cmdButton_Click(0)
      Case 5
         If IsNumeric(.Text) = False Then .Text = 0#
         .Text = Format(.Text, "#,##0.00")
         oPriceList.OtherAmount = CDbl(.Text)

         oPriceList.DownPayment(0) = CDbl(IIf(txtField(1) = "", 0, txtField(1)))
         oPriceList.DownPayment(1) = CDbl(IIf(txtField(9) = "", 0, txtField(9)))
         oPriceList.DownPayment(2) = CDbl(IIf(txtField(10) = "", 0, txtField(10)))
         oPriceList.DownPayment(3) = CDbl(IIf(txtField(11) = "", 0, txtField(11)))
         oPriceList.DownPayment(4) = CDbl(IIf(txtField(12) = "", 0, txtField(12)))
         
         txtField(1) = Format(oPriceList.DownPayment(0), "#,##0.00")
         txtField(9) = Format(oPriceList.DownPayment(1), "#,##0.00")
         txtField(10) = Format(oPriceList.DownPayment(2), "#,##0.00")
         txtField(11) = Format(oPriceList.DownPayment(3), "#,##0.00")
         txtField(12) = Format(oPriceList.DownPayment(4), "#,##0.00")
         
         txtField(3) = Format(oPriceList.MiscCharge(0), "#,##0.00")
         txtField(6) = Format(oPriceList.MiscCharge(1), "#,##0.00")
         txtField(7) = Format(oPriceList.MiscCharge(2), "#,##0.00")
         txtField(8) = Format(oPriceList.MiscCharge(3), "#,##0.00")
         txtField(13) = Format(oPriceList.MiscCharge(4), "#,##0.00")
         Call cmdButton_Click(0)
      Case 14
         If IsNumeric(.Text) = False Then .Text = 0#
         .Text = Format(.Text, "#,##0.00")
         oPriceList.Discount = CDbl(.Text)
         
         oPriceList.DownPayment(0) = CDbl(IIf(txtField(1) = "", 0, txtField(1)))
         oPriceList.DownPayment(1) = CDbl(IIf(txtField(9) = "", 0, txtField(9)))
         oPriceList.DownPayment(2) = CDbl(IIf(txtField(10) = "", 0, txtField(10)))
         oPriceList.DownPayment(3) = CDbl(IIf(txtField(11) = "", 0, txtField(11)))
         oPriceList.DownPayment(4) = CDbl(IIf(txtField(12) = "", 0, txtField(12)))
         
         txtField(1) = Format(oPriceList.DownPayment(0), "#,##0.00")
         txtField(9) = Format(oPriceList.DownPayment(1), "#,##0.00")
         txtField(10) = Format(oPriceList.DownPayment(2), "#,##0.00")
         txtField(11) = Format(oPriceList.DownPayment(3), "#,##0.00")
         txtField(12) = Format(oPriceList.DownPayment(4), "#,##0.00")
         
         txtField(3) = Format(oPriceList.MiscCharge(0), "#,##0.00")
         txtField(6) = Format(oPriceList.MiscCharge(1), "#,##0.00")
         txtField(7) = Format(oPriceList.MiscCharge(2), "#,##0.00")
         txtField(8) = Format(oPriceList.MiscCharge(3), "#,##0.00")
         txtField(13) = Format(oPriceList.MiscCharge(4), "#,##0.00")
         Call cmdButton_Click(0)
      End Select
   End With
End Sub

Private Sub LoadMaster()
   txtField(0) = oPriceList.CPModel
   txtField(2) = Format(oPriceList.Rebate, "#,##0.00")
   txtField(4) = Format(oPriceList.EndMortgage, "#,##0.00")
   txtField(5) = "0.00"
   txtField(1) = Format(oPriceList.MinimumDown(0), "#,##0.00")
   txtField(9) = Format(oPriceList.MinimumDown(1), "#,##0.00")
   txtField(10) = Format(oPriceList.MinimumDown(2), "#,##0.00")
   txtField(11) = Format(oPriceList.MinimumDown(3), "#,##0.00")
   txtField(12) = Format(oPriceList.MinimumDown(4), "#,##0.00")
   txtField(14) = "0.00"
   
   txtField(3) = Format(oPriceList.MiscCharge(0), "#,##0.00")
   txtField(6) = Format(oPriceList.MiscCharge(1), "#,##0.00")
   txtField(7) = Format(oPriceList.MiscCharge(2), "#,##0.00")
   txtField(8) = Format(oPriceList.MiscCharge(3), "#,##0.00")
   txtField(13) = Format(oPriceList.MiscCharge(4), "#,##0.00")
End Sub

Private Sub prcMonthlyAmort()
   If oPriceList.CPModel = "" Then Exit Sub
   
   Dim lnCtr As Integer
   Dim lnDownPaym(4) As Double
   Dim lanTerm(4) As Integer
   Dim lnFinance As Currency
   Dim lnSelPrce As Currency
   Dim lnPrcFrom As Currency
   Dim lnPrcThru As Currency
   Dim lnBaseAmt As Currency

   lanTerm(0) = 3
   lanTerm(1) = 6
   lanTerm(2) = 9
   lanTerm(3) = 12
   lanTerm(4) = 24

   lnDownPaym(0) = CDbl(txtField(1).Text)
   lnDownPaym(1) = CDbl(txtField(9).Text)
   lnDownPaym(2) = CDbl(txtField(10).Text)
   lnDownPaym(3) = CDbl(txtField(11).Text)
   lnDownPaym(4) = CDbl(txtField(12).Text)

   With MSFlexGrid1
      For lnCtr = 0 To UBound(lanTerm)
'         If lnDownPaym(lnCtr) <= 0 Then Exit Sub
            
         .TextMatrix(1, lnCtr + 2) = "N/A"
         .TextMatrix(2, lnCtr + 2) = "N/A"
         .TextMatrix(3, lnCtr + 2) = "N/A"
         .TextMatrix(4, lnCtr + 2) = "N/A"
         
         lnFinance = oPriceList.getMonthly(lnDownPaym(lnCtr), lanTerm(lnCtr), lnSelPrce, lnPrcFrom, lnPrcThru)
         If lnPrcFrom = 0 Then
            .TextMatrix(1, lnCtr + 2) = Format(lnFinance, "#,##0")
            .TextMatrix(2, lnCtr + 2) = Format(CDbl(.TextMatrix(1, lnCtr + 2)) - CDbl(txtField(2).Text), "#,##0")
            .TextMatrix(3, lnCtr + 2) = Format(CDbl(.TextMatrix(2, lnCtr + 2)) * lanTerm(lnCtr), "#,##0")
            .TextMatrix(4, lnCtr + 2) = Format(CDbl(.TextMatrix(3, lnCtr + 2)) * lanTerm(lnCtr), "#,##0")
         Else
            Select Case lnCtr
            Case 0
               If lnSelPrce >= lnPrcFrom Then
                  .TextMatrix(1, lnCtr + 2) = Format(lnFinance, "#,##0")
                  .TextMatrix(2, lnCtr + 2) = Format(CDbl(.TextMatrix(1, lnCtr + 2)) - CDbl(txtField(2).Text), "#,##0")
                  .TextMatrix(3, lnCtr + 2) = Format(CDbl(.TextMatrix(2, lnCtr + 2)) * lanTerm(lnCtr), "#,##0")
                  .TextMatrix(4, lnCtr + 2) = Format(CDbl(.TextMatrix(3, lnCtr + 2)) * lanTerm(lnCtr), "#,##0")
                  lnBaseAmt = lnPrcFrom
               End If
            Case 1, 2, 3, 4
               If lnSelPrce >= lnBaseAmt Then
                  .TextMatrix(1, lnCtr + 2) = Format(lnFinance, "#,##0")
                  .TextMatrix(2, lnCtr + 2) = Format(CDbl(.TextMatrix(1, lnCtr + 2)) - CDbl(txtField(2).Text), "#,##0")
                  .TextMatrix(3, lnCtr + 2) = Format(CDbl(.TextMatrix(2, lnCtr + 2)) * lanTerm(lnCtr), "#,##0")
                  .TextMatrix(4, lnCtr + 2) = Format(CDbl(.TextMatrix(3, lnCtr + 2)) * lanTerm(lnCtr), "#,##0")
                  
               End If
            End Select
         End If
         
         Select Case lnCtr
         Case 0
            txtField(3) = Format(IIf(lnFinance > 0, oPriceList.MiscCharge(0), 0), "#,##0.00")
            txtField(1) = Format(IIf(lnFinance > 0, oPriceList.DownPayment(0), 0), "#,##0.00")
         Case 1
            txtField(6) = Format(IIf(lnFinance > 0, oPriceList.MiscCharge(1), 0), "#,##0.00")
            txtField(9) = Format(IIf(lnFinance > 0, oPriceList.DownPayment(1), 0), "#,##0.00")
         Case 2
            txtField(7) = Format(IIf(lnFinance > 0, oPriceList.MiscCharge(2), 0), "#,##0.00")
            txtField(10) = Format(IIf(lnFinance > 0, oPriceList.DownPayment(2), 0), "#,##0.00")
         Case 3
            txtField(8) = Format(IIf(lnFinance > 0, oPriceList.MiscCharge(3), 0), "#,##0.00")
            txtField(11) = Format(IIf(lnFinance > 0, oPriceList.DownPayment(3), 0), "#,##0.00")
        Case 4
            txtField(13) = Format(IIf(lnFinance > 0, oPriceList.MiscCharge(4), 0), "#,##0.00")
            txtField(12) = Format(IIf(lnFinance > 0, oPriceList.DownPayment(4), 0), "#,##0.00")
         End Select
      Next
   End With
End Sub
