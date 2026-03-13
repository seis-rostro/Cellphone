VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmCPCredCardCalc 
   BorderStyle     =   0  'None
   Caption         =   "CP Credit Card Calculator"
   ClientHeight    =   4275
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10710
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4275
   ScaleWidth      =   10710
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   9390
      TabIndex        =   2
      Top             =   3570
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
      Picture         =   "frmCPCredCardCalc.frx":0000
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   3615
      Index           =   0
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   6376
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         CausesValidation=   0   'False
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
         Left            =   1380
         TabIndex        =   10
         Top             =   615
         Width           =   2490
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   5715
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "0.00"
         Top             =   615
         Width           =   1755
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   5715
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "0.00"
         Top             =   180
         Width           =   1755
      End
      Begin VB.TextBox txtField 
         Alignment       =   2  'Center
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
         Height          =   315
         Index           =   5
         Left            =   2070
         TabIndex        =   3
         Tag             =   "ht0"
         Top             =   5700
         Width           =   4515
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         CausesValidation=   0   'False
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
         Left            =   1380
         TabIndex        =   1
         Top             =   180
         Width           =   2490
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   2310
         Left            =   90
         TabIndex        =   5
         Tag             =   "et0;eb0;et0;bc2"
         Top             =   1185
         Width           =   8745
         _ExtentX        =   15425
         _ExtentY        =   4075
         _Version        =   393216
         Rows            =   9
         Cols            =   6
         Enabled         =   -1  'True
         MergeCells      =   1
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model Code"
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
         Index           =   3
         Left            =   150
         TabIndex        =   11
         Top             =   630
         Width           =   1020
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   1425
         Tag             =   "et0;ht2"
         Top             =   675
         Width           =   2490
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Discounted"
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
         Left            =   4635
         TabIndex        =   7
         Top             =   630
         Width           =   975
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SRP"
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
         Left            =   4710
         TabIndex        =   6
         Top             =   195
         Width           =   390
      End
      Begin VB.Line Line1 
         X1              =   75
         X2              =   7935
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
         Height          =   195
         Index           =   5
         Left            =   1020
         TabIndex        =   4
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
         Index           =   0
         Left            =   1425
         Tag             =   "et0;ht2"
         Top             =   240
         Width           =   2490
      End
      Begin VB.Label lblField 
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
         Index           =   0
         Left            =   150
         TabIndex        =   0
         Top             =   195
         Width           =   1065
      End
   End
End
Attribute VB_Name = "frmCPCredCardCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmCPCredCardCalc"

Private oPriceList As clsCPPriceList
Private oSkin As clsFormSkin

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String

   lsOldProc = "cmdButton_Click"
   ''On Error GoTo errProc
   Select Case Index
   Case 0
      ' process monthly payment
      Call loadZeroPercent
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
   ''On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualRight

   Set oPriceList = New clsCPPriceList
   Set oPriceList.AppDriver = oApp
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
      .Rows = 3
      .Font = "MS Sans Serif"

      'Column Title
      .TextMatrix(0, 0) = "Credit Card"
      .TextMatrix(0, 1) = "Payment"
      .TextMatrix(0, 2) = "3 Mos"
      .TextMatrix(0, 3) = "6 Mos"
      .TextMatrix(0, 4) = "12 Mos"
      .TextMatrix(0, 5) = "24 Mos"
      .TextMatrix(0, 6) = "36 Mos"
      
      .MergeCol(0) = True
      .MergeCol(1) = False
      .MergeCol(2) = False
      .MergeCol(3) = False
      .MergeCol(4) = False
      .MergeCol(5) = False
      .MergeCol(6) = False
      .WordWrap = True
      
      .Row = 0
      'Column Alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = 3
      Next
      .Row = 1
      
      'Column Width
      .ColWidth(0) = 1600
      .ColWidth(1) = 1000
      .ColWidth(2) = 1200
      .ColWidth(3) = 1200
      .ColWidth(4) = 1200
      .ColWidth(5) = 1200
      .ColWidth(6) = 1200

      .ColAlignment(2) = flexAlignRightCenter
      .ColAlignment(3) = flexAlignRightCenter
      .ColAlignment(4) = flexAlignRightCenter
      .ColAlignment(5) = flexAlignRightCenter
      .ColAlignment(6) = flexAlignRightCenter
   End With
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      If Index = 0 Then
         .SelStart = 0
         .SelLength = Len(.Text)
         .BackColor = oApp.getColor("HT1")
      Else
         .BackColor = oApp.getColor("HT1")
      End If
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

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   With txtField
      Select Case Index
      Case 0
         If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
            If txtField(0).Text = "" Then Exit Sub
            oPriceList.CPModel = txtField(0).Text
            
            If oPriceList.CPModel <> "" Then
               Call loadZeroPercent
            End If
            
            txtField(0).Text = oPriceList.CPModel
            txtField(3).Text = oPriceList.CPModelID
         End If
      Case 3
         If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
            If txtField(Index).Text = "" Then Exit Sub
            oPriceList.CPModelID = txtField(Index).Text
            
            If oPriceList.CPModelID <> "" Then
               Call loadZeroPercent
            End If
            
            txtField(Index).Text = oPriceList.CPModelID
            txtField(0).Text = oPriceList.CPModel
         End If
      End Select
   End With
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub loadZeroPercent()
   Dim lnCtr As Integer
   Dim lors As Recordset
   Dim lnRow As Long
   
   If oPriceList.CPModel = "" Then Exit Sub
   
   Set lors = oPriceList.loadZeroPercent()
   
   With MSFlexGrid1
      lnRow = 1
      .Rows = lors.RecordCount * 2 + 1
      
      Do Until lors.EOF()
         .TextMatrix(lnRow, 0) = lors("Bank Code")
         .TextMatrix(lnRow, 1) = "Monthly"
         .TextMatrix(lnRow, 2) = Format(lors("3 Mos"), "#,##0.00")
         .TextMatrix(lnRow, 3) = Format(lors("6 Mos"), "#,##0.00")
         .TextMatrix(lnRow, 4) = Format(lors("12 Mos"), "#,##0.00")
         .TextMatrix(lnRow, 5) = Format(lors("24 Mos"), "#,##0.00")
         .TextMatrix(lnRow, 6) = Format(lors("36 Mos"), "#,##0.00")
         
         .TextMatrix(lnRow + 1, 0) = lors("Bank Code")
         .TextMatrix(lnRow + 1, 1) = "Total"
         .TextMatrix(lnRow + 1, 2) = Format(lors("3 Gross"), "#,##0.00")
         .TextMatrix(lnRow + 1, 3) = Format(lors("6 Gross"), "#,##0.00")
         .TextMatrix(lnRow + 1, 4) = Format(lors("12 Gross"), "#,##0.00")
         .TextMatrix(lnRow + 1, 5) = Format(lors("24 Gross"), "#,##0.00")
         .TextMatrix(lnRow + 1, 6) = Format(lors("36 Gross"), "#,##0.00")
         
'         .MergeRow(lnRow) = True
'         .MergeRow(lnRow + 1) = True
         
         lnRow = lnRow + 2
         lors.MoveNext
      Loop
   
      For lnRow = 0 To .Rows - 1
         .Col = 0
         .Row = lnRow
         .CellFontBold = True
      
         .Col = 1
         .CellFontBold = True
      Next
   End With
   
   txtField(1) = Format(oPriceList.CashPrice(0, "nSelPrice"), "#,##0.00")
   txtField(2) = Format(oPriceList.CashPrice(0, "nLastPrce"), "#,##0.00")
End Sub
