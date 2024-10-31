VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0"; "xrGridControl.ocx"
Begin VB.Form frmConsignmentRegister 
   BorderStyle     =   0  'None
   Caption         =   "Consignment Register"
   ClientHeight    =   8895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13020
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8895
   ScaleWidth      =   13020
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   3135
      Index           =   1
      Left            =   120
      Tag             =   "wt0;fb0"
      Top             =   4950
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   5530
      BackColor       =   12632256
      Enabled         =   0   'False
      ClipControls    =   0   'False
      Begin xrGridEditor.GridEditor GridEditor1 
         Height          =   3000
         Left            =   45
         TabIndex        =   0
         Top             =   60
         Width           =   11160
         _ExtentX        =   19685
         _ExtentY        =   5292
         AllowBigSelection=   -1  'True
         AutoAdd         =   -1  'True
         AutoNumber      =   0   'False
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
         Object.HEIGHT          =   3000
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
         MOUSEICON       =   "frmConsignmentRegister.frx":0000
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
      Height          =   3870
      Index           =   0
      Left            =   135
      Tag             =   "wt0;fb0"
      Top             =   1050
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   6826
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   1410
         MaxLength       =   50
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   1395
         Width           =   5790
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   1425
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   165
         Width           =   1950
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   1425
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   705
         Width           =   2505
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   660
         Index           =   5
         Left            =   1380
         MaxLength       =   50
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   3075
         Width           =   5955
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   3
         Left            =   1425
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   1050
         Width           =   2505
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   4
         Left            =   4695
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   1050
         Width           =   2505
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   8
         Left            =   1410
         MaxLength       =   50
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   1740
         Width           =   2505
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   9
         Left            =   1410
         MaxLength       =   50
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   2085
         Width           =   2505
      End
      Begin VB.CheckBox chk 
         Caption         =   "VATable"
         Height          =   195
         Index           =   10
         Left            =   1410
         TabIndex        =   2
         Tag             =   "wt0;fb0"
         Top             =   2790
         Width           =   1530
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   11
         Left            =   1410
         MaxLength       =   50
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   2430
         Width           =   2505
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Name"
         Height          =   285
         Index           =   10
         Left            =   165
         TabIndex        =   20
         Top             =   1425
         Width           =   1200
      End
      Begin VB.Label lblField 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "unknown"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   375
         Index           =   12
         Left            =   7905
         TabIndex        =   19
         Tag             =   "eb0;et0"
         Top             =   360
         Width           =   2520
      End
      Begin VB.Shape Shape4 
         Height          =   600
         Index           =   0
         Left            =   7920
         Top             =   225
         Width           =   2490
      End
      Begin VB.Shape Shape3 
         Height          =   705
         Index           =   0
         Left            =   7875
         Top             =   180
         Width           =   2580
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transact. Date"
         Height          =   285
         Index           =   1
         Left            =   165
         TabIndex        =   18
         Top             =   735
         Width           =   1200
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   1
         Left            =   1515
         Tag             =   "et0;ht2"
         Top             =   270
         Width           =   1950
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction No"
         Height          =   285
         Index           =   0
         Left            =   165
         TabIndex        =   17
         Top             =   195
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   285
         Index           =   5
         Left            =   135
         TabIndex        =   16
         Top             =   3105
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date From"
         Height          =   285
         Index           =   2
         Left            =   150
         TabIndex        =   15
         Top             =   1080
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date Thru"
         Height          =   285
         Index           =   3
         Left            =   3930
         TabIndex        =   14
         Top             =   1095
         Width           =   795
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Discount"
         Height          =   285
         Index           =   6
         Left            =   165
         TabIndex        =   13
         Top             =   1770
         Width           =   990
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Add Discount"
         Height          =   285
         Index           =   7
         Left            =   165
         TabIndex        =   12
         Top             =   2130
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "With Held"
         Height          =   285
         Index           =   8
         Left            =   165
         TabIndex        =   11
         Top             =   2475
         Width           =   1200
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H8000000F&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   525
         Index           =   0
         Left            =   7980
         Tag             =   "et0;et0"
         Top             =   270
         Width           =   2400
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   705
      Index           =   2
      Left            =   135
      Tag             =   "wt0;fb0"
      Top             =   8100
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   1244
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   7
         Left            =   7680
         MaxLength       =   50
         TabIndex        =   22
         Tag             =   "ht0;ft0"
         Text            =   "0.00"
         Top             =   120
         Width           =   3435
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   6
         Left            =   2100
         MaxLength       =   50
         TabIndex        =   21
         Tag             =   "ht0;ft0"
         Top             =   105
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   9
         Left            =   6525
         TabIndex        =   24
         Top             =   180
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL QTY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   4
         Left            =   180
         TabIndex        =   23
         Top             =   165
         Width           =   1815
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   11655
      TabIndex        =   25
      Top             =   2400
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
      Picture         =   "frmConsignmentRegister.frx":001C
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   11655
      TabIndex        =   26
      Top             =   510
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
      Picture         =   "frmConsignmentRegister.frx":0796
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   11655
      TabIndex        =   27
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
      Picture         =   "frmConsignmentRegister.frx":0F10
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   540
      Index           =   3
      Left            =   120
      Tag             =   "wt0;fb0"
      Top             =   495
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   953
      BackColor       =   12632256
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
         Index           =   81
         Left            =   5355
         TabIndex        =   29
         Top             =   105
         Width           =   3810
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
         Index           =   80
         Left            =   1530
         TabIndex        =   28
         Top             =   90
         Width           =   1725
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Supplier Name"
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
         Index           =   20
         Left            =   3885
         TabIndex        =   31
         Top             =   135
         Width           =   1365
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Transaction No."
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
         Index           =   21
         Left            =   150
         TabIndex        =   30
         Top             =   120
         Width           =   1605
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   11655
      TabIndex        =   32
      Top             =   1140
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Print"
      AccessKey       =   "P"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmConsignmentRegister.frx":168A
   End
End
Attribute VB_Name = "frmConsignmentRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'jovanalic 2021-06-30 12:56pm
' started creating this ui
'she 2021-07-05 modified by she
Option Explicit
Private Const pxeMODULENAME = "frmConsignmentRegister"

Private WithEvents oTrans As ggcCPPurchasing.clsConsignmentTagging
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin

Dim pnIndex As Integer
Dim pbGridFocus As Boolean
Dim pnCtr As Integer
Dim pbSave As Boolean
Dim pbGridValidate As Boolean
Dim pbPosted As Boolean

Private Sub cmdButton_Click(Index As Integer)
   
   Dim lsOldProc As String
   Dim lnRep As Integer
   Dim lnMsg As String

   lsOldProc = "cmdButton_Click"
   ''On Error GoTo errProc

   txtField_LostFocus pnIndex
      Select Case Index
         Case 0 'Browse
            If oTrans.SearchTransaction() = True Then
               clearFields
               LoadMaster
               loadMasterDetail
            End If
         Case 1 'Print
            If oTrans.Master("cTranStat") = xeStateOpen Or oTrans.Master("cTranStat") = xeStateClosed Then
               lnRep = MsgBox("Do you want to PRINT Transfer?", vbYesNo)
               If lnRep = vbYes Then
                  PrintTrans
               Else
                  MsgBox "Unable to PRINT transaction.", vbCritical, "WARNING!"
               End If
            ElseIf oTrans.Master("cTranStat") = xeStatePosted Then
               MsgBox "Transaction was POsted!!!", vbCritical, "WARNING!"
            End If
         Case 2 ' Cancel
            If oTrans.CancelTransaction Then
            MsgBox "Transaction successfully cancelled!", vbInformation, "Information"
              clearFields
              setTransTat (-1)
            Else
               MsgBox "Unable to cancel transaction", vbCritical, "Error"
            End If
         Case 3 'Closed
            Unload Me
         
      End Select

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Sub clearFields()
   Dim loTxt As TextBox
   
   For Each loTxt In txtField
      loTxt = ""
   Next
     
   ClearDetail
   setTransTat (-1)
End Sub

Private Sub ClearDetail()
Dim lnCtr As Integer
   With GridEditor1
      .Rows = 2
      
      .TextMatrix(0, 0) = "No"
      .TextMatrix(0, 1) = "Branch"
      .TextMatrix(0, 2) = "S.I"
      .TextMatrix(0, 3) = "Barrcode"
      .TextMatrix(0, 4) = "Descript"
      .TextMatrix(0, 5) = "Qty"
      .TextMatrix(0, 6) = "Refer#"
      .TextMatrix(0, 7) = "U.Price"
      .TextMatrix(0, 8) = "Tag"
      .TextMatrix(0, 9) = "trans#"
      .TextMatrix(0, 10) = "sstockIDx"
      
      .TextMatrix(1, 0) = "1"
      .TextMatrix(1, 1) = ""
      .TextMatrix(1, 2) = ""
      .TextMatrix(1, 3) = ""
      .TextMatrix(1, 4) = ""
      .TextMatrix(1, 5) = ""
      .TextMatrix(1, 6) = ""
      .TextMatrix(1, 7) = ""
      .TextMatrix(1, 8) = ""
      .TextMatrix(1, 9) = ""
      .TextMatrix(1, 10) = ""
      
      For lnCtr = 1 To .Cols - 1
         .Col = lnCtr
         .CellBackColor = &HFFFFFF
         .CellFontBold = False
      Next
      
      .Row = 1
      .Col = 1
   End With
End Sub

Private Sub LoadMaster()
   With oTrans
      txtField(0).Text = IFNull(.Master("sTransNox"), "")
      txtField(1).Text = Format(IFNull(.Master("dTransact"), oApp.ServerDate), "MMMM DD, YYYY")
      txtField(2).Text = IFNull(.Master("sCompnyNm"), "")
      txtField(3).Text = Format(IFNull(.Master("dDateFrom"), oApp.ServerDate), "MMMM DD, YYYY")
      txtField(4).Text = Format(IFNull(.Master("dDateThru"), oApp.ServerDate), "MMMM DD, YYYY")
      txtField(5).Text = IFNull(.Master("sRemarksx"), "")
      txtField(7).Text = Format(IFNull(.Master("nTranTotl"), 0), "#,##0.00")
      txtField(8).Text = Format(IFNull(.Master("nDiscount"), 0), "#,##0.00")
      txtField(9).Text = Format(IFNull(.Master("nAddDiscx"), 0), "#,##0.00")
      txtField(11).Text = Format(IFNull(.Master("nTWithHld"), 0), "#,##0.00")
      If .Master("cVATaxabl") = 1 Then
         chk(10).Value = vbChecked
      Else
         chk(10).Value = vbUnchecked
      End If
      setTransTat (.Master("cTranStat"))
       txtField(80).Text = IFNull(.Master("sTransNox"), "")
       txtField(81).Text = IFNull(.Master("sCompnyNm"), "")
   End With
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
      
   End Select
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String

   lsOldProc = "Form_Load"
   '''On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oTrans = New ggcCPPurchasing.clsConsignmentTagging
   Set oTrans.AppDriver = oApp
   
    oTrans.TransStatus = "10234"
   oTrans.InitTransaction
   oTrans.NewTransaction

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualRight

   InitGrid
   InitForm
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
End Sub

Private Sub InitGrid()
   Dim lnCtr As Integer

   With GridEditor1
      .Cols = 11
      .Font = "MS Sans Serif"

      'Column Title
      .TextMatrix(0, 0) = "No"
      .TextMatrix(0, 1) = "Branch"
      .TextMatrix(0, 2) = "S.I"
      .TextMatrix(0, 3) = "Barrcode"
      .TextMatrix(0, 4) = "Descript"
      .TextMatrix(0, 5) = "Qty"
      .TextMatrix(0, 6) = "Refer#"
      .TextMatrix(0, 7) = "U.Price"
      .TextMatrix(0, 8) = "Tag"
      .TextMatrix(0, 9) = "trans#"
      .TextMatrix(0, 10) = "sStockIDx"
      .Row = 0

      'Column Alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = 4
      Next

      .ColWidth(0) = 450
      .ColWidth(1) = 2500
      .ColWidth(2) = 800
      .ColWidth(3) = 2500
      .ColWidth(4) = 2500
      .ColWidth(5) = 500
      .ColWidth(6) = 800
      .ColWidth(7) = 800
      .ColWidth(8) = 450
      .ColWidth(9) = 0
      .ColWidth(10) = 0

      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      .ColAlignment(4) = 1
      .ColAlignment(5) = 1
      .ColAlignment(6) = 1
      .ColAlignment(7) = 1
      .ColAlignment(8) = 1
      
      .Row = .Rows - 1
      .Col = 1
      
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

Private Sub oTrans_MasterRetrieved(ByVal Index As Integer)
   Select Case Index
   Case 15
      txtField(2).Text = oTrans.Master(Index)
   End Select
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   txtField(Index).BackColor = &HC0FFFF
   Select Case Index
      Case 1, 3, 4
      txtField(Index).Text = Format(txtField(Index).Text, "MM/DD/YY")
   End Select
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If Index = 80 Then
      If KeyCode = vbKeyReturn Or KeyCode = vbKeyF3 Then
         If oTrans.SearchTransaction(txtField(Index).Text, True) Then
            clearFields
            LoadMaster
            loadMasterDetail
         End If
      End If
   Else
      If KeyCode = vbKeyReturn Or KeyCode = vbKeyF3 Then
         If oTrans.SearchTransaction(txtField(Index).Text, False) Then
            clearFields
            LoadMaster
            loadMasterDetail
         End If
      End If
   End If
End Sub

Public Function setTransTat(nStat As Integer) As String
    Select Case nStat
       Case 0
          lblField(12) = "OPEN"
       Case 1
          lblField(12) = "CLOSED"
       Case 2
          lblField(12) = "POSTED"
       Case 3
          lblField(12) = "CANCELLED"
       Case 4
          lblField(12) = "VOID"
       Case Else
          lblField(12) = "UNKNOWN"
    End Select
End Function

Private Sub txtField_LostFocus(Index As Integer)
   txtField(Index).BackColor = &HFFFFFF
End Sub

Private Sub loadMasterDetail()
   Dim lnCtr As Integer
   Dim lsSQL As String
   Dim lorec As Recordset
   
   With GridEditor1
   If oTrans.ItemCount = 0 Then
      ClearDetail
      Else
         .Rows = 1
            For pnCtr = 0 To oTrans.ItemCount - 1
               .Rows = .Rows + 1
               
              lsSQL = "SELECT b.sBranchNm, a.sSalesInv, a.dTransact" & _
                        " FROM CP_SO_Master a" & _
                           " LEFT JOIN Branch b ON LEFT(a.sTransNox,4) = b.sBranchCd" & _
                        " WHERE a.sTransNox = " & strParm(IFNull(oTrans.Detail(lnCtr, "sSourceNo"), ""))
         
               lsSQL = lsSQL & _
                  " UNION ALL SELECT b.sBranchNm, a.sDocNmbrx sSalesInv, a.dTransact" & _
                        " FROM CP_Inventory_Adjustment a" & _
                           " LEFT JOIN Branch b ON LEFT(a.sTransNox,4) = b.sBranchCd" & _
                        " WHERE a.sTransNox = " & strParm(IFNull(oTrans.Detail(lnCtr, "sSourceNo"), ""))
               Set lorec = New Recordset
               lorec.Open lsSQL, oApp.Connection, , , adCmdText
               
               .TextMatrix(.Rows - 1, 0) = .Rows - 1
               If Not lorec.EOF Then
                  .TextMatrix(.Rows - 1, 1) = lorec("sBranchNm")
                  .TextMatrix(.Rows - 1, 2) = lorec("sSalesInv")
               Else
                  .TextMatrix(.Rows - 1, 1) = ""
                  .TextMatrix(.Rows - 1, 2) = ""
               End If
               
               .TextMatrix(.Rows - 1, 3) = IFNull(oTrans.Detail(pnCtr, "sBarrCode"), "")
               .TextMatrix(.Rows - 1, 4) = Format(oTrans.Detail(pnCtr, "sDescript"), "")
               .TextMatrix(.Rows - 1, 5) = IFNull(oTrans.Detail(pnCtr, "nItemQtyx"), 0)
               .TextMatrix(.Rows - 1, 6) = IFNull(oTrans.Detail(pnCtr, "sReferNox"), "")
               .TextMatrix(.Rows - 1, 7) = Format(IFNull(oTrans.Detail(pnCtr, "nUnitPrce"), 0), "#,##0.00")
               .TextMatrix(.Rows - 1, 8) = "Yes"
               .TextMatrix(.Rows - 1, 9) = IFNull(oTrans.Detail(pnCtr, "sSourceNo"), "")
               .TextMatrix(.Rows - 1, 10) = IFNull(oTrans.Detail(lnCtr, "sStockIDx"), "")
            Next
   End If
   End With
   ComputeTotal
End Sub


Private Sub InitForm()
   Dim lnCtr As Integer
   For lnCtr = 0 To txtField.Count - 1
      Select Case lnCtr
      Case 0
         txtField(lnCtr).Text = GetNextCode("CP_Consignment_Payment_Master", "sTransNox", True, oApp.Connection, True, oApp.BranchCode)
      Case 1, 3, 4
         txtField(lnCtr).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
      Case 2, 5
         txtField(lnCtr).Text = ""
      Case 6
          If Not IsNumeric(txtField(lnCtr)) Then
            txtField(lnCtr).Text = 0
         End If
      Case 7, 8, 9
         If Not IsNumeric(txtField(lnCtr)) Then
            txtField(lnCtr).Text = Format(0, "#,##0.00")
         End If
            
      txtField(11).Text = Format(0, "#,##0.00")
      End Select
   Next
   
'   CreateTempTable
End Sub

Private Sub ComputeTotal()
   Dim lnCtr As Integer
   Dim lnTotalQty As Double
   Dim lnTotalAmt As Currency
   
   lnTotalQty = 0
   lnTotalAmt = 0#
   With GridEditor1
      For lnCtr = 1 To .Rows - 1
         If .TextMatrix(lnCtr, 8) = "Yes" Then
            lnTotalQty = lnTotalQty + .TextMatrix(lnCtr, 5)
            lnTotalAmt = lnTotalAmt + (.TextMatrix(lnCtr, 5) * .TextMatrix(lnCtr, 7))
         End If
      Next
   End With
   txtField(6).Text = CInt(lnTotalQty)
   txtField(7).Text = Format(lnTotalAmt, "#,##0.00")
End Sub

Public Function PrintTrans() As Boolean
   Dim lrs As New ADODB.Recordset
   Dim lnCtr As Integer
   Dim lsOldProc As String
   Dim lsSQL As String
   Dim lorec As Recordset

   lsOldProc = "printTrans"
   ''On Error GoTo errProc
   
   PrintTrans = False
   
   Set lrs = New ADODB.Recordset
   
   lrs.Fields.Append "nField01", adInteger, 10
   lrs.Fields.Append "sField01", adVarChar, 200
   lrs.Fields.Append "sField02", adVarChar, 200
   lrs.Fields.Append "sField03", adVarChar, 200
   lrs.Fields.Append "sField04", adVarChar, 200
   lrs.Fields.Append "sField05", adVarChar, 200
   lrs.Fields.Append "lField01", adCurrency
   lrs.Fields.Append "lField02", adCurrency
   lrs.Open
      
   With GridEditor1
      For lnCtr = 0 To oTrans.ItemCount - 1
         lrs.AddNew
         lrs("nField01").Value = IFNull(oTrans.Detail(lnCtr, "nItemQtyx"), 0)
         
          lsSQL = "SELECT b.sBranchNm, a.sSalesInv, a.dTransact" & _
                        " FROM CP_SO_Master a" & _
                           " LEFT JOIN Branch b ON LEFT(a.sTransNox,4) = b.sBranchCd" & _
                        " WHERE a.sTransNox = " & strParm(IFNull(oTrans.Detail(lnCtr, "sSourceNo"), ""))
         
         lsSQL = lsSQL & _
               " UNION ALL SELECT b.sBranchNm, a.sDocNmbrx sSalesInv, a.dTransact" & _
                        " FROM CP_Inventory_Adjustment a" & _
                           " LEFT JOIN Branch b ON LEFT(a.sTransNox,4) = b.sBranchCd" & _
                        " WHERE a.sTransNox = " & strParm(IFNull(oTrans.Detail(lnCtr, "sSourceNo"), ""))

            Debug.Print lsSQL
               Set lorec = New Recordset
               lorec.Open lsSQL, oApp.Connection, , , adCmdText
         
         lrs("sField01").Value = lorec("sBranchNm")
         lrs("sField02").Value = lorec("dTransact")
         lrs("sField03").Value = lorec("sSalesInv")
         lrs("sField04").Value = IFNull(oTrans.Detail(lnCtr, "sBarrCode"), "")
         lrs("sField05").Value = IFNull(oTrans.Detail(lnCtr, "sDescript"), "")
         lrs("lField01").Value = Format(IFNull(oTrans.Detail(lnCtr, "nUnitPrce"), 0), "#,##0.00")
         lrs("lField02").Value = Format(IFNull(oTrans.Detail(lnCtr, "nUnitPrce"), 0), "#,##0.00") * IFNull(oTrans.Detail(lnCtr, "nItemQtyx"), 0)
      Next
   End With
   
   ' assign important info to the report
   Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\CPConsignment.rpt")
   oReport.DiscardSavedData
   oReport.FieldMappingType = crAutoFieldMapping
   oReport.Database.SetDataSource lrs

   oReport.Sections("RH").ReportObjects("txtTransNo").SetText Right(oTrans.Master("sTransNox"), 12)
   oReport.Sections("RH").ReportObjects("txtSupplier").SetText txtField(2).Text
   oReport.Sections("RH").ReportObjects("txtDate").SetText txtField(1).Text
   oReport.Sections("RH").ReportObjects("txtDFrom").SetText txtField(3).Text
   oReport.Sections("RH").ReportObjects("txtDThru").SetText txtField(4).Text
   oReport.Sections("RF").ReportObjects("txtNote").SetText txtField(5).Text
   
   oReport.PrintOutEx False, 1
   lrs.Close
   PrintTrans = True

endProc:
   If oTrans.Master("cTranStat") = xeStateOpen Then
      oTrans.CloseTransaction (oTrans.Master("sTransNox"))
   End If
   Set oReport = Nothing
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"
End Function

