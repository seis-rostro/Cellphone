VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0#1.0#0"; "xrGridControl.ocx"
Begin VB.Form frmSP_Transfer_Posting 
   BorderStyle     =   0  'None
   Caption         =   "Spareparts Transfer Posting"
   ClientHeight    =   6765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11685
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6765
   ScaleWidth      =   11685
   ShowInTaskbar   =   0   'False
   Begin xrGridEditor.GridEditor GridEditor1 
      Height          =   2940
      Left            =   165
      TabIndex        =   20
      Tag             =   "et0;eb0;et0;bc2"
      Top             =   2865
      Width           =   9840
      _ExtentX        =   17357
      _ExtentY        =   5186
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
      Object.HEIGHT          =   2940
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
      MOUSEICON       =   "frmCP_TransferPosting.frx":0000
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
      Height          =   5505
      Index           =   0
      Left            =   75
      Tag             =   "wt0;fb0"
      Top             =   1110
      Width           =   10065
      _ExtentX        =   17754
      _ExtentY        =   9710
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   9
         Left            =   8040
         TabIndex        =   19
         Tag             =   "ht0"
         Text            =   "0.00"
         Top             =   1215
         Width           =   1890
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   10
         Left            =   8040
         TabIndex        =   15
         Top             =   900
         Width           =   1890
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   8
         Left            =   1275
         MaxLength       =   8
         TabIndex        =   24
         Top             =   5055
         Width           =   4590
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   1230
         TabIndex        =   7
         Top             =   90
         Width           =   1950
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   8040
         TabIndex        =   11
         Top             =   585
         Width           =   1890
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   1245
         TabIndex        =   9
         Top             =   585
         Width           =   5415
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   1245
         MaxLength       =   15
         TabIndex        =   17
         Top             =   1215
         Width           =   5415
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   1245
         TabIndex        =   13
         Top             =   900
         Width           =   5415
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   1275
         MaxLength       =   30
         TabIndex        =   22
         Top             =   4755
         Width           =   4590
      End
      Begin VB.Shape Shape4 
         Height          =   330
         Index           =   0
         Left            =   6750
         Top             =   120
         Width           =   3150
      End
      Begin VB.Shape Shape3 
         Height          =   390
         Index           =   0
         Left            =   6720
         Top             =   90
         Width           =   3210
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "unknown"
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
         Left            =   6780
         TabIndex        =   30
         Tag             =   "eb0;et0"
         Top             =   150
         Width           =   3090
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gross Total"
         Height          =   195
         Index           =   3
         Left            =   6780
         TabIndex        =   18
         Top             =   1275
         Width           =   810
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Discount"
         Height          =   195
         Index           =   11
         Left            =   6780
         TabIndex        =   14
         Top             =   960
         Width           =   630
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Approved"
         Height          =   195
         Index           =   10
         Left            =   75
         TabIndex        =   23
         Top             =   5130
         Width           =   690
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   6225
         TabIndex        =   25
         Top             =   4890
         Width           =   705
      End
      Begin VB.Label lblTranTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   7020
         TabIndex        =   26
         Tag             =   "ht0;ft0"
         Top             =   4755
         Width           =   2790
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   285
         Left            =   1320
         Tag             =   "et0;ht2"
         Top             =   180
         Width           =   1950
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction No"
         Height          =   195
         Index           =   0
         Left            =   75
         TabIndex        =   6
         Top             =   150
         Width           =   1095
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Date"
         Height          =   195
         Index           =   1
         Left            =   6765
         TabIndex        =   10
         Top             =   645
         Width           =   1230
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Source/Origin"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   8
         Top             =   630
         Width           =   990
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   195
         Index           =   4
         Left            =   105
         TabIndex        =   12
         Top             =   960
         Width           =   570
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Request"
         Height          =   195
         Index           =   5
         Left            =   90
         TabIndex        =   16
         Top             =   1275
         Width           =   600
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   195
         Index           =   6
         Left            =   105
         TabIndex        =   21
         Top             =   4845
         Width           =   630
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H8000000F&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   285
         Index           =   0
         Left            =   6780
         Tag             =   "et0;et0"
         Top             =   150
         Width           =   3105
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   10350
      TabIndex        =   29
      Top             =   3285
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
      Picture         =   "frmCP_TransferPosting.frx":001C
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   10350
      TabIndex        =   27
      Top             =   2025
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
      Picture         =   "frmCP_TransferPosting.frx":0796
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   10350
      TabIndex        =   28
      Top             =   2655
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Post"
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
      Picture         =   "frmCP_TransferPosting.frx":0F10
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   525
      Index           =   1
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   926
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
         Index           =   13
         Left            =   7830
         MaxLength       =   50
         TabIndex        =   5
         Top             =   90
         Width           =   2070
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
         Index           =   12
         Left            =   3930
         MaxLength       =   50
         TabIndex        =   3
         Top             =   90
         Width           =   2505
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
         Index           =   11
         Left            =   1215
         TabIndex        =   1
         Top             =   90
         Width           =   1950
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Date Received"
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
         Left            =   6510
         TabIndex        =   4
         Top             =   135
         Width           =   1305
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Source"
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
         Left            =   3285
         TabIndex        =   2
         Top             =   135
         Width           =   600
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Transact. No"
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
         Left            =   60
         TabIndex        =   0
         Top             =   135
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmSP_Transfer_Posting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmSP_Transfer_Posting"

Private WithEvents oTrans As clsSPTransfer
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin

Dim pnCtr As Integer, pnIndex As Integer

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   Dim lnRep As Integer

   lsOldProc = "cmdButton_Click"
   On Error GoTo errProc

   txtField_LostFocus pnIndex
   With GridEditor1
      Select Case Index
      Case 0
         If oTrans.SearchAcceptance() Then
            LoadMaster
            LoadDetail
         Else
            If txtField(0).Text = "" Then ClearFields
         End If
         
         txtField(pnIndex).SetFocus
         .Refresh
      Case 1
         If txtField(0).Text <> "" Then
            If oTrans.Master("cTranStat") <> xeStatePosted Then
               lnRep = MsgBox("Post Transaction!!!", vbYesNo + vbQuestion, "Confrim")
   
               If lnRep = vbYes Then
                  If Not oTrans.AcceptDelivery(CDate(txtField(13).Text)) Then
                     MsgBox "Unable to Post Transaction!!!", vbCritical, "Warning"
                  Else
                     MsgBox "Transaction Post Successfully!!!", vbInformation, "Confirm"
                     If oTrans.OpenTransaction(oTrans.Master("sTransNox")) Then
                        LoadMaster
                        LoadDetail
                     End If
                  End If
               End If
            Else
               MsgBox "Transaction already posted!!!" & vbCrLf & _
                        "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
            End If
         Else
            MsgBox "Unable to Post Transaction!!!" & vbCrLf & _
                     "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
         End If
      Case 2
         Unload Me
      End Select
   End With
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0
   
   With GridEditor1
      .Refresh
   End With
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String
   
   lsOldProc = "Form_Load"
   On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oTrans = New clsSPTransfer
   Set oTrans.AppDriver = oApp
   oTrans.TransStatus = 10
   
   oTrans.InitTransaction

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransMaintenance
   
   InitForm
   ClearFields

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oTrans = Nothing
   Set oSkin = Nothing
End Sub

Private Sub oTrans_DetailRetrieved(ByVal Index As Integer)
   With GridEditor1
      .TextMatrix(.Row, Index) = oTrans.Detail(.Row - 1, Index)
   End With
End Sub

Private Sub oTrans_MasterRetrieved(ByVal Index As Integer)
   txtField(Index).Text = oTrans.Master(Index)
End Sub

Private Sub ClearFields()
   For pnCtr = 0 To 13
      Select Case pnCtr
      Case 9, 10
         txtField(pnCtr).Text = "0.00"
      Case 6, 7
      Case 13
         txtField(pnCtr).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
      Case Else
         txtField(pnCtr).Text = ""
         txtField(pnCtr).Tag = ""
      End Select
   Next

   Label2.Caption = "UNKNOWN"
   lblTranTotal.Caption = "0.00"

   With GridEditor1
      .Rows = 2
      .ColWidth(2) = 3450
 
      .TextMatrix(1, 1) = ""
      .TextMatrix(1, 2) = ""
      .TextMatrix(1, 3) = ""
      .TextMatrix(1, 4) = ""
      .TextMatrix(1, 5) = "0"
      .TextMatrix(1, 6) = "0"
      .TextMatrix(1, 7) = "0.00"
      .TextMatrix(1, 8) = "0.00"
   End With
 End Sub

Private Sub InitForm()
   Dim lnCtr As Integer

   With GridEditor1
      .Cols = 9
      .Rows = 2
      .Font = "MS Sans Serif"

      'Column Title
      .TextMatrix(0, 1) = "Barcode"
      .TextMatrix(0, 2) = "Description"
      .TextMatrix(0, 3) = "PT"
      .TextMatrix(0, 4) = "Model"
      .TextMatrix(0, 5) = "QOH"
      .TextMatrix(0, 6) = "Qty"
      .TextMatrix(0, 7) = "Unit Price"
      .TextMatrix(0, 8) = "Sub-Total"
      
      .Row = 0

      'Column Alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = 3
         .ColEnabled(lnCtr) = False
      Next

      'Column Width
      .ColWidth(0) = 300
      .ColWidth(1) = 1900
      .ColWidth(3) = 400
      .ColWidth(4) = 690
      .ColWidth(5) = 500
      .ColWidth(6) = 500
      .ColWidth(7) = 1000
      .ColWidth(8) = 1000
      
      .ColAlignment(1) = 1

      .ColNumberOnly(5) = True
      .ColNumberOnly(6) = True
      .ColNumberOnly(7) = True
      .ColNumberOnly(8) = True

      .ColFormat(5) = "#,##0"
      .ColFormat(6) = "#,##0"
      .ColFormat(7) = "#,###,##0.00"
      .ColFormat(8) = "#,###,##0.00"

      .ColDefault(5) = "0"
      .ColDefault(6) = "0"
      .ColDefault(7) = "0.00"
      .ColDefault(8) = "0.00"

      .TextMatrix(1, 1) = ""
      .TextMatrix(1, 2) = ""
      .TextMatrix(1, 3) = ""
      .TextMatrix(1, 4) = ""
      .TextMatrix(1, 5) = "0"
      .TextMatrix(1, 6) = "0"
      .TextMatrix(1, 7) = "0.00"
      .TextMatrix(1, 8) = "0.00"

      .Row = 1
      .Col = 1
   End With
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      If Index = 13 Then .Text = Format(.Text, "MM/DD/YY")
      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = oApp.getColor("HT1")
   End With
   
   pnIndex = Index
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lbFound As Boolean
   Dim lsOldProc As String
   
   lsOldProc = "txtField_KeyDown"
   On Error GoTo errProc
   
   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
      With txtField(Index)
         Select Case Index
         Case 11, 12
            If .Text = "" Then
               ClearFields
               Exit Sub
            End If

            If oTrans.SearchTransaction(IIf(Index = 11, CodeFormat(oApp.BranchCode, .Text), .Text), IIf(Index = 11, True, False)) = True Then
               LoadMaster
               LoadDetail
            Else
               ClearFields
            End If

            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
         End Select
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

Private Sub LoadMaster()
   For pnCtr = 0 To 12
      Select Case pnCtr
      Case 0, 11
         txtField(pnCtr).Text = Format(oTrans.Master("sTransNox"), "@@@@-@@@@@@")
         txtField(pnCtr).Tag = txtField(pnCtr).Text
      Case 1
         txtField(pnCtr).Text = Format(oTrans.Master("dTransact"), "MMMM DD, YYYY")
      Case 2, 12
         txtField(pnCtr).Text = oTrans.Master("xSourcexx")
         txtField(pnCtr).Tag = txtField(pnCtr).Text
      Case 3, 4, 5, 8
         txtField(pnCtr).Text = IIf(IsNull(oTrans.Master(pnCtr)), "", oTrans.Master(pnCtr))
      Case 9, 10
         txtField(pnCtr).Text = FormatNumber(IIf(IsNull(oTrans.Master(pnCtr)), "", oTrans.Master(pnCtr)), 2)
      Case Else
      End Select
   Next
   
   Label2.Caption = Format(TransStat(oTrans.Master("cTranStat")), ">")
   lblTranTotal.Caption = Format(oTrans.Master("nTranTotl"), "#,##0.00")
End Sub

Private Sub LoadDetail()
   Dim lnRow As Integer
   Dim lnCol As Integer

   With GridEditor1
      .Rows = IIf(oTrans.ItemCount = 0, 2, oTrans.ItemCount + 1)
     
      For lnRow = 0 To oTrans.ItemCount - 1
         For lnCol = 1 To 7
            .TextMatrix(lnRow + 1, lnCol) = oTrans.Detail(lnRow, lnCol)
         Next
         .TextMatrix(lnRow + 1, 8) = oTrans.Detail(lnRow, "nQuantity") * oTrans.Detail(lnRow, "nUnitPrce")
      Next
      ComputeTranTotal
   End With
End Sub

Private Sub ComputeTranTotal()
   Dim lnCtr As Integer

   With GridEditor1
      txtField(9).Text = 0#
      For lnCtr = 1 To .Rows - 1
         txtField(9).Text = FormatNumber(CDbl(txtField(9).Text) + CDbl(.TextMatrix(lnCtr, 8)), 2)
      Next
   oTrans.Master("nGrossAmt") = CDbl(txtField(9).Text)
   End With
   
   lblTranTotal = FormatNumber(CDbl(txtField(9).Text) * (100 - CDbl(txtField(10).Text)) / 100, 2)
   
   oTrans.Master("nTranTotl") = CDbl(lblTranTotal)
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   With txtField(Index)
      If Index = 13 Then
         If Not IsDate(.Text) Then .Text = oApp.ServerDate
         .Text = Format(.Text, "MMMM DD, YYYY")
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

