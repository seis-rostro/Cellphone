VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0#1.0#0"; "xrGridControl.ocx"
Begin VB.Form frmCP_TransferOrder 
   BorderStyle     =   0  'None
   Caption         =   "Spareparts Transfer"
   ClientHeight    =   7800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11940
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7800
   ScaleWidth      =   11940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin xrGridEditor.GridEditor GridEditor1 
      Height          =   4650
      Left            =   1635
      TabIndex        =   14
      Tag             =   "et0;eb0;et0;bc2"
      Top             =   2205
      Width           =   10140
      _ExtentX        =   17886
      _ExtentY        =   8202
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
      Object.HEIGHT          =   4650
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
      MOUSEICON       =   "frmCP_TransferOrder.frx":0000
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
      Height          =   7170
      Left            =   1560
      Tag             =   "wt0;fb0"
      Top             =   525
      Width           =   10290
      _ExtentX        =   18150
      _ExtentY        =   12647
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
         Left            =   8295
         TabIndex        =   13
         Text            =   "0.00"
         Top             =   1170
         Width           =   1890
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   10
         Left            =   8295
         TabIndex        =   11
         Top             =   840
         Width           =   1890
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   8
         Left            =   1290
         MaxLength       =   8
         TabIndex        =   18
         Top             =   6735
         Width           =   2625
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   1290
         TabIndex        =   1
         Top             =   90
         Width           =   1950
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   8295
         TabIndex        =   9
         Top             =   510
         Width           =   1890
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   1290
         TabIndex        =   3
         Top             =   540
         Width           =   5415
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   4
         Left            =   1290
         MaxLength       =   15
         TabIndex        =   7
         Top             =   1200
         Width           =   5415
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   1290
         TabIndex        =   5
         Top             =   870
         Width           =   5415
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   5
         Left            =   1290
         MaxLength       =   30
         TabIndex        =   16
         Top             =   6405
         Width           =   4590
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gross Total"
         Height          =   195
         Index           =   3
         Left            =   6930
         TabIndex        =   12
         Top             =   1230
         Width           =   810
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Discount"
         Height          =   195
         Index           =   11
         Left            =   6930
         TabIndex        =   10
         Top             =   900
         Width           =   630
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Approved"
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   17
         Top             =   6780
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
         Left            =   6540
         TabIndex        =   19
         Top             =   6465
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
         Left            =   7395
         TabIndex        =   20
         Tag             =   "et0;hb0"
         Top             =   6390
         Width           =   2790
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   315
         Left            =   1380
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
         Left            =   90
         TabIndex        =   0
         Top             =   150
         Width           =   1095
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Date"
         Height          =   195
         Index           =   1
         Left            =   6930
         TabIndex        =   8
         Top             =   555
         Width           =   1230
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Destination"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   2
         Top             =   600
         Width           =   795
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   4
         Top             =   930
         Width           =   570
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Request"
         Height          =   195
         Index           =   5
         Left            =   90
         TabIndex        =   6
         Top             =   1260
         Width           =   600
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   195
         Index           =   6
         Left            =   105
         TabIndex        =   15
         Top             =   6450
         Width           =   630
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   90
      TabIndex        =   23
      Top             =   4320
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
      Picture         =   "frmCP_TransferOrder.frx":001C
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   90
      TabIndex        =   21
      Top             =   3060
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
      Picture         =   "frmCP_TransferOrder.frx":0796
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   90
      TabIndex        =   22
      Top             =   3690
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
      Picture         =   "frmCP_TransferOrder.frx":0F10
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   90
      TabIndex        =   25
      Top             =   5580
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
      Picture         =   "frmCP_TransferOrder.frx":168A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   90
      TabIndex        =   24
      Top             =   4950
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "Print"
      AccessKey       =   "Print"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCP_TransferOrder.frx":1E04
   End
End
Attribute VB_Name = "frmCP_TransferOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmSP_TransferOrder"

Private WithEvents oSPTransfer As clsSPTransfer
Attribute oSPTransfer.VB_VarHelpID = -1
Private oSkin As clsFormSkin
Private pbCancel As Boolean

Dim pnCtr As Integer
Dim pnIndex As Integer
Dim pbGridFocus As Boolean
Dim pbPrint As Boolean

Property Get Cancel() As Boolean
   Cancel = pbCancel
End Property

Property Set SPTransfer(loSPTransfer As clsSPTransfer)
   Set oSPTransfer = loSPTransfer
End Property

Property Get SPTransfer() As clsSPTransfer
   Set SPTransfer = oSPTransfer
End Property

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   Dim lsRep As String
   
   lsOldProc = "cmdButton_Click"
   On Error GoTo errProc

   With GridEditor1
      Select Case Index
      Case 0 'Save
         If .Rows > 2 Then
            pnCtr = 0
            Do While pnCtr < .Rows
               If Trim(.TextMatrix(pnCtr, 1)) = "" Then
                  .Row = pnCtr
                  If oSPTransfer.DeleteDetail(.Row - 1) Then .DeleteRow
               Else
                  pnCtr = pnCtr + 1
               End If
            Loop
            
            .ColWidth(2) = 3450
            If .Rows > 14 Then .ColWidth(2) = 3250
         End If
         
         If isEntryOK Then
            If oSPTransfer.SaveTransaction = True Then
               MsgBox "Transaction Saved Successfully!!!", vbInformation, "Notice"
               txtField(8).Text = oApp.getUserName(IIf(IsNull(oSPTransfer.Master(8)), "", oSPTransfer.Master(8)))
               
               lsRep = MsgBox("Print Transaction!!!", vbYesNo + vbQuestion, "Confirm")
               If lsRep = vbYes Then
                  If Not PrintTrans Then MsgBox "Unable to Print Transaction!!!", vbCritical, "Warning"
               End If
               
               pbCancel = False
               pbPrint = True
            Else
               MsgBox "Unable to Save Transaction!!!", vbCritical, "Warning"
            End If
         End If
      Case 1 'Search
         If pbGridFocus Then
            If .Col = 1 Or .Col = 2 Then
               If oSPTransfer.SearchDetail(.Row - 1, .Col) Then .Col = 1
               .Refresh
               .SetFocus
            End If
         Else
            oSPTransfer.SearchMaster pnIndex
         End If
      Case 2 'Delete Row
         If .Rows > 2 Then
            If oSPTransfer.DeleteDetail(.Row - 1) Then .DeleteRow
            
            .ColWidth(2) = 3450
            If .Rows > 14 Then .ColWidth(2) = 3250
            ComputeTranTotal
         End If
      Case 3 'Cancel
         Unload Me
      Case 4 'Print
         If pbPrint Then
            If txtField(0).Text <> "" Then
               lsRep = MsgBox("Print Transaction!!!", vbYesNo + vbQuestion, "Confirm")
               If lsRep = vbYes Then
                  If Not PrintTrans Then MsgBox "Unable to Print Transaction!!!", vbCritical, "Warning"
               End If
            End If
         Else
            MsgBox "Unable to Print Transaction!!!" & vbCrLf & _
                     "Save transaction first to continue!!!", vbCritical, "Warning"
         End If
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
   
   InitValue
   pbCancel = True
   GridEditor1.Refresh
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
   On Error GoTo errProc
   
   CenterChildForm mdiMain, Me

   Set oSPTransfer = New clsSPTransfer
   Set oSPTransfer.AppDriver = oApp

   oSPTransfer.InitTransaction
   oSPTransfer.NewTransaction
   
   InitGrid

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransaction

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub GridEditor1_AddingRow(Cancel As Boolean)
   With GridEditor1
      If .TextMatrix(.Row, 1) = "" Then
         Cancel = True
      ElseIf .TextMatrix(.Row, 6) = 0 Then
         Cancel = True
      End If
      If Not Cancel Then oSPTransfer.AddDetail
      
      If .Rows > 14 Then .ColWidth(2) = 3450
   End With
End Sub

Private Sub GridEditor1_EditorValidate(Cancel As Boolean)
   With GridEditor1
      If .Col = 6 Then
         If Trim(.TextMatrix(.Row, 1)) <> "" Then
            If CDbl(.TextMatrix(.Row, .Col)) > CDbl(.TextMatrix(.Row, 5)) Then
               MsgBox "Item doesn't have enough stock!!!" & vbCrLf & vbCrLf & _
                      "Verify your entry then try again!!!", vbCritical, "Warning!"
               .TextMatrix(.Row, .Col) = 0
            End If
         End If
         oSPTransfer.Detail(.Row - 1, .Col) = CDbl(.TextMatrix(.Row, .Col))
      Else
         oSPTransfer.Detail(.Row - 1, .Col) = .TextMatrix(.Row, .Col)
      End If
      
      .TextMatrix(.Row, 8) = oSPTransfer.Detail(.Row - 1, "nQuantity") * oSPTransfer.Detail(.Row - 1, "nUnitPrce")
      ComputeTranTotal
   End With
End Sub

Private Sub GridEditor1_GotFocus()
   pbGridFocus = True
End Sub

Private Sub GridEditor1_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String
   
   lsOldProc = "GridEditor1_KeyDown"
   On Error GoTo errProc
   
   If KeyCode = vbKeyF3 Then
      With GridEditor1
         If .Col = 1 Or .Col = 2 Then
            If oSPTransfer.SearchDetail(.Row - 1, .Col, .TextMatrix(.Row, .Col)) Then .Col = 6
            .Refresh
            .SetFocus
         End If
      End With
      KeyCode = 0
   End If

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " _
                       & "  " & KeyCode _
                       & ", " & Shift _
                       & " )", True
End Sub

Private Sub oSPTransfer_DetailRetrieved(ByVal Index As Integer)
   With GridEditor1
      .TextMatrix(.Row, Index) = oSPTransfer.Detail(.Row - 1, Index)
   End With
End Sub

Private Sub oSPTransfer_MasterRetrieved(ByVal Index As Integer)
   txtField(Index).Text = oSPTransfer.Master(Index)
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   If Index = 1 Then txtField(Index).Text = Format(txtField(Index).Text, "MM/DD/YY")
   If txtField(Index) <> Empty Then
      txtField(Index).SelStart = 0
      txtField(Index).SelLength = Len(txtField(Index).Text)
   End If
   pbGridFocus = False
   pnIndex = Index
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String
   
   lsOldProc = "txtField_KeyDown"
   On Error GoTo errProc

   If KeyCode = vbKeyF3 Then
      If Index = 2 Then oSPTransfer.SearchMaster Index, txtField(Index).Text
      If txtField(Index).Text <> "" Then SetNextFocus
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

Private Sub InitGrid()
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
      Next

      'Column Width
      .ColWidth(0) = 300
      .ColWidth(1) = 1900
      .ColWidth(3) = 400
      .ColWidth(4) = 1000
      .ColWidth(5) = 500
      .ColWidth(6) = 500
      .ColWidth(7) = 1000
      .ColWidth(8) = 1000
      
      .ColAlignment(1) = 1
         
      .ColEnabled(3) = False
      .ColEnabled(4) = False
      .ColEnabled(5) = False
      .ColEnabled(7) = False
      .ColEnabled(8) = False

      .ColNumberOnly(5) = True
      .ColNumberOnly(6) = True
      .ColNumberOnly(7) = True
      .ColNumberOnly(8) = True

      .ColFormat(5) = "#,##0"
      .ColFormat(6) = "#,##0"
      .ColFormat(7) = "#,##0.00"
      .ColFormat(8) = "#,##0.00"

      .ColDefault(5) = 0
      .ColDefault(6) = 0
      .ColDefault(7) = 0#
      .ColDefault(8) = 0#

      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1

      .Row = 1
   End With
End Sub

Private Sub InitValue()
   For pnCtr = 0 To 10
      Select Case pnCtr
      Case 0
         txtField(pnCtr).Text = Format(oSPTransfer.Master(pnCtr), "@@@@-@@@@@@")
      Case 1
         txtField(pnCtr).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
      Case 6, 7
      Case 9, 10
         txtField(pnCtr).Text = Format(oSPTransfer.Master(pnCtr), "#,##0.00")
      Case Else
         txtField(pnCtr).Text = oSPTransfer.Master(pnCtr)
      End Select
   Next
   
   lblTranTotal.Caption = "0.00"
   
   With GridEditor1
      .Rows = IIf(oSPTransfer.ItemCount = 0, 2, oSPTransfer.ItemCount + 1)
      
      For pnCtr = 0 To oSPTransfer.ItemCount - 1
         .TextMatrix(pnCtr + 1, 1) = oSPTransfer.Detail(pnCtr, "sBarrCode")
         .TextMatrix(pnCtr + 1, 2) = oSPTransfer.Detail(pnCtr, "sDescript")
         .TextMatrix(pnCtr + 1, 3) = oSPTransfer.Detail(pnCtr, "sTypeCode")
         .TextMatrix(pnCtr + 1, 4) = oSPTransfer.Detail(pnCtr, "sModelNme")
         .TextMatrix(pnCtr + 1, 5) = oSPTransfer.Detail(pnCtr, "nQtyOnHnd")
         .TextMatrix(pnCtr + 1, 6) = oSPTransfer.Detail(pnCtr, "nQuantity")
         .TextMatrix(pnCtr + 1, 7) = oSPTransfer.Detail(pnCtr, "nUnitPrce")
         .TextMatrix(pnCtr + 1, 8) = oSPTransfer.Detail(pnCtr, "nQuantity") * oSPTransfer.Detail(pnCtr, "nUnitPrce")
      Next
      
      .ColWidth(2) = 3450
      If .Rows > 14 Then .ColWidth(2) = 3200
      
      .ColEnabled(1) = True
      .ColEnabled(2) = True
      .ColEnabled(6) = True
      .Row = .Rows - 1
      .Col = 1
   End With
   ComputeTranTotal
   pbPrint = False
End Sub

Private Sub ComputeTranTotal()
   Dim lnCtr As Integer

   With GridEditor1
      txtField(9).Text = 0#
      For lnCtr = 1 To .Rows - 1
         txtField(9).Text = FormatNumber(CDbl(txtField(9).Text) + CDbl(.TextMatrix(lnCtr, 8)), 2)
      Next
   oSPTransfer.Master("nGrossAmt") = CDbl(txtField(9).Text)
   End With
   
   'Convert CDbl to Val
   lblTranTotal = FormatNumber(CDbl(txtField(9).Text) * (100 - Val(txtField(10).Text)) / 100, 2)
   
   oSPTransfer.Master("nTranTotl") = CDbl(lblTranTotal)
End Sub

Private Function isEntryOK() As Boolean
   If txtField(2).Text = "" Then
      MsgBox "Invalid Destination Detected!!!" & vbCrLf & _
             "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
      txtField(2).SetFocus
      GoTo EntryNotOK
   End If

   With GridEditor1
      If .TextMatrix(1, 1) = "" Then
         MsgBox "Detail is required!!!" & vbCrLf & _
                "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
         .SetFocus
         .Row = 1
         .Col = 1
         GoTo EntryNotOK
      End If
   End With
   
EntryOK:
   isEntryOK = True
   Exit Function
EntryNotOK:
   isEntryOK = False
End Function

Public Function PrintTrans() As Boolean
   Dim lrs As New ADODB.Recordset
   Dim lnCtr As Integer
   Dim lsOldProc As String
   
   lsOldProc = "PrintTrans"
   On Error GoTo errProc

   PrintTrans = True
   
   Set lrs = New ADODB.Recordset
   
   lrs.Fields.Append "nEntryNo", adInteger, 3
   lrs.Fields.Append "sBarrCode", adVarChar, 23
   lrs.Fields.Append "sDescription", adVarChar, 60
   lrs.Fields.Append "sModel", adVarChar, 30
   lrs.Fields.Append "nQuantity", adInteger, 5
   lrs.Fields.Append "nUnitPrice", adDouble, 20
   lrs.Fields.Append "nTotal", adDouble, 20
   lrs.Open
      
   With oSPTransfer
      For lnCtr = 0 To .ItemCount - 1
         lrs.AddNew
         lrs("nEntryNo").Value = .Detail(lnCtr, "nEntryNox")
         lrs("sBarrCode").Value = .Detail(lnCtr, "sBarrCode")
         lrs("sDescription").Value = .Detail(lnCtr, "sDescript")
         lrs("sModel").Value = .Detail(lnCtr, "sModelNme")
         lrs("nQuantity").Value = .Detail(lnCtr, "nQuantity")
         lrs("nUnitPrice").Value = .Detail(lnCtr, "nUnitPrce")
         lrs("nTotal").Value = .Detail(lnCtr, "nUnitPrce") * .Detail(lnCtr, "nQuantity")
      Next
   End With
   
   ' assign important info to the report
   Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\Branch Transfer.rpt")
   oReport.DiscardSavedData
   oReport.FieldMappingType = crAutoFieldMapping
   oReport.Database.SetDataSource lrs

   oReport.Sections("RHa").ReportObjects("txtRefNo").SetText "SP" & "-" & Right(oSPTransfer.Master("sTransNox"), 8)
   oReport.Sections("RHa").ReportObjects("txtDate").SetText txtField(1).Text
   oReport.Sections("RHb").ReportObjects("txtToBranch").SetText txtField(2).Text
   oReport.Sections("RHb").ReportObjects("txtToAddress").SetText txtField(3).Text
   oReport.Sections("RF").ReportObjects("txtRemarks").SetText txtField(5).Text
   oReport.Sections("RF").ReportObjects("txtPrepared").SetText oApp.UserName
   oReport.Sections("RF").ReportObjects("txtApproved").SetText txtField(8).Text
   oReport.Sections("RF").ReportObjects("txtDisc").SetText txtField(10).Text
   oReport.Sections("RF").ReportObjects("txtNet").SetText lblTranTotal.Caption
   
   oReport.PrintOutEx False, 1
   lrs.Close

endProc:
   oSPTransfer.CloseTransaction oSPTransfer.Master("sTransNox")
   Set oReport = Nothing
   Exit Function
errProc:
   PrintTrans = False
   ShowError lsOldProc & "( " & " )"
End Function

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   With txtField(Index)
      .Text = TitleCase(.Text)
      Select Case Index
      Case 1
         If Not IsDate(.Text) Then .Text = oApp.ServerDate
         .Text = Format(.Text, "MMMM DD, YYYY")
         oSPTransfer.Master(Index) = .Text & " " & Format(oApp.ServerDate, "h:mm:ss AM/PM")
      Case 10
         If Not IsNumeric(.Text) Then .Text = 0
         .Text = Format(.Text, "#,##0.00")

         ComputeTranTotal
         oSPTransfer.Master(Index) = CDbl(.Text)
      Case Else
         oSPTransfer.Master(Index) = .Text
      End Select
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
