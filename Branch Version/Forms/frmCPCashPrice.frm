VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmCPCashPrice 
   BackColor       =   &H8000000C&
   BorderStyle     =   0  'None
   Caption         =   "CP Price List"
   ClientHeight    =   7800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11340
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7800
   ScaleWidth      =   11340
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame2 
      Height          =   525
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   570
      Width           =   9660
      _ExtentX        =   17039
      _ExtentY        =   926
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtFilter 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2535
         TabIndex        =   3
         Top             =   90
         Width           =   6975
      End
      Begin VB.ComboBox cmbSearch 
         Appearance      =   0  'Flat
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
         Left            =   90
         TabIndex        =   2
         Top             =   90
         Width           =   2400
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   6600
      Index           =   0
      Left            =   75
      Tag             =   "wt0;fb0"
      Top             =   1110
      Width           =   9690
      _ExtentX        =   17092
      _ExtentY        =   11642
      BorderStyle     =   1
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   6420
         Left            =   15
         TabIndex        =   1
         Tag             =   "et0;eb0;et0;bc2"
         Top             =   90
         Width           =   9630
         _ExtentX        =   16986
         _ExtentY        =   11324
         _Version        =   393216
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Left            =   9990
      TabIndex        =   0
      Top             =   7095
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
      Picture         =   "frmCPCashPrice.frx":0000
   End
End
Attribute VB_Name = "frmCPCashPrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmCashPrice"

Private oPriceList As clsCPPriceList
Private oSkin As clsFormSkin

Dim pbCtrlPress As Boolean
Dim lnSearch As Integer

Private Sub cmbSearch_Click()

   With MSFlexGrid1
      If cmbSearch.ListIndex = 0 Then 'Model Code
         .Col = 2
         lnSearch = 2
      ElseIf cmbSearch.ListIndex = 1 Then 'Model Name
         .Col = 1
         lnSearch = 1
      End If
   End With
End Sub

Private Sub cmdButton_Click()
   Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyReturn, vbKeyUp, vbKeyDown
      Select Case KeyCode
      Case vbKeyReturn
         If GetFocus = MSFlexGrid1.hwnd Then Exit Sub
         SetNextFocus
      Case vbKeyDown
         If pbCtrlPress Then
            With MSFlexGrid1
               If .Row < .Rows - 1 Then
                  .Row = .Row + 1
                  .Col = 1
                  .ColSel = .Cols - 1
                  If .Row > 11 Then .TopRow = .Row
                     
               End If
            End With
         Else
            SetNextFocus
         End If
      Case vbKeyUp
         If pbCtrlPress Then
            With MSFlexGrid1
               If .Row > 1 Then
                  If .Row = .TopRow Then .TopRow = .TopRow - 1
                  .Row = .Row - 1
                  .ColSel = .Cols - 1
               End If
            End With
         Else
            SetPreviousFocus
         End If
      End Select
   Case vbKeyControl
      pbCtrlPress = True
      
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
   DoEvents

   Call LoadPrice

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub InitForm()
   Dim lnCtr As Integer

   With MSFlexGrid1
      .Cols = 6
      .Rows = 2
      .Font = "MS Sans Serif"

      'Column Title
      .TextMatrix(0, 1) = "Model Name"
      .TextMatrix(0, 2) = "Model Code"
      .TextMatrix(0, 3) = "Brand"
      .TextMatrix(0, 4) = "Selling Price"
      .TextMatrix(0, 5) = "Last Price"
'      .TextMatrix(0, 5) = "SRP 2"
'      .TextMatrix(0, 6) = "SRP 3"
'      .TextMatrix(0, 7) = "SRP 4"
      .Row = 0

      'Column Alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = 3
      Next
      .Row = 1

      'Column Width
      .ColWidth(0) = 450
      .ColWidth(1) = 3450
      .ColWidth(2) = 1600
      .ColWidth(3) = 1600
      .ColWidth(4) = 1200
      .ColWidth(5) = 1200
'      .ColWidth(6) = 0
'      .ColWidth(7) = 0
'      .ColWidth(8) = 1200

      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
   End With

   cmbSearch.AddItem "Model Code"
   cmbSearch.AddItem "Model Name"
   
End Sub

Private Sub LoadPrice()
   Dim lnRow As Integer
   Dim lnCol As Integer
   Dim lnCtr As Integer
   Dim lnCell As Integer
   Dim ldLatest As Date
   Dim lsCategory As String
   Dim lsModelNme As String
   Dim lasFormat(5) As String

   lasFormat(0) = "@"
   lasFormat(1) = "@"
   lasFormat(2) = "@"
   lasFormat(3) = "@"
   lasFormat(4) = "#,##0.00"
   lasFormat(5) = "#,##0.00"
'   lasFormat(6) = "#,##0.00"
'   lasFormat(7) = "#,##0.00"
'   lasFormat(8) = "#,##0.00"

   With oPriceList
      Call .LoadCashPrice
      MSFlexGrid1.Rows = 1
      lsCategory = ""
      lsModelNme = ""
      lnCtr = 0
      lnRow = 1

      ldLatest = .CashLatestDate
      Do While lnCtr < .CashPriceCount
         If lsCategory <> .CashPrice(lnCtr, "sBrandNme") Then
            If lnRow + 1 > MSFlexGrid1.Rows Then MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1

            MSFlexGrid1.TextMatrix(lnRow, 1) = .CashPrice(lnCtr, "sBrandNme")

            MSFlexGrid1.Row = lnRow
            For lnCol = 1 To MSFlexGrid1.Cols - 1
               MSFlexGrid1.Col = lnCol

               MSFlexGrid1.CellFontBold = True
               MSFlexGrid1.CellBackColor = oApp.getColor("HT1")
            Next

            lsCategory = .CashPrice(lnCtr, "sBrandNme")
            lnRow = lnRow + 1
         End If

         If lnRow + 1 > MSFlexGrid1.Rows Then MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
         MSFlexGrid1.Row = lnRow
         For lnCol = 1 To MSFlexGrid1.Cols - 1
            MSFlexGrid1.Col = lnCol
            MSFlexGrid1.TextMatrix(lnRow, lnCol) = Format(IFNull(.CashPrice(lnCtr, lnCol), 0), lasFormat(lnCol))
            If ldLatest = .CashPrice(lnCtr, 6) Then MSFlexGrid1.CellFontBold = True
         Next
         lnRow = lnRow + 1
         lnCtr = lnCtr + 1
      Loop
                  
      If MSFlexGrid1.Rows > 26 Then
         MSFlexGrid1.ColWidth(1) = 3200
      Else
         MSFlexGrid1.ColWidth(1) = 3450
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

Private Function SearchOn(ByVal lsSeek, ByVal lnIndex) As Boolean
   Dim lnCtr As Long
   Dim lbFound As Boolean

   lbFound = False
   With MSFlexGrid1
      For lnCtr = 1 To .Rows
         If StrComp(Left(.TextMatrix(lnCtr, lnIndex), Len(lsSeek)), lsSeek, vbTextCompare) >= 0 Then
            .TopRow = lnCtr
            .Row = lnCtr
            .RowSel = lnCtr
            .Col = lnIndex
            .ColSel = .Cols - 1
            lbFound = True
            Exit For
         End If
      Next
      .ColSel = lnIndex
      .Sort = flexSortGenericAscending
   End With
   SearchOn = lbFound
End Function

Private Function ResultingText(iKeyAscii%) As String
   Dim sLeft As String
   Dim sSel As String
   Dim sRight As String
   Dim sResult As String

   On Error Resume Next

   With txtFilter
      sLeft = Left$(.Text, .SelStart)
      sSel = Mid$(.Text, .SelStart + 1, .SelLength)
      sRight = Mid$(.Text, .SelStart + .SelLength + 1)
   End With

   Select Case iKeyAscii
   Case vbKeyBack
      If Len(sSel) = 0 Then
         sResult = MinusRightChar(sLeft) & sRight
      Else
         sResult = sLeft & sRight
      End If
   Case vbKeyDelete
      If Len(sSel) = 0 Then
         sResult = sLeft & MinusLeftChar(sRight)
      Else
         sResult = sLeft & sRight
      End If
   Case Else
      sResult = sLeft & Chr$(iKeyAscii) & sRight
   End Select

   ResultingText = sResult
End Function

Private Sub txtFilter_KeyPress(keyascii As Integer)
   Dim lsSearchOn As String

   On Error Resume Next

   If keyascii = vbKeyReturn Or keyascii = vbKeyTab Then Exit Sub

   lsSearchOn = ResultingText(keyascii)
   If SearchOn(lsSearchOn, lnSearch) = False Then
      keyascii = 0
   End If
End Sub

