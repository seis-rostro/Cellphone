VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmCPTransfer2MC 
   BorderStyle     =   0  'None
   Caption         =   "CP Transfer to MC"
   ClientHeight    =   4725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14340
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4725
   ScaleWidth      =   14340
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1800
      Index           =   0
      Left            =   1560
      Tag             =   "wt0;fb0"
      Top             =   1050
      Width           =   4785
      _ExtentX        =   8440
      _ExtentY        =   3175
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   660
         Index           =   4
         Left            =   1410
         MultiLine       =   -1  'True
         TabIndex        =   7
         Text            =   "frmCPTransfer2MC.frx":0000
         Top             =   960
         Width           =   3255
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   1
         Left            =   1410
         TabIndex        =   5
         Text            =   "mmmm dd yy"
         Top             =   615
         Width           =   3255
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   1410
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   "C001-21-000001"
         Top             =   120
         Width           =   1860
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   465
         TabIndex        =   6
         Top             =   1050
         Width           =   840
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   840
         TabIndex        =   4
         Top             =   660
         Width           =   465
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   360
         Left            =   1485
         Tag             =   "et0;ht2"
         Top             =   195
         Width           =   1860
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trans. No.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   21
         Left            =   285
         TabIndex        =   2
         Top             =   195
         Width           =   1020
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1740
      Index           =   1
      Left            =   1560
      Tag             =   "wt0;fb0"
      Top             =   2880
      Width           =   4785
      _ExtentX        =   8440
      _ExtentY        =   3069
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtDetail 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   1410
         TabIndex        =   9
         Text            =   "10"
         Top             =   450
         Width           =   1095
      End
      Begin VB.TextBox txtDetail 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   1410
         TabIndex        =   8
         Text            =   "01234567890"
         Top             =   80
         Width           =   3255
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IMEI/Serial:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   210
         TabIndex        =   17
         Top             =   140
         Width           =   1095
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   510
         TabIndex        =   16
         Top             =   495
         Width           =   795
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   4110
      Index           =   2
      Left            =   6405
      Tag             =   "wt0;fb0"
      Top             =   510
      Width           =   7855
      _ExtentX        =   13864
      _ExtentY        =   7250
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   4110
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   7860
         _ExtentX        =   13864
         _ExtentY        =   7250
         _Version        =   393216
         FixedCols       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   90
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1755
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
      Picture         =   "frmCPTransfer2MC.frx":002F
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   90
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   495
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
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
      Picture         =   "frmCPTransfer2MC.frx":07A9
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   90
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1125
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
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
      Picture         =   "frmCPTransfer2MC.frx":0F23
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   90
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1125
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
      Picture         =   "frmCPTransfer2MC.frx":169D
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   510
      Index           =   3
      Left            =   1560
      Tag             =   "wt0;fb0"
      Top             =   510
      Width           =   4785
      _ExtentX        =   8440
      _ExtentY        =   900
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   1410
         TabIndex        =   1
         Text            =   "GMC Dagupan - Honda"
         Top             =   60
         Width           =   3255
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Destination :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   9
         Left            =   105
         TabIndex        =   0
         Top             =   120
         Width           =   1185
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   90
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   495
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
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
      Picture         =   "frmCPTransfer2MC.frx":1E17
   End
End
Attribute VB_Name = "frmCPTransfer2MC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pxeMODULENAME = "frmCPTransfer2MC"
Private Const pxeVisibleRow = 20

Private oSkin As clsFormSkin
Private WithEvents oTrans As clsCPTransfer2MC
Attribute oTrans.VB_VarHelpID = -1

Private pnActiveRow As Integer
Private pbControl As Boolean
Private pnIndex As Integer
Private pbLoaded As Boolean
Private poRS As Recordset
Private pnCtr As Integer
Private pnRow As Integer

Private pbClosedTrans As Boolean

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String

   lsOldProc = "cmdButton_Click"
   'On Error GoTo errPro

   Select Case Index
      Case 0 'OK/save
         If oTrans.Master("sDestinat") <> "" And txtField(2).Text <> "" Then
            With oTrans
               If .Detail(.ItemCount - 1, "sbarrcode") = "" Then .deleteDetail (.ItemCount - 1)
               If oTrans.SaveTransaction Then
                  MsgBox "Transaction Saved Successfully.", vbInformation, "Notice"
                  If MsgBox("Do you want to print transaction?", _
                     vbQuestion + vbYesNo, "Confirm") = vbYes Then
      
                     Call PrintTrans
                  End If
                  InitTransaction
                  ClearDetail
               Else
                  MsgBox "Unable to Save Transaction.", vbInformation, "Notice"
               End If
            End With
         Else
            MsgBox "Invalid Destination Branch!!!", vbCritical, "Warning"
            txtField(2).SetFocus
         End If
      Case 1 'del
         Call deleteDetail
         txtDetail(11).SetFocus
      Case 2 'cancel
         Unload Me
      Case 3 'new
         InitTransaction
      
   End Select

'endProc:
   'Exit Sub
'errProc:
   'ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Activate()
   Dim lsOldProc As String

   lsOldProc = "Form_Activate"
   ''On Error GoTo errProc

   oApp.MenuName = Me.Tag
   Me.ZOrder 0
   
   If Not pbLoaded Then pbLoaded = True
   
'endProc:
   'Exit Sub
'errProc:
   'ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   With MSFlexGrid1
      Select Case KeyCode
         Case vbKeyReturn, vbKeyDown
            If (GetFocus = txtDetail(1).hwnd And txtDetail(6).Enabled = False) Or _
               GetFocus = txtDetail(6).hwnd Then
               txtDetail(1).SetFocus
            End If
            
            SetNextFocus
         Case vbKeyUp
            If GetFocus = txtField(2).hwnd Then Exit Sub
            
            SetPreviousFocus
      End Select
   End With
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   If pbControl Then
      If KeyCode = pbControl Then pbControl = False
   End If
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String
   lsOldProc = "Form_Load"

'   'On Error GoTo errProc

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualLeft
   CenterChildForm mdiMain, Me

   Set oTrans = New clsCPTransfer2MC
   Set oTrans.AppDriver = oApp
   oTrans.Branch = oApp.BranchCode
   Call InitTransaction
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub InitGrid()
      Dim lnCtr As Integer

   With MSFlexGrid1
      .Cols = 4
      .Rows = 2
      .Clear

      pnActiveRow = 0
      .Row = 0
      .TextMatrix(0, 0) = "No."
      .TextMatrix(0, 1) = "Serial"
      .TextMatrix(0, 2) = "Description"
      .TextMatrix(0, 3) = "Qty"
      
      .Row = 0
      'Column Width
      .ColWidth(0) = 600
      .ColWidth(1) = 2500
      .ColWidth(2) = 3950
      .ColWidth(3) = 700

      'Column Alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = 3
      Next
      .Row = 1
      'Column Alignment
      .TextMatrix(1, 0) = 1

      .ColAlignment(0) = flexAlignCenterCenter
      .ColAlignment(1) = flexAlignLeftCenter
      .ColAlignment(2) = flexAlignLeftCenter
      .ColAlignment(3) = flexAlignCenterCenter

      .Col = 0
      .ColSel = .Cols - 1
   End With
End Sub



Private Sub MSFlexGrid1_Click()
   With MSFlexGrid1
      If pbLoaded Then setDetailInfo

      .Col = 0
      .ColSel = .Cols - 1
      
      txtDetail(1).SetFocus
   End With
End Sub

Private Sub oTrans_DetailRetrieved(ByVal Index As Integer)
   With MSFlexGrid1
      Select Case Index
         Case 1
            .TextMatrix(.Row, 1) = oTrans.Detail(.Row - 1, "xReferNox")
         Case 2
            .TextMatrix(.Row, 2) = oTrans.Detail(.Row - 1, "sDescript")
         Case 3
            .TextMatrix(.Row, 3) = IFNull(oTrans.Detail(.Row - 1, "nQuantity"), "1")
      End Select
   End With

End Sub

Private Sub oTrans_MasterRetrieved(ByVal Index As Integer)
   Select Case Index
   Case 1
      txtField(1) = Format(oTrans.Master(Index), "Mmm dd, yyyy")
   Case 9
      txtField(2) = IFNull(oTrans.Master(Index))
   End Select
End Sub


Private Sub txtDetail_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
   Case 6
      oTrans.Detail(pnRow - 1, Index) = txtField(Index)
      addDetail
   End Select
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   Select Case Index
      Case 1, 2, 4
         Call HighlightOn(Me.txtField(Index))
   End Select
   
   If Index = 1 Then
      txtField(Index) = Format(oTrans.Master(1), "mm-dd-yyyy")
   End If
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   With oTrans
      Select Case KeyCode
      Case vbKeyReturn, vbKeyF3
         Select Case Index
         Case 2
            .SearchMaster 2, txtField(Index)
         Case 5
            If txtField(Index) <> "" Then
               If txtField(Index).Tag <> txtField(Index) Then
                 
               End If
            End If
            txtField(Index).Tag = txtField(Index)
         End Select
      End Select
   End With
End Sub

Private Sub LoadMaster()
   With oTrans
      txtField(0) = Format(.Master(0), "@@@@-@@-@@@@@@")
      txtField(1) = Format(.Master(1), "Mmm dd, yyyy")
      txtField(2) = IFNull(.Master(9))
   End With
End Sub


Private Sub txtField_LostFocus(Index As Integer)
   Select Case Index
      Case 1, 2, 4
         Call HighlightOff(Me.txtField(Index))
   End Select
   
   If Index = 1 Then
      txtField(Index) = Format(oTrans.Master(1), "Mmm dd, yyyy")
   End If
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   With oTrans
      Select Case Index
      Case 0
      Case 1
         If Not IsDate(txtField(Index)) Then
            .Master("dTransact") = oApp.ServerDate
         Else
            .Master("dTransact") = CDate(txtField(Index))
         End If
      Case 3
         .Master("sBranchNm") = txtField(Index)
      Case Else
         .Master(Index) = txtField(Index)
      End Select
   End With
End Sub

Private Sub txtDetail_GotFocus(Index As Integer)
   Select Case Index
      Case 1, 6
         Call HighlightOn(Me.txtDetail(Index))
   End Select
End Sub

Private Sub txtDetail_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  With MSFlexGrid1
      If oTrans.Master("sBranchNm") = "" Or IsNull(oTrans.Master("sBranchNm")) Then Exit Sub
      Select Case Index
         Case 1
            Select Case KeyCode
               Case vbKeyF3
                  If oTrans.searchDetail(.Row - 1, IIf(Index = 1, "xReferNox", "sDescript"), txtDetail(1)) Then
                     txtDetail(1).SetFocus
                  End If
               Case vbKeyReturn
                  If .Row = .Rows - 1 Then
                     If oTrans.searchDetail(.Row - 1, IIf(Index = 1, "xReferNox", "sDescript"), txtDetail(1)) Then
                        txtDetail(1).SetFocus
                     End If
                     Call addDetail
                  Else
                     .Row = .Rows - 1
                     .Col = 0
                     .ColSel = .Cols - 1
                     setDetailInfo
               End If
            End Select
      End Select
   End With
End Sub

Private Sub addDetail()
   Dim lsStockIDx As String
   
   With MSFlexGrid1
      If oTrans.Detail(.Row - 1, "sStockIDx") = "" Then Exit Sub
      lsStockIDx = oTrans.Detail(.Row - 1, "sStockIDx")
      
      'find matched reference # on cp order
      If Not TypeName(poRS) = "Nothing" Then
         If poRS.RecordCount > 0 Then
            poRS.MoveFirst
            poRS.Find "sStockIDx = " & strParm(lsStockIDx), 0, adSearchForward, adBookmarkFirst
            If Not poRS.EOF Then
               If .Row = .Rows - 1 Then
                  poRS("nIssuedxx") = poRS("nIssuedxx") + 1
                  
                  If poRS("nIssuedxx") > poRS("nQuantity") Then
                     MsgBox "Quantity to issue exceeds the request."
                  End If
               End If
            End If
            poRS.Cancel
         End If
      End If
      
      If oTrans.addDetail Then
         .Rows = .Rows + 1
         .Row = .Rows - 1
         .Col = 0
         .ColSel = .Cols - 1
         
         If .Rows > 16 Then
            .ColWidth(2) = 3700
            .TopRow = .Rows - 14
         Else
            .ColWidth(2) = 3950
         End If
         
         .TextMatrix(.Row, 0) = .Row
         ClearDetail
      End If
   End With
End Sub

Private Sub setDetailInfo()
   Dim lnRow As Integer

   pnRow = MSFlexGrid1.Row

   With oTrans
      txtDetail(1) = oTrans.Detail(pnRow - 1, "xReferNox")
      txtDetail(6) = oTrans.Detail(pnRow - 1, "nQuantity")
      
      txtDetail(6).Enabled = Trim(.Detail(pnRow - 1, "sStockIDx")) <> "" And .Detail(pnRow - 1, "cHsSerial") = xeNo
   End With
End Sub


Private Sub txtDetail_LostFocus(Index As Integer)
   Select Case Index
      Case 1
         Call HighlightOff(Me.txtDetail(Index))
   End Select
End Sub

Private Sub ClearFields()
   Dim loTxt As TextBox

   For Each loTxt In txtField
      loTxt = ""
   Next
   
   ClearDetail
End Sub

Private Sub ClearDetail()
   Dim loTxt As TextBox

   For Each loTxt In txtDetail
      loTxt = ""
   Next
End Sub

Private Sub deleteDetail()
   Dim lnCtr As Integer
   Dim lnRow As Integer
   Dim lsStockIDx As String

   With MSFlexGrid1
      lsStockIDx = oTrans.Detail(.Row - 1, "sStockIDx")
      'find matched engine # on mc order
      If TypeName(poRS) = "Nothing" Then GoTo NopoRS
      
      poRS.Find "sStockIDx = " & strParm(lsStockIDx), 0, adSearchForward, adBookmarkFirst

      If oTrans.deleteDetail(.Row - 1) Then
         lnRow = oTrans.ItemCount
NopoRS:
         lnRow = oTrans.ItemCount
         If lnRow = 0 Then
            oTrans.addDetail
            lnRow = 1
         Else
            If oTrans.Detail(MSFlexGrid1.Row - 1, "sBarrCode") <> "" Then
               Call oTrans.deleteDetail(MSFlexGrid1.Row - 1)
               If oTrans.ItemCount = 0 Then Call oTrans.addDetail
            End If
           
         End If

         InitGrid

         lnRow = oTrans.ItemCount
         .Rows = lnRow + 1

         For lnCtr = 0 To lnRow - 1
            .TextMatrix(lnCtr + 1, 0) = lnCtr + 1
            .TextMatrix(lnCtr + 1, 1) = IFNull(oTrans.Detail(lnCtr, "sBarrcode"))
            .TextMatrix(lnCtr + 1, 2) = IFNull(oTrans.Detail(lnCtr, "sDescript"))
            .TextMatrix(lnCtr + 1, 3) = IFNull(oTrans.Detail(lnCtr, "sBrandNme"))
            .TextMatrix(lnCtr + 1, 4) = IFNull(oTrans.Detail(lnCtr, "sModelNme"), oTrans.Detail(lnCtr, "sModelCde"))
            
         Next

         .Row = .Rows - 1
         .Col = 0
         .ColSel = .Cols - 1

         setDetailInfo
         Call ClearDetail
      End If
   End With
End Sub

Private Sub InitTransaction()
   InitGrid
   ClearFields

   oTrans.InitTransaction
   oTrans.NewTransaction
   LoadMaster
   setDetailInfo
   
   cmdButton(3).Visible = False
End Sub

Private Function BranchAutomate(ByVal sBranchCd As String) As Boolean
   Dim lrs As Recordset
   
   Set lrs = New Recordset
   lrs.Open "SELECT * FROM Branch" & _
               " WHERE sBranchCd = " & strParm(sBranchCd) & _
                  " AND cAutomate = " & strParm(xeYes) _
   , oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   
   If Not lrs.EOF Then BranchAutomate = True
   Set lrs = Nothing
End Function

Private Function PrintTrans() As Boolean
   Dim loreport As frmRepViewer
   Dim lrs As Recordset
   Dim lors As Recordset
   Dim lnCtr As Integer
   Dim lsOldProc As String
   Dim lnTotlWSerial As Double
   Dim lnTotlWOSerial As Double
   Dim lsSourceNo As String
   
   lsOldProc = "PrinTrans"
   ''On Error GoTo errProc

   PrintTrans = True
   
   Set lrs = New ADODB.Recordset

   lrs.Fields.Append "nField01", adInteger, 3
   lrs.Fields.Append "nField02", adChar, 1
   lrs.Fields.Append "sField01", adVarChar, 20
   lrs.Fields.Append "sField02", adVarChar, 128
   lrs.Fields.Append "sField03", adVarChar, 30
   lrs.Fields.Append "sField04", adVarChar, 12
   lrs.Fields.Append "sField05", adVarChar, 100
   lrs.Fields.Append "lField01", adInteger, 6
   lrs.Open

   'lsSourceNo = IFNull(oTrans.StockReqSourceNo, "")
   
   With oTrans
      lnTotlWOSerial = 0
      lnTotlWSerial = 0
      
      For lnCtr = 0 To .ItemCount - 1
         lrs.AddNew
         lrs.Fields("nField01") = oTrans.Detail(lnCtr, "nQuantity")
         lrs.Fields("nField02") = oTrans.Detail(lnCtr, "cHsSerial")
         lrs.Fields("sField01") = oTrans.Detail(lnCtr, "sBarrCode")
         lrs.Fields("sField02") = oTrans.Detail(lnCtr, "sDescript")
         lrs.Fields("sField03") = oTrans.Detail(lnCtr, "sSerialNo")
         lrs.Fields("sField04") = oTrans.Detail(lnCtr, "sStockIDx")
         lrs.Fields("sField05") = oTrans.Detail(lnCtr, "sBrandNme")
         lrs.Fields("lField01") = oTrans.Detail(lnCtr, "nUnitPrce")
         If oTrans.Detail(lnCtr, "cHsSerial") = xeYes Then
            lnTotlWSerial = lnTotlWSerial + 1
         Else
            lnTotlWOSerial = lnTotlWOSerial + CDbl(oTrans.Detail(lnCtr, "nQuantity"))
         End If
      Next
      lrs.Sort = "nField02 DESC,sField05,sField05,sField03"
   End With

   'assign important info to the report
   Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\CP_ServicePhoneTransfer.rpt")
   oReport.DiscardSavedData
   oReport.FieldMappingType = crAutoFieldMapping
   oReport.Database.SetDataSource lrs

   Set lors = New ADODB.Recordset
   If lors.State = adStateOpen Then lors.Close

   lors.Open "SELECT" _
               & "  a.sAddressx" _
               & ", CONCAT(b.sTownName, ', ' , c.sProvName, ' ' , b.sZippCode) xTownName" _
               & ", a.sBranchNm" _
            & " FROM Branch a" _
               & " LEFT JOIN TownCity b" _
                  & " LEFT JOIN Province c" _
                     & " ON b.sProvIDxx = c.sProvIDxx" _
                  & " ON a.sTownIDxx = b.sTownIDxx" _
            & " WHERE a.sBranchCd = " & strParm(oTrans.Master("sDestinat")) _
            , oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText

   oReport.Sections("PHa").ReportObjects("txtRefNo").SetText "CP" & "-" & Right(oTrans.Master("sTransNox"), 10)
   oReport.Sections("PHa").ReportObjects("txtDate").SetText txtField(1).Text
   oReport.Sections("PHb").ReportObjects("txtTo").SetText lors("sBranchNm")
   oReport.Sections("PHb").ReportObjects("txtToAddress").SetText lors("sAddressx") & IFNull(lors("xTownName"), "")
   oReport.Sections("PHb").ReportObjects("txtFrom").SetText oApp.ClientName
   oReport.Sections("PHb").ReportObjects("txtFromAddress").SetText oApp.Address & ", " & oApp.TownCity & ", " & oApp.Province & " " & oApp.ZippCode
   oReport.Sections("RFb").ReportObjects("txtRemarks").SetText lsSourceNo & " " & txtField(4).Text
   oReport.Sections("RFb").ReportObjects("txtWithSerial").SetText IIf(lnTotlWSerial = 0, "", Format(lnTotlWSerial, "#,##0"))
   oReport.Sections("RFb").ReportObjects("txtWOutSerial").SetText IIf(lnTotlWOSerial = 0, "", Format(lnTotlWOSerial, "#,##0"))
   oReport.Sections("PF").ReportObjects("txtRptUser").SetText oApp.UserName

   Set loreport = New frmRepViewer
   Set loreport.ReportSource = oReport
   loreport.Show
   
   PrintTrans = True

endPoc:
   If Not pbClosedTrans Then
      If BranchAutomate(oTrans.Master("sDestinat")) Then
         If oTrans.CloseTransaction(oTrans.Master(0)) Then pbClosedTrans = True
      End If
   End If
   Set loreport = Nothing
   Set oReport = Nothing
   Set lrs = Nothing
   Set lors = Nothing
   Exit Function
errProc:
   PrintTrans = False
  ' ShowError lsOldProc & "( " & " )"
End Function

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




