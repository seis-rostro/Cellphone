VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmServicePhoneTagging 
   BorderStyle     =   0  'None
   Caption         =   "Service Phone Tagging"
   ClientHeight    =   4710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4710
   ScaleWidth      =   13680
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
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   4
         Left            =   1440
         MultiLine       =   -1  'True
         TabIndex        =   20
         Top             =   1080
         Width           =   3255
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   1
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "mmmm dd yy"
         Top             =   650
         Width           =   3255
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   0
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "C00121-000001"
         Top             =   120
         Width           =   2775
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
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
         Left            =   90
         TabIndex        =   3
         Top             =   1200
         Width           =   780
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
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
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   405
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   420
         Left            =   1750
         Tag             =   "et0;ht2"
         Top             =   200
         Width           =   2775
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction No."
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
         Left            =   90
         TabIndex        =   1
         Top             =   190
         Width           =   1485
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
      Begin VB.TextBox txtDetails 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   8
         Left            =   1800
         TabIndex        =   17
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txtDetails 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   5
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   930
         Width           =   2895
      End
      Begin VB.TextBox txtDetails 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   6
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   550
         Width           =   2895
      End
      Begin VB.TextBox txtDetails 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   11
         Left            =   1080
         TabIndex        =   14
         Top             =   80
         Width           =   3615
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
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
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Price"
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
         Index           =   5
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Purchase Price"
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
         Index           =   4
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IMEI"
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
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   405
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   4145
      Index           =   2
      Left            =   6400
      Tag             =   "wt0;fb0"
      Top             =   480
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   7303
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   4140
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Width           =   7155
         _ExtentX        =   12621
         _ExtentY        =   7303
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
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1740
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
      Picture         =   "frmServicePhoneTagging.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   90
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   480
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
      Picture         =   "frmServicePhoneTagging.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   90
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1110
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
      Picture         =   "frmServicePhoneTagging.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   90
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1110
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
      Picture         =   "frmServicePhoneTagging.frx":166E
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
         Left            =   1560
         TabIndex        =   21
         Top             =   60
         Width           =   3135
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
         Left            =   240
         TabIndex        =   13
         Top             =   120
         Width           =   1185
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   90
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   480
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
      Picture         =   "frmServicePhoneTagging.frx":1DE8
   End
End
Attribute VB_Name = "frmServicePhoneTagging"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pxeMODULENAME = "frmServicePhoneTagging"
Private Const pxeVisibleRow = 20

Private oSkin As clsFormSkin
Private WithEvents oTrans As clsCPServicePhoneTransfer
Attribute oTrans.VB_VarHelpID = -1

Private pnPrintRow As Integer
Private poPrinter As clsPrintDirect
Private Const pxeMaxLine As Integer = 65

Private pnActiveRow As Integer
Private pbControl As Boolean
Private pnIndex As Integer
Private pbLoaded As Boolean
Private poRS As Recordset
Private pnCtr As Integer

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
                  ClearOthers
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
         txtDetails(11).SetFocus
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
            If GetFocus = txtDetails(11).hwnd Then Exit Sub
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

   ''On Error GoTo errProc

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualLeft
   CenterChildForm mdiMain, Me

   Set oTrans = New clsCPServicePhoneTransfer
   Set oTrans.AppDriver = oApp
   oTrans.Branch = oApp.BranchCode
   Call InitTransaction
   pnPrintRow = 0
   
  
   
'endProc:
 '  Exit Sub
'errProc:
 '  ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub InitGrid()
      Dim lnCtr As Integer

   With MSFlexGrid1
      .Cols = 5
      .Rows = 2
      .Clear

      pnActiveRow = 0
      .Row = 0
      .TextMatrix(0, 0) = "No."
      .TextMatrix(0, 1) = "BRAND"
      .TextMatrix(0, 2) = "MODEL"
      .TextMatrix(0, 3) = "IMEI"
      .TextMatrix(0, 4) = "QTY."
      
      .Row = 0
      'Column Width
      .ColWidth(0) = 600
      .ColWidth(1) = 1800
      .ColWidth(2) = 1800
      .ColWidth(3) = 2150
      .ColWidth(4) = 700

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
      .ColAlignment(3) = flexAlignLeftCenter
      .ColAlignment(4) = flexAlignLeftCenter

      .Col = 0
      .ColSel = .Cols - 1
      
   End With
End Sub



Private Sub MSFlexGrid1_Click()
   With MSFlexGrid1
      If pbLoaded Then setDetailInfo

      .Col = 0
      .ColSel = .Cols - 1
      
      txtDetails(11).SetFocus
   End With
End Sub

Private Sub MSFlexGrid1_RowColChange()
   'If pbLoaded Then
    '  txtDetails(11).SetFocus
   'End If
End Sub

Private Sub oTrans_DetailRetrieved(ByVal Index As Integer)
   With MSFlexGrid1
      Select Case Index
         
         Case 3
            .TextMatrix(.Row, 3) = oTrans.Detail(.Row - 1, "xReferNox")
         Case 2
            .TextMatrix(.Row, 2) = oTrans.Detail(.Row - 1, "sBrandNme")
         Case 1
            .TextMatrix(.Row, 1) = IFNull(oTrans.Detail(.Row - 1, "sModelNme"), "") 'oTrans.Detail(.Row - 1, "sModelCde"))
         Case 4
            .TextMatrix(.Row, 4) = IFNull(oTrans.Detail(.Row - 1, "nQuantity"), "1")
      End Select
   End With

End Sub

Private Sub oTrans_MasterRetrieved(ByVal Index As Integer)
   Select Case Index
   Case 1
      txtField(Index) = strLongDate(oTrans.Master("dTransact"))
   Case 2
      txtField(Index) = IFNull(oTrans.Master("sBranchNm"))
   Case 9
      txtField(Index) = IFNull(oTrans.Master("sSourceNo"))
   Case 0
      

   End Select
End Sub


Private Sub txtField_GotFocus(Index As Integer)
   Select Case Index
      Case 1, 2, 4
         Call HighlightOn(Me.txtField(Index))
      Case 3
         If Len(txtField(Index)) <> 0 Then
            txtDetails(11).SetFocus
         Else
            Call HighlightOn(Me.txtField(Index))
         End If
   End Select
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   With oTrans
      Select Case KeyCode
      Case vbKeyReturn
         Select Case Index
         
         Case 2
               .SearchMaster 2, ""
         Case 5
            If txtField(Index) <> "" Then
               If txtField(Index).Tag <> txtField(Index) Then
                 
               End If
            End If
            txtField(Index).Tag = txtField(Index)
         End Select
      Case vbKeyF3
         Select Case Index
         Case 2
            .SearchMaster 2, ""
         Case 5
            txtField(Index) = txtField(Index).Tag
         End Select
      End Select
   End With
End Sub

Private Sub LoadMaster()
   With oTrans
      txtField(0) = Format(.Master("sTransNox"), "@@@@-@@-@@@@@@")
      txtField(1) = strLongDate(.Master("dTransact"))
      txtField(2) = IFNull(.Master(2))
      
      If Not pbLoaded Then Exit Sub
      
   End With
End Sub



Private Sub txtField_LostFocus(Index As Integer)
   Select Case Index
      Case 1, 2, 3, 4
         Call HighlightOff(Me.txtField(Index))
   End Select
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   With oTrans
      Select Case Index
      Case 1
         .Master("dTransact") = txtField(Index)
      Case 3
         .Master("sBranchNm") = txtField(Index)
      Case Else
         .Master(Index) = txtField(Index)
      End Select
   End With
End Sub

Private Sub txtDetails_GotFocus(Index As Integer)
   Select Case Index
      Case 1
         Call HighlightOn(Me.txtDetails(Index))
   End Select
End Sub

Private Sub txtDetails_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  With MSFlexGrid1
      If oTrans.Master("sBranchNm") = "" Or IsNull(oTrans.Master("sBranchNm")) Then Exit Sub
      Select Case Index
         Case 11
            Select Case KeyCode
               Case vbKeyF3
                  If oTrans.searchDetail(.Row - 1, IIf(Index = 11, "xReferNox", "sDescript"), txtDetails(11)) Then
                     txtDetails(11).SetFocus
                     Call addDetail
                  End If
               Case vbKeyReturn
                  If .Row = .Rows - 1 Then
                     If oTrans.searchDetail(.Row - 1, IIf(Index = 11, "xReferNox", "sDescript"), txtDetails(11)) Then
                        txtDetails(11).SetFocus
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
                  
                  'With MSFlexGrid2
                   '  .TextMatrix(poRS.AbsolutePosition, 7) = poRS("nIssuedxx")
                    ' .Row = poRS.AbsolutePosition
                     'If .Row > 14 Then .TopRow = .Row - 13
                     '.Col = 1
                     '.ColSel = .Cols - 1
                  'End With
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
            .ColWidth(1) = 2595
            .ColWidth(2) = 2595
            .TopRow = .Rows - 14
         Else
'            .ColWidth(1) = 2500
'            .ColWidth(2) = 3200
         End If
         
         .TextMatrix(.Row, 0) = .Row
         ClearOthers
      End If
   End With
End Sub

Private Sub setDetailInfo()
   Dim lnRow As Integer

   lnRow = MSFlexGrid1.Row

   With oTrans
      txtDetails(11) = oTrans.Detail(lnRow - 1, "xReferNox")
      txtDetails(6) = IFNull(oTrans.Detail(lnRow - 1, "nPurPrice"), "###,###.00")
      txtDetails(5) = IFNull(oTrans.Detail(lnRow - 1, "nUnitPrce"), "###,###.00")
      txtDetails(8) = oTrans.Detail(lnRow - 1, "nQuantity")
   End With
End Sub


Private Sub txtDetails_LostFocus(Index As Integer)
   Select Case Index
      Case 1
         Call HighlightOff(Me.txtDetails(Index))
   End Select
End Sub
Private Sub ClearOthers()
   Dim loTxt As TextBox

   For Each loTxt In txtDetails
      loTxt = ""
   Next
      txtField(4) = ""
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
         Call ClearOthers
      End If
   End With
End Sub

Private Sub InitTransaction()
   Call InitGrid

   'oTrans.IsCellphoneUnits = True
   oTrans.InitTransaction
   oTrans.NewTransaction
   Call LoadMaster
   cmdButton(3).Visible = False

End Sub

Private Function PrintTrans() As Boolean
   Dim loreport As frmRepViewer
   Dim lrs As Recordset
   Dim lors As Recordset
   Dim lnCtr As Integer
   Dim lsOldProc As String
   Dim lnTotlWSerial As Double
   Dim lnTotlWOSerial As Double
   Dim lsSourceNo As String
   
   lsOldProc = "PrintTrans"
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
   If oTrans.Master("cTranStat") = xeStateOpen Then
      If Not oTrans.CloseTransaction(oTrans.Master(0)) Then
         MsgBox "Unable to Close Transaction. Please inform MIS.", vbCritical, "Warning"
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





