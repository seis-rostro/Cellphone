VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0"; "xrGridControl.ocx"
Begin VB.Form frmSuppliesPurchasePosting 
   BorderStyle     =   0  'None
   Caption         =   "Supplies Purchase Posted"
   ClientHeight    =   6810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11085
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6810
   ScaleWidth      =   11085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   9705
      TabIndex        =   9
      Top             =   1770
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
      Picture         =   "frmSuppliesPurchasePosting.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   9705
      TabIndex        =   10
      Top             =   2385
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
      Picture         =   "frmSuppliesPurchasePosting.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   9705
      TabIndex        =   11
      Top             =   495
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
      Picture         =   "frmSuppliesPurchasePosting.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   9705
      TabIndex        =   12
      Top             =   1125
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
      Picture         =   "frmSuppliesPurchasePosting.frx":166E
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   465
      Index           =   1
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   495
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   820
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
         Height          =   315
         Index           =   8
         Left            =   1320
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   75
         Width           =   1590
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
         Height          =   315
         Index           =   9
         Left            =   4350
         MaxLength       =   50
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   75
         Width           =   4965
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Trans. No."
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
         Left            =   75
         TabIndex        =   14
         Top             =   105
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier"
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
         Left            =   3555
         TabIndex        =   13
         Top             =   105
         Width           =   1410
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   5745
      Index           =   0
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   990
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   10134
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   4
         Left            =   1215
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   1920
         Width           =   3705
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   5
         Left            =   1215
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   2340
         Width           =   1950
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   6
         Left            =   4125
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   2340
         Width           =   2010
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   7
         Left            =   7095
         TabIndex        =   8
         Text            =   "Text 1"
         Top             =   2340
         Width           =   1995
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   1200
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   945
         Width           =   1905
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   1200
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1320
         Width           =   4230
      End
      Begin VB.TextBox txtField 
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
         Index           =   0
         Left            =   1200
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   375
         Width           =   1905
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   630
         Index           =   3
         Left            =   5460
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   1005
         Width           =   3660
      End
      Begin xrGridEditor.GridEditor GridEditor1 
         Height          =   2775
         Left            =   90
         TabIndex        =   16
         Top             =   2850
         Width           =   9195
         _ExtentX        =   16219
         _ExtentY        =   4895
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
         Object.HEIGHT          =   2775
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
         MOUSEICON       =   "frmSuppliesPurchasePosting.frx":1DE8
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
      Begin VB.Shape Shape3 
         Height          =   390
         Index           =   0
         Left            =   5460
         Top             =   435
         Width           =   2505
      End
      Begin VB.Shape Shape4 
         Height          =   330
         Index           =   0
         Left            =   5490
         Top             =   465
         Width           =   2445
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
         Left            =   5505
         TabIndex        =   25
         Tag             =   "eb0;et0"
         Top             =   495
         Width           =   2385
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   315
         Left            =   1335
         Tag             =   "et0;ht2"
         Top             =   480
         Width           =   1920
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   24
         Top             =   1980
         Width           =   795
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Entry No"
         Height          =   195
         Index           =   3
         Left            =   270
         TabIndex        =   23
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
         Height          =   195
         Index           =   2
         Left            =   3300
         TabIndex        =   22
         Top             =   2400
         Width           =   585
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Price"
         Height          =   195
         Index           =   6
         Left            =   6270
         TabIndex        =   21
         Top             =   2400
         Width           =   690
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date "
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   20
         Top             =   990
         Width           =   810
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Trans. No."
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
         Left            =   240
         TabIndex        =   19
         Top             =   405
         Width           =   1110
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier"
         Height          =   195
         Index           =   9
         Left            =   240
         TabIndex        =   18
         Top             =   1380
         Width           =   570
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   195
         Index           =   5
         Left            =   4680
         TabIndex        =   17
         Top             =   990
         Width           =   630
      End
      Begin VB.Shape Shape2 
         Height          =   1620
         Index           =   2
         Left            =   105
         Top             =   135
         Width           =   9165
      End
      Begin VB.Shape Shape2 
         Height          =   1035
         Index           =   1
         Left            =   105
         Top             =   1770
         Width           =   9165
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H8000000F&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   285
         Index           =   0
         Left            =   5520
         Tag             =   "et0;et0"
         Top             =   495
         Width           =   2430
      End
   End
End
Attribute VB_Name = "frmSuppliesPurchasePosting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Private Const pxeMODULENAME = "frmCP_Purchasing_Posting"
'
'Private WithEvents oTrans As ggcSuppliesPurchase
'Private oSkin As clsFormSkin
'
'Dim pnCtr As Integer
'Dim pbLoad As Boolean
'Dim pnIndex As Integer
'
'Private Sub cmdButton_Click(Index As Integer)
'   Dim lsOldProc As String
'   Dim lsRep As String
'
'   lsOldProc = "cmdButton_Click"
'   'On Error GoTo errProc
'
'   txtField_LostFocus pnIndex
'   With GridEditor1
'      Select Case Index
'      Case 0 'Browse
'         If oTrans.SearchTransaction() = True Then
'            LoadMaster
'            LoadDetail
'         End If
'         txtField(9).SetFocus
'      Case 1 'Post
'         If pbLoad Then
'            If oTrans.Master("cTranStat") <> xeStatePosted Then
'               lsRep = MsgBox("Post Transaction!!!", vbYesNo + vbQuestion, "Confrim")
'
'               If lsRep = vbYes Then
'                  If Not oTrans.PostTransaction(oTrans.Master("sTransNox")) Then
'                     MsgBox "Unable to Post Transaction!!!", vbCritical, "Warning"
'                  Else
'                     MsgBox "Transaction Post Successfully!!!", vbInformation, "Confirm"
'                     If oTrans.OpenTransaction(oTrans.Master("sTransNox")) Then
'                        LoadMaster
'                        LoadDetail
'                     Else
'                        ClearFields
'                     End If
'                  End If
'               End If
'            Else
'               MsgBox "Unable to Post Transaction!!!" & vbCrLf & _
'                        "Transaction already posted!!!", vbCritical, "Warning"
'            End If
'         Else
'            MsgBox "Unable to Post Transaction!!!" & vbCrLf & _
'                   "No Transaction Loaded!!!", vbCritical, "Warning"
'         End If
'      Case 2 'Print
'         If oTrans.Master("cTranStat") = xeStatePosted Then
'            If txtField(0).Text <> "" Then
'               lsRep = MsgBox("Print Transaction!!!", vbYesNo + vbQuestion, "Confirm")
'               If lsRep = vbYes Then
'                  If Not PrintTrans Then MsgBox "Unable to Print Transaction!!!", vbCritical, "Warning"
'               End If
'            End If
'         Else
'            MsgBox "Unable to Print Transaction!!!", vbCritical, "Warning"
'         End If
'      Case 3 'Close
'         Unload Me
'      End Select
'   End With
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " & Index & " )", True
'End Sub
'
'Private Sub Form_Activate()
'   oApp.MenuName = Me.Tag
'   Me.ZOrder 0
'
'   GridEditor1.Refresh
'End Sub
'
'Private Sub Form_Load()
'   Dim lsOldProc As String
'
'   lsOldProc = "Form_Load"
'   'On Error GoTo errProc
'
'   CenterChildForm mdiMain, Me
'
'   Set oTrans = New ggcSuppliesPurchase
'   Set oTrans.AppDriver = oApp
'   oTrans.TransStatus = 10
'   'oTrans.InitTransaction
'
'   Set oSkin = New clsFormSkin
'   Set oSkin.AppDriver = oApp
'   Set oSkin.Form = Me
'   oSkin.ApplySkin xeFormTransMaintenance
'
'
'   ClearFields
'   InitGrid
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " & " )", True
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'   Set oTrans = Nothing
'   Set oSkin = Nothing
'End Sub
'
''Private Sub oTrans_DetailRetrieved(ByVal Index As Integer)
''   With GridEditor1
''      .TextMatrix(.Row, Index) = oTrans.Detail(.Row - 1, Index)
''   End With
''End Sub
''
''Private Sub oTrans_MasterRetrieved(ByVal Index As Integer)
''   txtField(Index).Text = oTrans.Master(Index)
''End Sub
'
'Private Sub ClearFields()
'   For pnCtr = 0 To txtField.Count - 1
'     Select Case pnCtr
'      Case 1
'         txtField(pnCtr).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
'      Case Else
'         txtField(pnCtr).Text = Empty
'         txtField(pnCtr).Tag = Empty
'      End Select
'   Next
'
'   Label2.Caption = "UNKNOWN"
'
'   With GridEditor1
'      .Rows = 2
'      .Col = 1
'      .ColWidth(2) = 5180
'
'      'empty row
'      .TextMatrix(1, 1) = ""
'      .TextMatrix(1, 2) = ""
'      .TextMatrix(1, 3) = ""
'      .TextMatrix(1, 4) = ""
'      .TextMatrix(1, 5) = 0
'   End With
'
'   pbLoad = False
' End Sub
'
'
'
'Private Sub txtField_GotFocus(Index As Integer)
'   With txtField(Index)
'      .SelStart = 0
'      .SelLength = Len(.Text)
'      .BackColor = oApp.getColor("HT1")
'   End With
'
'   pnIndex = Index
'End Sub
'
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'   Select Case KeyCode
'   Case vbKeyReturn, vbKeyUp, vbKeyDown
'      Select Case KeyCode
'      Case vbKeyReturn, vbKeyDown
'         If GetFocus = GridEditor1.hwnd Then Exit Sub
'         SetNextFocus
'      Case vbKeyUp
'         SetPreviousFocus
'      End Select
'   End Select
'End Sub
'
'Private Sub LoadMaster()
'   For pnCtr = 0 To txtField.Count - 1
'      Select Case pnCtr
'      Case 0, 9
'         txtField(pnCtr).Text = Format(oTrans.Master(0), "@@@@-@@@@@@")
'         txtField(pnCtr).Tag = txtField(pnCtr).Text
'      Case 1, 7
'         txtField(pnCtr).Text = Format(oTrans.Master(pnCtr), "MMMM DD, YYYY")
'      Case 2, 10
'         txtField(pnCtr).Text = oTrans.Master(2)
'         txtField(pnCtr).Tag = txtField(pnCtr).Text
'      Case Else
'         txtField(pnCtr).Text = IIf(IsNull(oTrans.Master(pnCtr)), "", oTrans.Master(pnCtr))
'      End Select
'   Next
'
'   Label2.Caption = Format(TransStat(oTrans.Master("cTranStat")), ">")
'   pbLoad = True
'End Sub
'
'Private Sub LoadDetail()
'   Dim lnCtr As Integer
'
'   With GridEditor1
'      .Rows = IIf(oTrans.ItemCount = 0, 2, oTrans.ItemCount + 1)
'
'      .ColWidth(2) = 5180
'      If .Rows > 16 Then .ColWidth(2) = 4940
'      For pnCtr = 0 To oTrans.ItemCount - 1
'         For lnCtr = 1 To 5
'            .TextMatrix(pnCtr + 1, lnCtr) = oTrans.Detail(pnCtr, lnCtr)
'         Next
'      Next
'   End With
'End Sub
'
'Public Function PrintTrans() As Boolean
'   Dim loreport As frmRepViewer
'
'   Dim lors As ADODB.Recordset
'   Dim lrs As ADODB.Recordset
'   Dim lnCtr As Integer
'   Dim lsOldProc As String
'
'   lsOldProc = "PrintTrans"
'   'On Error GoTo errProc
'
'   PrintTrans = True
'
'   Set lors = New ADODB.Recordset
'
'   lors.Fields.Append "nQuantity", adInteger, 3
'   lors.Fields.Append "sModel", adVarChar, 50
'   lors.Fields.Append "sColor", adVarChar, 50
'   lors.Fields.Append "sDescription", adVarChar, 50
'   lors.Fields.Append "sBarrCode", adVarChar, 25
'   lors.Open
'
'   With GridEditor1
'      For lnCtr = 1 To .Rows - 1
'         lors.AddNew
'         lors("nQuantity").Value = .TextMatrix(lnCtr, 5)
'         lors("sModel").Value = .TextMatrix(lnCtr, 3)
'         lors("sColor").Value = .TextMatrix(lnCtr, 4)
'         lors("sDescription").Value = .TextMatrix(lnCtr, 2)
'         lors("sBarrCode").Value = .TextMatrix(lnCtr, 1)
'      Next
'   End With
'
'   ' assign important info to the report
'   Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\Purchase.rpt")
'   oReport.DiscardSavedData
'   oReport.FieldMappingType = crAutoFieldMapping
'   oReport.Database.SetDataSource lors
'
'   Set lrs = New ADODB.Recordset
'   lrs.Open "Select" _
'               & "  CONCAT(b.sTownName, ', ', c.sProvName, ' ', b.sZippCode) xAddressx" _
'            & " From Branch a" _
'               & ", TownCity b" _
'                  & " Left Join Province c" _
'                     & " On b.sProvIDxx = c.sProvIDxx" _
'            & " Where a.sTownIDxx = b.sTownIDxx" _
'               & " And a.sBranchCd = " & strParm(oTrans.Master("sBranchCd")) _
'            , oApp.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'   If Not lrs.EOF Then oReport.Sections("PH").ReportObjects("txtDeliver").SetText "           " & txtField(4).Text & vbCrLf & " " & lrs("xAddressx")
'   oReport.Sections("PH").ReportObjects("txtSupplier").SetText txtField(2).Text
'   oReport.Sections("PH").ReportObjects("txtTerm").SetText txtField(6).Text
'   oReport.Sections("PH").ReportObjects("txtDDate").SetText txtField(7).Text
'   oReport.Sections("PH").ReportObjects("txtDate").SetText txtField(1).Text
'   oReport.Sections("PF").ReportObjects("txtUserRpt").SetText oApp.UserName
'
'
'         Set loreport = New frmRepViewer
'         Set loreport.ReportSource = oReport
'         loreport.Show
'
'
''185      oReport.PrintOutEx False, 1
'   lors.Close
'   lrs.Close
'
'endProc:
'   oTrans.CloseTransaction (oTrans.Master(0))
'   Set oReport = Nothing
'   Set lors = Nothing
'   Set lrs = Nothing
'   Exit Function
'errProc:
'   PrintTrans = False
'   ShowError lsOldProc & "( " & " )"
'End Function
'
'Private Sub txtField_LostFocus(Index As Integer)
'   With txtField(Index)
'      .BackColor = oApp.getColor("EB")
'   End With
'End Sub
'
'Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
'   Dim lsOldProc As String
'
'   lsOldProc = "txtField_Validate"
'   'On Error GoTo errProc
'
'   With txtField(Index)
'      .Text = TitleCase(.Text)
'      Select Case Index
'      Case 9, 10
'         If .Text = "" Then
'            ClearFields
'            Exit Sub
'         End If
'
'         If .Text <> .Tag Then
'            If oTrans.SearchTransaction _
'            (IIf(Index = 9, CodeFormat(oApp.BranchCode, .Text), .Text) _
'            , IIf(Index = 9, True, False)) Then
'               LoadMaster
'               LoadDetail
'            Else
'               ClearFields
'               .SetFocus
'            End If
'         End If
'      End Select
'   End With
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " _
'                       & "  " & Index _
'                       & ", " & Cancel _
'                       & " )", True
'End Sub
'
'Private Sub ShowError(ByVal lsProcName As String, Optional bEnd As Boolean = False)
'   With oApp
'      .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
'      If bEnd Then
'         .xShowError
'         End
'      Else
'         With Err
'            .Raise .Number, .Source, .Description
'         End With
'      End If
'   End With
'End Sub
'Private Sub InitGrid()
'   With GridEditor1
'      .Row = 2
'      .Cols = 5
'      .Font = "MS Sans Serif"
'      'Cols Title
'      .TextMatrix(0, 1) = "Entry No"
'      .TextMatrix(0, 2) = "Description"
'      .TextMatrix(0, 3) = "Quantity"
'      .TextMatrix(0, 4) = "Qty on Hnd"
'      .TextMatrix(0, 5) = "Act on Hnd"
'
'      .Row = 0
'
'      'Col Alignment
'      For pnCtr = 0 To .Cols - 1
'         .Col = pnCtr
'         .CellFontBold = True
'         .CellAlignment = 3
'      Next
'      'Cols Width
'      .ColWidth(0) = 330
'      .ColWidth(1) = 1000
'      .ColWidth(2) = 1000
'      .ColWidth(3) = 1000
'      .ColWidth(4) = 1000
'      .ColWidth(5) = 1000
'
'      .ColFormat(3) = "#,##0"
'      .ColFormat(4) = "#,##0.00"
'      .ColFormat(5) = "#,##0.00"
'      .ColDefault(3) = 0
'      .ColDefault(4) = 0
'      .ColDefault(5) = 0
'
'      .ColAlignment(1) = 1
'      .ColAlignment(2) = 1
'      .ColAlignment(3) = 1
'      .ColAlignment(5) = 1
'      .EditorBackColor = oApp.getColor("HT1")
'      .Row = 1
'      .Col = 1
'      End With
'End Sub
'
