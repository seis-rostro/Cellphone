VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0"; "xrGridControl.ocx"
Begin VB.Form frmCP_BackLoad 
   BorderStyle     =   0  'None
   Caption         =   "Cellphone Backload Transfer"
   ClientHeight    =   7800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11715
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   11715
   ShowInTaskbar   =   0   'False
   Begin xrGridEditor.GridEditor GridEditor1 
      Height          =   4815
      Left            =   1575
      TabIndex        =   12
      Tag             =   "et0;eb0;et0;bc2"
      Top             =   2850
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   8493
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
      Object.HEIGHT          =   4815
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
      MOUSEICON       =   "frmCP_BackLoad.frx":0000
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
      Height          =   1320
      Index           =   1
      Left            =   1575
      Tag             =   "wt0;fb0"
      Top             =   1500
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   2328
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   1365
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   150
         Width           =   2310
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   645
         Index           =   4
         Left            =   1365
         MaxLength       =   128
         MultiLine       =   -1  'True
         TabIndex        =   11
         Text            =   "frmCP_BackLoad.frx":001C
         Top             =   480
         Width           =   8475
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   3
         Left            =   4770
         MaxLength       =   8
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   150
         Width           =   1575
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   4980
         MaxLength       =   50
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   795
         Width           =   1665
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transact. Date"
         Height          =   285
         Index           =   1
         Left            =   165
         TabIndex        =   4
         Top             =   180
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transmittal No"
         Height          =   285
         Index           =   3
         Left            =   3690
         TabIndex        =   8
         Top             =   180
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   300
         Index           =   7
         Left            =   570
         TabIndex        =   10
         Top             =   510
         Width           =   660
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Destination"
         Height          =   285
         Index           =   4
         Left            =   4065
         TabIndex        =   6
         Top             =   945
         Width           =   855
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   900
      Index           =   0
      Left            =   1575
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   10020
      _ExtentX        =   17674
      _ExtentY        =   1588
      BackColor       =   12632256
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
         Index           =   6
         Left            =   1365
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   435
         Width           =   8460
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
         Left            =   7830
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   105
         Width           =   1995
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
         Index           =   5
         Left            =   1365
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   105
         Width           =   4965
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
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
         Left            =   165
         TabIndex        =   22
         Top             =   480
         Width           =   1305
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Branch Origin"
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
         Index           =   19
         Left            =   165
         TabIndex        =   0
         Top             =   150
         Width           =   1305
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction No"
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
         Left            =   6435
         TabIndex        =   2
         Top             =   150
         Width           =   1440
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   90
      TabIndex        =   13
      Top             =   4950
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
      Picture         =   "frmCP_BackLoad.frx":0032
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   90
      TabIndex        =   14
      Top             =   3690
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
      Picture         =   "frmCP_BackLoad.frx":07AC
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   90
      TabIndex        =   15
      Top             =   4320
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
      Picture         =   "frmCP_BackLoad.frx":0F26
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   90
      TabIndex        =   16
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
      Picture         =   "frmCP_BackLoad.frx":16A0
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   90
      TabIndex        =   17
      Top             =   3690
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
      Picture         =   "frmCP_BackLoad.frx":1E1A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   6
      Left            =   90
      TabIndex        =   18
      Top             =   5580
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
      Picture         =   "frmCP_BackLoad.frx":2594
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   90
      TabIndex        =   19
      Top             =   4950
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
      Picture         =   "frmCP_BackLoad.frx":2D0E
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   7
      Left            =   90
      TabIndex        =   20
      Top             =   4320
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Save To"
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
      Picture         =   "frmCP_BackLoad.frx":3488
   End
End
Attribute VB_Name = "frmCP_BackLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Private Const pxeMODULENAME = "frmCP_Branch2Branch_Transfer"
'
'Private WithEvents oTrans As clsCPTransfer
'Private oSkin As clsFormSkin
'Private oBranch As clsBranch
'
'Dim pnIndex As Integer
'Dim pbGridFocus As Boolean
'Dim pnCtr As Integer
'Dim pbSave As Boolean
'Dim pbGridValidate As Boolean
'
'Private Sub chkField_Click()
''   oTrans.DiskTransaction = IIf(chkField.Value = 1, True, False)
'End Sub
'
'Private Sub cmdButton_Click(Index As Integer)
'   Dim lsOldProc As String
'   Dim lnRep As String
'
'   lsOldProc = "cmdButton_Click"
'   'On Error GoTo errProc
'
'   With GridEditor1
'      Select Case Index
'      Case 0
'         If .Rows > 2 Then
'            pnCtr = 0
'            Do While pnCtr < .Rows
'               If Trim(.TextMatrix(pnCtr, 1)) = "" Then
'                  .Row = pnCtr
'                  If oTrans.deleteDetail(.Row - 1) Then .DeleteRow
'               Else
'                  pnCtr = pnCtr + 1
'               End If
'            Loop
'
'            .ColWidth(3) = 3150
'            If .Rows > 18 Then .ColWidth(3) = 2950
'         End If
'
'         If isEntryOK Then
'            If oTrans.SaveTransaction Then
'               MsgBox "Record Updated Successfully!!!", vbInformation, "Confirm"
'               lnRep = MsgBox("Post Transaction!!!", vbYesNo + vbQuestion, "Confirm")
'
'               If lnRep = vbYes Then
'                  If Not oTrans.AcceptDelivery(oTrans.Master("dTransact")) Then MsgBox "Unable to Post Transaction!!!", vbCritical, "Warning"
'               End If
'               InitButton xeModeReady
'               lnRep = MsgBox("Print Transaction!!!", vbYesNo + vbQuestion, "Confirm")
'               If lnRep = vbYes Then
'                  If Not PrintTrans Then MsgBox "Unable to Print Transaction!!!", vbCritical, "Warning"
'               End If
'               pbSave = True
'            Else
'               MsgBox "Unable to Update Record!!!", vbCritical, "Warning"
'            End If
'         End If
'      Case 1
'         If pbGridFocus Then
'            If Trim(txtField(5).Text) = "" Then Exit Sub
'            If oTrans.searchDetail(.Row - 1, 1) Then .Col = 1
'            .Refresh
'            .SetFocus
'         Else
'            If pnIndex = 5 Then
'               If oBranch.SearchRecord("", False) Then
'                  oTrans.Branch = oBranch.Master("sBranchCd")
'                  oTrans.InitTransaction
'                  oTrans.NewTransaction
'                  ClearFields
'
'                  txtField(pnIndex).Text = oBranch.Master("sBranchNm")
'                  txtField(6).Text = oBranch.Master("sAddressx")
'                  txtField(0).Text = Format(oTrans.Master("sTransNox"), "@@@@-@@@@@@")
'                  txtField(1).SetFocus
'               Else
'                  If Trim(.Tag) <> "" Then
'                     .Text = .Tag
'                     Exit Sub
'                  End If
'
'                  ClearFields
'                  .SetFocus
'               End If
'            Else
'               oTrans.SearchMaster pnIndex
'            End If
'         End If
'      Case 2
'         If .Rows > 2 Then
'            If oTrans.deleteDetail(.Row - 1) Then .DeleteRow
'
'            For pnCtr = 1 To .Rows - 1
'               .TextMatrix(pnCtr, 0) = pnCtr
'            Next
'
'            .ColWidth(3) = 3150
'            If .Rows > 18 Then .ColWidth(3) = 2950
'         End If
'      Case 3
'         lnRep = MsgBox("Transaction is in Update Mode!!!" & vbCrLf & _
'                        "Do you want to Cancel Transaction!!!", vbYesNo + vbQuestion, "Confirm")
'
'         If lnRep = vbYes Then
'            oTrans.NewTransaction
'            ClearFields
'            InitButton xeModeReady
'         Else
'            txtField(pnIndex).SetFocus
'         End If
'         pbSave = False
'      Case 4
'         oTrans.NewTransaction
'         ClearFields
'         InitButton xeModeAddNew
'         txtField(5).SetFocus
'      Case 5
'         If pbSave Then
'            lnRep = MsgBox("Print Transaction!!!", vbYesNo + vbQuestion, "Confirm")
'            If lnRep = vbYes Then
'               If Not PrintTrans Then MsgBox "Unable to Print Transaction!!!", vbCritical, "Warning"
'            End If
'         Else
'            MsgBox "Unable to Print Transaction!!!", vbCritical, "Warning"
'         End If
'      Case 6
'         Unload Me
'      Case 7
''         If chkField.Value = 1 Then
''            oTrans.DiskTransaction = True
''            If oTrans.CreateDiskTransfer Then
''               MsgBox "Transaction was Successfully Save to Mobile Disk!!!", vbInformation, "Notice"
''            Else
''               MsgBox "Unable to Save Transaction to Mobile Disk!!!", vbCritical, "Warning"
''            End If
''         Else
''            MsgBox "MC Delivery Capture Was Not Yet Set!!!" & vbCrLf & _
''               "Please Checked 'Save to Mobile Disk' then Try Exporting Delivery Again!!!", _
''               vbCritical, "Warning"
''         End If
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
'   GridEditor1.Refresh
'
'   oApp.MenuName = Me.Tag
'   Me.ZOrder 0
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
'   Set oTrans = New clsCPTransfer
'   Set oTrans.AppDriver = oApp
'
''   oTrans.DiskTransaction = False
'   oTrans.InitTransaction
'   oTrans.NewTransaction
'
'   Set oSkin = New clsFormSkin
'   Set oSkin.AppDriver = oApp
'   Set oSkin.Form = Me
'   oSkin.ApplySkin xeFormTransaction
'
'   Set oBranch = New clsBranch
'   Set oBranch.AppDriver = oApp
'   oBranch.Filter = "sBranchCd <> " & strParm(oApp.BranchCode) _
'                     & " AND cAutomate = " & strParm(xeNo)
'   oBranch.InitRecord
'   oBranch.NewRecord
'
'   InitGrid
'   ClearFields
'   InitButton xeModeAddNew
'
'   txtField(3).MaxLength = oTrans.MasFldSize(3)
'   txtField(4).MaxLength = oTrans.MasFldSize(4)
'
'   pbGridValidate = False
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
'Private Sub GridEditor1_AddingRow(Cancel As Boolean)
'   With GridEditor1
'      If .TextMatrix(.Row, 1) = "" Then
'         Cancel = True
'      ElseIf .TextMatrix(.Row, 6) = "0" Then
'         Cancel = True
'      End If
'      If Not Cancel Then
'         If .Row = .Rows - 1 Then oTrans.addDetail
'      End If
'
'      If .Rows > 20 Then .ColWidth(3) = 2950
'   End With
'End Sub
'
'Private Sub GridEditor1_EditorValidate(Cancel As Boolean)
'   Dim lsOldProc As String
'
'   lsOldProc = "GridEditor1_EditorValidate"
'   'On Error GoTo errProc
'
'   With GridEditor1
'      If Trim(txtField(5).Text) = "" Then Exit Sub
'      If pbGridValidate Then
'         pbGridValidate = False
'         Exit Sub
'      End If
'
'      If .Col = 1 Or .Col = 2 Then
'         .TextMatrix(.Row, .Col) = compareSerial(.TextMatrix(.Row, .Col), .Row)
'      End If
'      oTrans.Detail(.Row - 1, .Col) = .TextMatrix(.Row, .Col)
'
'      If oTrans.Detail(.Row - 1, "cHsSerial") = xeYes Then
'         Select Case .Col
'         Case 1
'            If .TextMatrix(.Row, .Col) <> "" Then
'               oTrans.Detail(.Row - 1, "nQuantity") = 1
'               .TextMatrix(.Row, 6) = oTrans.Detail(.Row - 1, "nQuantity")
'               If .Row = .Rows - 1 Then
'                  .Rows = .Rows + 1
'                  oTrans.addDetail
'               End If
'
'               .Row = .Rows - 1
'               .Col = 0
'            End If
'         Case 6
'            If CDbl(.TextMatrix(.Row, .Col)) > 1 Then .TextMatrix(.Row, .Col) = 1
'         End Select
'      End If
'   End With
'   pbGridValidate = True
'
'endProc:
'   GridEditor1.Refresh
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " & Cancel & " )", True
'End Sub
'
'Private Sub GridEditor1_GotFocus()
'   With GridEditor1
'      .EditorBackColor = oApp.getColor("HT1")
'   End With
'   pbGridFocus = True
'End Sub
'
'Private Sub GridEditor1_KeyDown(KeyCode As Integer, Shift As Integer)
'   Dim lsOldProc As String
'
'   lsOldProc = "GridEditor1_KeyDown"
'   'On Error GoTo errProc
'
'   If KeyCode = vbKeyF3 Then
'      If Trim(txtField(5).Text) = "" Then
'         KeyCode = 0
'         Exit Sub
'      End If
'
'      With GridEditor1
'         If .Col = 1 Or .Col = 2 Then
'            If oTrans.searchDetail(.Row - 1, .Col, .TextMatrix(.Row, .Col)) Then
'               If oTrans.Detail(.Row - 1, "cHsSerial") = xeYes Then
'                  .TextMatrix(.Row, 6) = 1
'                  oTrans.Detail(.Row - 1, "nQuantity") = 1
'                  If .Row = .Rows - 1 Then
'                     .Rows = .Rows + 1
'                     oTrans.addDetail
'                  End If
'
'                  .Row = .Rows - 1
'                  .Col = 1
'               Else
'                  .Col = 6
'               End If
'            Else
'               oTrans.Detail(.Row - 1, .Col) = .TextMatrix(.Row, .Col)
'               .Col = 1
'            End If
'         End If
'
'         .Refresh
'         .SetFocus
'         KeyCode = 0
'      End With
'   End If
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " _
'                       & "  " & KeyCode _
'                       & ", " & Shift _
'                       & " )", True
'End Sub
'
'Private Sub GridEditor1_LostFocus()
'   With GridEditor1
'      .EditorBackColor = oApp.getColor("EB")
'      oTrans.Detail(.Row - 1, .Col) = .TextMatrix(.Row, .Col)
'   End With
'
'   pbGridValidate = False
'End Sub
'
'Private Sub oTrans_DetailRetrieved(ByVal Index As Integer)
'   With GridEditor1
'      .TextMatrix(.Row, Index) = oTrans.Detail(.Row - 1, Index)
'   End With
'End Sub
'
'Private Sub oTrans_MasterRetrieved(ByVal Index As Integer)
'   txtField(Index).Text = oTrans.Master(Index)
'End Sub
'
'Private Sub InitGrid()
'   With GridEditor1
'      .Rows = 2
'      .Cols = 7
'      .Font = "MS Sans Serif"
'
'      'column title
'      .TextMatrix(0, 1) = "BarrCode"
'      .TextMatrix(0, 2) = "Description"
'      .TextMatrix(0, 3) = "Model"
'      .TextMatrix(0, 6) = "Qty"
'
'      .Row = 0
'
'      'column alignment
'      For pnCtr = 0 To .Cols - 1
'         .Col = pnCtr
'         .CellFontBold = True
'         .CellAlignment = 3
'      Next
'
'      'column width
'      .ColWidth(0) = 330
'      .ColWidth(1) = 2670
'      .ColWidth(2) = 3000
'      .ColWidth(4) = 0
'      .ColWidth(5) = 0
'      .ColWidth(6) = 800
'
'      .ColFormat(6) = "#,##0"
'      .ColNumberOnly(6) = True
'      .ColDefault(6) = 0
'
'      .ColAlignment(1) = 1
'      .ColAlignment(2) = 1
'      .ColAlignment(3) = 1
'      .ColAlignment(6) = 6
'
'      .ColEnabled(3) = False
'      .ColEnabled(4) = False
'      .ColEnabled(5) = False
'
'      .EditorBackColor = oApp.getColor("HT1")
'
'      .Row = 1
'      .Col = 1
'   End With
'End Sub
'
'Private Sub txtField_GotFocus(Index As Integer)
'   With txtField(Index)
'      If Index = 1 Then .Text = Format(.Text, "MM/DD/YY")
'
'      .SelStart = 0
'      .SelLength = Len(.Text)
'      .BackColor = oApp.getColor("HT1")
'   End With
'
'   pbGridFocus = False
'   pnIndex = Index
'End Sub
'
'Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'   Dim lsOldProc As String
'
'   lsOldProc = "txtField_KeyDown"
'   'On Error GoTo errProc
'
'   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
'      With txtField(Index)
'         Select Case Index
'         Case 5
'            Call txtField_Validate(Index, False)
'         Case Else
'            If KeyCode = vbKeyF3 Then
'               oTrans.SearchMaster Index, .Text
'               If .Text <> "" Then SetNextFocus
'            Else
'               If .Text <> "" Then oTrans.SearchMaster Index, .Text
'            End If
'         End Select
'      End With
'      KeyCode = 0
'   End If
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " _
'                       & "  " & Index _
'                       & ", " & KeyCode _
'                       & ", " & Shift _
'                       & " )", True
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
'Private Sub InitButton(lnStat As Integer)
'   Dim lbShow As Boolean
'
'   lbShow = IIf(lnStat = 0, False, True)
'   cmdButton(4).Visible = Not lbShow
'   cmdButton(5).Visible = Not lbShow
'   cmdButton(6).Visible = Not lbShow
'   cmdButton(7).Visible = Not lbShow
'
'   cmdButton(0).Visible = lbShow
'   cmdButton(1).Visible = lbShow
'   cmdButton(2).Visible = lbShow
'   cmdButton(3).Visible = lbShow
'
'   For pnCtr = 1 To txtField.Count - 1
'      If pnCtr <> 2 Then txtField(pnCtr).Enabled = lbShow
'   Next
'
'   With GridEditor1
'      .ColEnabled(1) = lbShow
'      .ColEnabled(2) = lbShow
'      .ColEnabled(6) = lbShow
'   End With
'
'   xrFrame1(0).Enabled = lbShow
'   xrFrame1(1).Enabled = lbShow
'
'   If Not lbShow Then cmdButton(4).SetFocus
'End Sub
'
'Private Function PrintTrans() As Boolean
'   Dim lrs As Recordset
'   Dim lors As Recordset
'   Dim lnCtr As Integer
'   Dim lsOldProc As String
'
'   lsOldProc = "PrinTrans"
'   'On Error GoTo errProc
'
'   PrintTrans = True
'
'   Set lrs = New ADODB.Recordset
'
'   lrs.Fields.Append "nField01", adInteger, 3
'   lrs.Fields.Append "sField01", adVarChar, 20
'   lrs.Fields.Append "sField02", adVarChar, 128
'   lrs.Fields.Append "sField03", adVarChar, 25
'   lrs.Fields.Append "sField04", adVarChar, 12
'   lrs.Fields.Append "sField05", adVarChar, 20
'   lrs.Open
'
'   With oTrans
'      For lnCtr = 0 To .ItemCount - 1
'        lrs.Fields("nField01") = oTrans.Detail(lnCtr, "nQuantity")
'        lrs.Fields("sField01") = oTrans.Detail(lnCtr, "sBarrCode")
'        lrs.Fields("sField02") = oTrans.Detail(lnCtr, "sDescript")
'        lrs.Fields("sField03") = oTrans.Detail(lnCtr, "sBrandNme")
'        lrs.Fields("sField04") = oTrans.Detail(lnCtr, "sStockIDx")
'        lrs.Fields("sField05") = oTrans.Detail(lnCtr, "sSerialNo")
'      Next
'   End With
'
'   'assign important info to the report
'   Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\CP_Transfer.rpt")
'   oReport.DiscardSavedData
'   oReport.FieldMappingType = crAutoFieldMapping
'   oReport.Database.SetDataSource lrs
'
'   Set lors = New ADODB.Recordset
'   If lors.State = adStateOpen Then lors.Close
'
'   lors.Open "SELECT" _
'               & "  CONCAT(a.sAddressx, ', ', b.sTownName, ', ', c.sProvName, ' ', b.sZippCode) as Address" _
'               & ", d.sCompnyNm" _
'            & " From Branch a" _
'               & ", TownCity b" _
'               & ", Province c" _
'               & ", Company d" _
'            & " WHERE a.sBranchCd = " & strParm(oTrans.Master("sDestinat")) _
'               & " AND a.sTownIDxx = b.sTownIDxx" _
'               & " AND b.sProvIDxx = c.sProvIDxx" _
'               & " AND a.sCompnyID = d.sCompnyID" _
'            , oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
'
'   oReport.Sections("RH").ReportObjects("txtRefNo").SetText "CP" & "-" & Right(oTrans.Master("sTransNox"), 10)
'   oReport.Sections("RH").ReportObjects("txtDate").SetText txtField(1).Text
'   oReport.Sections("PH").ReportObjects("txtTo").SetText lors("sCompnyNm")
'   oReport.Sections("PH").ReportObjects("txtToAddress").SetText lors("Address")
'   oReport.Sections("PH").ReportObjects("txtFrom").SetText oApp.ClientName
'   oReport.Sections("PH").ReportObjects("txtFromAddress").SetText oApp.Address & ", " & oApp.TownCity & ", " & oApp.Province & " " & oApp.ZippCode
'   oReport.Sections("PF").ReportObjects("txtPrepared").SetText oApp.UserName
'   oReport.Sections("RF").ReportObjects("txtNote").SetText txtField(4).Text
'
'   oReport.PrintOutEx False, 1
'   lors.Close
'
'   PrintTrans = True
'
'endPoc:
'   oTrans.CloseTransaction (oTrans.Master(0))
'   Set oReport = Nothing
'   Set lrs = Nothing
'   Set lors = Nothing
'   Exit Function
'errProc:
'   PrintTrans = False
'   ShowError lsOldProc & "( " & " )"
'End Function
'
'
'Private Sub ClearFields()
'   For pnCtr = 0 To txtField.Count - 1
'      Select Case pnCtr
'      Case 0
'         txtField(pnCtr).Text = ""
'      Case 1
'         txtField(pnCtr).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
'      Case Else
'         txtField(pnCtr).Text = ""
'         txtField(pnCtr).Tag = ""
'      End Select
'   Next
'
'   With GridEditor1
'      .Rows = 2
'      .ColWidth(3) = 3150
'
'      'empty row
'      .TextMatrix(1, 1) = ""
'      .TextMatrix(1, 2) = ""
'      .TextMatrix(1, 3) = ""
'      .TextMatrix(1, 6) = "0"
'   End With
'
'   pbSave = False
'End Sub
'
'Private Sub txtField_LostFocus(Index As Integer)
'   With txtField(Index)
'      .BackColor = oApp.getColor("EB")
'   End With
'End Sub
'
'Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
'   With txtField(Index)
'      .Text = TitleCase(.Text)
'      Select Case Index
'      Case 1
'         If Not IsDate(.Text) Then .Text = oApp.ServerDate
'         .Text = Format(.Text, "MMMM DD, YYYY")
'      Case 3
'         .Text = Format(.Text, ">")
'      Case 5
'         If Trim(.Text) = "" Then
'            ClearFields
'            Exit Sub
'         End If
'
'         If Trim(LCase(.Text)) <> Trim(LCase(.Tag)) Then
'            If oBranch.SearchRecord(.Text, False) Then
'               oTrans.Branch = oBranch.Master("sBranchCd")
'               oTrans.InitTransaction
'               oTrans.NewTransaction
'               ClearFields
'
'               .Text = oBranch.Master("sBranchNm")
'               txtField(6).Text = oBranch.Master("sAddressx")
'               txtField(0).Text = Format(oTrans.Master("sTransNox"), "@@@@-@@@@@@")
'            Else
'               If Trim(.Tag) <> "" Then
'                  .Text = .Tag
'                  Exit Sub
'               End If
'
'               ClearFields
'               .SetFocus
'            End If
'         End If
'
'         .Tag = .Text
'      End Select
'
'      If Index < 5 Then oTrans.Master(Index) = .Text
'   End With
'End Sub
'
'Private Function isEntryOK() As Boolean
'   If txtField(5).Text = "" Then
'      MsgBox "Branch not found!!!" & vbCrLf & _
'             "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
'      txtField(5).SetFocus
'      GoTo EntryNotOK
'   End If
'
'   If txtField(0).Text = "" Then
'      MsgBox "Transaction not found!!!" & vbCrLf & _
'             "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
'      txtField(5).SetFocus
'      GoTo EntryNotOK
'   End If
'
'   If Left(oTrans.Master("sTransNox"), Len(oApp.BranchCode)) = oApp.BranchCode Then
'      MsgBox "Unable to create transaction!!!" & vbCrLf & _
'             "Invalid source/origin branch!!!", vbCritical, "Warning"
'      txtField(5).SetFocus
'      GoTo EntryNotOK
'   End If
'
'   With GridEditor1
'      If Trim(.TextMatrix(1, 1)) = "" Then
'         MsgBox "Detail is required!!!" & vbCrLf & _
'                "Please Verify your Entry then Try Again", vbCritical, "Warning"
'         .SetFocus
'         .Row = 1
'         .Col = 1
'         GoTo EntryNotOK
'      End If
'   End With
'
'   oTrans.Master("sDestinat") = oApp.BranchCode
'
'EntryOK:
'   isEntryOK = True
'   Exit Function
'EntryNotOK:
'   isEntryOK = False
'End Function
'
'Private Function compareSerial(Value As String, Row As Integer) As String
'   Dim lnRep As Integer
'   Dim lnCtr As Integer
'   Dim lsValue As Integer
'   Dim lnValue As Integer
'
'   If Trim(Value) = "" Then
'      compareSerial = ""
'      Exit Function
'   End If
'
'   With GridEditor1
'      For lnCtr = 1 To .Rows - 1
'         If .TextMatrix(lnCtr, 1) = Value And lnCtr <> Row Then
'            If oTrans.Detail(lnCtr - 1, "cHsSerial") = xeYes Then
'               MsgBox "Duplicate Serial No!!!" & vbCrLf & _
'                        "Please Verify your entry then try again!!!", vbCritical, "Warning"
'            Else
'               lnRep = MsgBox("Duplicate Serial No!!!" & vbCrLf & _
'                                 "Item automatically add from existing serial!!!", vbInformation, "CONFIRM")
'               If lnRep = vbYes Then
'                  lsValue = InputBox("Please specify quantity for serial " & Value, "Quantity", 0)
'                  lnValue = IIf(lsValue = "", 0, lsValue)
'
'                  .TextMatrix(lnCtr, 6) = .TextMatrix(lnCtr, 6) + lnValue
'                  oTrans.Detail(lnCtr - 1, "nQuantity") = CDbl(.TextMatrix(lnCtr, 6))
'               End If
'            End If
'            compareSerial = ""
'         Else
'            compareSerial = Value
'         End If
'      Next
'   End With
'End Function
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
