VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0"; "xrGridControl.ocx"
Begin VB.Form frmCP_SRP 
   BorderStyle     =   0  'None
   Caption         =   "CP SRP"
   ClientHeight    =   6960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11415
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6960
   ScaleWidth      =   11415
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   2235
      Left            =   1575
      Tag             =   "wt0;fb0"
      Top             =   525
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   3942
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   1320
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   990
         Width           =   4380
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   4
         Left            =   6660
         MaxLength       =   8
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   660
         Width           =   2940
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   720
         Index           =   3
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   7
         Text            =   "frmCP_SRP.frx":0000
         Top             =   1320
         Width           =   8295
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   1320
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   660
         Width           =   2265
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
         Left            =   1320
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   150
         Width           =   2265
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   4
         Top             =   1050
         Width           =   630
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Approved"
         Height          =   195
         Index           =   8
         Left            =   5835
         TabIndex        =   8
         Top             =   720
         Width           =   810
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   195
         Index           =   5
         Left            =   90
         TabIndex        =   6
         Top             =   1395
         Width           =   630
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transact. Date"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   2
         Top             =   720
         Width           =   1065
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transact. No."
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
         Left            =   90
         TabIndex        =   0
         Top             =   210
         Width           =   1185
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   315
         Left            =   1410
         Tag             =   "et0;ht2"
         Top             =   255
         Width           =   2265
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   90
      TabIndex        =   17
      Top             =   4125
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
      Picture         =   "frmCP_SRP.frx":0006
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   90
      TabIndex        =   14
      Top             =   4125
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
      Picture         =   "frmCP_SRP.frx":0780
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   630
      Index           =   0
      Left            =   90
      TabIndex        =   11
      Top             =   2175
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   1111
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
      Picture         =   "frmCP_SRP.frx":0EFA
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   90
      TabIndex        =   13
      Top             =   3495
      Width           =   1260
      _ExtentX        =   2223
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
      Picture         =   "frmCP_SRP.frx":1674
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   90
      TabIndex        =   15
      Top             =   4755
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
      Picture         =   "frmCP_SRP.frx":1DEE
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   90
      TabIndex        =   16
      Top             =   3495
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
      Picture         =   "frmCP_SRP.frx":2568
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   6
      Left            =   90
      TabIndex        =   18
      Top             =   4755
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
      Picture         =   "frmCP_SRP.frx":2CE2
   End
   Begin xrGridEditor.GridEditor GridEditor1 
      Height          =   4020
      Left            =   1590
      TabIndex        =   10
      Tag             =   "et0;eb0;et0;bc2"
      Top             =   2805
      Width           =   9705
      _ExtentX        =   17119
      _ExtentY        =   7091
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
      Object.HEIGHT          =   4020
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
      MOUSEICON       =   "frmCP_SRP.frx":345C
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
   Begin xrControl.xrButton cmdButton 
      Height          =   630
      Index           =   7
      Left            =   75
      TabIndex        =   12
      Top             =   2835
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   1111
      Caption         =   "&Load"
      AccessKey       =   "L"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCP_SRP.frx":3478
   End
End
Attribute VB_Name = "frmCP_SRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmCP_SRP"

Private WithEvents oTrans As clsCPSRP
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin

Dim pnCtr As Integer
Dim pnIndex As Integer
Dim pbGridFocus As Boolean
Dim pbSave As Boolean

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   Dim lsRep As String

   lsOldProc = "cmdButton_Click"
   'On Error GoTo errProc
   
   With GridEditor1
      Select Case Index
      Case 0 'Save
         If .Rows > 2 Then
            pnCtr = 0
            Do While pnCtr < .Rows
               If Trim(.TextMatrix(pnCtr, 9)) = "x" Then
                  .Row = pnCtr
                  If oTrans.deleteDetail(.Row - 1) Then .deleteRow
               Else
                  pnCtr = pnCtr + 1
               End If
            Loop
         
            .ColWidth(2) = 2450
            .ColWidth(3) = 2600
            If .Rows > 16 Then
               .ColWidth(2) = 2350
               .ColWidth(3) = 2500
            End If
         End If
            
         If isEntryOk Then
            If oTrans.SaveTransaction = True Then
               MsgBox "Transaction Saved Successfully!!!", vbInformation, "Notice"
               initButton xeModeReady
               
'               lsRep = MsgBox("Print Transaction!!!", vbYesNo + vbQuestion, "Confirm")
'               If lsRep = vbYes Then
'                  If Not PrintTrans Then MsgBox "Unable to Print Transaction!!!", vbCritical, "Warning"
'               End If
'               pbSave = True
            Else
               MsgBox "Unable to Save Transaction!!!", vbCritical, "Warning"
            End If
         End If
      Case 1 'Search
         If Not pbGridFocus Then
            oTrans.SearchMaster pnIndex
         End If
      Case 2 'Delete row
         If .Rows > 2 Then
            If oTrans.deleteDetail(.Row - 1) Then .deleteRow
            
            .ColWidth(3) = 3600
            If .Rows > 20 Then .ColWidth(3) = 3400
         End If
      Case 3 'Cancel
         lsRep = MsgBox("Transaction is in Update Mode!!!" & vbCrLf & _
                        "Do you want to Cancel Transaction!!!", vbYesNo + vbQuestion, "Confirm")
         
         If lsRep = vbYes Then
            oTrans.NewTransaction
            ClearFields
            initButton xeModeReady
         Else
            txtField(pnIndex).SetFocus
         End If
         pbSave = False
      Case 4 'New
         oTrans.NewTransaction
         ClearFields
         initButton xeModeAddNew
         txtField(2).SetFocus
      Case 5 'Print
         If pbSave Then
            lsRep = MsgBox("Print Transaction!!!", vbYesNo + vbQuestion, "Confirm")
            If lsRep = vbYes Then
               If Not PrintTrans Then MsgBox "Unable to Print Transaction!!!", vbCritical, "Warning"
            End If
         Else
            MsgBox "Unable to Print Transaction!!!", vbCritical, "Warning"
         End If
      Case 6 '
         Unload Me
      Case 7 ' Load Detail
         If oTrans.Master("sCategrID") <> "" Then
            If oTrans.getBarrCode Then LoadDetail
            .SetFocus
            .Row = 1
            .Col = 8
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
   'On Error GoTo errProc
   
   CenterChildForm mdiMain, Me

   Set oTrans = New clsCPSRP
   Set oTrans.AppDriver = oApp

   oTrans.InitTransaction
   oTrans.NewTransaction
   
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransaction
   
   InitGrid
   ClearFields
   initButton xeModeAddNew

   txtField(4).MaxLength = oTrans.MasFldSize(4)
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oTrans = Nothing
   Set oSkin = Nothing
End Sub

Private Sub GridEditor1_AddingRow(Cancel As Boolean)
   With GridEditor1
      If .TextMatrix(.Row, 1) = "" Then
         Cancel = True
      ElseIf .TextMatrix(.Row, 6) = 0 Then
         Cancel = True
      End If
      If Not Cancel Then oTrans.addDetail
 
      If .Rows > 16 Then .ColWidth(3) = 3400
   End With
End Sub

Private Sub GridEditor1_EditorValidate(Cancel As Boolean)
   With GridEditor1
      oTrans.Detail(.Row - 1, "nNewSRPxx") = CDbl(.TextMatrix(.Row, .Col))
      If CDbl(.TextMatrix(.Row, 7)) <> CDbl(.TextMatrix(.Row, 8)) Then .TextMatrix(.Row, 9) = ""
   End With
End Sub

Private Sub GridEditor1_GotFocus()
   With GridEditor1
      .EditorBackColor = oApp.getColor("HT1")
   End With
   pbGridFocus = True
End Sub

Private Sub GridEditor1_LostFocus()
   With GridEditor1
      .EditorBackColor = oApp.getColor("EB")
      oTrans.Detail(.Row - 1, .Col) = .TextMatrix(.Row, .Col)
   End With
End Sub

Private Sub oTrans_DetailRetrieved(ByVal Index As Integer)
   With GridEditor1
      .TextMatrix(.Row, Index) = oTrans.Detail(.Row - 1, Index)
   End With
End Sub

Private Sub oTrans_MasterRetrieved(ByVal Index As Integer)
   txtField(Index).Text = oTrans.Master(Index)
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      If Index = 1 Or Index = 6 Then .Text = Format(.Text, "MM/DD/YY")
      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = oApp.getColor("HT1")
   End With
   
   pbGridFocus = False
   pnIndex = Index
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String
   
   lsOldProc = "txtField_KeyDown"
   'On Error GoTo errProc
   
   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
      With txtField(Index)
         If KeyCode = vbKeyF3 Then
            oTrans.SearchMaster Index, .Text
            If .Text <> "" Then SetNextFocus
         Else
            If .Text <> "" Then oTrans.SearchMaster Index, .Text
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

Private Sub initButton(lnStat As Integer)
   Dim lbShow As Boolean
   Dim lnCtr As Integer

   lbShow = IIf(lnStat = 0, False, True)
   cmdButton(4).Visible = Not lbShow
   cmdButton(5).Visible = Not lbShow
   cmdButton(6).Visible = Not lbShow
   
   cmdButton(0).Visible = lbShow
   cmdButton(1).Visible = lbShow
   cmdButton(2).Visible = lbShow
   cmdButton(3).Visible = lbShow
   cmdButton(7).Visible = lbShow

   xrFrame1.Enabled = lbShow
   
   With GridEditor1
      .ColEnabled(8) = lbShow
   End With
   
   If Not lbShow Then cmdButton(4).SetFocus
End Sub

Private Sub InitGrid()
   Dim lnCtr As Integer

   With GridEditor1
      .Cols = 10
      .Rows = 2
      .Font = "MS Sans Serif"

      'Column Title
      .TextMatrix(0, 1) = "Barcode"
      .TextMatrix(0, 2) = "Description"
      .TextMatrix(0, 3) = "Brand"
      .TextMatrix(0, 6) = "QOH"
      .TextMatrix(0, 7) = "Old SRP"
      .TextMatrix(0, 8) = "New SRP"
      .TextMatrix(0, 9) = "Stat"
      
      .Row = 0

      'Column Alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = 3
      Next

      'Column Width
      .ColWidth(0) = 300
      .ColWidth(1) = 1500
      .ColWidth(2) = 2450
      .ColWidth(3) = 2600
      .ColWidth(4) = 0
      .ColWidth(5) = 0
      .ColWidth(6) = 500
      .ColWidth(7) = 900
      .ColWidth(8) = 900
      .ColWidth(9) = 500
      
      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
         
      .ColEnabled(1) = False
      .ColEnabled(2) = False
      .ColEnabled(3) = False
      .ColEnabled(4) = False
      .ColEnabled(5) = False
      .ColEnabled(6) = False
      .ColEnabled(7) = False
      .ColEnabled(9) = False
      
      .ColNumberOnly(6) = True
      .ColNumberOnly(7) = True
      .ColNumberOnly(8) = True
      
      .ColFormat(6) = "#,##0"
      .ColFormat(7) = "#,##0.00"
      .ColFormat(8) = "#,##0.00"
      
      .ColDefault(6) = 0
      .ColDefault(7) = 0#
      .ColDefault(8) = 0#
      .ColDefault(9) = "x"
      
      .EditorBackColor = oApp.getColor("HT1")

      .Col = 1
      .Row = 1
   End With
End Sub

Private Sub ClearFields()
   For pnCtr = 0 To txtField.Count - 1
      Select Case pnCtr
      Case 0
         txtField(pnCtr).Text = Format(oTrans.Master(pnCtr), "@@@@-@@@@@@")
      Case 1
         txtField(pnCtr).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
      Case Else
         txtField(pnCtr).Text = oTrans.Master(pnCtr)
      End Select
   Next

   With GridEditor1
      .Rows = 2
      .Col = 1
      
      .TextMatrix(1, 1) = ""
      .TextMatrix(1, 2) = ""
      .TextMatrix(1, 3) = ""
      .TextMatrix(1, 4) = ""
      .TextMatrix(1, 5) = ""
      .TextMatrix(1, 6) = 0
      .TextMatrix(1, 7) = 0#
      .TextMatrix(1, 8) = 0#
      .TextMatrix(1, 9) = "x"
   End With
End Sub

Private Function isEntryOk() As Boolean
   If txtField(2).Text = "" Then
      MsgBox "Invalid Category Detected!!!" & vbCrLf & _
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
   isEntryOk = True
   Exit Function
EntryNotOK:
   isEntryOk = False
End Function

Public Function PrintTrans() As Boolean
   Dim loRS As ADODB.Recordset
   Dim lrs As ADODB.Recordset
   Dim lnCtr As Integer
   Dim lsOldProc As String
   
   lsOldProc = "PrintTrans"
   'On Error GoTo errProc

   PrintTrans = True
   
   Set loRS = New ADODB.Recordset
   
   loRS.Fields.Append "nQuantity", adInteger, 3
   loRS.Fields.Append "sModel", adVarChar, 50
   loRS.Fields.Append "sColor", adVarChar, 50
   loRS.Fields.Append "sDescription", adVarChar, 50
   loRS.Fields.Append "sBarrCode", adVarChar, 25
   loRS.Open
   
   With GridEditor1
      For lnCtr = 1 To .Rows - 1
         loRS.AddNew
         loRS("nQuantity").Value = .TextMatrix(lnCtr, 5)
         loRS("sModel").Value = .TextMatrix(lnCtr, 3)
         loRS("sColor").Value = .TextMatrix(lnCtr, 4)
         loRS("sDescription").Value = .TextMatrix(lnCtr, 2)
         loRS("sBarrCode").Value = .TextMatrix(lnCtr, 1)
      Next
   End With
   
   ' assign important info to the report
   Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\Purchase.rpt")
   oReport.DiscardSavedData
   oReport.FieldMappingType = crAutoFieldMapping
   oReport.Database.SetDataSource loRS
   
   Set lrs = New ADODB.Recordset
   lrs.Open "Select" _
               & "  CONCAT(b.sTownName, ', ', c.sProvName, ' ', b.sZippCode) xAddressx" _
            & " From Branch a" _
               & ", TownCity b" _
                  & " Left Join Province c" _
                     & " On b.sProvIDxx = c.sProvIDxx" _
            & " Where a.sTownIDxx = b.sTownIDxx" _
               & " And a.sBranchCd = " & strParm(oTrans.Master("sBranchCd")) _
            , oApp.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
   
   If Not lrs.EOF Then oReport.Sections("PH").ReportObjects("txtDeliver").SetText "           " & txtField(4).Text & vbCrLf & " " & lrs("xAddressx")
   oReport.Sections("PH").ReportObjects("txtSupplier").SetText txtField(2).Text
   oReport.Sections("PH").ReportObjects("txtTerm").SetText txtField(6).Text
   oReport.Sections("PH").ReportObjects("txtDDate").SetText txtField(7).Text
   oReport.Sections("PH").ReportObjects("txtDate").SetText txtField(1).Text
   oReport.Sections("PF").ReportObjects("txtUserRpt").SetText oApp.UserName

   oReport.PrintOutEx False, 1
   loRS.Close
   lrs.Close

endProc:
   oTrans.CloseTransaction (oTrans.Master(0))
   Set oReport = Nothing
   Set loRS = Nothing
   Set lrs = Nothing
   Exit Function
errProc:
   PrintTrans = False
   ShowError lsOldProc & "( " & " )"
End Function

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   With txtField(Index)
      .Text = TitleCase(.Text)
      Select Case Index
      Case 1
         If Not IsDate(.Text) Then .Text = oApp.ServerDate
         .Text = Format(.Text, "MMMM DD, YYYY")
      End Select
      oTrans.Master(Index) = .Text
   End With
End Sub

Private Sub LoadDetail()
   Dim lnCtr As Integer
   Dim lsOldProc As String
   
   lsOldProc = "LoadDetail"
   'On Error GoTo errProc
   
   If oTrans.Master("sCategrID") = "" Then Exit Sub
   
   With GridEditor1
      .Rows = oTrans.ItemCount + 1
   

      For lnCtr = 0 To oTrans.ItemCount - 1
         .TextMatrix(lnCtr + 1, 1) = oTrans.Detail(lnCtr, "sBarrCode")
         .TextMatrix(lnCtr + 1, 2) = oTrans.Detail(lnCtr, "sDescript")
         .TextMatrix(lnCtr + 1, 3) = oTrans.Detail(lnCtr, "sBrandNme") & " " & _
                                       oTrans.Detail(lnCtr, "sModelNme") & " " & _
                                       oTrans.Detail(lnCtr, "sColorNme")
         .TextMatrix(lnCtr + 1, 6) = oTrans.Detail(lnCtr, "nQtyOnHnd")
         .TextMatrix(lnCtr + 1, 7) = oTrans.Detail(lnCtr, "nOldSRPxx")
         .TextMatrix(lnCtr + 1, 8) = oTrans.Detail(lnCtr, "nNewSRPxx")
         .TextMatrix(lnCtr + 1, 9) = "x"
      Next
      
      If .Rows > 16 Then
         .ColWidth(2) = 2350
         .ColWidth(3) = 2500
      Else
         .ColWidth(2) = 2450
         .ColWidth(3) = 2600
      End If
   End With
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "()", False
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
