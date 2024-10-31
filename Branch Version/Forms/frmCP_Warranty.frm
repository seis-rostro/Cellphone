VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0"; "xrGridControl.ocx"
Begin VB.Form frmCP_Warranty 
   BorderStyle     =   0  'None
   Caption         =   "Warranty Transfer"
   ClientHeight    =   6900
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
   ScaleHeight     =   6900
   ScaleWidth      =   11715
   ShowInTaskbar   =   0   'False
   Begin xrGridEditor.GridEditor GridEditor1 
      Height          =   3780
      Left            =   1575
      TabIndex        =   12
      Tag             =   "et0;eb0;et0;bc2"
      Top             =   3000
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   6668
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
      Object.HEIGHT          =   3780
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
      MOUSEICON       =   "frmCP_Warranty.frx":0000
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
      Height          =   600
      Index           =   2
      Left            =   90
      TabIndex        =   15
      Top             =   4050
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
      Picture         =   "frmCP_Warranty.frx":001C
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   90
      TabIndex        =   13
      Top             =   2790
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
      Picture         =   "frmCP_Warranty.frx":0796
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   90
      TabIndex        =   14
      Top             =   3420
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
      Picture         =   "frmCP_Warranty.frx":0F10
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   90
      TabIndex        =   18
      Top             =   4680
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
      Picture         =   "frmCP_Warranty.frx":168A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   90
      TabIndex        =   16
      Top             =   3420
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
      Picture         =   "frmCP_Warranty.frx":1E04
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   6
      Left            =   90
      TabIndex        =   19
      Top             =   4680
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
      Picture         =   "frmCP_Warranty.frx":257E
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   90
      TabIndex        =   17
      Top             =   4050
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
      Picture         =   "frmCP_Warranty.frx":2CF8
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   2385
      Index           =   0
      Left            =   1575
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   10020
      _ExtentX        =   17674
      _ExtentY        =   4207
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   660
         Index           =   4
         Left            =   1230
         MultiLine       =   -1  'True
         TabIndex        =   11
         Text            =   "frmCP_Warranty.frx":3472
         Top             =   1575
         Width           =   8520
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   6660
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   765
         Width           =   3090
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   465
         Index           =   3
         Left            =   1230
         MultiLine       =   -1  'True
         TabIndex        =   5
         Text            =   "frmCP_Warranty.frx":347A
         Top             =   1095
         Width           =   4545
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
         Left            =   1230
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   120
         Width           =   2265
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   1230
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   765
         Width           =   4545
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   6660
         TabIndex        =   9
         Tag             =   "ht0;ft0"
         Top             =   1095
         Width           =   3090
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   195
         Index           =   4
         Left            =   315
         TabIndex        =   10
         Top             =   1605
         Width           =   735
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   195
         Index           =   1
         Left            =   6120
         TabIndex        =   6
         Top             =   810
         Width           =   345
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   195
         Index           =   6
         Left            =   315
         TabIndex        =   4
         Top             =   1125
         Width           =   570
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   315
         Left            =   1320
         Tag             =   "et0;ht2"
         Top             =   225
         Width           =   2265
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
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
         Index           =   0
         Left            =   150
         TabIndex        =   0
         Top             =   150
         Width           =   915
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier"
         Height          =   195
         Index           =   2
         Left            =   315
         TabIndex        =   2
         Top             =   810
         Width           =   570
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         Height          =   195
         Index           =   3
         Left            =   6120
         TabIndex        =   8
         Top             =   1125
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmCP_Warranty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmCP_PO_Return"

Private WithEvents oTrans As clsCPPOReturn
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
   Dim lsRep As String
   
   lsOldProc = "cmdButton_Click"
   On Error Goto errProc
   
   If Not pbGridFocus And Index = 0 Then Call txtField_Validate(pnIndex, False)
   With GridEditor1
      Select Case Index
      Case 0
         If .Rows > 2 Then
            pnCtr = 0
            Do While pnCtr < .Rows
               If Trim(.TextMatrix(pnCtr, 1)) = "" Then
                  .Row = pnCtr
                  If oTrans.DeleteDetail(.Row - 1) Then .DeleteRow
               Else
                  pnCtr = pnCtr + 1
               End If
            Loop

            .ColWidth(3) = 3100
            If .Rows > 16 Then .ColWidth(3) = 2900
         End If

         If isEntryOK Then
            If oTrans.SaveTransaction Then
               MsgBox "Record Updated Successfully!!!", vbInformation, "Confirm"
               InitButton xeModeReady

               lsRep = MsgBox("Print Transaction!!!", vbYesNo + vbQuestion, "Confirm")
               If lsRep = vbYes Then
                  If Not PrintTrans Then MsgBox "Unable to Print Transaction!!!", vbCritical, "Warning"
               End If
               pbSave = True
            Else
               MsgBox "Unable to Update Record!!!", vbCritical, "Warning"
            End If
         End If
      Case 1
         If pbGridFocus Then
            If oTrans.SearchDetail(.Row - 1, 1) Then .Col = 1
            .Refresh
            .SetFocus
         Else
            oTrans.SearchMaster pnIndex
         End If
      Case 2
         If .Rows > 2 Then
            If oTrans.DeleteDetail(.Row - 1) Then .DeleteRow

            For pnCtr = 1 To .Rows - 1
               .TextMatrix(pnCtr, 0) = pnCtr
            Next

            .ColWidth(3) = 3100
            If .Rows > 16 Then .ColWidth(3) = 2900
         End If
      Case 3
         lsRep = MsgBox("Transaction is in Update Mode!!!" & vbCrLf & _
                        "Do you want to Cancel Transaction!!!", vbYesNo + vbQuestion, "Confirm")
         
         If lsRep = vbYes Then
            oTrans.NewTransaction
            ClearFields
            InitButton xeModeReady
         Else
            txtField(pnIndex).SetFocus
         End If
         pbSave = False
      Case 4
         oTrans.NewTransaction
         ClearFields
         InitButton xeModeAddNew
         txtField(1).SetFocus
      Case 5
         If pbSave Then
            lsRep = MsgBox("Print Transaction!!!", vbYesNo + vbQuestion, "Confirm")
            If lsRep = vbYes Then
               If Not PrintTrans Then MsgBox "Unable to Print Transaction!!!", vbCritical, "Warning"
            End If
         Else
            MsgBox "Unable to Print Transaction!!!", vbCritical, "Warning"
         End If
      Case 6
         Unload Me
      End Select
   End With
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Sub Form_Activate()
   GridEditor1.Refresh

   oApp.MenuName = Me.Tag
   Me.ZOrder 0
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String
   
   lsOldProc = "Form_Load"
   On Error Goto errProc
   
   CenterChildForm mdiMain, Me

   Set oTrans = New clsCPPOReturn
   Set oTrans.AppDriver = oApp
   
   oTrans.InitTransaction
   oTrans.NewTransaction
   
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransaction
   
   InitGrid
   ClearFields
   InitButton xeModeAddNew
   
   txtField(4).MaxLength = oTrans.MasFldSize(4)
   
   pbGridValidate = False
   
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
      ElseIf .TextMatrix(.Row, 6) = "0" Then
         Cancel = True
      End If
      If Not Cancel Then
         If .Row = .Rows - 1 Then oTrans.AddDetail
      End If
 
      If .Rows > 16 Then .ColWidth(3) = 2900
   End With
End Sub

Private Sub GridEditor1_EditorValidate(Cancel As Boolean)
   Dim lsOldProc As String
   
   lsOldProc = "GridEditor1_EditorValidate"
   On Error Goto errProc
   
   With GridEditor1
      If pbGridValidate Then
         pbGridValidate = False
         Exit Sub
      End If

      If .Col = 1 Or .Col = 2 Then
         .TextMatrix(.Row, .Col) = compareSerial(.TextMatrix(.Row, .Col), .Row)
      End If
      
      oTrans.Detail(.Row - 1, .Col) = .TextMatrix(.Row, .Col)
      If oTrans.Detail(.Row - 1, "cHsSerial") = xeYes Then
         Select Case .Col
         Case 1, 2
            If .TextMatrix(.Row, .Col) <> "" Then
               oTrans.Detail(.Row - 1, "nQuantity") = 1
               .TextMatrix(.Row, 6) = oTrans.Detail(.Row - 1, "nQuantity")
               If .Row = .Rows - 1 Then
                  .Rows = .Rows + 1
                  oTrans.AddDetail
                  .Col = 0
               End If

               .Row = .Rows - 1
            End If
         Case 6
'            If CDbl(.TextMatrix(.Row, 6)) > CDbl(.TextMatrix(.Row, 5)) Then
'               .TextMatrix(.Row, .Col) = 0
'            End If
            
            If CDbl(.TextMatrix(.Row, .Col)) > 1 Then .TextMatrix(.Row, .Col) = 1
         End Select
      End If
               
      If .Rows > 16 Then
         .TopRow = .Rows - 1
         .ColWidth(3) = 2900
      End If
   End With
   pbGridValidate = True
   
endProc:
   GridEditor1.Refresh
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Cancel & " )", True
End Sub

Private Sub GridEditor1_GotFocus()
   With GridEditor1
      .EditorBackColor = oApp.getColor("HT1")
   End With
   pbGridFocus = True
End Sub

Private Sub GridEditor1_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String
   
   lsOldProc = "GridEditor1_KeyDown"
   On Error Goto errProc
   
   If KeyCode = vbKeyF3 Then
      With GridEditor1
         If oTrans.SearchDetail(.Row - 1, .Col, .TextMatrix(.Row, .Col)) Then
            If oTrans.Detail(.Row - 1, "cHsSerial") = xeYes Then
               .TextMatrix(.Row, 6) = 1
               oTrans.Detail(.Row - 1, "nQuantity") = 1
               If .Row = .Rows - 1 Then
                  .Rows = .Rows + 1
                  oTrans.AddDetail
               End If

               .Row = .Rows - 1
               .Col = 1
            Else
               .Col = 6
            End If
         Else
            oTrans.Detail(.Row - 1, .Col) = .TextMatrix(.Row, .Col)
            .Col = 1
         End If
         
         .Refresh
         .SetFocus
         If .Rows > 16 Then
            .TopRow = .Rows - 1
            .ColWidth(3) = 2900
         End If
         KeyCode = 0
      End With
   End If
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " _
                       & "  " & KeyCode _
                       & ", " & Shift _
                       & " )", True
End Sub

Private Sub GridEditor1_LostFocus()
   With GridEditor1
      .EditorBackColor = oApp.getColor("EB")
      If cmdButton(0).Visible Then oTrans.Detail(.Row - 1, .Col) = .TextMatrix(.Row, .Col)
      If .Rows > 16 Then .TopRow = .Rows - 1
   End With
   
   pbGridValidate = False
End Sub

Private Sub oTrans_DetailRetrieved(ByVal Index As Integer)
   With GridEditor1
      .TextMatrix(.Row, Index) = oTrans.Detail(.Row - 1, Index)
   End With
End Sub

Private Sub oTrans_MasterRetrieved(ByVal Index As Integer)
   txtField(Index).Text = oTrans.Master(Index)
End Sub

Private Sub InitGrid()
   With GridEditor1
      .Rows = 2
      .Cols = 7
      .Font = "MS Sans Serif"

      'column title
      .TextMatrix(0, 1) = "BarrCode"
      .TextMatrix(0, 2) = "Description"
      .TextMatrix(0, 3) = "Model"
      .TextMatrix(0, 4) = "Unit Price"
      .TextMatrix(0, 5) = "QOH"
      .TextMatrix(0, 6) = "Qty"
      
      .Row = 0

      'column alignment
      For pnCtr = 0 To .Cols - 1
         .Col = pnCtr
         .CellFontBold = True
         .CellAlignment = 3
      Next

      'column width
      .ColWidth(0) = 330
      .ColWidth(1) = 2000
      .ColWidth(2) = 2500
      .ColWidth(4) = 1020
      .ColWidth(5) = 500
      .ColWidth(6) = 500
      
      .ColFormat(4) = "#,##0.00"
      .ColFormat(5) = "#,##0"
      .ColFormat(6) = "#,##0"
      .ColNumberOnly(6) = True
      .ColDefault(4) = 0#
      .ColDefault(5) = 0
      .ColDefault(6) = 0
      
      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      .ColAlignment(4) = 6
      .ColAlignment(5) = 6
      .ColAlignment(6) = 6
      
      .ColEnabled(3) = False
      .ColEnabled(4) = False
      .ColEnabled(5) = False
      
      .EditorBackColor = oApp.getColor("HT1")
      
      .Row = 1
      .Col = 1
   End With
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      If Index = 1 Then .Text = Format(.Text, "MM/DD/YY")
      
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
   On Error Goto errProc
   
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

Private Sub InitButton(lnStat As Integer)
   Dim lbShow As Boolean

   lbShow = IIf(lnStat = 0, False, True)
   cmdButton(4).Visible = Not lbShow
   cmdButton(5).Visible = Not lbShow
   cmdButton(6).Visible = Not lbShow
   
   cmdButton(0).Visible = lbShow
   cmdButton(1).Visible = lbShow
   cmdButton(2).Visible = lbShow
   cmdButton(3).Visible = lbShow
   
   For pnCtr = 1 To txtField.Count - 3
      txtField(pnCtr).Enabled = lbShow
   Next
   
   With GridEditor1
      .ColEnabled(1) = lbShow
      .ColEnabled(2) = lbShow
      .ColEnabled(6) = lbShow
   End With
   
   If Not lbShow Then cmdButton(4).SetFocus
End Sub

Private Function PrintTrans() As Boolean
   Dim loreport As frmRepViewer

   Dim lrs As ADODB.Recordset
   Dim lors As ADODB.Recordset
   Dim lnCtr As Integer
   Dim lsOldProc As String
   Dim lsStockIDx As String
   
   lsOldProc = "PrintTrans"
   On Error Goto errProc

   PrintTrans = True
   oTrans.OpenTransaction "0108000001"
   Set lrs = New ADODB.Recordset
   lrs.Fields.Append "nField01", adInteger, 3
   lrs.Fields.Append "nField02", adChar, 1
   lrs.Fields.Append "sField01", adVarChar, 20
   lrs.Fields.Append "sField02", adVarChar, 128
   lrs.Fields.Append "sField03", adVarChar, 20
   lrs.Fields.Append "sField04", adVarChar, 128
   lrs.Fields.Append "sField05", adVarChar, 100
   lrs.Fields.Append "sField06", adVarChar, 25
   lrs.Fields.Append "sField07", adVarChar, 25
   lrs.Open

   With oTrans
      For lnCtr = 0 To .ItemCount - 1
         lrs.AddNew
         lrs.Fields("nField01") = oTrans.Detail(lnCtr, "nQuantity")
         lrs.Fields("nField02") = oTrans.Detail(lnCtr, "cHsSerial")
         lrs.Fields("sField01") = oTrans.Detail(lnCtr, "sTransNox")
         lrs.Fields("sField02") = oTrans.Detail(lnCtr, "sStockIDx")
         lrs.Fields("sField03") = oTrans.Detail(lnCtr, "sBarrCode")
         lrs.Fields("sField04") = oTrans.Detail(lnCtr, "sDescript")
         lrs.Fields("sField05") = oTrans.Detail(lnCtr, "sBrandNme")
         lrs.Fields("sField06") = IFNull(oTrans.Detail(lnCtr, "sSerialNo"), "")
      Next
      lrs.Sort = "nField02,sField05,sField03,sField06"
   End With

   Set lors = New ADODB.Recordset
   If lors.State = adStateOpen Then lors.Close

   lors.Open "SELECT" _
               & "  CONCAT(a.sAddressx, ', ', b.sTownName, ', ', c.sProvName, ' ', b.sZippCode) as xAddressx" _
               & ", a.sBranchNm" _
            & " From Branch a" _
               & ", TownCity b" _
               & ", Province c" _
            & " WHERE a.sBranchCd = " & strParm(Left(oTrans.Master("sTransNox"), 2)) _
               & " AND a.sTownIDxx = b.sTownIDxx" _
               & " AND b.sProvIDxx = c.sProvIDxx" _
            , oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
      
   Set lors = New ADODB.Recordset
   If lors.State = adStateOpen Then lors.Close

   lors.Open "SELECT" _
               & "  CONCAT(a.sAddressx, ', ', b.sTownName, ', ', c.sProvName, ' ', b.sZippCode) as xAddressx" _
               & ", a.sBranchNm" _
            & " From Branch a" _
               & ", TownCity b" _
               & ", Province c" _
            & " WHERE a.sBranchCd = " & strParm(Left(oTrans.Master("sTransNox"), 2)) _
               & " AND a.sTownIDxx = b.sTownIDxx" _
               & " AND b.sProvIDxx = c.sProvIDxx" _
            , oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText

   Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\CPPurchaseReturnFormOld.rpt")
   'assign important info to the report
   oReport.DiscardSavedData
   oReport.FieldMappingType = crAutoFieldMapping
   oReport.Database.SetDataSource lrs
   
   oReport.Sections("PHa").ReportObjects("txtTransNox").SetText "CP-" & Right(oTrans.Master("sTransNox"), 8)
   oReport.Sections("PHa").ReportObjects("txtDate").SetText txtField(1).Text
   oReport.Sections("PHb").ReportObjects("txtTo").SetText txtField(2).Text
   oReport.Sections("PHb").ReportObjects("txtToAddress").SetText txtField(3).Text
   oReport.Sections("PHb").ReportObjects("txtFrom").SetText lors("sBranchNm")
   oReport.Sections("PHb").ReportObjects("txtFromAddress").SetText lors("xAddressx")
   oReport.Sections("RFb").ReportObjects("txtRemarks").SetText txtField(4).Text
   oReport.Sections("PF").ReportObjects("txtRptUser").SetText oApp.UserName
   
   Set loreport = New frmRepViewer
   Set loreport.ReportSource = oReport
   loreport.Show

endPoc:
   If Not pbPosted Then
      oTrans.CloseTransaction (oTrans.Master(0))
      pbPosted = True
   End If
   Set oReport = Nothing
   Set lrs = Nothing
   Set lors = Nothing
   Set loreport = Nothing
   Exit Function
errProc:
   PrintTrans = False
   ShowError lsOldProc & "( " & " )"
End Function

Private Function PrintTrans1() As Boolean
   Const lnRows As Integer = 10
   
   Dim loreport As frmRepViewer
   Dim CRXSubreport As Report
   Dim CRXSections As Sections
   Dim CRXSection As Section
   Dim CRXSubreportObj As SubreportObject
   Dim CRXReportObjects As ReportObjects
   Dim CRXReportObject As Object

   Dim lrs As ADODB.Recordset
   Dim lors As ADODB.Recordset
   Dim lrsMaster As Recordset
   Dim lnCtr As Integer
   Dim lsOldProc As String
   Dim loSubReport As Report
   Dim lnTotalItem As Integer
   Dim lsStockIDx As String
   Dim lbInserted As Boolean
   
   lsOldProc = "PrintTrans"
   On Error Goto errProc

   PrintTrans1 = True
   
   Set lrs = New ADODB.Recordset
   lrs.Fields.Append "nField01", adInteger, 3
   lrs.Fields.Append "nField02", adChar, 1
   lrs.Fields.Append "sField01", adVarChar, 20
   lrs.Fields.Append "sField02", adVarChar, 128
   lrs.Fields.Append "sField03", adVarChar, 20
   lrs.Fields.Append "sField04", adVarChar, 128
   lrs.Fields.Append "sField05", adVarChar, 100
   lrs.Fields.Append "sField06", adVarChar, 25
   lrs.Fields.Append "sField07", adVarChar, 25
   lrs.Fields.Append "cField01", adChar, 1
   lrs.Open
   
   Set lrsMaster = New ADODB.Recordset
   lrsMaster.Fields.Append "sField01", adVarChar, 20
   lrsMaster.Open
   
   With oTrans
      lnTotalItem = IIf(.ItemCount \ lnRows = 0, lnRows, _
                     IIf(.ItemCount Mod lnRows = 0, _
                     (.ItemCount / lnRows) * lnRows, _
                     ((.ItemCount \ lnRows) + 1) * lnRows))
      
      lbInserted = False
      For lnCtr = 0 To lnTotalItem - 1
         lrs.AddNew
         If .ItemCount > lnCtr Then
            lrs.Fields("nField01") = oTrans.Detail(lnCtr, "nQuantity")
            lrs.Fields("nField02") = oTrans.Detail(lnCtr, "cHsSerial")
            lrs.Fields("sField01") = oTrans.Detail(lnCtr, "sTransNox")
            lrs.Fields("sField02") = oTrans.Detail(lnCtr, "sStockIDx")
            lrs.Fields("sField03") = oTrans.Detail(lnCtr, "sBarrCode")
            lrs.Fields("sField04") = oTrans.Detail(lnCtr, "sDescript")
            lrs.Fields("sField05") = oTrans.Detail(lnCtr, "sBrandNme")
            lrs.Fields("sField06") = oTrans.Detail(lnCtr, "sSerialNo")
            lrs.Fields("sField07") = ""
            lrs.Fields("cField01") = 0
         Else
'            lrs.Fields("nField01") = 1
            lrs.Fields("nField02") = 3
            lrs.Fields("sField01") = oTrans.Master("sTransNox")
            lrs.Fields("sField02") = lnCtr
            lrs.Fields("sField03") = ""
            lrs.Fields("sField04") = ""
            lrs.Fields("sField05") = ""
            lrs.Fields("sField06") = ""
            lrs.Fields("sField07") = ""
            If Not lbInserted Then
               lrs.Fields("sField07") = "==========XXX=========="
               lbInserted = True
            End If
            lrs.Fields("cField01") = 1
         End If
      Next
      lrs.Sort = "nField02,sField05,sField03,sField06"
   End With

   Set lors = New ADODB.Recordset
   If lors.State = adStateOpen Then lors.Close

   lors.Open "SELECT" _
               & "  CONCAT(a.sAddressx, ', ', b.sTownName, ', ', c.sProvName, ' ', b.sZippCode) as xAddressx" _
               & ", a.sBranchNm" _
            & " From Branch a" _
               & ", TownCity b" _
               & ", Province c" _
            & " WHERE a.sBranchCd = " & strParm(Left(oTrans.Master("sTransNox"), 2)) _
               & " AND a.sTownIDxx = b.sTownIDxx" _
               & " AND b.sProvIDxx = c.sProvIDxx" _
            , oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
      
   Set lors = New ADODB.Recordset
   If lors.State = adStateOpen Then lors.Close

   lors.Open "SELECT" _
               & "  CONCAT(a.sAddressx, ', ', b.sTownName, ', ', c.sProvName, ' ', b.sZippCode) as xAddressx" _
               & ", a.sBranchNm" _
            & " From Branch a" _
               & ", TownCity b" _
               & ", Province c" _
            & " WHERE a.sBranchCd = " & strParm(Left(oTrans.Master("sTransNox"), 2)) _
               & " AND a.sTownIDxx = b.sTownIDxx" _
               & " AND b.sProvIDxx = c.sProvIDxx" _
            , oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText

   Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\CPPurchaseReturnForm.rpt")
   Set CRXSections = oReport.Sections
   For Each CRXSection In CRXSections
      Set CRXReportObjects = CRXSection.ReportObjects
      For Each CRXReportObject In CRXReportObjects
         If CRXReportObject.Kind = crSubreportObject Then
            Set CRXSubreportObj = CRXReportObject
            Set CRXSubreport = CRXSubreportObj.OpenSubreport
            Select Case CRXSubreportObj.Name
            Case "SubMaster1"
               CRXSubreport.Sections("RHa").ReportObjects("txtDate").SetText txtField(1).Text
               CRXSubreport.Sections("RHb").ReportObjects("txtTo").SetText txtField(2).Text
               CRXSubreport.Sections("RHb").ReportObjects("txtToAddress").SetText txtField(3).Text
               CRXSubreport.Sections("RHb").ReportObjects("txtFrom").SetText lors("sBranchNm")
               CRXSubreport.Sections("RHb").ReportObjects("txtFromAddress").SetText lors("xAddressx")
               CRXSubreport.Sections("RFb").ReportObjects("txtPrepared").SetText oApp.UserName
               CRXSubreport.Sections("RFb").ReportObjects("txtRemarks").SetText txtField(4).Text
               CRXSubreport.Database.SetDataSource lrs
            Case "SubMaster2"
               CRXSubreport.Sections("RHa").ReportObjects("txtDate").SetText txtField(1).Text
               CRXSubreport.Sections("RHb").ReportObjects("txtTo").SetText txtField(2).Text
               CRXSubreport.Sections("RHb").ReportObjects("txtToAddress").SetText txtField(3).Text
               CRXSubreport.Sections("RHb").ReportObjects("txtFrom").SetText lors("sBranchNm")
               CRXSubreport.Sections("RHb").ReportObjects("txtFromAddress").SetText lors("xAddressx")
               CRXSubreport.Sections("RFb").ReportObjects("txtPrepared").SetText oApp.UserName
               CRXSubreport.Sections("RFb").ReportObjects("txtRemarks").SetText txtField(4).Text
               CRXSubreport.Database.SetDataSource lrs
            End Select
         End If
      Next CRXReportObject
   Next CRXSection
   
   lrsMaster.AddNew
   lrsMaster.Fields("sField01") = oTrans.Master("sTransNox")
   
   'assign important info to the report
   oReport.DiscardSavedData
   oReport.FieldMappingType = crAutoFieldMapping
   oReport.Database.SetDataSource lrsMaster
   
   Set loreport = New frmRepViewer
   Set loreport.ReportSource = oReport
   loreport.Show

endPoc:
   If Not pbPosted Then
      oTrans.CloseTransaction (oTrans.Master(0))
      pbPosted = True
   End If
   Set oReport = Nothing
   Set lrs = Nothing
   Set lors = Nothing
   Set loreport = Nothing
   Exit Function
errProc:
   PrintTrans1 = False
   ShowError lsOldProc & "( " & " )"
End Function

Private Sub ClearFields()
   For pnCtr = 0 To txtField.Count - 1
      Select Case pnCtr
      Case 0
         txtField(pnCtr).Text = Format(oTrans.Master(pnCtr), "@@@@-@@@@@@")
      Case 1
         txtField(pnCtr).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
      Case Else
         txtField(pnCtr).Text = Empty
      End Select
   Next
   
   With GridEditor1
      .Rows = 2
      .ColWidth(3) = 3100
      
      'empty row
      .TextMatrix(1, 1) = ""
      .TextMatrix(1, 2) = ""
      .TextMatrix(1, 3) = ""
      .TextMatrix(1, 4) = "0.00"
      .TextMatrix(1, 5) = "0"
      .TextMatrix(1, 6) = "0"
   End With
   
   pbSave = False
   pbPosted = False
   lblTotal.Caption = "0.00"
End Sub

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
      Case 3
         .Text = Format(.Text, ">")
      End Select
      
      oTrans.Master(Index) = .Text
   End With
End Sub

Private Function isEntryOK() As Boolean
   If txtField(2).Text = "" Then
      MsgBox "Supplier not found!!!" & vbCrLf & _
             "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
      txtField(2).SetFocus
      GoTo EntryNotOK
   End If
   
   With GridEditor1
      If Trim(.TextMatrix(1, 1)) = "" Then
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

Private Function compareSerial(Value As String, Row As Integer) As String
   Dim lnRep As Integer
   Dim lnCtr As Integer
   Dim lsValue As String
   Dim lnValue As Integer
   
   If Trim(Value) = "" Then
      compareSerial = ""
      Exit Function
   End If
   
   With GridEditor1
      For lnCtr = 1 To .Rows - 1
         If .TextMatrix(lnCtr, 1) = Value And lnCtr <> Row Then
            If oTrans.Detail(lnCtr - 1, "cHsSerial") = xeYes Then
               MsgBox "Duplicate Serial No!!!" & vbCrLf & _
                        "Please Verify your entry then try again!!!", vbCritical, "Warning"
            Else
               lnRep = MsgBox("Duplicate Serial No!!!" & vbCrLf & _
                                 "Item automatically add from existing serial!!!", vbYesNo + vbQuestion, "CONFIRMATION")
               If lnRep = vbYes Then
                  lsValue = InputBox("Please specify quantity for serial " & Value & vbCrLf & _
                                       .TextMatrix(lnCtr, 2) & vbCrLf & _
                                       .TextMatrix(lnCtr, 3), "Quantity", 0)
                  lnValue = IIf(lsValue = "", 0, lsValue)

                  .TextMatrix(lnCtr, 6) = .TextMatrix(lnCtr, 6) + lnValue
                  oTrans.Detail(lnCtr - 1, "nQuantity") = CDbl(.TextMatrix(lnCtr, 6))
               End If
            End If
            compareSerial = ""
         Else
            compareSerial = Value
         End If
      Next
   End With
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
