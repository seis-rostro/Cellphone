VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0"; "xrGridControl.ocx"
Begin VB.Form frmCP_Load_Receiving 
   BorderStyle     =   0  'None
   Caption         =   "Load Receiving"
   ClientHeight    =   7965
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11940
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   11940
   ShowInTaskbar   =   0   'False
   Begin xrGridEditor.GridEditor GridEditor1 
      Height          =   4290
      Left            =   1575
      TabIndex        =   19
      Top             =   3105
      Width           =   10245
      _ExtentX        =   18071
      _ExtentY        =   7567
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
      Object.HEIGHT          =   4290
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
      MOUSEICON       =   "frmCP_Load_Receiving.frx":0000
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
      TabIndex        =   22
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
      Picture         =   "frmCP_Load_Receiving.frx":001C
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   2490
      Index           =   0
      Left            =   1575
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   10245
      _ExtentX        =   18071
      _ExtentY        =   4392
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   8
         Left            =   7545
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   1650
         Width           =   2505
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   6
         Left            =   7545
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   1980
         Width           =   780
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   675
         Index           =   3
         Left            =   1365
         MultiLine       =   -1  'True
         TabIndex        =   5
         Text            =   "frmCP_Load_Receiving.frx":0796
         Top             =   990
         Width           =   4770
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   7545
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   660
         Width           =   2505
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   1365
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   660
         Width           =   4770
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   615
         Index           =   7
         Left            =   1365
         MultiLine       =   -1  'True
         TabIndex        =   7
         Text            =   "frmCP_Load_Receiving.frx":079E
         Top             =   1680
         Width           =   4770
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
         Left            =   1365
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   165
         Width           =   1950
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   5
         Left            =   7545
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   990
         Width           =   2505
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   4
         Left            =   7545
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   1320
         Width           =   2505
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   8385
         TabIndex        =   18
         Top             =   2055
         Width           =   1080
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Invoice"
         Height          =   285
         Index           =   8
         Left            =   6300
         TabIndex        =   14
         Top             =   1695
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Discount"
         Height          =   285
         Index           =   5
         Left            =   6300
         TabIndex        =   16
         Top             =   2025
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   285
         Index           =   4
         Left            =   165
         TabIndex        =   4
         Top             =   1020
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transact. Date"
         Height          =   285
         Index           =   1
         Left            =   6300
         TabIndex        =   8
         Top             =   705
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "P.O. No"
         Height          =   285
         Index           =   3
         Left            =   6300
         TabIndex        =   12
         Top             =   1350
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   285
         Index           =   7
         Left            =   165
         TabIndex        =   6
         Top             =   1695
         Width           =   1200
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   1
         Left            =   1455
         Tag             =   "et0;ht2"
         Top             =   270
         Width           =   1950
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Term"
         Height          =   285
         Index           =   6
         Left            =   6300
         TabIndex        =   10
         Top             =   1020
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier"
         Height          =   285
         Index           =   2
         Left            =   165
         TabIndex        =   2
         Top             =   705
         Width           =   1200
      End
      Begin VB.Label Label1 
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
         Height          =   285
         Index           =   0
         Left            =   165
         TabIndex        =   0
         Top             =   195
         Width           =   1200
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   90
      TabIndex        =   20
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
      Picture         =   "frmCP_Load_Receiving.frx":07A6
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   90
      TabIndex        =   21
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
      Picture         =   "frmCP_Load_Receiving.frx":0F20
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   90
      TabIndex        =   24
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
      Picture         =   "frmCP_Load_Receiving.frx":169A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   90
      TabIndex        =   23
      Top             =   4950
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
      Picture         =   "frmCP_Load_Receiving.frx":1E14
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   90
      TabIndex        =   25
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
      Picture         =   "frmCP_Load_Receiving.frx":258E
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   480
      Left            =   1575
      Tag             =   "wt0;fb0"
      Top             =   7395
      Width           =   10230
      _ExtentX        =   18045
      _ExtentY        =   847
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H000000FF&
         Height          =   360
         Index           =   9
         Left            =   8640
         TabIndex        =   26
         Text            =   "0.00"
         Top             =   45
         Width           =   1485
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "TTL Amt"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   14
         Left            =   7770
         TabIndex        =   27
         Top             =   105
         Width           =   810
      End
   End
End
Attribute VB_Name = "frmCP_Load_Receiving"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmCP_Load_Receiving"

Private WithEvents oTrans As clsCPLoadReceiving
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin

Dim pnIndex As Integer
Dim pbGridFocus As Boolean
Dim pnCtr As Integer

Private Sub cmdButton_Click(Index As Integer)
   Dim lsRep As String
   Dim lsOldProc As String
   Dim lnCtr As Integer

   lsOldProc = "cmdButton_Click"
   'On Error GoTo errProc

   With GridEditor1
      Select Case Index
      Case 0
         If .Rows > 2 Then
            pnCtr = 1
            Do While pnCtr <= .Rows
               If CDbl(.TextMatrix(pnCtr, 5)) = 0# And pnCtr <> 1 Then
                  .Row = pnCtr
                  If oTrans.deleteDetail(.Row - 1) Then .deleteRow
               Else
                  pnCtr = pnCtr + 1
               End If
                If pnCtr = .Rows Then Exit Do
            Loop

            .ColWidth(3) = 2750
            If .Rows > 16 Then .ColWidth(3) = 2550
         End If

         If isEntryOK Then
            If oTrans.SaveTransaction Then
               MsgBox "Record Updated Successfully!!!", vbInformation, "Notice"
               initButton xeModeReady
            Else
               MsgBox "Unable to Update Record!!!", vbCritical, "Warning"
            End If
         End If
      Case 1
         If pbGridFocus Then
            If oTrans.searchDetail(.Row - 1, 1) Then
               .Col = 5
            Else
               .Col = 1
            End If

            .Refresh
            .SetFocus
         Else
            oTrans.SearchMaster pnIndex
         End If
      Case 2
         If .Rows > 2 Then
            If oTrans.deleteDetail(.Row - 1) Then .deleteRow

            For pnCtr = 1 To .Rows - 1
               .TextMatrix(pnCtr, 0) = pnCtr
            Next

            .ColWidth(3) = 2750
            If .Rows > 16 Then .ColWidth(3) = 2550
         End If
      Case 3
         lsRep = MsgBox("Transaction is in Update Mode!!!" & vbCrLf & _
                        "Do you want to Cancel Transaction!!!", vbYesNo + vbQuestion, "Confirm")

         If lsRep = vbYes Then
            oTrans.NewTransaction
            ClearFields
            initButton xeModeReady
         Else
            txtField(pnIndex).SetFocus
         End If
      Case 4
         oTrans.NewTransaction

         ClearFields
         initButton xeModeAddNew
         txtField(2).SetFocus
      Case 5
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
   'On Error GoTo errProc

   CenterChildForm mdiMain, Me
   Set oTrans = New clsCPLoadReceiving
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

   For pnCtr = 1 To txtField.Count - 1
      txtField(pnCtr).MaxLength = oTrans.MasFldSize(pnCtr)
   Next

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
      ElseIf .TextMatrix(.Row, 5) = "0" Then
         Cancel = True
      End If
      If Not Cancel Then
         Call ComputeTotal
         oTrans.addDetail
      End If
      
      If .Rows > 16 Then .ColWidth(3) = 2550
   End With
End Sub

Private Sub GridEditor1_EditorValidate(Cancel As Boolean)
   Dim lsOldProc As String

   lsOldProc = "GridEditor1_EditorValidate"
   'On Error GoTo errProc

   With GridEditor1
      If .Col = 5 Then
         oTrans.Detail(.Row - 1, .Col) = CDbl(.TextMatrix(.Row, .Col))
      Else
         oTrans.Detail(.Row - 1, .Col) = .TextMatrix(.Row, .Col)
         If .Col = 1 Or .Col = 2 Then
            If .TextMatrix(.Row, .Col) <> "" Then .Col = 2
         End If
      End If
   End With

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
   'On Error GoTo errProc

   If KeyCode = vbKeyF3 Then
      With GridEditor1
         If oTrans.searchDetail(.Row - 1, .Col, .TextMatrix(.Row, .Col)) Then
            .Col = 5
         Else
            .Col = 1
         End If
         
         .Refresh
         .SetFocus
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
      Select Case .Col
      Case 5
         oTrans.Detail(.Row - 1, .Col) = CDbl(.TextMatrix(.Row, .Col))
      Case Else
         oTrans.Detail(.Row - 1, .Col) = .TextMatrix(.Row, .Col)
      End Select
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

Private Sub InitGrid()
   With GridEditor1
      .Rows = 2
      .Cols = 6
      .Font = "MS Sans Serif"

      'column title
      .TextMatrix(0, 1) = "BarrCode"
      .TextMatrix(0, 2) = "Description"
      .TextMatrix(0, 3) = "Model"
      .TextMatrix(0, 4) = "AOH"
      .TextMatrix(0, 5) = "Amount"

      .Row = 0

      'column alignment
      For pnCtr = 0 To .Cols - 1
         .Col = pnCtr
         .CellFontBold = True
         .CellAlignment = 3
      Next

      'column width
      .ColWidth(0) = 330
      .ColWidth(1) = 2100
      .ColWidth(2) = 2600
      .ColWidth(4) = 1200
      .ColWidth(5) = 1200

      .ColFormat(4) = "#,##0.00"
      .ColFormat(5) = "#,##0.00"
      .ColNumberOnly(6) = True
      .ColDefault(4) = 0#
      .ColDefault(5) = 0#

      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      .ColAlignment(4) = 6
      .ColAlignment(5) = 6

      .ColEnabled(3) = False
      .ColEnabled(4) = False

      .EditorBackColor = oApp.getColor("HT1")

      .Row = 1
      .Col = 1
   End With
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      Select Case Index
      Case 1
         .Text = Format(.Text, "MM/DD/YY")
      End Select

      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = oApp.getColor("HT1")
   End With

   pbGridFocus = False
   pnIndex = Index
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

Private Sub initButton(lnStat As Integer)
   Dim lbShow As Boolean

   lbShow = IIf(lnStat = 0, False, True)
   cmdButton(4).Visible = Not lbShow
   cmdButton(5).Visible = Not lbShow

   cmdButton(0).Visible = lbShow
   cmdButton(1).Visible = lbShow
   cmdButton(2).Visible = lbShow
   cmdButton(3).Visible = lbShow

   For pnCtr = 1 To txtField.Count - 1
      If pnCtr <> 3 Then txtField(pnCtr).Enabled = lbShow
   Next

   With GridEditor1
      .ColEnabled(1) = lbShow
      .ColEnabled(2) = lbShow
      .ColEnabled(5) = lbShow
   End With

   If Not lbShow Then cmdButton(4).SetFocus
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

Private Sub ClearFields()
   For pnCtr = 0 To txtField.Count - 1
      Select Case pnCtr
      Case 0
         txtField(pnCtr).Text = Format(oTrans.Master(0), "@@@@-@@@@@@")
      Case 1
         txtField(pnCtr).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
      Case 6
         txtField(pnCtr).Text = "0.00"
      Case Else
         txtField(pnCtr).Text = oTrans.Master(pnCtr)
      End Select
   Next

   With GridEditor1
      .Rows = 2
      .ColWidth(3) = 2750

      'empty row
      .TextMatrix(1, 1) = ""
      .TextMatrix(1, 2) = ""
      .TextMatrix(1, 3) = ""
      .TextMatrix(1, 4) = "0.00"
      .TextMatrix(1, 5) = "0.00"
   End With
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
      Case 6, 9
         If Not IsNumeric(.Text) Then .Text = 0#
         .Text = Format(.Text, "#,##0.00")
      Case 4, 8
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

   If txtField(4).Text = "" Then
      MsgBox "Unknown PO Number!!!" & vbCrLf & _
             "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
      txtField(4).SetFocus
      GoTo EntryNotOK
   End If

   If txtField(5).Text = "" Then
      MsgBox "Invalid Term Detected!!!" & vbCrLf & _
             "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
      txtField(5).SetFocus
      GoTo EntryNotOK
   End If
   
   With GridEditor1
      If Trim(.TextMatrix(1, 1)) = "" Or .TextMatrix(1, 5) = 0 Then
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

Private Function ComputeTotal() As Double
   Dim lnCtr As Integer
   Dim lnSum As Double

   lnSum = 0#
   With GridEditor1
      For lnCtr = 1 To .Rows - 1
         lnSum = CDbl(.TextMatrix(lnCtr, 5)) + lnSum
      Next
   End With

   txtField(9).Text = Format(lnSum, "#,##0.00")

   oTrans.Master("nTranTotl") = CDbl(txtField(9).Text)
End Function
