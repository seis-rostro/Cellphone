VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0"; "xrGridControl.ocx"
Begin VB.Form frmCP_Purchasing 
   BorderStyle     =   0  'None
   Caption         =   "Purchase Order"
   ClientHeight    =   7650
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12240
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   12240
   ShowInTaskbar   =   0   'False
   Begin xrGridEditor.GridEditor GridEditor1 
      Height          =   3780
      Left            =   1575
      TabIndex        =   8
      Tag             =   "et0;eb0;et0;bc2"
      Top             =   3705
      Width           =   10530
      _ExtentX        =   18574
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
      MOUSEICON       =   "frmCP_Purchasing.frx":0000
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
      TabIndex        =   20
      Top             =   1785
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
      Picture         =   "frmCP_Purchasing.frx":001C
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   90
      TabIndex        =   9
      Top             =   525
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
      Picture         =   "frmCP_Purchasing.frx":0796
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   90
      TabIndex        =   19
      Top             =   1155
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
      Picture         =   "frmCP_Purchasing.frx":0F10
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   90
      TabIndex        =   23
      Top             =   2415
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
      Picture         =   "frmCP_Purchasing.frx":168A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   90
      TabIndex        =   21
      Top             =   1155
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
      Picture         =   "frmCP_Purchasing.frx":1E04
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   6
      Left            =   90
      TabIndex        =   24
      Top             =   2415
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
      Picture         =   "frmCP_Purchasing.frx":257E
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   90
      TabIndex        =   22
      Top             =   1785
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
      Picture         =   "frmCP_Purchasing.frx":2CF8
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   3135
      Index           =   0
      Left            =   1575
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   10530
      _ExtentX        =   18574
      _ExtentY        =   5530
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   24
         Left            =   1095
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   765
         Width           =   3030
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   16
         Left            =   1095
         TabIndex        =   29
         Top             =   1770
         Width           =   5265
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   13
         Left            =   7305
         TabIndex        =   7
         Top             =   1830
         Width           =   3030
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   12
         Left            =   7305
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   1470
         Width           =   3030
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   4
         Left            =   1095
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   2100
         Width           =   5265
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   6
         Left            =   7305
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   1110
         Width           =   3030
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   540
         Index           =   5
         Left            =   1095
         MultiLine       =   -1  'True
         TabIndex        =   4
         Text            =   "frmCP_Purchasing.frx":3472
         Top             =   2430
         Width           =   5265
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   7305
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   765
         Width           =   3030
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   1095
         MultiLine       =   -1  'True
         TabIndex        =   14
         Text            =   "frmCP_Purchasing.frx":347A
         Top             =   1440
         Width           =   5265
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
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   120
         Width           =   2265
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   1095
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   1110
         Width           =   5265
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company"
         Height          =   195
         Index           =   11
         Left            =   150
         TabIndex        =   31
         Top             =   810
         Width           =   660
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Brand"
         Height          =   195
         Index           =   10
         Left            =   105
         TabIndex        =   30
         Top             =   1800
         Width           =   420
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PO #"
         Height          =   195
         Index           =   9
         Left            =   6435
         TabIndex        =   28
         Top             =   1830
         Width           =   375
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Term"
         Height          =   195
         Index           =   8
         Left            =   6435
         TabIndex        =   27
         Top             =   1545
         Width           =   360
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Delivered To"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   26
         Top             =   2100
         Width           =   915
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "D. Delivery"
         Height          =   195
         Index           =   5
         Left            =   6435
         TabIndex        =   25
         Top             =   1185
         Width           =   780
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
         Height          =   495
         Left            =   7305
         TabIndex        =   17
         Tag             =   "ht0;ft0"
         Top             =   2190
         Width           =   3030
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   195
         Index           =   4
         Left            =   105
         TabIndex        =   18
         Top             =   2550
         Width           =   735
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   195
         Index           =   1
         Left            =   6435
         TabIndex        =   15
         Top             =   810
         Width           =   345
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   195
         Index           =   6
         Left            =   105
         TabIndex        =   13
         Top             =   1470
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
         TabIndex        =   10
         Top             =   150
         Width           =   915
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier"
         Height          =   195
         Index           =   2
         Left            =   105
         TabIndex        =   12
         Top             =   1155
         Width           =   570
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         Height          =   195
         Index           =   3
         Left            =   6435
         TabIndex        =   16
         Top             =   2355
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmCP_Purchasing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmCP_PO_Return"

Private WithEvents oTrans As ggcCPPurchasing.clsCPPurchasing
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
   '''On Error GoTo errProc

   If Not pbGridFocus And Index = 0 Then Call txtField_Validate(pnIndex, False)
   With GridEditor1
      Select Case Index
      Case 0
         If .Rows > 2 Then
            pnCtr = 0
            Do While pnCtr < .Rows
               If Trim(.TextMatrix(pnCtr, 1)) = "" Then
                  .Row = pnCtr
                  If oTrans.deleteDetail(.Row - 1) Then .deleteRow
               Else
                  pnCtr = pnCtr + 1
               End If
            Loop

            .ColWidth(3) = 2300
            If .Rows > 16 Then .ColWidth(3) = 2100
         End If

         If isEntryOk Then
            If oTrans.SaveTransaction Then
               MsgBox "Record Updated Successfully!!!", vbInformation, "Confirm"
               initButton xeModeReady

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
            If oTrans.searchDetail(.Row - 1, 1) Then .Col = 1
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

            .ColWidth(3) = 2300
            If .Rows > 16 Then .ColWidth(3) = 2100
         End If
      Case 3
         lsRep = MsgBox("Transaction is in Update Mode!!!" & vbCrLf & _
                        "Do you want to Cancel Transaction!!!", vbYesNo + vbQuestion, "Confirm")

         If lsRep = vbYes Then
            oTrans.NewTransaction
            clearFields
            initButton xeModeReady
         Else
            txtField(pnIndex).SetFocus
         End If
         pbSave = False
      Case 4
         oTrans.NewTransaction
         clearFields
         initButton xeModeAddNew
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
   '''On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oTrans = New ggcCPPurchasing.clsCPPurchasing
   Set oTrans.AppDriver = oApp

   oTrans.InitTransaction
   oTrans.NewTransaction

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualLeft
   
   InitGrid
   clearFields
   
   initButton xeModeAddNew

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
      End If

      If Not Cancel Then
         If .Row = .Rows - 1 Then oTrans.addDetail
      End If

      If .Rows > 16 Then .ColWidth(3) = 2100
   End With
End Sub

Private Sub GridEditor1_EditorValidate(Cancel As Boolean)
   Dim lsOldProc As String

   lsOldProc = "GridEditor1_EditorValidate"
   '''On Error GoTo errProc

   With GridEditor1
      If pbGridValidate Then
         pbGridValidate = False
         Exit Sub
      End If
      
      If .Col = 5 Then
         oTrans.Detail(.Row - 1, "nQuantity") = .TextMatrix(.Row, .Col)
      ElseIf .Col = 6 Then
         oTrans.Detail(.Row - 1, "nUnitPrce") = .TextMatrix(.Row, .Col)
      Else
         oTrans.Detail(.Row - 1, .Col) = .TextMatrix(.Row, .Col)
      End If
      If .Rows > 16 Then
         .TopRow = .Rows - 1
         .ColWidth(3) = 2100
      End If
   End With
   pbGridValidate = True
endProc:
'   GridEditor1.Refresh
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
   '''On Error GoTo errProc

   If KeyCode = vbKeyF3 Then
      With GridEditor1
         If oTrans.searchDetail(.Row - 1, .Col, .TextMatrix(.Row, .Col)) Then
            .Col = 5
         Else
            oTrans.Detail(.Row - 1, .Col) = .TextMatrix(.Row, .Col)
            .Col = 1
         End If

         .Refresh
         .SetFocus
         If .Rows > 16 Then
            .TopRow = .Rows - 1
            .ColWidth(3) = 2100
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

Private Sub oTrans_DetailRetrieved(ByVal Index As Integer, ByVal Value As Variant)
   With GridEditor1
      Select Case Index
      Case 1, 2, 3, 4
         .TextMatrix(.Row, Index) = oTrans.Detail(.Row - 1, Index)
      End Select
   End With
End Sub

Private Sub InitGrid()
   With GridEditor1
      .Rows = 2
      .Cols = 7
      .Font = "MS Sans Serif"

      'column title
      .TextMatrix(0, 1) = "BarrCode"
      .TextMatrix(0, 2) = "Description"
      .TextMatrix(0, 3) = "Brand"
      .TextMatrix(0, 4) = "Model"
      .TextMatrix(0, 5) = "Quantity"
      .TextMatrix(0, 6) = "Unit Prc"
      
      .Row = 0

      'column alignment
      For pnCtr = 0 To .Cols - 1
         .Col = pnCtr
         .CellFontBold = True
         .CellAlignment = 3
      Next

      'column width
      .ColWidth(0) = 330
      .ColWidth(1) = 1800
      .ColWidth(2) = 3000
      .ColWidth(3) = 1200
      .ColWidth(4) = 1200
      .ColWidth(5) = 850
      .ColWidth(6) = 950
      
      .ColFormat(5) = 0#
      .ColFormat(6) = "0.00"
            
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

Private Sub oTrans_MasterRetrieved(ByVal Index As Variant, ByVal Value As Variant)
   Select Case Index
   Case 4
      txtField(Index).Text = oTrans.Master(8)
   Case Else
      txtField(Index).Text = oTrans.Master(Index)
   End Select
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
   '''On Error GoTo errProc

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

Private Sub initButton(lnStat As Integer)
   Dim lbShow As Boolean

   lbShow = IIf(lnStat = 0, False, True)
   cmdButton(4).Visible = Not lbShow
   cmdButton(5).Visible = Not lbShow
   cmdButton(6).Visible = Not lbShow

   cmdButton(0).Visible = lbShow
   cmdButton(1).Visible = lbShow
   cmdButton(2).Visible = lbShow
   cmdButton(3).Visible = lbShow

   With GridEditor1
      .ColEnabled(1) = lbShow
      .ColEnabled(2) = lbShow
      .ColEnabled(7) = lbShow
   End With

   If Not lbShow Then cmdButton(4).SetFocus
End Sub

Public Function PrintTrans() As Boolean
   Dim lsOldProc As String
   Dim lnCtr As Integer
   Dim lrs As ADODB.Recordset
   Dim loRS As Recordset
   Dim lsAddress As String
   Dim lsSQL As String
   
   lsOldProc = "InitReport"

   Set lrs = New ADODB.Recordset
   With lrs
      .Fields.Append "nField01", adInteger, 3
      .Fields.Append "sField01", adVarChar, 64
      .Fields.Append "sField02", adVarChar, 64, adFldIsNullable
      .Fields.Append "sField03", adVarChar, 64, adFldIsNullable
      .Fields.Append "lField01", adCurrency
      .Open
      
      For lnCtr = 0 To oTrans.ItemCount - 1
         lrs.AddNew
         .Fields("nField01").Value = oTrans.Detail(lnCtr, "nQuantity")
         .Fields("sField01").Value = oTrans.Detail(lnCtr, "sModelNme")
         .Fields("sField02").Value = oTrans.Detail(lnCtr, "sColorNme")
         .Fields("sField03").Value = oTrans.Detail(lnCtr, "sModelCde")
         .Fields("lField01").Value = oTrans.Detail(lnCtr, "nUnitPrce")
      Next
   End With
   
   lsSQL = "SELECT CONCAT(a.sAddressx, ', ', b.sTownName, ' ', c.sProvName)" & _
                     " FROM Branch a, TownCity b, Province c" & _
                     " WHERE a.sBranchNm = " & strParm(txtField(3)) & _
                        " AND a.sTownIDxx = b.sTownIDxx AND b.sProvIDxx = c.sProvIDxx"
   Set loRS = New Recordset
   
   With loRS
      .Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
      If Not .EOF Then
         lsAddress = loRS(0)
      Else
         lsAddress = ""
      End If
   End With
   
   Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\MP-Purchase4Supplier.rpt")
    
   With oReport
      'assign important info to the report
      .DiscardSavedData
      .FieldMappingType = crAutoFieldMapping
      .Database.SetDataSource lrs
      
    .Sections("PH").ReportObjects("txtSupplier").SetText txtField(2)
    .Sections("PH").ReportObjects("txtDeliverTo").SetText txtField(4)
    .Sections("PH").ReportObjects("txtTerm").SetText txtField(9)
    .Sections("PH").ReportObjects("txtDateTransact").SetText txtField(1)
    .Sections("PH").ReportObjects("txtInvoiceTo").SetText oTrans.Master("sCompnyNm")
    .Sections("PH").ReportObjects("txtPONo").SetText oTrans.Master("sReferNox")
    .Sections("PF").ReportObjects("txtRequested").SetText "JULIE MARTINEZ" 'oApp.UserName
 
      oReport.PrintOutEx False, 1
   End With

   If oTrans.Master("cTranStat") = xeStateOpen Then
      PrintTrans = oTrans.CloseTransaction
   Else
      PrintTrans = True
   End If
End Function

Private Sub clearFields()
   For pnCtr = 0 To txtField.Count - 1
      Select Case pnCtr
      Case 0
         txtField(pnCtr).Text = Format(oTrans.Master(pnCtr), "@@@@-@@@@@@")
      Case 1, 6
         txtField(pnCtr).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
      Case 2, 3, 4, 5
         txtField(pnCtr).Text = ""
      End Select
   Next
   
   txtField(12).Text = ""
   txtField(13).Text = ""
   txtField(24).Text = ""
   
   
   With GridEditor1
      .Rows = 2
      .ColWidth(3) = 2300

      'empty row
      .TextMatrix(1, 1) = ""
      .TextMatrix(1, 2) = ""
      .TextMatrix(1, 3) = ""
      .TextMatrix(1, 4) = ""
      .TextMatrix(1, 5) = "0"
      .TextMatrix(1, 6) = "0.00"
      
      .ColFormat(6) = "0.00"
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
      Case 1, 6
         If Not IsDate(.Text) Then .Text = oApp.ServerDate
         .Text = Format(.Text, "MMMM DD, YYYY")
      Case 3
         .Text = Format(.Text, ">")
      End Select
      
      If Index = 6 Then
         oTrans.Master(11) = txtField(Index)
      ElseIf Index = 13 Then
         oTrans.Master(14) = txtField(Index)
      Else
         oTrans.Master(Index) = txtField(Index)
      End If
   End With
End Sub

Private Function isEntryOk() As Boolean
   If txtField(2).Text = "" Then
      MsgBox "Supplier not found!!!" & vbCrLf & _
             "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
      txtField(2).SetFocus
      GoTo EntryNotOK
   End If
   
   If txtField(12).Text = "" Then
      MsgBox "Invalid Term info!!!" & vbCrLf & _
             "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
      txtField(12).SetFocus
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
   isEntryOk = True
   Exit Function
EntryNotOK:
   isEntryOk = False
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


