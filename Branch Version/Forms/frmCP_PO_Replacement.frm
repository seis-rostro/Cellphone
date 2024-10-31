VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmCP_PO_Replacement 
   BorderStyle     =   0  'None
   Caption         =   "PO Replacement"
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
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3795
      Left            =   5130
      TabIndex        =   24
      Top             =   2955
      Width           =   6465
      _ExtentX        =   11404
      _ExtentY        =   6694
      _Version        =   393216
      FocusRect       =   0
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   3780
      Left            =   1605
      Tag             =   "wt0;fb0"
      Top             =   2970
      Width           =   3480
      _ExtentX        =   6138
      _ExtentY        =   6668
      BackColor       =   12632256
      ClipControls    =   0   'False
      BorderStyle     =   3
      Begin VB.TextBox txtFieldDetail 
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   3
         Left            =   1230
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   1320
         Width           =   1260
      End
      Begin VB.TextBox txtFieldDetail 
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   2
         Left            =   1230
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   1710
         Width           =   1260
      End
      Begin VB.TextBox txtFieldDetail 
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   1
         Left            =   1230
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   930
         Width           =   2130
      End
      Begin VB.TextBox txtFieldDetail 
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   0
         Left            =   1230
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   540
         Width           =   2130
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Price"
         Height          =   195
         Index           =   9
         Left            =   195
         TabIndex        =   28
         Top             =   1395
         Width           =   690
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QTY"
         Height          =   195
         Index           =   8
         Left            =   180
         TabIndex        =   27
         Top             =   1785
         Width           =   330
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   26
         Top             =   1005
         Width           =   795
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Code/IMEI"
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   25
         Top             =   615
         Width           =   780
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   90
      TabIndex        =   19
      Top             =   1815
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
      Picture         =   "frmCP_PO_Replacement.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   90
      TabIndex        =   17
      Top             =   555
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
      Picture         =   "frmCP_PO_Replacement.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   90
      TabIndex        =   18
      Top             =   1185
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
      Picture         =   "frmCP_PO_Replacement.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   90
      TabIndex        =   22
      Top             =   2445
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
      Picture         =   "frmCP_PO_Replacement.frx":166E
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   90
      TabIndex        =   20
      Top             =   1185
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
      Picture         =   "frmCP_PO_Replacement.frx":1DE8
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   6
      Left            =   90
      TabIndex        =   23
      Top             =   2445
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
      Picture         =   "frmCP_PO_Replacement.frx":2562
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   90
      TabIndex        =   21
      Top             =   1815
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
      Picture         =   "frmCP_PO_Replacement.frx":2CDC
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
         Height          =   315
         Index           =   6
         Left            =   6660
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1095
         Width           =   3090
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   660
         Index           =   4
         Left            =   1230
         MultiLine       =   -1  'True
         TabIndex        =   3
         Text            =   "frmCP_PO_Replacement.frx":3456
         Top             =   1575
         Width           =   4545
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   6660
         TabIndex        =   1
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
         TabIndex        =   12
         Text            =   "frmCP_PO_Replacement.frx":345E
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
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   120
         Width           =   2265
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   1230
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   765
         Width           =   4545
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Refer #"
         Height          =   195
         Index           =   10
         Left            =   6045
         TabIndex        =   29
         Top             =   1140
         Width           =   540
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
         TabIndex        =   15
         Tag             =   "ht0;ft0"
         Top             =   1440
         Width           =   3090
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   195
         Index           =   4
         Left            =   315
         TabIndex        =   16
         Top             =   1605
         Width           =   735
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   195
         Index           =   1
         Left            =   6045
         TabIndex        =   13
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
         TabIndex        =   11
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
         TabIndex        =   8
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
         TabIndex        =   10
         Top             =   810
         Width           =   570
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         Height          =   195
         Index           =   3
         Left            =   6045
         TabIndex        =   14
         Top             =   1470
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmCP_PO_Replacement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmCP_PO_Replacement"

Private WithEvents oTrans As ggcCPPurchasing.clsCPPOReplacement
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin

Dim pnIndex As Integer
Dim pbGridGotFocus As Boolean
Dim pnCtr As Integer
Dim pbSave As Boolean
Dim pbGridValidate As Boolean
Dim pbPosted As Boolean
Dim pbForm As Boolean
Dim pnRow As Integer

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   Dim lsRep As String

   lsOldProc = "cmdButton_Click"
   ''On Error GoTo errProc

   If Not pbGridGotFocus And Index = 0 Then Call txtField_Validate(pnIndex, False)
   
   With MSFlexGrid1
      Select Case Index
      Case 0
         If .Rows > 2 Then
            pnCtr = 0
            Do While pnCtr < .Rows
               If Trim(.TextMatrix(pnCtr, 1)) = "" Then
                  .Row = pnCtr
                  If oTrans.deleteDetail(.Row - 1) Then deleteRow
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
         If pbGridGotFocus Then
            If oTrans.searchDetail(.Row - 1, 1) Then .Col = 1
            .Refresh
            .SetFocus
         Else
            oTrans.SearchMaster pnIndex
         End If
      Case 2
         If .Rows > 2 Then
            If oTrans.deleteDetail(.Row - 1) Then deleteRow

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
            ClearFields
            initButton xeModeReady
         Else
            txtField(pnIndex).SetFocus
         End If
         pbSave = False
      Case 4
         oTrans.NewTransaction
         ClearFields
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
   MSFlexGrid1.Refresh

   oApp.MenuName = Me.Tag
   Me.ZOrder 0
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String

   lsOldProc = "Form_Load"
   ''On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oTrans = New ggcCPPurchasing.clsCPPOReplacement
   Set oTrans.AppDriver = oApp

   oTrans.InitTransaction
   oTrans.NewTransaction

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualLeft

   InitGrid
   ClearFields
   initButton xeModeAddNew

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

Private Sub MSFlexGrid1_GotFocus()
   pbGridGotFocus = True
End Sub

Private Sub MSFlexGrid1_RowColChange()
   pnRow = MSFlexGrid1.Row
End Sub

Private Sub MSFlexGrid1_SelChange()
   With MSFlexGrid1
      txtFieldDetail(0) = .TextMatrix(pnRow, 0 + 1)
      txtFieldDetail(1) = .TextMatrix(pnRow, 0 + 2)
      txtFieldDetail(2) = .TextMatrix(pnRow, 0 + 6)
      txtFieldDetail(3) = .TextMatrix(pnRow, 0 + 4)
   End With
End Sub

Private Sub oTrans_DetailRetrieved(ByVal Index As Integer)
   With MSFlexGrid1
      Select Case Index
      Case 1
         txtFieldDetail(0) = oTrans.Detail(pnRow - 1, "xReferNox")
      Case 2
         txtFieldDetail(1) = oTrans.Detail(pnRow - 1, "sDescript")
      Case 4
         txtFieldDetail(3) = oTrans.Detail(pnRow - 1, "nUnitPrce")
      End Select
      .TextMatrix(pnRow, 3) = IFNull(oTrans.Detail(pnRow - 1, "sModelNme"), "No Model")
      .TextMatrix(pnRow, 5) = oTrans.Detail(pnRow - 1, "sReferNox")
   End With
   
End Sub

Private Sub oTrans_MasterRetrieved(ByVal Index As Integer)
   txtField(Index).Text = oTrans.Master(Index)
End Sub

Private Sub InitGrid()
   With MSFlexGrid1
      .Rows = 2
      .Cols = 7
      .Font = "MS Sans Serif"

      'column title
      .TextMatrix(0, 1) = "BarrCode"
      .TextMatrix(0, 2) = "Description"
      .TextMatrix(0, 3) = "Model"
      .TextMatrix(0, 4) = "Unit Price"
      .TextMatrix(0, 5) = "ReferNo"
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
      .ColWidth(2) = 2000
      .ColWidth(4) = 1020
      .ColWidth(5) = 1300
      .ColWidth(6) = 500

      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      .ColAlignment(4) = 6
      .ColAlignment(5) = 1
      .ColAlignment(6) = 6

      .Row = 1
      .Col = 1
      .ColSel = .Cols - 1
   End With
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      If Index = 1 Then .Text = Format(.Text, "MM/DD/YY")

      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = oApp.getColor("HT1")
   End With

   pbGridGotFocus = False
   pnIndex = Index
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String

   lsOldProc = "txtField_KeyDown"
   ''On Error GoTo errProc

   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
      With txtField(Index)
         If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
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
         If GetFocus = MSFlexGrid1.hwnd Then Exit Sub
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

   For pnCtr = 1 To txtField.Count - 3
      txtField(pnCtr).Enabled = lbShow
   Next

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
   ''On Error GoTo errProc

   PrintTrans = True
   Set lrs = New ADODB.Recordset
   lrs.Fields.Append "lField01", adCurrency
   lrs.Fields.Append "nField01", adInteger, 3
   lrs.Fields.Append "nField02", adChar, 1
   lrs.Fields.Append "sField01", adVarChar, 20
   lrs.Fields.Append "sField02", adVarChar, 128
   lrs.Fields.Append "sField03", adVarChar, 20
   lrs.Fields.Append "sField04", adVarChar, 128
   lrs.Fields.Append "sField05", adVarChar, 100
   lrs.Fields.Append "sField06", adVarChar, 30
   lrs.Fields.Append "sField07", adVarChar, 25
   lrs.Fields.Append "sField08", adVarChar, 25
   lrs.Open

   With oTrans
      For lnCtr = 0 To .ItemCount - 1
         lrs.AddNew
         lrs.Fields("lField01") = 0#  'oTrans.Detail(lnCtr, "nUnitPrce") 'she 2017-12-22 validation of unit price is accounting
         lrs.Fields("nField01") = oTrans.Detail(lnCtr, "nQuantity")
         lrs.Fields("nField02") = oTrans.Detail(lnCtr, "cHsSerial")
         lrs.Fields("sField01") = oTrans.Detail(lnCtr, "sTransNox")
         lrs.Fields("sField02") = oTrans.Detail(lnCtr, "sStockIDx")
         lrs.Fields("sField03") = oTrans.Detail(lnCtr, "sBarrCode")
         lrs.Fields("sField04") = oTrans.Detail(lnCtr, "sDescript")
         lrs.Fields("sField05") = oTrans.Detail(lnCtr, "sBrandNme")
         lrs.Fields("sField06") = IFNull(oTrans.Detail(lnCtr, "sSerialNo"), "")
         lrs.Fields("sField07") = IFNull(oTrans.Detail(lnCtr, "sReferNox"), "")
         lrs.Fields("sField08") = IFNull(oTrans.Master("sReferNox"), "")
         
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
            & " WHERE a.sBranchCd = " & strParm(Left(oTrans.Master("sTransNox"), Len(oApp.BranchCode))) _
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
            & " WHERE a.sBranchCd = " & strParm(Left(oTrans.Master("sTransNox"), Len(oApp.BranchCode))) _
               & " AND a.sTownIDxx = b.sTownIDxx" _
               & " AND b.sProvIDxx = c.sProvIDxx" _
            , oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText

   Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\CPPurchaseReplacement.rpt")
   'assign important info to the report
   oReport.DiscardSavedData
   oReport.FieldMappingType = crAutoFieldMapping
   oReport.Database.SetDataSource lrs

   oReport.Sections("PHa").ReportObjects("txtTransNox").SetText "CP-" & Right(oTrans.Master("sTransNox"), 10)
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

Private Sub ClearFields()
   For pnCtr = 0 To txtField.Count - 1
      Select Case pnCtr
      Case 0
         txtField(pnCtr).Text = Format(oTrans.Master(pnCtr), "@@@@-@@@@@@")
      Case 1
         txtField(pnCtr).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
      Case 2, 3, 4
         txtField(pnCtr).Text = Empty
      End Select
   Next
   
      txtField(6).Text = ""
      
  For pnCtr = 0 To txtFieldDetail.Count - 1
      Select Case pnCtr
      Case 0, 1, 2
         txtFieldDetail(pnCtr).Text = ""
      Case 3
         txtFieldDetail(pnCtr).Text = "0.00"
      End Select
   Next

   With MSFlexGrid1
      .Rows = 2
      .ColWidth(3) = 2300

      'empty row
      .TextMatrix(1, 1) = ""
      .TextMatrix(1, 2) = ""
      .TextMatrix(1, 3) = ""
      .TextMatrix(1, 4) = "0.00"
      .TextMatrix(1, 5) = ""
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

Private Function isEntryOk() As Boolean
   If txtField(2).Text = "" Then
      MsgBox "Supplier not found!!!" & vbCrLf & _
             "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
      txtField(2).SetFocus
      GoTo EntryNotOK
   End If
   
   If txtField(6).Text = "" Then
      MsgBox "Refer No. not found!!!" & vbCrLf & _
             "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
      txtField(6).SetFocus
      GoTo EntryNotOK
   End If
   
   With MSFlexGrid1
      If Trim(.TextMatrix(1, 1)) = "" Then  'Or Trim(.TextMatrix(1, 4)) = 0# And oApp.UserLevel < xeEngineer
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

Private Sub txtFieldDetail_Change(Index As Integer)
   With MSFlexGrid1
      Select Case Index
      Case 0
         .TextMatrix(pnRow, 1) = txtFieldDetail(Index)
      Case 1
         .TextMatrix(pnRow, 2) = txtFieldDetail(Index)
      Case 2
         .TextMatrix(pnRow, 6) = txtFieldDetail(Index)
      Case 3
         .TextMatrix(pnRow, 4) = txtFieldDetail(Index)
      End Select
               
   End With
End Sub

Private Sub txtFieldDetail_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

   With MSFlexGrid1
      Select Case Index
      Case 0, 1
         If KeyCode = vbKeyF3 Then
            oTrans.searchDetail .Row - 1, Index + 1, txtFieldDetail(Index)
            If txtFieldDetail(Index) <> "" Then SetNextFocus
         Else
            If txtFieldDetail(Index) <> "" Then oTrans.searchDetail Index + 1, txtFieldDetail(Index)
         End If
      Case 2
         If KeyCode = vbKeyReturn Then
            Call txtFieldDetail_Validate(Index, False)
            If Trim(oTrans.Detail(pnRow - 1, "xReferNox")) <> "" And _
                  Trim(oTrans.Detail(pnRow - 1, "sDescript")) <> "" And _
                  oTrans.Detail(pnRow - 1, "nUnitPrce") <> 0# And _
                  oTrans.Detail(pnRow - 1, "nQuantity") > 0 Then
               Call addDetail
            Else
               oTrans.Detail(pnRow - 1, "nQuantity") = 0
               txtFieldDetail(Index) = 0
            End If
         End If
      End Select
   End With
   
End Sub

Private Sub txtFieldDetail_Validate(Index As Integer, Cancel As Boolean)
   With txtFieldDetail(Index)
      Select Case Index
      Case 2
         If Not IsNumeric(.Text) Then .Text = 0
            oTrans.Detail(pnRow - 1, "nQuantity") = CDbl(.Text)
'         Call computeTotalDep
      Case 3
         If Not IsNumeric(.Text) Then .Text = 0#
            .Text = Format(.Text, "#,##0.00")
            If .Text = 0# And oApp.UserLevel < xeEngineer Then
               MsgBox "Invalid Unit Price Detected", vbInformation, "INFO"
               Exit Sub
               .SetFocus
            End If
         oTrans.Detail(pnRow - 1, "nUnitPrce") = CDbl(.Text)
'         Call computeTotalDep
      Case Else
         .Text = TitleCase(.Text)
         oTrans.Detail(pnRow - 1, Index + 1) = .Text
      End Select
   End With
End Sub

Private Sub addDetail()
Dim lsOldProc As String

   lsOldProc = pxeMODULENAME & "addDetail"
   ''On Error GoTo errProc

   With MSFlexGrid1
      If oTrans.addDetail Then
         .Rows = .Rows + 1
         pnRow = .Rows - 2
         
         txtFieldDetail(0) = oTrans.Detail(pnRow - 1, "xReferNox")
         txtFieldDetail(1) = oTrans.Detail(pnRow - 1, "sDescript")
         txtFieldDetail(2) = Format(oTrans.Detail(pnRow - 1, "nQuantity"), "#,##0")
         txtFieldDetail(3) = Format(oTrans.Detail(pnRow - 1, "nUnitPrce"), "#,##0.0000")
         
         .Row = .Rows - 1
         .ColSel = .Cols - 1
         txtFieldDetail(0).SetFocus
      End If
      
      .Refresh
   End With
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub deleteRow()
   Dim lnLastRow As Integer
   Dim lnCtr As Integer

   With MSFlexGrid1
      .Rows = .Rows - 1

      lnLastRow = .Rows - 1
      For pnCtr = lnLastRow To oTrans.ItemCount - 1
         .TextMatrix(pnCtr + 1, 0) = pnCtr + 1
         .TextMatrix(pnCtr + 1, 1) = oTrans.Detail(pnCtr, "xReferNox")
         .TextMatrix(pnCtr + 1, 2) = oTrans.Detail(pnCtr, "sDescript")
         .TextMatrix(pnCtr + 1, 3) = oTrans.Detail(pnCtr, "sModelNme")
         .TextMatrix(pnCtr + 1, 4) = oTrans.Detail(pnCtr, "sReferNox")
         .TextMatrix(pnCtr + 1, 5) = oTrans.Detail(pnCtr, "nUnitPrce")
         .TextMatrix(pnCtr + 1, 6) = oTrans.Detail(pnCtr, "nQuantity")
         

         .Row = pnCtr + 1
         If (pnCtr + 1) Mod 2 = 0 Then
            For lnCtr = 1 To .Cols - 1
               .Col = lnCtr
               .CellBackColor = oApp.getColor("fb0")
            Next
         End If
      Next

'      pnRow = lnLastRow - 1
'      .Row = pnRow - 1
      pnRow = pnRow - 1
      .Col = 1
      .ColSel = .Cols - 1
   End With
End Sub
