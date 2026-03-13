VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0"; "xrGridControl.ocx"
Begin VB.Form frmCP_PO_Receiving 
   BorderStyle     =   0  'None
   Caption         =   "Delivery Acceptance"
   ClientHeight    =   8085
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11940
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   11940
   ShowInTaskbar   =   0   'False
   Begin xrGridEditor.GridEditor GridEditor1 
      Height          =   3900
      Left            =   1575
      TabIndex        =   24
      Top             =   4155
      Width           =   10245
      _ExtentX        =   18071
      _ExtentY        =   6879
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
      Object.HEIGHT          =   3900
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
      MOUSEICON       =   "frmCP_PO_Receiving.frx":0000
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
      TabIndex        =   27
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
      Picture         =   "frmCP_PO_Receiving.frx":001C
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   3540
      Index           =   0
      Left            =   1575
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   10245
      _ExtentX        =   18071
      _ExtentY        =   6244
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   12
         Left            =   7545
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   3045
         Width           =   2505
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   11
         Left            =   7545
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   2715
         Width           =   2505
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   10
         Left            =   7545
         TabIndex        =   35
         Text            =   "Text1"
         Top             =   1395
         Width           =   2505
      End
      Begin VB.CheckBox chkField 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Backload Unit"
         Height          =   195
         Index           =   1
         Left            =   6285
         TabIndex        =   34
         Tag             =   "et0;fb0"
         Top             =   240
         Width           =   1365
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   9
         Left            =   1365
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   735
         Width           =   2820
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   8
         Left            =   7545
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   1725
         Width           =   2505
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   480
         Index           =   3
         Left            =   1365
         MultiLine       =   -1  'True
         TabIndex        =   7
         Text            =   "frmCP_PO_Receiving.frx":0796
         Top             =   1395
         Width           =   4770
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   5
         Left            =   7545
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   2385
         Width           =   2505
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   7545
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   735
         Width           =   2505
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   1365
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   1065
         Width           =   4770
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   1455
         Index           =   7
         Left            =   1365
         MultiLine       =   -1  'True
         TabIndex        =   9
         Text            =   "frmCP_PO_Receiving.frx":079E
         Top             =   1890
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
         Index           =   6
         Left            =   7545
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   1065
         Width           =   2505
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   4
         Left            =   7545
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   2055
         Width           =   2505
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "ADDTL. DISC."
         Height          =   285
         Index           =   12
         Left            =   6315
         TabIndex        =   22
         Top             =   3105
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "DISCOUNT"
         Height          =   285
         Index           =   11
         Left            =   6315
         TabIndex        =   20
         Top             =   2775
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Order No"
         Height          =   285
         Index           =   10
         Left            =   6300
         TabIndex        =   36
         Top             =   1425
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Company"
         Height          =   285
         Index           =   8
         Left            =   165
         TabIndex        =   2
         Top             =   750
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Refer. Date"
         Height          =   285
         Index           =   5
         Left            =   6300
         TabIndex        =   14
         Top             =   1785
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   285
         Index           =   4
         Left            =   165
         TabIndex        =   6
         Top             =   1425
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Inv. No."
         Height          =   285
         Index           =   9
         Left            =   6300
         TabIndex        =   18
         Top             =   2400
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transact. Date"
         Height          =   285
         Index           =   1
         Left            =   6300
         TabIndex        =   10
         Top             =   780
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Reference No"
         Height          =   285
         Index           =   3
         Left            =   6300
         TabIndex        =   16
         Top             =   2085
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   285
         Index           =   7
         Left            =   165
         TabIndex        =   8
         Top             =   1920
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
         TabIndex        =   12
         Top             =   1095
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier"
         Height          =   285
         Index           =   2
         Left            =   165
         TabIndex        =   4
         Top             =   1110
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
      TabIndex        =   25
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
      Picture         =   "frmCP_PO_Receiving.frx":07A6
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   90
      TabIndex        =   26
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
      Picture         =   "frmCP_PO_Receiving.frx":0F20
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   90
      TabIndex        =   32
      Top             =   4335
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
      Picture         =   "frmCP_PO_Receiving.frx":169A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   90
      TabIndex        =   28
      Top             =   1815
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
      Picture         =   "frmCP_PO_Receiving.frx":1E14
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   6
      Left            =   90
      TabIndex        =   33
      Top             =   4335
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
      Picture         =   "frmCP_PO_Receiving.frx":258E
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   600
      Index           =   5
      Left            =   90
      TabIndex        =   29
      Top             =   2445
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "Seria&l"
      AccessKey       =   "l"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCP_PO_Receiving.frx":2D08
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   600
      Index           =   7
      Left            =   90
      TabIndex        =   30
      Top             =   3075
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "Free U."
      AccessKey       =   "Free U."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCP_PO_Receiving.frx":3482
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   600
      Index           =   8
      Left            =   90
      TabIndex        =   31
      Top             =   3705
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "Demo U."
      AccessKey       =   "Demo U."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCP_PO_Receiving.frx":3BFC
   End
End
Attribute VB_Name = "frmCP_PO_Receiving"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmDeliverAcceptance"

Private WithEvents oTrans As clsCPPOReceiving
Attribute oTrans.VB_VarHelpID = -1
Private oFormSerialNew As frmDASerialNew
Private oFormSerialBackload As frmDASerialBackload
Private oSkin As clsFormSkin

Dim pnIndex As Integer
Dim pbGridFocus As Boolean
Dim pnCtr As Integer

Private Sub chkField_Click(Index As Integer)
   If Index = 1 Then oTrans.BackLoad = chkField(Index).Value
End Sub

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
            pnCtr = 0
            Do While pnCtr < .Rows
               If .TextMatrix(pnCtr, 6) = 0 Then
                  .Row = pnCtr
                  If oTrans.deleteDetail(.Row - 1) Then .deleteRow
               Else
                  pnCtr = pnCtr + 1
               End If
            Loop

            .ColWidth(3) = 3100
            If .Rows > 16 Then .ColWidth(3) = 2850
         End If

         If isEntryOk Then
            If oTrans.SaveTransaction Then
               MsgBox "Record Updated Successfully!!!", vbInformation, "Notice"
               If oTrans.OpenTransaction(oTrans.Master("sTransNox")) Then initButton xeModeReady
            Else
               MsgBox "Unable to Update Record!!!", vbCritical, "Warning"
            End If
         End If
      Case 1
         If pbGridFocus Then
            If oTrans.searchDetail(.Row - 1, 1) Then
               .Col = 3
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

            .ColWidth(3) = 3100
            If .Rows > 16 Then .ColWidth(3) = 2850
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
         If pbGridFocus Then
            If .TextMatrix(.Row, 3) <> 0 Then
               If chkField(1).Value = xeYes Then
                  AcceptSerialBackload
               Else
                  AcceptSerialNew
               End If
               .Col = 1
               .SetFocus
            End If
         End If
      Case 6
         Unload Me
      Case 7 'Free Units
         oTrans.GetFreeSerial
      Case 8 'Demo Units
         oTrans.GetDemoSerial
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

   Set oFormSerialNew = New frmDASerialNew
   Set oFormSerialBackload = New frmDASerialBackload
   Set oTrans = New clsCPPOReceiving
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

   For pnCtr = 1 To txtField.Count - 1
      txtField(pnCtr).MaxLength = oTrans.MasFldSize(pnCtr)
   Next

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oFormSerialNew = Nothing
   Set oFormSerialBackload = Nothing
   Set oTrans = Nothing
   Set oSkin = Nothing
End Sub

Private Sub GridEditor1_AddingRow(Cancel As Boolean)
   Dim lsUserID As String
   Dim lsUserName As String
   Dim lnUserRights As Integer
   Dim lasRights() As String

   With GridEditor1
      If .TextMatrix(.Row, 1) = "" Then
         Cancel = True
      ElseIf CDbl(.TextMatrix(.Row, 4)) = 0# Then
      'she 2015-11-12
      ' get approval for giveaways with 0.00 unit price
         lasRights = Split(oApp.mdiMain.Controls(oApp.MenuName).Tag, "»")
         If GetApproval(oApp, lnUserRights, lsUserID, lsUserName, oApp.MenuName) = False Then 'lasRights(0)
            Cancel = True
            Exit Sub
         End If

         If lnUserRights < xeManager Then
            MsgBox "Approving Officer Has no Right to Cancel this transaction!!!" & vbCrLf & _
               "Request can not be granted!!!", vbCritical, "Warning"
            Cancel = True
            Exit Sub
         End If
      ElseIf CDbl(.TextMatrix(.Row, 4)) = 0# Then
         Cancel = True
      ElseIf .TextMatrix(.Row, 6) <= 0 Then
         Cancel = True
      End If
      If Not Cancel Then oTrans.addDetail

      If .Rows > 16 Then .ColWidth(3) = 2850

'      MsgBox oTrans.Detail(.Row - 1, "cUnitType")
   End With
End Sub

Private Sub GridEditor1_EditorValidate(Cancel As Boolean)
   Dim lsOldProc As String

   lsOldProc = "GridEditor1_EditorValidate"
   'On Error GoTo errProc

   With GridEditor1
      If .Col = 6 Then
         If oTrans.Detail(.Row - 1, "cHsSerial") = xeYes Then
            If CLng(.TextMatrix(.Row, .Col)) <> CLng(oTrans.Detail(.Row - 1, .Col)) Then
               If .TextMatrix(.Row, 1) <> Empty Then
                  oTrans.Detail(.Row - 1, .Col) = .TextMatrix(.Row, .Col)
                  If chkField(1).Value = xeYes Then
                     Cancel = AcceptSerialBackload
                  Else
                     Cancel = AcceptSerialNew
                  End If

                  If Not Cancel Then
                     .TextMatrix(.Row, .Col) = 0
                     oTrans.Detail(.Row - 1, .Col) = CDbl(.TextMatrix(.Row, .Col))
                     GoTo endProc
                  End If
               End If
            End If
         End If
      End If

      If .Col = 4 Or .Col = 6 Then
         oTrans.Detail(.Row - 1, .Col) = CDbl(.TextMatrix(.Row, .Col))
      Else
         oTrans.Detail(.Row - 1, .Col) = .TextMatrix(.Row, .Col)
         If .Col = 1 Or .Col = 2 Then
            If .TextMatrix(.Row, .Col) <> "" Then .Col = 2
         End If
      End If
      
'      With GridEditor1
         If .Col = 7 Then
            If .TextMatrix(.Row, 7) = "LDU" Then
                oTrans.Detail(.Row - 1, "cUnitType") = 0
            ElseIf .TextMatrix(.Row, 7) = "Regular" Then
                oTrans.Detail(.Row - 1, "cUnitType") = 1
            ElseIf .TextMatrix(.Row, 7) = "free" Then
                oTrans.Detail(.Row - 1, "cUnitType") = 2
            ElseIf .TextMatrix(.Row, 7) = "Service" Then
                oTrans.Detail(.Row - 1, "cUnitType") = 4
            ElseIf .TextMatrix(.Row, 7) = "RDU" Then
                oTrans.Detail(.Row - 1, "cUnitType") = 5
            Else
               .TextMatrix(.Row, 7) = "Regular"
                oTrans.Detail(.Row - 1, 7) = 1
            End If
'         MsgBox oTrans.Detail(.Row - 1, "cUnitType")
         ElseIf oTrans.searchDetail(.Row - 1, .Col, .TextMatrix(.Row, .Col)) Then
            .Col = 4
         Else
            .Col = 1
         End If
         .Refresh
         .SetFocus
'         KeyCode = 0
'      End With
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
   
   With GridEditor1
      If KeyCode = 0 Or KeyCode = vbKeyF3 Then
         Select Case .Col
         Case 1, 2
            If .TextMatrix(.Row, .Col) <> "" Then
               If oTrans.searchDetail(.Row - 1, .Col, .TextMatrix(.Row, .Col)) Then
                  .Col = 4
               End If
            End If
         Case 6
            If .TextMatrix(.Row, .Col) > 0 Then
               .Col = 7
            End If
         Case Else
            .Col = 1
         End Select
      End If
   End With
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
      Case 4, 6
         oTrans.Detail(.Row - 1, .Col) = CDbl(.TextMatrix(.Row, .Col))
      Case Else
         oTrans.Detail(.Row - 1, .Col) = .TextMatrix(.Row, .Col)
      End Select
   End With
End Sub

Private Sub oTrans_DetailRetrieved(ByVal Index As Integer)
   With GridEditor1
      .TextMatrix(.Row, Index) = oTrans.Detail(.Row - 1, Index)
      .TextMatrix(.Row, 7) = IIf(oTrans.Detail(.Row - 1, "cUnitType") = 1, "Regular", "LDU")
   End With
End Sub

Private Sub oTrans_MasterRetrieved(ByVal Index As Integer)
  If Index = 26 Then
      txtField(10).Text = oTrans.Master("sOrderNox")
  Else
      txtField(Index).Text = oTrans.Master(Index)
  End If
End Sub

Private Sub InitGrid()
   With GridEditor1
      .Rows = 2
      .Cols = 8
      .Font = "MS Sans Serif"

      'column title
      .TextMatrix(0, 1) = "BarrCode"
      .TextMatrix(0, 2) = "Description"
      .TextMatrix(0, 3) = "Model"
      .TextMatrix(0, 4) = "Unit Price"
      .TextMatrix(0, 5) = "QOH"
      .TextMatrix(0, 6) = "Qty"
      .TextMatrix(0, 7) = "U.Type"
    
      .Row = 0

      'column alignment
      For pnCtr = 0 To .Cols - 1
         .Col = pnCtr
         .CellFontBold = True
         .CellAlignment = 3
      Next

      'column width
      .ColWidth(0) = 430
      .ColWidth(1) = 2200
      .ColWidth(2) = 2540
      .ColWidth(3) = 1000
      .ColWidth(4) = 800
      .ColWidth(5) = 500
      .ColWidth(6) = 500

      .ColFormat(4) = "#,##0.0000"
      .ColFormat(5) = "#,##0"
      .ColFormat(6) = "#,##0"
      .ColNumberOnly(6) = True
      .ColDefault(4) = 0#
      .ColDefault(5) = 0
      .ColDefault(6) = 0
      .ColDefault(7) = "#,##0.00"
      

      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      .ColAlignment(4) = 6
      .ColAlignment(5) = 6
      .ColAlignment(6) = 6

      .ColEnabled(3) = False
      .ColEnabled(5) = False

      .EditorBackColor = oApp.getColor("HT1")

      .Row = 1
      .Col = 1
   End With
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      Select Case Index
      Case 1, 8
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
   cmdButton(6).Visible = Not lbShow

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
      .ColEnabled(4) = lbShow
      .ColEnabled(6) = lbShow
   End With

   If Not lbShow Then cmdButton(4).SetFocus
End Sub

Private Function AcceptSerialNew() As Boolean
   Dim lnRow As Long

   With GridEditor1
      lnRow = .Row
      If .TextMatrix(lnRow, 1) = "" Or .TextMatrix(lnRow, 6) = 0 Then Exit Function
      oFormSerialNew.GridEditor1.Rows = .TextMatrix(lnRow, 6) + 1
   End With

   With oFormSerialNew
      .InitGrid1
      .EntryNo = False
      .EditMode = xeModeAddNew

      For pnCtr = 1 To GridEditor1.TextMatrix(lnRow, 6)
         .GridEditor1.TextMatrix(pnCtr, 1) = oTrans.Serial(lnRow - 1, oTrans.Detail(lnRow - 1, "sStockIDx"), pnCtr - 1, 1)
      Next

      .Show 1

      If .Cancel = 0 Then
         For pnCtr = 1 To GridEditor1.TextMatrix(lnRow, 6)
            oTrans.Serial(lnRow - 1, oTrans.Detail(lnRow - 1, "sStockIDx"), pnCtr - 1, 1) = .GridEditor1.TextMatrix(pnCtr, 1)
         Next
      End If
      AcceptSerialNew = .Cancel = 0
   End With
End Function

Private Function AcceptSerialBackload() As Boolean
   Dim lnRow As Long

   If oTrans.Master("sSupplier") = "" Then
      MsgBox "Supplier is required!!!" & vbCrLf & _
               "Please verify your entry then try again!!!", vbCritical, "Warning"
      Exit Function
   End If

   With GridEditor1
      lnRow = .Row
      If .TextMatrix(lnRow, 1) = "" Or .TextMatrix(lnRow, 6) = 0 Then Exit Function
      oFormSerialBackload.GridEditor1.Rows = .TextMatrix(lnRow, 6) + 1
   End With

   With oFormSerialBackload
      .InitGrid1
      .EntryNo = False
      .Supplier = oTrans.Master("sSupplier")

      For pnCtr = 1 To GridEditor1.TextMatrix(lnRow, 6)
         .GridEditor1.TextMatrix(pnCtr, 1) = oTrans.Serial(lnRow - 1, oTrans.Detail(lnRow - 1, "sStockIDx"), pnCtr - 1, 1)
         .GridEditor1.TextMatrix(pnCtr, 2) = oTrans.Serial(lnRow - 1, oTrans.Detail(lnRow - 1, "sStockIDx"), pnCtr - 1, 2)
      Next

      .Show 1

      If .Cancel = 0 Then
         For pnCtr = 1 To GridEditor1.TextMatrix(lnRow, 6)
            oTrans.Serial(lnRow - 1, oTrans.Detail(lnRow - 1, "sStockIDx"), pnCtr - 1, 1) = .GridEditor1.TextMatrix(pnCtr, 1)
            oTrans.Serial(lnRow - 1, oTrans.Detail(lnRow - 1, "sStockIDx"), pnCtr - 1, 2) = .GridEditor1.TextMatrix(pnCtr, 2)
            oTrans.Serial(lnRow - 1, oTrans.Detail(lnRow - 1, "sStockIDx"), pnCtr - 1, 3) = .GridEditor1.TextMatrix(pnCtr, 3)
            oTrans.Serial(lnRow - 1, oTrans.Detail(lnRow - 1, "sStockIDx"), pnCtr - 1, 4) = .GridEditor1.TextMatrix(pnCtr, 4)
         Next
      End If
      AcceptSerialBackload = .Cancel = 0
   End With
End Function

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String

   lsOldProc = "txtField_KeyDown"
   'On Error GoTo errProc

   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
      With txtField(Index)
         If Index = 10 Then
            If KeyCode = vbKeyF3 Then
               oTrans.SearchMaster 26, .Text
               If .Text <> "" Then SetNextFocus
            Else
               If .Text <> "" Then oTrans.SearchMaster 26, .Text
            End If
         Else
            If KeyCode = vbKeyF3 Then
               oTrans.SearchMaster Index, .Text
               If .Text <> "" Then SetNextFocus
            Else
               If .Text <> "" Then oTrans.SearchMaster Index, .Text
            End If
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
         txtField(pnCtr).Text = Format(oTrans.Master(0), "@@@@-@@@@@@@@")
      Case 1, 8
         txtField(pnCtr).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
      Case 10
         txtField(pnCtr).Text = ""
      Case 11, 12
         txtField(pnCtr).Text = "0.00"
      Case Else
         txtField(pnCtr).Text = oTrans.Master(pnCtr)
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
      .TextMatrix(1, 7) = "0.00"
   End With

   chkField(1).Value = 0

   oTrans.BackLoad = chkField(1).Value
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
      Case 1, 8
         If Not IsDate(.Text) Then .Text = oApp.ServerDate
         .Text = Format(.Text, "MMMM DD, YYYY")
      Case 4
         .Text = Format(.Text, ">")
      Case 11, 12
         If Not IsNumeric(.Text) Then .Text = 0#
         .Text = Format(.Text, "#,##0.00")
      
      End Select
      If Index = 10 Then
         oTrans.Master("sOrderNox") = .Text
      Else
         oTrans.Master(Index) = .Text
      End If
   End With
End Sub

Private Function isEntryOk() As Boolean
   If txtField(9).Text = "" Then
      MsgBox "Company not found!!!" & vbCrLf & _
             "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
      txtField(9).SetFocus
      GoTo EntryNotOK
   End If

   If txtField(2).Text = "" Then
      MsgBox "Supplier not found!!!" & vbCrLf & _
             "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
      txtField(2).SetFocus
      GoTo EntryNotOK
   End If

   If txtField(4).Text = "" Then
      MsgBox "Unknown Reference Number!!!" & vbCrLf & _
             "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
      txtField(4).SetFocus
      GoTo EntryNotOK
   End If
   
   If txtField(6).Text = "" Then
      MsgBox "Invalid Term Detected!!!" & vbCrLf & _
             "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
      txtField(6).SetFocus
      GoTo EntryNotOK
   End If

'   With GridEditor1
'      If Trim(.TextMatrix(1, 1)) = "" Or .TextMatrix(1, 6) = 0 Then
'         MsgBox "Detail is required!!!" & vbCrLf & _
'                "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
'         .SetFocus
'         .Row = 1
'         .Col = 1
'         GoTo EntryNotOK
'      End If
'   End With

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
