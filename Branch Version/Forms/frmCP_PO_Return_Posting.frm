VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0"; "xrGridControl.ocx"
Begin VB.Form frmCP_PO_Return_Posting 
   BorderStyle     =   0  'None
   Caption         =   "PO Return"
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
      Height          =   3255
      Left            =   120
      TabIndex        =   16
      Tag             =   "et0;eb0;et0;bc2"
      Top             =   3510
      Width           =   9990
      _ExtentX        =   17621
      _ExtentY        =   5741
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
      Object.HEIGHT          =   3255
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
      MOUSEICON       =   "frmCP_PO_Return_Posting.frx":0000
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
      Height          =   2385
      Index           =   0
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   1095
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   4207
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   5
         Left            =   6765
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   1095
         Width           =   3090
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   660
         Index           =   4
         Left            =   1095
         MultiLine       =   -1  'True
         TabIndex        =   15
         Text            =   "frmCP_PO_Return_Posting.frx":001C
         Top             =   1575
         Width           =   4545
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   6765
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   765
         Width           =   3090
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   465
         Index           =   3
         Left            =   1095
         MultiLine       =   -1  'True
         TabIndex        =   9
         Text            =   "frmCP_PO_Return_Posting.frx":0024
         Top             =   1095
         Width           =   4545
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   1095
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   120
         Width           =   2265
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   1095
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   765
         Width           =   4545
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Refer #"
         Height          =   195
         Index           =   5
         Left            =   6135
         TabIndex        =   28
         Top             =   1170
         Width           =   540
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
         Left            =   6825
         TabIndex        =   26
         Tag             =   "eb0;et0"
         Top             =   210
         Width           =   2970
      End
      Begin VB.Shape Shape4 
         Height          =   330
         Index           =   0
         Left            =   6795
         Top             =   165
         Width           =   3030
      End
      Begin VB.Shape Shape3 
         Height          =   390
         Index           =   0
         Left            =   6765
         Top             =   135
         Width           =   3090
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
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
         Left            =   6765
         TabIndex        =   13
         Tag             =   "ht0;ft0"
         Top             =   1410
         Width           =   3090
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   14
         Top             =   1605
         Width           =   735
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   195
         Index           =   1
         Left            =   6135
         TabIndex        =   10
         Top             =   810
         Width           =   345
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   8
         Top             =   1125
         Width           =   570
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   315
         Left            =   1185
         Tag             =   "et0;ht2"
         Top             =   225
         Width           =   2265
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trans. No."
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   4
         Top             =   150
         Width           =   750
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   6
         Top             =   810
         Width           =   570
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         Height          =   195
         Index           =   3
         Left            =   6135
         TabIndex        =   12
         Top             =   1440
         Width           =   360
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H8000000F&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   285
         Index           =   0
         Left            =   6825
         Tag             =   "et0;et0"
         Top             =   195
         Width           =   2985
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   10365
      TabIndex        =   24
      Top             =   3075
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
      Picture         =   "frmCP_PO_Return_Posting.frx":0033
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   10365
      TabIndex        =   23
      Top             =   2445
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
      Picture         =   "frmCP_PO_Return_Posting.frx":07AD
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   10365
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
      Picture         =   "frmCP_PO_Return_Posting.frx":0F27
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   10365
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
      Picture         =   "frmCP_PO_Return_Posting.frx":16A1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   10365
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
      Picture         =   "frmCP_PO_Return_Posting.frx":1E1B
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   6
      Left            =   10365
      TabIndex        =   25
      Top             =   3075
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
      Picture         =   "frmCP_PO_Return_Posting.frx":2595
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   10365
      TabIndex        =   20
      Top             =   555
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
      Picture         =   "frmCP_PO_Return_Posting.frx":2D0F
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   7
      Left            =   10365
      TabIndex        =   21
      Top             =   1185
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Update"
      AccessKey       =   "U"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCP_PO_Return_Posting.frx":3489
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   8
      Left            =   10365
      TabIndex        =   22
      Top             =   1815
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
      Picture         =   "frmCP_PO_Return_Posting.frx":3C03
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   525
      Index           =   1
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   10020
      _ExtentX        =   17674
      _ExtentY        =   926
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
         Index           =   6
         Left            =   1095
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   90
         Width           =   2250
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
         Index           =   7
         Left            =   4380
         MaxLength       =   50
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   90
         Width           =   5505
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
         Index           =   10
         Left            =   150
         TabIndex        =   0
         Top             =   135
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
         Left            =   3570
         TabIndex        =   2
         Top             =   135
         Width           =   795
      End
   End
End
Attribute VB_Name = "frmCP_PO_Return_Posting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmCP_PO_Return"

Private WithEvents oTrans As ggcCPPurchasing.clsCPPOReturn
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin

Dim pnIndex As Integer
Dim pbGridFocus As Boolean
Dim pnCtr As Integer
Dim pbSave As Boolean
Dim pbGridValidate As Boolean
Dim pbEditMode As Boolean
Dim pbPosted As Boolean

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   Dim lsRep As String
   Dim lsUserID As String
   Dim lsUserName As String
   Dim lnUserRights As Integer
   Dim lasRights() As String

   lsOldProc = "cmdButton_Click"
   ''On Error GoTo errProc

   If Not pbGridFocus And Index = 0 Then Call txtField_Validate(pnIndex, False)
   With GridEditor1
      Select Case Index
      Case 0 'Save
         If .Rows > 2 Then
            pnCtr = 0
            Do While pnCtr < .Rows
               If Trim(.TextMatrix(pnCtr, 1)) = "" Then
                  .Row = pnCtr
                  If oTrans.deleteDetail(.Row - 1) Then .DeleteRow
               Else
                  pnCtr = pnCtr + 1
               End If
            Loop

            .ColWidth(3) = 2300
            If .Rows > 16 Then .ColWidth(3) = 2100
         End If

         If isEntryOK Then
            If oTrans.SaveTransaction Then
               If oTrans.OpenTransaction(oTrans.Master("sTransNox")) Then
                  MsgBox "Record Updated Successfully!!!", vbInformation, "Confirm"
                  initButton xeModeReady

                  lsRep = MsgBox("Print Transaction!!!", vbYesNo + vbQuestion, "Confirm")
                  If lsRep = vbYes Then
                     If Not PrintTrans Then MsgBox "Unable to Print Transaction!!!", vbCritical, "Warning"
                  End If
                  pbSave = True
               End If
            Else
               MsgBox "Unable to Update Record!!!", vbCritical, "Warning"
            End If
         End If
      Case 1 'Search
         If pbGridFocus Then
            If oTrans.SearchDetail(.Row - 1, 1) Then .Col = 1
            .Refresh
            .SetFocus
         Else
            oTrans.SearchMaster pnIndex
         End If
      Case 2 'Delete Row
         If .Rows > 2 Then
            If oTrans.deleteDetail(.Row - 1) Then .DeleteRow

            For pnCtr = 1 To .Rows - 1
               .TextMatrix(pnCtr, 0) = pnCtr
            Next

            .ColWidth(3) = 2300
            If .Rows > 16 Then .ColWidth(3) = 2100
         End If
      Case 3 'Cancel Update
         lsRep = MsgBox("Transaction is in Update Mode!!!" & vbCrLf & _
                        "Do you want to Cancel Transaction!!!", vbYesNo + vbQuestion, "Confirm")

         If lsRep = vbYes Then
            oTrans.NewTransaction
            Clearfields
            initButton xeModeReady
         Else
            txtField(pnIndex).SetFocus
         End If
         pbSave = False
      Case 4 'Browse
         If oTrans.SearchTransaction() Then
            LoadMaster
            LoadDetail
         End If

         txtField(7).SetFocus
      Case 5 'Print
         If pbSave Then
            lsRep = MsgBox("Print Transaction!!!", vbYesNo + vbQuestion, "Confirm")
            If lsRep = vbYes Then
               If Not PrintTrans Then MsgBox "Unable to Print Transaction!!!", vbCritical, "Warning"
            End If
         Else
            MsgBox "Unable to Print Transaction!!!", vbCritical, "Warning"
         End If
      Case 6 'Close
         Unload Me
      Case 7 'Update
         If txtField(0).Text <> "" Then
            oTrans.UpdateTransaction
            initButton xeModeUpdate
            txtField(2).SetFocus
            pbEditMode = True
         Else
            MsgBox "No Transaction to Update!!!" & vbCrLf & _
                   "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
         End If
      Case 8 'Cancel Transaction
         If txtField(0).Text <> "" Then
            lasRights = Split(oApp.mdiMain.Controls(oApp.MenuName).Tag, "»")
            Debug.Print lasRights(3)
            If GetApproval(oApp, lnUserRights, lsUserID, lsUserName, lasRights(3)) = False Then GoTo endProc

'            If (lnUserRights And (xeSupervisor + xeSysAdmin)) = 0 Then
'               MsgBox "Approving Officer Has no Right to Cancel this transaction!!!" & vbCrLf & _
'                  "Request can not be granted!!!", vbCritical, "Warning"
'               GoTo endProc
'            End If

            If oTrans.CancelTransaction Then
               MsgBox "Transaction Cancelled Successfully!!!", vbInformation, "Notice"
               Label2.Caption = Format(TransStat(oTrans.Master("cTranStat")), ">")
            Else
               MsgBox "Unable to Cancel Transaction!!!", vbCritical, "Warning"
            End If
         End If
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
   ''On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oTrans = New ggcCPPurchasing.clsCPPOReturn
   Set oTrans.AppDriver = oApp
   
   oTrans.TransStatus = 10
   oTrans.InitTransaction
   oTrans.NewTransaction

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualRight

   InitGrid
   Clearfields
   initButton xeModeReady

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
      ElseIf .TextMatrix(.Row, 7) = "0" Then
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
   ''On Error GoTo errProc

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
               If .Row = .Rows - 1 Then
                  .Rows = .Rows + 1
                  oTrans.addDetail
                  .Col = 0
               End If

               .Row = .Rows - 1
            End If
         Case 7
            If CDbl(.TextMatrix(.Row, .Col)) <> 1 Then .TextMatrix(.Row, .Col) = 1
         End Select
      End If

      If .Rows > 16 Then
         .TopRow = .Rows - 1
         .ColWidth(3) = 2100
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
   ''On Error GoTo errProc

   If KeyCode = vbKeyF3 Then
      With GridEditor1
         If oTrans.SearchDetail(.Row - 1, .Col, .TextMatrix(.Row, .Col)) Then
            If oTrans.Detail(.Row - 1, "cHsSerial") = xeYes Then
               .TextMatrix(.Row, 6) = 1
               oTrans.Detail(.Row - 1, "nQuantity") = 1
               If .Row = .Rows - 1 Then
                  .Rows = .Rows + 1
                  oTrans.addDetail
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
      .Cols = 8
      .Font = "MS Sans Serif"

      'column title
      .TextMatrix(0, 1) = "BarrCode"
      .TextMatrix(0, 2) = "Description"
      .TextMatrix(0, 3) = "Model"
      .TextMatrix(0, 4) = "Unit Price"
      .TextMatrix(0, 5) = "QOH"
      .TextMatrix(0, 6) = "ReferNo"
      .TextMatrix(0, 7) = "Qty"

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
      .ColWidth(5) = 500
      .ColWidth(6) = 1290
      .ColWidth(7) = 500

      .ColFormat(4) = "#,##0.00"
      .ColFormat(5) = "#,##0"
      .ColFormat(7) = "#,##0"
      .ColNumberOnly(7) = True
      .ColDefault(4) = 0#
      .ColDefault(5) = 0
      .ColDefault(7) = 0

      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      .ColAlignment(4) = 6
      .ColAlignment(5) = 6
      .ColAlignment(6) = 1
      .ColAlignment(7) = 6

      .ColEnabled(3) = False
      .ColEnabled(4) = False
      .ColEnabled(5) = False
      .ColEnabled(6) = False

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
   ''On Error GoTo errProc

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
         If GetFocus = GridEditor1.hWnd Then Exit Sub
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
   cmdButton(7).Visible = Not lbShow
   cmdButton(8).Visible = Not lbShow

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
   lrs.Fields.Append "nField01", adInteger, 3
   lrs.Fields.Append "nField02", adChar, 1
   lrs.Fields.Append "sField01", adVarChar, 20
   lrs.Fields.Append "sField02", adVarChar, 128
   lrs.Fields.Append "sField03", adVarChar, 20
   lrs.Fields.Append "sField04", adVarChar, 128
   lrs.Fields.Append "sField05", adVarChar, 100
   lrs.Fields.Append "sField06", adVarChar, 25
   lrs.Fields.Append "sField07", adVarChar, 25
   lrs.Fields.Append "sField08", adVarChar, 25
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

   Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\CPPurchaseReturnFormOld.rpt")
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

Private Sub Clearfields()
   For pnCtr = 0 To txtField.Count - 1
      Select Case pnCtr
      Case 1
         txtField(pnCtr).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
      Case Else
         txtField(pnCtr).Text = ""
         txtField(pnCtr).Tag = txtField(pnCtr).Text
      End Select
   Next
      txtField(6).Text = ""
   With GridEditor1
      .Rows = 2
      .ColWidth(3) = 2300

      'empty row
      .TextMatrix(1, 1) = ""
      .TextMatrix(1, 2) = ""
      .TextMatrix(1, 3) = ""
      .TextMatrix(1, 4) = "0.00"
      .TextMatrix(1, 5) = "0"
      .TextMatrix(1, 6) = ""
      .TextMatrix(1, 7) = "0"
   End With

   pbSave = False
   pbPosted = False
   pbEditMode = True
   lblTotal.Caption = "0.00"
   Label2.Caption = "UNKNOWN"
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

      If Index < 5 Then oTrans.Master(Index) = .Text
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

Private Sub LoadMaster()
   For pnCtr = 0 To txtField.Count - 1
      Select Case pnCtr
      Case 0, 5
         txtField(pnCtr).Text = Format(oTrans.Master(0), "@@@@-@@@@@@")
         txtField(pnCtr).Tag = txtField(pnCtr).Text
      Case 1
         txtField(pnCtr).Text = Format(oTrans.Master(pnCtr), "MMMM DD, YYYY")
      Case 2, 7
         txtField(pnCtr).Text = oTrans.Master(2)
         txtField(pnCtr).Tag = txtField(pnCtr).Text
      Case Else
         txtField(pnCtr).Text = IFNull(oTrans.Master(pnCtr), "")
      End Select
   Next

   pbSave = True
   Label2.Caption = Format(TransStat(oTrans.Master("cTranStat")), ">")
End Sub

Private Sub LoadDetail()
   Dim lnCtr As Integer

   With GridEditor1
      .Rows = IIf(oTrans.ItemCount = 0, 2, oTrans.ItemCount + 1)

      .ColWidth(3) = 2300
      If .Rows > 16 Then .ColWidth(3) = 2100

      For pnCtr = 0 To oTrans.ItemCount - 1
         For lnCtr = 1 To .Cols - 1
            .TextMatrix(pnCtr + 1, lnCtr) = oTrans.Detail(pnCtr, lnCtr)
         Next
      Next
   End With
End Sub

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
