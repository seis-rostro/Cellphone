VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0"; "xrGridControl.ocx"
Begin VB.Form frmCP_JO_Branch_Transfer 
   BorderStyle     =   0  'None
   Caption         =   "Job Order Delivery"
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
      Height          =   3555
      Left            =   1575
      TabIndex        =   25
      Top             =   3255
      Width           =   10050
      _ExtentX        =   17727
      _ExtentY        =   6271
      AllowBigSelection=   -1  'True
      AutoAdd         =   0   'False
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
      Object.HEIGHT          =   3555
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
      MOUSEICON       =   "frmCP_JO_Branch_Transfer.frx":0000
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
      Index           =   8
      Left            =   90
      TabIndex        =   18
      Top             =   4050
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Tag All"
      AccessKey       =   "T"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCP_JO_Branch_Transfer.frx":001C
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   90
      TabIndex        =   17
      Top             =   3420
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Retrieve"
      AccessKey       =   "R"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCP_JO_Branch_Transfer.frx":0796
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   465
      Left            =   1575
      Tag             =   "wt0;fb0"
      Top             =   2745
      Width           =   10050
      _ExtentX        =   17727
      _ExtentY        =   820
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   6
         Left            =   5070
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   60
         Width           =   4755
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   5
         Left            =   1350
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   60
         Width           =   2505
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Branch/ASC"
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
         Index           =   2
         Left            =   3960
         TabIndex        =   13
         Top             =   105
         Width           =   1050
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Brand Name"
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
         Left            =   195
         TabIndex        =   11
         Top             =   105
         Width           =   1155
      End
   End
   Begin VB.CheckBox chkField 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Save to &Mobile Disk"
      Height          =   195
      Left            =   9585
      TabIndex        =   24
      Tag             =   "et0;fb0"
      Top             =   735
      Width           =   1725
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   2190
      Index           =   0
      Left            =   1575
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   10050
      _ExtentX        =   17727
      _ExtentY        =   3863
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.CheckBox chkIsBranch 
         BackColor       =   &H00C0C0C0&
         Caption         =   "GUANZON BRANCH"
         Height          =   195
         Left            =   5055
         TabIndex        =   6
         Tag             =   "et0;fb0"
         Top             =   765
         Width           =   1995
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   5055
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   990
         Width           =   4770
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   1365
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   660
         Width           =   2505
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   660
         Index           =   4
         Left            =   1365
         MaxLength       =   1028
         MultiLine       =   -1  'True
         TabIndex        =   10
         Text            =   "frmCP_JO_Branch_Transfer.frx":0F10
         Top             =   1320
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
         Left            =   1365
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   165
         Width           =   1950
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   3
         Left            =   1365
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   990
         Width           =   2505
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
         Index           =   9
         Left            =   165
         TabIndex        =   0
         Top             =   210
         Width           =   1185
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transact. Date"
         Height          =   285
         Index           =   1
         Left            =   165
         TabIndex        =   2
         Top             =   690
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transmittal No"
         Height          =   285
         Index           =   3
         Left            =   195
         TabIndex        =   4
         Top             =   1065
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   285
         Index           =   7
         Left            =   585
         TabIndex        =   9
         Top             =   1410
         Width           =   795
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
         Caption         =   "Destination"
         Height          =   285
         Index           =   4
         Left            =   4185
         TabIndex        =   7
         Top             =   1020
         Width           =   1200
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   90
      TabIndex        =   15
      Top             =   2160
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
      Picture         =   "frmCP_JO_Branch_Transfer.frx":0F26
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   90
      TabIndex        =   16
      Top             =   2790
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
      Picture         =   "frmCP_JO_Branch_Transfer.frx":16A0
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   90
      TabIndex        =   19
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
      Picture         =   "frmCP_JO_Branch_Transfer.frx":1E1A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   90
      TabIndex        =   20
      Top             =   2790
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
      Picture         =   "frmCP_JO_Branch_Transfer.frx":2594
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   6
      Left            =   90
      TabIndex        =   23
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
      Picture         =   "frmCP_JO_Branch_Transfer.frx":2D0E
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   90
      TabIndex        =   22
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
      Picture         =   "frmCP_JO_Branch_Transfer.frx":3488
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   7
      Left            =   90
      TabIndex        =   21
      Top             =   3420
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
      Picture         =   "frmCP_JO_Branch_Transfer.frx":3C02
   End
End
Attribute VB_Name = "frmCP_JO_Branch_Transfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmCP_JO_Branch_Transfer"

Private WithEvents oTrans As clsJobOrderTransfer
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin

Dim pnIndex As Integer
Dim pbGridFocus As Boolean
Dim pnCtr As Integer
Dim pbSave As Boolean
Dim pbGridValidate As Boolean
Dim pbClosedTrans As Boolean

Private Sub chkField_Click()
'   oTrans.DiskTransaction = IIf(chkField.Value = 1, True, False)
End Sub

Private Sub chkIsBranch_Click()
   oTrans.IsBranch = chkIsBranch.Value = Checked
End Sub

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   Dim lnRep As String
   Dim lnCtr As Integer
   
   lsOldProc = "cmdButton_Click"
   ''On Error GoTo errProc
   
   With GridEditor1
      Select Case Index
      Case 0
         If oTrans.ItemCount > 1 Then
           If .Rows > 2 Then
               pnCtr = 0
               Do While pnCtr < .Rows
                  If Trim(.TextMatrix(pnCtr, 5)) = "NO" Then
                     .Row = pnCtr
                     If oTrans.deleteDetail(.Row - 1) Then .deleteRow
                  Else
                     pnCtr = pnCtr + 1
                  End If
               Loop
            End If
         End If
         
         If isEntryOk Then
            If oTrans.SaveTransaction Then
               MsgBox "Record Updated Successfully!!!", vbInformation, "Confirm"
               If Not BranchAutomate(oTrans.Master("sDestinat")) Then
                  If Not oTrans.AcceptDelivery(oTrans.Master("dTransact")) Then
                     MsgBox "Automatic Posting encountered error!!!" & vbCrLf & _
                              "Please contact GMC/GGC SEG for assistance!!!", vbCritical, "Warning"
                  End If
               End If
               
               Call oTrans.OpenTransaction(oTrans.Master("sTransNox"))
               Call LoadDetail("YES")

               initButton xeModeReady
               lnRep = MsgBox("Print Transaction!!!", vbYesNo + vbQuestion, "Confirm")
               If lnRep = vbYes Then
                  If Not PrintTrans Then MsgBox "Unable to Print Transaction!!!", vbCritical, "Warning"
               End If
               pbSave = True
            Else
               MsgBox "Unable to Update Record!!!", vbCritical, "Warning"
            End If
         End If
      Case 1
         If Not pbGridFocus Then
            oTrans.SearchMaster pnIndex
         End If
      Case 2
         If txtField(2).Text = "" Then
            MsgBox "Invalid Destination!!!" & vbCrLf & _
                     "Please Verify your Entry then Try Again", vbCritical, "Warning"
            txtField(2).SetFocus
         Else
            oTrans.Brand = Trim(txtField(5).Text)
            oTrans.BranchName = Trim(txtField(6).Text)
            If oTrans.LoadDetail Then
               Call LoadDetail("NO")
            Else
               Call ClearDetail
            End If
         End If
      Case 3
         lnRep = MsgBox("Transaction is in Update Mode!!!" & vbCrLf & _
                        "Do you want to Cancel Transaction!!!", vbYesNo + vbQuestion, "Confirm")
         
         If lnRep = vbYes Then
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
            lnRep = MsgBox("Print Transaction!!!", vbYesNo + vbQuestion, "Confirm")
            If lnRep = vbYes Then
               If Not PrintTrans Then MsgBox "Unable to Print Transaction!!!", vbCritical, "Warning"
            End If
         Else
            MsgBox "Unable to Print Transaction!!!", vbCritical, "Warning"
         End If
      Case 6
         Unload Me
      Case 7
'         If chkField.Value = 1 Then
'            oTrans.DiskTransaction = True
'            If oTrans.CreateDiskTransfer Then
'               MsgBox "Transaction was Successfully Save to Mobile Disk!!!", vbInformation, "Notice"
'            Else
'               MsgBox "Unable to Save Transaction to Mobile Disk!!!", vbCritical, "Warning"
'            End If
'         Else
'            MsgBox "MC Delivery Capture Was Not Yet Set!!!" & vbCrLf & _
'               "Please Checked 'Save to Mobile Disk' then Try Exporting Delivery Again!!!", _
'               vbCritical, "Warning"
'         End If
      Case 8
         Call TagAll
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

   Set oTrans = New clsJobOrderTransfer
   Set oTrans.AppDriver = oApp
   
'   oTrans.DiskTransaction = False
   oTrans.InitTransaction
   oTrans.NewTransaction
   
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransaction
   
   InitGrid
   ClearFields
   initButton xeModeAddNew
   
   txtField(3).MaxLength = oTrans.MasFldSize(3)
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

Private Sub GridEditor1_DblClick()
   With GridEditor1
      If .TextMatrix(.Row, 5) = "NO" Then
         .TextMatrix(.Row, 5) = "YES"
      Else
         .TextMatrix(.Row, 5) = "NO"
      End If
   End With
End Sub

Private Sub GridEditor1_GotFocus()
   pbGridFocus = True
End Sub

Private Sub GridEditor1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeySpace Then
      With GridEditor1
         If oTrans.Detail(.Row - 1, "sSerialNo") <> "" Then
            If .TextMatrix(.Row, 5) = "NO" Then
               .TextMatrix(.Row, 5) = "YES"
            Else
               .TextMatrix(.Row, 5) = "NO"
            End If
         End If
      End With
      KeyCode = 0
   End If
End Sub

Private Sub GridEditor1_LostFocus()
   pbGridValidate = False
End Sub

Private Sub oTrans_DetailRetrieved(ByVal Index As Integer)
   With GridEditor1
      .TextMatrix(.Row, Index) = oTrans.Detail(.Row - 1, Index)
   End With
End Sub

Private Sub oTrans_MasterRetrieved(ByVal Index As Integer)
   Select Case Index
   Case 2, 5
      If chkIsBranch.Value = 0 Then
         txtField(2).Text = oTrans.Master(5)
      Else
         txtField(Index).Text = oTrans.Master(Index)
      End If
   Case Else
      txtField(Index).Text = oTrans.Master(Index)
   End Select
End Sub

Private Sub InitGrid()
   With GridEditor1
      .Rows = 2
      .Cols = 6
      .Font = "MS Sans Serif"

      'column title
      .TextMatrix(0, 1) = "Branch"
      .TextMatrix(0, 2) = "JO No"
      .TextMatrix(0, 3) = "IMEI"
      .TextMatrix(0, 4) = "Brand"
      .TextMatrix(0, 5) = "FWRD"
      
      .Row = 0

      'column alignment
      For pnCtr = 0 To .Cols - 1
         .Col = pnCtr
         .CellFontBold = True
         .CellAlignment = 3
         .ColEnabled(pnCtr) = False
      Next

      'column width
      .ColWidth(0) = 330
      .ColWidth(2) = 1000
      .ColWidth(3) = 2200
      .ColWidth(5) = 700
      
      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      .ColAlignment(4) = 1
      .ColAlignment(5) = 3
      
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

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String
   Dim lsSearchOn As String
   
   lsOldProc = "txtField_KeyDown"
   ''On Error GoTo errProc
   
   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
      With txtField(Index)
         Select Case Index
         Case 2
            If KeyCode = vbKeyF3 Then
               oTrans.SearchMaster Index, .Text
               If .Text <> "" Then SetNextFocus
            Else
               If .Text <> "" Then oTrans.SearchMaster Index, .Text
            End If
         End Select
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
   cmdButton(7).Visible = Not lbShow
   
   cmdButton(0).Visible = lbShow
   cmdButton(1).Visible = lbShow
   cmdButton(2).Visible = lbShow
   cmdButton(3).Visible = lbShow
   cmdButton(8).Visible = lbShow
   
   For pnCtr = 1 To txtField.Count - 1
      txtField(pnCtr).Enabled = lbShow
   Next
   
   If Not lbShow Then cmdButton(4).SetFocus
End Sub

Private Function PrintTrans() As Boolean
   Dim loreport As frmRepViewer
   Dim lrs As Recordset
   Dim lors As Recordset
   Dim lnCtr As Integer
   Dim lsOldProc As String
   Dim lnTotlWSerial As Double
   Dim lnTotlWOSerial As Double
   Dim lsSQL As String

   lsOldProc = "PrinTrans"
   ''On Error GoTo errProc

   PrintTrans = True

   Set lrs = New ADODB.Recordset

   lrs.Fields.Append "nField01", adInteger, 3
   lrs.Fields.Append "nField02", adChar, 1
   lrs.Fields.Append "sField01", adVarChar, 20
   lrs.Fields.Append "sField02", adVarChar, 128
   lrs.Fields.Append "sField03", adVarChar, 20
   lrs.Fields.Append "sField04", adVarChar, 12
   lrs.Fields.Append "sField05", adVarChar, 100
   lrs.Open

   With oTrans
      lnTotlWOSerial = 0
      lnTotlWSerial = 0

      For lnCtr = 0 To .ItemCount - 1
         lrs.AddNew
         lrs.Fields("nField01") = 1
         lrs.Fields("nField02") = oTrans.Detail(lnCtr, "cHsSerial")
         lrs.Fields("sField01") = oTrans.Detail(lnCtr, "sBrandNme")
         lrs.Fields("sField02") = oTrans.Detail(lnCtr, "sModelNme")
         lrs.Fields("sField03") = IFNull(oTrans.Detail(lnCtr, "sSerialNo"), "")
         lrs.Fields("sField04") = oTrans.Detail(lnCtr, "sReferNox")
         lnTotlWSerial = lnTotlWSerial + 1
      Next
      lrs.Sort = "nField02 DESC,sField05,sField03"
   End With

   'assign important info to the report
   Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\CP_JO_Transfer.rpt")
   oReport.DiscardSavedData
   oReport.FieldMappingType = crAutoFieldMapping
   oReport.Database.SetDataSource lrs

   Set lors = New ADODB.Recordset
   If lors.State = adStateOpen Then lors.Close

   If chkIsBranch.Value = Checked Then
      lsSQL = "SELECT" _
                  & "  a.sAddressx" _
                  & ", CONCAT(b.sTownName, ', ' , c.sProvName, ' ' , b.sZippCode) xTownName" _
                  & ", a.sBranchNm xSourcexx" _
               & " FROM Branch a" _
                  & " LEFT JOIN TownCity b" _
                     & " LEFT JOIN Province c" _
                        & " ON b.sProvIDxx = c.sProvIDxx" _
                     & " ON a.sTownIDxx = b.sTownIDxx" _
               & " WHERE a.sBranchCd = " & strParm(oTrans.Master("sDestinat"))
   Else
      lsSQL = "SELECT" _
                  & "  a.sAddressx" _
                  & ", CONCAT(b.sTownName, ', ' , c.sProvName, ' ' , b.sZippCode) xTownName" _
                  & ", a.sCompnyNm xSourcexx" _
               & " FROM Client_Master a" _
                  & " LEFT JOIN TownCity b" _
                     & " LEFT JOIN Province c" _
                        & " ON b.sProvIDxx = c.sProvIDxx" _
                     & " ON a.sTownIDxx = b.sTownIDxx" _
               & " WHERE a.sClientID = " & strParm(oTrans.Master("sLocation"))
   End If
   lors.Open lsSQL, oApp.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText

   oReport.Sections("PHa").ReportObjects("txtRefNo").SetText "CP" & "-" & Right(oTrans.Master("sTransNox"), 10)
   oReport.Sections("PHa").ReportObjects("txtDate").SetText txtField(1).Text
   oReport.Sections("PHb").ReportObjects("txtTo").SetText lors("xSourcexx")
   oReport.Sections("PHb").ReportObjects("txtToAddress").SetText lors("sAddressx") & IFNull(lors("xTownName"), "")
   oReport.Sections("PHb").ReportObjects("txtFrom").SetText oApp.ClientName
   oReport.Sections("PHb").ReportObjects("txtFromAddress").SetText oApp.Address & ", " & oApp.TownCity & ", " & oApp.Province & " " & oApp.ZippCode
   oReport.Sections("RFb").ReportObjects("txtRemarks").SetText txtField(4).Text
   oReport.Sections("RFb").ReportObjects("txtWithSerial").SetText IIf(lnTotlWSerial = 0, "", Format(lnTotlWSerial, "#,##0"))
   oReport.Sections("RFb").ReportObjects("txtWOutSerial").SetText IIf(lnTotlWOSerial = 0, "", Format(lnTotlWOSerial, "#,##0"))
   oReport.Sections("PF").ReportObjects("txtRptUser").SetText oApp.UserName

   Set loreport = New frmRepViewer
   Set loreport.ReportSource = oReport
   loreport.Show
   lors.Close

   PrintTrans = True

endPoc:
'   If Not pbClosedTrans Then
'      If Not BranchAutomate(oTrans.Master("sDestinat")) Then
'         If oTrans.CloseTransaction(oTrans.Master(0)) Then pbClosedTrans = True
'      End If
'   End If
   Set loreport = Nothing
   Set oReport = Nothing
   Set lrs = Nothing
   Set lors = Nothing
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
      Case 5, 6
         txtField(pnCtr).Text = ""
      Case Else
         txtField(pnCtr).Text = oTrans.Master(pnCtr)
      End Select
   Next

   Call ClearDetail
   chkField.Value = 0
   chkIsBranch.Value = 1
   pbSave = False
   pbClosedTrans = False
End Sub

Private Sub ClearDetail()
   With GridEditor1
      .Rows = 2
      .ColWidth(1) = 3000
      .ColWidth(4) = 2700
      
      'empty row
      .TextMatrix(1, 0) = "1"
      .TextMatrix(1, 1) = ""
      .TextMatrix(1, 2) = ""
      .TextMatrix(1, 3) = ""
      .TextMatrix(1, 4) = ""
      .TextMatrix(1, 5) = "NO"
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
      Case 3
         .Text = Format(.Text, ">")
      End Select
      
      If Index < 5 Then oTrans.Master(Index) = txtField(Index).Text
   End With
End Sub

Private Function isEntryOk() As Boolean
   Dim lbHsDetail As Boolean
   
   If txtField(2).Text = "" Then
      MsgBox "Destination not found!!!" & vbCrLf & _
             "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
      txtField(2).SetFocus
      GoTo EntryNotOK
   End If
   
   With GridEditor1
      For pnCtr = 1 To .Rows - 1
         If Not lbHsDetail Then
            If UCase(.TextMatrix(pnCtr, 5)) = "YES" Then
               lbHsDetail = True
               Exit For
            End If
         End If
      Next
   End With
   If Not lbHsDetail Then GoTo EntryNotOK
   
EntryOK:
   isEntryOk = True
   Exit Function
EntryNotOK:
   isEntryOk = False
End Function

Private Sub LoadDetail(ByVal sForwarded As String)
   Dim lnCtr As Integer

   With GridEditor1
      .Rows = IIf(oTrans.ItemCount = 0, 2, oTrans.ItemCount + 1)
      .ColWidth(1) = 3000
      .ColWidth(4) = 2700
      If .Rows > 14 Then
         .ColWidth(1) = 2900
         .ColWidth(4) = 2600
      End If
            
      For pnCtr = 0 To oTrans.ItemCount - 1
         .TextMatrix(pnCtr + 1, 1) = oTrans.Detail(pnCtr, "sBranchNm")
         .TextMatrix(pnCtr + 1, 2) = oTrans.Detail(pnCtr, "sReferNox")
         .TextMatrix(pnCtr + 1, 3) = oTrans.Detail(pnCtr, "sSerialNo")
         .TextMatrix(pnCtr + 1, 4) = oTrans.Detail(pnCtr, "sBrandNme")
         .TextMatrix(pnCtr + 1, 5) = sForwarded
      Next
   End With
End Sub

Private Function BranchAutomate(ByVal sBranchCd As String) As Boolean
   Dim lrs As Recordset
   
   If chkIsBranch.Value = Checked Then BranchAutomate = True
   
   Set lrs = New Recordset
   lrs.Open "SELECT * FROM Branch" & _
               " WHERE sBranchCd = " & strParm(sBranchCd) & _
                  " AND cAutomate = " & strParm(xeYes) _
   , oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   
   If Not lrs.EOF Then BranchAutomate = True
   Set lrs = Nothing
End Function

Private Sub TagAll()
   With GridEditor1
      If oTrans.ItemCount = 1 And oTrans.Detail(0, "sSerialNo") = "" Then Exit Sub
      
      For pnCtr = 0 To oTrans.ItemCount - 1
         If .TextMatrix(pnCtr + 1, 5) = "NO" Then
            .TextMatrix(pnCtr + 1, 5) = "YES"
         Else
            .TextMatrix(pnCtr + 1, 5) = "NO"
         End If
      Next
   End With
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
