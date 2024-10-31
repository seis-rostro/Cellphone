VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmCPInvTypeTrans 
   BorderStyle     =   0  'None
   Caption         =   "CP Inventory Type Transfer"
   ClientHeight    =   6180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10080
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6180
   ScaleWidth      =   10080
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame2 
      Height          =   2820
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   2280
      Width           =   4350
      _ExtentX        =   7673
      _ExtentY        =   4974
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   4
         Left            =   1305
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   2145
         Width           =   2760
      End
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   1305
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   1785
         Width           =   2760
      End
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   1305
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   1440
         Width           =   2760
      End
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   1305
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   315
         Width           =   2760
      End
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   1305
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   660
         Width           =   2760
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Color"
         Height          =   195
         Index           =   7
         Left            =   375
         TabIndex        =   14
         Top             =   2160
         Width           =   810
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
         Height          =   195
         Index           =   6
         Left            =   375
         TabIndex        =   12
         Top             =   1830
         Width           =   810
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Brand"
         Height          =   195
         Index           =   3
         Left            =   375
         TabIndex        =   10
         Top             =   1500
         Width           =   810
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "BARCODE"
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
         Index           =   5
         Left            =   180
         TabIndex        =   6
         Top             =   360
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
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
         Index           =   4
         Left            =   180
         TabIndex        =   8
         Top             =   720
         Width           =   975
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1710
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   4350
      _ExtentX        =   7673
      _ExtentY        =   3016
      BorderStyle     =   1
      Begin VB.ComboBox cmbField 
         Height          =   315
         ItemData        =   "frmCPInvTypeTrans.frx":0000
         Left            =   1575
         List            =   "frmCPInvTypeTrans.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1095
         Width           =   2640
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   1575
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   750
         Width           =   2625
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
         Left            =   1575
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   315
         Width           =   1530
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   315
         Left            =   1650
         Tag             =   "et0;ht2"
         Top             =   405
         Width           =   1530
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Inventory Type"
         Height          =   195
         Index           =   0
         Left            =   105
         TabIndex        =   4
         Top             =   1155
         Width           =   1065
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transact Date"
         Height          =   195
         Index           =   1
         Left            =   105
         TabIndex        =   2
         Top             =   810
         Width           =   1020
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction No."
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
         Index           =   2
         Left            =   105
         TabIndex        =   0
         Top             =   360
         Width           =   1455
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   0
      Left            =   9210
      TabIndex        =   22
      Top             =   5400
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmCPInvTypeTrans.frx":0033
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   2
      Left            =   6915
      TabIndex        =   18
      Top             =   5400
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmCPInvTypeTrans.frx":07AD
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   1
      Left            =   6915
      TabIndex        =   17
      Top             =   5400
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmCPInvTypeTrans.frx":0F27
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   5
      Left            =   9210
      TabIndex        =   21
      Top             =   5400
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmCPInvTypeTrans.frx":16A1
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   4
      Left            =   7680
      TabIndex        =   20
      Top             =   5400
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmCPInvTypeTrans.frx":1E1B
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   3
      Left            =   8445
      TabIndex        =   19
      Top             =   5400
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmCPInvTypeTrans.frx":2595
      PicturePos      =   1
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4545
      Left            =   4530
      TabIndex        =   16
      Top             =   555
      Width           =   5430
      _ExtentX        =   9578
      _ExtentY        =   8017
      _Version        =   393216
      FocusRect       =   0
      SelectionMode   =   1
   End
End
Attribute VB_Name = "frmCPInvTypeTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmCPInvTypeTrans"

Private WithEvents oTrans As clsCPInvTypTrans
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin
Private bLoaded As Boolean

Dim poRSDet As Recordset

Dim pbMoveCombo As Boolean
Dim pnIndex As Integer
Dim pnRow As Integer
Dim pnEditMode As Integer

Private Sub cmbField_Click()
   With cmbField
      If .ListIndex = -1 Then
         oTrans.Master("cInvTypex") = Null
      Else
         oTrans.Master("cInvTypex") = .ListIndex
      End If
   End With
End Sub

Private Sub cmbField_GotFocus()
   pbMoveCombo = True
End Sub

Private Sub cmbField_LostFocus()
   If cmbField.ListIndex < 0 Then cmbField.ListIndex = -1
   pbMoveCombo = False
End Sub

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   
   lsOldProc = "cmdButton_Click"
   ''On Error GoTo errProc
   
   txtField_LostFocus pnIndex
   
   Select Case Index
   Case 0 'cancel
      If oTrans.InitTransaction Then
         ClearFields
         InitGrid
         initButton oTrans.EditMode
         cmdButton(1).SetFocus
      End If
   Case 1 'new
      If oTrans.NewTransaction Then
         ClearFields
         InitGrid
         initButton oTrans.EditMode
         txtField(1).SetFocus
      End If
   Case 2 'save
      If cmbField.ListIndex = -1 Then
         MsgBox "Invalid inventory type detected.", vbCritical, "Warning"
         Exit Sub
      End If
      'delete last row of detail if barcode is empty
      If oTrans.Detail(MSFlexGrid1.Row - 1, "sBarrCode") = "" Then oTrans.deleteDetail (MSFlexGrid1.Row - 1)
      
      If oTrans.SaveTransaction Then
         MsgBox "Transaction was saved successfuly.", vbInformation, "Notice"
      Else
         MsgBox "Unable to save transaction.", vbCritical, "Warning"
      End If
   Case 3 'search
      Call txtOthers_KeyDown(0, vbKeyF3, 0)
      txtOthers(0).SetFocus
   Case 4 'delete
      Call deleteDetail
   Case 5
      Unload Me
   End Select

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Sub initButton(ByVal fnEdit As Integer)
   cmdButton(0).Visible = fnEdit = 1
   cmdButton(2).Visible = fnEdit = 1
   cmdButton(1).Visible = Not fnEdit = 1
   cmdButton(5).Visible = Not fnEdit = 1
   
   xrFrame1.Enabled = fnEdit = 1
   xrFrame2.Enabled = fnEdit = 1
End Sub

Private Sub Form_Activate()
   Dim lsOldProc As String
   
   lsOldProc = "Form_Activate"
   ''On Error GoTo errProc
   
   MSFlexGrid1.Refresh

   oApp.MenuName = Me.Tag
   Me.ZOrder 0
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String
   
   lsOldProc = "Form_Load"
'   ''On Error GoTo errProc
   
   CenterChildForm mdiMain, Me
   
   bLoaded = False
   
   Set oTrans = New clsCPInvTypTrans
   Set oTrans.AppDriver = oApp

   oTrans.InitTransaction
   oTrans.NewTransaction
   
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormMaintenance
   
   InitGrid
   LoadDetailDes
   ClearFields

   initButton oTrans.EditMode
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
End Sub

Private Sub MSFlexGrid1_Click()
   With MSFlexGrid1
      If oTrans.ItemCount = 1 And oTrans.Detail(.Row - 1, "sBarrCode") = "" Then Exit Sub
   
      txtOthers(0) = .TextMatrix(.Row, 1)
      txtOthers(1) = .TextMatrix(.Row, 2)
      txtOthers(2) = oTrans.Detail(.Row - 1, "sBrandNme")
      txtOthers(3) = oTrans.Detail(.Row - 1, "sModelNme")
      txtOthers(4) = oTrans.Detail(.Row - 1, "sColorNme")
      
      txtOthers(0).SetFocus
      
      pnRow = .Row
   End With
End Sub

Private Sub oTrans_DetailRetrieved(ByVal Index As Integer)
   Select Case Index
   Case 0
      txtOthers(Index) = oTrans.Detail(pnRow - 1, "sBarrCode")
   Case 1
      txtOthers(Index) = oTrans.Detail(pnRow - 1, "sDescript")
   Case 2
      txtOthers(Index) = oTrans.Detail(pnRow - 1, "sBrandNme")
   Case 3
      txtOthers(Index) = oTrans.Detail(pnRow - 1, "sModelNme")
   Case 4
      txtOthers(Index) = oTrans.Detail(pnRow - 1, "sColorNme")
   End Select
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      If Index = 1 Then .Text = Format(.Text, "MM/DD/YY")
      
      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = oApp.getColor("HT1")
   End With
   pnIndex = Index
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   With txtField(Index)
      If Not IsDate(.Text) Then .Text = oApp.ServerDate
      .Text = Format(.Text, "MMMM DD, YYYY")
      
      oTrans.Master(Index) = .Text
   End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyUp, vbKeyDown, vbKeyReturn
      If Not txtField(1).hwnd And KeyCode = vbKeyReturn Then
         Call txtField_Validate(1, False)
         cmbField.SetFocus
      End If
      
      Select Case KeyCode
      Case vbKeyDown
         SetNextFocus
      Case vbKeyUp
         SetPreviousFocus
      End Select
   End Select
End Sub

Private Sub InitGrid()
   With MSFlexGrid1
      .Rows = 1
      .AddItem ""
      
      .Cols = 3
      .TextMatrix(0, 0) = "No"
      .TextMatrix(0, 1) = "Barcode"
      .TextMatrix(0, 2) = "Description"
         
      .ColWidth(0) = 600
      .ColWidth(1) = 2150
      .ColWidth(2) = 2570
      
      .ColAlignment(0) = 1
      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      
      .Row = 1
      .Col = 1
      .ColSel = .Cols - 1
   End With
End Sub

Private Sub ClearFields()
   With txtField
      .Item(0).Text = Format(oTrans.Master("sTransNox"), "@@@@-@@@@@@@@")
      .Item(1).Text = Format(oTrans.Master("dTransact"), "MMMM DD, YYYY")
   End With
   
   With MSFlexGrid1
      .Rows = 2
      
      'empty row
      .TextMatrix(1, 1) = ""
      .TextMatrix(1, 2) = ""
   End With
   
   If oTrans.EditMode = xeModeReady Then cmbField.ListIndex = -1
   
   With txtOthers
      .Item(0).Text = ""
      .Item(1).Text = ""
      .Item(2).Text = ""
      .Item(3).Text = ""
      .Item(4).Text = ""
   End With
   
   pnRow = 1
End Sub

Private Sub txtOthers_GotFocus(Index As Integer)
   With txtOthers(Index)
      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = oApp.getColor("HT1")
   End With
End Sub

   Private Sub txtOthers_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String
   
   lsOldProc = "txtOthers_KeyDown"
   ''On Error GoTo errProc
   
   Select Case KeyCode
      Case vbKeyF3
         With txtOthers(Index)
            If KeyCode = vbKeyF3 Then
               oTrans.searchDetail pnRow - 1, Index + 1, .Text
               If .Text <> "" Then SetNextFocus
            End If
         End With
         KeyCode = 0
      Case vbKeyReturn
         If Index = 0 Or Index = 1 Then
            Call addDetail
         
            KeyCode = 0
         End If
   End Select

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " _
                       & "  " & Index _
                       & ", " & KeyCode _
                       & ", " & Shift _
                       & " )", True
End Sub

Private Sub txtOthers_LostFocus(Index As Integer)
   With txtOthers(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub addDetail()
   Dim lsOldProc As String
   Dim lnCtr As Integer
   
   lsOldProc = "addDetail"
   
   With oTrans
      If .EditMode <> xeModeAddNew Then Exit Sub
      If .Detail(pnRow - 1, "sBarrCode") = "" Then Exit Sub
         
      'Search for existing detail with the same barcode
      If .ItemCount > 1 Then
         For lnCtr = 0 To .ItemCount - 2 'dont search for the last row, since its empty
            If txtOthers(0) = .Detail(lnCtr, "sBarrCode") Then GoTo movetolast
         Next
      End If
      
      Call oTrans.addDetail
   End With
   
movetolast:
   Call ClearFields
   Call LoadDetailDes

   With MSFlexGrid1
      pnRow = .Rows - 1
      .Row = .Rows - 1
      .Col = 1
      .ColSel = .Cols - 1
   End With
   
   txtOthers(0).SetFocus
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub deleteDetail()
   With oTrans
      If .EditMode <> xeModeAddNew Then Exit Sub
      If .Detail(pnRow - 1, "sBarrCode") <> "" Then Call .deleteDetail(pnRow - 1)
   End With
   
   Call ClearFields
   Call LoadDetailDes

   With MSFlexGrid1
      pnRow = .Rows - 1
      .Row = .Rows - 1
      .Col = 1
      .ColSel = .Cols - 1
   End With
   
   txtOthers(0).SetFocus
End Sub
Private Sub LoadDetailDes()
   Dim lnCtr As Integer
   Dim lnRow As Integer
   
   lnRow = oTrans.ItemCount
   
   With MSFlexGrid1
      .Rows = lnRow + 1
      
      For lnCtr = 0 To lnRow - 1
         .TextMatrix(lnCtr + 1, 0) = lnCtr + 1
         .TextMatrix(lnCtr + 1, 1) = oTrans.Detail(lnCtr, "sBarrCode")
         .TextMatrix(lnCtr + 1, 2) = oTrans.Detail(lnCtr, "sDescript")
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
