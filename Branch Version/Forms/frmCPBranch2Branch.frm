VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0#1.0#0"; "xrGridControl.ocx"
Begin VB.Form frmCPBranch2Branch 
   BorderStyle     =   0  'None
   Caption         =   "Motorcycle Branch to Branch Transfer"
   ClientHeight    =   7800
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
   ScaleHeight     =   7800
   ScaleWidth      =   11940
   ShowInTaskbar   =   0   'False
   Begin xrGridEditor.GridEditor GridEditor1 
      Height          =   5535
      Left            =   1575
      TabIndex        =   12
      Tag             =   "et0;eb0;et0;bc2"
      Top             =   2130
      Width           =   10245
      _ExtentX        =   18071
      _ExtentY        =   9763
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
      Object.HEIGHT          =   5535
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
      MOUSEICON       =   "frmCPBranch2Branch.frx":0000
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
      Picture         =   "frmCPBranch2Branch.frx":001C
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   990
      Index           =   0
      Left            =   1575
      Tag             =   "wt0;fb0"
      Top             =   1125
      Width           =   10245
      _ExtentX        =   18071
      _ExtentY        =   1746
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   5235
         MaxLength       =   50
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   150
         Width           =   4860
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   1365
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   150
         Width           =   2505
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   4
         Left            =   5235
         MaxLength       =   128
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   480
         Width           =   4860
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   3
         Left            =   1365
         MaxLength       =   8
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   480
         Width           =   2505
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
         Left            =   180
         TabIndex        =   8
         Top             =   510
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   285
         Index           =   7
         Left            =   4140
         TabIndex        =   10
         Top             =   510
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Destination"
         Height          =   285
         Index           =   4
         Left            =   4140
         TabIndex        =   6
         Top             =   180
         Width           =   1200
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   90
      TabIndex        =   13
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
      Picture         =   "frmCPBranch2Branch.frx":0796
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   90
      TabIndex        =   14
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
      Picture         =   "frmCPBranch2Branch.frx":0F10
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
      Picture         =   "frmCPBranch2Branch.frx":168A
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
      Picture         =   "frmCPBranch2Branch.frx":1E04
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   7
      Left            =   90
      TabIndex        =   20
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
      Picture         =   "frmCPBranch2Branch.frx":257E
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   6
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
      Picture         =   "frmCPBranch2Branch.frx":2CF8
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   90
      TabIndex        =   18
      Top             =   4320
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Access."
      AccessKey       =   "A"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCPBranch2Branch.frx":33DA
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   555
      Index           =   1
      Left            =   1575
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   10245
      _ExtentX        =   18071
      _ExtentY        =   979
      BackColor       =   12632256
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
         Index           =   0
         Left            =   8100
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
         Width           =   5220
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
         Left            =   6705
         TabIndex        =   2
         Top             =   150
         Width           =   1440
      End
   End
End
Attribute VB_Name = "frmCPBranch2Branch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pxeMODULENAME As String = "frmMCBranch2Branch"

Private WithEvents oTrans As clsMCDelivery
Attribute oTrans.VB_VarHelpID = -1
Private oForm As frmMCDeliveryAccess
Private oSkin As clsFormSkin

Dim pnIndex As Integer
Dim pbGridFocus As Boolean
Dim pnCtr As Integer
Dim pbSave As Boolean
Dim psBranch As String
Dim pnCancelFocus As Long
Dim pnSearchFocus As Long

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   Dim lsRep As String
   
   lsOldProc = "cmdButton_Click"
   On Error GoTo errProc
   
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
            
            .ColWidth(3) = 3840
            If .Rows > 22 Then .ColWidth(3) = 3640
         End If
         
         If isEntryOK Then
            If oTrans.AccCount > 0 Then
'               If Not AcceptAccess Then Exit Sub
            End If
            If oTrans.SaveTransaction = True Then
               MsgBox "Record Updated Successfully!!!", vbInformation, "Confirm"
               lsRep = MsgBox("Post Transaction!!!", vbYesNo + vbQuestion, "Confirm")
               
               If lsRep = vbYes Then
                  If Not oTrans.AcceptDelivery(oTrans.Master("dTransact")) Then MsgBox "Unable to Post Transaction!!!", vbCritical, "Warning"
               End If
               InitButton xeModeReady
               pbSave = True
            Else
               MsgBox "Unable to Update Record!!!", vbCritical, "Warning"
            End If
         End If
      Case 1
         If pbGridFocus = True Then
            If oTrans.SearchDetail(.Row - 1, 1) Then .Col = 1
            .Refresh
            .SetFocus
         Else
            oTrans.SearchMaster pnIndex
         End If
      Case 2
         If .Rows > 2 Then
            If oTrans.DeleteDetail(.Row - 1) Then .DeleteRow
            
            'column width
            .ColWidth(3) = 3840
            If .Rows > 22 Then .ColWidth(3) = 3640
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
         txtField(5).SetFocus
      Case 5
'         If pbSave Then AcceptAccess
      Case 6
         If pbSave Then
            lsRep = MsgBox("Print Transaction!!!", vbYesNo + vbQuestion, "Confirm")
            If lsRep = vbYes Then
               If Not PrintTrans Then MsgBox "Unable to Print Transaction!!!", vbCritical, "Warning"
            End If
         Else
            MsgBox "Unable to Print Transaction!!!", vbCritical, "Warning"
         End If
      Case 7
         Unload Me
      End Select
   End With
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   GridEditor1.Refresh
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String
   
   lsOldProc = "Form_Load"
   On Error GoTo errProc
   CenterChildForm mdiMain, Me

   Set oTrans = New clsMCDelivery
   Set oTrans.AppDriver = oApp
   
   Set oForm = New frmMCDeliveryAccess
   
   oTrans.InitTransaction
   oTrans.NewTransaction
   
   InitGrid
   ClearFields
   InitButton xeModeAddNew
   txtField(0).Enabled = False
   
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransaction
   pnCancelFocus = cmdButton(3).hWnd
   pnSearchFocus = cmdButton(1).hWnd
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oTrans = Nothing
   Set oSkin = Nothing
   Set oForm = Nothing
End Sub

Private Sub GridEditor1_AddingRow(Cancel As Boolean)
   With GridEditor1
      If .TextMatrix(.Row, 1) = "" Then Cancel = True
      If Not Cancel Then oTrans.AddDetail
      
      If .Rows > 22 Then .ColWidth(3) = 3640
   End With
End Sub

Private Sub GridEditor1_EditorValidate(Cancel As Boolean)
   If GetFocus = pnCancelFocus Or GetFocus = pnSearchFocus Then Exit Sub
   With GridEditor1
      oTrans.Detail(.Row - 1, .Col) = .TextMatrix(.Row, .Col)
   End With
End Sub

Private Sub GridEditor1_GotFocus()
   pbGridFocus = True
End Sub

Private Sub GridEditor1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Then
      With GridEditor1
         If oTrans.SearchDetail(.Row - 1, 1, .TextMatrix(.Row, 1)) Then .Col = 1
         .Refresh
         .SetFocus
         KeyCode = 0
      End With
   End If
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
      .Cols = 4
      .Font = "MS Sans Serif"
      
      'column title
      .TextMatrix(0, 1) = "Engine No"
      .TextMatrix(0, 2) = "Frame No"
      .TextMatrix(0, 3) = "Model"
 
      .Row = 0
      
      'column alignment
      For pnCtr = 0 To .Cols - 1
         .Col = pnCtr
         .CellFontBold = True
         .CellAlignment = 1
      Next

      'column width
      .ColWidth(0) = 330
      .ColWidth(1) = 3000
      .ColWidth(2) = 3000

      .ColEnabled(2) = False
      .ColEnabled(3) = False

      .ColFormat(1) = ">"
      .ColFormat(2) = ">"
      
      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      
      .Row = 1
      .Col = 1
   End With
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   If Index = 1 Then txtField(Index).Text = Format(txtField(Index).Text, "MM/DD/YY")
   If txtField(Index) <> Empty Then
      txtField(Index).SelStart = 0
      txtField(Index).SelLength = Len(txtField(Index).Text)
   End If
   pbGridFocus = False
   pnIndex = Index
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Then
      Select Case Index
      Case 2
         oTrans.SearchMaster Index, txtField(Index).Text
      Case 5
         SearchBranch False, txtField(Index).Text
      End Select
      If txtField(Index).Text <> "" Then SetNextFocus
      KeyCode = 0
   End If
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

Private Sub InitButton(lnStat As Integer)
   Dim lbShow As Boolean

   lbShow = IIf(lnStat = 0, False, True)
   cmdButton(0).Visible = lbShow
   cmdButton(1).Visible = lbShow
   cmdButton(2).Visible = lbShow
   cmdButton(3).Visible = lbShow
   
   For pnCtr = 1 To txtField.Count - 1
      txtField(pnCtr).Enabled = lbShow
   Next
   
   With GridEditor1
      .ColEnabled(1) = lbShow
   End With
   
   cmdButton(4).Visible = Not lbShow
   cmdButton(5).Visible = Not lbShow
   cmdButton(6).Visible = Not lbShow
   cmdButton(7).Visible = Not lbShow
   
   xrFrame1(0).Enabled = Not lbShow
   xrFrame1(1).Enabled = lbShow
   
   If Not lbShow Then cmdButton(4).SetFocus
End Sub

Private Function PrintTrans() As Boolean
   Dim lrs As ADODB.Recordset
   Dim lors As ADODB.Recordset
   Dim lnCtr As Integer
   Dim lsMCInvID As String
   Dim lasMCInv() As String
   Dim lanMCInv() As Integer
   Dim lsIncluded As String
   Dim lsExcluded As String
   Dim lnQuantity As Integer
   Dim lnGivenxxx  As Integer
   Dim lsAcsModID As String
   Dim lasAcsMod() As String
   Dim lbFirst As Boolean

   PrintTrans = True
   On Error GoTo errProc
   
   Set lrs = New ADODB.Recordset

   lrs.Fields.Append "nField01", adInteger, 3
   lrs.Fields.Append "sField01", adVarChar, 50
   lrs.Fields.Append "sField02", adVarChar, 28
   lrs.Fields.Append "sField03", adVarChar, 20
   lrs.Fields.Append "sField04", adVarChar, 4
   lrs.Open

   With oTrans
      lsMCInvID = ""
      For pnCtr = 0 To .ItemCount - 1
         If .Detail(pnCtr, "sEngineNo") = "" Then Exit For

         If InStr(1, lsMCInvID, .Detail(pnCtr, "sMCInvIDx"), vbTextCompare) = 0 Then
            lsMCInvID = lsMCInvID & "�" & .Detail(pnCtr, "sMCInvIDx")
         End If
      Next
      
      lsMCInvID = Mid(lsMCInvID, 2)
      lasMCInv = Split(lsMCInvID, "�")
      ReDim lanMCInv(UBound(lasMCInv)) As Integer
      
      For pnCtr = 0 To .ItemCount - 1
         If .Detail(pnCtr, "sEngineNo") = "" Then Exit For
         
         lnCtr = InStr(1, lsMCInvID, .Detail(pnCtr, "sMCInvIDx"))
         lnCtr = lnCtr \ 7
         If lnCtr > 6 Then lnCtr = lnCtr - 1
         lanMCInv(lnCtr) = lanMCInv(lnCtr) + 1
      Next

      For pnCtr = 0 To UBound(lanMCInv)
         lbFirst = True
         For lnCtr = 0 To .ItemCount - 1
            If .Detail(lnCtr, "sMCInvIDx") = lasMCInv(pnCtr) Then
               lrs.AddNew
               If lbFirst Then
                  lrs("nField01").Value = lanMCInv(pnCtr)
                  lrs("sField01").Value = .Detail(lnCtr, "sModelNme")
                  lbFirst = False
               End If
               lrs("sField02").Value = .Detail(lnCtr, "sEngineNo")
               lrs("sField03").Value = .Detail(lnCtr, "sFrameNox")
               lrs("sField04").Value = .Detail(lnCtr, "sCompnyNm")
            End If
         Next
      Next
      
      lsAcsModID = ""
      For pnCtr = 0 To .AccCount - 1
         If InStr(1, lsAcsModID, .Accessory(pnCtr, "sDescript"), vbTextCompare) = 0 Then
            lsAcsModID = lsAcsModID & "�" & .Accessory(pnCtr, "sDescript")
         End If
      Next
      
      lsAcsModID = Mid(lsAcsModID, 2)
      lasAcsMod = Split(lsAcsModID, "�")

      lsIncluded = ""
      lsExcluded = ""
      For pnCtr = 0 To UBound(lasAcsMod)
         lnGivenxxx = 0
         lnQuantity = 0
         For lnCtr = 0 To .AccCount - 1
            If .Accessory(lnCtr, "sDescript") = lasAcsMod(pnCtr) Then
               lnQuantity = lnQuantity + .Accessory(lnCtr, "nQuantity")
               lnGivenxxx = lnGivenxxx + .Accessory(lnCtr, "nGivenxxx")
            End If
         Next
         If lnGivenxxx > 0 Then lsIncluded = lsIncluded & ", " & Trim(Str(lnGivenxxx)) & " " & lasAcsMod(pnCtr)
         If lnQuantity > lnGivenxxx Then
            lsExcluded = lsExcluded & ", " & Trim(Str(lnQuantity - lnGivenxxx)) & " " & lasAcsMod(pnCtr)
         End If
      Next
      
      If lsIncluded <> "" Then lsIncluded = Mid(lsIncluded, 2)
      If lsExcluded <> "" Then lsExcluded = Mid(lsExcluded, 2)
   End With

   'assign important info to the report
   Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\Inter-BranchStockTransfer.rpt")
   oReport.DiscardSavedData
   oReport.FieldMappingType = crAutoFieldMapping
   oReport.Database.SetDataSource lrs

   Set lors = New ADODB.Recordset
   If lors.State = adStateOpen Then lors.Close
   
   lors.Open "SELECT" _
               & "  CONCAT(a.sAddressx, ', ', b.sTownName, ', ', c.sProvName, ' ', b.sZippCode) as Address" _
               & ", d.sCompnyNm" _
            & " From Branch a" _
               & ", TownCity b" _
               & ", Province c" _
               & ", Company d " _
            & " WHERE a.sBranchCd = " & strParm(oTrans.Master("sDestinat")) _
               & " AND a.sTownIDxx = b.sTownIDxx" _
               & " AND b.sProvIDxx = c.sProvIDxx" _
               & " AND a.sCompnyID = d.sCompnyID" _
            , oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   
   oReport.Sections("RH").ReportObjects("txtRefNo").SetText "MC" & "-" & Right(oTrans.Master("sTransNox"), 8)
   oReport.Sections("RH").ReportObjects("txtDate").SetText txtField(1).Text
   oReport.Sections("PH").ReportObjects("txtTo").SetText lors("sCompnyNm")
   oReport.Sections("PH").ReportObjects("txtToAddress").SetText lors("Address")
   oReport.Sections("PH").ReportObjects("txtFrom").SetText oApp.ClientName
   oReport.Sections("PH").ReportObjects("txtFromAddress").SetText oApp.Address & ", " & oApp.TownCity & ", " & oApp.Province & " " & oApp.ZippCode
   oReport.Sections("PF").ReportObjects("txtPrepared").SetText oApp.UserName
   
   If lsIncluded = "" Then
      lsIncluded = lsExcluded
      lsExcluded = ""
   End If
   
   If lsExcluded <> "" Then
      oReport.Sections("RF").ReportObjects("txtNote").SetText "Accessories" & "(" & lsIncluded & " )" _
                                                              & ", " & vbCrLf & "The Items" _
                                                              & "(" & lsExcluded & " ) " & "will follow..." & vbCrLf & txtField(4).Text
   Else
      oReport.Sections("RF").ReportObjects("txtNote").SetText "With Complete Accessories" & "(" & lsIncluded & " )" & vbCrLf & txtField(4).Text
   End If
   
   oReport.PrintOutEx False, 1
   lors.Close

endPoc:
   oTrans.CloseTransaction (oTrans.Master(0))
   Set oReport = Nothing
   Set lrs = Nothing
   Set lors = Nothing
   Exit Function
errProc:
   PrintTrans = False
   MsgBox Err.Description, vbCritical, "Warning"
End Function

Private Sub ClearFields()
   For pnCtr = 0 To txtField.Count - 1
      txtField(pnCtr).Text = ""
   Next
   
   txtField(1).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
   txtField(5).Tag = ""
   With GridEditor1
      .Rows = 2
      
      .ColWidth(3) = 3840
      'empty row
      .TextMatrix(1, 1) = ""
      .TextMatrix(1, 2) = ""
      .TextMatrix(1, 3) = ""
   End With
   pbSave = False
   psBranch = ""
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

'Private Function AcceptAccess() As Boolean
'   Dim lnCtr As Integer
'
'   With oForm
'      .GridEditor1.Rows = oTrans.AccCount + 1
'      For lnCtr = 0 To oTrans.AccCount - 1
'         .GridEditor1.TextMatrix(lnCtr + 1, 1) = oTrans.Accessory(lnCtr, "sModelNme")
'         .GridEditor1.TextMatrix(lnCtr + 1, 2) = oTrans.Accessory(lnCtr, "sDescript")
'         .GridEditor1.TextMatrix(lnCtr + 1, 3) = oTrans.Accessory(lnCtr, "nQtyOnHnd")
'         .GridEditor1.TextMatrix(lnCtr + 1, 4) = oTrans.Accessory(lnCtr, "nQuantity")
'         .GridEditor1.TextMatrix(lnCtr + 1, 5) = oTrans.Accessory(lnCtr, "nGivenxxx")
'      Next
'
'      .Show 1
'
'      If .Cancel = 0 Then
'         For lnCtr = 1 To .GridEditor1.Rows - 1
'            oTrans.Accessory(lnCtr - 1, "sModelNme") = .GridEditor1.TextMatrix(lnCtr, 1)
'            oTrans.Accessory(lnCtr - 1, "sDescript") = .GridEditor1.TextMatrix(lnCtr, 2)
'            oTrans.Accessory(lnCtr - 1, "nQtyOnHnd") = .GridEditor1.TextMatrix(lnCtr, 3)
'            oTrans.Accessory(lnCtr - 1, "nQuantity") = .GridEditor1.TextMatrix(lnCtr, 4)
'            oTrans.Accessory(lnCtr - 1, "nGivenxxx") = .GridEditor1.TextMatrix(lnCtr, 5)
'         Next
'      End If
'
'      AcceptAccess = .Cancel = 0
'   End With
'End Function

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
   Case 1
      If IsDate(txtField(Index).Text) = False Then
         txtField(Index).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
      Else
         txtField(Index).Text = Format(txtField(Index).Text, "MMMM DD, YYYY")
      End If
   Case 3
      txtField(Index).Text = Format(txtField(Index).Text, ">")
   Case 5
      If GetFocus = pnCancelFocus Then Exit Sub
      If txtField(Index).Text = "" Then
         ClearFields
      Else
         If txtField(Index).Text <> txtField(Index).Tag Then SearchBranch True, txtField(Index).Text
      End If
      txtField(Index).Tag = txtField(Index).Text
   End Select
   
   If Index = 1 Then
      oTrans.Master(Index) = txtField(Index).Text & " " & Format(oApp.ServerDate, "hh:mm:ss AM/PM")
   Else
      If Index <> 5 Then oTrans.Master(Index) = txtField(Index).Text
   End If
End Sub

Private Function isEntryOK() As Boolean
   If txtField(5).Text = "" Then
      MsgBox "Branch not found!!!", vbCritical, "Warning"
      txtField(5).SetFocus
      GoTo EntryNotOK
   End If
   
   If txtField(2).Text = "" Then
      MsgBox "Destination not found!!!", vbCritical, "Warning"
      txtField(2).SetFocus
      GoTo EntryNotOK
   End If

'   If txtField(3).Text = "" Then
'      MsgBox "Unknown Transmittal Number!!!", vbCritical, "Warning"
'      txtField(3).SetFocus
'      GoTo EntryNotOK
'   End If

   With GridEditor1
      If Trim(.TextMatrix(1, 1)) = "" Then
         MsgBox "Detail is required!!!", vbCritical, "Warning"
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

Private Sub SearchBranch(ByVal Search As Boolean, Optional Branch As Variant)
   Dim lrs As ADODB.Recordset
   Dim lsBrowse As String
   Dim lsSelected() As String
   Dim lsSQL As String
   
   lsSQL = "Select" _
               & "  sBranchCd" _
               & ", sBranchNm" _
            & " From Branch" _
            & " Where cRecdStat = " & strParm(xeRecStateActive)
               
   If Not Search Then
      If Not IsMissing(Branch) Then lsSQL = lsSQL & " And sBranchNm LIKE " & strParm(Branch & "%")
   Else
      lsSQL = lsSQL & " And sBranchNm = " & strParm(Branch)
   End If
   
   Set lrs = New ADODB.Recordset
   
   lrs.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   
   If lrs.EOF Then
      txtField(5).Text = ""
      psBranch = ""
   ElseIf lrs.RecordCount = 1 Then
      txtField(5).Text = lrs("sBranchNm")
      psBranch = lrs("sBranchCd")
   Else
      lsBrowse = KwikBrowse(oApp, lrs _
                           , "sBranchCd�sBranchNm" _
                           , "BranchCd�Branch Name" _
                           , , False)
      
      If lsBrowse <> "" Then
         lsSelected = Split(lsBrowse, "�")
         psBranch = lsSelected(0)
         txtField(5).Text = lsSelected(1)
      End If
   End If
   
   txtField(5).Tag = txtField(5).Text
   txtField(5).SelStart = 0
   txtField(5).SelLength = Len(txtField(5).Text)
   lrs.Close
   
   If psBranch = "" Then txtField(5).Text = ""
   Set lrs = Nothing
   
   oTrans.Branch = Format(psBranch, "00")
   If txtField(5).Text = "" Then ClearFields
   oTrans.InitTransaction
   oTrans.NewTransaction
   
   txtField(0).Text = Format(oTrans.Master("sTransNox"), "@@-@@@@@@@@")
   If txtField(5).Text = "" Then ClearFields
   If txtField(5) <> "" Then
      xrFrame1(0).Enabled = True
         xrFrame1(1).Enabled = False
   Else
      xrFrame1(0).Enabled = False
      xrFrame1(1).Enabled = True
   End If
End Sub

