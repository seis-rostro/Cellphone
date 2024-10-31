VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0"; "xrGridControl.ocx"
Begin VB.Form frmCP_Branch_Transfer_Reg 
   BorderStyle     =   0  'None
   Caption         =   "Delivery"
   ClientHeight    =   7800
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
   ScaleHeight     =   7800
   ScaleWidth      =   11715
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   8
      Left            =   10365
      TabIndex        =   17
      Top             =   1785
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
      Picture         =   "frmCP_Branch_Transfer_Reg.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   6
      Left            =   10365
      TabIndex        =   24
      Top             =   4290
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
      Picture         =   "frmCP_Branch_Transfer_Reg.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   9
      Left            =   10365
      TabIndex        =   21
      Top             =   2415
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
      Picture         =   "frmCP_Branch_Transfer_Reg.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   10365
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
      Picture         =   "frmCP_Branch_Transfer_Reg.frx":166E
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   10365
      TabIndex        =   15
      Top             =   525
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
      Picture         =   "frmCP_Branch_Transfer_Reg.frx":1DE8
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   7
      Left            =   10365
      TabIndex        =   16
      Top             =   1155
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
      Picture         =   "frmCP_Branch_Transfer_Reg.frx":2562
   End
   Begin xrGridEditor.GridEditor GridEditor1 
      Height          =   4380
      Left            =   105
      TabIndex        =   14
      Tag             =   "et0;eb0;et0;bc2"
      Top             =   3315
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   7726
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
      Object.HEIGHT          =   4380
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
      MOUSEICON       =   "frmCP_Branch_Transfer_Reg.frx":2CDC
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
      Height          =   2190
      Index           =   0
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   1095
      Width           =   10050
      _ExtentX        =   17727
      _ExtentY        =   3863
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.CheckBox chkField 
         Caption         =   "Save to &Mobile Disk"
         Height          =   195
         Left            =   5040
         TabIndex        =   26
         Tag             =   "et0;fb0"
         Top             =   120
         Width           =   1725
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   5055
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   660
         Width           =   4770
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   1365
         TabIndex        =   7
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
         TabIndex        =   13
         Text            =   "frmCP_Branch_Transfer_Reg.frx":2CF8
         Top             =   1320
         Width           =   8460
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   1365
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   165
         Width           =   1950
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   3
         Left            =   1365
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   990
         Width           =   2505
      End
      Begin VB.Shape Shape3 
         Height          =   390
         Index           =   0
         Left            =   7320
         Top             =   120
         Width           =   2505
      End
      Begin VB.Shape Shape4 
         Height          =   330
         Index           =   0
         Left            =   7350
         Top             =   150
         Width           =   2445
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
         Left            =   7380
         TabIndex        =   27
         Tag             =   "eb0;et0"
         Top             =   195
         Width           =   2385
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transact. No."
         Height          =   285
         Index           =   9
         Left            =   165
         TabIndex        =   4
         Top             =   210
         Width           =   1185
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transact. Date"
         Height          =   285
         Index           =   1
         Left            =   165
         TabIndex        =   6
         Top             =   690
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transmittal No"
         Height          =   285
         Index           =   3
         Left            =   195
         TabIndex        =   10
         Top             =   1065
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   285
         Index           =   7
         Left            =   585
         TabIndex        =   12
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
         TabIndex        =   8
         Top             =   690
         Width           =   1200
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H8000000F&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   285
         Index           =   0
         Left            =   7380
         Tag             =   "et0;et0"
         Top             =   180
         Width           =   2400
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   525
      Index           =   1
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   10050
      _ExtentX        =   17727
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
         Left            =   4590
         MaxLength       =   50
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   90
         Width           =   5250
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
         Top             =   90
         Width           =   1950
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Branch"
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
         Left            =   3840
         TabIndex        =   2
         Top             =   135
         Width           =   705
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
         Left            =   105
         TabIndex        =   0
         Top             =   135
         Width           =   1365
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   10365
      TabIndex        =   18
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
      Picture         =   "frmCP_Branch_Transfer_Reg.frx":2D0E
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   10365
      TabIndex        =   25
      Top             =   4290
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
      Picture         =   "frmCP_Branch_Transfer_Reg.frx":3488
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   10365
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
      Picture         =   "frmCP_Branch_Transfer_Reg.frx":3C02
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   10365
      TabIndex        =   22
      Top             =   3045
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
      Picture         =   "frmCP_Branch_Transfer_Reg.frx":437C
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   10
      Left            =   10365
      TabIndex        =   23
      Top             =   3660
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Freight"
      AccessKey       =   "F"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCP_Branch_Transfer_Reg.frx":4AF6
   End
End
Attribute VB_Name = "frmCP_Branch_Transfer_Reg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmCP_Branch_Transfer"

Private WithEvents oTrans As clsCPTransfer
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin

Dim pnIndex As Integer
Dim pbGridFocus As Boolean
Dim pnCtr As Integer
Dim pbSave As Boolean
Dim pbEditMode As Boolean
Dim pbGridValidate As Boolean
Dim pbClosedTrans As Boolean

Private Sub chkField_Click()
'   oTrans.DiskTransaction = IIf(chkField.Value = 1, True, False)
End Sub

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   Dim lsRep As String
   Dim lsUserID As String
   Dim lsUserName As String
   Dim lnUserRights As Integer
   Dim lasRights() As String
   
   lsOldProc = "cmdButton_Click"
   ''On Error GoTo errProc
   
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

            .ColWidth(3) = 3100
            If .Rows > 16 Then .ColWidth(3) = 2850
         End If
               
         If isEntryOk Then
            If oTrans.SaveTransaction Then
               If oTrans.OpenTransaction(oTrans.Master("sTransNox")) Then
                  MsgBox "Record Updated Successfully!!!", vbInformation, "Notice"
                  initButton xeModeReady

                  pbEditMode = False
                  txtField(5).SetFocus
               End If
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
            
            .ColWidth(3) = 3470
            If .Rows > 19 Then .ColWidth(3) = 3170
         End If
      Case 3
         lsRep = MsgBox("Transaction is in Update Mode!!!" & vbCrLf & _
                        "Do you want to Cancel Transaction!!!", vbYesNo + vbQuestion, "Confirm")
         
         If lsRep = vbYes Then
            If oTrans.OpenTransaction(oTrans.Master(0)) Then
               LoadMaster
               LoadDetail
            Else
               If txtField(0).Text = "" Then ClearFields
            End If
         
            initButton xeModeReady
            txtField(6).SetFocus
            pbEditMode = False
         Else
            txtField(pnIndex).SetFocus
         End If
      Case 4
         If oTrans.SearchTransaction() Then
            LoadMaster
            LoadDetail
         Else
            If txtField(0).Text = "" Then ClearFields
         End If
      Case 5
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
      Case 6
         Unload Me
      Case 7
'  she
'         If txtField(0).Text <> "" Then
'            If oTrans.UpdateTransaction Then
'               InitButton xeModeUpdate
'               pbEditMode = True
'               txtField(1).SetFocus
'            End If
'         Else
'            MsgBox "No Transaction to Update!!!" & vbCrLf & _
'                   "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
'         End If
      Case 8
         If txtField(0).Text <> "" Then
            lsRep = MsgBox("Are you sure you want to Cancel Transaction!!!", vbYesNo + vbQuestion, "Confirm")
            If lsRep = vbNo Then GoTo endProc
'            lasRights = Split(oApp.mdiMain.Controls(oApp.MenuName).Tag, "»")
'            If GetApproval(oApp, lnUserRights, lsUserID, lsUserName, lasRights(0)) = False Then GoTo endProc
            
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
         Else
            MsgBox "No Transaction to Cancel!!!", vbInformation, "Notice"
         End If
      Case 9
         If pbEditMode = False Then
            If txtField(0).Text <> "" Then
               lsRep = MsgBox("Print Transaction!!!", vbYesNo + vbQuestion, "Confirm")
               If lsRep = vbYes Then
                 If Not PrintTrans Then MsgBox "Unable to Print Transaction!!!", vbCritical, "Warning"
               End If
            End If
         Else
            MsgBox "Unable to Print Transaction!!!" & vbCrLf & _
                   "Save Transaction first to continue printing!!!", vbCritical, "Warning"
         End If
      Case 10
         If Not oApp.IsWarehouse Then Exit Sub
         If txtField(0).Text <> "" Then
            lsRep = MsgBox("Print Transaction!!!", vbYesNo + vbQuestion, "Confirm")
            If lsRep = vbYes Then
              If Not FreightTrans Then MsgBox "Unable to Print Transaction!!!", vbCritical, "Warning"
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

   Set oTrans = New clsCPTransfer
   Set oTrans.AppDriver = oApp
   
'   oTrans.DiskTransaction = False
   oTrans.InitTransaction
   oTrans.NewTransaction
   
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransMaintenance
   
   InitGrid
   ClearFields
   initButton xeModeReady
   
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

Private Sub GridEditor1_AddingRow(Cancel As Boolean)
   With GridEditor1
      If .TextMatrix(.Row, 1) = "" Then
         Cancel = True
      ElseIf .TextMatrix(.Row, 6) = "0" Then
         Cancel = True
      End If
      If Not Cancel Then
         If .Row = .Rows - 1 Then oTrans.addDetail
      End If
 
      If .Rows > 20 Then .ColWidth(3) = 2850
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
               oTrans.Detail(.Row - 1, "nQuantity") = 1
               .TextMatrix(.Row, 6) = oTrans.Detail(.Row - 1, "nQuantity")
               If .Row = .Rows - 1 Then
                  .Rows = .Rows + 1
                  oTrans.addDetail
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
         .ColWidth(3) = 2850
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
         If oTrans.searchDetail(.Row - 1, .Col, .TextMatrix(.Row, .Col)) Then
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
            .ColWidth(3) = 2850
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
   cmdButton(7).Visible = Not lbShow
   cmdButton(8).Visible = Not lbShow
   xrFrame1(1).Enabled = Not lbShow
   
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
   xrFrame1(0).Enabled = lbShow
End Sub

Private Function PrintTrans() As Boolean
   Dim loreport As frmRepViewer
   Dim lrs As Recordset
   Dim lors As Recordset
   Dim lnCtr As Integer
   Dim lsOldProc As String
   Dim lnTotlWSerial As Double
   Dim lnTotlWOSerial As Double
   
   lsOldProc = "PrinTrans"
   ''On Error GoTo errProc

   PrintTrans = True
   
   Set lrs = New ADODB.Recordset

   lrs.Fields.Append "nField01", adInteger, 3
   lrs.Fields.Append "nField02", adChar, 1
   lrs.Fields.Append "sField01", adVarChar, 20
   lrs.Fields.Append "sField02", adVarChar, 128
   lrs.Fields.Append "sField03", adVarChar, 30
   lrs.Fields.Append "sField04", adVarChar, 12
   lrs.Fields.Append "sField05", adVarChar, 100
   lrs.Fields.Append "lField01", adCurrency
   lrs.Open

   With oTrans
      lnTotlWOSerial = 0
      lnTotlWSerial = 0
      
      For lnCtr = 0 To .ItemCount - 1
         lrs.AddNew
         lrs.Fields("nField01") = oTrans.Detail(lnCtr, "nQuantity")
         lrs.Fields("nField02") = oTrans.Detail(lnCtr, "cHsSerial")
         lrs.Fields("sField01") = oTrans.Detail(lnCtr, "sBarrCode")
         lrs.Fields("sField02") = oTrans.Detail(lnCtr, "sDescript")
         lrs.Fields("sField03") = IFNull(oTrans.Detail(lnCtr, "sSerialNo"), "")
         lrs.Fields("sField04") = oTrans.Detail(lnCtr, "sStockIDx")
         lrs.Fields("sField05") = oTrans.Detail(lnCtr, "sBrandNme")
         lrs.Fields("lField01") = oTrans.Detail(lnCtr, "nSelPrice")
         
         If oTrans.Detail(lnCtr, "cHsSerial") = xeYes Then
            lnTotlWSerial = lnTotlWSerial + 1
         Else
            lnTotlWOSerial = lnTotlWOSerial + CDbl(oTrans.Detail(lnCtr, "nQuantity"))
         End If
      Next
      lrs.Sort = "nField02 DESC,sField05,sField03"
   End With

   'assign important info to the report
   Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\CP_Transfer.rpt")
   oReport.DiscardSavedData
   oReport.FieldMappingType = crAutoFieldMapping
   oReport.Database.SetDataSource lrs

   Set lors = New ADODB.Recordset
   If lors.State = adStateOpen Then lors.Close

   lors.Open "SELECT" _
               & "  a.sAddressx" _
               & ", CONCAT(b.sTownName, ', ' , c.sProvName, ' ' , b.sZippCode) xTownName" _
               & ", a.sBranchNm" _
            & " FROM Branch a" _
               & " LEFT JOIN TownCity b" _
                  & " LEFT JOIN Province c" _
                     & " ON b.sProvIDxx = c.sProvIDxx" _
                  & " ON a.sTownIDxx = b.sTownIDxx" _
            & " WHERE a.sBranchCd = " & strParm(oTrans.Master("sDestinat")) _
            , oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText

   oReport.Sections("PHa").ReportObjects("txtRefNo").SetText "CP" & "-" & Right(oTrans.Master("sTransNox"), 10)
   oReport.Sections("PHa").ReportObjects("txtDate").SetText txtField(1).Text
   oReport.Sections("PHb").ReportObjects("txtTo").SetText lors("sBranchNm")
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

   If oTrans.Master("cTranStat") = xeStateOpen Then
      If oTrans.CloseTransaction(oTrans.Master(0)) Then pbClosedTrans = True
   End If

   Set loreport = Nothing
   Set oReport = Nothing
   Set lrs = Nothing
   Set lors = Nothing
   Exit Function
errProc:
   PrintTrans = False
   ShowError lsOldProc & "( " & " )"
End Function

Private Function FreightTrans() As Boolean
   Dim loreport As frmRepViewer
   Dim lrs As Recordset
   Dim lors As Recordset
   Dim lrsTemp As Recordset
   Dim lnCtr As Integer
   Dim lsOldProc As String
   Dim lnTotlWSerial As Double
   Dim lnTotlWOSerial As Double
   Dim lsSQL As String
   
   lsOldProc = "FreightTrans"
   ''On Error GoTo errProc

   FreightTrans = True
   
   Set lrs = New ADODB.Recordset

   lrs.Fields.Append "nField01", adInteger, 3
   lrs.Fields.Append "nField02", adChar, 1
   lrs.Fields.Append "sField01", adVarChar, 20
   lrs.Fields.Append "sField02", adVarChar, 128
   lrs.Fields.Append "sField03", adVarChar, 30
   lrs.Fields.Append "sField04", adVarChar, 12
   lrs.Fields.Append "sField05", adVarChar, 100
   lrs.Fields.Append "lField01", adCurrency
   lrs.Fields.Append "lField02", adCurrency
   lrs.Open

   With oTrans
      lnTotlWOSerial = 0
      lnTotlWSerial = 0
      
      For lnCtr = 0 To .ItemCount - 1
         lrs.AddNew
         lrs.Fields("nField01") = oTrans.Detail(lnCtr, "nQuantity")
         lrs.Fields("nField02") = oTrans.Detail(lnCtr, "cHsSerial")
         lrs.Fields("sField01") = oTrans.Detail(lnCtr, "sBarrCode")
         lrs.Fields("sField02") = oTrans.Detail(lnCtr, "sDescript")
         lrs.Fields("sField03") = IFNull(oTrans.Detail(lnCtr, "sSerialNo"), "")
         lrs.Fields("sField04") = oTrans.Detail(lnCtr, "sStockIDx")
         lrs.Fields("sField05") = oTrans.Detail(lnCtr, "sBrandNme")
         
         lsSQL = "SELECT" & _
                     "  SUM(xUnitPrce) xUnitPrce" & _
                     ", SUM(yUnitPrce) yUnitPrce" & _
                     ", SUM(zUnitPrce) zUnitPrce" & _
                        " FROM (SELECT" & _
                        "  b.sStockIDx" & _
                        ", b.nUnitPrce xUnitPrce" & _
                        ", '0' yUnitPrce" & _
                        ", '0' zUnitPrce" & _
                     " FROM CP_PO_Receiving_Serial a" & _
                           ", CP_PO_Receiving_Detail b" & _
                     " WHERE a.sSerialID = " & strParm(oTrans.Detail(lnCtr, "sSerialID")) & _
                        " AND a.sTransNox = b.sTransNox" & _
                        " AND a.nEntryNox = b.nEntryNox"
         lsSQL = lsSQL & _
                     " UNION" & _
                        " (SELECT" & _
                              "  a.sStockIDx" & _
                              ", '0' xUnitPrce" & _
                              ", a.nUnitPrce yUnitPrce" & _
                              ", '0' zUnitPrce" & _
                           " FROM CP_PO_Receiving_Detail a" & _
                              ", CP_PO_Receiving_Master b" & _
                           " WHERE a.sStockIDx = " & strParm(oTrans.Detail(lnCtr, "sStockIDx")) & _
                              " AND a.nUnitPrce > 0" & _
                              " AND a.sTransNox = b.sTransNox" & _
                           " ORDER BY b.dTransact DESC" & _
                           " LIMIT 1)" & _
                     " UNION" & _
                        " SELECT" & _
                              "  sStockIDx" & _
                              ", '0' xUnitPrce" & _
                              ", '0' yUnitPrce" & _
                              ", nSelPrice yUnitPrce" & _
                           " FROM CP_Inventory" & _
                           " WHERE sStockIDx = " & strParm(oTrans.Detail(lnCtr, "sStockIDx")) & ") xSourceTable" & _
                     " GROUP BY sStockIDx"
         
         Set lrsTemp = New Recordset
         lrsTemp.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
         
         If oTrans.Detail(lnCtr, "cHsSerial") = xeYes Then
            lnTotlWSerial = lnTotlWSerial + 1
            If Not lrsTemp.EOF Then
               If lrsTemp("xUnitPrce") > 0 Then
                  lrs.Fields("lField01") = lrsTemp("xUnitPrce")
               ElseIf lrsTemp("yUnitPrce") > 0 Then
                  lrs.Fields("lField01") = lrsTemp("yUnitPrce")
               ElseIf lrsTemp("zUnitPrce") > 0 Then
                  lrs.Fields("lField01") = lrsTemp("zUnitPrce")
               End If
            End If
         Else
            lnTotlWOSerial = lnTotlWOSerial + CDbl(oTrans.Detail(lnCtr, "nQuantity"))
            If Not lrsTemp.EOF Then
               If lrsTemp("yUnitPrce") > 0 Then
                  lrs.Fields("lField02") = lrsTemp("yUnitPrce")
               ElseIf lrsTemp("zUnitPrce") > 0 Then
                  lrs.Fields("lField02") = lrsTemp("zUnitPrce")
               End If
            End If
         End If
      Next
      lrs.Sort = "nField02 DESC,sField05,sField03"
   End With

   'assign important info to the report
   Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\CP_Transfer_Freight.rpt")
   oReport.DiscardSavedData
   oReport.FieldMappingType = crAutoFieldMapping
   oReport.Database.SetDataSource lrs

   Set lors = New ADODB.Recordset
   If lors.State = adStateOpen Then lors.Close

   lors.Open "SELECT" _
               & "  a.sAddressx" _
               & ", CONCAT(b.sTownName, ', ' , c.sProvName, ' ' , b.sZippCode) xTownName" _
               & ", a.sBranchNm" _
            & " FROM Branch a" _
               & " LEFT JOIN TownCity b" _
                  & " LEFT JOIN Province c" _
                     & " ON b.sProvIDxx = c.sProvIDxx" _
                  & " ON a.sTownIDxx = b.sTownIDxx" _
            & " WHERE a.sBranchCd = " & strParm(oTrans.Master("sDestinat")) _
            , oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText

   oReport.Sections("PHa").ReportObjects("txtRefNo").SetText "CP" & "-" & Right(oTrans.Master("sTransNox"), 10)
   oReport.Sections("PHa").ReportObjects("txtDate").SetText txtField(1).Text
   oReport.Sections("PHb").ReportObjects("txtTo").SetText lors("sBranchNm")
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
   
   FreightTrans = True

endPoc:
   Set loreport = Nothing
   Set oReport = Nothing
   Set lrs = Nothing
   Set lors = Nothing
   Exit Function
errProc:
   FreightTrans = False
   ShowError lsOldProc & "( " & " )"
End Function

Private Sub ClearFields()
   For pnCtr = 0 To txtField.Count - 1
      Select Case pnCtr
      Case 1
         txtField(pnCtr).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
      Case Else
         txtField(pnCtr).Text = ""
         txtField(pnCtr).Tag = ""
      End Select
   Next

   With GridEditor1
      .Rows = 2
      .ColWidth(3) = 3100
      
      'empty row
      .TextMatrix(1, 1) = ""
      .TextMatrix(1, 2) = ""
      .TextMatrix(1, 3) = "0.00"
      .TextMatrix(1, 6) = "0"
   End With
   
   chkField.Value = 0
   pbSave = False
   pbClosedTrans = False
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
      Case 5, 6
         If .Text = "" Then
            ClearFields
            Exit Sub
         End If
         
         If Trim(LCase(.Text)) <> Trim(LCase(.Tag)) Then
            If oTrans.SearchTransaction _
            (IIf(Index = 5, CodeFormat(oApp.BranchCode, .Text), .Text) _
            , IIf(Index = 5, True, False)) Then
               LoadMaster
               LoadDetail
            Else
               ClearFields
               .SetFocus
            End If
         End If
      End Select
      
      If Index < 5 Then oTrans.Master(Index) = .Text
   End With
End Sub

Private Function isEntryOk() As Boolean
   If txtField(2).Text = "" Then
      MsgBox "Destination not found!!!" & vbCrLf & _
             "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
      txtField(2).SetFocus
      GoTo EntryNotOK
   End If
   
   With GridEditor1
      If Trim(.TextMatrix(1, 1)) = "" Then
         MsgBox "Detail is required!!!" & vbCrLf & _
                "Please Verify your Entry then Try Again", vbCritical, "Warning"
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

Private Sub LoadMaster()
   For pnCtr = 0 To txtField.Count - 1
      Select Case pnCtr
      Case 0, 5
         txtField(pnCtr).Text = Format(oTrans.Master(0), "@@@@-@@@@@@")
         txtField(pnCtr).Tag = txtField(pnCtr).Text
      Case 1
         txtField(pnCtr).Text = Format(oTrans.Master(pnCtr), "MMMM DD, YYYY")
      Case 2, 6
         txtField(pnCtr).Text = oTrans.Master(2)
         txtField(pnCtr).Tag = txtField(pnCtr).Text
      Case Else
         txtField(pnCtr).Text = IIf(IsNull(oTrans.Master(pnCtr)), "", oTrans.Master(pnCtr))
      End Select
   Next
   
   Label2.Caption = Format(TransStat(oTrans.Master("cTranStat")), ">")
   If oTrans.Master("cTranStat") = xeStateClosed Then pbClosedTrans = True
   chkField.Value = 0
   pbEditMode = False
End Sub

Private Sub LoadDetail()
   Dim lnCtr As Integer
   
   With GridEditor1
      .Rows = IIf(oTrans.ItemCount = 0, 2, oTrans.ItemCount + 1)
      
      .ColWidth(3) = 3100
      If .Rows > 16 Then .ColWidth(3) = 2850
      
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

Private Function BranchAutomate(ByVal sBranchCd As String) As Boolean
   Dim lrs As Recordset
   
   Set lrs = New Recordset
   lrs.Open "SELECT * FROM Branch" & _
               " WHERE sBranchCd = " & strParm(sBranchCd) & _
                  " AND cAutomate = " & strParm(xeYes) _
   , oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   
   If Not lrs.EOF Then BranchAutomate = True
   Set lrs = Nothing
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
