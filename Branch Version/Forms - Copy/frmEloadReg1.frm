VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmEloadReg1 
   BorderStyle     =   0  'None
   Caption         =   "CP Load"
   ClientHeight    =   6090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6915
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6090
   ScaleWidth      =   6915
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   525
      Index           =   1
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   926
      BackColor       =   12632256
      ClipControls    =   0   'False
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
         Index           =   9
         Left            =   3165
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   105
         Width           =   1980
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
         Index           =   8
         Left            =   1320
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   105
         Width           =   1215
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ref. #"
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
         Index           =   3
         Left            =   2565
         TabIndex        =   2
         Top             =   120
         Width           =   555
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   0
         Top             =   120
         Width           =   1185
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   4860
      Index           =   0
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   1095
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   8573
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   1320
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   705
         Width           =   1905
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   7
         Left            =   2850
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   3045
         Width           =   2070
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   6
         Left            =   2850
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   2535
         Width           =   2070
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   1320
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   1035
         Width           =   1905
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   1320
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   150
         Width           =   1215
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   3
         Left            =   1320
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   1365
         Width           =   3720
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   4
         Left            =   1320
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   1695
         Width           =   3720
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   5
         Left            =   1320
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   2025
         Width           =   3720
      End
      Begin VB.Shape Shape4 
         Height          =   330
         Index           =   0
         Left            =   180
         Top             =   4275
         Width           =   2445
      End
      Begin VB.Shape Shape3 
         Height          =   390
         Index           =   0
         Left            =   150
         Top             =   4245
         Width           =   2505
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
         Left            =   210
         TabIndex        =   29
         Tag             =   "eb0;et0"
         Top             =   4305
         Width           =   2385
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H8000000F&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   285
         Index           =   0
         Left            =   210
         Tag             =   "et0;et0"
         Top             =   4305
         Width           =   2430
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transact Date"
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   6
         Top             =   720
         Width           =   1020
      End
      Begin VB.Shape Shape2 
         Height          =   1815
         Index           =   1
         Left            =   150
         Top             =   2415
         Width           =   4875
      End
      Begin VB.Label lblChangeAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   19.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   2850
         TabIndex        =   21
         Top             =   3480
         Width           =   2085
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Change Amt."
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
         Index           =   10
         Left            =   1320
         TabIndex        =   20
         Top             =   3540
         Width           =   1095
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amt. Tendered"
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
         Index           =   8
         Left            =   1320
         TabIndex        =   18
         Top             =   3075
         Width           =   1260
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Load Amount"
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
         Index           =   6
         Left            =   1320
         TabIndex        =   16
         Top             =   2580
         Width           =   1125
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reference No."
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   8
         Top             =   1080
         Width           =   1050
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   315
         Left            =   1410
         Tag             =   "et0;ht2"
         Top             =   255
         Width           =   1215
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transact. No."
         Height          =   195
         Index           =   11
         Left            =   90
         TabIndex        =   4
         Top             =   210
         Width           =   975
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Barcode"
         Height          =   195
         Index           =   9
         Left            =   150
         TabIndex        =   10
         Top             =   1425
         Width           =   600
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   7
         Left            =   150
         TabIndex        =   12
         Top             =   1755
         Width           =   795
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile Number"
         Height          =   195
         Index           =   5
         Left            =   150
         TabIndex        =   14
         Top             =   2085
         Width           =   1065
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   5565
      TabIndex        =   23
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
      Picture         =   "frmEloadReg1.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   5565
      TabIndex        =   22
      Top             =   540
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
      Picture         =   "frmEloadReg1.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   6
      Left            =   5565
      TabIndex        =   25
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
      Picture         =   "frmEloadReg1.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   5565
      TabIndex        =   24
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
      Picture         =   "frmEloadReg1.frx":166E
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   5565
      TabIndex        =   26
      Top             =   1185
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
      Picture         =   "frmEloadReg1.frx":1DE8
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   5565
      TabIndex        =   28
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
      Picture         =   "frmEloadReg1.frx":2562
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   5565
      TabIndex        =   27
      Top             =   1815
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
      Picture         =   "frmEloadReg1.frx":2CDC
   End
End
Attribute VB_Name = "frmEloadReg1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmEloadReg"

Private WithEvents oTrans As clsCPLoad
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin

Dim psTransNox As String
Dim pnIndex As Integer
Dim pnCtr As Integer
Dim pbCancel As Boolean

Property Get loadTrans() As clsCPLoad
   Set loadTrans = oTrans
End Property

Property Get Cancelled() As Boolean
   Cancelled = pbCancel
End Property

Property Let TransNox(lsTransNox As String)
   psTransNox = lsTransNox
End Property

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   Dim lnRep As Integer

   lsOldProc = "cmdButton_Click"
   'On Error GoTo errProc

   txtField_LostFocus pnIndex
   Select Case Index
   Case 0 'Save
      If isEntryOk Then
         If oTrans.SaveTransaction Then
            Me.Hide
            pbCancel = False
         Else
            MsgBox "Unable to Save Transaction!!!", vbCritical, "Warning"
         End If
      End If
   Case 1 'Search
         If pnIndex = 5 Or pnIndex = 6 Then
            oTrans.searchBarrcode pnIndex, ""
            txtField(pnIndex).SetFocus
         End If
         txtField(pnIndex).SetFocus
   Case 2 'Cancel Edit
      Me.Hide
      pbCancel = True
   Case 3 'Browse
      If oTrans.SearchTransaction() Then
         LoadMaster
      Else
         If txtField(0).Text = "" Then ClearFields
      End If

      txtField(pnIndex).SetFocus
   Case 4 'Update
      If Not oTrans.UpdateTransaction Then Exit Sub
      initButton xeModeUpdate
      txtField(2).SetFocus
   Case 5 'Cancel Trans/Delete
      If txtField(0).Text <> "" Then
         If oTrans.CancelTransaction Then
            MsgBox "Transaction Cancelled Successfully!!!", vbInformation, "Notice"
            If oTrans.OpenTransaction(oTrans.Master("sTransNox")) Then LoadMaster
         Else
            MsgBox "Unable to Cancel Transaction!!!", vbCritical, "Warning"
         End If
      Else
         MsgBox "No Transaction to Cancel!!!" & vbCrLf & _
               "Please Verify your entry then try again!!!", vbCritical, "Warning"
      End If
   Case 6 'Close
      Unload Me
   End Select

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0
   
   If oTrans.OpenTransaction(psTransNox) Then
      Call LoadMaster
      Call cmdButton_Click(4)
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyReturn, vbKeyUp, vbKeyDown
      Select Case KeyCode
      Case vbKeyReturn, vbKeyDown
         SetNextFocus
      Case vbKeyUp
         SetPreviousFocus
      End Select
   Case vbKeyF1
   Case vbKeyF5
'      If Not isEntryOK Then Exit Sub
'      Call txtField_Validate(pnIndex, False)
'      If Not oTrans.SaveTransaction Then
'         MsgBox "Unable to Save Transaction!!!" & vbCrLf & _
'                  "Please contanct GMC/GGC SEG for assistance!!!", vbCritical, "Warning"
'         Exit Sub
'      End If
'
'      oTrans.NewTransaction
'      ClearFields
'      txtField(1).SetFocus
   Case vbKeyEscape
      Unload Me
   End Select
End Sub

Private Sub Form_Load()
   CenterChildForm mdiMain, Me
   
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransMaintenance
   
   Set oTrans = New clsCPLoad
   Set oTrans.AppDriver = oApp

   oTrans.InitTransaction
   oTrans.NewTransaction
   
   initButton xeModeReady
   ClearFields
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
   Set oTrans = Nothing
End Sub

Private Sub ClearFields()
   Dim lnCtr As Integer
   
   For lnCtr = 0 To txtField.Count - 1
      Select Case lnCtr
      Case 2
         txtField(lnCtr).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
      Case 6, 7
         txtField(lnCtr).Text = "0.00"
      Case Else
         txtField(lnCtr).Text = ""
         txtField(lnCtr).Tag = ""
      End Select
   Next
   
   lblChangeAmount.Caption = "0.00"
   Label2.Caption = "UNKNOWN"
End Sub

Private Sub initButton(lnStat As Integer)
   Dim lbShow As Boolean

   lbShow = IIf(lnStat = 0, False, True)
   cmdButton(3).Visible = Not lbShow
   cmdButton(4).Visible = Not lbShow
   cmdButton(5).Visible = Not lbShow
   cmdButton(6).Visible = Not lbShow
   
   xrFrame1(1).Enabled = Not lbShow
   
   cmdButton(0).Visible = lbShow
   cmdButton(1).Visible = lbShow
   cmdButton(2).Visible = lbShow
   
   xrFrame1(0).Enabled = lbShow
End Sub

Private Sub oTrans_MasterRetrieved(ByVal Index As Integer)
   With oTrans
      Select Case Index
      Case 3
         txtField(3).Text = oTrans.Master("sBarrCode")
      Case 4
         txtField(4).Text = oTrans.Master("sDescript")
      End Select
   End With
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = oApp.getColor("HT1")
   End With
   
   pnIndex = Index
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String
   
   lsOldProc = "txtField_KeyDown"
   'On Error GoTo errProc
   
   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
      With txtField(pnIndex)
         Select Case Index
         Case 3
            If KeyCode = vbKeyF3 Then
               If oTrans.searchBarrcode(Index, .Text) Then
                  If .Text <> "" Then SetNextFocus
               End If
            Else
               If .Text <> "" Then oTrans.searchBarrcode Index, .Text
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

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   Dim lnChange As Currency
   
   With txtField(Index)
      Select Case Index
      Case 1
         .Text = Format(.Text, ">")
      Case 2
         If Not IsDate(.Text) Then .Text = oApp.ServerDate
         .Text = Format(.Text, "MMMM DD, YYYY")
      Case 6, 7
         If Not IsNumeric(.Text) Then .Text = 0#
         .Text = Format(.Text, "#,##0.00")
         lnChange = CDbl(txtField(7).Text) - CDbl(txtField(6).Text)
         lblChangeAmount.Caption = Format(IIf(lnChange > 0#, lnChange, 0#), "#,##0.00")
      End Select
      If Index < 8 Then oTrans.Master(Index) = .Text
   End With
End Sub

Private Sub LoadMaster()
   For pnCtr = 0 To txtField.Count - 1
      Select Case pnCtr
      Case 0, 8
         txtField(pnCtr).Text = Format(oTrans.Master("sTransNox"), "@@@@-@@@@@@@@")
         txtField(pnCtr).Tag = txtField(pnCtr).Text
      Case 2
         txtField(pnCtr).Text = Format(oTrans.Master(pnCtr), "MMMM DD, YYYY")
      Case 1, 9
         txtField(pnCtr).Text = oTrans.Master("sReferNox")
         txtField(pnCtr).Tag = txtField(pnCtr).Text
      Case 6, 7
         txtField(pnCtr).Text = Format(oTrans.Master(pnCtr), "#,##0.00")
      Case Else
         txtField(pnCtr).Text = IIf(IsNull(oTrans.Master(pnCtr)), "", oTrans.Master(pnCtr))
      End Select
   Next
   
   lblChangeAmount.Caption = Format(CDbl(txtField(7).Text) - CDbl(txtField(6).Text), "#,##0.00")
   Label2.Caption = Format(TransStat(oTrans.Master("cTranStat")), ">")
End Sub

Private Function isEntryOk() As Boolean
   Dim lnCtr As Integer
   Dim lsUserID As String, lsUserName As String, lsOldProc As String
   Dim lnUserRights As Integer, lnRep As String

   If txtField(3).Text = "" Then
      MsgBox "Barrcode not found!!!", vbCritical, "Warning"
      txtField(3).SetFocus
      GoTo endProc
   End If
   
   If txtField(5).Text = "" Then
      MsgBox "Phone Number not found!!!", vbCritical, "Warning"
      txtField(5).SetFocus
      GoTo endProc
   End If
   
   If CDbl(txtField(7).Text) < CDbl(txtField(6).Text) Then
      MsgBox "Invalid amount paid!!!", vbCritical, "Warning"
      txtField(7).SetFocus
      GoTo endProc
   End If
   
   isEntryOk = True

endProc:
   Exit Function
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
