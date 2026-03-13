VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmCP_SCJobOrder 
   BorderStyle     =   0  'None
   Caption         =   "Job Order"
   ClientHeight    =   5055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10995
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5055
   ScaleWidth      =   10995
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   4380
      Left            =   1575
      Tag             =   "wt0;fb0"
      Top             =   525
      Width           =   9315
      _ExtentX        =   16431
      _ExtentY        =   7726
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   9
         Left            =   5865
         TabIndex        =   18
         Top             =   1590
         Width           =   1440
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   13
         Left            =   7425
         TabIndex        =   28
         Top             =   2835
         Width           =   1665
      End
      Begin VB.OptionButton chkServiceType 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Back Job (J.O. #)"
         Height          =   195
         Index           =   2
         Left            =   7395
         TabIndex        =   27
         Tag             =   "wt0;fb0"
         Top             =   2625
         Width           =   1545
      End
      Begin VB.OptionButton chkServiceType 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Under Limited Warranty"
         Height          =   195
         Index           =   1
         Left            =   5355
         TabIndex        =   26
         Tag             =   "wt0;fb0"
         Top             =   2895
         Width           =   2100
      End
      Begin VB.OptionButton chkServiceType 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Void Warranty"
         Height          =   195
         Index           =   0
         Left            =   5355
         TabIndex        =   25
         Tag             =   "wt0;fb0"
         Top             =   2625
         Width           =   1440
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   840
         Index           =   15
         Left            =   1005
         MaxLength       =   512
         MultiLine       =   -1  'True
         TabIndex        =   30
         Top             =   3225
         Width           =   8160
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   11
         Left            =   7920
         TabIndex        =   22
         Top             =   1920
         Width           =   1170
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   10
         Left            =   5865
         TabIndex        =   20
         Top             =   1920
         Width           =   1440
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   12
         Left            =   5865
         MaxLength       =   25
         TabIndex        =   24
         Top             =   2250
         Width           =   3225
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   3
         Left            =   1005
         TabIndex        =   7
         Top             =   1185
         Width           =   4200
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   840
         Index           =   4
         Left            =   1005
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   1515
         Width           =   4200
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   7
         Left            =   5865
         MaxLength       =   25
         TabIndex        =   14
         Top             =   930
         Width           =   3225
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   8
         Left            =   5865
         TabIndex        =   16
         Top             =   1260
         Width           =   1440
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   1005
         TabIndex        =   3
         Top             =   855
         Width           =   1830
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   3630
         MaxLength       =   10
         TabIndex        =   5
         Top             =   855
         Width           =   1575
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
         Left            =   990
         TabIndex        =   1
         Top             =   240
         Width           =   1620
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   840
         Index           =   14
         Left            =   1005
         MaxLength       =   512
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   2370
         Width           =   4200
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
         Height          =   195
         Index           =   15
         Left            =   5355
         TabIndex        =   17
         Top             =   1635
         Width           =   435
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "CELLPHONE INFORMATION"
         Height          =   195
         Index           =   8
         Left            =   5340
         TabIndex        =   12
         Tag             =   "wt0;fb0"
         Top             =   705
         Width           =   2130
      End
      Begin VB.Shape Shape2 
         Height          =   2355
         Left            =   5235
         Top             =   855
         Width           =   3930
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   195
         Index           =   3
         Left            =   315
         TabIndex        =   29
         Top             =   3195
         Width           =   630
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ref No."
         Height          =   195
         Index           =   14
         Left            =   7335
         TabIndex        =   21
         Top             =   2010
         Width           =   555
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DOP"
         Height          =   195
         Index           =   12
         Left            =   5355
         TabIndex        =   19
         Top             =   1965
         Width           =   345
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dealer"
         Height          =   195
         Index           =   13
         Left            =   5355
         TabIndex        =   23
         Top             =   2295
         Width           =   465
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
         Height          =   195
         Index           =   11
         Left            =   285
         TabIndex        =   6
         Top             =   1215
         Width           =   660
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cust. Add."
         Height          =   195
         Index           =   10
         Left            =   210
         TabIndex        =   8
         Top             =   1545
         Width           =   735
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IMEI"
         Height          =   195
         Index           =   7
         Left            =   5355
         TabIndex        =   13
         Top             =   975
         Width           =   330
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Accessory"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   10
         Top             =   2415
         Width           =   735
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   315
         Left            =   1080
         Tag             =   "et0;ht2"
         Top             =   330
         Width           =   1620
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "J.O. No."
         Height          =   195
         Index           =   18
         Left            =   2955
         TabIndex        =   4
         Top             =   915
         Width           =   585
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Brand"
         Height          =   195
         Index           =   5
         Left            =   5355
         TabIndex        =   15
         Top             =   1320
         Width           =   420
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trans. Date"
         Height          =   195
         Index           =   1
         Left            =   105
         TabIndex        =   2
         Top             =   900
         Width           =   840
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trans. #"
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
         Top             =   300
         Width           =   735
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   90
      TabIndex        =   39
      Top             =   4305
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
      Picture         =   "frmCP_SCJobOrder.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   90
      TabIndex        =   33
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
      Picture         =   "frmCP_SCJobOrder.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   8
      Left            =   90
      TabIndex        =   38
      Top             =   4305
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
      Picture         =   "frmCP_SCJobOrder.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   90
      TabIndex        =   37
      Top             =   3675
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
      Picture         =   "frmCP_SCJobOrder.frx":166E
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   90
      TabIndex        =   35
      Top             =   2415
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "Back&Out"
      AccessKey       =   "O"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCP_SCJobOrder.frx":1DE8
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   7
      Left            =   90
      TabIndex        =   34
      Top             =   1785
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
      Picture         =   "frmCP_SCJobOrder.frx":2562
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   6
      Left            =   90
      TabIndex        =   36
      Top             =   3045
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
      Picture         =   "frmCP_SCJobOrder.frx":2CDC
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   90
      TabIndex        =   31
      Top             =   2415
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
      Picture         =   "frmCP_SCJobOrder.frx":3456
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   90
      TabIndex        =   32
      Top             =   3045
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
      Picture         =   "frmCP_SCJobOrder.frx":3BD0
   End
End
Attribute VB_Name = "frmCP_SCJobOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmCP_JobOrder"

Private WithEvents oTrans As clsJobOrder
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin

Dim pbLoadRecord As Boolean
Dim pnCtr As Integer, pnIndex As Integer

Private Sub chkServiceType_Click(Index As Integer)
   oTrans.Master("cJOTypexx") = Index
End Sub

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   Dim lnRep As Integer

   lsOldProc = "cmdButton_Click"
   On Error GoTo errProc

   txtField_LostFocus pnIndex
   Select Case Index
   Case 0 'Save
      If isEntryOK Then
         If oTrans.SaveTransaction Then
            MsgBox "Record Updated Successfully!!!", vbInformation, "Notice"
            InitButton xeModeReady
         Else
            MsgBox "Unable to Update Record!!!", vbCritical, "Warning"
         End If
      End If
   Case 1 ' Search
      oTrans.SearchMaster pnIndex, txtField(pnIndex).Text
      txtField(pnIndex).SetFocus
   Case 2 ' Browse
      If oTrans.SearchTransaction() Then
         LoadMaster

         If cmdButton(4).Visible = False Then
            InitButton xeModeReady
            cmdButton(4).SetFocus
         End If
      End If
   Case 3 ' Cancel Update
      lnRep = MsgBox("Cancel Current Transaction!!!?", vbYesNo + vbQuestion, "Confirm")
      If lnRep = vbYes Then
         InitButton xeModeReady
         If pbLoadRecord Then
            oTrans.OpenTransaction oTrans.Master("sTransNox")
            LoadMaster
         Else
            ClearFields
         End If
         cmdButton(4).SetFocus
      Else
         txtField(pnIndex).SetFocus
      End If
   Case 4 ' New
      oTrans.NewTransaction
      InitButton xeModeAddNew
      ClearFields
      
      txtField(1).SetFocus
   Case 5 ' BackOut - Cancel Transaction
      If pbLoadRecord Then
         If oTrans.CancelTransaction Then
            oTrans.NewTransaction
            InitButton xeModeAddNew
            ClearFields
            txtField(1).SetFocus
         Else
            MsgBox "Unable to BackOut Transaction!!!", vbCritical, "Warning"
         End If
      Else
         MsgBox "Unable to BackOut Transaction!!!" & vbCrLf & _
                "No Transaction is Loaded!!!", vbCritical, "Warning"
      End If
   Case 6 ' Print
   Case 7 ' Update
      If oTrans.UpdateTransaction Then
         InitButton xeModeAddNew
         txtField(1).SetFocus
      Else
         MsgBox "Unable to Update Transaction!!!", vbCritical, "Warning"
      End If
   Case 8 ' Closed
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
   End Select
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String

   lsOldProc = "Form_Load"
   On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oTrans = New clsJobOrder
   Set oTrans.AppDriver = oApp

   oTrans.JOStatus = xeJOStateOpen
   oTrans.InitTransaction
   oTrans.NewTransaction

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualLeft

   ClearFields
   InitButton xeModeAddNew

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oTrans = Nothing
   Set oSkin = Nothing
End Sub

Private Sub InitButton(lnStat As Integer)
   Dim lbShow As Boolean

   lbShow = IIf(lnStat = 0, False, True)
   cmdButton(4).Visible = Not lbShow
   cmdButton(5).Visible = Not lbShow
   cmdButton(6).Visible = Not lbShow
   cmdButton(7).Visible = Not lbShow
   cmdButton(8).Visible = Not lbShow

   cmdButton(0).Visible = lbShow
   cmdButton(1).Visible = lbShow
   cmdButton(3).Visible = lbShow

   txtField(1).Enabled = lbShow
   txtField(2).Enabled = lbShow
   txtField(3).Enabled = lbShow
   txtField(7).Enabled = lbShow
   txtField(8).Enabled = lbShow
   txtField(9).Enabled = lbShow
   txtField(10).Enabled = lbShow
   txtField(11).Enabled = lbShow
   txtField(12).Enabled = lbShow
   txtField(13).Enabled = lbShow
   txtField(14).Enabled = lbShow
   txtField(15).Enabled = lbShow

   If Not lbShow Then cmdButton(4).SetFocus
End Sub

Private Sub oTrans_MasterRetrieved(ByVal Index As Integer)
   With oTrans
      txtField(Index).Text = IFNull(.Master(Index), "")
   End With
End Sub

Private Sub ClearFields()
   For pnCtr = 0 To txtField.Count - 1
      Select Case pnCtr
      Case 0
         txtField(pnCtr).Text = Format(oTrans.Master(pnCtr), "@@@@-@@@@@@")
      Case 1, 10
         txtField(pnCtr).Text = Format(oApp.ServerDate, "MMM-DD-YYYY")
      Case 2
         txtField(pnCtr).Text = oTrans.Master(pnCtr)
      Case 5, 6
      Case Else
         txtField(pnCtr).Text = ""
      End Select
   Next
   
   chkServiceType(oTrans.Master("cJOTypexx")).Value = True
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
   Dim lsValue As String

   lsOldProc = "txtField_KeyDown"
   On Error GoTo errProc

   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
      With txtField(Index)
         Select Case Index
         Case 3, 7
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
   ShowError lsOldProc & "( " & Index _
                       & ", " & KeyCode _
                       & ", " & Shift & " )", True
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Function isEntryOK() As Boolean
   If Trim(txtField(7).Text) = "" Then
      MsgBox "Serial not found!!!", vbCritical, "Warning"
      txtField(7).SetFocus
      GoTo EntryNotOK
   End If

EntryOK:
   isEntryOK = True
   Exit Function
EntryNotOK:
   isEntryOK = False
End Function

Private Sub LoadMaster()
   For pnCtr = 0 To txtField.Count - 1
      Select Case pnCtr
      Case 0
         txtField(pnCtr).Text = Format(oTrans.Master(pnCtr), "@@@@-@@@@@@")
      Case 1, 10
         txtField(pnCtr).Text = Format(oTrans.Master(pnCtr), "MMM-DD-YYYY")
      Case 5, 6
      Case Else
         txtField(pnCtr).Text = IIf(IsNull(oTrans.Master(pnCtr)), "", oTrans.Master(pnCtr))
      End Select
   Next

   pbLoadRecord = True
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   With txtField(Index)
      .Text = TitleCase(.Text)
      Select Case Index
      Case 1, 10
         If Not IsDate(.Text) Then .Text = Date
         .Text = Format(.Text, "MMM-DD-YYYY")
      Case 2, 7, 9, 11, 13
         .Text = UCase(.Text)
      End Select
      oTrans.Master(Index) = .Text
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
