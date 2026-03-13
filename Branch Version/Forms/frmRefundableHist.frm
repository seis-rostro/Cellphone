VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmRefundableHist 
   BorderStyle     =   0  'None
   Caption         =   "Transaction History"
   ClientHeight    =   6555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6555
   ScaleWidth      =   9405
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   4395
      Index           =   0
      Left            =   120
      Tag             =   "wt0;fb0"
      Top             =   2010
      Width           =   7725
      _ExtentX        =   13626
      _ExtentY        =   7752
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   9
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   3150
         Width           =   1995
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   10
         Left            =   1800
         MaxLength       =   40
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   3550
         Width           =   5745
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   1800
         MaxLength       =   40
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   1980
         Width           =   5745
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   1800
         MaxLength       =   40
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   1590
         Width           =   5745
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   630
         Width           =   2655
      End
      Begin VB.ComboBox cmbField 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         ItemData        =   "frmRefundableHist.frx":0000
         Left            =   1800
         List            =   "frmRefundableHist.frx":0013
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2370
         Width           =   1995
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   1800
         MaxLength       =   40
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   8
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   2760
         Width           =   1995
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   1800
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   240
         Width           =   2655
      End
      Begin xrControl.xrFrame xrFrame2 
         Height          =   375
         Left            =   4920
         Top             =   240
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         BackColor       =   12632256
         ClipControls    =   0   'False
         Begin VB.Label Label 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "UNKNOWN"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   13
            Left            =   540
            TabIndex        =   0
            Top             =   30
            Width           =   1365
         End
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Amt Paid"
         Height          =   315
         Index           =   6
         Left            =   120
         TabIndex        =   18
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   195
         Index           =   14
         Left            =   120
         TabIndex        =   17
         Top             =   3600
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Particular"
         Height          =   195
         Index           =   13
         Left            =   120
         TabIndex        =   16
         Top             =   2060
         Width           =   660
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Reference No"
         Height          =   285
         Index           =   3
         Left            =   120
         TabIndex        =   15
         Top             =   1280
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transact. Date"
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Branch"
         Height          =   195
         Index           =   11
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Width           =   510
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Type"
         Height          =   285
         Index           =   5
         Left            =   120
         TabIndex        =   12
         Top             =   2440
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Check Amt"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   2880
         Width           =   1095
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
         Left            =   120
         TabIndex        =   10
         Top             =   300
         Width           =   1410
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   315
         Left            =   1920
         Tag             =   "et0;ht2"
         Top             =   315
         Width           =   2655
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1515
      Index           =   1
      Left            =   120
      Tag             =   "wt0;fb0"
      Top             =   480
      Width           =   7725
      _ExtentX        =   13626
      _ExtentY        =   2672
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   1800
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   210
         Width           =   5775
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   1800
         MaxLength       =   40
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   1015
         Width           =   5775
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   610
         Width           =   2655
      End
      Begin VB.ComboBox cmbField 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   2370
         Width           =   1995
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   99
         Left            =   1800
         MaxLength       =   40
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   1590
         Width           =   5745
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   12
         Left            =   1800
         MaxLength       =   40
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   1980
         Width           =   5745
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
         Index           =   2
         Left            =   120
         TabIndex        =   30
         Top             =   250
         Width           =   810
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Reference No"
         Height          =   285
         Index           =   10
         Left            =   120
         TabIndex        =   29
         Top             =   640
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Particular"
         Height          =   285
         Index           =   12
         Left            =   120
         TabIndex        =   28
         Top             =   1045
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Type"
         Height          =   285
         Index           =   7
         Left            =   120
         TabIndex        =   27
         Top             =   2440
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Branch"
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   26
         Top             =   1680
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Particular"
         Height          =   195
         Index           =   15
         Left            =   120
         TabIndex        =   25
         Top             =   2060
         Width           =   660
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   8070
      TabIndex        =   31
      Top             =   1755
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
      Picture         =   "frmRefundableHist.frx":0041
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   8070
      TabIndex        =   32
      Top             =   510
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
      Picture         =   "frmRefundableHist.frx":07BB
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   8070
      TabIndex        =   33
      Top             =   1140
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Void"
      AccessKey       =   "V"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmRefundableHist.frx":0F35
   End
End
Attribute VB_Name = "frmRefundableHist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
   Option Explicit
Private Const pxeMODULENAME = "frmRefundableHist"

Private WithEvents oTrans As clsRefundableDep
Attribute oTrans.VB_VarHelpID = -1

Private oSkin As clsFormSkin


Dim pbSave As Boolean
Dim psSelected() As String
Dim pnIndex As Integer
Dim pnCtr As Integer
Dim pnRow As Integer
Dim pbFormLoad As Boolean
Dim pbMasterGotFocus As Boolean
Dim pbGridGotFocus As Boolean
Dim pnActiveRow As Integer


Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0
End Sub
Private Sub initButton(lnStat As Integer)
   Dim lbShow As Boolean


   lbShow = IIf(lnStat = 0, False, True)
   cmdButton(4).Visible = Not lbShow
   'cmdButton(5).Visible = Not lbShow

   cmdButton(1).Visible = Not lbShow 'new
   cmdButton(2).Visible = lbShow 'save
   cmdButton(3).Visible = lbShow 'cancel
   cmdButton(4).Visible = Not lbShow 'close

'   xrFrame2.Enabled = lbShow
'   xrFrame3.Enabled = lbShow

   If Not lbShow Then cmdButton(1).SetFocus
End Sub
Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   Dim lnRep As Integer

   lsOldProc = "cmdButton_Click"
   ''On Error GoTo errProc
'   txtField_LostFocus pnIndex
      Select Case Index
        Case 0 'Browse
'
             If oTrans.SearchTransaction(txtField(0).Text, _
                                         txtField(1).Text, _
                                         txtField(2).Text, "", False) Then
                LoadDetail
             End If
'        Case 1 'confirm
'            If oTrans.CloseTransaction(oTrans.Master(0)) Then
'                MsgBox "Transaction has been confirmed successfully!", vbInformation, "Notice"
'                ClearAll
'            End If
            
        Case 2 'void
            If oTrans.CancelTransaction(oTrans.Master(0)) Then
                MsgBox "Transaction has been void successfully!", vbInformation, "Notice"
                ClearAll
            End If
        Case 3 'Close
            Unload Me
      End Select

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Sub LoadMaster()
   With oTrans
      txtField(0).Text = .Master(0)
      txtField(1).Text = strLongDate(.Master(1))
   End With
End Sub



Private Sub Form_Load()
   Dim lsOldProc As String

   lsOldProc = "Form_Load"
   'On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oTrans = New clsRefundableDep
   Set oTrans.AppDriver = oApp
   
   oTrans.Branch = oApp.BranchCode
   oTrans.InitTransaction
   oTrans.NewTransaction

   
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransMaintenance
   
    ClearAll
'   InitGrid
'   InitFields
'   LoadDetail

   'cmbTranType.ListIndex = 1
'   ClearFields
'   initButton xeModeAddNew

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub
Private Sub clearFields()
   Dim loTxt As TextBox
   Dim lnCtr As Integer
   
   For lnCtr = 1 To 3
      txtField(lnCtr) = ""
   Next
   
   For Each loTxt In txtField
      loTxt.BackColor = oApp.getColor("EB")
   Next

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oTrans = Nothing
   Set oSkin = Nothing
End Sub

Private Sub ClearAll()

    cmbField(1).ListIndex = 0
'    MSFlexGrid1.Clear
'    MSFlexGrid1.Rows = 1
    txtField(0) = ""
    txtField(1) = ""
    txtField(2) = ""
    txtField(3) = ""
    txtField(4) = ""
    txtField(5) = ""
    txtField(6) = ""
    txtField(7) = ""
    txtField(8) = "0.00"
    txtField(9) = "0.00"
    txtField(10) = ""
    Label(13) = "UNKNOWN"
End Sub

Private Sub LoadDetail()
    ' Search
    txtField(0) = IIf(IsNull(oTrans.Master(2)) Or Trim(oTrans.Master(2)) = "", "", oTrans.Master(2)) ' Branch
    txtField(1) = IIf(IsNull(oTrans.Master(17)) Or Trim(oTrans.Master(17)) = "", "", oTrans.Master(17)) ' Reference no
    txtField(2) = IIf(IsNull(oTrans.Master(3)) Or Trim(oTrans.Master(3)) = "", "", oTrans.Master(3)) ' Particular
    ' Details
    txtField(3) = IIf(IsNull(oTrans.Master(0)) Or Trim(oTrans.Master(0)) = "", "", Format(oTrans.Master(0), "@@@@@@-@@@@@@")) ' Transaction no
    txtField(4) = IIf(IsNull(oTrans.Master(1)) Or Trim(oTrans.Master(1)) = "", "", Format(oTrans.Master(1), "MMMM DD, YYYY")) ' Tran date
    txtField(5) = IIf(IsNull(oTrans.Master(17)) Or Trim(oTrans.Master(17)) = "", "", oTrans.Master(17)) ' Refer no
    txtField(6) = IIf(IsNull(oTrans.Master(2)) Or Trim(oTrans.Master(2)) = "", "", oTrans.Master(2)) ' Branch
    txtField(7) = IIf(IsNull(oTrans.Master(3)) Or Trim(oTrans.Master(3)) = "", "", oTrans.Master(3)) ' Particular
    txtField(8) = Format(IFNull(oTrans.Master(4), "0.00"), "#,##0.00")   'check amt
    txtField(9) = Format(IFNull(oTrans.Master(5), "0.00"), "#,##0.00")   ' amt paid
    txtField(10) = IIf(IsNull(oTrans.Master(7)) Or Trim(oTrans.Master(7)) = "", "", oTrans.Master(7)) ' Remarks
    
    Select Case oTrans.Master(12)
    Case 0
        Label(13) = "OPENED"
    Case 1
        Label(13) = "CONFIRMED"
    Case 2
        Label(13) = "APPROVED"
    Case 3
        Label(13) = "VOID"
    Case 4
        Label(13) = "COLLECTED"
    Case Else
        Label(13) = "UNKNOWN" ' Default to empty for null or unrecognized values
        
    End Select
    
    cmbField(1).ListIndex = oTrans.Master(6)
End Sub


Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String
   
   lsOldProc = "txtField_KeyDown"
   On Error GoTo errProc
   
   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
      With txtField(Index)
         If Index >= 0 And Index <= 2 Then
            ' Call the SearchTransaction function with the relevant text fields
            If oTrans.SearchTransaction(txtField(0).Text, _
                                         txtField(1).Text, _
                                         txtField(2).Text, 0, False) Then
                ' If search is successful, load the details
                LoadDetail
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
