VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmSplitLoadPosting 
   BorderStyle     =   0  'None
   Caption         =   "Load Transfer"
   ClientHeight    =   8070
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8070
   ScaleWidth      =   7920
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1140
      Index           =   0
      Left            =   120
      Tag             =   "wt0;fb0"
      Top             =   570
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   2011
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   2
         Left            =   1110
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   630
         Width           =   4875
      End
      Begin VB.TextBox txtSearch 
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
         Height          =   330
         Index           =   1
         Left            =   3735
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   225
         Width           =   2235
      End
      Begin VB.TextBox txtSearch 
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
         Height          =   330
         Index           =   0
         Left            =   1110
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   225
         Width           =   2115
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Destination"
         Height          =   210
         Index           =   2
         Left            =   210
         TabIndex        =   5
         Top             =   675
         Width           =   885
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   3300
         TabIndex        =   4
         Top             =   285
         Width           =   885
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Trans. No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   180
         TabIndex        =   3
         Top             =   270
         Width           =   885
      End
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   2475
      Left            =   120
      Tag             =   "wt0;fb0"
      Top             =   5445
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   4366
      BackColor       =   12632256
      Enabled         =   0   'False
      ClipControls    =   0   'False
      Begin VB.TextBox txtFied 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   4
         Left            =   1065
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   450
         Width           =   1980
      End
      Begin VB.TextBox txtFied 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   5
         Left            =   1065
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   840
         Width           =   1980
      End
      Begin VB.TextBox txtFied 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   8
         Left            =   1065
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   1245
         Width           =   1290
      End
      Begin VB.TextBox txtFied 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Index           =   10
         Left            =   4065
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   1245
         Width           =   1290
      End
      Begin VB.TextBox txtFied 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   6
         Left            =   4065
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   435
         Width           =   1980
      End
      Begin VB.TextBox txtFied 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   7
         Left            =   4065
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   840
         Width           =   1980
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Orig Code"
         Height          =   210
         Index           =   5
         Left            =   120
         TabIndex        =   17
         Top             =   450
         Width           =   840
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Load Amt"
         Height          =   210
         Index           =   7
         Left            =   120
         TabIndex        =   16
         Top             =   1290
         Width           =   675
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Trnsfer Amt"
         Height          =   210
         Index           =   8
         Left            =   3165
         TabIndex        =   15
         Top             =   1290
         Width           =   840
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Supersede"
         Height          =   210
         Index           =   9
         Left            =   3165
         TabIndex        =   14
         Top             =   465
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   210
         Index           =   6
         Left            =   120
         TabIndex        =   13
         Top             =   870
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   210
         Index           =   10
         Left            =   3165
         TabIndex        =   12
         Top             =   840
         Width           =   975
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   3690
      Index           =   1
      Left            =   120
      Tag             =   "wt0;fb0"
      Top             =   1740
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   6509
      BackColor       =   12632256
      Enabled         =   0   'False
      ClipControls    =   0   'False
      Begin VB.TextBox txtFied 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Index           =   0
         Left            =   1095
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   225
         Width           =   1740
      End
      Begin VB.TextBox txtFied 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   1
         Left            =   1200
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   1095
         Width           =   1740
      End
      Begin VB.TextBox txtFied 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   2
         Left            =   1200
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   1845
         Width           =   4455
      End
      Begin VB.TextBox txtFied 
         Appearance      =   0  'Flat
         Height          =   450
         Index           =   3
         Left            =   1200
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   2235
         Width           =   4455
      End
      Begin VB.TextBox txtFied 
         Appearance      =   0  'Flat
         Height          =   765
         Index           =   9
         Left            =   1200
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   2745
         Width           =   4455
      End
      Begin VB.TextBox txtFied 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   14
         Left            =   1200
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   1470
         Width           =   1740
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3810
         TabIndex        =   33
         Tag             =   "eb0;et0"
         Top             =   240
         Width           =   2070
      End
      Begin VB.Shape Shape3 
         Height          =   360
         Index           =   0
         Left            =   3765
         Top             =   210
         Width           =   2160
      End
      Begin VB.Shape Shape4 
         Height          =   420
         Index           =   0
         Left            =   3735
         Top             =   180
         Width           =   2220
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Trans. No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   14
         Left            =   180
         TabIndex        =   29
         Top             =   270
         Width           =   885
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   210
         Index           =   13
         Left            =   285
         TabIndex        =   28
         Top             =   1155
         Width           =   885
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Destination"
         Height          =   210
         Index           =   12
         Left            =   300
         TabIndex        =   27
         Top             =   1890
         Width           =   885
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   210
         Index           =   3
         Left            =   315
         TabIndex        =   26
         Top             =   2295
         Width           =   885
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   210
         Index           =   4
         Left            =   315
         TabIndex        =   25
         Top             =   2805
         Width           =   885
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Refer No"
         Height          =   210
         Index           =   11
         Left            =   285
         TabIndex        =   24
         Top             =   1530
         Width           =   885
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   6585
      TabIndex        =   30
      Top             =   1875
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
      Picture         =   "frmSplitLoadPosting.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   6585
      TabIndex        =   31
      Top             =   615
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
      Picture         =   "frmSplitLoadPosting.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   6585
      TabIndex        =   32
      Top             =   1245
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Post"
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
      Picture         =   "frmSplitLoadPosting.frx":0EF4
   End
End
Attribute VB_Name = "frmSplitLoadPosting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pxeMODULENAME = "frmSplitLoadPosting"

Private WithEvents oTrans As clsCPWallet2Retail
Attribute oTrans.VB_VarHelpID = -1

Private oSkin As clsFormSkin
Private bLoaded As Boolean

Dim pbFormLoad As Boolean

Private Sub Form_Activate()
   Dim lsOldProc As String

   lsOldProc = "Form_Activate"
   ''On Error GoTo errProc

   oApp.MenuName = Me.Tag
   Me.ZOrder 0

   If Not pbFormLoad = False Then
      pbFormLoad = True
   End If

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String

   lsOldProc = "Form_Load"
   ''On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oTrans = New clsCPWallet2Retail
   Set oTrans.AppDriver = oApp

   oTrans.InitTransaction
   oTrans.NewTransaction

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualRight

   ClearOthers
   InitFields


endProc:
   Exit Sub
errProc:
   ShowError lsOldProc
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oTrans = Nothing
   Set oSkin = Nothing

   pbFormLoad = False
End Sub

Private Sub InitForm(ByVal fnEdit As Integer)
   Dim loTxt As TextBox

   cmdButton(4).Visible = Not (fnEdit = 0)
   cmdButton(3).Visible = Not (fnEdit = 0)

   cmdButton(0).Visible = (fnEdit = 0)
   cmdButton(1).Visible = (fnEdit = 0)
   cmdButton(2).Visible = (fnEdit = 0)

   For Each loTxt In txtFied
      loTxt = ""
   Next

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

Private Sub InitFields()
   With oTrans
'      txtFied(0) = Format(.Master("sTransNox"), "@@@@@@-@@@@@@")
'      txtFied(1) = Format(.Master("dTransact"), "MMMM DD, YYYY")
      txtFied(8) = Format(.Master("nAmountxx"), "#,##0.00")
      txtFied(10) = Format(.Master("nAmountxx"), "#,##0.00")
   End With

End Sub

Private Sub initButton(lnStat As Integer)
   Dim lbShow As Boolean

   lbShow = IIf(lnStat = 0, False, True)
   cmdButton(1).Visible = lbShow
   cmdButton(3).Visible = Not lbShow

   cmdButton(0).Visible = Not lbShow
   cmdButton(2).Visible = lbShow

End Sub

Private Sub ClearOthers()
   Dim loTxt As TextBox
   Dim lnCtr As Integer

   For lnCtr = 0 To 10
      txtFied(lnCtr) = ""
   Next

   txtFied(14).Text = ""

   For lnCtr = 0 To 2
      txtSearch(lnCtr) = ""
   Next

   For Each loTxt In txtFied
      loTxt.BackColor = oApp.getColor("EB")
   Next
End Sub

Private Sub cmdButton_Click(Index As Integer)
   Dim lnRep As Integer

   Select Case Index
   Case 0 'Browse
      oTrans.SearchAcceptance
      Call LoadData
   Case 1 'Post
      If oTrans.Master("cTranStat") = xeStateOpen Then
         If oTrans.AcceptDelivery(txtSearch(1).Text) Then
            MsgBox "Transaction Post Succesfully!!!", vbInformation
         End If
      Else
         MsgBox "Unable to post Transaction." & vbCrLf & _
                        "Please Inform SSG/SEG.", vbInformation, "INFO"
      End If
   Case 2 'Close
      Unload Me
   End Select
End Sub

Private Sub oTrans_MasterRetrieved(ByVal Index As Variant, ByVal Value As Variant)
   Select Case Index
   Case 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 14
      txtFied(Index) = Value
   End Select
End Sub

Private Sub txtFied_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
      Select Case Index
      Case 2, 4, 6
         oTrans.SearchMaster Index, txtFied(Index).Text
         SetNextFocus
      Case 8
         txtFied(8).Text = Format(txtFied(8).Text, "#,##0.00")
         txtFied(10).Text = Format(txtFied(8).Text, "#,##0.00")
      Case Else
         SetNextFocus
      End Select
   End If
End Sub

Private Sub txtFied_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
   Case 8
      If Not IsNumeric(txtFied(Index).Text) Then txtFied(Index).Text = 0#
      oTrans.Master(Index) = Trim(txtFied(Index).Text)
   Case 10
      txtFied(Index) = txtFied(8)
   Case 14
      txtFied(Index) = txtFied(14)
      oTrans.Master(Index) = Trim(txtFied(Index).Text)
   Case Else
      oTrans.Master(Index) = txtFied(Index)
   End Select
End Sub

Private Sub txtSearch_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
      Select Case Index
      Case 0
         oTrans.SearchAcceptance txtSearch(Index), True
         Call LoadData
      Case 2
         oTrans.SearchAcceptance txtSearch(Index), False
         Call LoadData
      End Select

   End If
End Sub

Private Sub LoadData()

   txtFied(0) = oTrans.Master(0)
   txtFied(1) = oTrans.Master(1)
   txtFied(2) = oTrans.Master(2)
   txtFied(3) = IFNull(oTrans.Master(3), "")
   txtFied(4) = oTrans.Master(4)
   txtFied(5) = oTrans.Master(5)
   txtFied(6) = IFNull(oTrans.Master(6), "")
   txtFied(7) = IFNull(oTrans.Master(7), "")
   txtFied(8) = Format(oTrans.Master(8), "#,##0.00")
   txtFied(9) = oTrans.Master(9)
   txtFied(10) = Format(oTrans.Master(8), "#,##0.00")
   txtFied(14) = IFNull(oTrans.Master(14), "")

   txtSearch(0) = oTrans.Master(0)
   txtSearch(2) = oTrans.Master(2)
   Label2.Caption = TransStat(oTrans.Master("cTranStat"))

End Sub

Private Sub txtSearch_Validate(Index As Integer, Cancel As Boolean)
   If Not IsDate(txtSearch(1).Text) Then
      txtSearch(1).Text = Format(oApp.ServerDate, "MMM DD, YYYY")
      oTrans.Master("dReceived") = txtSearch(1).Text
   End If
End Sub
