VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmCP_Purchasing_Posting 
   BorderStyle     =   0  'None
   Caption         =   "Cellphone Purchase Order"
   ClientHeight    =   7125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11925
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7125
   ScaleWidth      =   11925
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   10575
      TabIndex        =   19
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
      Picture         =   "frmCP_Purchasing_Posting.frx":0000
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   2190
      Index           =   0
      Left            =   120
      Tag             =   "wt0;fb0"
      Top             =   1110
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   3863
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   1335
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   75
         Width           =   2265
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   8220
         TabIndex        =   15
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   570
         Width           =   1815
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   1320
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   570
         Width           =   5295
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   1320
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   915
         Width           =   5295
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   465
         Index           =   5
         Left            =   1320
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   1605
         Width           =   5295
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   6
         Left            =   8235
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   915
         Width           =   1800
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   4
         Left            =   1320
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   1260
         Width           =   5295
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   315
         Left            =   1350
         Tag             =   "et0;ht2"
         Top             =   165
         Width           =   2265
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction No"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   4
         Top             =   165
         Width           =   1095
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Date"
         Height          =   195
         Index           =   1
         Left            =   6960
         TabIndex        =   14
         Top             =   675
         Width           =   1230
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company Name"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   6
         Top             =   675
         Width           =   1125
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   8
         Top             =   1005
         Width           =   570
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   195
         Index           =   5
         Left            =   90
         TabIndex        =   12
         Top             =   1665
         Width           =   630
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Delivery"
         Height          =   195
         Index           =   6
         Left            =   6960
         TabIndex        =   16
         Top             =   1005
         Width           =   960
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Delivered To"
         Height          =   195
         Index           =   7
         Left            =   90
         TabIndex        =   10
         Top             =   1335
         Width           =   915
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
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
         Height          =   255
         Left            =   7050
         TabIndex        =   22
         Tag             =   "eb0;et0"
         Top             =   165
         Width           =   2940
      End
      Begin VB.Shape Shape3 
         Height          =   390
         Index           =   0
         Left            =   6975
         Top             =   105
         Width           =   3075
      End
      Begin VB.Shape Shape4 
         Height          =   330
         Index           =   0
         Left            =   7005
         Top             =   135
         Width           =   3015
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   10575
      TabIndex        =   21
      Top             =   1800
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
      Picture         =   "frmCP_Purchasing_Posting.frx":077A
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   540
      Index           =   1
      Left            =   120
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   10200
      _ExtentX        =   17992
      _ExtentY        =   953
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
         Left            =   1320
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   90
         Width           =   2145
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
         Left            =   5010
         MaxLength       =   50
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   90
         Width           =   5025
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
         Index           =   9
         Left            =   75
         TabIndex        =   0
         Top             =   120
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Company Name"
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
         Left            =   3615
         TabIndex        =   2
         Top             =   120
         Width           =   1410
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   10575
      TabIndex        =   20
      Top             =   1170
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
      Picture         =   "frmCP_Purchasing_Posting.frx":0EF4
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3630
      Left            =   135
      TabIndex        =   18
      Top             =   3345
      Width           =   10170
      _ExtentX        =   17939
      _ExtentY        =   6403
      _Version        =   393216
      FocusRect       =   0
      SelectionMode   =   1
   End
End
Attribute VB_Name = "frmCP_Purchasing_Posting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmCP_Purchasing"

Private WithEvents oTrans As clsCPPurchasing
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin

Dim pnCtr As Integer
Dim pnIndex As Integer
Dim pbGridFocus As Boolean
Dim pbSave As Boolean
Dim pbEditMode As Boolean

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

Private Sub InitGrid()
   With MSFlexGrid1
      .Rows = 2
      .Cols = 6
      .Font = "MS Sans Serif"

      'column title
      .TextMatrix(0, 1) = "BarrCode"
      .TextMatrix(0, 2) = "Description"
      .TextMatrix(0, 3) = "Brand"
      .TextMatrix(0, 4) = "Model"
      .TextMatrix(0, 5) = "Quantity"
      
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
      .ColWidth(2) = 3550
      .ColWidth(3) = 1650
      .ColWidth(4) = 1600
      .ColWidth(5) = 930
            
      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      .ColAlignment(4) = 6
      .ColAlignment(5) = 6

      .Row = 1
      .Col = 1
      .ColSel = .ColSel - 1
   End With
End Sub

Private Sub LoadMaster()
   For pnCtr = 0 To txtField.Count - 1
      Select Case pnCtr
      Case 0
         txtField(pnCtr).Text = Format(oTrans.Master(pnCtr), "@@@@-@@@@@@")
      Case 1
         txtField(pnCtr).Text = Format(oTrans.Master(pnCtr), "MMMM DD, YYYY")
      Case 2
         txtField(pnCtr).Text = oTrans.Master(2)
         txtField(pnCtr).Tag = txtField(pnCtr).Text
      Case 4
         txtField(pnCtr).Text = oTrans.Master(8)
      Case 7
         txtField(pnCtr).Text = Format(oTrans.Master(0), "@@@@-@@@@@@")
      Case 8
         txtField(pnCtr).Text = oTrans.Master(2)
      Case Else
         txtField(pnCtr).Text = IIf(IsNull(oTrans.Master(pnCtr)), "", oTrans.Master(pnCtr))
      End Select
   Next

   Label2.Caption = Format(TransStat(oTrans.Master("cTranStat")), ">")
   pbSave = True
End Sub

Private Sub LoadDetail()
   Dim lnCtr As Integer

   With MSFlexGrid1
      .Rows = IIf(oTrans.ItemCount = 0, 2, oTrans.ItemCount + 1)

      For pnCtr = 0 To oTrans.ItemCount - 1
         For lnCtr = 1 To .Cols - 1
            If lnCtr = 5 Then
               .TextMatrix(pnCtr + 1, lnCtr) = oTrans.Detail(pnCtr, 6)
            Else
               .TextMatrix(pnCtr + 1, lnCtr) = oTrans.Detail(pnCtr, lnCtr)
            End If
         Next
      Next
   End With
End Sub

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   Dim lsRep As String

   lsOldProc = "cmdButton_Click"
   On Error GoTo errProc
      
   Select Case Index
   Case 0 ' Browse
      If oTrans.SearchTransaction() Then
         LoadMaster
         LoadDetail
      End If
   Case 1 'post
      If txtField(0).Text <> "" Then
         If oTrans.Master("cTransTat") = 1 Then
            lsRep = MsgBox("Do you want Post this Transaction?", vbYesNo + vbQuestion, "Confirm")
               If lsRep = vbYes Then
                  If oTrans.PostTransaction(oTrans.Master("sTransNox")) = True Then
                     MsgBox "Post Transaction Successfully!!!", vbInformation
                  Else
                     MsgBox "Unable to Post Transaction!!!", vbCritical, "Warning"
                  End If
               Else
                     MsgBox "Unable to Post Transaction!!!", vbCritical, "Warning"
               End If
         End If
      End If
   Case 2 'Close
      Unload Me
   End Select

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Sub Form_Activate()
   MSFlexGrid1.Refresh

   oApp.MenuName = Me.Tag
   Me.ZOrder 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyReturn, vbKeyUp, vbKeyDown
      Select Case KeyCode
      Case vbKeyReturn, vbKeyDown
         If GetFocus = MSFlexGrid1.hWnd Then Exit Sub
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

   Set oTrans = New clsCPPurchasing
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

   pbEditMode = True

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oTrans = Nothing
   Set oSkin = Nothing
End Sub

Private Sub Clearfields()
   For pnCtr = 0 To txtField.Count - 1
      Select Case pnCtr
      Case 1, 6
         txtField(pnCtr).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
      Case Else
         txtField(pnCtr).Text = ""
         txtField(pnCtr).Tag = ""
      End Select
   Next

   With MSFlexGrid1
      .Rows = 2

      'empty row
      .TextMatrix(1, 1) = ""
      .TextMatrix(1, 2) = ""
      .TextMatrix(1, 3) = ""
      
      .Row = 1
      .Col = 1
      .ColSel = .Cols - 1
   End With

   pbSave = False
End Sub

Private Sub oTrans_DetailRetrieved(ByVal Index As Integer, ByVal Value As Variant)
      With MSFlexGrid1
      If Index = 5 Then
         .TextMatrix(.Row, Index) = oTrans.Detail(.Row, 6)
      Else
         .TextMatrix(.Row, Index) = Value
      End If
   End With
End Sub

Private Sub oTrans_MasterRetrieved(ByVal Index As Variant, ByVal Value As Variant)
   txtField(Index).Text = Value
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      Select Case Index
      Case 1, 7
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

   lsOldProc = "txtField_KeyDown"
   On Error GoTo errProc

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

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   With txtField(Index)
      .Text = TitleCase(.Text)
      Select Case Index
      Case 1, 7
         If Not IsDate(.Text) Then .Text = oApp.ServerDate
         .Text = Format(.Text, "MMMM DD, YYYY")
      Case 5
         .Text = Format(.Text, ">")
      Case 10
         If .Text = "" Then
            Clearfields
            Exit Sub
         End If

         If Trim(LCase(.Text)) <> Trim(LCase(.Tag)) Then
            If oTrans.SearchTransaction(.Text, IIf(Index = 9, True, False)) Then
               LoadMaster
               LoadDetail
            Else
               Clearfields
               .SetFocus
            End If
         End If
      End Select

      If Index < 9 Then oTrans.Master(Index) = .Text
   End With
End Sub
'

