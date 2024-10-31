VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmCPDelSched 
   BorderStyle     =   0  'None
   Caption         =   "Mobile Phone Delivery Schedule"
   ClientHeight    =   5595
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13920
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   13920
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2340
      Left            =   1575
      TabIndex        =   18
      Top             =   3150
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   4128
      _Version        =   393216
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   90
      TabIndex        =   21
      Top             =   1770
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
      Picture         =   "frmCPDelSched.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   90
      TabIndex        =   19
      Top             =   1155
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
      Picture         =   "frmCPDelSched.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   90
      TabIndex        =   20
      Top             =   540
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
      Picture         =   "frmCPDelSched.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   90
      TabIndex        =   22
      Top             =   2400
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
      Picture         =   "frmCPDelSched.frx":166E
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   90
      TabIndex        =   23
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
      Picture         =   "frmCPDelSched.frx":1DE8
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   90
      TabIndex        =   24
      Top             =   2400
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
      Picture         =   "frmCPDelSched.frx":2562
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   2595
      Index           =   0
      Left            =   1575
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   7020
      _ExtentX        =   12383
      _ExtentY        =   4577
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.ComboBox cmbOthers 
         Height          =   315
         ItemData        =   "frmCPDelSched.frx":2CDC
         Left            =   4230
         List            =   "frmCPDelSched.frx":2CEF
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1785
         Width           =   1845
      End
      Begin VB.CommandButton cmdOthers 
         Caption         =   "-"
         Height          =   330
         Index           =   1
         Left            =   6540
         TabIndex        =   17
         Top             =   2145
         Width           =   330
      End
      Begin VB.CommandButton cmdOthers 
         Caption         =   "+"
         Height          =   330
         Index           =   0
         Left            =   6150
         TabIndex        =   16
         Top             =   2145
         Width           =   360
      End
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   3
         Left            =   885
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   2145
         Width           =   5205
      End
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   885
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   1815
         Width           =   2265
      End
      Begin VB.ComboBox cmbField 
         Height          =   315
         ItemData        =   "frmCPDelSched.frx":2D23
         Left            =   915
         List            =   "frmCPDelSched.frx":2D2D
         TabIndex        =   3
         Text            =   "Division"
         Top             =   630
         Width           =   1875
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   3
         Left            =   915
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   1320
         Width           =   1875
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   645
         Index           =   4
         Left            =   3165
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   990
         Width           =   3690
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
         Index           =   2
         Left            =   915
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   990
         Width           =   1875
      End
      Begin VB.Shape Shape4 
         Height          =   330
         Index           =   0
         Left            =   4380
         Top             =   180
         Width           =   2445
      End
      Begin VB.Shape Shape3 
         Height          =   390
         Index           =   0
         Left            =   4350
         Top             =   150
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
         Left            =   4395
         TabIndex        =   26
         Tag             =   "eb0;et0"
         Top             =   225
         Width           =   2400
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Notes"
         Height          =   285
         Index           =   5
         Left            =   180
         TabIndex        =   14
         Top             =   2160
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Truck Size"
         Height          =   285
         Index           =   2
         Left            =   3420
         TabIndex        =   12
         Top             =   1830
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cluster"
         Height          =   285
         Index           =   0
         Left            =   180
         TabIndex        =   10
         Top             =   1830
         Width           =   1200
      End
      Begin VB.Line Line1 
         X1              =   90
         X2              =   6885
         Y1              =   1695
         Y2              =   1695
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
         Caption         =   "Division"
         Height          =   285
         Index           =   1
         Left            =   165
         TabIndex        =   2
         Top             =   690
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   285
         Index           =   3
         Left            =   180
         TabIndex        =   4
         Top             =   1020
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   285
         Index           =   7
         Left            =   3165
         TabIndex        =   8
         Top             =   765
         Width           =   1200
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
         Caption         =   "Schedule"
         Height          =   285
         Index           =   4
         Left            =   180
         TabIndex        =   6
         Top             =   1350
         Width           =   1200
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H8000000F&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   285
         Index           =   2
         Left            =   4410
         Tag             =   "et0;et0"
         Top             =   210
         Width           =   2400
      End
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   4950
      Left            =   8655
      TabIndex        =   25
      Top             =   540
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   8731
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   6
      Left            =   90
      TabIndex        =   27
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
      Picture         =   "frmCPDelSched.frx":2D4B
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   7
      Left            =   90
      TabIndex        =   28
      Top             =   1770
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
      Picture         =   "frmCPDelSched.frx":34C5
   End
End
Attribute VB_Name = "frmCPDelSched"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmCPDelSched"

Private WithEvents oTrans As ggcCPOrder.clsDeliverySchedule
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin

Dim pnOthNdx As Integer
Dim pbGridFocus As Boolean
Dim pnRow As Integer
Dim pnCtr As Integer
Dim pbSave As Boolean
Dim pbShown As Boolean

Private Sub cmbField_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
      SetNextFocus
   End If
End Sub

Private Sub cmbOthers_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
      SetNextFocus
   End If
End Sub

Private Sub cmbOthers_Validate(Cancel As Boolean)
   If cmbOthers.ListIndex >= 0 Then
      oTrans.Detail(pnRow, 2) = cmbOthers.ListIndex
      MSFlexGrid1.TextMatrix(pnRow + 1, 2) = oTrans.Detail(pnRow, 2)
   End If
End Sub

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   Dim lsRep As String

   lsOldProc = "cmdButton_Click"
   ''On Error GoTo errProc

   txtOthers_LostFocus pnOthNdx
   With MSFlexGrid1
      Select Case Index
      Case 0
         If .Rows > 2 Then
            pnCtr = 0
            Do While pnCtr < .Rows - 1
               If Trim(.TextMatrix(pnCtr + 1, 1)) = "" Then
                  pnRow = pnCtr
                  Call delDetail
               Else
                  pnCtr = pnCtr + 1
               End If
            Loop
         End If
         .ColWidth(3) = 3125
         If .Rows > 7 Then .ColWidth(3) = 2925
         
         oTrans.Master("cDivision") = cmbField.ListIndex

         If isEntryOk Then
            If oTrans.SaveTransaction Then
               MsgBox "Record Updated Successfully!!!", vbInformation, "Confirm"
               initButton xeModeReady
            Else
               MsgBox "Unable to Update Record!!!", vbCritical, "Warning"
            End If
         End If
      Case 1
         If pnOthNdx = 1 Then
            If oTrans.searchDetail(pnRow, .Col) Then .Col = .Col
            .Refresh
            .SetFocus
         End If
      Case 2
         If .Rows > 2 Then
            .ColWidth(3) = 3125
            If .Rows > 7 Then .ColWidth(3) = 2925
         End If
      Case 3 ' Cancel
         lsRep = MsgBox("Transaction is in Update Mode!!!" & vbCrLf & _
                        "Do you want to Cancel Transaction!!!", vbYesNo + vbQuestion, "Confirm")

         If lsRep = vbYes Then
            oTrans.InitTransaction

            ClearFields
            initButton xeModeReady
         Else
            txtOthers(pnOthNdx).SetFocus
         End If

         pbSave = False
      Case 4 ' New
         ClearFields
         oTrans.NewTransaction
         initButton xeModeAddNew
         
         InitFields
         txtField(2).SetFocus
      Case 5
         Unload Me
      Case 6
         If oTrans.SearchTransaction(oApp.BranchCode) Then
            Call InitFields
            Call LoadDetail
            Call loadCluster
         End If
      Case 7
         If txtField(0) = "" Then Exit Sub
         
         If oTrans.PostTransaction(txtField(0)) Then
            MsgBox "Transaction Posted Successfuly.", vbInformation, "Success"
         End If
      End Select
   End With

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Sub cmdOthers_Click(Index As Integer)
   If Index = 0 Then
      Call addDetail

      txtOthers(1).SetFocus
   Else
      txtOthers(1).SetFocus
      Call delDetail
   End If
End Sub

Private Sub Form_Activate()
   MSFlexGrid1.Refresh

   oApp.MenuName = Me.Tag
   Me.ZOrder 0
   pbShown = True
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String

   lsOldProc = "Form_Load"
   ''On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oTrans = New ggcCPOrder.clsDeliverySchedule
   Set oTrans.AppDriver = oApp

   oTrans.InitTransaction
   oTrans.NewTransaction

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualLeft

   InitForm
   ClearFields

   InitFields
   initButton xeModeAddNew
   txtField(0).Enabled = False

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oTrans = Nothing
   Set oSkin = Nothing
End Sub

Private Sub MSFlexGrid1_RowColChange()
   If Not pbShown Then Exit Sub

   With MSFlexGrid1
      If pnRow = .Row - 1 Then Exit Sub

      pnRow = .Row - 1
      If .TextMatrix(pnRow + 1, 1) = "" Then
         Call InitOthers
      Else
         Call setFieldInfo
      End If
   End With
End Sub

Private Sub oTrans_DetailRetrieved(ByVal Index As Integer, ByVal Value As Variant)
   Dim lsOldProc As String

   lsOldProc = "oTrans_DetailRetrieved"
   ''On Error GoTo errProc

   With MSFlexGrid1
      .TextMatrix(.Row, Index) = Value

      If Index = 1 Then
         Call loadCluster

         txtOthers(Index) = Value
      ElseIf Index = 2 Then
      Else
         txtOthers(Index) = Value
      End If
   End With

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", False
End Sub

Private Sub oTrans_MasterRetrieved(ByVal Index As Integer, ByVal Value As Variant)
   Dim lsOldProc As String

   lsOldProc = "oTrans_Master"
   ''On Error GoTo errProc

   Select Case Index
   Case 15
      Label2.Caption = TransStat(oTrans.Master("cTranStat"))
   Case Else
      txtField(Index).Text = Value
   End Select
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", False
End Sub

Private Sub InitForm()
   With MSFlexGrid1
      .Rows = 2
      .Cols = 4
      .Font = "MS Sans Serif"

      'column title
      .TextMatrix(0, 1) = "Cluster Name"
      .TextMatrix(0, 2) = "Truck Size"
      .TextMatrix(0, 3) = "Remarks"

      .Row = 0

      'column alignment
      For pnCtr = 0 To .Cols - 1
         .Col = pnCtr
         .CellFontBold = True
         .CellAlignment = 1
      Next

      'column width
      .ColWidth(0) = 300
      .ColWidth(1) = 2300
      .ColWidth(2) = 1200
      .ColWidth(3) = 2000

      .Row = 1
      .Col = 1
   End With
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      Select Case Index
      Case 1, 2
         .Text = Format(.Text, "MM/DD/YY")
      End Select

      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = oApp.getColor("HT1")
   End With
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

   xrFrame1(0).Enabled = lbShow

   If Not lbShow Then cmdButton(4).SetFocus
End Sub

Private Sub InitFields()
   Dim lsOldProc As String

   lsOldProc = "initFields"
   ''On Error GoTo errProc

   txtField(0).Text = Format(oTrans.Master("sTransNox"), "@@@@@@-@@@@@@")
   txtField(2).Text = Format(oTrans.Master("dTransact"), "mmmm dd, yyyy")
   txtField(3).Text = Format(oTrans.Master("dSchedule"), "mmmm dd, yyyy")
   txtField(4).Text = ""
   
   Label2.Caption = TransStat(oTrans.Master("cTranStat"))
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "()", False
End Sub

Private Sub ClearFields()
   Dim lsOldProc As String

   lsOldProc = "ClearFields"
   ''On Error GoTo errProc

   txtField(0).Text = ""
   txtField(2).Text = ""
   txtField(3).Text = ""
   txtField(4).Text = ""
   
   Label2.Caption = TransStat(-1)

   Call InitOthers
   TreeView1.Nodes.Clear

   With MSFlexGrid1
      .Rows = 2
      .Col = 1
      .ColWidth(3) = 3125

      .TextMatrix(1, 1) = ""
      .TextMatrix(1, 2) = ""
      .TextMatrix(1, 3) = ""
   End With
   
   cmbField.ListIndex = 0
   cmbField.Enabled = False

   pbSave = False
   pnRow = 0

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "()", False
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
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

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   Dim lsOldProc As String

   lsOldProc = "txtField_Validate"
   ''On Error GoTo errProc

   With txtField(Index)
      Select Case Index
      Case 2, 3
         If Not IsDate(.Text) Then .Text = oTrans.Master(Index)

         .Text = Format(.Text, "MMMM DD, YYYY")
      End Select

      oTrans.Master(Index) = .Text
   End With

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & ", " & Cancel & " )", True
End Sub

Private Function isEntryOk() As Boolean
   With MSFlexGrid1
      If Trim(.TextMatrix(1, 1)) = "" Then
         MsgBox "Detail is required!!!" & vbCrLf & _
                "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
         .Row = 1
         .Col = 1
         .SetFocus
         GoTo endProc
      End If
   End With

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

Private Sub txtOthers_GotFocus(Index As Integer)
   With txtOthers(Index)
      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = oApp.getColor("HT1")
   End With

   pnOthNdx = Index
End Sub

Private Sub txtOthers_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyF3, vbKeyReturn, vbKeyUp, vbKeyDown
      Select Case KeyCode
      Case vbKeyF3, vbKeyReturn
         If oTrans.searchDetail(pnRow, Index, txtOthers(Index).Text) Then
            Call loadCluster
         End If
         Call SetNextFocus
      Case vbKeyDown
         SetNextFocus
      Case vbKeyUp
         SetPreviousFocus
      End Select
   End Select
End Sub

Private Sub txtOthers_LostFocus(Index As Integer)
   With txtOthers(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub setDetail()
   With MSFlexGrid1
      .TextMatrix(pnRow + 1, 1) = oTrans.Detail(pnRow, 1)
      .TextMatrix(pnRow + 1, 2) = oTrans.Detail(pnRow, 2)
      .TextMatrix(pnRow + 1, 3) = oTrans.Detail(pnRow, 3)
   End With
End Sub

Private Sub loadCluster()
   Dim lors As Recordset
   Dim loNode As Node
   Dim lsClustrID As String

   With TreeView1
      .Nodes.Clear

      For pnCtr = 0 To oTrans.ItemCount - 1
         lsClustrID = oTrans.Detail(pnCtr, "sClustrID")
         Set lors = oTrans.getClusterMembers(lsClustrID)

         If TypeName(lors) <> "Nothing" Then
            Set loNode = .Nodes.Add(, , lsClustrID, oTrans.Detail(pnCtr, "sClustrDs"))

            Do Until lors.EOF()
               Call .Nodes.Add(lsClustrID, tvwChild, lors("sBranchCd"), lors("sBranchNm"))

               lors.MoveNext
            Loop
         End If
         loNode.Expanded = True
      Next
   End With
End Sub

Private Sub LoadDetail()
   Dim lnCtr As Integer
   
   InitForm
   
   With MSFlexGrid1
      .Col = 1
      .ColWidth(3) = 3125

      .TextMatrix(1, 1) = ""
      .TextMatrix(1, 2) = ""
      .TextMatrix(1, 3) = ""
   
      .Rows = oTrans.ItemCount + 1
      For lnCtr = 0 To oTrans.ItemCount - 1
         .TextMatrix(lnCtr + 1, 0) = lnCtr + 1
         .TextMatrix(lnCtr + 1, 1) = oTrans.Detail(lnCtr, "sClustrDs")
         .TextMatrix(lnCtr + 1, 2) = oTrans.Detail(lnCtr, "cTrckSize")
         .TextMatrix(lnCtr + 1, 3) = oTrans.Detail(lnCtr, "sRemarksx")
      Next
   End With
End Sub

Private Sub addDetail()
   With MSFlexGrid1
      If .TextMatrix(pnRow + 1, 1) = "" Then Exit Sub

      oTrans.addDetail
      .Rows = .Rows + 1

      .Row = .Rows - 1
      pnRow = .Row - 1
      Call InitOthers
   End With
End Sub

Private Sub delDetail()
   Dim lnCtr As Integer

   Call oTrans.deleteDetail(pnRow)

   With MSFlexGrid1
      For lnCtr = pnRow + 1 To .Rows - 2
         .TextMatrix(lnCtr, 1) = .TextMatrix(lnCtr + 1, 1)
         .TextMatrix(lnCtr, 2) = .TextMatrix(lnCtr + 1, 2)
         .TextMatrix(lnCtr, 3) = .TextMatrix(lnCtr + 1, 3)
      Next
      .Rows = .Rows - 1
      .Row = .Rows - 1

      pnRow = .Row - 1

      Call setFieldInfo
      Call loadCluster
   End With
End Sub

Private Sub InitOthers()
   pnOthNdx = 1
   txtOthers(1).Text = ""
   txtOthers(3).Text = ""
   cmbOthers.ListIndex = 0

   With MSFlexGrid1
      .TextMatrix(pnRow + 1, 1) = ""
      .TextMatrix(pnRow + 1, 2) = ""
      .TextMatrix(pnRow + 1, 3) = ""
   End With
End Sub

Private Sub setFieldInfo()
   With MSFlexGrid1
      txtOthers(1).Text = .TextMatrix(pnRow + 1, 1)
      txtOthers(3).Text = .TextMatrix(pnRow + 1, 3)

      cmbOthers.ListIndex = oTrans.TruckSizeCode(.TextMatrix(pnRow + 1, 2))
   End With
End Sub

Private Sub txtOthers_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
   Case 1
      If txtOthers(Index).Text <> "" Then
         oTrans.Detail(pnRow, Index) = txtOthers(Index).Text
         Call loadCluster
      End If
   Case 3
      oTrans.Detail(pnRow, Index) = txtOthers(Index).Text
      MSFlexGrid1.TextMatrix(pnRow + 1, Index) = txtOthers(Index).Text
   End Select
End Sub
