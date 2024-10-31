VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0"; "xrGridControl.ocx"
Begin VB.Form frmCP_AccessJobOrderReg 
   BorderStyle     =   0  'None
   Caption         =   "Warranty to Service Center Maintenance"
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10995
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7200
   ScaleWidth      =   10995
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   9645
      TabIndex        =   20
      Top             =   2505
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
      Picture         =   "frmCP_AccessJobOrderReg.frx":0000
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   3810
      Index           =   2
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   3285
      Width           =   9315
      _ExtentX        =   16431
      _ExtentY        =   6720
      BackColor       =   12632256
      BorderStyle     =   1
      Begin xrGridEditor.GridEditor GridEditor1 
         Height          =   3660
         Left            =   45
         TabIndex        =   15
         Tag             =   "et0;eb0;et0;bc2"
         Top             =   60
         Width           =   9180
         _ExtentX        =   16193
         _ExtentY        =   6456
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
         Object.HEIGHT          =   3660
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
         MOUSEICON       =   "frmCP_AccessJobOrderReg.frx":077A
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
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   2205
      Index           =   0
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   1065
      Width           =   9315
      _ExtentX        =   16431
      _ExtentY        =   3889
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   645
         Index           =   4
         Left            =   1005
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   1350
         Width           =   4455
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   3
         Left            =   1005
         TabIndex        =   10
         Top             =   1020
         Width           =   4455
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   1005
         TabIndex        =   6
         Top             =   690
         Width           =   1830
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   3630
         MaxLength       =   10
         TabIndex        =   8
         Top             =   690
         Width           =   1830
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
         TabIndex        =   4
         Top             =   150
         Width           =   1620
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   1305
         Index           =   5
         Left            =   5490
         MaxLength       =   512
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   690
         Width           =   3675
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
         Left            =   6615
         TabIndex        =   24
         Tag             =   "eb0;et0"
         Top             =   225
         Width           =   2505
      End
      Begin VB.Shape Shape4 
         Height          =   330
         Index           =   0
         Left            =   6585
         Top             =   180
         Width           =   2565
      End
      Begin VB.Shape Shape3 
         Height          =   390
         Index           =   0
         Left            =   6555
         Top             =   150
         Width           =   2625
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   195
         Index           =   4
         Left            =   375
         TabIndex        =   11
         Top             =   1380
         Width           =   570
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   195
         Index           =   2
         Left            =   5460
         TabIndex        =   13
         Top             =   495
         Width           =   675
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "S. Center"
         Height          =   195
         Index           =   6
         Left            =   285
         TabIndex        =   9
         Top             =   1065
         Width           =   660
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   315
         Left            =   1080
         Tag             =   "et0;ht2"
         Top             =   240
         Width           =   1620
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "J.O. No."
         Height          =   195
         Index           =   18
         Left            =   2955
         TabIndex        =   7
         Top             =   720
         Width           =   585
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trans. Date"
         Height          =   195
         Index           =   1
         Left            =   105
         TabIndex        =   5
         Top             =   705
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
         TabIndex        =   3
         Top             =   210
         Width           =   735
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H8000000F&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   285
         Index           =   0
         Left            =   6615
         Tag             =   "et0;et0"
         Top             =   210
         Width           =   2520
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   525
      Index           =   1
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   525
      Width           =   9315
      _ExtentX        =   16431
      _ExtentY        =   926
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
         Index           =   6
         Left            =   975
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   90
         Width           =   1620
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
         Index           =   7
         Left            =   3630
         MaxLength       =   50
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   90
         Width           =   5565
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&J.O. No."
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
         Left            =   135
         TabIndex        =   25
         Top             =   135
         Width           =   720
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Company"
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
         Left            =   2790
         TabIndex        =   1
         Top             =   135
         Width           =   780
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   9645
      TabIndex        =   23
      Top             =   3765
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
      Picture         =   "frmCP_AccessJobOrderReg.frx":0796
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   9645
      TabIndex        =   19
      Top             =   1875
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
      Picture         =   "frmCP_AccessJobOrderReg.frx":0F10
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   9645
      TabIndex        =   21
      Top             =   3135
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&History"
      AccessKey       =   "H"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCP_AccessJobOrderReg.frx":168A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   9645
      TabIndex        =   16
      Top             =   1245
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
      Picture         =   "frmCP_AccessJobOrderReg.frx":1E04
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   9645
      TabIndex        =   17
      Top             =   1875
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
      Picture         =   "frmCP_AccessJobOrderReg.frx":257E
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   6
      Left            =   9645
      TabIndex        =   22
      Top             =   3765
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
      Picture         =   "frmCP_AccessJobOrderReg.frx":2CF8
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   7
      Left            =   9645
      TabIndex        =   18
      Top             =   2505
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
      Picture         =   "frmCP_AccessJobOrderReg.frx":3472
   End
End
Attribute VB_Name = "frmCP_AccessJobOrderReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmCP_JobOrder"

Private WithEvents oTrans As clsAccessJobOrder
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin

Dim pbGridFocus As Boolean
Dim pbLoadRecord As Boolean
Dim pnCtr As Integer, pnIndex As Integer

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   Dim lnRep As Integer

   lsOldProc = "cmdButton_Click"
   ''On Error GoTo errProc

   txtField_LostFocus pnIndex
   With GridEditor1
      Select Case Index
      Case 0 'Browse
         If oTrans.SearchTransaction() Then
            LoadMaster
            LoadDetail
         End If
         .Refresh
      Case 1 'Update
         If txtField(0).Text = "" Then
            MsgBox "No Transaction to Update!!!" & vbCrLf & _
                   "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
            GoTo endProc
         End If

         If oTrans.UpdateTransaction Then
            initButton xeModeUpdate
   
            txtField(1).SetFocus
         End If
      Case 2 'History
         If txtField(0).Text = "" Then
            MsgBox "No Record is Loaded!!!" & vbCrLf & _
                   "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
            GoTo endProc
         End If
         
'         With oFormLedger
'            .ClientID = oTrans.Master("sClientID")
'            .Show 1
'         End With
      Case 3 'Close
         Unload Me
      Case 4 'Save
         If isEntryOk Then
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
            End If

            If oTrans.SaveTransaction Then
               MsgBox "Record Updated Successfully!!!", vbInformation, "Notice"
               initButton xeModeReady
               txtField(6).SetFocus
            Else
               MsgBox "Unable to Update Record!!!", vbCritical, "Warning"
            End If
         End If
      Case 5 ' Search
         If Not pbGridFocus Then
            oTrans.SearchMaster pnIndex, txtField(pnIndex).Text
            txtField(pnIndex).SetFocus
         End If
         .Refresh
      Case 6 ' Cancel Update
         lnRep = MsgBox("Cancel Current Transaction!!!?", vbYesNo + vbQuestion, "Confirm")
         If lnRep = vbYes Then
            initButton xeModeReady
            If pbLoadRecord Then
               oTrans.OpenTransaction oTrans.Master("sTransNox")
               LoadMaster
               LoadDetail
            Else
               ClearFields
            End If
            txtField(6).SetFocus
         Else
            txtField(pnIndex).SetFocus
         End If
         .Refresh
      Case 7 ' Delete Row
         If .Rows = 2 Then
            If oTrans.deleteDetail(.Row - 1) Then
               .TextMatrix(1, 1) = ""
               .TextMatrix(1, 2) = ""
            End If
         Else
            If oTrans.deleteDetail(.Row - 1) Then .deleteRow
         End If
      End Select
   End With

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0

   GridEditor1.Refresh
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

Private Sub Form_Load()
   Dim lsOldProc As String

   lsOldProc = "Form_Load"
   ''On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oTrans = New clsAccessJobOrder
   Set oTrans.AppDriver = oApp

   oTrans.InitTransaction
   oTrans.NewTransaction

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransMaintenance
   
   InitGrid
   ClearFields
   initButton xeModeReady

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oTrans = Nothing
   Set oSkin = Nothing
End Sub

Private Sub initButton(lnStat As Integer)
   Dim lbShow As Boolean

   lbShow = IIf(lnStat = 0, False, True)
   cmdButton(0).Visible = Not lbShow
   cmdButton(1).Visible = Not lbShow
   cmdButton(3).Visible = Not lbShow
   
   cmdButton(4).Visible = lbShow
   cmdButton(5).Visible = lbShow
   cmdButton(6).Visible = lbShow
   cmdButton(7).Visible = lbShow

   txtField(1).Enabled = lbShow
   txtField(2).Enabled = lbShow
   txtField(3).Enabled = lbShow
   txtField(5).Enabled = lbShow

    With GridEditor1
      .ColEnabled(1) = lbShow
      .ColEnabled(2) = lbShow
   End With

   xrFrame1(0).Enabled = lbShow
   xrFrame1(2).Enabled = lbShow
   xrFrame1(1).Enabled = Not lbShow
End Sub

Private Sub GridEditor1_AddingRow(Cancel As Boolean)
   With GridEditor1
      If .TextMatrix(.Row, 1) = "" Then
         Cancel = True
      End If
      If Not Cancel Then oTrans.addDetail
   End With
End Sub

Private Sub GridEditor1_EditorValidate(Cancel As Boolean)
   With GridEditor1
      If .Col = 6 Then
         If .TextMatrix(.Row, .Col) = 0 Then .TextMatrix(.Row, .Col) = .TextMatrix(.Row, .Cols - 1)
         If .TextMatrix(.Row, .Col) > .TextMatrix(.Row, .Cols - 1) Then .TextMatrix(.Row, .Col) = 0
         oTrans.Detail(.Row - 1, .Col) = CDbl(.TextMatrix(.Row, .Col))
      End If
   End With
End Sub

Private Sub GridEditor1_LostFocus()
   With GridEditor1
      .EditorBackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub GridEditor1_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String
   
   lsOldProc = "GridEditor1_KeyDown"
   ''On Error GoTo errProc
   
   If KeyCode = vbKeyF3 Then
      With GridEditor1
         If oTrans.searchDetail(.Row - 1, .Col, .TextMatrix(.Row, .Col)) Then
            .Col = 2
         Else
            oTrans.Detail(.Row - 1, .Col) = .TextMatrix(.Row, .Col)
            .Col = 1
         End If
         
         .Refresh
         .SetFocus
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

Private Sub GridEditor1_GotFocus()
   With GridEditor1
      .EditorBackColor = oApp.getColor("HT1")
   End With
   pbGridFocus = True
End Sub

Private Sub oTrans_DetailRetrieved(ByVal Index As Integer)
   With GridEditor1
      .TextMatrix(.Row, Index) = oTrans.Detail(.Row - 1, Index)
   End With
End Sub

Private Sub oTrans_MasterRetrieved(ByVal Index As Integer)
   With oTrans
      txtField(Index).Text = IFNull(.Master(Index), "")
   End With
End Sub

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
      .ColWidth(3) = 2780
      
      'empty row
      .TextMatrix(1, 1) = ""
      .TextMatrix(1, 2) = ""
      .TextMatrix(1, 3) = ""
      .TextMatrix(1, 4) = 0
      .TextMatrix(1, 5) = 0
   End With
   
   Label2.Caption = "UNKNOWN"
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
      .TextMatrix(0, 4) = "QOH"
      .TextMatrix(0, 5) = "Qty"
      .TextMatrix(0, 6) = "Rcv"
      
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
      .ColWidth(4) = 500
      .ColWidth(5) = 500
      .ColWidth(6) = 500
      
      .ColFormat(4) = "#,##0"
      .ColFormat(5) = "#,##0"
      .ColFormat(6) = "#,##0"
      .ColNumberOnly(6) = True
      .ColDefault(4) = 0
      .ColDefault(5) = 0
      .ColDefault(6) = 0
      
      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      .ColAlignment(4) = 6
      .ColAlignment(5) = 6
      .ColAlignment(6) = 6
      
      .ColEnabled(1) = False
      .ColEnabled(2) = False
      .ColEnabled(3) = False
      .ColEnabled(4) = False
      .ColEnabled(5) = False
      
      .EditorBackColor = oApp.getColor("HT1")
      
      .Row = 1
      .Col = 6
   End With
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   Dim lsPlateNo() As String

   With txtField(Index)
      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = oApp.getColor("HT1")
   End With
   pbGridFocus = False
   pnIndex = Index
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String
   Dim lsValue As String

   lsOldProc = "txtField_KeyDown"
   ''On Error GoTo errProc

   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
      With txtField(Index)
         Select Case Index
         Case 3
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

Private Function isEntryOk() As Boolean
   If Trim(txtField(2).Text) = "" Then
      MsgBox "J.O. No. not found!!!", vbCritical, "Warning"
      txtField(2).SetFocus
      GoTo EntryNotOK
   End If

   If Trim(txtField(3).Text) = "" Then
      MsgBox "Company not found!!!", vbCritical, "Warning"
      txtField(3).SetFocus
      GoTo EntryNotOK
   End If

EntryOK:
   isEntryOk = True
   Exit Function
EntryNotOK:
   isEntryOk = False
End Function

Private Sub LoadMaster()
   For pnCtr = 0 To 5
      Select Case pnCtr
      Case 0
         txtField(pnCtr).Text = Format(oTrans.Master(pnCtr), "@@@@-@@@@@@")
      Case 2
         txtField(pnCtr).Text = oTrans.Master(pnCtr)
         txtField(6).Text = txtField(pnCtr).Text
         txtField(6).Tag = txtField(6).Text
      Case 3
         txtField(pnCtr).Text = oTrans.Master(pnCtr)
         txtField(7).Text = txtField(pnCtr).Text
         txtField(7).Tag = txtField(7).Text
      Case 1
         txtField(pnCtr).Text = Format(oTrans.Master(pnCtr), "MMMM DD, YYYY")
      Case Else
         txtField(pnCtr).Text = IIf(IsNull(oTrans.Master(pnCtr)), "", oTrans.Master(pnCtr))
      End Select
   Next
   
   Select Case oTrans.Master("cTranStat")
   Case 0
      Label2.Caption = "JOB ORDER"
   Case 1
      Label2.Caption = "FOR REPAIR"
   Case 2
      Label2.Caption = "RELEASED"
   Case 3
      Label2.Caption = "CANCELLED"
   Case 4
      Label2.Caption = "FORWARDED"
   Case 5
      Label2.Caption = "REPAIRED"
   End Select

   pbLoadRecord = True
End Sub

Private Sub LoadDetail()
   Dim lnCtr As Integer
   
   With GridEditor1
      .Rows = IIf(oTrans.ItemCount = 0, 2, oTrans.ItemCount + 1)
      
      .ColWidth(3) = 2780
      If .Rows > 16 Then .ColWidth(3) = 2580
      
      For pnCtr = 0 To oTrans.ItemCount - 1
         For lnCtr = 1 To .Cols - 1
            Select Case lnCtr
            Case 6
               .TextMatrix(pnCtr + 1, lnCtr) = oTrans.Detail(pnCtr, "nQuantity")
               oTrans.Detail(pnCtr, "nReceived") = .TextMatrix(pnCtr + 1, lnCtr)
            Case Else
               .TextMatrix(pnCtr + 1, lnCtr) = oTrans.Detail(pnCtr, lnCtr)
            End Select
         Next
      Next
   End With
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   With txtField(Index)
      .Text = TitleCase(.Text)
      Select Case Index
      Case 1
         If Not IsDate(.Text) Then .Text = Date
         .Text = Format(.Text, "MMMM DD, YYYY")
      Case 2, 5
         .Text = UCase(.Text)
      Case 6, 7
         If .Text = "" Then
            ClearFields
            Exit Sub
         End If
                           
         If Trim(LCase(.Text)) <> Trim(LCase(.Tag)) Then
            If oTrans.SearchTransaction(.Text, IIf(Index = 6, True, False)) Then
               LoadMaster
               LoadDetail
            Else
               ClearFields
               .SetFocus
            End If
         End If
      End Select
      If Index < 6 Then oTrans.Master(Index) = .Text
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
