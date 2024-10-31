VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0"; "xrGridControl.ocx"
Begin VB.Form frmCP_AccessJobOrder 
   BorderStyle     =   0  'None
   Caption         =   "Warranty to Service Center"
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11715
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7200
   ScaleWidth      =   11715
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame2 
      Height          =   4365
      Left            =   1575
      Tag             =   "wt0;fb0"
      Top             =   2730
      Width           =   10065
      _ExtentX        =   17754
      _ExtentY        =   7699
      BackColor       =   12632256
      BorderStyle     =   1
      Begin xrGridEditor.GridEditor GridEditor1 
         Height          =   4230
         Left            =   45
         TabIndex        =   12
         Tag             =   "et0;eb0;et0;bc2"
         Top             =   45
         Width           =   9945
         _ExtentX        =   17542
         _ExtentY        =   7461
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
         Object.HEIGHT          =   4230
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
         MOUSEICON       =   "frmCP_AccessJobOrder.frx":0000
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
      Height          =   2190
      Left            =   1575
      Tag             =   "wt0;fb0"
      Top             =   525
      Width           =   10065
      _ExtentX        =   17754
      _ExtentY        =   3863
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
         TabIndex        =   9
         Top             =   1350
         Width           =   4455
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   3
         Left            =   1005
         TabIndex        =   7
         Top             =   1020
         Width           =   4455
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   1005
         TabIndex        =   3
         Top             =   690
         Width           =   1830
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   3630
         MaxLength       =   10
         TabIndex        =   5
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
         TabIndex        =   1
         Top             =   150
         Width           =   1620
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   1305
         Index           =   5
         Left            =   5520
         MaxLength       =   512
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   690
         Width           =   4455
      End
      Begin VB.Label lblField 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   195
         Index           =   4
         Left            =   375
         TabIndex        =   8
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
         Left            =   5490
         TabIndex        =   10
         Top             =   480
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
         TabIndex        =   6
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
         TabIndex        =   4
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
         TabIndex        =   2
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
         TabIndex        =   0
         Top             =   210
         Width           =   735
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   90
      TabIndex        =   23
      Top             =   5010
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
      Picture         =   "frmCP_AccessJobOrder.frx":001C
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   6
      Left            =   90
      TabIndex        =   16
      Top             =   1230
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
      Picture         =   "frmCP_AccessJobOrder.frx":0796
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   90
      TabIndex        =   22
      Top             =   5010
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
      Picture         =   "frmCP_AccessJobOrder.frx":0F10
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   90
      TabIndex        =   21
      Top             =   4380
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
      Picture         =   "frmCP_AccessJobOrder.frx":168A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   9
      Left            =   90
      TabIndex        =   19
      Top             =   3120
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
      Picture         =   "frmCP_AccessJobOrder.frx":1E04
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   7
      Left            =   90
      TabIndex        =   17
      Top             =   1860
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
      Picture         =   "frmCP_AccessJobOrder.frx":257E
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   10
      Left            =   90
      TabIndex        =   20
      Top             =   3750
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
      Picture         =   "frmCP_AccessJobOrder.frx":2CF8
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   90
      TabIndex        =   15
      Top             =   3750
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
      Picture         =   "frmCP_AccessJobOrder.frx":3472
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   90
      TabIndex        =   13
      Top             =   2490
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
      Picture         =   "frmCP_AccessJobOrder.frx":3BEC
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   90
      TabIndex        =   14
      Top             =   3120
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
      Picture         =   "frmCP_AccessJobOrder.frx":4366
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   8
      Left            =   90
      TabIndex        =   18
      Top             =   2490
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Conform"
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
      Picture         =   "frmCP_AccessJobOrder.frx":4AE0
   End
End
Attribute VB_Name = "frmCP_AccessJobOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmCP_AccessJobOrder"

Private WithEvents oTrans As clsAccessJobOrder
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin

Dim pbGridFocus As Boolean
Dim pbLoadRecord As Boolean
Dim pnCtr As Integer, pnIndex As Integer

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   Dim lnRep As Integer
   Dim oFormDate As frmDateCriteria

   lsOldProc = "cmdButton_Click"
   ''On Error GoTo errProc

   txtField_LostFocus pnIndex
   Set oFormDate = New frmDateCriteria
   Set oFormDate.AppDriver = oApp
   With GridEditor1
      Select Case Index
      Case 0 'Save
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
               lnRep = MsgBox("Do you want to conform transaction", vbYesNo + vbQuestion, "Confirm")
               If lnRep = vbYes Then
                  oFormDate.Show 1
                  pbLoadRecord = True
                  If oFormDate.Cancelled = True Then GoTo endProc
                  If oTrans.Fowarded(oTrans.Master("sTransNox"), oFormDate.DateEntry) Then
                     oTrans.NewTransaction
                     initButton xeModeAddNew
                     ClearFields
                     txtField(1).SetFocus
                  Else
                     MsgBox "Unable to Conform Transaction!!!", vbCritical, "Warning"
                  End If
               End If
            Else
               MsgBox "Unable to Update Record!!!", vbCritical, "Warning"
            End If
         End If
      Case 1 ' Search
         If Not pbGridFocus Then
            oTrans.SearchMaster pnIndex, txtField(pnIndex).Text
            txtField(pnIndex).SetFocus
         End If
         .Refresh
      Case 2 ' Delete Row
         If .Rows = 2 Then
            If oTrans.deleteDetail(.Row - 1) Then
               .TextMatrix(1, 1) = ""
               .TextMatrix(1, 2) = ""
            End If
         Else
            If oTrans.deleteDetail(.Row - 1) Then .deleteRow
         End If
      Case 3 ' Browse
         If oTrans.SearchTransaction() Then
            LoadMaster
            LoadDetail

            If cmdButton(6).Visible = False Then
               initButton xeModeReady
               cmdButton(6).SetFocus
            End If
         End If
         .Refresh
      Case 4 ' Cancel Update
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
            cmdButton(6).SetFocus
         Else
            txtField(pnIndex).SetFocus
         End If
         .Refresh
      Case 5 ' Close
         Unload Me
      Case 6 ' New
         oTrans.NewTransaction
         initButton xeModeAddNew
         ClearFields
         txtField(1).SetFocus
      Case 7 ' Update
         If oTrans.UpdateTransaction Then
            initButton xeModeAddNew
            txtField(1).SetFocus
         Else
            MsgBox "Unable to Update Transaction!!!", vbCritical, "Warning"
         End If
      Case 8 ' Conform - Forward to Service Center
         If pbLoadRecord Then
            oFormDate.Show 1
            If oFormDate.Cancelled = True Then GoTo endProc
            If oTrans.Fowarded(oTrans.Master("sTransNox"), oFormDate.DateEntry) Then
               oTrans.NewTransaction
               initButton xeModeAddNew
               ClearFields
               txtField(1).SetFocus
            Else
               MsgBox "Unable to Conform Transaction!!!", vbCritical, "Warning"
            End If
         Else
            MsgBox "Unable to Conform Transaction!!!" & vbCrLf & _
                   "No Transaction is Loaded!!!", vbCritical, "Warning"
         End If
      Case 9 ' BackOut - Cancel Transaction
         If pbLoadRecord Then
            If oTrans.CancelTransaction Then
               oTrans.NewTransaction
               initButton xeModeAddNew
               ClearFields
               txtField(1).SetFocus
            Else
               MsgBox "Unable to BackOut Transaction!!!", vbCritical, "Warning"
            End If
         Else
            MsgBox "Unable to BackOut Transaction!!!" & vbCrLf & _
                   "No Transaction is Loaded!!!", vbCritical, "Warning"
         End If
      Case 10 ' Print
      
      End Select
   End With

endProc:
   Set oFormDate = Nothing
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

   oTrans.JOStatus = xeJOStateJobOrder
   oTrans.InitTransaction
   oTrans.NewTransaction

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransaction

   InitGrid
   ClearFields
   initButton xeModeAddNew

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
   cmdButton(5).Visible = Not lbShow
   cmdButton(6).Visible = Not lbShow
   cmdButton(7).Visible = Not lbShow
   cmdButton(8).Visible = Not lbShow
   cmdButton(9).Visible = Not lbShow
   cmdButton(10).Visible = Not lbShow

   cmdButton(0).Visible = lbShow
   cmdButton(1).Visible = lbShow
   cmdButton(2).Visible = lbShow
   cmdButton(4).Visible = lbShow

   txtField(1).Enabled = lbShow
   txtField(2).Enabled = lbShow
   txtField(3).Enabled = lbShow
   txtField(5).Enabled = lbShow

    With GridEditor1
      .ColEnabled(1) = lbShow
      .ColEnabled(2) = lbShow
   End With

   If Not lbShow Then cmdButton(6).SetFocus
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
      oTrans.Detail(.Row - 1, .Col) = .TextMatrix(.Row, .Col)
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
      Case 0
         txtField(pnCtr).Text = Format(oTrans.Master(pnCtr), "@@@@-@@@@@@")
      Case 1
         txtField(pnCtr).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
      Case Else
         txtField(pnCtr).Text = ""
      End Select
   Next

   With GridEditor1
      .Rows = 2
      .ColWidth(3) = 4050
      
      'empty row
      .TextMatrix(1, 1) = ""
      .TextMatrix(1, 2) = ""
      .TextMatrix(1, 3) = ""
      .TextMatrix(1, 4) = 0
      .TextMatrix(1, 5) = 0
   End With
End Sub

Private Sub InitGrid()
   With GridEditor1
      .Rows = 2
      .Cols = 6
      .Font = "MS Sans Serif"

      'column title
      .TextMatrix(0, 1) = "BarrCode"
      .TextMatrix(0, 2) = "Description"
      .TextMatrix(0, 3) = "Model"
      .TextMatrix(0, 4) = "QOH"
      .TextMatrix(0, 5) = "Qty"
      
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
      
      .ColFormat(4) = "#,##0"
      .ColFormat(5) = "#,##0"
      .ColNumberOnly(5) = True
      .ColDefault(4) = 0
      .ColDefault(5) = 0
      
      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      .ColAlignment(4) = 6
      .ColAlignment(5) = 6
      
      .ColEnabled(3) = False
      .ColEnabled(4) = False
      
      .EditorBackColor = oApp.getColor("HT1")
      
      .Row = 1
      .Col = 1
   End With
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   Dim lsPlateNo() As String

   With txtField(Index)
      If Index = 6 Then
         .MaxLength = 6
         If .Text <> "" And _
            Len(.Text) = 7 Then
            lsPlateNo = Split(.Text, "-")
            If UBound(lsPlateNo) > 0 Then .Text = lsPlateNo(0) & lsPlateNo(1)
         End If
      End If

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
         Case 3, 5
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
      MsgBox "J.O. No not found!!!", vbCritical, "Warning"
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
   For pnCtr = 0 To txtField.Count - 1
      Select Case pnCtr
      Case 0
         txtField(pnCtr).Text = Format(oTrans.Master(pnCtr), "@@@@-@@@@@@")
      Case 1
         txtField(pnCtr).Text = Format(oTrans.Master(pnCtr), "MMMM DD, YYYY")
      Case Else
         txtField(pnCtr).Text = IIf(IsNull(oTrans.Master(pnCtr)), "", oTrans.Master(pnCtr))
      End Select
   Next

   pbLoadRecord = True
End Sub

Private Sub LoadDetail()
   Dim lnCtr As Integer
   
   With GridEditor1
      .Rows = IIf(oTrans.ItemCount = 0, 2, oTrans.ItemCount + 1)
      
      .ColWidth(3) = 4050
      If .Rows > 16 Then .ColWidth(3) = 3850
      
      For pnCtr = 0 To oTrans.ItemCount - 1
         For lnCtr = 1 To .Cols - 1
            .TextMatrix(pnCtr + 1, lnCtr) = oTrans.Detail(pnCtr, lnCtr)
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
      Case 2
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
