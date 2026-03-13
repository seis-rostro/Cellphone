VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0"; "xrGridControl.ocx"
Begin VB.Form frmCPTransGAways 
   BorderStyle     =   0  'None
   Caption         =   "Giveaways Issuance"
   ClientHeight    =   7320
   ClientLeft      =   0
   ClientTop       =   2715
   ClientWidth     =   10080
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   10080
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   90
      TabIndex        =   16
      Top             =   1190
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
      Picture         =   "frmCPTransGAways.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   90
      TabIndex        =   15
      Top             =   555
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
      Picture         =   "frmCPTransGAways.frx":077A
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   525
      Index           =   1
      Left            =   1560
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   8400
      _ExtentX        =   14817
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
         Index           =   5
         Left            =   3615
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   90
         Width           =   4635
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
         Index           =   4
         Left            =   840
         TabIndex        =   1
         Top             =   90
         Width           =   1635
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Customer"
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
         Left            =   2745
         TabIndex        =   2
         Top             =   135
         Width           =   780
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&SI No"
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
         Index           =   20
         Left            =   165
         TabIndex        =   0
         Top             =   135
         Width           =   1365
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   7
      Left            =   90
      TabIndex        =   17
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
      Picture         =   "frmCPTransGAways.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   90
      TabIndex        =   18
      Top             =   1190
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
      Picture         =   "frmCPTransGAways.frx":166E
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   6
      Left            =   90
      TabIndex        =   20
      Top             =   2450
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
      Picture         =   "frmCPTransGAways.frx":1DE8
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   90
      TabIndex        =   19
      Top             =   555
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
      Picture         =   "frmCPTransGAways.frx":2562
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   570
      Index           =   2
      Left            =   1560
      Tag             =   "wt0;fb0"
      Top             =   6600
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   1005
      Enabled         =   0   'False
      BorderStyle     =   1
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "LEGEND: STATUS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   13
         Left            =   135
         TabIndex        =   10
         Top             =   75
         Width           =   1785
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "-Open"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   12
         Left            =   2490
         TabIndex        =   12
         Top             =   165
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "-Released"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   10
         Left            =   4065
         TabIndex        =   14
         Top             =   165
         Width           =   870
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   2287
         TabIndex        =   11
         Tag             =   "ht0;fb0"
         Top             =   75
         Width           =   240
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   3825
         TabIndex        =   13
         Tag             =   "ht0;fb0"
         Top             =   75
         Width           =   240
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   5430
      Index           =   0
      Left            =   1560
      Tag             =   "wt0;fb0"
      Top             =   1110
      Width           =   8400
      _ExtentX        =   14817
      _ExtentY        =   9578
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   6525
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   120
         Width           =   1665
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   1395
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   735
         Width           =   3960
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   465
         Index           =   3
         Left            =   1395
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   "frmCPTransGAways.frx":2CDC
         Top             =   1065
         Width           =   3960
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   1395
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   165
         Width           =   1950
      End
      Begin xrGridEditor.GridEditor GridEditor2 
         Height          =   1725
         Left            =   120
         TabIndex        =   24
         Tag             =   "et0;eb0;et0;bc2"
         Top             =   1635
         Width           =   8130
         _ExtentX        =   14340
         _ExtentY        =   3043
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
         FOCUSRECT       =   0
         EDITORBACKCOLOR =   -2147483643
         EDITORFORECOLOR =   -2147483640
         FORECOLOR       =   -2147483640
         FORECOLORFIXED  =   -2147483630
         FORECOLORSEL    =   -2147483634
         FORMATSTRING    =   ""
         Object.HEIGHT          =   1725
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
         MOUSEICON       =   "frmCPTransGAways.frx":2CE4
         MOUSEPOINTER    =   0
         REDRAW          =   -1  'True
         RIGHTTOLEFT     =   0   'False
         ROWS            =   2
         SCROLLBARS      =   3
         SCROLLTRACK     =   0   'False
         SELECTIONMODE   =   1
         Object.TOOLTIPTEXT     =   ""
         WORDWRAP        =   0   'False
      End
      Begin xrGridEditor.GridEditor GridEditor1 
         Height          =   1725
         Left            =   120
         TabIndex        =   25
         Tag             =   "et0;eb0;et0;bc2"
         Top             =   3600
         Width           =   8130
         _ExtentX        =   14340
         _ExtentY        =   3043
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
         Object.HEIGHT          =   1725
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
         MOUSEICON       =   "frmCPTransGAways.frx":2D00
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
      Begin VB.Line Line1 
         X1              =   0
         X2              =   8520
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Transact. Date"
         Height          =   285
         Index           =   1
         Left            =   5400
         TabIndex        =   22
         Top             =   120
         Width           =   1065
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction No."
         Height          =   285
         Index           =   14
         Left            =   135
         TabIndex        =   5
         Top             =   180
         Width           =   1350
      End
      Begin VB.Shape Shape2 
         Height          =   990
         Index           =   0
         Left            =   120
         Top             =   630
         Width           =   8130
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cust. Address"
         Height          =   195
         Index           =   11
         Left            =   195
         TabIndex        =   8
         Top             =   1140
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name"
         Height          =   285
         Index           =   3
         Left            =   195
         TabIndex        =   6
         Top             =   780
         Width           =   1200
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   1
         Left            =   1485
         Tag             =   "et0;ht2"
         Top             =   270
         Width           =   1950
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   90
      TabIndex        =   26
      Top             =   1815
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Confirm"
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
      Picture         =   "frmCPTransGAways.frx":2D1C
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   90
      TabIndex        =   23
      Top             =   1815
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
      Picture         =   "frmCPTransGAways.frx":3496
   End
End
Attribute VB_Name = "frmCPTransGAways"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Private Const pxeMODULENAME = "frmCPTransGAways"
'
'Private WithEvents oTrans As clsCPGiveAway
'Private oSkin As clsFormSkin
'
'Dim pnIndex As Integer, pnCtr As Integer
'Dim pbEditMode As Boolean
'Dim pbGridFocus As Boolean
'Dim pbHsSerial As Boolean
'
'Private Sub cmdButton_Click(Index As Integer)
'   Dim lsRep As String
'   Dim lsOldProc As String
'   Dim lsBarrCode As String
'   Dim lsQty As String
'   Dim lnCtr As Integer
'   Dim lnQty As Integer
'   Dim lbDuplicate As Boolean
'
'    lsOldProc = "cmdButton_Click"
'   'On Error GoTo errProc
'
'   Select Case Index
'    Case 0 'browse
'    If pnIndex >= 4 Then
'       If txtField(pnIndex).Text <> "" Then
'            If oTrans.SearchTransaction((txtField(pnIndex).Text), IIf(pnIndex = 4, True, False)) Then
'             LoadMaster
'             LoadDetail
'            End If
'      End If
'    Else
'        If txtField(4).Text <> "" Then
'            If oTrans.SearchTransaction((txtField(4).Text), True) Then
'             LoadMaster
'             LoadDetail
'            End If
'        End If
'    End If
'
'   Case 1 'update
'      If txtField(2).Text <> "" Then
'         If oTrans.UpdateTransaction Then
'
'            initButton xeModeUpdate
'            GridEditor1.SetFocus
'            GridEditor1.Col = 5
'            GridEditor1.Rows = oTrans.ItemCountOthers + 2
'         Else
'            MsgBox "Unable to update transaction.", vbCritical, "Notice'"
'         End If
'      Else
'         MsgBox "Unable to update transaction." & vbCrLf & _
'            "Please verify you entry.", vbCritical, "Notice"
'      End If
'
'   Case 2 'Confirm
'        If MsgBox("Do you want to confirm transaction?", _
'                  vbQuestion + vbYesNo, "Confirm") = vbYes Then
'
'            If oTrans.CloseTransaction(oTrans.Master("sTransNox")) Then
'                If oTrans.OpenTransaction(oTrans.Master("sTransNox")) Then
'                    LoadMaster
'                    LoadDetail
'                    initButton xeModeReady
'                    txtField(5).SetFocus
'                    MsgBox "Transaction Released Successfully!!!", vbInformation, "Notice"
'                End If
'            Else
'                MsgBox "Unable to Release GiveAway!!!"
'            End If
'
'         Else
'            MsgBox "Unable to Release GiveAway!!!"
'         End If
'    Case 3 'Search
'      If pbEditMode Then
'          With GridEditor1
'                lsBarrCode = oTrans.SearchReferNo(IIf(.Col > 2, 1, .Col), .TextMatrix(.Row, .Col))
'                lnQty = 0
'
'                If lsBarrCode <> "" Then
'                    For lnCtr = 1 To .Rows - 1
'                      If Trim(LCase(lsBarrCode)) = Trim(LCase(.TextMatrix(lnCtr, 1))) Then
'                         lbDuplicate = True
'                      End If
'                    Next
'
'                    If Not lbDuplicate Then
'                       If Trim(.Text) <> "" Then Call InsertDetail(lnQty, lsBarrCode)
'                    End If
'             End If
'          End With
'    End If
'   Case 4 'save
'      If txtField(0).Text <> "" Then
'         If oTrans.SaveTransaction Then
'            initButton xeModeReady
'            If oTrans.OpenTransaction(oTrans.Master("sTransNox")) Then
'               LoadMaster
'               LoadDetail
'               initButton xeModeReady
'               txtField(5).SetFocus
'            End If
'         Else
'           MsgBox "Unable to save transaction!!!" & vbCrLf & _
'                    "Please contact GMC/GGC SEG for assistance!!!", vbCritical, "Warning"
'         End If
'      End If
'
'    Case 5 'delete row
'        With GridEditor1
'            If oTrans.deleteDetail(.Row) Then .deleteRow
'            .ColWidth(2) = 3970
'            If .Rows > 6 Then .ColWidth(2) = 3720
'         End With
'    Case 6 'Cancel
'     lsRep = MsgBox("Cancel Current Transaction?", vbYesNo + vbQuestion, "Confirm")
'      If lsRep = vbYes Then
'         If oTrans.OpenTransaction(oTrans.Master("sTransNox")) Then
'            LoadMaster
'            LoadDetail
'            initButton xeModeReady
'            txtField(5).SetFocus
'         End If
'      End If
'
'   Case 7 'close
'      Unload Me
'   End Select
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " & Index & " )", True
'End Sub
'
'Private Sub Form_Activate()
'   oApp.MenuName = Me.Tag
'   Me.ZOrder 0
'
'   With GridEditor1
'      .Refresh
'   End With
'End Sub
'
'Private Sub Form_Load()
'   Dim lsOldProc As String
'
'   lsOldProc = "Form_Load"
'   'On Error GoTo errProc
'
'   CenterChildForm mdiMain, Me
'
'   Set oTrans = New clsCPGiveAway
'   Set oTrans.AppDriver = oApp
'   oTrans.InitTransaction
'
'   Set oSkin = New clsFormSkin
'   Set oSkin.AppDriver = oApp
'   Set oSkin.Form = Me
'   oSkin.ApplySkin xeFormTransaction
'
'   InitGrid
'   InitValue
'   initButton xeModeReady
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " & " )", True
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'   Set oTrans = Nothing
'   Set oSkin = Nothing
'End Sub
'
'Private Sub GridEditor1_GotFocus()
'   With GridEditor1
'      .EditorBackColor = oApp.getColor("HT1")
'   End With
'   pbGridFocus = True
'End Sub
'
'Private Sub GridEditor1_KeyDown(KeyCode As Integer, Shift As Integer)
'   Dim lsOldProc As String
'   Dim lsBarrCode As String
'   Dim lsQty As String
'   Dim lnCtr As Integer
'   Dim lnQty As Integer
'   Dim lbDuplicate As Boolean
'
'   lsOldProc = "GridEditor1_KeyDown"
'   'On Error GoTo errProc
'
'   Select Case KeyCode
'   Case vbKeyF3, vbKeyReturn
'      With GridEditor1
'         If .Col = 1 Or .Col = 2 Then
'
'            lsBarrCode = oTrans.SearchReferNo(.Col, .TextMatrix(.Row, .Col))
'            lnQty = 0
'
'            If lsBarrCode <> "" Then
'                For lnCtr = 1 To .Rows - 1
'                  If Trim(LCase(lsBarrCode)) = Trim(LCase(.TextMatrix(lnCtr, 1))) Then
'                     lbDuplicate = True
'                  End If
'                Next
'
'                If Not lbDuplicate Then
'                   If Trim(.Text) <> "" Then Call InsertDetail(lnQty, lsBarrCode)
'                End If
'
'            End If
'         End If
'      End With
'   Case vbKeyDown
'      If GridEditor1.Row = GridEditor1.Rows - 1 Then Exit Sub
'      GridEditor1.Row = GridEditor1.Row + 1
'      GridEditor1.Col = GridEditor1.Cols - 1
'      GridEditor1.SetFocus
'   Case vbKeyUp
'      If GridEditor1.Row = 1 Then Exit Sub
'      GridEditor1.Row = GridEditor1.Row - 1
'      GridEditor1.Col = GridEditor1.Cols - 1
'      GridEditor1.SetFocus
'   End Select
'
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " _
'                       & "  " & KeyCode _
'                       & ", " & Shift _
'                       & " )", True
'End Sub
'
'
'
'Private Sub GridEditor1_LostFocus()
'   With GridEditor1
'      .EditorBackColor = oApp.getColor("EB")
'   End With
'End Sub
'
'Private Sub GridEditor1_RowAdded()
'   With GridEditor1
'      .TextMatrix(.Rows - 1, 1) = oTrans.Detail(.Rows - 2, "xReferNox")
'      .TextMatrix(.Rows - 1, 2) = oTrans.Detail(.Rows - 2, "sDescript")
'      .TextMatrix(.Rows - 1, 3) = oTrans.Detail(.Rows - 2, "nQtyOnHnd")
'      .TextMatrix(.Rows - 1, 4) = oTrans.Detail(.Rows - 2, "nQuantity")
''      .TextMatrix(.Rows - 1, 5) = oTrans.Detail(.Rows - 2, "dModified")
'   End With
'End Sub
'
'Private Sub GridEditor1_RowColChange()
'   If Not pbGridFocus Then Exit Sub
'   With GridEditor1
'      .ColEnabled(1) = True
'      .ColEnabled(2) = True
'   End With
'End Sub
'
'Private Sub oTrans_DetailRetrieved(ByVal Index As Integer)
'    With GridEditor1
'        If oTrans.Detail(.Row - 1, "sCategID1") = "C001013" Or _
'                       oTrans.Detail(.Row - 1, "sCategID1") = "C001045" Or _
'                       oTrans.Detail(.Row - 1, "sCategID1") = "M029001" Or _
'                       oTrans.Detail(.Row - 1, "nQuantity") <= 0 Then
'
'            Select Case Index
'                Case 1, 2, 3
'                   .TextMatrix(.Row, Index) = oTrans.Detail(.Row - 1, Index)
'                Case 4
'                   .TextMatrix(.Row, Index) = oTrans.Detail(.Row - 1, Index)
'            End Select
'        End If
'
'        .ColWidth(2) = 3970
'        If .Rows > 6 Then .ColWidth(2) = 3720
'    End With
'End Sub
'
'Private Sub GridEditor1_AddingRow(Cancel As Boolean)
'   With GridEditor1
'      If .TextMatrix(.Row, 1) = Empty Then
'         ' empty record is not allowed
'         Cancel = True
'      Else
'         Cancel = Not oTrans.addDetail()
'      End If
'   End With
'End Sub
'
'Private Sub txtField_GotFocus(Index As Integer)
'   With txtField(Index)
'      .SelStart = 0
'      .SelLength = Len(.Text)
'   End With
'   pbGridFocus = False
'End Sub
'
'Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'   Dim lsOldProc As String
'
'   lsOldProc = "txtField_KeyDown"
'   'On Error GoTo errProc
'
'   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
'        Select Case Index
'            Case 4, 5
'                If oTrans.SearchTransaction(txtField(Index).Text, (Index = 4)) = True Then
'                    LoadMaster
'                    LoadDetail
'                End If
'                KeyCode = 0 ' Prevent default action
'        End Select
'    End If
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " _
'                       & "  " & Index _
'                       & ", " & KeyCode _
'                       & ", " & Shift _
'                       & " )", True
'End Sub
'
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'   Select Case KeyCode
'   Case vbKeyReturn, vbKeyDown
'      If GetFocus = GridEditor1.hwnd Then Exit Sub
'      SetNextFocus
'   Case vbKeyUp
'      SetPreviousFocus
'   End Select
'End Sub
'
'Private Sub initButton(lnStat As Integer)
'   Dim lbShow As Boolean
'
'   lbShow = IIf(lnStat = 0, False, True)
'   xrFrame1(1).Enabled = Not lbShow
'   cmdButton(0).Visible = Not lbShow
'   cmdButton(1).Visible = Not lbShow
'   cmdButton(2).Visible = Not lbShow
'   cmdButton(7).Visible = Not lbShow
'
'
'   cmdButton(3).Visible = lbShow
'   cmdButton(4).Visible = lbShow
'   cmdButton(5).Visible = lbShow
'   cmdButton(6).Visible = lbShow
'
'   xrFrame1(0).Enabled = lbShow
'
'   With GridEditor1
'      .ColEnabled(1) = lbShow
'      .ColEnabled(2) = lbShow
'   End With
'
'
'
'   pbEditMode = lbShow
'End Sub
'
'Private Sub InitValue()
'   For pnCtr = 0 To 5
'      Select Case pnCtr
'      Case 0 To 5
'         txtField(pnCtr).Text = ""
'         txtField(pnCtr).Tag = ""
'      Case 1
'         txtField(pnCtr).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
'      End Select
'   Next
'
'   With GridEditor1
'      .Rows = 2
'      .TextMatrix(1, 1) = ""
'      .TextMatrix(1, 2) = ""
'      .TextMatrix(1, 3) = 0
'      .TextMatrix(1, 4) = 0
'   End With
'   With GridEditor2
'      .Rows = 2
'      .TextMatrix(1, 1) = ""
'      .TextMatrix(1, 2) = ""
'      .TextMatrix(1, 3) = 0
'      .TextMatrix(1, 4) = 0
'   End With
'   pbGridFocus = False
'End Sub
'
'Private Sub InitGrid()
'   Dim lnCtr As Integer
'
'   With GridEditor1
'      .Rows = 2
'      .Cols = 5
'      .Font = "MS Sans Serif"
'
'      'Column Title
'      .TextMatrix(0, 1) = "IMEI No - Barcode"
'      .TextMatrix(0, 2) = "Description"
'      .TextMatrix(0, 3) = "QOH"
'      .TextMatrix(0, 4) = "Qty"
''      .TextMatrix(0, 5) = "Date Added"
'      .Row = 0
'
'      'Column Alignment
'      For lnCtr = 0 To .Cols - 1
'         .Col = lnCtr
'         .CellFontBold = True
'         .CellAlignment = 3
'      Next
'
'      .ColAlignment(1) = 1
'
'      'Column Width
'      .ColWidth(0) = 330
'      .ColWidth(1) = 2850
'      .ColWidth(2) = 3970
'      .ColWidth(3) = 500
'      .ColWidth(4) = 400
'      .ColWidth(5) = 2500
'
'      .ColDefault(4) = 0
'      .ColNumberOnly(4) = True
'      .ColMaxValue(4) = 999
'
'      .ColEnabled(3) = False
'      .ColEnabled(5) = False
'
'      .TextMatrix(1, 1) = ""
'      .TextMatrix(1, 2) = ""
'      .TextMatrix(1, 3) = 0
'      .TextMatrix(1, 4) = 0
'
'      .EditorBackColor = oApp.getColor("HT1")
'
'      .Row = 1
'      .Col = 1
'   End With
'
'
'   With GridEditor2
'      .Rows = 2
'      .Cols = 5
'      .Font = "MS Sans Serif"
'      'Column Title
'      .TextMatrix(0, 1) = "IMEI No - Barcode"
'      .TextMatrix(0, 2) = "Description"
'      .TextMatrix(0, 3) = "Price"
'      .TextMatrix(0, 4) = "Qty"
'      .Row = 0
'
'      'Column Alignment
'      For lnCtr = 0 To .Cols - 1
'         .Col = lnCtr
'         .CellFontBold = True
'         .CellAlignment = 3
'      Next
'
'      .ColAlignment(1) = 1
'      'Column Width
'      .ColWidth(0) = 330
'      .ColWidth(1) = 2850
'      .ColWidth(2) = 3380
'      .ColWidth(3) = 1000
'      .ColWidth(4) = 500
'
'      .ColDefault(3) = 0
'      .ColDefault(4) = 0
'      .ColNumberOnly(4) = True
'      .ColMaxValue(4) = 999
'
'      .ColEnabled(1) = False
'      .ColEnabled(2) = False
'      .ColEnabled(3) = False
'      .ColEnabled(4) = False
'
'      .TextMatrix(1, 1) = ""
'      .TextMatrix(1, 2) = ""
'      .TextMatrix(1, 3) = 0
'      .TextMatrix(1, 4) = 0
'
'      .EditorBackColor = oApp.getColor("HT1")
'
'      .Row = 1
'      .Col = 1
'   End With
'End Sub
'
'Private Sub LoadMaster()
'   For pnCtr = 3 To 4
'      txtField(pnCtr) = oTrans.Master(pnCtr)
'   Next
'   txtField(0).Text = Format(oTrans.Master(0), IIf(Len(oApp.BranchCode) = 2, "@@@@-@@@@@@", "@@@@@@-@@@@@@"))
'   txtField(1).Text = Format(oTrans.Master(1), "MMMM DD, YYYY")
'   txtField(2).Text = Format(oTrans.Master("xfullname"))
'   txtField(3).Text = Format(oTrans.Master("xaddressx"))
'
'   txtField(4).Text = oTrans.Master(2)
'   txtField(5).Text = txtField(2).Text
'
'   txtField(4).Tag = oTrans.Master(2)
'   txtField(5).Tag = txtField(2).Text
'End Sub
'
'Private Sub LoadDetail()
'   Dim lnRow As Integer
'   Dim lsRep As String
'   Dim lnDetail As String
'    lnDetail = 0
'      GridEditor2.Rows = oTrans.ItemCount + 1
'      For lnRow = 0 To oTrans.ItemCount - 1
'        If oTrans.Detail(lnRow, "sCategID1") <> "C001013" And _
'            oTrans.Detail(lnRow, "sCategID1") <> "C001045" And _
'            oTrans.Detail(lnRow, "sCategID1") <> "M029001" And _
'            oTrans.Detail(lnRow, "nQuantity") > 0 Then
'           lnDetail = lnDetail + 1
'            'add to main detail
'            With GridEditor2
'                .Rows = lnDetail + 1
'                 .TextMatrix(lnDetail, 1) = oTrans.Detail(lnRow, "xReferNox")
'                 .TextMatrix(lnDetail, 2) = oTrans.Detail(lnRow, "sDescript")
'                 .TextMatrix(lnDetail, 3) = Format(oTrans.Detail(lnRow, "nSelPrice"), "#,##0.00")
'                 .TextMatrix(lnDetail, 4) = oTrans.Detail(lnRow, "nQuantity")
'                 .ColWidth(2) = 3380
'                If .Rows > 6 Then .ColWidth(2) = 3130
'            End With
'        End If
'      Next
'
'      GridEditor1.Rows = oTrans.ItemCountOthers + 1
'      For lnRow = 0 To oTrans.ItemCountOthers - 1
'            With GridEditor1
'                .Rows = oTrans.ItemCountOthers + 1
'                .TextMatrix(lnRow + 1, 1) = oTrans.OtherDetail(lnRow, "xReferNox")
'                .TextMatrix(lnRow + 1, 2) = oTrans.OtherDetail(lnRow, "sDescript")
'                .TextMatrix(lnRow + 1, 3) = oTrans.OtherDetail(lnRow, "nQtyOnHnd")
'                .TextMatrix(lnRow + 1, 4) = oTrans.OtherDetail(lnRow, "nQuantity")
''                .TextMatrix(lnRow + 1, 5) = oTrans.OtherDetail(lnRow, "dModified")
'                .ColWidth(2) = 3970
'                If .Rows > 6 Then .ColWidth(2) = 3720
'            End With
'      Next
'
'
'
'      GridEditor1.Row = 1
'      GridEditor1.Col = 1
''      GridEditor1.SetFocus
'End Sub
'
'Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
'   Dim lsOldProc As String
'
'   lsOldProc = "txtField_Validate"
'   'On Error GoTo errProc
'
'   With txtField(Index)
'      If .Text = "" Then
'         InitValue
'         GoTo endProc
'      End If
'
'      Select Case Index
'      Case 4, 5
'         If LCase(.Tag) <> LCase(.Text) Then
'            If oTrans.SearchTransaction(.Text, IIf(Index = 4, True, False)) = True Then
'               LoadMaster
'               LoadDetail
'            Else
'               InitValue
'               InitGrid
'            End If
'         End If
'
'         .Tag = .Text
'      Case 1
'         If Not IsDate(.Text) Then .Text = oApp.ServerDate
'         .Text = Format(.Text, "MMMM DD, YYYY")
'         oTrans.Master("dTransact") = .Text
'      End Select
'   End With
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " _
'                       & "  " & Index _
'                       & ", " & Cancel _
'                       & " )", True
'End Sub
'
'Private Sub ShowError(ByVal lsProcName As String, Optional bEnd As Boolean = False)
'   With oApp
'      .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
'      If bEnd Then
'         .xShowError
'         End
'      Else
'         With Err
'            .Raise .Number, .Source, .Description
'         End With
'      End If
'   End With
'End Sub
'
'
'Private Sub InsertDetail(ByVal Quantity As Integer, ByVal Value As String)
'   With GridEditor1
'            If oTrans.ItemCountOthers <= .Row Then
'                .Rows = .Rows + 1
'                oTrans.addDetail
'            End If
'            oTrans.OtherDetail(.Row - 1, "xReferNox") = Value
'            If oTrans.OtherDetail(.Row - 1, "xReferNox") <> "" Then
'               .TextMatrix(.Row, 1) = Value
'               .TextMatrix(.Row, 0) = .Row
'            Else
'               oTrans.deleteDetail .Row
'               Exit Sub
'            End If
'
'      oTrans.OtherDetail(.Row - 1, "nQuantity") = Quantity
'      .TextMatrix(.Row, 1) = oTrans.OtherDetail(.Row - 1, "xReferNox")
'      .TextMatrix(.Row, 2) = oTrans.OtherDetail(.Row - 1, "sDescript")
'      .TextMatrix(.Row, 3) = oTrans.OtherDetail(.Row - 1, "nQtyOnHnd")
'      .TextMatrix(.Row, 4) = Quantity
''      .TextMatrix(.Row, 5) = oTrans.OtherDetail(.Row - 1, "dModified")
'
'      If Not pbHsSerial Then pbHsSerial = oTrans.OtherDetail(.Row - 1, "cHsSerial") = xeYes
'   End With
'End Sub
'
