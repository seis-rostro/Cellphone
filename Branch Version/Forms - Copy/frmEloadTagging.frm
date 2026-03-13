VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0"; "xrGridControl.ocx"
Begin VB.Form frmELoadTagging 
   BorderStyle     =   0  'None
   Caption         =   "Eload Tagging"
   ClientHeight    =   7800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7380
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7800
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   Begin xrGridEditor.GridEditor GridEditor1 
      Height          =   5895
      Left            =   90
      TabIndex        =   2
      Top             =   1155
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   10398
      AllowBigSelection=   -1  'True
      AutoAdd         =   0   'False
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
      Object.HEIGHT          =   5895
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
      MOUSEICON       =   "frmEloadTagging.frx":0000
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
      Height          =   585
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   1032
      BackColor       =   12632256
      BorderStyle     =   1
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
         Left            =   1515
         TabIndex        =   1
         Text            =   "Text"
         Top             =   120
         Width           =   4005
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Transact Date"
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
         Index           =   0
         Left            =   60
         TabIndex        =   0
         Top             =   150
         Width           =   1230
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   6030
      TabIndex        =   10
      Top             =   3705
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
      Picture         =   "frmEloadTagging.frx":001C
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   6030
      TabIndex        =   5
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
      Picture         =   "frmEloadTagging.frx":0796
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   6030
      TabIndex        =   8
      Top             =   2445
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
      Picture         =   "frmEloadTagging.frx":0F10
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   6030
      TabIndex        =   9
      Top             =   3075
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
      Picture         =   "frmEloadTagging.frx":168A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   6030
      TabIndex        =   6
      Top             =   1185
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Load"
      AccessKey       =   "L"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmEloadTagging.frx":1E04
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   660
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   7050
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   1164
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   1
         Left            =   3165
         MaxLength       =   30
         TabIndex        =   4
         Tag             =   "ht0"
         Top             =   45
         Width           =   2430
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   8
         Left            =   2340
         TabIndex        =   3
         Top             =   120
         Width           =   720
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   6030
      TabIndex        =   7
      Top             =   1815
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Delete"
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
      Picture         =   "frmEloadTagging.frx":257E
   End
End
Attribute VB_Name = "frmELoadTagging"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmEloadTagging"
Private WithEvents oTrans As clsCPLoadTagging
Attribute oTrans.VB_VarHelpID = -1
Private oLoadTrans As clsCPLoad

Private oSkin As clsFormSkin
Private pbGridFocus As Boolean
Private pnLastRow As Integer

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   Dim lnRep As Integer

   lsOldProc = "cmdButton_Click"
   ''On Error GoTo errProc

   With GridEditor1
      Select Case Index
      Case 0 'Browse
         If oTrans.SearchTransaction Then LoadDetail
         .SetFocus
      Case 1 'Load
         If oTrans.OpenTransaction(CDate(txtField(0).Text)) Then LoadDetail
         .SetFocus
      Case 2 'Delete
         If oLoadTrans.OpenTransaction(.TextMatrix(.Row, 5)) Then
            If oLoadTrans.DeleteTransaction Then
               txtField(1).Text = Format(oTrans.Master("nTranAmtx") - CDbl(.TextMatrix(.Row, 4)), "#,##0.00")
               .deleteRow
            End If
         End If
      Case 3 'Update
         If oLoadTrans.OpenTransaction(.TextMatrix(.Row, 5)) Then
            Dim oFormEload As frmEloadReg1
            
            Set oFormEload = New frmEloadReg1
            
            oFormEload.TransNox = oLoadTrans.Master("sTransNox")
            oFormEload.Show 1
            
            If oFormEload.Cancelled = False Then
               .TextMatrix(.Row, 1) = oFormEload.loadTrans.Master("sReferNox")
               .TextMatrix(.Row, 2) = oFormEload.loadTrans.Master("sBarrCode")
               .TextMatrix(.Row, 3) = oFormEload.loadTrans.Master("sPhoneNum")
               .TextMatrix(.Row, 4) = Format(oFormEload.loadTrans.Master("nTranAmtx"), "#,##0.00")
            End If

            Set oFormEload = Nothing
         End If
      Case 4 'Print
      Case 5 'Unload Me
         Unload Me
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
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String
   
   lsOldProc = "Form_Load"
   ''On Error GoTo errProc
   
   CenterChildForm mdiMain, Me
   
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransMaintenance
   
   Set oTrans = New clsCPLoadTagging
   Set oTrans.AppDriver = oApp

   oTrans.InitTransaction
   oTrans.NewTransaction
   
   Set oLoadTrans = New clsCPLoad
   Set oLoadTrans.AppDriver = oApp
   oLoadTrans.InitTransaction
   
   Call InitGrid
   Call ClearFields
   Call GridEditor1_Click
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub ClearFields()
   With GridEditor1
      .TextMatrix(1, 1) = ""
      .TextMatrix(1, 2) = ""
      .TextMatrix(1, 3) = ""
      .TextMatrix(1, 4) = "0.00"
      .Row = 1
   End With
   
   txtField(0).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
   txtField(1).Text = "0"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
   Set oTrans = Nothing
   Set oLoadTrans = Nothing
End Sub

Public Sub InitGrid()
   Dim lnCtr As Integer
   
   With GridEditor1
      .Cols = 6
      .Font = "MS Sans Serif"

      'column title
      .TextMatrix(0, 1) = "Reference No"
      .TextMatrix(0, 2) = "Barcode"
      .TextMatrix(0, 3) = "Mobile No"
      .TextMatrix(0, 4) = "Amount"
      .Row = 0

      'column alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = 3
         .ColEnabled(lnCtr) = False
      Next
      
      .ColWidth(0) = 330
      
      'column format
      .ColFormat(1) = ">"
      
      'column alignment
      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 6
      .ColAlignment(4) = 6
      
      'column width
      .ColWidth(1) = 1300
      .ColWidth(2) = 1750
      .ColWidth(3) = 1300
      .ColWidth(4) = 900
      .ColWidth(5) = 0
   End With
End Sub

Private Sub GridEditor1_Click()
   Dim lnCtr As Integer
   Dim lnRow As Integer

   With GridEditor1
      lnRow = .Row
      If pnLastRow <> 0 Then
         .Row = pnLastRow
         
         For lnCtr = 1 To .Cols - 1
            .Col = lnCtr
            .CellBackColor = &HFFFFFF
         Next
      End If
      
      .Row = lnRow
      For lnCtr = 1 To .Cols - 1
         .Col = lnCtr
         .CellBackColor = &HC0C0C0
      Next
      
      If .Row <> 0 Then pnLastRow = .Row
      .Col = 1
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
      oTrans.Detail(.Row - 1, .Col) = .TextMatrix(.Row, .Col)
   End With
End Sub

Private Sub GridEditor1_SelChange()
   Call GridEditor1_Click
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      If Index = 0 Then .Text = Format(.Text, "MM/DD/YYYY")
      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = oApp.getColor("HT1")
   End With

   pbGridFocus = False
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   With txtField(Index)
      If Not IsDate(.Text) Then .Text = oApp.ServerDate
      .Text = Format(.Text, "MMMM DD, YYYY")
   End With
End Sub

Private Sub GridEditor1_GotFocus()
   With GridEditor1
      .EditorBackColor = oApp.getColor("HT1")
   End With
   pbGridFocus = True
End Sub

Private Sub LoadDetail()
   Dim lnCtr As Integer
   Dim lnTotal As Currency
   
   With GridEditor1
      .Rows = oTrans.ItemCount + 1
      
      For lnCtr = 0 To oTrans.ItemCount - 1
         .TextMatrix(lnCtr + 1, 1) = oTrans.Detail(lnCtr, "sReferNox")
         .TextMatrix(lnCtr + 1, 2) = oTrans.Detail(lnCtr, "sBarrCode")
         .TextMatrix(lnCtr + 1, 3) = oTrans.Detail(lnCtr, "sPhoneNum")
         .TextMatrix(lnCtr + 1, 4) = Format(oTrans.Detail(lnCtr, "nAmountxx"), "#,##0.00")
         .TextMatrix(lnCtr + 1, 5) = oTrans.Detail(lnCtr, "sTransNox")
         lnTotal = lnTotal + oTrans.Detail(lnCtr, "nAmountxx")
      Next
      
      txtField(0).Text = Format(oTrans.Master("dTransact"), "MMMM DD, YYYY")
      txtField(1).Text = Format(oTrans.Master("nTranAmtx"), "#,##0.00")
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
