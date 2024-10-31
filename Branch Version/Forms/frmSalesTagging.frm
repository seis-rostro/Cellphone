VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0"; "xrGridControl.ocx"
Begin VB.Form frmSalesTagging 
   BorderStyle     =   0  'None
   Caption         =   "Sales Tagging"
   ClientHeight    =   7800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11715
   FillColor       =   &H00FFFFFF&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7800
   ScaleWidth      =   11715
   ShowInTaskbar   =   0   'False
   Begin xrGridEditor.GridEditor GridEditor1 
      Height          =   5895
      Left            =   90
      TabIndex        =   4
      Top             =   1155
      Width           =   10050
      _ExtentX        =   17727
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
      MOUSEICON       =   "frmSalesTagging.frx":0000
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
      Width           =   10035
      _ExtentX        =   17701
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
         Index           =   2
         Left            =   4785
         TabIndex        =   3
         Top             =   105
         Width           =   5115
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
         Left            =   1395
         TabIndex        =   1
         Text            =   "Text"
         Top             =   105
         Width           =   2400
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
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
         Index           =   1
         Left            =   3900
         TabIndex        =   2
         Top             =   150
         Width           =   780
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
      Index           =   3
      Left            =   10365
      TabIndex        =   10
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
      Picture         =   "frmSalesTagging.frx":001C
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   10365
      TabIndex        =   8
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
      Picture         =   "frmSalesTagging.frx":0796
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   10365
      TabIndex        =   9
      Top             =   1815
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
      Picture         =   "frmSalesTagging.frx":0F10
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   10365
      TabIndex        =   7
      Top             =   555
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
      Picture         =   "frmSalesTagging.frx":168A
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   660
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   7050
      Width           =   10050
      _ExtentX        =   17727
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
         Left            =   7485
         MaxLength       =   30
         TabIndex        =   6
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
         Left            =   6660
         TabIndex        =   5
         Top             =   120
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmSalesTagging"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pxeMODULENAME = "frmSalesTagging"
Private oFormPOSReg As frmCP_POSReg

Private oSkin As clsFormSkin
Private pbGridFocus As Boolean
Private pnLastRow As Integer
Private psCategID As String

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   Dim lnRep As Integer

   lsOldProc = "cmdButton_Click"
   ''On Error GoTo errProc

   With GridEditor1
      Select Case Index
      Case 0 'Load
         Call loadTransaction
         .SetFocus
      Case 1 'Update
         oFormPOSReg.TransNox = .TextMatrix(.Row, 1)
         oFormPOSReg.Show
      Case 2 'Print
      Case 3 'Close
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
   ''On Error GoTo errProc
   
   CenterChildForm mdiMain, Me
   Set oFormPOSReg = New frmCP_POSReg
    
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransMaintenance
   
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
      .TextMatrix(1, 5) = ""
      .Row = 1
   End With
   
   txtField(0).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
   txtField(1).Text = "0"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
   Set oFormPOSReg = Nothing
End Sub

Public Sub InitGrid()
   Dim lnCtr As Integer
   
   With GridEditor1
      .Cols = 6
      .Font = "MS Sans Serif"

      'column title
      .TextMatrix(0, 1) = "Trans No"
      .TextMatrix(0, 2) = "Full Name"
      .TextMatrix(0, 3) = "Sales Inv"
      .TextMatrix(0, 4) = "Amount"
      .TextMatrix(0, 5) = "Salesman"
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
      .ColAlignment(3) = 1
      .ColAlignment(4) = 6
      .ColAlignment(5) = 1
      
      'column width
      .ColWidth(1) = 1300
      .ColWidth(2) = 3500
      .ColWidth(3) = 1300
      .ColWidth(4) = 1000
      .ColWidth(5) = 2500
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

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If Index = 2 Then
      With txtField(Index)
         If KeyCode = vbKeyF3 Then
            SearchCategory False
            KeyCode = 0
         ElseIf KeyCode = vbKeyReturn Then
            If .Text <> "" Then SearchCategory False
            KeyCode = 0
         End If
      End With
   End If
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   With txtField(Index)
      If Index = 0 Then
         If Not IsDate(.Text) Then .Text = oApp.ServerDate
         .Text = Format(.Text, "MMMM DD, YYYY")
      ElseIf Index = 2 Then
         If Trim(.Text) = "" Then psCategID = ""
      End If
   End With
End Sub

Private Sub GridEditor1_GotFocus()
   With GridEditor1
      .EditorBackColor = oApp.getColor("HT1")
   End With
   pbGridFocus = True
End Sub

Private Sub SearchCategory(ByVal lbEqual As Boolean)
   Dim lors As ADODB.Recordset
   Dim lsSelect As String
   Dim lasSelect() As String
   Dim lsSQL As String
   Dim lnCtr As Integer

   With txtField(2)
      lsSQL = "SELECT" & _
                  "  a.sCategrID" & _
                  ", a.sCategrNm" & _
               " FROM Category a" & _
                  ", CP_Inventory b" & _
               " WHERE a.cRecdStat = " & strParm(xeRecStateActive) & _
                  " AND a.sCategrNm" & _
                     IIf(lbEqual, _
                        " = " & strParm(Trim(.Text)), _
                        " LIKE " & strParm((.Text) & "%")) & _
                  " AND a.sCategrID = b.sCategID1" & _
               " GROUP BY a.sCategrID" & _
               " ORDER BY a.sCategrNm"
      
      Set lors = New Recordset
      lors.Open lsSQL, oApp.Connection, , , adCmdText
   
      If lors.EOF Then
         .Text = ""
         psCategID = Empty
      ElseIf lors.RecordCount = 1 Then
         .Text = lors(1)
         psCategID = lors(0)
      Else
         lsSelect = KwikBrowse(oApp, lors _
                              , "sCategrID»sCategrNm" _
                              , "Code»Category")
      
         If lsSelect <> "" Then
            lasSelect = Split(lsSelect, "»")
            .Text = lasSelect(1)
            psCategID = lasSelect(0)
         Else
            If psCategID <> "" Then .Text = .Tag
         End If
      End If
      .Tag = .Text
      .SelStart = 0
      .SelLength = Len(.Text)
   End With

endProc:
   lors.Close
   Set lors = Nothing
   Exit Sub
End Sub

Private Sub loadTransaction()
   Dim lsProcName As String
   Dim lors As Recordset
   Dim lsSQL As String
   Dim lnCtr As Integer
   Dim lnTotl As Double
   
   lsProcName = "loadTransaction"
''On Error GoTo errProc
   
   lsSQL = "SELECT DISTINCT" & _
               "  a.sTransNox" & _
               ", CONCAT(b.sLastName, ', ', b.sFrstName, ' ', b.sMiddName) xFullName" & _
               ", a.dTransact" & _
               ", a.sSalesInv" & _
               ", a.nTranTotl" & _
               ", CONCAT(c.sFrstName, ' ', c.sLastName) xSalesman" & _
            " FROM CP_SO_Master a" & _
               " LEFT JOIN Client_Master b" & _
                  " ON a.sClientID = b.sClientID" & _
               " LEFT JOIN Salesman c" & _
                  " ON a.sSalesman = c.sEmployID" & _
               ", CP_SO_Detail d" & _
               ", CP_Inventory e" & _
            " WHERE a.sTransNox LIKE " & strParm(oApp.BranchCode & "%") & _
               " AND a.sTransNox = d.sTransNox" & _
               " AND d.sStockIDx = e.sStockIDx" & _
               " AND a.dTransact Between " & dateParm(txtField(0).Text) & _
                  " AND " & dateParm(txtField(0).Text & " 23:59:59") & _
               IIf(psCategID <> "", " AND e.sCategID1 = " & strParm(psCategID), "") & _
            " ORDER BY a.sTransNox"
   
   Set lors = New Recordset
   lors.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   
   With GridEditor1
      If lors.EOF Then
         .Rows = 2
         .TextMatrix(1, 1) = ""
         .TextMatrix(1, 2) = ""
         .TextMatrix(1, 3) = ""
         .TextMatrix(1, 4) = "0.00"
         .TextMatrix(1, 5) = ""
         
         MsgBox "No Record found!!!" & vbCrLf & _
                  "Please verify you entry then try again!!!", vbInformation, "Warning"
      Else
         .Rows = lors.RecordCount + 1
         lnTotl = 0#
         For lnCtr = 1 To lors.RecordCount
            .TextMatrix(lnCtr, 1) = lors("sTransNox")
            .TextMatrix(lnCtr, 2) = lors("xFullName")
            .TextMatrix(lnCtr, 3) = IFNull(lors("sSalesInv"), "")
            .TextMatrix(lnCtr, 4) = Format(lors("nTranTotl"), "#,##0.00")
            .TextMatrix(lnCtr, 5) = lors("xSalesman")
            lnTotl = lnTotl + lors("nTranTotl")
            lors.MoveNext
         Next
         txtField(1).Text = Format(lnTotl, "#,##0.00")
      End If
   End With
   
endProc:
   Set lors = Nothing
   Exit Sub
errProc:
   ShowError lsProcName
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
